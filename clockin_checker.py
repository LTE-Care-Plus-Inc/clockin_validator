import io
from datetime import datetime, date, time, timedelta
from zoneinfo import ZoneInfo
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

TZ = ZoneInfo("America/New_York")

st.set_page_config(page_title="BT Session Start Checker", layout="wide")
st.title("BT Session Start Checker")

st.markdown("""
Upload the two files below.  
**File 1 (HiRasmus)** must include: `Start Time`, `Appointment ID`, `Status`  
**File 2 (Aloha Sessions)** must include: `Staff Name`, `Client Name`, `Appointment ID`, `Appt. Start Time`, `Service Name`  

**Logic overview**  
1) Filter Aloha to `Service Name == "Direct Service BT"`  
2) Join on `Appointment ID`  
3) Compare `Appt. Start Time` (scheduled) to `Start Time` (actual) with a ± minutes tolerance (default 5).  
4) “Flagged” = no start yet for a past session or outside tolerance.  
5) “Future” = scheduled in the future.  
6) “Valid” = started within tolerance.
""")

# --- Sidebar controls ---
with st.sidebar:
    st.header("Options")
    tolerance_min = st.number_input("Tolerance (± minutes)", min_value=0, max_value=120, value=5, step=1)
    only_check_current_minute = st.checkbox("Only check the current minute (e.g., 3:05 checks 3:05)", value=False)
    assume_today_if_no_date = st.checkbox("If no Appt. Date column, assume today (local time)", value=True)
    show_tables = st.checkbox("Preview results in app", value=True)

# --- Uploaders ---
col1, col2 = st.columns(2)
with col1:
    f_hirasmus = st.file_uploader("Upload File 1: HiRasmus", type=["csv", "xlsx"])
with col2:
    f_aloha = st.file_uploader("Upload File 2: Aloha Sessions", type=["csv", "xlsx"])


def read_any(file):
    if file is None:
        return None
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file, dtype=str)
    return pd.read_excel(file, dtype=str, engine="openpyxl")


def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()
    return df


def parse_dt_series(series, fmt=None, utc=False):
    # robust parser
    return pd.to_datetime(series, format=fmt, errors="coerce", utc=utc)


def drop_tz(series):
    s = pd.to_datetime(series, errors="coerce")
    try:
        # If tz-aware → convert to ET (for consistency) then drop tz
        return s.dt.tz_convert(TZ).dt.tz_localize(None)
    except Exception:
        # If already tz-naive (or NaT), just return as naive
        return s.dt.tz_localize(None) if getattr(s.dt, "tz", None) is not None else s
    
def parse_time_cell(val):
            s = (str(val) if val is not None else "").strip()
            # Try common time formats
            for fmt in ("%I:%M:%S %p", "%I:%M %p", "%H:%M:%S", "%H:%M"):
                try:
                    return datetime.strptime(s, fmt).time()
                except Exception:
                    pass
            return None  # unparseable

def combine_date_and_time(date_series: pd.Series, time_series: pd.Series):
    # Combines (possibly-na) date + time into localized datetimes
    dt = []
    for d, t in zip(date_series, time_series):
        t_parsed = pd.to_datetime(str(t), format="%I:%M:%S %p", errors="coerce")
        if pd.isna(t_parsed):
            dt.append(pd.NaT)
            continue
        if pd.notna(d):
            # If d is already datetime-like, cast to date
            if isinstance(d, pd.Timestamp):
                d = d.date()
            else:
                try:
                    d = pd.to_datetime(d, errors="coerce").date()
                except Exception:
                    d = None
        if d is None or pd.isna(d):
            dt.append(pd.NaT)
        else:
            dt.append(datetime.combine(d, t_parsed.time()).replace(tzinfo=TZ))
    return pd.to_datetime(pd.Series(dt), errors="coerce")


def local_now():
    return datetime.now(TZ).replace(second=0, microsecond=0)


def round_to_minute(dt_obj: datetime):
    if pd.isna(dt_obj):
        return dt_obj
    return dt_obj.replace(second=0, microsecond=0)


# --- Process after upload ---
if f_hirasmus is not None and f_aloha is not None:
    try:
        df_hi = read_any(f_hirasmus)
        df_al = read_any(f_aloha)
        df_hi = normalize_cols(df_hi)
        df_al = normalize_cols(df_al)

        # Validate required columns
        req_hi = {"Start Time", "Appointment ID", "Status"}
        req_al = {"Staff Name", "Client Name", "Appointment ID", "Appt. Start Time", "Service Name"}
        missing_hi = req_hi - set(df_hi.columns)
        missing_al = req_al - set(df_al.columns)
        if missing_hi:
            st.error(f"HiRasmus is missing required columns: {sorted(missing_hi)}")
            st.stop()
        if missing_al:
            st.error(f"Aloha Sessions is missing required columns: {sorted(missing_al)}")
            st.stop()

        # Step 1: filter Aloha to Direct Service BT
        df_al = df_al[df_al["Service Name"] == "Direct Service BT"].copy()
        df_al = df_al[df_al["Service Name"] == "Direct Service BT"].copy()
        df_al.reset_index(drop=True, inplace=True)

        # --- Robust scheduled datetime from Appt. Start Time (+ optional date) ---

        # Try to detect an appointment date column if present
        date_cols = [c for c in df_al.columns if c.strip().lower() in {"appt. date", "appointment date", "date"}]
        appt_date_col = date_cols[0] if date_cols else None

        # Parse optional date column (if any) *after* reset_index so it aligns by position
        if appt_date_col:
            appt_dates = pd.to_datetime(df_al[appt_date_col], errors="coerce")
        else:
            appt_dates = pd.Series(pd.NaT, index=df_al.index, dtype="datetime64[ns]")

        today_local = datetime.now(TZ).date()

        scheduled_list = []
        for pos in range(len(df_al)):
            row = df_al.iloc[pos]
            t = parse_time_cell(row.get("Appt. Start Time", None))
            if t is None:
                scheduled_list.append(pd.NaT)
                continue

            d = appt_dates.iloc[pos]
            if pd.isna(d):
                if assume_today_if_no_date:
                    d = today_local
                else:
                    scheduled_list.append(pd.NaT)
                    continue
            else:
                d = d.date()

            dt_local = datetime.combine(d, t).replace(tzinfo=TZ)
            scheduled_list.append(dt_local)

        df_al["_scheduled_dt"] = pd.to_datetime(scheduled_list, errors="coerce")
        df_al["_scheduled_dt"] = df_al["_scheduled_dt"].apply(round_to_minute)

        # Join prep: For HiRasmus, parse Start Time and localize to ET
        actual = pd.to_datetime(df_hi["Start Time"], errors="coerce")  # tz-naive
        if getattr(actual.dt, "tz", None) is None:
            actual = actual.dt.tz_localize(TZ)   # make it timezone-aware
        else:
            actual = actual.dt.tz_convert(TZ)    # if already aware, convert to our TZ
        df_hi["_actual_start_dt"] = actual.apply(round_to_minute)

        # If multiple rows per Appointment ID in HiRasmus, keep earliest start
        df_hi_sorted = df_hi.sort_values(by=["Appointment ID", "_actual_start_dt"])
        df_hi_min = df_hi_sorted.groupby("Appointment ID", as_index=False).first()

        # Step 2: join on Appointment ID
        merged = pd.merge(
            df_al,
            df_hi_min[["Appointment ID", "_actual_start_dt", "Start Time", "Status"]],
            on="Appointment ID",
            how="left",
            suffixes=("", "_hi"),
        )

        # --- Normalize TZ so both columns are tz-aware in ET ---
        def to_et(series):
            s = pd.to_datetime(series, errors="coerce")
            # If tz-naive, localize; if tz-aware, convert
            if getattr(s.dt, "tz", None) is None:
                return s.dt.tz_localize(TZ)
            else:
                return s.dt.tz_convert(TZ)

        merged["_scheduled_dt"] = to_et(merged["_scheduled_dt"])
        merged["_actual_start_dt"] = to_et(merged["_actual_start_dt"])

        # Compute minutes difference (actual - scheduled)
        merged["minutes_diff"] = (
            merged["_actual_start_dt"] - merged["_scheduled_dt"]
        ).dt.total_seconds() / 60.0

        # Current time logic
        now_local = local_now()

        # Only check the current minute if requested
        to_check = merged.copy()
        if only_check_current_minute:
            to_check = to_check[to_check["_scheduled_dt"] == now_local]

        # Classify rows
        future_mask = to_check["_scheduled_dt"] > now_local
        live_mask = ~future_mask

        # Flag reasons
        def reason_row(row):
            # If Appt. Start Time cell itself is missing/unparseable, call that out
            if pd.isna(row.get("_scheduled_dt")):
                raw_time = str(row.get("Appt. Start Time", "")).strip()
                return "Missing/invalid Appt. Start Time" if not raw_time else "Invalid Appt. Start Time format"

            if row["_scheduled_dt"] > now_local:
                return "Future session"
            if pd.isna(row.get("_actual_start_dt")):
                return "No start recorded (not started)"
            if pd.isna(row.get("minutes_diff")):
                return "Missing time difference"

            if abs(row["minutes_diff"]) > tolerance_min:
                delta = int(round(row["minutes_diff"]))
                return f"Outside tolerance (Δ={delta} min)"
            return "Within tolerance"

        to_check["Reason"] = to_check.apply(reason_row, axis=1)

        # Split buckets
        df_future = to_check[future_mask].copy()
        df_past_now = to_check[live_mask].copy()
        df_valid = df_past_now[df_past_now["Reason"] == "Within tolerance"].copy()
        df_flagged = df_past_now[df_past_now["Reason"] != "Within tolerance"].copy()

        # Arrange columns for output
        common_cols = [
            "Staff Name", "Client Name", "Appointment ID",
            "Service Name", "Appt. Start Time", "Start Time", "Status",
            "_scheduled_dt", "_actual_start_dt", "minutes_diff", "Reason"
        ]

        def safe_cols(df):
            return [c for c in common_cols if c in df.columns]

        out_valid = df_valid[safe_cols(df_valid)].sort_values(
            by=["_scheduled_dt", "Staff Name", "Client Name"], na_position="last"
        )
        out_flagged = df_flagged[safe_cols(df_flagged)].sort_values(
            by=["_scheduled_dt", "Staff Name", "Client Name"], na_position="last"
        )
        out_future = df_future[safe_cols(df_future)].sort_values(
            by=["_scheduled_dt", "Staff Name", "Client Name"], na_position="last"
        )

        # Optional preview
        if show_tables:
            st.subheader("Valid")
            st.dataframe(out_valid, use_container_width=True, hide_index=True)
            st.subheader("Flagged")
            st.dataframe(out_flagged, use_container_width=True, hide_index=True)
            st.subheader("Future")
            st.dataframe(out_future, use_container_width=True, hide_index=True)

        buffer = io.BytesIO()

        # Make export-safe (tz-naive) copies
        def prepare_for_excel(df):
            df2 = df.copy()
            for col in ["_scheduled_dt", "_actual_start_dt"]:
                if col in df2.columns:
                    df2[col] = drop_tz(df2[col])
            return df2

        export_valid = prepare_for_excel(out_valid)
        export_flagged = prepare_for_excel(out_flagged)
        export_future = prepare_for_excel(out_future)

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_valid.to_excel(writer, sheet_name="Valid", index=False)
            export_flagged.to_excel(writer, sheet_name="Flagged", index=False)
            export_future.to_excel(writer, sheet_name="Future", index=False)

            # Autosize columns (openpyxl)
            for sheet_name, df_for in {
                "Valid": export_valid,
                "Flagged": export_flagged,
                "Future": export_future,
            }.items():
                ws = writer.sheets[sheet_name]
                for idx, col_name in enumerate(df_for.columns, start=1):
                    sample = [str(col_name)] + [str(v) for v in df_for[col_name].head(500).tolist()]
                    width = min(max(len(s) for s in sample), 48) + 2
                    ws.column_dimensions[get_column_letter(idx)].width = width

        st.download_button(
            "Download Excel (Valid / Flagged / Future)",
            data=buffer.getvalue(),
            file_name=f"bt_session_check_{datetime.now(TZ).strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Info footer
        st.caption(
            f"Now (local): **{now_local.strftime('%Y-%m-%d %I:%M %p')}** • "
            f"Tolerance: ±{tolerance_min} min • "
            f"{'Checking only current minute' if only_check_current_minute else 'Checking all sessions up to now'} • "
            f"{'Assuming today for scheduled times without dates' if (appt_date_col is None and assume_today_if_no_date) else ''}"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload both files to proceed.")
