"""
PESRP SIS Data Scraper - Streamlit Dashboard (Cloud & Local Compatible)
========================================================================
A web application to fetch school statistics (enrollment, teachers, school levels)
from the Punjab Education Sector Reform Programme (PESRP) SIS portal using EMIS codes.
"""

import io
import re
import time
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st
import urllib3
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Disable SSL warnings if insecure mode is enabled
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# Configuration & Constants
# ==========================================
BASE_URL = "https://sis.pesrp.edu.pk"
DASHBOARD_URL = f"{BASE_URL}/dashboard"
ENROLLMENT_URL = f"{BASE_URL}/dashboard_revamp/get_gender_summary_pie"
TEACHER_URL = f"{BASE_URL}/dashboard_revamp/get_teacher_designation_stats"
SCHOOL_LEVEL_URL = f"{BASE_URL}/dashboard_revamp/get_school_levels_count"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "en-US,en;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
}

DEFAULT_EMIS_CODES = [
    34510001,
    34510002,
    34510003,
    34510004,
    34510006,
    34510007,
    34510008,
    34510009,
    34510010,
    34510011,
    34510012,
    34510013,
    34510014,
    34510015,
    34510016,
    34510017,
    34510018,
    34510019,
    34510020,
    34510021,
    34510022,
    34510023,
    34510024,
    34510027,
    34510028,
    34510029,
    34510031,
    34510032,
    34510033,
    34510045,
    34510046,
    34510048,
    34510055,
    34510060,
    34510061,
    34510062,
    34510063,
    34510064,
    34510065,
    34510068,
    34510287,
    34510457,
    34510474,
    34510486,
    34510494,
    34510515,
    34510602,
    34510604,
    34510622,
    34510623,
    34510659,
    34520001,
    34520002,
    34520003,
    34520005,
    34520006,
    34520007,
    34520008,
    34520009,
    34520010,
    34520011,
    34520012,
    34520013,
    34520014,
    34520015,
    34520016,
    34520017,
    34520018,
    34520020,
    34520023,
    34520024,
    34520025,
    34520028,
    34520030,
    34520031,
    34520032,
    34520033,
    34520034,
    34520037,
    34520039,
    34520042,
    34520043,
    34520160,
    34520201,
    34520235,
    34520415,
    34520417,
    34520418,
    34520419,
    34520616,
    34520620,
]


def parse_int_value(val: Any) -> int:
    """Safely convert string/numeric values into clean integers."""
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str):
        cleaned = re.sub(r"[^\d-]", "", val.strip())
        return int(cleaned) if cleaned else 0
    return 0


class PESRPScraper:
    """Encapsulates HTTP session management and scraping operations for PESRP SIS."""

    def __init__(
        self,
        request_timeout: int = 15,
        verify_ssl: bool = True,
        proxy_url: Optional[str] = None,
    ):
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)
        self.timeout = request_timeout
        self.verify_ssl = verify_ssl
        self.csrf_token: Optional[str] = None

        if proxy_url:
            self.session.proxies = {
                "http": proxy_url,
                "https": proxy_url,
            }

    def initialize_session(self) -> str:
        """Establishes initial session and fetches the active CSRF token."""
        try:
            response = self.session.get(
                DASHBOARD_URL,
                timeout=self.timeout,
                verify=self.verify_ssl,
            )
            response.raise_for_status()
        except requests.exceptions.Timeout:
            raise RuntimeError(
                "Connection Timed Out! The target server 'sis.pesrp.edu.pk' took too long to respond. "
                "This usually happens when deployed on cloud hosting (Streamlit Cloud/AWS) "
                "because PESRP restricts or blocks foreign cloud IP addresses."
            )
        except requests.exceptions.HTTPError as e:
            if e.response.status_code in (403, 406):
                raise RuntimeError(
                    f"HTTP Error {e.response.status_code}: Access Denied! The PESRP server is "
                    "actively blocking requests from this cloud server's IP address (WAF/Geo-blocking)."
                )
            raise RuntimeError(f"HTTP Error {e.response.status_code}: {e}")
        except requests.exceptions.SSLError:
            raise RuntimeError(
                "SSL Certificate Verification Failed! Uncheck 'Verify SSL Certificates' in the sidebar settings."
            )
        except requests.exceptions.ConnectionError:
            raise RuntimeError(
                "Failed to establish connection to sis.pesrp.edu.pk. The target server may be blocking cloud IPs."
            )

        soup = BeautifulSoup(response.text, "html.parser")
        csrf_input = soup.find("input", {"name": "csrf_test_name"})

        if csrf_input and csrf_input.get("value"):
            self.csrf_token = csrf_input["value"]
        else:
            csrf_meta = soup.find("meta", {"name": "csrf_test_name"})
            if csrf_meta and csrf_meta.get("content"):
                self.csrf_token = csrf_meta["content"]
            else:
                raise ValueError("CSRF token could not be found on PESRP dashboard.")

        return self.csrf_token

    def _post_chart_endpoint(self, endpoint_url: str, emis_code: str) -> Dict[str, Any]:
        """Sends POST request to a given chart endpoint, keeping CSRF tokens fresh."""
        if not self.csrf_token:
            self.initialize_session()

        payload = {
            "csrf_test_name": self.csrf_token,
            "wing": "",
            "district": "",
            "tehsil": "",
            "markaz": "",
            "school": "",
            "classes": "",
            "s_id_emis_code": str(emis_code),
        }

        headers = {"Referer": DASHBOARD_URL}

        response = self.session.post(
            endpoint_url,
            data=payload,
            headers=headers,
            timeout=self.timeout,
            verify=self.verify_ssl,
        )
        response.raise_for_status()

        data = response.json()

        new_token = data.get("csrf_test_name")
        if new_token:
            self.csrf_token = new_token

        return data

    def fetch_school_record(
        self, emis_code: str, fetch_extended: bool = True
    ) -> Dict[str, Any]:
        """Fetches enrollment and optional teacher/level data for an EMIS code."""
        emis_str = str(emis_code).strip()

        enrollment_resp = self._post_chart_endpoint(ENROLLMENT_URL, emis_str)

        total_enrollment = parse_int_value(enrollment_resp.get("total"))
        boys = parse_int_value(
            enrollment_resp.get("male") or enrollment_resp.get("boys")
        )
        girls = parse_int_value(
            enrollment_resp.get("female") or enrollment_resp.get("girls")
        )

        record = {
            "EMIS Code": emis_str,
            "Total Enrollment": total_enrollment,
        }

        if fetch_extended:
            record["Boys"] = boys
            record["Girls"] = girls

            try:
                teacher_resp = self._post_chart_endpoint(TEACHER_URL, emis_str)
                total_teachers = parse_int_value(teacher_resp.get("total"))
                record["Total Teachers"] = total_teachers
            except Exception:
                record["Total Teachers"] = "N/A"

            try:
                level_resp = self._post_chart_endpoint(SCHOOL_LEVEL_URL, emis_str)
                school_level = (
                    level_resp.get("level_name")
                    or level_resp.get("school_level")
                    or "N/A"
                )
                record["School Level"] = school_level
            except Exception:
                record["School Level"] = "N/A"

        return record


# ==========================================
# Excel Formatting Helper
# ==========================================
def export_to_styled_excel(df: pd.DataFrame) -> bytes:
    """Generates a styled Excel report in bytes."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="School Stats")
        workbook = writer.book
        worksheet = writer.sheets["School Stats"]

        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        row_alt_fill = PatternFill(
            start_color="F2F5F9", end_color="F2F5F9", fill_type="solid"
        )
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        center_align = Alignment(horizontal="center", vertical="center")

        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        for row_idx, row in enumerate(
            worksheet.iter_rows(
                min_row=2, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)
            ),
            start=2,
        ):
            is_alt = row_idx % 2 == 0
            for cell in row:
                cell.border = border_thin
                if is_alt:
                    cell.fill = row_alt_fill

                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = center_align

        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    return output.getvalue()


# ==========================================
# Streamlit User Interface
# ==========================================
def main():
    st.set_page_config(
        page_title="PESRP School Data Scraper",
        page_icon="🏫",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("🏫 PESRP SIS School Data Extractor")
    st.markdown(
        "Extract live enrollment, teacher stats, and school levels from the "
        "**Punjab Education Sector Reform Programme (PESRP)** portal."
    )

    st.sidebar.header("⚙️ Configuration")

    input_mode = st.sidebar.radio(
        "EMIS Input Source:",
        ["Use Default List (91 Schools)", "Custom EMIS Codes"],
    )

    if input_mode == "Custom EMIS Codes":
        raw_input = st.sidebar.text_area(
            "Enter EMIS Codes (separated by comma, space, or new line):",
            value="34510001\n34510002\n34510003",
            height=150,
        )
        parsed_codes = re.findall(r"\b\d+\b", raw_input)
        emis_list = [int(c) for c in parsed_codes]
    else:
        emis_list = DEFAULT_EMIS_CODES

    st.sidebar.subheader("Extraction Options")
    fetch_extended = st.sidebar.checkbox(
        "Fetch Extended Data (Teachers, Gender, Level)",
        value=True,
    )
    request_delay = st.sidebar.slider(
        "Request Delay (seconds):",
        min_value=0.0,
        max_value=2.0,
        value=0.3,
        step=0.1,
    )

    st.sidebar.subheader("Network & Cloud Settings")
    verify_ssl = st.sidebar.checkbox("Verify SSL Certificates", value=True)
    timeout_val = st.sidebar.number_input(
        "Request Timeout (sec):", value=15, min_value=5, max_value=60
    )

    use_proxy = st.sidebar.checkbox("Use HTTP/HTTPS Proxy (Geo-Bypass)")
    proxy_url = None
    if use_proxy:
        proxy_url = st.sidebar.text_input(
            "Proxy URL (e.g., http://user:pass@proxy.example.com:8080):",
            value="",
        )

    st.sidebar.markdown("---")
    st.sidebar.info(f"📍 **Total EMIS Codes to process:** `{len(emis_list)}`")

    start_btn = st.button(
        "🚀 Start Scraping Data", type="primary", use_container_width=True
    )

    if "scraped_df" not in st.session_state:
        st.session_state.scraped_df = None

    if start_btn:
        if not emis_list:
            st.error("Please provide at least one valid numeric EMIS code.")
            return

        scraper = PESRPScraper(
            request_timeout=timeout_val,
            verify_ssl=verify_ssl,
            proxy_url=proxy_url if use_proxy and proxy_url else None,
        )

        with st.status("Connecting to PESRP SIS Portal...", expanded=True) as status:
            try:
                scraper.initialize_session()
                st.write("✅ Established session and retrieved active CSRF token.")
            except Exception as e:
                status.update(label="❌ Connection to PESRP failed", state="error")
                st.error(f"**Error Message:** {e}")
                return

            progress_bar = st.progress(0)
            status_text = st.empty()

            scraped_records = []
            total_count = len(emis_list)

            for idx, emis in enumerate(emis_list, start=1):
                status_text.text(f"Fetching record [{idx}/{total_count}]: EMIS {emis}")

                try:
                    record = scraper.fetch_school_record(
                        emis, fetch_extended=fetch_extended
                    )
                    scraped_records.append(record)
                except Exception as err:
                    st.warning(f"⚠️ Failed to fetch EMIS {emis}: {err}")

                progress_bar.progress(idx / total_count)
                if request_delay > 0:
                    time.sleep(request_delay)

            status.update(
                label=f"🎉 Successfully scraped {len(scraped_records)} schools!",
                state="complete",
            )

        if scraped_records:
            df = pd.DataFrame(scraped_records)
            df = df.sort_values(by="EMIS Code").reset_index(drop=True)
            st.session_state.scraped_df = df

    if st.session_state.scraped_df is not None:
        df = st.session_state.scraped_df

        st.markdown("---")
        st.subheader("📊 Scraped Results")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Schools Scraped", len(df))
        m2.metric("Total Enrollment", f"{df['Total Enrollment'].sum():,}")

        if "Boys" in df.columns and "Girls" in df.columns:
            m3.metric("Total Boys", f"{df['Boys'].sum():,}")
            m4.metric("Total Girls", f"{df['Girls'].sum():,}")

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

        col_dl1, col_dl2 = st.columns(2)

        excel_bytes = export_to_styled_excel(df)
        col_dl1.download_button(
            label="📥 Download Excel Report (.xlsx)",
            data=excel_bytes,
            file_name="PESRP_School_Statistics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        csv_bytes = df.to_csv(index=False).encode("utf-8")
        col_dl2.download_button(
            label="📄 Download CSV File (.csv)",
            data=csv_bytes,
            file_name="PESRP_School_Statistics.csv",
            mime="text/csv",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
