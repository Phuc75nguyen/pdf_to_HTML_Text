"""
Parser module for extracting booking information from OTA booking PDF emails.

This module supports parsing booking confirmation and cancellation emails from
Expedia and Agoda.  It extracts the most relevant fields (such as guest name,
booking dates, room type, rates and payment details) and normalises them
according to a fixed schema defined by ``FIELD_ORDER``.  It also includes
functions to render the parsed data into a simple HTML report and a plain
text file.

Usage::

    from parser import process_pdf

    txt_path, html_path = process_pdf('/path/to/file.pdf')

The ``process_pdf`` function will automatically detect whether the PDF
belongs to Expedia or Agoda based on its contents.  If the source cannot
be identified or the parser cannot extract any information, it will
raise a ``ValueError``.

Note: this module does not perform OCR.  It relies on textual content
extracted by ``pdfplumber``.  Therefore, PDFs containing only images will
not be parsed correctly.
"""

from __future__ import annotations

import os
import re
import calendar
from datetime import datetime
from collections import OrderedDict
from typing import Dict, Tuple

try:
    # pdfplumber is required for reading PDF text.  Some environments may
    # not have it installed.  In that case, process_pdf will raise a
    # descriptive ImportError.
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # type: ignore

# ---------------------------------------------------------------------------
# Configuration: field order and billing subfield order
#
# The keys defined here correspond exactly to the required output schema.  If
# you change these names you must also update any downstream consumers.

# Exact order of the fields in the final output (JSON-like).  This order is
# preserved in the generated text and HTML reports.
FIELD_ORDER = [
    "Status booking Reservation",
    "Customer First Name",
    "Customer Last Name",
    "Email Customer",
    "BookingID",
    "Has Prepaid",
    "Booked on",
    "Check in",
    "Check out",
    "Special Request",
    "Room Type Code",
    "No. of room",
    "Occupancy Adult",
    "Occupancy Childrent",
    "Daily Rate",
    "Total Booking",
    "Amount to Charge Expedia",
    "Billing Details:",
]

# Order of subfields inside the billing details dictionary
BILLING_ORDER = [
    "Card Number",
    "Activation Date",
    "Expiration Date",
    "Validation Code",
]

# Allowed currency codes for rate extraction.  If your OTA uses additional
# currency codes, add them here.
CURRENCY_PATTERN = r"(VND|USD|EUR|JPY|THB|SGD|AUD|GBP|KRW|CNY)"

# Month name mapping for converting textual month names to numeric values.
_MONTH_MAP = {
    'jan': 1, 'january': 1,
    'feb': 2, 'february': 2,
    'mar': 3, 'march': 3,
    'apr': 4, 'april': 4,
    'may': 5,
    'jun': 6, 'june': 6,
    'jul': 7, 'july': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'september': 9,
    'oct': 10, 'october': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12,
}


def ordered_output(data: Dict[str, any]) -> OrderedDict:
    """Return a new OrderedDict with keys sorted according to FIELD_ORDER.

    Missing keys are added with a sensible default (empty string or False).

    Args:
        data: A dict containing parsed booking information.

    Returns:
        OrderedDict: Data sorted and padded according to FIELD_ORDER.
    """
    result = OrderedDict()
    for key in FIELD_ORDER:
        if key == "Billing Details:":
            sub = data.get(key, {}) or {}
            # ensure subfields are ordered and present
            bill = OrderedDict()
            for subkey in BILLING_ORDER:
                bill[subkey] = sub.get(subkey, "")
            result[key] = bill
        else:
            # boolean fields use False as default; others use empty string
            default = False if key == "Has Prepaid" else ""
            result[key] = data.get(key, default)
    return result


def norm_text(text: str) -> str:
    """Normalise whitespace in extracted PDF text.

    This helper collapses multiple spaces and tabs into a single space and
    compresses consecutive newlines to a single newline.  It also strips
    leading/trailing whitespace.  Non-breaking spaces (\xa0) are replaced
    with regular spaces.

    Args:
        text: The raw extracted text from a PDF page.

    Returns:
        A normalised string.
    """
    if not text:
        return ""
    # Replace non-breaking spaces
    text = text.replace("\xa0", " ")
    # Collapse runs of spaces and tabs
    text = re.sub(r"[ \t]+", " ", text)
    # Collapse multiple newlines
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def parse_month_day_year(month: str, day: str, year: str) -> str:
    """Convert textual month/day/year parts to an MM/DD/YYYY string.

    Args:
        month: Month name (abbreviated or full).
        day: Day of month as a string.
        year: Year as a string.

    Returns:
        A string formatted as MM/DD/YYYY if the month is recognised,
        otherwise the original concatenation of inputs.
    """
    try:
        mm = _MONTH_MAP[month.strip().lower()]
        dd = int(day)
        yyyy = int(year)
        return f"{mm:02d}/{dd:02d}/{yyyy}"
    except Exception:
        return f"{month} {day}, {year}"


def mdy_to_mmddyyyy(date_str: str) -> str:
    """Convert a date string like 'November 16, 2025' to '11/16/2025'.

    If conversion fails, the original string is returned unchanged.
    """
    if not date_str:
        return ""
    date_str = date_str.strip()
    # Remove extraneous commas or periods
    date_str = date_str.replace("\u00a0", " ")  # nbsp
    try:
        # Attempt abbreviated month (Nov)
        dt = datetime.strptime(date_str, "%b %d, %Y")
        return dt.strftime("%m/%d/%Y")
    except Exception:
        pass
    try:
        # Attempt full month (November)
        dt = datetime.strptime(date_str, "%B %d, %Y")
        return dt.strftime("%m/%d/%Y")
    except Exception:
        pass
    # Fallback: attempt manual parse
    parts = re.match(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", date_str)
    if parts:
        return parse_month_day_year(parts.group(1), parts.group(2), parts.group(3))
    return date_str


def parse_vi_sent_datetime(text: str) -> str:
    """Extract a booking 'sent' date from Vietnamese header lines.

    This function looks for common date patterns in Vietnamese booking
    emails, such as ``Ngày T2 10/11/2025 10:51`` or ``Đã gửi: Thứ Hai, 10 tháng 11, 2025 ...``.
    It returns the date in MM/DD/YYYY format.  If nothing is found,
    returns an empty string.
    """
    if not text:
        return ""
    # First try to match a common DD/MM/YYYY pattern which appears in Vietnamese
    # headers such as "Ngày T2 10/11/2025 10:51".  Because the typical
    # ordering in Vietnamese is day/month/year, we convert to month/day/year.
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", text)
    if m:
        dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{mm:02d}/{dd:02d}/{yyyy}"
    # Next try to match a pattern like "DD tháng MM, YYYY", which also
    # expresses day-of-month followed by the word "tháng" (month) and then
    # the month number and year.  This pattern appears in the Vietnamese
    # version of "Đã gửi" lines.
    m = re.search(r"(\d{1,2})\s+th[aá]ng\s+(\d{1,2}),\s*(\d{4})", text, re.I)
    if m:
        dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{mm:02d}/{dd:02d}/{yyyy}"
    return ""


def find_agoda_first_daily_rate(text: str) -> str:
    """Find the first daily rate in an Agoda booking email.

    Agoda rates are usually presented in a table following a 'From - To' and
    'Rates' header.  Each date line is followed by a currency code ('VND',
    etc.) on its own line and then the numeric amount on the next line.

    This function locates the first occurrence of a pattern like ``VND\n1,197,000.00``
    after the 'From - To' table header.  Only numbers with comma thousands
    separators are considered.  The fractional portion (``.00``) is removed.

    Returns:
        A string like ``1,197,000 VND`` or an empty string if not found.
    """
    if not text:
        return ""
    # Locate the table header to limit our search scope
    anchor = re.search(r"From\s*-\s*To\s*Rates", text, re.I)
    tail = text[anchor.end():] if anchor else text
    # Look for VND on its own line followed by a line with a number containing commas
    m = re.search(
        r"\bVND\b\s*\n\s*(\d{1,3}(?:,\d{3})+)(?:\.\d+)?",
        tail, re.I)
    if m:
        return f"{m.group(1)} VND"
    return ""


def find_agoda_total_gross(text: str) -> str:
    """Find the gross (sell) rate in an Agoda booking email.

    Agoda confirmation emails typically include both a 'Reference sell rate'
    (gross) and a 'Net rate'.  The gross rate is the total amount the guest
    paid before commission.  This function returns the gross rate if found.

    Returns:
        A string like ``7,581,000 VND`` or an empty string if not found.
    """
    if not text:
        return ""
    # Search for 'Reference sell rate' followed by a currency and a number
    m = re.search(
        r"Reference sell rate.*?\bVND\b\s*(\d{1,3}(?:,\d{3})+)(?:\.\d+)?",
        text, re.I | re.S)
    if m:
        return f"{m.group(1)} VND"
    return ""


def find_expedia_amount(t: str) -> str:
    """Extract the amount to charge Expedia from an Expedia booking email.

    In Expedia cancellation or confirmation emails the amount to charge
    may be presented either on the same line as the label or on the line(s)
    following it, optionally with the word 'Group'.  This helper will
    recognise several common patterns and return the numerical amount with
    currency.  If no currency is found, just the number is returned.
    """
    amount = None
    # Patterns where the amount and currency are close together
    patterns = [
        rf"Amount to Charge Expedia(?:\s*Group)?\s*:\s*([\d,.]+)\s*{CURRENCY_PATTERN}",
        rf"Amount to Charge Expedia(?:\s*Group)?\s*:\s*[\r\n ]+([\d,.]+)\s*{CURRENCY_PATTERN}",
        rf"Amount to Charge Expedia(?:\s*Group)?\s+([\d,.]+)\s*{CURRENCY_PATTERN}",
    ]
    for pat in patterns:
        m = re.search(pat, t, re.I | re.S)
        if m:
            amount = f"{m.group(1)} {m.group(2)}"
            break
    # Fallback: just find the first number after the label (without currency)
    if amount is None:
        anchor = re.search(r"Amount to Charge Expedia", t, re.I)
        if anchor:
            tail = t[anchor.end():anchor.end() + 200]
            m = re.search(r"(?<!\d)(\d{1,3}(?:[.,]\d{3})+|\d+)(?![\d,])", tail)
            if m:
                amount = m.group(1)
    return amount or ""


class ExpediaParser:
    """Parser for Expedia booking emails.

    This class extracts data from Expedia confirmation and cancellation
    PDFs.  It assumes the PDF text has already been normalised.
    """

    def __init__(self, text: str) -> None:
        self.text = norm_text(text)
        self.data: Dict[str, any] = {}

    def parse(self) -> OrderedDict:
        t = self.text
        # Status (Cancelled or Confirmed)
        self.data["Status booking Reservation"] = (
            "Cancelled" if re.search(r"\b(Cancellation|Cancelled on)\b", t, re.I) else "Confirmed"
        )
        # Guest email
        m = re.search(r"Guest Email:\s*([^\s]+@[^\s]+)", t, re.I)
        if m:
            self.data["Email Customer"] = m.group(1)
        # Booking ID
        m = re.search(r"Reservation ID:\s*(\d+)", t, re.I)
        if m:
            self.data["BookingID"] = m.group(1)
        # Has Prepaid
        self.data["Has Prepaid"] = bool(re.search(r"Guest has PRE-PAID", t, re.I))
        # Booked on (convert to MM/DD/YYYY)
        m = re.search(r"Booked on:\s*(.+?)\n", t, re.I)
        if m:
            self.data["Booked on"] = mdy_to_mmddyyyy(m.group(1))
        # Room Type Code (may be labelled as 'Room Type Code' or 'Room Type Name')
        m = re.search(r"Room Type Code:\s*(.+?)\n", t)
        if not m:
            m = re.search(r"Room Type Name:\s*(.+?)(?:\s*-\s*Non-refundable)?\s*\n", t)
        if m:
            self.data["Room Type Code"] = m.group(1).strip()
        # Rates
        # Daily Rate (match daily base rate line)
        m = re.search(r"Daily Base Rate.*?-\s*([\d,.]+)\s*" + CURRENCY_PATTERN, t, re.I | re.S)
        if m:
            self.data["Daily Rate"] = f"{m.group(1)} {m.group(2)}"
        # Total Booking Amount or Price
        m = re.search(r"(Total Booking Amount|Total Booking Price)\s*:?[\s]*([\d,.]+)\s*" + CURRENCY_PATTERN, t, re.I | re.S)
        if m:
            self.data["Total Booking"] = f"{m.group(2)} {m.group(3)}"
        # Amount to Charge Expedia
        amount = find_expedia_amount(t)
        if amount:
            self.data["Amount to Charge Expedia"] = amount
        # Stay details: Check in/out and occupancy
        stay = re.search(
            r"Check-In\s+Check-Out\s+Adults\s+Kids/Ages.*?\n"
            r"([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\s+"
            r"([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\s+"
            r"(\d+)\s+"
            r"(\d+)",
            t, re.I | re.S)
        if stay:
            self.data["Check in"] = mdy_to_mmddyyyy(stay.group(1))
            self.data["Check out"] = mdy_to_mmddyyyy(stay.group(2))
            self.data["Occupancy Adult"] = stay.group(3)
            self.data["Occupancy Childrent"] = stay.group(4)
        # Special request (Expedia emails seldom include this; leave empty)
        self.data.setdefault("Special Request", "")
        # No. of room (Expedia emails typically do not specify; assume 1)
        self.data.setdefault("No. of room", "1")
        # Billing details (Expedia virtual credit card)
        billing: Dict[str, str] = {}
        m = re.search(r"Card Number\s+([\d-]+)", t, re.I)
        if m:
            billing["Card Number"] = m.group(1)
        m = re.search(r"Activation Date\s+([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})", t, re.I)
        if m:
            billing["Activation Date"] = mdy_to_mmddyyyy(m.group(1))
        m = re.search(r"Expiration Date\s*(.+?)\n", t, re.I)
        if m:
            # Some emails include location after the date; we strip extra tokens
            raw = m.group(1).strip()
            # If the line contains a comma and letters (e.g. 'Sep 2030 Seattle'), take the first two tokens
            parts = raw.split()
            if len(parts) >= 2:
                # If month and year only, we need to compute last day of month
                possible_month = parts[0]
                possible_year = parts[1].rstrip(',')
                try:
                    month_num = _MONTH_MAP[possible_month.lower()[:3]]
                    year_num = int(possible_year)
                    last_day = calendar.monthrange(year_num, month_num)[1]
                    billing["Expiration Date"] = f"{month_num:02d}/{last_day:02d}/{year_num}"
                except Exception:
                    billing["Expiration Date"] = raw
            else:
                billing["Expiration Date"] = raw
        m = re.search(r"Validation Code\s+(\d+)", t, re.I)
        if m:
            billing["Validation Code"] = m.group(1)
        self.data["Billing Details:"] = billing
        return ordered_output(self.data)


class AgodaParser:
    """Parser for Agoda booking emails.

    This parser extracts booking information from Agoda confirmation emails.
    Agoda emails have a different structure than Expedia: they often include
    both gross and net rates, and they hide credit card details.  The
    extracted data is normalised to the same output schema, with certain
    fields left blank if not applicable.
    """

    def __init__(self, text: str) -> None:
        self.text = norm_text(text)
        self.data: Dict[str, any] = {}

    def parse(self) -> OrderedDict:
        t = self.text
        # Booking status: Agoda confirmations for prepaid bookings
        self.data["Status booking Reservation"] = "Confirmed"
        # Has Prepaid if 'PREPAID' appears
        self.data["Has Prepaid"] = bool(re.search(r"\bPREPAID\b", t, re.I))
        # Booking ID
        m = re.search(r"Booking ID\s*(\d+)", t, re.I)
        if m:
            self.data["BookingID"] = m.group(1)
        # Guest name (first & last)
        mfirst = re.search(r"Customer First Name\s+([A-ZÀ-Ỹ' \-]+)", t)
        mlast = re.search(r"Customer Last Name\s+([A-ZÀ-Ỹ' \-]+)", t)
        if mfirst:
            self.data["Customer First Name"] = mfirst.group(1).title()
        if mlast:
            self.data["Customer Last Name"] = mlast.group(1).title()
        # Email
        m = re.search(r"Email:\s*([^\s]+@[^\s]+agoda-messaging\.com)", t, re.I)
        if m:
            self.data["Email Customer"] = m.group(1)
        # Booked on: parse Vietnamese date header or English variant
        booked = parse_vi_sent_datetime(t)
        if booked:
            self.data["Booked on"] = booked
        # Check-in and Check-out (handle newline between month and day)
        ci = re.search(
            r"Check[- ]in\s+([A-Za-z]{3,9})\s*(\d{1,2}),\s*(\d{4})",
            t, re.I | re.S)
        if ci:
            self.data["Check in"] = parse_month_day_year(ci.group(1), ci.group(2), ci.group(3))
        co = re.search(
            r"Check[- ]out\s+([A-Za-z]{3,9})\s*(\d{1,2}),\s*(\d{4})",
            t, re.I | re.S)
        if co:
            self.data["Check out"] = parse_month_day_year(co.group(1), co.group(2), co.group(3))
        # Room type and occupancy details
        room_match = re.search(
            r"Room Type\s+No\. of Rooms\s+Occupancy.*?\n([^\n]+)",
            t, re.I)
        if room_match:
            line = room_match.group(1).strip()
            tokens = line.split()
            room_tokens = []
            numbers = []
            for tok in tokens:
                if not numbers and not tok.isdigit():
                    # part of the room name
                    room_tokens.append(tok)
                elif tok.isdigit():
                    numbers.append(tok)
            if room_tokens:
                self.data["Room Type Code"] = " ".join(room_tokens)
            if numbers:
                # First number = number of rooms
                self.data["No. of room"] = numbers[0]
                # Second number = occupancy adult if present
                if len(numbers) >= 2:
                    self.data["Occupancy Adult"] = numbers[1]
        # Occupancy Adult (fallback)
        if "Occupancy Adult" not in self.data:
            m = re.search(r"(\d+)\s+Adult", t)
            if m:
                self.data["Occupancy Adult"] = m.group(1)
        # Occupancy children
        m = re.search(r"(\d+)\s+Child", t)
        if m:
            self.data["Occupancy Childrent"] = m.group(1)
        else:
            # Default child occupancy to 0 if not specified
            self.data.setdefault("Occupancy Childrent", "0")
        # Default occupancy adult to 1 if not found
        self.data.setdefault("Occupancy Adult", "1")
        # If number of rooms was not found, default to 1
        self.data.setdefault("No. of room", "1")
        # Special request (Agoda emails seldom include; leave empty)
        self.data.setdefault("Special Request", "")
        # Daily rate: first daily rate after 'From - To Rates'
        dr = find_agoda_first_daily_rate(t)
        if dr:
            self.data["Daily Rate"] = dr
        # Total Booking: prefer gross (reference sell rate), fall back to net rate
        total_gross = find_agoda_total_gross(t)
        if total_gross:
            self.data["Total Booking"] = total_gross
        else:
            # Fall back to net rate (look for 'Net rate (incl. taxes & fees)' followed by 'VND' and amount)
            m = re.search(
                r"Net rate\s*\(incl\. taxes & fees\)\s*\n\s*VND\s*\s*(\d{1,3}(?:,\d{3})+)(?:\.\d+)?",
                t, re.I)
            if m:
                self.data["Total Booking"] = f"{m.group(1)} VND"
        # Amount to Charge Expedia: not applicable for Agoda (leave blank)
        self.data["Amount to Charge Expedia"] = ""
        # Billing details: Agoda hides UPC; return empty fields
        billing: Dict[str, str] = {}
        self.data["Billing Details:"] = billing
        return ordered_output(self.data)


def save_data_as_txt(data_dict: OrderedDict, filepath: str) -> None:
    """Write ordered booking data to a plain text file.

    The fields appear one per line in the order defined by FIELD_ORDER.  The
    billing details are grouped under a header.  This function does not
    return anything; it simply writes the file.
    """
    with open(filepath, "w", encoding="utf-8") as f:
        for key, value in data_dict.items():
            if key != "Billing Details:":
                f.write(f"{key}: {value}\n")
        f.write("\n--- Billing Details: ---\n")
        for subkey, subval in data_dict["Billing Details:"].items():
            f.write(f"  {subkey}: {subval}\n")


def build_html(data: OrderedDict) -> str:
    """Render booking data as a simple HTML string.

    This implementation does not rely on external template files.  It
    produces a minimal, clean layout using a table.  The status is
    displayed as a badge next to the title.  The document is encoded as
    UTF-8 and includes basic styling for readability.
    """
    status = data.get("Status booking Reservation", "")
    rows = []
    for k, v in data.items():
        if k == "Billing Details:":
            continue
        rows.append(f"<tr><td>{k}</td><td>{v}</td></tr>")
    bill_rows = []
    for subkey in BILLING_ORDER:
        bill_rows.append(f"<tr><td>{subkey}</td><td>{data['Billing Details:'].get(subkey, '')}</td></tr>")
    html = f"""<!doctype html>
<html lang="en"><meta charset="utf-8"><title>{data.get('BookingID','')}-report</title>
<style>
body{{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;line-height:1.45}}
h1{{font-size:20px;margin:0 0 8px}}
.badge{{display:inline-block;padding:2px 8px;border-radius:999px;background:#fee2e2;color:#991b1b;
       font-weight:700;font-size:12px;margin-left:6px}}
table{{border-collapse:collapse;min-width:720px;max-width:980px;box-shadow:0 2px 8px rgba(0,0,0,.06)}}
td,th{{border:1px solid #e5e7eb;padding:8px 10px;vertical-align:top}}
td:first-child{{background:#f9fafb;font-weight:600;width:260px}}
tfoot td{{border:none;color:#6b7280;padding-top:10px}}
</style>
<h1>Normalized Booking <span class="badge">{status}</span></h1>
<table><tbody>
{''.join(rows)}
<tr><th colspan="2" style="text-align:left">Billing Details</th></tr>
{''.join(bill_rows)}
</tbody>
<tfoot><tr><td colspan="2">Source: OTA confirmation email PDF.</td></tr></tfoot>
</table></html>"""
    return html


def process_pdf(filepath: str) -> Tuple[str, str]:
    """Parse a PDF file and generate a text and HTML report.

    Args:
        filepath: The path to the PDF file to be processed.

    Returns:
        A tuple ``(txt_path, html_path)`` with the absolute paths of the
        generated text and HTML files.  The output files are saved in
        the same directory as the input PDF.

    Raises:
        ValueError: If the PDF cannot be identified as either an Expedia or
            Agoda email.
    """
    # Check that pdfplumber is available
    if pdfplumber is None:
        raise ImportError(
            "pdfplumber is not installed. Please install the pdfplumber package to use process_pdf()."
        )
    # Extract text from all pages of the PDF
    with pdfplumber.open(filepath) as pdf:
        full_text = "\n".join([(page.extract_text() or "") for page in pdf.pages])
    text_norm = norm_text(full_text)
    low = text_norm.lower()
    # Determine source based on keywords
    if "expedia" in low or "expediapartnercentral" in low:
        parser = ExpediaParser(text_norm)
    elif "agoda" in low and "booking id" in low:
        parser = AgodaParser(text_norm)
    else:
        raise ValueError("Cannot identify OTA source (supported: Expedia, Agoda).")
    data = parser.parse()
    # Determine output file names
    base = os.path.basename(filepath).rsplit('.', 1)[0]
    out_dir = os.path.dirname(filepath) or "."
    txt_path = os.path.join(out_dir, f"{base}_extracted.txt")
    html_path = os.path.join(out_dir, f"{base}_report.html")
    # Write files
    save_data_as_txt(data, txt_path)
    html_content = build_html(data)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return txt_path, html_path