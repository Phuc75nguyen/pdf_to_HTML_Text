from collections import OrderedDict
import re, calendar
import pdfplumber
import os
from datetime import datetime

# Thứ tự khóa EXACT như JSON mẫu
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
    "Billing Details:",  # nested dict giữ thứ tự con riêng
]

BILLING_ORDER = ["Card Number", "Activation Date", "Expiration Date", "Validation Code"]

def ordered_output(data: dict) -> OrderedDict:
    """Ép dữ liệu vào đúng thứ tự & đủ khoá, điền rỗng cho phần thiếu."""
    out = OrderedDict()
    for k in FIELD_ORDER:
        if k == "Billing Details:":
            sub = data.get(k, {}) or {}
            bill = OrderedDict()
            for b in BILLING_ORDER:
                bill[b] = sub.get(b, "")
            out[k] = bill
        else:
            out[k] = data.get(k, "" if k not in ("Has Prepaid",) else False)
    return out



#chuẩn hóa ngày (các helper function)
def norm_text(s: str) -> str:
    s = (s or "").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def mdy_to_mmddyyyy(s: str) -> str:
    try:
        return datetime.strptime(s.strip(), "%b %d, %Y").strftime("%m/%d/%Y")
    except Exception:
        return s

def booked_on_to_mmddyyyy(s: str) -> str:
    m = re.search(r"([A-Za-z]{3}\s+\d{1,2},\s+\d{4})", s)
    return mdy_to_mmddyyyy(m.group(1)) if m else s

def month_year_to_lastday(s: str) -> str:
    m = re.search(r"([A-Za-z]{3})\s+(\d{4})", s)
    if not m:
        return s
    dt = datetime.strptime(f"01 {m.group(1)} {m.group(2)}", "%d %b %Y")
    last = calendar.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last).strftime("%m/%d/%Y")

#regex patterns
def extract_guest_name(t: str) -> tuple[str, str]:
    # Cho phép Unicode tên người (có dấu, dấu gạch, dấu nháy)
    pat = (
        r"Guest:\s*"
        r"(?P<name>[\w .\-'À-ỹ]+?)"
        r"(?=\s+(Booked on:|Reservation ID:|Guest Email:|\n))"
    )
    m = re.search(pat, t, flags=re.I)
    if not m:
        return "", ""
    fullname = m.group("name").strip()
    parts = fullname.split()
    if len(parts) >= 2:
        first = " ".join(parts[:-1])
        last = parts[-1]
    else:
        first, last = fullname, ""
    return first, last
# parser cho file Expedia PDF
class ExpediaParser:
    def __init__(self, text):
        self.t = norm_text(text)
        self.data = {}

    def parse(self):
        t = self.t

        # 1) Status
        self.data["Status booking Reservation"] = (
            "Cancelled" if re.search(r"\b(Cancellation|Cancelled on)\b", t, re.I) else "Confirmed"
        )

        # 2) Email
        m = re.search(r"Guest Email:\s*([^\s]+@[^\s]+)", t, re.I)
        if m: self.data["Email Customer"] = m.group(1)

        # 3) Booking ID
        m = re.search(r"Reservation ID:\s*(\d+)", t, re.I)
        if m: self.data["BookingID"] = m.group(1)

        # 4) Has Prepaid
        self.data["Has Prepaid"] = bool(re.search(r"Guest has PRE-PAID", t, re.I))

        # 5) Booked on (MM/DD/YYYY)
        m = re.search(r"Booked on:\s*(.+?)\n", t, re.I)
        if m: self.data["Booked on"] = booked_on_to_mmddyyyy(m.group(1))

        # 6) Room Type Code
        m = re.search(r"Room Type Code:\s*(.+?)\n", t)
        if not m:
            m = re.search(r"Room Type Name:\s*(.+?)(?:\s*-\s*Non-refundable)?\s*\n", t)
        if m: self.data["Room Type Code"] = m.group(1).strip()

        #7: giá
        # ===== 7) Giá =====
        CURRENCY = r"(VND|USD|EUR|JPY|THB|SGD|AUD|GBP|KRW|CNY)"

        # Daily Rate
        m = re.search(r"Daily Base Rate.*?- *([\d,.]+)\s*" + CURRENCY, t, re.I | re.S)
        if m:
            self.data["Daily Rate"] = f"{m.group(1)} {m.group(2)}"

        # Total Booking (Amount/Price đều hỗ trợ)
        m = re.search(r"(Total Booking Amount|Total Booking Price)\s*:?\s*([\d,.]+)\s*" + CURRENCY, t, re.I | re.S)
        if m:
            self.data["Total Booking"] = f"{m.group(2)} {m.group(3)}"

        # Amount to Charge Expedia — số tiền có thể nằm cùng dòng, currency ở vài dòng sau
        amount = None

# 1) Thử các pattern khi số + currency ở gần nhau (nếu may mắn PDF không bẻ dòng “Group”)
        patterns = [
        r"Amount to Charge Expedia(?:\s*Group)?\s*:\s*([\d,.]+)\s*" + CURRENCY,
        r"Amount to Charge Expedia(?:\s*Group)?\s*:\s*[\r\n ]+([\d,.]+)\s*" + CURRENCY,
        r"Amount to Charge Expedia(?:\s*Group)?\s+([\d,.]+)\s*" + CURRENCY,
]
        for pat in patterns:
            m = re.search(pat, t, re.I | re.S)
            if m:
                amount = f"{m.group(1)} {m.group(2)}"
                break

# 2) Fallback chắc chắn: LẤY CHỈ SỐ TIỀN ngay sau anchor, bỏ qua currency
        if amount is None:
            anchor = re.search(r"Amount to Charge Expedia", t, re.I)
        if anchor:
            tail = t[anchor.end(): anchor.end() + 200]  # quét phần sau nhãn
        # Bắt số đầu tiên dạng 4,604,688 (hoặc 4604688) trong cửa sổ này
        m = re.search(r"(?<!\d)(\d{1,3}(?:[.,]\d{3})+|\d+)(?![\d,])", tail)
        if m:
            amount = m.group(1)  # chỉ số tiền, không kèm VND

        if amount:
            self.data["Amount to Charge Expedia"] = amount



        # 8) Bảng ngày & số khách
        tbl = re.search(
            r"Check-In\s+Check-Out\s+Adults\s+Kids/Ages.*?\n"
            r"([A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+"
            r"([A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+"
            r"(\d+)\s+(\d+)",
            t, re.I | re.S
        )
        if tbl:
            self.data["Check in"] = mdy_to_mmddyyyy(tbl.group(1))
            self.data["Check out"] = mdy_to_mmddyyyy(tbl.group(2))
            self.data["Occupancy Adult"] = tbl.group(3)
            self.data["Occupancy Childrent"] = tbl.group(4)

        # 9) Special Request & No. of room
        self.data["Special Request"] = self.data.get("Special Request", "")
        self.data["No. of room"] = self.data.get("No. of room", "1")

        # 10) Guest name
        first, last = extract_guest_name(t)
        self.data["Customer First Name"] = first
        self.data["Customer Last Name"] = last

        # 11) Billing Details
        bill = {}
        m = re.search(r"Card Number\s*([\d-]+)", t, re.I)
        if m: bill["Card Number"] = m.group(1)
        m = re.search(r"Activation Date\s*([A-Za-z]{3}\s+\d{1,2},\s+\d{4})", t, re.I)
        if m: bill["Activation Date"] = mdy_to_mmddyyyy(m.group(1))
        m = re.search(r"Expiration Date\s*(.+?)\n", t, re.I)
        if m: bill["Expiration Date"] = month_year_to_lastday(m.group(1).strip())
        m = re.search(r"Validation Code\s*(\d+)", t, re.I)
        if m: bill["Validation Code"] = m.group(1)
        self.data["Billing Details:"] = bill

        # 12) Ép về đúng thứ tự & đủ khoá
        return ordered_output(self.data)

#save data
def save_data_as_txt(data_dict, filepath):
    from collections import OrderedDict
    data = ordered_output(dict(data_dict))  # đảm bảo thứ tự

    with open(filepath, "w", encoding="utf-8") as f:
        for k, v in data.items():
            if k != "Billing Details:":
                f.write(f"{k}: {v}\n")
        f.write("\n--- Billing Details: ---\n")
        for kk, vv in data["Billing Details:"].items():
            f.write(f"  {kk}: {vv}\n")

#hàm để process PDF
# --- THÊM CÁC IMPORT NÀY Ở ĐẦU FILE (nếu chưa có) ---
import os
import pdfplumber

# --- THÊM CUỐI FILE: HTML builder + process_pdf ---

def build_html(data: dict) -> str:
    """Render HTML tối giản theo đúng thứ tự FIELD_ORDER/BILLING_ORDER."""
    # đảm bảo đúng thứ tự & đủ khóa
    ordered = ordered_output(dict(data))

    # phần body (trừ Billing Details)
    rows = []
    for k, v in ordered.items():
        if k == "Billing Details:":
            continue
        rows.append(f"<tr><td>{k}</td><td>{v}</td></tr>")

    # billing
    bill = ordered.get("Billing Details:", {})
    bill_rows = []
    for kk in BILLING_ORDER:
        bill_rows.append(f"<tr><td>{kk}</td><td>{bill.get(kk,'')}</td></tr>")

    status = ordered.get("Status booking Reservation", "")
    html = f"""<!doctype html>
<html lang="en"><meta charset="utf-8"><title>{ordered.get('BookingID','')}-report</title>
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
<h1>Expedia – Normalized Booking <span class="badge">{status}</span></h1>
<table><tbody>
{''.join(rows)}
<tr><th colspan="2" style="text-align:left">Billing Details</th></tr>
{''.join(bill_rows)}
</tbody>
<tfoot><tr><td colspan="2">Source: Expedia cancellation email PDF.</td></tr></tfoot>
</table></html>"""
    return html

def process_pdf(filepath: str):
    """
    Mở PDF, trích xuất text, chọn parser phù hợp,
    ghi TXT/HTML theo đúng schema & thứ tự, trả về (txt_path, html_path).
    """
    with pdfplumber.open(filepath) as pdf:
        full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

    text_norm = norm_text(full_text)

    # Detect nguồn (tạm thời chỉ Expedia; nếu cần sẽ bổ sung Agoda)
    if ("expedia" in text_norm.lower()) or ("expediapartnercentral" in text_norm.lower()):
        parser = ExpediaParser(text_norm)
    else:
        # Nếu chưa hỗ trợ định dạng khác, raise để biết lý do
        raise ValueError("Không nhận diện được định dạng PDF (chỉ hỗ trợ Expedia hiện tại).")

    data = parser.parse()  # đã ordered_output trong ExpediaParser.parse()

    # Lưu cạnh file gốc
    base = os.path.basename(filepath).rsplit(".", 1)[0]
    out_dir = os.path.dirname(filepath) or "."
    txt_path = os.path.join(out_dir, f"{base}_extracted.txt")
    html_path = os.path.join(out_dir, f"{base}_report.html")

    # Ghi TXT theo thứ tự cố định
    save_data_as_txt(data, txt_path)

    # Ghi HTML inline (không cần template ngoài)
    html = build_html(data)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    return txt_path, html_path
