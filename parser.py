import pdfplumber
import re  # Thư viện cho Biểu thức Chính quy (Regex)
import json
import os
from jinja2 import Environment, FileSystemLoader

# Thiết lập môi trường Jinja2
# (Giả sử template.html nằm cùng thư mục)
try:
    file_loader = FileSystemLoader('.') 
    env = Environment(loader=file_loader)
    template = env.get_template('template.html')
except Exception:
    print("Lỗi: Không tìm thấy 'template.html'. Hãy chắc chắn file này tồn tại.")
    template = None

def parse_expedia(full_text, pages):
    """
    Hàm này nhận text thô từ file Expedia và trích xuất dữ liệu.
    """
    data = {} # Tạo một dictionary rỗng để chứa kết quả

    # Dùng Regex để "săn" dữ liệu
    # (m) = chế độ đa dòng, (i) = không phân biệt hoa thường
    
    # "Status booking Reservation": "Cancelled"
    if re.search(r"Cancellation", full_text, re.I): 
        data["Status_booking_Reservation"] = "Cancelled"
    else:
        # (Chúng ta có thể thêm logic cho "Confirmed" sau)
        data["Status_booking_Reservation"] = "Confirmed" 

    # "Customer First Name": "Minh Tam", "Customer Last Name": "Diep"
    # Regex: Tìm chữ "Guest:", sau đó lấy 2 "từ" đầu tiên, và "từ" cuối cùng
    match = re.search(r"Guest: (\w+\s+\w+) (\w+)", full_text) 
    if match:
        data["Customer_First_Name"] = match.group(1) # Phần trong ngoặc () đầu tiên
        data["Customer_Last_Name"] = match.group(2)  # Phần trong ngoặc () thứ hai
    else:
        # Thử một mẫu khác nếu mẫu trên thất bại
        match_guest = re.search(r"Guest: (.*)", full_text)
        if match_guest:
             # Nếu chỉ có 1 tên, ta sẽ phải xử lý (ví dụ: "Guest: JohnDoe")
             # Tạm thời để trống nếu không khớp
             full_name = match_guest.group(1).split()
             data["Customer_First_Name"] = full_name[0] if len(full_name) > 0 else ""
             data["Customer_Last_Name"] = " ".join(full_name[1:]) if len(full_name) > 1 else ""
        else:
             data["Customer_First_Name"] = ""
             data["Customer_Last_Name"] = ""


    # "Email Customer": "2p4a290mm8@m.expediapartnercentral.com"
    match = re.search(r"Guest Email: ([\S]+@[\S]+)", full_text) 
    if match:
        data["Email_Customer"] = match.group(1)

    # "BookingID": "2307501514"
    match = re.search(r"Reservation ID: (\d+)", full_text) 
    if match:
        data["BookingID"] = match.group(1)

    # "Has Prepaid": true
    data["Has_Prepaid"] = bool(re.search(r"Guest has PRE-PAID", full_text, re.I)) 

    # "Booked on": "Oct 9, 2025" (Regex sẽ lấy toàn bộ, ta sẽ cần chuẩn hóa sau)
    match = re.search(r"Booked on: (.*) PST", full_text) 
    if match:
        data["Booked_on"] = match.group(1).strip() # "Oct 9, 2025 8:11 AM"

    # "Room Type Code"
    match = re.search(r"Room Type Name: (.*?)( - Non-refundable)?$", full_text, re.M) 
    if match:
        data["Room_Type_Code"] = match.group(1).strip()

    # "Total Booking": "5,755,860 VND"
    match = re.search(r"Total Booking Amount:\s+([\d,.]+) (\w+)", full_text, re.M) 
    if match:
        data["Total_Booking"] = f"{match.group(1)} {match.group(2)}"

    # "Amount to Charge Expedia": "4,604,688 VND"
    match = re.search(r"Amount to Charge Expedia\s+([\d,.]+) (\w+)", full_text, re.M) 
    if match:
        data["Amount_to_Charge_Expedia"] = f"{match.group(1)} {match.group(2)}"

    # --- Xử lý bảng (Cách tốt nhất) ---
    # pdfplumber rất mạnh trong việc đọc bảng
    
    # Trích xuất bảng Check-in/Check-out
    try:
        page_one = pages[0] # Lấy trang 1
        tables = page_one.extract_tables()
        
        # Bảng 1 (Chi tiết check-in)
        checkin_table = tables[0]
        # Dòng 1 là dữ liệu (dòng 0 là tiêu đề)
        data["Check_in"] = checkin_table[1][0].replace('\n', ' ') # "Nov 12, 2025"
        data["Check_out"] = checkin_table[1][1].replace('\n', ' ') # "Nov 14, 2025"
        data["Occupancy_Adult"] = checkin_table[1][2] # "3"
        data["Occupancy_Childrent"] = checkin_table[1][3] # "0"
        
        # Bảng 2 (Billing Details)
        billing_table = tables[1]
        data["Billing_Details"] = {
            "Card_Number": billing_table[1][1],      # "5556-5703-2498-5802"
            "Activation_Date": billing_table[2][1],  # "Nov 11, 2025"
            "Expiration_Date": billing_table[3][1],  # "Sep 2030"
        }
    except Exception as e:
        print(f"Lỗi khi đọc bảng: {e}")
        data["Billing_Details"] = {}

    # "Validation Code" (Không nằm trong bảng)
    match = re.search(r"Validation Code\s+(\d+)", full_text, re.M) 
    if match:
        data.setdefault("Billing_Details", {})["Validation_Code"] = match.group(1)

    # Các trường còn thiếu (bạn có thể bổ sung Regex sau)
    data.setdefault("Special_Request", "") 
    data.setdefault("No_of_room", "1") # Tạm thời
    
    # "Daily Rate"
    match = re.search(r"Daily Base Rate - Package - ([\d,.]+) (\w+)", full_text, re.M) 
    if match:
        data["Daily_Rate"] = f"{match.group(1)} {match.group(2)}"
    else:
        data.setdefault("Daily_Rate", "") 

    return data

def process_pdf(filepath):
    """
    Hàm chính để xử lý một file PDF.
    Nó sẽ "nhận diện" file và gọi hàm parser_... tương ứng.
    """
    try:
        with pdfplumber.open(filepath) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += page.extract_text() + "\n"
            
            # --- Nhận diện loại PDF ---
            data_dict = {}
            if "expedia.com" in full_text or "Expedia Group" in full_text: 
                data_dict = parse_expedia(full_text, pdf.pages)
            elif "agoda.com" in full_text: 
                # (Chúng ta sẽ tạo hàm parse_agoda() sau)
                data_dict = {"Error": "Parser Agoda chưa được xây dựng."}
            else:
                data_dict = {"Error": "Không nhận diện được loại PDF."}

            # --- Lưu file ---
            base_filename = os.path.basename(filepath).replace('.pdf', '')
            output_folder = os.path.dirname(filepath) # Lưu cùng chỗ

            # 1. Lưu file JSON (giống hệt .txt nhưng có cấu trúc)
            json_filename = os.path.join(output_folder, f"{base_filename}_extracted.json")
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(data_dict, f, ensure_ascii=False, indent=4)
                
            # 2. Lưu file HTML
            html_filename = os.path.join(output_folder, f"{base_filename}_report.html")
            if template:
                output_html = template.render(data_dict)
                with open(html_filename, "w", encoding="utf-8") as f:
                    f.write(output_html)
            else:
                html_filename = None # Không tạo được HTML nếu thiếu template

            return json_filename, html_filename

    except Exception as e:
        print(f"Lỗi nghiêm trọng khi xử lý file {filepath}: {e}")
        return None, None