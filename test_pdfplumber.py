import pdfplumber
import os

pdf_filename = 'Inbox - ota-booking - Outlook.pdf'

if not os.path.exists(pdf_filename):
    print(f"Lỗi: Không tìm thấy file '{pdf_filename}'.")
else:
    print(f"--- Đang mở file: {pdf_filename} ---")
    
    # Tạo một biến để "gom" text của tất cả các trang
    full_text = "" 
    
    with pdfplumber.open(pdf_filename) as pdf:
        
        # Dùng vòng lặp for để duyệt qua từng trang
        # enumerate giúp chúng ta lấy cả số thứ tự (i) và đối tượng trang (page)
        for i, page in enumerate(pdf.pages):
            
            print(f"--- Đang đọc trang {i + 1} ---")
            
            # Trích xuất văn bản từ trang hiện tại
            text_content = page.extract_text()
            
            # Cộng dồn văn bản vào biến full_text
            if text_content: # Chỉ cộng nếu trang có text
                full_text += text_content + "\n" # Thêm ký tự xuống dòng để tách biệt các trang

    print("\n--- NỘI DUNG ĐẦY ĐỦ CỦA FILE PDF ---")
    print(full_text)
    print("-----------------------------------")