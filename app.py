import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import queue
import time  # Chúng ta dùng time.sleep() để "giả lập" việc xử lý PDF
import parser 

# Định nghĩa lớp ứng dụng chính
class PdfExtractorApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Extractor (Dự án PDF to HTML/Text)")
        self.root.geometry("700x550") # Kích thước cửa sổ
        
        # Danh sách để lưu các đường dẫn file đã chọn
        self.file_list = []
        
        # Queue để giao tiếp giữa luồng xử lý và giao diện
        self.log_queue = queue.Queue()
        
        # Tạo các thành phần giao diện
        self.create_widgets()
        
        # Khởi động bộ đếm thời gian để kiểm tra queue
        self.check_log_queue()

    def create_widgets(self):
        # --- Phần 1: Tải file lên ---
        frame_select = ttk.Frame(self.root, padding="10")
        frame_select.pack(fill='x', expand=False)

        btn_select_files = ttk.Button(frame_select, text="[ 1. Chọn File PDF ]", command=self.select_files)
        btn_select_files.pack(side='left', padx=5, pady=5)
        
        # (Tùy chọn, bạn có thể thêm nút "Chọn Thư mục" sau)
        # btn_select_folder = ttk.Button(frame_select, text="Chọn Thư mục", command=self.select_folder)
        # btn_select_folder.pack(side='left', padx=5, pady=5)

        # Listbox để hiển thị file đã chọn
        frame_list = ttk.Frame(self.root, padding="10")
        frame_list.pack(fill='both', expand=True)
        
        self.listbox = tk.Listbox(frame_list, height=10, selectmode=tk.EXTENDED)
        self.listbox.pack(side='left', fill='both', expand=True)
        
        # Thêm thanh cuộn cho Listbox
        scrollbar = ttk.Scrollbar(frame_list, orient='vertical', command=self.listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox.config(yscrollcommand=scrollbar.set)

        # --- Phần 2: Xử lý ---
        frame_controls = ttk.Frame(self.root, padding="10")
        frame_controls.pack(fill='x', expand=False)

        self.btn_start = ttk.Button(frame_controls, text="[ 2. Bắt đầu Chuyển đổi ]", command=self.start_conversion)
        self.btn_start.pack(side='left', padx=5, pady=5)
        
        btn_clear = ttk.Button(frame_controls, text="Xóa Danh sách", command=self.clear_list)
        btn_clear.pack(side='right', padx=5, pady=5)

        # --- Phần 3: Trả về kết quả (Log) ---
        frame_log = ttk.Frame(self.root, padding="10")
        frame_log.pack(fill='both', expand=True)

        ttk.Label(frame_log, text="[ 3. Kết quả (Log): ]").pack(anchor='w')
        
        self.log_text = tk.Text(frame_log, height=10, state='disabled', wrap=tk.WORD)
        self.log_text.pack(side='left', fill='both', expand=True)
        
        # Thêm thanh cuộn cho Log
        log_scrollbar = ttk.Scrollbar(frame_log, orient='vertical', command=self.log_text.yview)
        log_scrollbar.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=log_scrollbar.set)

    # -------------------------------------
    # Các hàm chức năng
    # -------------------------------------

    def select_files(self):
        # Mở hộp thoại chọn file
        # 'askopenfilenames' cho phép chọn nhiều file
        filenames = filedialog.askopenfilenames(
            title="Chọn file PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filenames:
            for f in filenames:
                if f not in self.file_list:
                    self.file_list.append(f)
                    self.listbox.insert(tk.END, os.path.basename(f)) # Chỉ hiển thị tên file
            self.log_message(f"Đã thêm {len(filenames)} file vào hàng đợi.")

    def clear_list(self):
        self.file_list.clear()
        self.listbox.delete(0, tk.END)
        self.log_message("Đã xóa sạch danh sách.")

    def log_message(self, message):
        # Đây là hàm an toàn để *GỬI* log từ bất kỳ luồng nào
        # Nó đưa tin nhắn vào queue
        current_time = time.strftime("%H:%M:%S")
        self.log_queue.put(f"[{current_time}] {message}\n")

    def check_log_queue(self):
        # Hàm này chạy trên GIAO DIỆN CHÍNH (main thread)
        # Nó kiểm tra queue mỗi 100ms
        try:
            while True:
                # Lấy tin nhắn từ queue
                message = self.log_queue.get_nowait()
                
                # Hiển thị tin nhắn lên Text box
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, message)
                self.log_text.see(tk.END) # Tự động cuộn xuống
                self.log_text.config(state='disabled')
                
        except queue.Empty:
            # Nếu queue rỗng thì không làm gì cả
            pass
        finally:
            # Lên lịch để tự gọi lại chính nó sau 100ms
            self.root.after(100, self.check_log_queue)

    def start_conversion(self):
        if not self.file_list:
            messagebox.showwarning("Chưa có file", "Vui lòng chọn ít nhất một file PDF để xử lý.")
            return

        # Vô hiệu hóa nút "Start" để tránh bấm 2 lần
        self.btn_start.config(state='disabled')
        self.log_message(f"Bắt đầu xử lý {len(self.file_list)} file...")

        # *** Đây là phần quan trọng ***
        # Chúng ta chạy hàm xử lý logic trên một luồng (thread) riêng
        # để không làm "đơ" giao diện
        self.thread = threading.Thread(target=self.process_files_thread, daemon=True)
        self.thread.start()

    def process_files_thread(self):
        # *** HÀM NÀY CHẠY TRÊN LUỒNG NGẦM (BACKGROUND THREAD) ***
        try:
            total_files = len(self.file_list)
            for i, filepath in enumerate(self.file_list):
                filename = os.path.basename(filepath)
                self.log_message(f"({i+1}/{total_files}) Đang xử lý: {filename}...")
                
                # --- PHẦN LÕI LOGIC THẬT ---
                # Gọi "bộ não" parser.py của chúng ta
                # Thay vì time.sleep(1)
                
                json_file, html_file = parser.process_pdf(filepath)
                
                # -----------------------------
                
                if json_file:
                    self.log_message(f"    -> XUẤT FILE: {os.path.basename(json_file)}")
                if html_file:
                    self.log_message(f"    -> XUẤT FILE: {os.path.basename(html_file)}")
                if not json_file and not html_file:
                    self.log_message(f"    -> LỖI: Không thể xử lý {filename}")

            self.log_message(f"=== ĐÃ XỬ LÝ XONG {total_files}/{total_files} file ===")

        except Exception as e:
            # Gửi lỗi về log chính
            import traceback
            error_details = traceback.format_exc()
            self.log_message(f"LỖI NGHIÊM TRỌNG TRONG LUỒNG: {e}\n{error_details}")
            # Vẫn nên hiện messagebox vì đây là lỗi nghiêm trọng
            # Nhưng phải gọi nó một cách an toàn
            self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Đã xảy ra lỗi nghiêm trọng: {e}"))
        
        finally:
            # Gửi tin nhắn cho GUI để kích hoạt lại nút "Start"
            self.log_queue.put("ENABLE_START_BUTTON")

    def check_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                
                # Xử lý các "tin nhắn đặc biệt"
                if message == "ENABLE_START_BUTTON":
                    self.btn_start.config(state='normal')
                else:
                    # Hiển thị tin nhắn log bình thường
                    self.log_text.config(state='normal')
                    self.log_text.insert(tk.END, message)
                    self.log_text.see(tk.END)
                    self.log_text.config(state='disabled')
                
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_log_queue)

# -------------------------------------
# Khởi chạy ứng dụng
# -------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = PdfExtractorApp(root)
    root.mainloop()