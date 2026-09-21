import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import khối xử lý lõi (Core)
from core.rule_engine import RuleEngine

class AppWindow:
    def __init__(self, root):
        self.root = root
        
        # Các biến trạng thái của giao diện
        self.file_path_var = tk.StringVar()
        self.add_note_var = tk.BooleanVar(value=True) # Mặc định bật tính năng chèn ghi chú (Inline Note)
        
        # Khởi tạo các thành phần giao diện
        self.create_widgets()

    def create_widgets(self):
        # ==========================================
        # KHUNG TRÊN: CHỌN FILE
        # ==========================================
        frame_top = tk.Frame(self.root, pady=10, padx=20)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Chọn file công văn (.docx):", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        
        entry_path = tk.Entry(frame_top, textvariable=self.file_path_var, width=60, state='readonly')
        entry_path.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        btn_browse = tk.Button(frame_top, text="Duyệt...", command=self.browse_file, width=10)
        btn_browse.pack(side=tk.LEFT)

        # ==========================================
        # KHUNG GIỮA: BẢNG TRACKING LỖI (TREEVIEW)
        # ==========================================
        frame_middle = tk.Frame(self.root, pady=10, padx=20)
        frame_middle.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame_middle, text="Chi tiết các lỗi phát hiện:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        # Khởi tạo cột cho Treeview
        columns = ("vị_trí", "nội_dung", "lỗi_nghị_định")
        self.tree = ttk.Treeview(frame_middle, columns=columns, show="headings", height=15)
        
        # Cấu hình tiêu đề và độ rộng các cột
        self.tree.heading("vị_trí", text="Vị trí (Đoạn)")
        self.tree.column("vị_trí", width=100, anchor=tk.CENTER)
        
        self.tree.heading("nội_dung", text="Trích xuất nội dung sai")
        self.tree.column("nội_dung", width=300, anchor=tk.W)
        
        self.tree.heading("lỗi_nghị_định", text="Lý do / Căn cứ Nghị định")
        self.tree.column("lỗi_nghị_định", width=400, anchor=tk.W)

        # Tích hợp thanh cuộn dọc (Scrollbar)
        scrollbar = ttk.Scrollbar(frame_middle, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ==========================================
        # KHUNG DƯỚI: ĐIỀU KHIỂN & CẤU HÌNH
        # ==========================================
        frame_bottom = tk.Frame(self.root, pady=15, padx=20)
        frame_bottom.pack(fill=tk.X)

        # Checkbox bật/tắt ghi chú vào file
        chk_note = tk.Checkbutton(frame_bottom, text="Đính kèm ghi chú lỗi (Inline Note) vào file xuất", 
                                  variable=self.add_note_var, font=("Arial", 10))
        chk_note.pack(side=tk.LEFT)

        # Nút xử lý chính
        btn_process = tk.Button(frame_bottom, text="Kiểm tra & Chuẩn hóa", 
                                command=self.process_document, 
                                font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", width=20, height=2)
        btn_process.pack(side=tk.RIGHT)

    # ==========================================
    # CÁC HÀM XỬ LÝ SỰ KIỆN (EVENT HANDLERS)
    # ==========================================
    def browse_file(self):
        """Mở hộp thoại chọn file Word"""
        file_path = filedialog.askopenfilename(
            title="Chọn file công văn",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def process_document(self):
        """Chạy thuật toán kiểm tra, ghi log lên Treeview và xuất file"""
        target_file = self.file_path_var.get()
        if not target_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một file .docx trước khi kiểm tra!")
            return
            
        # Xóa dữ liệu cũ trên bảng Treeview của lần chạy trước
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            # Khởi tạo lõi xử lý (RuleEngine) và chạy kiểm tra
            engine = RuleEngine(target_file)
            
            # Truyền trạng thái của Checkbox vào hàm xử lý
            errors = engine.check_and_format(add_note=self.add_note_var.get())
            
            # Lưu file đã highlight ra đĩa
            output_path = engine.save_output()
            
            # Render kết quả lên giao diện
            if errors:
                for err in errors:
                    self.tree.insert("", tk.END, values=(err["vi_tri"], err["noi_dung"], err["loi"]))
                messagebox.showwarning("Phát hiện lỗi", f"Đã tìm thấy {len(errors)} lỗi định dạng.\nFile đối chiếu được tạo tại:\n{output_path}")
            else:
                messagebox.showinfo("Hoàn hảo", "Công văn của bạn đã chuẩn thể thức! Không phát hiện lỗi.")
                
            # Tự động mở file Word vừa được chuẩn hóa để người dùng xem luôn
            os.startfile(output_path)
            
        except Exception as e:
            messagebox.showerror("Lỗi hệ thống", f"Có lỗi xảy ra trong quá trình xử lý file:\n{str(e)}")