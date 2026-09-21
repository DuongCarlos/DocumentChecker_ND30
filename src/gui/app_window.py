import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core.rule_engine import RuleEngine

class AppWindow:
    def __init__(self, root):
        self.root = root
        self.file_path_var = tk.StringVar()
        self.add_note_var = tk.BooleanVar(value=True) 
        self.create_widgets()

    def create_widgets(self):
        frame_top = tk.Frame(self.root, pady=10, padx=20)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Chọn file công văn (.docx):", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        
        entry_path = tk.Entry(frame_top, textvariable=self.file_path_var, width=60, state='readonly')
        entry_path.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        btn_browse = tk.Button(frame_top, text="Duyệt...", command=self.browse_file, width=10)
        btn_browse.pack(side=tk.LEFT)

        frame_middle = tk.Frame(self.root, pady=10, padx=20)
        frame_middle.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame_middle, text="Chi tiết các lỗi phát hiện:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        # Đã bổ sung thêm cột "cách_sửa" vào cấu trúc bảng
        columns = ("vị_trí", "nội_dung", "lỗi_nghị_định", "cách_sửa")
        self.tree = ttk.Treeview(frame_middle, columns=columns, show="headings", height=15)
        
        self.tree.heading("vị_trí", text="Vị trí (Đoạn)")
        self.tree.column("vị_trí", width=100, anchor=tk.CENTER)
        
        self.tree.heading("nội_dung", text="Trích xuất nội dung sai")
        self.tree.column("nội_dung", width=250, anchor=tk.W)
        
        self.tree.heading("lỗi_nghị_định", text="Lý do / Căn cứ Nghị định")
        self.tree.column("lỗi_nghị_định", width=350, anchor=tk.W)

        self.tree.heading("cách_sửa", text="Hướng dẫn sửa (Tổ hợp phím)")
        self.tree.column("cách_sửa", width=250, anchor=tk.W)

        scrollbar = ttk.Scrollbar(frame_middle, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        frame_bottom = tk.Frame(self.root, pady=15, padx=20)
        frame_bottom.pack(fill=tk.X)

        chk_note = tk.Checkbutton(frame_bottom, text="Đính kèm ghi chú lỗi (Inline Note) vào file xuất", 
                                  variable=self.add_note_var, font=("Arial", 10))
        chk_note.pack(side=tk.LEFT)

        btn_process = tk.Button(frame_bottom, text="Kiểm tra & Chuẩn hóa", 
                                command=self.process_document, 
                                font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", width=20, height=2)
        btn_process.pack(side=tk.RIGHT)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file công văn",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def process_document(self):
        target_file = self.file_path_var.get()
        if not target_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một file .docx trước khi kiểm tra!")
            return
            
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            engine = RuleEngine(target_file)
            errors = engine.check_and_format(add_note=self.add_note_var.get())
            output_path = engine.save_output()
            
            if errors:
                for err in errors:
                    # Trích xuất dữ liệu cách sửa để đẩy vào cột thứ 4
                    fix = err.get("cach_sua", "")
                    self.tree.insert("", tk.END, values=(err["vi_tri"], err["noi_dung"], err["loi"], fix))
                messagebox.showwarning("Phát hiện lỗi", f"Đã tìm thấy {len(errors)} lỗi định dạng.\nFile đối chiếu được tạo tại:\n{output_path}")
            else:
                messagebox.showinfo("Hoàn hảo", "Công văn của bạn đã chuẩn thể thức! Không phát hiện lỗi.")
                
            os.startfile(output_path)
            
        except Exception as e:
            messagebox.showerror("Lỗi hệ thống", f"Có lỗi xảy ra trong quá trình xử lý file:\n{str(e)}")