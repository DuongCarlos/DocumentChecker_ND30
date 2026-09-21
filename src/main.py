import tkinter as tk
from gui.app_window import AppWindow

def main():
    # Khởi tạo cửa sổ chính
    root = tk.Tk()
    root.title("Công cụ Chuẩn hóa Công văn - Nghị định 30/2020/NĐ-CP")
    
    # Thiết lập kích thước cửa sổ và vị trí xuất hiện ở giữa màn hình
    window_width = 900
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x_cordinate = int((screen_width/2) - (window_width/2))
    y_cordinate = int((screen_height/2) - (window_height/2))
    root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
    
    # Chặn thay đổi kích thước quá nhỏ
    root.minsize(800, 500)

    # Khởi tạo Giao diện (AppWindow)
    app = AppWindow(root)

    # Chạy vòng lặp sự kiện
    root.mainloop()

if __name__ == "__main__":
    main()