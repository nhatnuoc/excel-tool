import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button, Label, Entry
import os
import re

def sanitize_filename(filename):
    """Xóa ký tự không hợp lệ khỏi tên file."""
    return re.sub(r'[\\/:"*?<>|]+', "_", str(filename))

def split_excel_by_column(column_name):
    """Hàm chia file Excel theo tên cột."""
    # Chọn tệp Excel
    file_path = filedialog.askopenfilename(
        title="Chọn file Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return
    
    # Chọn thư mục lưu
    output_folder = filedialog.askdirectory(title="Chọn thư mục để lưu các file")
    if not output_folder:
        return
    
    try:
        # Đọc file Excel
        data = pd.read_excel(file_path)
        
        # Kiểm tra cột tồn tại
        if column_name not in data.columns:
            messagebox.showerror("Lỗi", f"Cột '{column_name}' không tồn tại trong file Excel.")
            return
        
        # Lấy các giá trị duy nhất trong cột
        unique_values = data[column_name].dropna().unique()
        
        # Chia file và lưu từng file
        for value in unique_values:
            filtered_data = data[data[column_name] == value]
            sanitized_value = sanitize_filename(value)
            file_name = f"Danh sách thẻ hết hạn - {sanitized_value}.xlsx"
            save_path = os.path.join(output_folder, file_name)
            filtered_data.to_excel(save_path, index=False)
        
        messagebox.showinfo("Thành công", f"Đã chia tách file thành công và lưu vào: {output_folder}")
    
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

# Tạo giao diện bằng Tkinter
def create_gui():
    """Tạo giao diện chính."""
    def on_split_click():
        column_name = column_entry.get().strip()
        if not column_name:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên cột để chia tách!")
            return
        split_excel_by_column(column_name)
    
    root = Tk()
    root.title("Công cụ chia tách file Excel")
    root.geometry("400x250")

    Label(root, text="Công cụ chia tách file Excel", font=("Arial", 14)).pack(pady=10)
    
    Label(root, text="Nhập tên cột để chia tách:", font=("Arial", 10)).pack(pady=5)
    column_entry = Entry(root, width=30, font=("Arial", 10))
    column_entry.pack(pady=5)

    Button(root, text="Chọn và Chia File Excel", command=on_split_click, font=("Arial", 10)).pack(pady=10)
    Button(root, text="Thoát", command=root.quit, font=("Arial", 10)).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
