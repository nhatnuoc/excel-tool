import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button, Label, StringVar, OptionMenu, Entry
import os
import re

def sanitize_filename(filename):
    """Xóa ký tự không hợp lệ khỏi tên file."""
    return re.sub(r'[\\/:"*?<>|]+', "_", str(filename))

def load_excel_columns(file_path, header_row):
    """Đọc danh sách cột từ file Excel, với dòng tiêu đề được chỉ định."""
    try:
        data = pd.read_excel(file_path, header=header_row)
        columns = list(data.columns)
        return data, columns
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")
        return None, None

def split_excel_by_column(file_path, column_name, header_row):
    """Chia tách file Excel theo tên cột."""
    # Chọn thư mục lưu
    output_folder = filedialog.askdirectory(title="Chọn thư mục để lưu các file")
    if not output_folder:
        return
    
    try:
        # Đọc file Excel
        data = pd.read_excel(file_path, header=header_row)
        
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
    file_path = None  # Khởi tạo biến file_path trong phạm vi hàm create_gui

    def on_select_file():
        """Xử lý chọn file và hiển thị danh sách cột."""
        nonlocal file_path  # Biến nonlocal tham chiếu file_path trong hàm cha
        try:
            header_row = int(header_row_var.get()) - 1  # Dòng tiêu đề từ người dùng (giảm 1 vì bắt đầu từ 0)
        except ValueError:
            messagebox.showerror("Lỗi", "Dòng tiêu đề phải là số nguyên.")
            return
        
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        _, columns = load_excel_columns(file_path, header_row)
        if columns:
            column_var.set("")
            column_menu["menu"].delete(0, "end")
            for col in columns:
                column_menu["menu"].add_command(label=col, command=lambda value=col: column_var.set(value))
            label_columns.config(text="\n".join(columns))
    
    def on_split_click():
        """Thực hiện chia tách file theo cột được chọn."""
        if not file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file Excel trước!")
            return
        column_name = column_var.get().strip()
        if not column_name:
            messagebox.showerror("Lỗi", "Vui lòng chọn cột để chia tách!")
            return
        try:
            header_row = int(header_row_var.get()) - 1  # Dòng tiêu đề từ người dùng (giảm 1 vì bắt đầu từ 0)
        except ValueError:
            messagebox.showerror("Lỗi", "Dòng tiêu đề phải là số nguyên.")
            return
        split_excel_by_column(file_path, column_name, header_row)
    
    # Khởi tạo giao diện
    root = Tk()
    root.title("Công cụ chia tách file Excel")
    root.geometry("500x500")

    Label(root, text="Công cụ chia tách file Excel", font=("Arial", 14)).pack(pady=10)
    
    # Nhập dòng tiêu đề
    Label(root, text="Nhập số dòng chứa tiêu đề cột (mặc định là 1):", font=("Arial", 10)).pack(pady=5)
    header_row_var = StringVar(root, value="1")
    Entry(root, textvariable=header_row_var, font=("Arial", 10), width=10).pack(pady=5)

    # Nút chọn file Excel
    Button(root, text="Chọn File Excel", command=on_select_file, font=("Arial", 10)).pack(pady=10)

    # Hiển thị danh sách cột
    Label(root, text="Danh sách các cột trong file:", font=("Arial", 10)).pack(pady=5)
    label_columns = Label(root, text="", font=("Arial", 9), justify="left", wraplength=400)
    label_columns.pack(pady=5)

    # Dropdown chọn cột
    Label(root, text="Chọn cột để chia tách:", font=("Arial", 10)).pack(pady=5)
    column_var = StringVar(root)
    column_menu = OptionMenu(root, column_var, "")
    column_menu.pack(pady=10)

    # Nút thực hiện chia file
    Button(root, text="Thực hiện Chia File", command=on_split_click, font=("Arial", 10)).pack(pady=10)
    Button(root, text="Thoát", command=root.quit, font=("Arial", 10)).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
