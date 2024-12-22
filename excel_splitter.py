import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button, Label
import os

def split_excel_by_column():
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
        
        # Xác định cột để chia
        column_name = "ĐV phát hành"
        if column_name not in data.columns:
            messagebox.showerror("Lỗi", f"Cột '{column_name}' không tồn tại trong file Excel.")
            return
        
        # Lấy các giá trị duy nhất trong cột
        unique_values = data[column_name].dropna().unique()
        
        # Chia file và lưu từng file
        for value in unique_values:
            filtered_data = data[data[column_name] == value]
            file_name = f"Danh sách thẻ hết hạn - {value}.xlsx"
            save_path = os.path.join(output_folder, file_name)
            filtered_data.to_excel(save_path, index=False)
        
        messagebox.showinfo("Thành công", f"Đã chia tách file thành công và lưu vào: {output_folder}")
    
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

# Tạo giao diện bằng Tkinter
def create_gui():
    root = Tk()
    root.title("Công cụ chia tách file Excel")
    root.geometry("400x200")

    Label(root, text="Chia tách file Excel theo cột 'ĐV phát hành'", font=("Arial", 12)).pack(pady=20)
    
    Button(root, text="Chọn và Chia File Excel", command=split_excel_by_column, font=("Arial", 10)).pack(pady=10)
    Button(root, text="Thoát", command=root.quit, font=("Arial", 10)).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
