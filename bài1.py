import tkinter as tk
from tkinter import ttk, messagebox
import csv
import pandas as pd
from datetime import datetime


# Hàm ghi dữ liệu vào file CSV
def save_to_csv():
    data = {
        "Mã": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Đơn vị": entry_donvi.get(),
        "Chức danh": entry_chucdanh.get(),
        "Ngày sinh": entry_ngaysinh.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_cmnd.get(),
        "Ngày cấp": entry_ngaycap.get(),
        "Nơi cấp": entry_noicap.get(),
    }
    # Ghi dữ liệu vào file CSV
    with open("employee_data.csv", mode="a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=data.keys())
        if file.tell() == 0:  # Ghi header nếu file trống
            writer.writeheader()
        writer.writerow(data)
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu!")
    clear_entries()

# Hàm xóa các ô nhập liệu
def clear_entries():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_donvi.delete(0, tk.END)
    entry_chucdanh.delete(0, tk.END)
    entry_ngaysinh.delete(0, tk.END)
    entry_cmnd.delete(0, tk.END)
    entry_ngaycap.delete(0, tk.END)
    entry_noicap.delete(0, tk.END)
    gender_var.set("")

# Hàm lọc nhân viên có sinh nhật hôm nay
def filter_birthday_today():
    today = datetime.now().strftime("%d/%m")
    try:
        df = pd.read_csv("employee_data.csv", encoding="utf-8")
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y', errors='coerce')
        today_birthdays = df[df['Ngày sinh'].dt.strftime('%d/%m') == today]
        if not today_birthdays.empty:
            messagebox.showinfo("Sinh nhật hôm nay", today_birthdays.to_string(index=False))
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên!")

# Hàm xuất danh sách ra file Excel và sắp xếp theo tuổi giảm dần
def export_to_excel():
    try:
        df = pd.read_csv("employee_data.csv", encoding="utf-8")
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y', errors='coerce')
        df = df.sort_values(by='Ngày sinh', ascending=True)  # Tuổi giảm dần
        df.to_excel("employee_list.xlsx", index=False, engine='openpyxl')
        messagebox.showinfo("Thành công", "Danh sách đã được xuất ra file employee_list.xlsx!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Không có dữ liệu để xuất!")

# Tạo giao diện Tkinter
root = tk.Tk()
root.title("Quản lý nhân viên")

# Các biến lưu trữ dữ liệu
gender_var = tk.StringVar()

# Tạo các label và entry
tk.Label(root, text="Mã:").grid(row=0, column=0)
entry_ma = tk.Entry(root)
entry_ma.grid(row=0, column=1)

tk.Label(root, text="Tên:").grid(row=0, column=2)
entry_ten = tk.Entry(root)
entry_ten.grid(row=0, column=3)

tk.Label(root, text="Đơn vị:").grid(row=1, column=0)
entry_donvi = tk.Entry(root)
entry_donvi.grid(row=1, column=1)

tk.Label(root, text="Chức danh:").grid(row=1, column=2)
entry_chucdanh = tk.Entry(root)
entry_chucdanh.grid(row=1, column=3)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY):").grid(row=2, column=0)
entry_ngaysinh = tk.Entry(root)
entry_ngaysinh.grid(row=2, column=1)

tk.Label(root, text="Giới tính:").grid(row=2, column=2)
tk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").grid(row=2, column=3, sticky="w")
tk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").grid(row=2, column=3, sticky="e")

tk.Label(root, text="Số CMND:").grid(row=3, column=0)
entry_cmnd = tk.Entry(root)
entry_cmnd.grid(row=3, column=1)

tk.Label(root, text="Ngày cấp:").grid(row=3, column=2)
entry_ngaycap = tk.Entry(root)
entry_ngaycap.grid(row=3, column=3)

tk.Label(root, text="Nơi cấp:").grid(row=4, column=0)
entry_noicap = tk.Entry(root)
entry_noicap.grid(row=4, column=1)

# Các nút chức năng
tk.Button(root, text="Lưu thông tin", command=save_to_csv).grid(row=5, column=0, pady=10)
tk.Button(root, text="Sinh nhật hôm nay", command=filter_birthday_today).grid(row=5, column=1, pady=10)
tk.Button(root, text="Xuất danh sách", command=export_to_excel).grid(row=5, column=2, pady=10)
tk.Button(root, text="Thoát", command=root.quit).grid(row=5, column=3, pady=10)

root.mainloop()
