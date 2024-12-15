
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import pandas as pd

file_name= "nhanvien.csv"
try:
    with open(file_name, "x", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Số CMND", "Nơi cấp", "Ngày cấp"])
except FileExistsError:
    pass

def save_data():
    data = [
        ma_entry.get(), ten_entry.get(), donvi_entry.get(), chucdanh_entry.get(),
        ngaysinh_entry.get(), cmnd_entry.get(), noicap_entry.get(), ngaycap_entry.get()
    ]
    if "" in data:
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin")
        return
    with open(file_name, "a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu!")
    clear_entries()

def clear_entries():
    for entry in all_entries:
        entry.delete(0, tk.END)

def show_birthdays_today():
    today = datetime.now().strftime("%d/%m/%Y")[:5]  
    birthdays = []
    with open(file_name, "r", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row["Ngày sinh"][:5] == today:
                birthdays.append(row)
    if birthdays:
        result = "\n".join([f"{nv['Tên']} - {nv['Ngày sinh']}" for nv in birthdays])
        messagebox.showinfo("Sinh nhật hôm nay", result)
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")

def export_to_excel():
    df = pd.read_csv(file_name)
    df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format="%d/%m/%Y", errors='coerce')
    df = df.sort_values(by='Ngày sinh', ascending=True)
    df.to_excel("danhsach_nhanvien.xlsx", index=False)
    messagebox.showinfo("Thành công", "Danh sách đã được xuất ra file Excel: danhsach_nhanvien.xlsx")

root = tk.Tk()
root.title("Quản lý thông tin nhân viên")

tk.Label(root, text="Mã:").grid(row=0, column=0, padx=5, pady=5)
ma_entry = tk.Entry(root)
ma_entry.grid(row=0, column=1)

tk.Label(root, text="Tên:").grid(row=0, column=2, padx=5, pady=5)
ten_entry = tk.Entry(root)
ten_entry.grid(row=0, column=3)

tk.Label(root, text="Đơn vị:").grid(row=1, column=0, padx=5, pady=5)
donvi_entry = tk.Entry(root)
donvi_entry.grid(row=1, column=1)

tk.Label(root, text="Chức danh:").grid(row=1, column=2, padx=5, pady=5)
chucdanh_entry = tk.Entry(root)
chucdanh_entry.grid(row=1, column=3)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY):").grid(row=2, column=0, padx=5, pady=5)
ngaysinh_entry = tk.Entry(root)
ngaysinh_entry.grid(row=2, column=1)

tk.Label(root, text="Số CMND:").grid(row=2, column=2, padx=5, pady=5)
cmnd_entry = tk.Entry(root)
cmnd_entry.grid(row=2, column=3)

tk.Label(root, text="Nơi cấp:").grid(row=3, column=0, padx=5, pady=5)
noicap_entry = tk.Entry(root)
noicap_entry.grid(row=3, column=1)

tk.Label(root, text="Ngày cấp (DD/MM/YYYY):").grid(row=3, column=2, padx=5, pady=5)
ngaycap_entry = tk.Entry(root)
ngaycap_entry.grid(row=3, column=3)

all_entries = [ma_entry, ten_entry, donvi_entry, chucdanh_entry, ngaysinh_entry, cmnd_entry, noicap_entry, ngaycap_entry]

tk.Button(root, text="Lưu thông tin", command=save_data).grid(row=4, column=0, columnspan=2, pady=10)
tk.Button(root, text="Sinh nhật ngày hôm nay", command=show_birthdays_today).grid(row=4, column=2, columnspan=2, pady=10)
tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel).grid(row=5, column=0, columnspan=4, pady=10)

root.mainloop()
