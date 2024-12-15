import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from datetime import datetime
import pandas as pd
import os

FILE_NAME = "nhanvien.csv"

def save_to_csv():
    data = {
        "Mã": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Ngày sinh": entry_ngaysinh.get(),
        "Giới tính": gender_var.get(),
        "Đơn vị": entry_donvi.get(),
        "Số CMND": entry_cmnd.get(),
        "Ngày cấp": entry_ngaycap.get(),
        "Nơi cấp": entry_noicap.get(),
        "Chức danh": entry_chucdanh.get()
    }

    if not all(data.values()):
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin!")
        return
    
    file_exists = os.path.isfile(FILE_NAME)
    with open(FILE_NAME, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=data.keys())
        if not file_exists:
            writer.writeheader()
        writer.writerow(data)
    
    messagebox.showinfo("Thành công", "Đã lưu thông tin nhân viên!")
    clear_fields()

def clear_fields():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_ngaysinh.delete(0, tk.END)
    entry_donvi.delete(0, tk.END)
    entry_cmnd.delete(0, tk.END)
    entry_ngaycap.delete(0, tk.END)
    entry_noicap.delete(0, tk.END)
    entry_chucdanh.delete(0, tk.END)


def show_today_birthdays():
    if not os.path.isfile(FILE_NAME):
        messagebox.showerror("Lỗi", "Chưa có dữ liệu!")
        return

    today = datetime.now().strftime("%d/%m")
    birthdays = []

    with open(FILE_NAME, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row["Ngày sinh"][:5] == today:
                birthdays.append(row["Tên"])

    if birthdays:
        messagebox.showinfo("Sinh nhật hôm nay", "\n".join(birthdays))
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")


def export_to_excel():
    if not os.path.isfile(FILE_NAME):
        messagebox.showerror("Lỗi", "Chưa có dữ liệu!")
        return


    df = pd.read_csv(FILE_NAME)
    

    df['Tuổi'] = df['Ngày sinh'].apply(lambda x: datetime.now().year - int(x[-4:]) 
                                        if isinstance(x, str) and len(x) >= 4 and x[-4:].isdigit() 
                                        else None)
    
    df = df.dropna(subset=['Tuổi'])

    df = df.sort_values(by="Tuổi", ascending=False)

    
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Chọn vị trí lưu file"
    )
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Thành công", f"File Excel đã được lưu tại:\n{file_path}")


root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("700x400")


gender_var = tk.StringVar(value="Nam")


tk.Label(root, text="Mã:").place(x=20, y=20)
entry_ma = tk.Entry(root)
entry_ma.place(x=80, y=20, width=200)

tk.Label(root, text="Tên:").place(x=300, y=20)
entry_ten = tk.Entry(root)
entry_ten.place(x=350, y=20, width=300)

tk.Label(root, text="Ngày sinh:").place(x=20, y=60)
entry_ngaysinh = tk.Entry(root)
entry_ngaysinh.place(x=80, y=60, width=200)

tk.Label(root, text="Giới tính:").place(x=300, y=60)
tk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").place(x=370, y=60)
tk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").place(x=430, y=60)

tk.Label(root, text="Đơn vị:").place(x=20, y=100)
entry_donvi = tk.Entry(root)
entry_donvi.place(x=80, y=100, width=200)

tk.Label(root, text="Số CMND:").place(x=300, y=100)
entry_cmnd = tk.Entry(root)
entry_cmnd.place(x=370, y=100, width=200)

tk.Label(root, text="Ngày cấp:").place(x=20, y=140)
entry_ngaycap = tk.Entry(root)
entry_ngaycap.place(x=80, y=140, width=200)

tk.Label(root, text="Nơi cấp:").place(x=300, y=140)
entry_noicap = tk.Entry(root)
entry_noicap.place(x=370, y=140, width=200)

tk.Label(root, text="Chức danh:").place(x=20, y=180)
entry_chucdanh = tk.Entry(root)
entry_chucdanh.place(x=100, y=180, width=470)


tk.Button(root, text="Lưu thông tin", command=save_to_csv).place(x=100, y=230, width=100)
tk.Button(root, text="Sinh nhật hôm nay", command=show_today_birthdays).place(x=220, y=230, width=150)
tk.Button(root, text="Xuất file Excel", command=export_to_excel).place(x=400, y=230, width=150)

root.mainloop()
