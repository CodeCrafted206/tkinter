import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import csv
from datetime import datetime
import pandas as pd


# Hàm lưu dữ liệu vào file CSV
def save_to_csv(data):
    with open('employees.csv', mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(data)


# Hàm xử lý nút Lưu
def save_employee():
    employee_data = [
        entry_id.get(),
        entry_name.get(),
        combobox_department.get(),
        entry_position.get(),
        entry_birthdate.get(),
        gender_var.get(),
        entry_id_card.get(),
        entry_issue_date.get(),
        entry_issue_place.get()
    ]

    if all(employee_data):  # Kiểm tra nếu tất cả thông tin đều đầy đủ
        save_to_csv(employee_data)
        messagebox.showinfo("Thông báo", "Đã lưu thông tin nhân viên thành công!")
        clear_form()
    else:
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ thông tin!")


# Hàm xóa dữ liệu trong form
def clear_form():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combobox_department.set('')
    entry_position.delete(0, tk.END)
    entry_birthdate.delete(0, tk.END)
    gender_var.set('Nam')
    entry_id_card.delete(0, tk.END)
    entry_issue_date.delete(0, tk.END)
    entry_issue_place.delete(0, tk.END)


# Hàm hiển thị nhân viên có sinh nhật hôm nay
def show_today_birthdays():
    try:
        today = datetime.now().strftime('%d/%m/%Y')
        with open('employees.csv', mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            employees = [row for row in reader if row[4] == today]

        if employees:
            result = "\n".join([f"{emp[1]} - {emp[4]}" for emp in employees])
            messagebox.showinfo("Sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên!")


# Hàm xuất danh sách toàn bộ nhân viên
def export_all_employees():
    try:
        df = pd.read_csv('employees.csv', header=None, names=[
            'Mã', 'Tên', 'Đơn vị', 'Chức danh', 'Ngày sinh', 'Giới tính', 'Số CMND', 'Ngày cấp', 'Nơi cấp'
        ])
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y')
        df['Tuổi'] = (datetime.now() - df['Ngày sinh']).dt.days // 365
        df.sort_values(by='Tuổi', ascending=False, inplace=True)

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Thông báo", f"Xuất danh sách thành công tới {file_path}")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {e}")


# Tạo giao diện
root = tk.Tk()
root.title("Quản lý nhân viên")

# Biến lưu trữ giá trị
gender_var = tk.StringVar(value="Nam")

# Tạo các thành phần trong giao diện
tk.Label(root, text="Mã:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_id = tk.Entry(root)
entry_id.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Tên:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_name = tk.Entry(root)
entry_name.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Đơn vị:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
combobox_department = ttk.Combobox(root, values=["Phân xưởng que hàn", "Phân xưởng khác"])
combobox_department.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Chức danh:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
entry_position = tk.Entry(root)
entry_position.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
entry_birthdate = tk.Entry(root)
entry_birthdate.grid(row=4, column=1, padx=5, pady=5)

tk.Label(root, text="Giới tính:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
tk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").grid(row=5, column=1, sticky="w")
tk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").grid(row=5, column=2, sticky="w")

tk.Label(root, text="Số CMND:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
entry_id_card = tk.Entry(root)
entry_id_card.grid(row=6, column=1, padx=5, pady=5)

tk.Label(root, text="Ngày cấp (DD/MM/YYYY):").grid(row=7, column=0, padx=5, pady=5, sticky="w")
entry_issue_date = tk.Entry(root)
entry_issue_date.grid(row=7, column=1, padx=5, pady=5)

tk.Label(root, text="Nơi cấp:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
entry_issue_place = tk.Entry(root)
entry_issue_place.grid(row=8, column=1, padx=5, pady=5)

# Nút Lưu
btn_save = tk.Button(root, text="Lưu", command=save_employee)
btn_save.grid(row=9, column=0, padx=5, pady=5)

# Nút Sinh nhật hôm nay
btn_birthday = tk.Button(root, text="Sinh nhật hôm nay", command=show_today_birthdays)
btn_birthday.grid(row=9, column=1, padx=5, pady=5)

# Nút Xuất toàn bộ danh sách
btn_export = tk.Button(root, text="Xuất toàn bộ danh sách", command=export_all_employees)
btn_export.grid(row=9, column=2, padx=5, pady=5)

root.mainloop()
