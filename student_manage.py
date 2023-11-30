from tkinter import *
from tkinter import ttk
import sys,mysql.connector
from openpyxl import Workbook
from tkinter import messagebox
#Hàm connect db.
add_window = None

def connection():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="123456",
            database="quan_ly_diem"
        )
        return conn
    except mysql.connector.Error as e:
        print("Error connecting to the database:", e)
        return None
#thêm sinh viên
def add_sinhvien(ma_sv, hoten_sv, ngay_sinh, gioi_tinh, dan_toc, noi_sinh, ma_lop):
    conn = connection()
    cur = conn.cursor()
    try:
        # Kiểm tra xem mã sinh viên đã tồn tại hay chưa
        cur.execute("SELECT ma_sv FROM sinhvien WHERE ma_sv = %s", (ma_sv,))
        existing_sv = cur.fetchone()

        if existing_sv:
            conn.close()
            messagebox.showerror("Lỗi", "Mã sinh viên đã tồn tại")
        else:
            # Thêm sinh viên nếu mã sinh viên chưa tồn tại
            cur.execute("""
                INSERT INTO sinhvien (ma_sv, hoten_sv, ngay_sinh, gioi_tinh, dan_toc, noi_sinh, ma_lop) 
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (ma_sv, hoten_sv, ngay_sinh, gioi_tinh, dan_toc, noi_sinh, ma_lop))
            conn.commit()
            conn.close()
            messagebox.showinfo("Thông báo", "Thêm sinh viên thành công")
    except mysql.connector.Error as e:
        print("Error adding data to database:", e)
        conn.rollback()
        conn.close()
#xoá sinh viên 
def delete_sinhvien(ma_sv):
    conn = connection()
    cur = conn.cursor()
    try:
        # Kiểm tra xem sinh viên tồn tại hay không trước khi xoá
        cur.execute("SELECT ma_sv FROM sinhvien WHERE ma_sv = %s", (ma_sv,))
        existing_sv = cur.fetchone()

        if existing_sv:
            # Nếu sinh viên tồn tại, thực hiện xoá
            cur.execute("DELETE FROM sinhvien WHERE ma_sv = %s", (ma_sv,))
            conn.commit()
            conn.close()
            messagebox.showinfo("Thông báo", "Xoá sinh viên thành công")
        else:
            # Nếu sinh viên không tồn tại, hiển thị thông báo lỗi
            conn.close()
            messagebox.showerror("Lỗi", "Không tìm thấy sinh viên có mã này")
    except mysql.connector.Error as e:
        print("Error deleting data from database:", e)
        conn.rollback()
        conn.close()
#end
#Điểm học phần Start################################################################################

def view_diemhocphan():
    # Sử dụng biến global.
    global tree
    global search_diemhocphan_entry
    # Xóa nội dung cũ trong Text widget

    # Tạo input search và button search và clear search
    search_diemhocphan_entry = Entry(root, width=40)
    search_diemhocphan_entry.grid(row=6, column=0, padx=0, pady=5)

    search_diemhocphan_button = Button(root, text="Tìm kiếm", command=search_diemhocphan, width=20)
    search_diemhocphan_button.grid(row=7, column=0, padx=0, pady=5)

    add_diemhocphan_button = Button(root, text="Thêm điểm học phần", command=add_diemhocphan_window, width=20)
    add_diemhocphan_button.grid(row=12, column=0, padx=0, pady=5)

    conn = connection()
    cur = conn.cursor()
    cur.execute("select diemhocphan.ma_sv,hoten_sv,ma_lop,diemhocphan.ma_mon,ten_mon,sotinchi,diem_giua_ky,diem_thi_hp from diemhocphan inner join sinhvien on diemhocphan.ma_sv=sinhvien.ma_sv inner join monhocphan on monhocphan.ma_mon=diemhocphan.ma_mon")
    data = cur.fetchall()
    conn.close()

    # Tạo một Treeview widget
    tree = ttk.Treeview(root, columns=(1, 2, 3, 4, 5, 6, 7, 8), show="headings", height=20)

    # Đặt tên các cột
    tree.heading(1, text="Mã SV")
    tree.heading(2, text="Họ và tên SV")
    tree.heading(3, text="Mã lớp")
    tree.heading(4, text="Mã học phần")
    tree.heading(5, text="Tên học phần")
    tree.heading(6, text="Số tín chỉ")
    tree.heading(7, text="Điểm giữa kỳ")
    tree.heading(8, text="Điểm cuối kỳ")
    
    # Đặt lại chiều rộng của mã SV, Mã lớp, số tín chỉ.
    tree.column(1, width=60)
    tree.column(3, width=90)
    tree.column(6, width=60)
    

    # Hiển thị dữ liệu trong Treeview
    for row in data:
        tree.insert("", "end", values=row)

    # Thêm Treeview vào cửa sổ chính
    tree.grid(row=0, column=1, pady=0, padx=10,rowspan=26)

    # Tạo và thiết lập Scrollbar
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    scrollbar.grid(row=0, column=2, sticky="ns",rowspan=26)
    tree.configure(yscrollcommand=scrollbar.set)

    def on_item_click(event):
        # Lấy thông tin của dòng được chọn
        selected_item = tree.selection()[0]
        values = tree.item(selected_item, 'values')

        # Hiển thị cửa sổ chi tiết điểm học phần
        # view_diemhocphan_details(values)
        view_Diemhp_details(values)
    # Gán hàm on_item_click khi click vào dòng
    tree.bind('<ButtonRelease-1>', on_item_click)
# tổng số tín chỉ đã học 
def total_credit_hours_of_student(ma_sv):
    conn = connection()
    cur = conn.cursor()
    try:
        # Truy vấn cơ sở dữ liệu để lấy tổng số tín chỉ đã học của sinh viên theo mã sinh viên
        cur.execute("""
            SELECT SUM(sotinchi) 
            FROM diemhocphan 
            INNER JOIN monhocphan ON diemhocphan.ma_mon = monhocphan.ma_mon 
            WHERE diemhocphan.ma_sv = %s
        """, (ma_sv,))
        total_credit_hours = cur.fetchone()[0]  # Lấy tổng số tín chỉ từ kết quả truy vấn

        conn.close()

        return total_credit_hours if total_credit_hours else 0  # Trả về tổng số tín chỉ, nếu không có trả về 0
    except mysql.connector.Error as e:
        print("Error retrieving total credit hours:", e)
        conn.close()
        return 0  # Trả về 0 nếu có lỗi khi truy vấn cơ sở dữ liệu


def view_diemhocphan_details(student_info):
    # Tạo một cửa sổ mới
    details_window = Toplevel(root)
    details_window.title("SV:" + student_info[1]) 

    # Hiển thị thông tin chi tiết của dòng được chọn
    Label(details_window, text="Mã SV:").grid(row=0, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[0]).grid(row=0, column=1, padx=10, pady=5)

    Label(details_window, text="Họ và tên SV:").grid(row=1, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[1]).grid(row=1, column=1, padx=10, pady=5)

    Label(details_window, text="Năm Sinh:").grid(row=2, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[2]).grid(row=2, column=1, padx=10, pady=5)

    Label(details_window, text="Giới Tính:").grid(row=3, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[3]).grid(row=3, column=1, padx=10, pady=5)

    Label(details_window, text="Dân Tộc:").grid(row=4, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[4]).grid(row=4, column=1, padx=10, pady=5)

    Label(details_window, text="Địa Chỉ:").grid(row=5, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[5]).grid(row=5, column=1, padx=10, pady=5)

    # Sử dụng Entry để cho phép chỉnh sửa
    Label(details_window, text="Mã lớp:").grid(row=6, column=0, padx=10, pady=5)
    midterm_entry = Entry(details_window)
    midterm_entry.grid(row=6, column=1, padx=10, pady=5)
    midterm_entry.insert(0, student_info[6])  # Hiển thị giá trị hiện tại
    total_credit_hours = total_credit_hours_of_student(student_info[0])
    Label(details_window, text="Tổng số tín chỉ đã học :").grid(row=7, column=0, padx=10, pady=5)
    final_entry = Entry(details_window)
    Label(details_window, text=total_credit_hours).grid(row=7, column=1, padx=10, pady=5)
    final_entry.insert(0, student_info[7])  # Hiển thị giá trị hiện tại
    delete_button = Button(details_window, text="Xoá", command=lambda: delete_sinhvien(student_info[0]))
    delete_button.grid(row=8, columnspan=2, column=1, padx=10, pady=5)

def view_Diemhp_details(student_info):
    # Tạo một cửa sổ mới
    details_window = Toplevel(root)
    details_window.title("SV:" + student_info[1]) 

    # Hiển thị thông tin chi tiết của dòng được chọn
    Label(details_window, text="Mã SV:").grid(row=0, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[0]).grid(row=0, column=1, padx=10, pady=5)

    Label(details_window, text="Họ và tên SV:").grid(row=1, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[1]).grid(row=1, column=1, padx=10, pady=5)

    Label(details_window, text="Mã lớp:").grid(row=2, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[2]).grid(row=2, column=1, padx=10, pady=5)

    Label(details_window, text="Mã học phần:").grid(row=3, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[3]).grid(row=3, column=1, padx=10, pady=5)

    Label(details_window, text="Tên học phần:").grid(row=4, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[4]).grid(row=4, column=1, padx=10, pady=5)

    Label(details_window, text="Số tín chỉ:").grid(row=5, column=0, padx=10, pady=5)
    Label(details_window, text=student_info[5]).grid(row=5, column=1, padx=10, pady=5)

    # Sử dụng Entry để cho phép chỉnh sửa
    Label(details_window, text="Điểm giữa kỳ:").grid(row=6, column=0, padx=10, pady=5)
    midterm_entry = Entry(details_window)
    midterm_entry.grid(row=6, column=1, padx=10, pady=5)
    midterm_entry.insert(0, student_info[6])  # Hiển thị giá trị hiện tại

    Label(details_window, text="Điểm thi HP:").grid(row=7, column=0, padx=10, pady=5)
    final_entry = Entry(details_window)
    final_entry.grid(row=7, column=1, padx=10, pady=5)
    final_entry.insert(0, student_info[7])  # Hiển thị giá trị hiện tại

    # Hàm cập nhật điểm
    def update_grade():
        # Lấy giá trị mới từ các ô nhập liệu
        new_midterm = midterm_entry.get()
        new_final = final_entry.get()

        # Thực hiện cập nhật vào cơ sở dữ liệu
        conn = connection()
        cur = conn.cursor()
        cur.execute("UPDATE diemhocphan SET diem_giua_ky=%s, diem_thi_hp=%s WHERE ma_sv=%s AND ma_mon=%s", (new_midterm, new_final, student_info[0], student_info[3]))
        conn.commit()
        conn.close()

        # Cập nhật lại dữ liệu trong Treeview
        view_diemhocphan()

        # Đóng cửa sổ chi tiết sau khi cập nhật
        details_window.destroy()


    # Nút cập nhật điểm
    update_button = Button(details_window, text="Cập nhật điểm", command=update_grade, width=20)
    update_button.grid(row=8, column=1, padx=10, pady=10)

#danh sách học phần END
    # Hàm cập nhật điểm
    def update_grade():
        # Lấy giá trị mới từ các ô nhập liệu
        new_midterm = midterm_entry.get()
        new_final = final_entry.get()

        # Thực hiện cập nhật vào cơ sở dữ liệu
        conn = connection()
        cur = conn.cursor()
        cur.execute("UPDATE diemhocphan SET diem_giua_ky=%s, diem_thi_hp=%s WHERE ma_sv=%s AND ma_mon=%s", (new_midterm, new_final, student_info[0], student_info[3]))
        conn.commit()
        conn.close()

        # Cập nhật lại dữ liệu trong Treeview
        view_diemhocphan()

        # Đóng cửa sổ chi tiết sau khi cập nhật
        details_window.destroy()
    # Nút cập nhật điểm
    update_button = Button(details_window, text="Cập nhật điểm", command=update_grade, width=20)
    update_button.grid(row=8, column=1, padx=10, pady=10)


def search_diemhocphan():

    global tree 

    search_query = search_diemhocphan_entry.get().lower()
    conn = connection()
    cur = conn.cursor()
    cur.execute("SELECT diemhocphan.ma_sv,hoten_sv,ma_lop,diemhocphan.ma_mon,ten_mon,sotinchi,diem_giua_ky,diem_thi_hp FROM diemhocphan INNER JOIN sinhvien ON diemhocphan.ma_sv=sinhvien.ma_sv INNER JOIN monhocphan ON monhocphan.ma_mon=diemhocphan.ma_mon WHERE LOWER(hoten_sv) LIKE %s OR diemhocphan.ma_sv LIKE %s OR diemhocphan.ma_mon LIKE %s OR ma_lop LIKE %s OR LOWER(ten_mon) LIKE %s", ('%' + search_query + '%', '%' + search_query + '%', '%' + search_query + '%', '%' + search_query + '%', '%' + search_query + '%'))
    data = cur.fetchall()
    conn.close()

    # Clear the Treeview
    for item in tree.get_children():
        tree.delete(item)

    # Display the filtered data in the Treeview
    for row in data:
        tree.insert("", "end", values=row)

def add_diemhocphan_window():
    global add_window
    add_window = Toplevel(root)
    add_window.title("Thêm Điểm Học Phần")

    # Thêm các Label và Entry để nhập thông tin điểm học phần mới
    Label(add_window, text="Mã SV:").grid(row=0, column=0, padx=10, pady=5)
    ma_sv_entry = Entry(add_window, width=40)
    ma_sv_entry.grid(row=0, column=1, padx=10, pady=5)

    Label(add_window, text="Mã Môn Học:").grid(row=1, column=0, padx=10, pady=5)
    ma_mon_values = get_ma_mon_values()
    ma_mon_combobox = ttk.Combobox(add_window, values=ma_mon_values, width=37)
    ma_mon_combobox.grid(row=1, column=1, padx=10, pady=5)

    Label(add_window, text="Điểm Giữa Kỳ:").grid(row=2, column=0, padx=10, pady=5)
    diem_giua_ky_entry = Entry(add_window, width=40)
    diem_giua_ky_entry.grid(row=2, column=1, padx=10, pady=5)

    Label(add_window, text="Điểm Thi HP:").grid(row=3, column=0, padx=10, pady=5)
    diem_thi_hp_entry = Entry(add_window, width=40)
    diem_thi_hp_entry.grid(row=3, column=1, padx=10, pady=5)

    # Thêm nút để thực hiện thêm điểm học phần
    add_button.grid(row=4, column=1, padx=10, pady=10)
def add_diemhocphan(ma_sv, ma_mon, diem_giua_ky, diem_thi_hp):
    global add_window 
    conn = connection()
    cur = conn.cursor()
    try:
        # Try to insert the data, if there's a duplicate key, update the values
        cur.execute("""
            INSERT INTO diemhocphan (ma_sv, ma_mon, diem_giua_ky, diem_thi_hp) 
            VALUES (%s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE 
            diem_giua_ky = VALUES(diem_giua_ky), diem_thi_hp = VALUES(diem_thi_hp)
        """, (ma_sv, ma_mon, diem_giua_ky, diem_thi_hp))
        conn.commit()
        view_diemhocphan()  # Refresh the view after adding/updating data
        add_window.destroy()  # Close the add_window if the operation is successful
        if add_window:
            add_window.destroy()
            messagebox.showinfo("Success", "Cập nhật thành công!")
        conn.close()
        # Check if the record was updated or inserted successfully
        
        conn.close()
    except mysql.connector.Error as e:
        print("Thêm thất bại", e)
        conn.rollback()
        conn.close()
#Điểm học phần END ################################################################################

#Sinh viên Start#########################################################
def view_sinhvien():
    # Sử dụng biến global.
    global tree
    global search_sinhvien_entry
    # Xóa nội dung cũ trong Text widget

    # Tạo input search và button search và clear search
    search_sinhvien_entry = Entry(root, width=40)
    search_sinhvien_entry.grid(row=6, column=0, padx=0, pady=5)

    search_sinhvien_button = Button(root, text="Tìm kiếm", command=search_diemhocphan, width=20)
    search_sinhvien_button.grid(row=7, column=0, padx=0, pady=5)

    add_sinhvien_button = Button(root, text="Thêm sinh viên", command=add_sinhvien_window, width=20)
    add_sinhvien_button.grid(row=12, column=0, padx=0, pady=5)

    conn = connection()
    cur = conn.cursor()
    cur.execute("select * from sinhvien")
    data = cur.fetchall()
    conn.close()

    # Tạo một Treeview widget
    tree = ttk.Treeview(root, columns=(1, 2, 3, 4, 5, 6, 7), show="headings", height=20)

    # Đặt tên các cột
    tree.heading(1, text="Mã SV")
    tree.heading(2, text="Họ và tên SV")
    tree.heading(3, text="Ngày sinh")
    tree.heading(4, text="Giới tính")
    tree.heading(5, text="Dân tộc")
    tree.heading(6, text="Nơi sinh")
    tree.heading(7, text="Mã lớp")
    
    # Đặt lại chiều rộng của mã SV, Mã lớp, số tín chỉ.
    tree.column(1, width=60)
    tree.column(6, width=150)
    

    # Hiển thị dữ liệu trong Treeview
    for row in data:
        tree.insert("", "end", values=row)

    # Thêm Treeview vào cửa sổ chính
    tree.grid(row=0, column=1, pady=0, padx=10,rowspan=26)

    # Tạo và thiết lập Scrollbar
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    scrollbar.grid(row=0, column=2, sticky="ns",rowspan=26)
    tree.configure(yscrollcommand=scrollbar.set)

    def on_item_click(event):
        # Lấy thông tin của dòng được chọn
        selected_item = tree.selection()[0]
        values = tree.item(selected_item, 'values')

        # Hiển thị cửa sổ chi tiết điểm học phần
        view_diemhocphan_details(values)

    # Gán hàm on_item_click khi click vào dòng
    tree.bind('<ButtonRelease-1>', on_item_click)
#Sinh viên End#########################################################
#Học phần Start########################################################
def get_ma_mon_values():
    conn = connection()
    if conn:
        try:
            cur = conn.cursor()
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT ma_mon FROM monhocphan")
            ma_mon_values = [row[0] for row in cur.fetchall()]
            return ma_mon_values
        except mysql.connector.Error as e:
            print("Error fetching ma_mon values:", e)
        finally:
            conn.close()

    # Return an empty list if there was an issue with the database connection
    return []
#Học phần End##########################################################

def update_info_window():
    # Tạo một cửa sổ mới
    update_window = Toplevel(root)
    update_window.title("Update Student Info")

    # Các biến lưu giữ thông tin cập nhật
    update_student_name = StringVar()
    update_branch = StringVar()
    update_phone = StringVar()
    update_father = StringVar()
    update_address = StringVar()

    # Các nhãn và ô nhập liệu trong cửa sổ cập nhật
    Label(update_window, text="New Student Name:").grid(row=0, column=0, padx=10, pady=5)
    Entry(update_window, textvariable=update_student_name).grid(row=0, column=1, padx=10, pady=5)

    Label(update_window, text="New Branch:").grid(row=1, column=0, padx=10, pady=5)
    Entry(update_window, textvariable=update_branch).grid(row=1, column=1, padx=10, pady=5)

    Label(update_window, text="New Phone Number:").grid(row=2, column=0, padx=10, pady=5)
    Entry(update_window, textvariable=update_phone).grid(row=2, column=1, padx=10, pady=5)

    Label(update_window, text="New Father Name:").grid(row=3, column=0, padx=10, pady=5)
    Entry(update_window, textvariable=update_father).grid(row=3, column=1, padx=10, pady=5)

    Label(update_window, text="New Address:").grid(row=4, column=0, padx=10, pady=5)
    Entry(update_window, textvariable=update_address).grid(row=4, column=1, padx=10, pady=5)

    # Nút "Update" trong cửa sổ cập nhật
    Button(update_window, text="Update", command=clse, width=20).grid(row=5, column=1, padx=10, pady=10) 
def xuat_tatca_sinhvien_diem():
    workbook = Workbook()
    global add_window 
    # Tạo worksheet cho danh sách sinh viên
    ws_sinhvien = workbook.create_sheet(title="Danh sách sinh viên")
    headers_sinhvien = ["Mã SV", "Họ và tên SV", "Ngày sinh", "Giới tính", "Dân tộc", "Nơi sinh", "Mã lớp"]
    ws_sinhvien.append(headers_sinhvien)

    conn = connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sinhvien")
    data_sinhvien = cur.fetchall()
    conn.close()

    for row in data_sinhvien:
        ws_sinhvien.append(row)

    # Tạo worksheet cho bảng điểm học phần
    ws_diemhocphan = workbook.create_sheet(title="Bảng điểm học phần")
    headers_diemhocphan = [
        "Mã SV", "Họ và tên SV", "Mã lớp", "Mã học phần", "Tên học phần", "Số tín chỉ", "Điểm giữa kỳ", "Điểm cuối kỳ"
    ]
    ws_diemhocphan.append(headers_diemhocphan)

    conn = connection()
    cur = conn.cursor()
    cur.execute("SELECT diemhocphan.ma_sv,hoten_sv,ma_lop,diemhocphan.ma_mon,ten_mon,sotinchi,diem_giua_ky,diem_thi_hp FROM diemhocphan INNER JOIN sinhvien ON diemhocphan.ma_sv=sinhvien.ma_sv INNER JOIN monhocphan ON monhocphan.ma_mon=diemhocphan.ma_mon")
    data_diemhocphan = cur.fetchall()
    conn.close()

    for row in data_diemhocphan:
        ws_diemhocphan.append(row)

    # Lưu file Excel với tên file
    excel_filename = "TatCaSinhVien_DiemHocPhan.xlsx"
    workbook.save(excel_filename)
    print(f"Excel file '{excel_filename}' exported successfully!")
def clse():
    sys.exit() 
def add_sinhvien_window():
    add_window = Toplevel(root)
    add_window.title("Thêm Sinh Viên")

    Label(add_window, text="Mã SV:").grid(row=0, column=0, padx=10, pady=5)
    ma_sv_entry = Entry(add_window, width=40)
    ma_sv_entry.grid(row=0, column=1, padx=10, pady=5)

    Label(add_window, text="Họ và tên SV:").grid(row=1, column=0, padx=10, pady=5)
    hoten_sv_entry = Entry(add_window, width=40)
    hoten_sv_entry.grid(row=1, column=1, padx=10, pady=5)

    Label(add_window, text="Ngày sinh (YYYY-MM-DD):").grid(row=2, column=0, padx=10, pady=5)
    ngay_sinh_entry = Entry(add_window, width=40)
    ngay_sinh_entry.grid(row=2, column=1, padx=10, pady=5)

    Label(add_window, text="Giới tính (Nam/Nữ):").grid(row=3, column=0, padx=10, pady=5)
    gioi_tinh_entry = Entry(add_window, width=40)
    gioi_tinh_entry.grid(row=3, column=1, padx=10, pady=5)

    Label(add_window, text="Dân tộc:").grid(row=4, column=0, padx=10, pady=5)
    dan_toc_entry = Entry(add_window, width=40)
    dan_toc_entry.grid(row=4, column=1, padx=10, pady=5)

    Label(add_window, text="Nơi sinh:").grid(row=5, column=0, padx=10, pady=5)
    noi_sinh_entry = Entry(add_window, width=40)
    noi_sinh_entry.grid(row=5, column=1, padx=10, pady=5)

    Label(add_window, text="Mã lớp:").grid(row=6, column=0, padx=10, pady=5)
    ma_lop_entry = Entry(add_window, width=40)
    ma_lop_entry.grid(row=6, column=1, padx=10, pady=5)

    add_button = Button(add_window, text="Thêm Sinh Viên", command=lambda: add_sinhvien(
        ma_sv_entry.get(), hoten_sv_entry.get(), ngay_sinh_entry.get(),
        gioi_tinh_entry.get(), dan_toc_entry.get(), noi_sinh_entry.get(), ma_lop_entry.get()
    ), width=20)
    add_button.grid(row=7, column=1, padx=10, pady=10)
    # update_info_window()
# Hãy thay đổi màu sắc cho các nút button trong hàm main:
if __name__ == "__main__":
    root = Tk()
    root.title("Quản lý điểm sinh viên")

    # ... (các dòng code khác)

    # Đặt màu sắc và font cho các nút button
    button_color = '#3498db'  # Màu xanh dương
    button_font = ('Arial', 12)

    # Danh sách sinh viên
    b1 = Button(root, text="Danh sách sinh viên", command=view_sinhvien, width=40)
    b1.grid(row=0, column=0)
    b1.config(bg=button_color, fg='white', font=button_font)  # Cấu hình màu và font

    # Thêm sinh viên
    b2 = Button(root, text="Thêm sinh viên", command=add_sinhvien_window, width=40)
    b2.grid(row=1, column=0)
    b2.config(bg=button_color, fg='white', font=button_font)  # Cấu hình màu và font

    # Danh sách học phần
    b3 = Button(root, text="Danh sách học phần", command=view_diemhocphan, width=40)
    b3.grid(row=2, column=0)
    b3.config(bg=button_color, fg='white', font=button_font)  # Cấu hình màu và font

    # Danh sách điểm học phần
    b4 = Button(root, text="Danh sách điểm học phần", command=view_diemhocphan, width=40)
    b4.grid(row=3, column=0)
    b4.config(bg=button_color, fg='white', font=button_font)  # Cấu hình màu và font

    # Xuất excel
    b8 = Button(root, text="Xuất excel", command=xuat_tatca_sinhvien_diem, width=40)
    b8.grid(row=4, column=0)
    b8.config(bg=button_color, fg='white', font=button_font)  # Cấu hình màu và font

    # Đóng ứng dụng
    b7 = Button(root, text="Đóng ứng dụng", command=clse, width=40)
    b7.grid(row=23, column=0)
    b7.config(bg='red', fg='white', font=button_font)  # Cấu hình màu và font

    root.resizable(False, False)
    root.mainloop()
#cách 2 
# if __name__=="__main__":
#     root=Tk()
#     root.title("Quản lý điểm sinh viên")
    
#     t1=Text(root,width=140,height=25)
#     t1.grid(row=0,column=1,rowspan=26, padx=10)
#     def change_color(button):
#         button['bg'] = 'blue'

#     b1=Button(root,text="Danh sách sinh viên",command=view_sinhvien,width=40)
#     b1.grid(row=0,column=0)

#     b2=Button(root,text="Thêm sinh viên",command=add_sinhvien_window,width=40)
#     b2.grid(row=1,column=0)

#     b3=Button(root,text="Danh sách học phần",command=view_diemhocphan,width=40)
#     b3.grid(row=2,column=0)

#     b4=Button(root,text="Danh sách điểm học phần",command=view_diemhocphan,width=40)
#     b4.grid(row=3,column=0)
#     b8=Button(root,text="Xuất excel",command=xuat_tatca_sinhvien_diem,width=40)
#     b8.grid(row=4,column=0)
#     b7=Button(root,text="Đóng ứng dụng",command=clse,width=40)
#     b7.grid(row=23,column=0)
#     b7['bg'] = 'red'
#     # b1['bg'] = 'blue'
#     # b2['bg'] = 'blue'
#     # b4['bg'] = 'blue'
#     # b3['bg'] = 'blue'
#     # b8['bg'] = 'blue'

#     root.resizable(False, False)
#     root.mainloop()