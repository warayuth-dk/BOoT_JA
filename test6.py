import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyperclip

# ฟังก์ชันสำหรับโหลดข้อมูลจาก Excel (ฐานข้อมูล)
def load_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"ไม่สามารถโหลดไฟล์ Excel ได้: {e}")
        return None

# ฟังก์ชันสำหรับ query ข้อมูลตามค่าที่ค้นหา
def query_data(df, search_value):
    filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_value, case=False).any(), axis=1)]
    return filtered_df

# ฟังก์ชันสำหรับเปิดไฟล์ Excel
def open_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = load_excel(file_path)
        if df is not None:
            messagebox.showinfo("Success", "ไฟล์ Excel ถูกโหลดเรียบร้อยแล้ว")

# ฟังก์ชันสำหรับการ query และแสดงผล
def query_and_display():
    if df is None:
        messagebox.showerror("Error", "กรุณาโหลดไฟล์ Excel ก่อน")
        return
    search_value = search_entry.get()
    result = query_data(df, search_value)

    # ล้างข้อมูลในตารางผลลัพธ์ก่อน
    for i in tree.get_children():
        tree.delete(i)
    
    # ล้างข้อมูลใน Text widget ก่อน
    output_text.delete(1.0, tk.END)

    if not result.empty:
        selected_type = type_dropdown.get()
        
        # เติมค่าใน Dropdown ตามประเภทที่ต้องการค้นหา ถ้า selected_type เป็น CDIT
        if selected_type == "CDIT":
            dropdown_values = result.iloc[:, 0].unique().tolist()  # สมมติว่าใช้คอลัมน์แรกในการกรอง
            type_dropdown['values'] = dropdown_values  # อัปเดตค่าใน Dropdown
            
            query_values = []
            for _, row in result.iterrows():
                k_value = row[10] if pd.notna(row[10]) else ""  # คอลัมน์ K
                if any(keyword in k_value for keyword in ["SSD", "BU", "WLAN", "OST"]):
                    query_values.append(str(row[9]))  # คอลัมน์ J
            
            # แยกค่าและจัดกลุ่มตามลำดับที่ต้องการ
            ordered_keywords = ["SSD", "BU", "WLAN", "OS"]
            ordered_results = {key: [] for key in ordered_keywords}
            
            for _, row in result.iterrows():
                k_value = row[10] if pd.notna(row[10]) else ""
                for keyword in ordered_keywords:
                    if keyword in k_value:
                        ordered_results[keyword].append(str(row[9]))  # คอลัมน์ J

            # สร้างข้อความผลลัพธ์ตามลำดับ
            output_text_value = []
            for keyword in ordered_keywords:
                output_text_value.extend(ordered_results[keyword])  # เพิ่มค่าที่อยู่ในลิสต์ตามลำดับ
            
            if output_text_value:
                output_text.insert(tk.END, "/".join(output_text_value) + "\n")  # แสดงผลใน Text widget
        else:
            for _, row in result.iterrows():
                # แสดงข้อมูลใน Treeview
                tree.insert("", "end", values=(row[1], row[9], row[10]))  # แสดงข้อมูลจากคอลัมน์ B, J และ K
                # แสดงข้อมูลใน Text widget
                output_text.insert(tk.END, f"{row[1]}: {row[9]} - {row[10]}\n")  # แสดงข้อมูลตามที่มีในคอลัมน์ B, J และ K     
    else:
        messagebox.showinfo("Info", "ไม่พบข้อมูลที่ตรงกับการค้นหา")

# ฟังก์ชันสำหรับคัดลอกข้อมูลที่เลือก
def copy_selection():
    selected_text = output_text.get("1.0", tk.END).strip()
    pyperclip.copy(selected_text)

# สร้างหน้าต่างหลักของ GUI
root = tk.Tk()
root.title("Excel Database Query Program")

df = None

# ปรับการตั้งค่าของ grid ให้ขยายได้
root.grid_rowconfigure(6, weight=1)  # ทำให้แถวที่ 6 ขยาย
root.grid_columnconfigure(0, weight=1)  # ทำให้คอลัมน์ที่ 0 ขยาย
root.grid_columnconfigure(1, weight=1)  # ทำให้คอลัมน์ที่ 1 ขยาย

tk.Label(root, text="ค้นหาข้อมูล:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
search_entry = tk.Entry(root)
search_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

tk.Button(root, text="เปิดไฟล์ Excel", command=open_file).grid(row=1, column=0, columnspan=2, padx=10, pady=10)
tk.Button(root, text="Query ข้อมูล", command=query_and_display).grid(row=2, column=0, columnspan=2, padx=10, pady=10)
tk.Button(root, text="คัดลอกข้อมูล", command=copy_selection).grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# สร้าง Dropdown สำหรับประเภทที่ต้องการค้นหา
tk.Label(root, text="ประเภทที่ต้องการค้นหา:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
type_dropdown = ttk.Combobox(root)
type_dropdown['values'] = ["CDIT","CSAN"]  # เพิ่ม CDIT เป็นค่าใน dropdown
type_dropdown.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

# สร้างตารางแสดงผลลัพธ์ (Treeview)
tree = ttk.Treeview(root, columns=("ColB", "ColJ", "ColK"), show='headings', height=10)
tree.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")  # ใช้ sticky เพื่อให้เต็มพื้นที่

tree.heading("ColB", text="คอลัมน์ B")
tree.heading("ColJ", text="คอลัมน์ J")
tree.heading("ColK", text="คอลัมน์ K")

# สร้าง Text widget สำหรับแสดงข้อความเพิ่มเติม
output_text = tk.Text(root, height=10, wrap=tk.WORD)
output_text.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")  # ใช้ sticky เพื่อให้เต็มพื้นที่

root.mainloop()
