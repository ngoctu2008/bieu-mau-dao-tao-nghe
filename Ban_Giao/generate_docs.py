import openpyxl
from docxtpl import DocxTemplate
import os
import shutil

# 1. Read class information from the MENU sheet
wb = openpyxl.load_workbook("Quan_Ly_Hoc_Vien.xlsx", data_only=True)
menu_ws = wb['MENU']

class_info = {
    'nghe_dao_tao': menu_ws['E3'].value or "",  # Row 3 (index 2) is 'Tên nghề đào tạo'
    'khoa': menu_ws['E4'].value or "",          # Row 4 is 'Khóa'
    'nam_dao_tao': menu_ws['E5'].value or "",   # Row 5 is 'Năm đào tạo'
    'xa': menu_ws['E6'].value or "",            # Row 6 is 'Địa điểm đào tạo - Xã'
    'huyen': menu_ws['E7'].value or "",         # Row 7 is 'Địa điểm đào tạo - Huyện'
    'loai_nghe': menu_ws['E8'].value or ""      # Row 8 is 'Là nghề'
}

# 2. Read student information from the "Thông tin hoc sinh" sheet
hs_ws = wb['Thông tin hoc sinh']

students = []
# Assuming the data starts at row 3 (based on previous inspections)
for row in hs_ws.iter_rows(min_row=3, max_row=200, values_only=True):
    stt = row[0]
    ho_ten = row[1]

    # If there is no name, we reached the end of the list
    if not ho_ten:
        continue

    nam_sinh_raw = row[2]
    # Simple logic to split birth year if it's a full date, else treat as year
    ngay_sinh, thang_sinh, nam_sinh = "...", "...", str(nam_sinh_raw) if nam_sinh_raw else "..."

    gioi_tinh = row[3] or ""
    dan_toc = row[4] or ""
    nhom_doi_tuong = row[9] or ""
    nguyen_quan = row[10] or ""
    noi_cu_tru = row[11] or ""
    trinh_do_van_hoa = row[8] or ""

    student_data = {
        'ho_ten': ho_ten,
        'ngay_sinh': ngay_sinh,
        'thang_sinh': thang_sinh,
        'nam_sinh': nam_sinh,
        'gioi_tinh': gioi_tinh,
        'dan_toc': dan_toc,
        'nhom_doi_tuong': nhom_doi_tuong,
        'nguyen_quan': nguyen_quan,
        'noi_cu_tru': noi_cu_tru,
        'trinh_do_van_hoa': trinh_do_van_hoa,
        'cccd': '......................'
    }
    # Merge class info into student info
    student_data.update(class_info)
    students.append(student_data)

# 3. Generate the documents
templates_dir = "word_templates"
output_dir = "Ho_So_Da_Dien"

if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir)

# Helper function to generate documents
def generate_student_docs(student):
    safe_name = str(student['ho_ten']).replace("/", "_").replace("\\", "_")
    student_dir = os.path.join(output_dir, safe_name)
    os.makedirs(student_dir, exist_ok=True)

    for filename in os.listdir(templates_dir):
        if filename.endswith(".docx"):
            template_path = os.path.join(templates_dir, filename)
            doc = DocxTemplate(template_path)

            # The context is the student data
            doc.render(student)

            output_path = os.path.join(student_dir, filename)
            doc.save(output_path)

print(f"Found {len(students)} students. Generating documents...")
for student in students:
    generate_student_docs(student)
print("Done generating documents!")
