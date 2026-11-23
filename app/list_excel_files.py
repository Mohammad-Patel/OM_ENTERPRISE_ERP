import os
import openpyxl

def generate_excel_list(root_dir, output_filename="project_excel_files.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Project Files"
    
    # Headers
    ws.append(["Sr. No.", "File Name", "Relative Path", "Directory"])
    
    counter = 1
    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            # Agar sirf Excel files ya saari files ki list chahiye toh condition laga sakte hain
            if file.endswith(('.xlsx', '.xls', '.csv')):
                rel_path = os.path.relpath(os.path.join(dirpath, file), root_dir)
                ws.append([counter, file, rel_path, dirpath])
                counter += 1
                
    wb.save(output_filename)
    print(f"List generated successfully: {output_filename}")

if __name__ == "__main__":
    # Apne project ka root path yahan dein (current directory by default)
    generate_excel_list(".")