import openpyxl

def load_excel(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        return workbook
    except Exception as e:
        print(f"Error loading file {file_path}: {e}")
        return None

def save_excel(workbook, file_path):
    try:
        workbook.save(file_path)
        print(f"File saved successfully to {file_path}")
    except Exception as e:
        print(f"Error saving file {file_path}: {e}")
