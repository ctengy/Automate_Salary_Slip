import openpyxl as opxl
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import os
import io
import sys
import msoffcrypto as mso
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox

def show_text_to_GUI(text):
    lable_detail.insert(tk.END, text)
    window.update()

def load_encrypted_excel(encrypted_filename: str):#載入有密碼的excel
    
    decrypted = io.BytesIO()
    password = input_password.get()
    with open(encrypted_filename, "rb") as f:
        file = mso.OfficeFile(f)
        file.load_key(password=password)
        file.decrypt(decrypted)
    return decrypted

def replace_word_text(word_file, old_text: str, new_text: str):
    try:
        for paragraph in word_file.paragraphs:
            paragraph.text = paragraph.text.replace(old_text, new_text)

        for table in word_file.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = cell.text.replace(old_text, new_text)
    except:
        show_text_to_GUI('填入錯誤\n')

def get_password(path_password):
    try:
        show_text_to_GUI('正在取得密碼\n')
        password_workbook = opxl.load_workbook(path_password)
        password_sheet = password_workbook.worksheets[0]
        password_dic = {
            password_sheet.cell(row=pwrow, column=1).value: password_sheet.cell(row=pwrow, column=2).value
            for pwrow in range(1, password_sheet.max_row + 1)}
        show_text_to_GUI('取得密碼成功\n')
        return password_dic
    except:
        show_text_to_GUI('取得密碼失敗\n')
        raise

def encrypt_files(file_path, password_dic: dict, name):
    try:
        doc = word.Documents.Open(file_path)
        doc.Password = str(password_dic[f'{name}'])
        doc.Save()
        doc.Close()
        show_text_to_GUI(f'{name}薪資單已加密完畢\n')
    except KeyError:
        pass
    except:
        show_text_to_GUI('加密失敗\n')
        raise

def files_path():
    global source_file, path_docx, path_password, generate_path
    try:
        root = tk.Tk()
        root.withdraw()

        source_file = filedialog.askopenfilename(title='選擇轉換檔案')
        path_docx = filedialog.askopenfilename(title='選擇範例word')
        path_password = filedialog.askopenfilename(title='選擇密碼表')
        generate_path = filedialog.askdirectory(title='選擇生成檔案資料夾')

    except:
        show_text_to_GUI('檢查資料夾中檔案是否正確\n')
        messagebox.showerror("錯誤", "檢查資料夾中的檔案是否正確")
        raise

def load_files():
    global sheet, paydate
    try:
        workbook = opxl.load_workbook(source_file, data_only=True)
        sheet = workbook['薪資總表']
        paydate = str(sheet['A2'].value)[0:11]
    except:
        try:
            decryptedwb = load_encrypted_excel(source_file)
            workbook = opxl.load_workbook(decryptedwb, data_only=True)
            sheet = workbook['薪資總表']
            paydate = str(sheet['A2'].value)[0:11]
        except:
            show_text_to_GUI('讀取錯誤\n檢查EXCEL檔案或密碼\n')
            messagebox.showerror("錯誤", "讀取錯誤，請檢查EXCEL檔案或密碼")
            raise

def get_user_input():
    global star_row, end_row
    while True:
        try:
            star_row = int(input_start_row.get())
            end_row = int(input_end_row.get())
            messagebox.showinfo("提示", "關閉 Excel 檔案後按下 Enter 鍵開始生成")
            break
        except ValueError:
            messagebox.showerror("錯誤", "請輸入整數！")

def table_replace():
    global word
    load_files()
    get_user_input()
    try:
        password = get_password(path_password)
        word = win32.Dispatch('Word.Application')
        for table_row in range(star_row + 1, end_row + 1):
            word_file = Document(path_docx)
            for table_col in range(3, sheet.max_column + 1):
                old_text = str(sheet.cell(row=star_row, column=table_col).value)
                new_value = sheet.cell(row=table_row, column=table_col).value
                new_text = str(new_value) if new_value not in [None, 0] else ""
                replace_word_text(word_file, old_text, new_text)
            name = str(sheet.cell(row=table_row, column=6).value)
            generate_file_name = f"{paydate}薪資單-{name}.docx"
            title_text = word_file.paragraphs[0].text.replace('#發薪日期', paydate)
            word_file.paragraphs[0].text = title_text
            title_text_style = word_file.paragraphs[0].runs[0]
            title_text_style.font.name = "Arial"
            title_text_style.bold = True
            title_text_style._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
            title_text_style.font.size = Pt(14)
            word_file.save(os.path.join(generate_path, generate_file_name))
            show_text_to_GUI(f'{name}薪資單生成完畢\n')
            path_unprotectdocx = os.path.normpath(os.path.join(generate_path, generate_file_name))
            encrypt_files(path_unprotectdocx, password, name)
        word.Quit()
        show_text_to_GUI('薪資單均已加密完畢\n')
        messagebox.showerror("完成","完成!!")
    except:
        show_text_to_GUI( '填入過程中出現錯誤\n')
        messagebox.showerror("錯誤", "填入過程中出現錯誤")
        raise

def close_window():
    window.destroy()
    sys.exit()

def main():
    global input_start_row, input_end_row, detail, window,lable_detail, input_password
    
    window = tk.Tk()
    window.title("薪資單生成程式")
    width = 600
    height = 400
    left = 0
    top = 0
    window.geometry(f'{width}x{height}+{left}+{top}') 
    detail = tk.StringVar()
    button_load_files = tk.Button(window, text="選擇檔案", command=files_path)
    button_load_files.pack()
    label_password = tk.Label(window, text="檔案密碼")
    label_password.pack()
    input_password = tk.Entry(window)
    input_password.pack()
    label_start_row = tk.Label(window, text="薪資表表頭位置:")
    label_start_row.pack()
    input_start_row = tk.Entry(window)
    input_start_row.pack()
    label_end_row = tk.Label(window, text="人員名單結束位置:")
    label_end_row.pack()
    input_end_row = tk.Entry(window)
    input_end_row.pack()
    button_generate = tk.Button(window, text = '開始生成', command=table_replace)
    button_generate.pack()
    close_button = tk.Button(window, text="Close", command= close_window)
    close_button.pack()
    lable_detail = tk.Text(window)
    lable_detail.pack()

    window.mainloop()


main()
