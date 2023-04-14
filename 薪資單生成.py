import openpyxl as opxl
from docx import Document
import os
import win32com.client as win32
word = win32.Dispatch('Word.Application')

def replace_word_text(word_file, old_text:str, new_text:str):#文字替換函數
    try:
        for paragraph in word_file.paragraphs:#填入所有表格中paragraphs的部分
            paragraph.text = paragraph.text.replace(old_text, new_text)
                
        for table in word_file.tables:#填入所有表格中table的部分
            for row in table.rows:
                for cell in row.cells:
                    cell.text = cell.text.replace(old_text, new_text)
    except:
        print('填入錯誤')

def get_password(path_password):#從密碼表取密碼
    try:
        print('正在取得密碼')
        password_workbook = opxl.load_workbook(path_password)
        password_sheet = password_workbook.worksheets[0]
        password_dic = {password_sheet.cell(row=pwrow, column=1).value: password_sheet.cell(row=pwrow, column=2).value
                        for pwrow in range(1, password_sheet.max_row+1)}
        print('取得密碼成功')
        return password_dic
                  
    except:
        print('取得密碼失敗')
        raise

def encrypt_files(file_path, password_dic:dict, name):#加密檔案
    try:
        doc = word.Documents.Open(file_path)
        doc.Password = str(password_dic[f'{name}'])
        doc.Save()
        doc.Close()
        print(name,'薪資單已加密完畢')
    except:
        print('加密失敗')    

try:#指定檔案位置
    path = os.path.dirname(__file__)#當前目錄
    generate_path = os.path.join(path, '生成文件')
    os.makedirs(generate_path, exist_ok=True)#如果沒有自動生成資料夾
    path_docx = os.path.join(path,'範例薪資單.docx')
    path_xlsx = os.path.join(path,'薪資表.xlsx')
    path_password = os.path.join(path,'密碼表.xlsx')
except:
    print('檢查資料夾中檔案是否正確')
    raise
try:#載入檔案
    workbook = opxl.load_workbook(path_xlsx,data_only= True) #讀取WORKBOOK並且設定data_only= True，使公式格回傳計算後結果
    sheet = workbook['薪資總表'] #選擇SHEET
    paydate = str(sheet['A2'].value)[0:11]#從表格取得日期字串
except:
    print('讀取錯誤檢查檔案EXCEL檔案')
    raise
while True:
    try:
        star_row = int(input('薪資表表頭位置: '))
        end_row = int(input('人員名單結束位置: '))
        break
    except ValueError:
        print('請輸入整數！')
try:
    password = get_password(path_password)#取得密碼函數
    for table_row in range(star_row, end_row):#指定ROW循環範圍
        word_file = Document(path_docx)
        for table_col in range (3, sheet.max_column+1):#指定COL循環範圍
            old_text = str(sheet.cell(row= star_row, column=table_col).value) #取得標題名當作範本WORD中的待替換文本
            new_value = sheet.cell(row=table_row, column=table_col).value  # 獲取欲替換單元格內容值
            new_text = str(new_value) if new_value not in [None, 0] else ""  # 將結果轉為字符串並且當結果為NONE或0時替換為空白
            replace_word_text(word_file,old_text,new_text)#調用替換函數
        name = str(sheet.cell(row=table_row, column=6).value)#設定NAME變數
        generate_file_name =  f"{paydate}薪資單-{name}.docx"
        title_text = word_file.paragraphs[0].text.replace('#發薪日期', paydate)#取範本WORD中第一段文字動態替換日期
        word_file.paragraphs[0].text = title_text#將變數存回文件中!!!
        word_file.save(os.path.join(generate_path, generate_file_name))#儲存並命名標題完成一次循環
        print(name,'薪資單生成完畢')
        #加密薪資單
        path_unprotectdocx = os.path.join(generate_path, generate_file_name)#加密文件的路徑
        encrypt_files(path_unprotectdocx, password, name)#加密函數
    word.Quit()#關掉wiv32模組
    print('薪資單均已加密完畢')
except:
    print('填入過程中出現錯誤')
    raise
