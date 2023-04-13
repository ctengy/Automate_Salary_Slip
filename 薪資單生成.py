import openpyxl as opxl
from docx import Document
import os
import win32com.client as win32
word = win32.Dispatch('Word.Application')
#指定檔案位置
try:
    path = os.getcwd()#當前工作目錄
    generate_path = os.path.join(path, '生成文件')
    os.makedirs(generate_path, exist_ok=True)#如果沒有自動生成資料夾
    path_docx = os.path.join(path,'範例薪資單.docx')
    path_xlsx = os.path.join(path,'薪資表.xlsx')
except:
    print('檢查資料夾中檔案是否正確')
    raise
#載入檔案
try:
    workbook = opxl.load_workbook(path_xlsx,data_only= True) #讀取WORKBOOK並且設定data_only= True，使公式格回傳計算後結果
    sheet = workbook.active #選擇最近保存的SHEET
    paydate = str(sheet['A2'].value)[0:11]#從表格取得日期字串
except:
    print('讀取錯誤檢查檔案EXCEL檔案')
    raise

#批量處理
try:
    #自EXCEL中取值
    for table_row in range(5, 35):#指定ROW循環範圍
        word_file = Document(path_docx)
        for table_col in range (3, 25):#指定COL循環範圍
            old_text = str(sheet.cell(row=4, column=table_col).value) #取得標題名當作範本WORD中的待替換文本
            new_value = sheet.cell(row=table_row, column=table_col).value  # 獲取欲替換單元格內容值
            new_text = str(new_value) if new_value not in [None, 0] else ""  # 將結果轉為字符串並且當結果為NONE或0時替換為空白
            #以下自範例WORD中填入
            for paragraph in word_file.paragraphs:#填入所有表格中paragraphs的部分
                paragraph.text = paragraph.text.replace(old_text, new_text)
            
            for table in word_file.tables:#填入所有表格中table的部分
                for row in table.rows:
                    for cell in row.cells:
                        cell.text = cell.text.replace(old_text, new_text)

        name = str(sheet.cell(row=table_row, column=6).value)#設定NAME變數
        title_text = word_file.paragraphs[0].text.replace('#發薪日期', paydate)#取範本WORD中第一段文字動態替換日期
        word_file.paragraphs[0].text = title_text#將變數存回文件中!!!
        word_file.save(os.path.join(generate_path, f"{paydate}薪資單-{name}.docx"))#儲存並命名標題完成一次循環
    print('薪資單生成完畢')
except:
    print('填入過程中出現錯誤')
    raise
#加密所有檔案
try:
    unencyptfiles = os.listdir(generate_path)
    for file_name in unencyptfiles:
        file_path = os.path.join(generate_path, file_name)
        doc = word.Documents.Open(file_path)
        doc.Password = "0000"
        doc.Save()
        doc.Close()
        print('已加密',file_name)
    word.Quit()
    print('薪資單加密完畢')
except:
    print('獲取檔名失敗')
    raise