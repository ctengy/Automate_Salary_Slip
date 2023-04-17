import openpyxl as opxl
from docx import Document
from docx.oxml.ns import qn #調整中文字體
from docx.shared import  Pt #調整字體大小用
import os
import io
import msoffcrypto as mso
import win32com.client as win32

def loadencrypted_excel(encrypted_filename:str):#有密碼的話

    decrypted = io.BytesIO()
    password = input('請輸入密碼\n')
    with open(encrypted_filename, "rb") as f:
        file = mso.OfficeFile(f)
        file.load_key(password= password)  # Use password
        file.decrypt(decrypted)
    return decrypted   
      
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
    except KeyError:
        pass
    except:
        print('加密失敗')    
        raise

try:#指定檔案位置
    soure_file = input('轉換檔案名\n')
    path = os.path.dirname(__file__)#當前目錄
    generate_path = os.path.join(path, f'{soure_file}生成文件')
    os.makedirs(generate_path, exist_ok=True)#如果沒有自動生成資料夾
    path_docx = input('輸入範例word\n')
    path_xlsx = soure_file
    print(path_xlsx)
    path_password = input('密碼表\n')
except:
    print('檢查資料夾中檔案是否正確')
    input("按下 Enter 鍵以結束程式")
    raise

try:#載入檔案
    workbook = opxl.load_workbook(path_xlsx,data_only= True) #讀取WORKBOOK並且設定data_only= True，使公式格回傳計算後結果
    sheet = workbook['薪資總表'] #選擇SHEET
    paydate = str(sheet['A2'].value)[0:11]#從表格取得日期字串
except:
    try:
        decryptedwb = loadencrypted_excel(soure_file)
        workbook = opxl.load_workbook(decryptedwb,data_only= True) #讀取WORKBOOK並且設定data_only= True，使公式格回傳計算後結果
        sheet = workbook['薪資總表'] #選擇SHEET
        paydate = str(sheet['A2'].value)[0:11]
    except:
        print('讀取錯誤\n檢查EXCEL檔案或密碼')
        input("按下 Enter 鍵結束")
        raise

while True:#使用者指定資料
    try:
        star_row = int(input('薪資表表頭位置: '))
        end_row = int(input('人員名單結束位置: '))
        input("關閉excel檔案後 按下 Enter 鍵開始生成")
        break
    except ValueError:
        print('請輸入整數！')
        input("按下 Enter 鍵繼續")

try:#取值並填入
    password = get_password(path_password)#取得密碼函數
    word = win32.Dispatch('Word.Application')
    for table_row in range(star_row+1, end_row+1):#指定ROW循環範圍
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
        #調整標題格式
        title_text_style = word_file.paragraphs[0].runs[0]
        title_text_style.font.name = "Arial"#不知道是因為有數字還是怎樣，必須要加上這行font
        title_text_style.bold = True
        title_text_style._element.rPr.rFonts.set(qn('w:eastAsia'),'微軟正黑體')
        title_text_style.font.size = Pt(14)#一樣沒font不行
        word_file.save(os.path.join(generate_path, generate_file_name))#儲存並命名標題完成一次循環
        print(name,'薪資單生成完畢')
        #加密薪資單
        path_unprotectdocx = os.path.join(generate_path, generate_file_name)#加密文件的路徑
        encrypt_files(path_unprotectdocx, password, name)#加密函數
    word.Quit()
    print('薪資單均已加密完畢')
    input("按下 Enter 鍵繼續")
except:
    print('填入過程中出現錯誤')
    input("按下 Enter 鍵以結束程式")
    raise
