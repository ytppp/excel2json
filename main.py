import os
import io
import json
import sys
from openpyxl.reader.excel import load_workbook

def app_path():
    if hasattr(sys, 'frozen'):
        return os.path.dirname(sys.executable)  # 使用pyinstaller打包后的exe目录
    return os.path.dirname(__file__)  # 没打包前的py目录

def excel2json(excel_file, path):
    wb = load_workbook(excel_file) # 加载excel表格
    for sheet in wb.worksheets: # wb.worksheets: 获取所有工作表对象
        print(sheet.title)
        result = {}
        lang_list = []
        # 生成语种列表
        for column in range(sheet.max_column):
            # 排除 key
            if column > 0:
                lang = sheet.cell(1, column + 1).value
                lang_list.append(lang)
                result[lang] = {}

        for row in range(sheet.max_row): # sheet.max_row: 获取行数
            # 排除表头
            if row > 0:
                for lang in lang_list:
                    key = sheet.cell(row + 1, 1).value
                    val = sheet.cell(row + 1, lang_list.index(lang) + 2).value
                    if val:
                        result[lang][key] = val
        save_json_file(sheet.title, path, lang_list, result)
    wb.close()



def save_json_file(customer, path, lang_list, result):
    customer_dir = f"{path}\{customer}"
    if os.path.exists(customer_dir) == False:
        os.makedirs(customer_dir)
    for lang in lang_list:
        file = io.open(f"{customer_dir}\{lang}.json", 'w', encoding='utf-8')
        # 把对象转化为json对象
        # indent: 参数根据数据格式缩进显示，读起来更加清晰
        # ensure_ascii = True：默认输出ASCII码，如果把这个该成False, 就可以输出中文。
        txt = json.dumps(result[lang], indent=2, ensure_ascii=False)
        file.write(txt)
        file.close()
    print("输出成功")

if __name__ == '__main__':
    excel = r'.\lang.xlsx'
    trans = r'.\trans'
    excel_path = input("请输入 excel 文件所在的路径（支持相对路径，不填写则默认当前文件夹）：")
    trans_path = input("请输入翻译文件输出路径（不填写则默认当前文件夹）：")
    if excel_path:
        excel = excel_path
    if trans_path:
        trans = trans_path
    app_path()
    excel2json(excel, trans)
    input('Press Enter to exit...')
