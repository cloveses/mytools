from pony.orm import *

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

@db_session
def dump(table_obj,column_titles,all_keys,according_key=None):
    """
    导出数据数据到xlsx文件
    table_obj ORM对象
    column_titles 表的标题列表
    all_keys 导出的字段列表
    according_key 分类依据字段名(默认为None时全部数据导出在results.xlsx文件中)
    """
    if according_key:
        according_values = select(getattr(e,according_key) for e in table_obj)
        for according_value in according_values:
            datas = [column_titles,]
            table_objects = select(s for s in table_obj).filter(lambda e:getattr(e,according_key) == according_value)
            for table_object in table_objects:
                row = []
                for key in all_keys:
                    row.append(getattr(table_object,key))
                datas.append(row)
            if datas:
                save_datas_xlsx(according_value+'.xlsx',datas)
    else:
        datas = [column_titles,]
        table_objects = select(s for s in table_obj
            if getattr(e,according_key) == according_key)
        for table_object in table_objects:
            row = []
            for key in all_keys:
                row.append(getattr(table_object,key))
            datas.append(row)
        if datas:
            save_datas_xlsx('results.xlsx',datas)
