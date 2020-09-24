import os
import xlrd
import xlsxwriter

'''
依据某数据中的某列，到大量数据文件中查找对应某列的所有数据
将要查找到的整行添加到依据文件的数据行中
'''

# 对应不同的项目应修改此函数
def deal_row(row):
    if not isinstance(row[6], str):
        print(row, 'error')
    return row

def get_file_datas(filename,row_deal_function=None,grid_end=0,start_row=1):
    """start_row＝1 有一行标题行；gred_end=1 末尾行不导入"""
    """row_del_function 为每行的数据类型处理函数，不传则对数据类型不作处理 """
    wb = xlrd.open_workbook(filename)
    ws = wb.sheets()[0]
    nrows = ws.nrows
    datas = []
    for i in range(start_row,nrows-grid_end):
        row = ws.row_values(i)
        # print(row)
        if row_deal_function:
            row = row_deal_function(row)
        datas.append(row)
    return datas

def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

def main(src_file, look_up_dir, src_col, look_up_col):
    # 查找数据依据文件 src_file
    # 被提取数据文件所在目录
    # 查找数据依据文件中依据列序号
    # 被提取数据文件中依据列序号
    src_datas  = get_file_datas(src_file)

    datas = []
    for file in get_files(look_up_dir):
        ds = get_file_datas(file, deal_row)
        datas.extend(ds)
    # print(datas)
    datas = {data[look_up_col]:data for data in datas}

    new_datas = []
    for src_data in src_datas:
        print()
        if src_data[src_col] in datas:
            d = src_data[:]
            d.extend(datas[src_data[src_col]])
            new_datas.append(d)
    save_datas_xlsx('res.xlsx', new_datas)

if __name__ == '__main__':
    main('src.xlsx', './datas', 1, 1)