import os
import xlrd
import xlsxwriter


def get_rows_data(filename):
     #既可以打开xls类型的文件，也可以打开xlsx类型的文件
    #w = xlrd.open_workbook('text.xls')
    #w = xlrd.open_workbook('acs.xlsx')
    datas = []
    w = xlrd.open_workbook(filename)
    ws = w.sheets()[0]
    nrows = ws.nrows
    for i in range(nrows):
        data = ws.row_values(i)
        datas.append(data)
    #    print(datas)
    return datas

def get_cols_data(filename,headline_row_num=0):
    # 按列获取数据
    datas = []
    w = xlrd.open_workbook(filename)
    ws = w.sheets()[0]
    ncols = ws.ncols
    for i in range(ncols):
        data = ws.col_values(i)[headline_row_num:]
        datas.append(data)
    #    print(datas)
    return datas

def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files
    
def countit(datas):
    # 对(datas)列表中的数据项进行分类统计并输出
    from collections import Counter
    c = Counter(datas)
    for k,v in c.items():
        print(k,v,sep='\t')
    return c

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

def save_data_sheets_xlsx(filename,datass,sheetnames):
    #将对应表中的信息列表写入对应的work_sheet中
    #例：datass=[sheet1表的信息列表，sheet2表的信息列表],sheetnames=['sheet1','sheet2']
    w = xlsxwriter.Workbook(filename)
    sheets = [w.add_sheet(sheetname) for sheetname in sheetnames]
    for ws,datas in zip(sheets,datass):
        for rowi,row in enumerate(datas):
            for coli,celld in enumerate(row):
                ws.write(rowi,coli,celld)
    w.close()

def summary_col(filename,col_seq_num,res_filename='sum.xlsx'):
    # 统计指定电子表格文件中的指定序号列到指定的文件中
    # filename 指定电子表格文件
    # col_seq_num 列序号从0开始
    # res_filename 统计结果存放文件名
    cols_data = get_data_cols(filename)
    cols_data = cols_data[col_seq_num]
    res = countit(cols_data)
    res = [[k,v] for k,v in res.items()]
    save_datas_xlsx(res_filename,res)

def merge_files_data(mydir,res_filename,headline_rows=0):
    # 合并指定目录(mydir)下的分表数据到一个电子表格文件(res_filename)中的一张表中
    if not os.path.exists(mydir):
        print('Directory is not exist.')
        return
    filenames = get_files(mydir)
    datass = []
    for filename in filenames:
        datas = get_data(filename)
        datass.extend(datas)
    if datass:
        if len(datass) > 1 and headline_rows > 0:
            tails = datass[1:]
            datass = datass[:1]
            if headline_rows > 0:
                no_headline = [data[headline_rows:] for data in tails]
            datass.extend(tails)
        save_datas_xlsx(res_filename,datass)

if __name__ == '__main__':
    pass
