import os
import xlrd
from models import *


def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files

@db_session
def gath_data(my_dir,table_obj,column_names,column_types=None,column_num=-1,start_row=1,grid_end=0):
    """
    my_dir 导入数据文件所在目录；
    table_obj 自定义的ORM类；
    column_names 列名称列表;
    column_types 各列数据类型(默认是字符串) 例：{'weight':float}
    column_num 总列数，－1表示导入所有列；
    start_row＝1 有一行标题行；
    gred_end=1 末尾行导入;
    """

    files = get_files(my_dir)
    for file in files:
        wb = xlrd.open_workbook(file)
        ws = wb.sheets()[0]
        nrows = ws.nrows
        for i in range(start_row,nrows-grid_end):
            datas = ws.row_values(i)
            if column_num != -1:
                datas = datas[:column_num]
            datas = {k:v for k,v in zip(column_names,datas) if v}
            for k,v in datas.items():
                if isinstance(v,float):
                    datas[k] = str(int(v))
                elif isinstance(v,int):
                    datas[k] = str(v)

            # 以下为依据参数column_types类型转换代码
            if column_types:
                for k,t in column_types.items():
                    if isinstance(t,float) and not isinstance(datas[k],float):
                        if datas[k].strip().replace('.','',1).isdigit():
                            datas[k] == float(datas[k])
                    elif isinstance(t,int) and not isinstance(datas[k],int):
                        if isinstance(datas[k],str):
                            if datas[k].strip().isdigit():
                                datas[k] = int(datas[k])
                        elif isinstance(datas[k],float):
                            datas[k] = int(datas[k])

            table_obj(**datas)

if __name__ == '__main__':
    column_names = ('seq', 'checkResult', 'city', 'xian', 'zhen','cun',
        'name','idcode', 'sex','health', 'sch','grade','sclass',
        'statu','hoster','hosterTel','sq','sr',
        'ss','st','su','sv','sw','sx','memo')

    gath_data('result',Stud,column_names) # 末尾行无多余数据
