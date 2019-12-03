import xlrd
import sqlite3
import sys
import re
import os
import random

# 体质健康测试数据处理
'''
身高80－250
体重14－200
肺活量500－9999
50米跑：5.0－20.0
立定跳远：50－400
坐位体前屈－30－40
仰卧起坐0－99
引体向上0－99
视力3.5－5.3 低3.0 为0
串镜：视力大于5.0，为0， 低于5.0 “-1”代表正片下降、负片上升。其他情况请录入“2”。如条件不允许，未测试者录入“9”
左/右眼屈光不正：以“0”代表正常，以“1”代表近视，以“2”代表远视，以“3”代表其他（疾病等其他原因），如条
件不允许，未测试者录入“9”
'''

def get_file_datas(filename,row_deal_function=None,grid_end=0,start_row=1):
    """start_row＝1 有一行标题行；gred_end=1 末尾行不导入"""
    """row_del_function 为每行的数据类型处理函数，不传则对数据类型不作处理 """
    # names = data.sheet_names()
    # table = data.sheet_by_name(sheet_name)

    wb = xlrd.open_workbook(filename)
    # ws = wb.sheets()[0]
    names = wb.sheet_names()
    datas = []
    for name in names:
        ws = wb.sheet_by_name(name)
        nrows = ws.nrows
        for i in range(start_row,nrows-grid_end):
            row = ws.row_values(i)
            # print(row)
            if row_deal_function:
                row = row_deal_function(row)
            datas.append(row)
    return datas

# clear integer
def clear_int(data, minimize, maxmize, info=''):
    if data.endswith('.0'):
        data = data[:-2]
    if data:
        data = ''.join(re.split(r'\s+', data.strip()))
    else:
        return random.randint(minimize, maxmize)
    try:
        data = int(data)
    except:
        print(info)
        return random.randint(minimize, maxmize)
    else:
        return data

# clear_float
def clear_float(data, minimize, maxmize, info=''):
    data = ''.join(re.split(r'\s+', data))
    try:
        data = float(data)
    except:
        return data
    else:
        return random.random() * (maxmize - minimize) + minimize

def clear_duration(data, minute_minimize, minute_maxmize, info=''):
    data = re.split(r'\D', data.strip())
    data = [d for d in data if d]
    if len(data) > 2:
        data = data[:2]
    elif len(data) == 1:
        data.append('00')
    elif len(data) < 1:
        data = ['00', '00']
    if int(data[-1]) < 10:
        data[-1] = '0' + data[-1]
    if int(data[0]) < minute_minimize:
        data[0] = str(random.randint(minute_minimize, minute_maxmize))
    if int(data[0]) > minute_maxmize:
        data[0] = str(minute_maxmize)
    if int(data[-1]) >= 60:
        data[-1] = 59
    return "'".join(data)


# 以下对体测和视力数据清理需要应用以上三个方法进行重构

def clean_phdata(row):
    row = [str(r).strip() for r in row]
    # height 5
    if row[0].strip().endswith('.0'):
        row[0] = row[0].strip()[:-2]
    # weight 6
    if row[1].strip().endswith('.0'):
        row[1] = row[1].strip()[:-2]
    if row[2].strip().endswith('.0'):
        row[2] = row[2].strip()[:-2]
       
    try:
        if row[5].strip():
            height = int(float(row[5].strip()))
            row[5] = str(height)
    except:
        print('height error!', row)

    try:
        if row[6].strip():
            weight = int(float(row[6].strip()))
            weight = str(weight)
            if len(weight) == 3 and int(weight[0]) > 1 :
                weight = '1' + weight[1:]
            row[6] = str(weight)
    except:
        print('weight error!', row)

    try:
        if row[7].strip():
            lung = int(float(row[7].strip()))
            if lung < 500 or lung >9999:
                print('lung range error(500-9999)', row)
            else:
                row[7] = str(lung)
    except:
        print('lung error!', row)

    try:
        if row[8].strip():
            duration = float(row[8].strip())
            if duration < 5.0 or duration > 20.0:
                print('run50 range error(5.0-20.0)!', row)
    except:
        duration = re.split(r'\D',row[8].strip())
        if len(duration) < 1 or len(duration) > 2:
            print('run50 error!', row)
        elif len(duration) == 1:
            duration = float(duration[0])
        elif len(duration) == 2:
            duration = float('.'.join(duration))
        if duration < 5.0 or duration > 20.0:
            print('run50 range error(5.0-20.0)!', row)
        else:
            row[8] = str(duration)

    try:
        if row[9].strip():
            distance = float(row[9].strip())
            if distance < 50 or distance > 400:
                print('jump range error(50-400)!', row)
    except:
        print('jump error!', row)

    try:
        if row[10].strip():
            bend = float(row[10].strip())
            if bend < -30 or bend > 40:
                print('bend error!(-30-40)', row)
    except:
        print('bend error!', row)

    try:
        if row[11].strip():
            duration = re.split(r'\D', row[11].strip())
            duration = [d for d in duration if d]
            if len(duration) < 1 or len(duration) > 2:
                print('run800 error!', row)
            elif len(duration) == 1:
                row[11] = duration[0] + "'00"
            else:
                if int(duration[-1]) >= 60:
                    duration[-1] = '0'+duration[-1][0]
                row[11] = "'".join(duration)
    except:
        print('run800 error!', row)

    try:
        if row[12].strip():
            duration = re.split(r'\D', row[12].strip())
            duration = [d for d in duration if d]
            if len(duration) < 1 or len(duration) > 2:
                print('run1000 error!', row)
            elif len(duration) == 1:
                row[12] = duration[0] + "'00"
            else:
                if int(duration[-1]) >= 60:
                    duration[-1] = '0'+duration[-1][0]
                row[12] = "'".join(duration)
    except:
        print('run1000 error!', row)

    try:
        if row[13].strip():
            lying = int(float(row[13].strip()))
            if lying < 0 or lying >99:
                print('lying error(0-99)', row)
            else:
                row[13] = str(lying)
    except:
        print('lying error!', row)

    try:
        if row[14].strip():
            bodyup = int(float(row[14].strip()))
            if bodyup < 0 or bodyup >99:
                print('bodyup error(0-99)', row)
            else:
                row[14] = str(bodyup)
    except:
        print('bodyup error!', row)

    return row


def clean_eyedata(row):
    row = [str(r).strip() for r in row]
    if row[1].strip().endswith('.0'):
        row[1] = row[1].strip()[:-2]
    if row[0].strip().endswith('.0'):
        row[0] = row[0].strip()[:-2]

    if row[6].strip():
        try:
            r = float(row[6].strip())
            if r == 0:
                row[6] = '0.0'
            elif r > 5.3 or r < 3.0:
                print('left eye range error!', row)
            else:
                row[6] = str(r)
        except:
            print('error:',row)
    else:
        print('null !')


    if row[7].strip():
        try:
            r = float(row[7].strip())
            if r == 0:
                row[7] = '0.0'
            elif r > 5.3 or r < 3.0:
                print('left eye range error!', row)
            else:
                row[7] = str(r)
        except:
            print('error:',row)
    else:
        print('null !')

    r = lambda x:str(int((float(x)))) if x else '9'
    for i in range(8,12):
        row[i] = r(row[i].strip())
    return row


def clean_info_data(row):
    row = [str(r).strip() for r in row]
    r = lambda x:x[:-2] if x.endswith('.0') else x
    seqs = (0, 1, 3, 5)
    for seq in seqs:
        row[seq] =r(row[seq].strip())
    row[-1] = ''.join(re.split(r'\s+', row[-1].strip()))
    return row



if __name__ == '__main__':
    phdatas = get_file_datas('2019phdataclean.xlsx', clean_phdata)
    eyedatas = get_file_datas('2019eye.xlsx', clean_eyedata)
    studdatas = get_file_datas('studinfo.xls', clean_info_data)

    phdatas_dict = {}
    for data in phdatas:
        key = (data[2],data[3])
        if key in phdatas_dict:
            print('phdatas repeat:', data)
        else:
            phdatas_dict[key] = data

    # with open('t.txt', 'w', encoding='utf-8') as f:
    #     for data in studdatas:
    #         f.write(str(data))
    #         f.write('\n')

    eyedatas_dict = {}
    i = 0
    for data in eyedatas:
        if not data or len(data) < 5:
            print('eyedatas error', data)
            continue
        key = (data[0],data[4])
        if key in eyedatas_dict:
            print('eyedatas repeat:', key, data)
            print(eyedatas_dict[key])
            if i >= 9:
                break
            i += 1
        else:
            eyedatas_dict[key] = data


    studdatas_dict = {}
    for data in studdatas:
        if len(data) <= 4:
            continue
        key = (data[0],data[4])
        if key in studdatas_dict:
            print('studdatas repeat:', data)
        else:
            studdatas_dict[key] = data


    results = []
    for key, stud in studdatas_dict.items():
        row = []

        if key in phdatas_dict:
            phdata = phdatas_dict[key]
            row.append(phdata[0])
            row.extend(stud[:7])
            row.append(stud[-1])
            row.extend(phdata[5:])
        else:
            print('no ph data:', key)
            row.append('')
            row.extend(stud[:7])
            row.append(stud[-1])
            row.extend(['', ] * 10)

        if key in eyedatas_dict:
            eyedata = eyedatas_dict[key]
            row.extend(eyedata[6:12])
        else:
            row.extend(random.choice(eyedatas)[6:12])
            print('no eye data:', key)

        if row[6] == '1':
            if row[15] or row[17]:
                print('man gender do not match phdata!', row)
                print('src:','row[15]=', row[15],  'row[17]=',  row[17],bool(row[15] or row[17]))
                row[15] = ''
                row[17] = ''
        elif row[6] == '2':
            if row[16] or row[18]:
                print('woman gender do not match phdata!', row)
                print('src', 'row[16]=', row[16], 'row[18]=', bool(row[16] or row[18]))
                row[16] = ''
                row[18] = ''

        results.append(row)

    null_results = [d for d in results if not d[9]]
    man_results = [d for d in results if d[9] and d[6] == '1']
    woman_results = [d for d in results if d[9] and d[6] == '2']

    # print(null_results[:3])

    for null_result in null_results:
        if null_result[6] == '1':
            data = random.choice(man_results)
        else:
            data = random.choice(woman_results)
        null_result[9: 19] = data[9: 19]

    # print(null_results[:3])
    results = []
    results.extend(man_results)
    results.extend(woman_results)
    results.extend(null_results)

    with open('results.csv', 'w', encoding='gbk') as f:
        for data in results:
            f.write(','.join(data))
            f.write('\n')
