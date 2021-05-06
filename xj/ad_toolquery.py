from selenium import webdriver
import xlrd
import time
import xlsxwriter

# 学籍批量高级查询

def save_datas_xlsx(filename,datas):
    #将一张表的信息写入电子表格中XLSX文件格式
    import xlsxwriter
    w = xlsxwriter.Workbook(filename)
    w_sheet = w.add_worksheet('sheet1')
    for rowi,row in enumerate(datas):
        for coli,celld in enumerate(row):
            w_sheet.write(rowi,coli,celld)
    w.close()

##C:\Users\djx\AppData\Local\Google\Chrome\User Data\Default

def get_explorer(purl):
    option = webdriver.ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    # option.add_argument('--user-data-dir=C:\\Users\\djx\\AppData\\Local\\Google\\Chrome\\User Data\\Default') 
    br = webdriver.Chrome(chrome_options=option)
    br.get(purl)
    return br

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
        datas.append(row)
    return datas

def main():
    ds = []
    br = get_explorer('https://xj.ahjygl.gov.cn/SMS.UI/Pages/Common/Login.aspx')
    # br.implicitly_wait(20)
    input('手工登录完成？')
    datas = get_file_datas('in.xlsx')
    try:
        for data in datas:
            dd = data[:]
            print(data[0])
            br.switch_to.frame('right')
            br.switch_to.frame('UpperHalf')

            id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_CheckBox4"]')
            id_html.click()


            id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_IDNO"]')
            id_html.clear()
            id_html.send_keys(data[1])

            # 点击查询
            query_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Query"]')
            query_html.click()

            # 等待结果
            time.sleep(30)
            br.implicitly_wait(5)
            br.switch_to.default_content()
            br.switch_to.frame('right')
            br.switch_to.frame('UpperHalf')

            result_html = br.find_element_by_xpath('//table[@id="ctl00_ContentPlaceHolder_GridView_StuList"]')
            result_texts = result_html.text.split()
            if len(result_texts) > 10:
                trs = result_html.find_elements_by_xpath('.//tr')[1:]
                result_texts = []
                for tr in trs:
                    row = []
                    tds = tr.find_elements_by_xpath('.//td')
                    for td in tds[6:]:
                        if td.text:
                            row.append(td.text)
                    result_texts.append(','.join(row))
                if result_texts:
                    result = ';'.join(result_texts)
                else:
                    result = ''
                dd.append(result)
            ds.append(dd)

            time.sleep(0.2)
            br.implicitly_wait(2)

            # 返回
            br.find_element_by_xpath('//input[@id="Button_Back"]').click()
            br.implicitly_wait(8)
            br.switch_to.default_content()

    except Exception as e:
        print(e)
        if ds:
                save_datas_xlsx('rr.xlsx', ds)
    save_datas_xlsx('rr.xlsx', ds)
    br.close()


if __name__ == '__main__':
    main()
