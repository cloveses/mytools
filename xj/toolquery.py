from selenium import webdriver
import xlrd
import time
import xlsxwriter

# 批量查询全国学籍

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
    option.add_argument('--user-data-dir=C:\\Users\\djx\\AppData\\Local\\Google\\Chrome\\User Data\\Default') 
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
    for data in datas:
        dd = data[:]
        print(data[0])
        br.switch_to.frame('right')
        br.switch_to.frame('UpperHalf')
        id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_Name"]')
        id_html.clear()
        id_html.send_keys(data[0])

        id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_IDNO"]')
        id_html.clear()
        id_html.send_keys(data[1])

        query_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Query"]')
        query_html.click()

        time.sleep(0.2)
        br.implicitly_wait(2)
        result_html = br.find_element_by_xpath('//span[@id="ctl00_ContentPlaceHolder_Label_Result"]')
        result = result_html.text
        if '没有查询到学生相关信息，请联系学生所在学校进行查询' not in result:
            dd.append(result)
        ds.append(dd)
        br.implicitly_wait(30)
        # br.switch_to_default_content()
        # br.switch_to.frame('right')
        # br.switch_to.frame('LowerHalf')
        # query_html = br.find_element_by_xpath('//a[@id="ctl00_ContentPlaceHolder_GridView_StudentList_ctl02_linkButton_Result"]')
        # query_html.click()
        br.switch_to_default_content()
        # input('回车继续')
    save_datas_xlsx('rr.xlsx', ds)
    br.close()


if __name__ == '__main__':
    main()
