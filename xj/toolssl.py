from selenium import webdriver
import xlrd
import time

def get_explorer(purl):
    br = webdriver.Firefox()
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
    br = get_explorer('https://xj.ahjygl.gov.cn/SMS.UI/Pages/Common/Login.aspx')
    # br.implicitly_wait(20)
    input('手工登录完成？')
    datas = get_file_datas('in.xls')
    for data in datas:
        print(data[3])
        br.switch_to_frame('right')
        br.switch_to_frame('UpperHalf')
        id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_StudentIDNO"]')
        id_html.clear()
        id_html.send_keys(data[5])
        query_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Search"]')
        query_html.click()

        # time.sleep(0.5)
        # br.switch_to_default_content()
        # br.switch_to_frame('right')
        # br.switch_to_frame('LowerHalf')
        # query_html = br.find_element_by_xpath('//a[@id="ctl00_ContentPlaceHolder_GridView_StudentList_ctl02_linkButton_Result"]')
        # query_html.click()
        input('请点击核查按钮')

        skip = input('是否跳过（y/n）:')
        if skip == 'y':
            continue
        br.switch_to_default_content()
        br.switch_to_frame('right')
        br.switch_to_frame('LowerHalf')
        result_html = br.find_element_by_xpath('//select[@id="ctl00_ContentPlaceHolder_DropDownList_HCJG"]')
        result_html.click()

        value = data[14]
        if value.strip() == '非留守':
            value = 4
        else:
            value = 3
        if value == 4:
            result_html = br.find_element_by_xpath('//option[@value="%s"]' % value)
            result_html.click()

        checker_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_HCR"]')
        checker_html.send_keys(data[16])
        unit_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_SZDW"]')
        unit_html.send_keys(data[17])
        phnumber_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_SJHM"]')
        p = str(int(data[18])) if isinstance(data[18], float) else data[18]
        phnumber_html.send_keys(p)
        parent_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_JZXM"]')
        parent_html.send_keys(data[19])
        gx_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_YXSGX"]')
        gx_html.send_keys(data[20])
        gxph_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_LXDH"]')
        p = str(int(data[21])) if isinstance(data[21], float) else data[21]
        gxph_html.send_keys(p)
        memo_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_HCQKSM"]')
        memo_html.send_keys(data[15])

        btn_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Save"]')
        btn_html.click()

        # input('点击保存和弹出对话框。')
        time.sleep(0.3)
        alert = br.switch_to_alert()
        alert.accept()
        time.sleep(1)
        br.switch_to_default_content()


if __name__ == '__main__':
    main()