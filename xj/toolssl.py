from selenium import webdriver
import xlrd
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, NoSuchFrameException

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

def deal_tele(data):
    # print(data[22], data[25], data[28])
    for index in (22, 25, 28):
        if data[index] and isinstance(data[index], str) and not data[index].strip().isdigit():
            data[index] = ''
    # print(data[22], data[25], data[28])

def main():
    br = get_explorer('https://xj.ahjygl.gov.cn/SMS.UI/Pages/Common/Login.aspx')
    br.implicitly_wait(20)
    input('手工登录完成？')
    datas = get_file_datas('res.xlsx')
    # for data in datas:
    i = 0
    total = len(datas)
    while True:
        if i >= total:
            br.quit()
            break
        data = datas[i]
        deal_tele(data)
        if not ((data[20] or data[23] or data[26]) and (data[21] or data[24] or data[27]) and (data[22] or data[25] or data[28])):
            i += 1
            continue
        print(data[4], data[6])

        try:
            br.switch_to_frame('right')
            br.switch_to_frame('UpperHalf')
            id_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_StudentIDNO"]')
            id_html.clear()
            # 输入身份证查询
            id_html.send_keys(data[6])
            query_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Search"]')
            query_html.click()

            time.sleep(3)
            # br.switch_to_default_content()
            # br.switch_to_frame('right')
            # br.switch_to_frame('LowerHalf')
            try:
                br.find_element_by_xpath('//a[@id="ctl00_ContentPlaceHolder_GridView_StudentList_ctl02_linkButton_Detail"]')
            except:
                pass
            else:
                print('已填写！')
                continue

            br.find_element_by_xpath('//a[@id="ctl00_ContentPlaceHolder_GridView_StudentList_ctl02_linkButton_Result"]').click()
            # query_html
            # input('请点击核查按钮')


            print('等待查询...')
            time.sleep(6)
            # skip = input('是否跳过（y/n）:')
            # if skip == 'y':
            #     continue
            br.switch_to_default_content()
            br.switch_to_frame('right')
            br.switch_to_frame('LowerHalf')

            WebDriverWait(br, 500).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder_DropDownList_HCJG')))

            result_html = br.find_element_by_xpath('//select[@id="ctl00_ContentPlaceHolder_DropDownList_HCJG"]')
            result_html.click()

            value = data[14]
            if value.strip() == '已关爱':
                value = 5
            elif value.strip() == '未关爱':
                value = 6
            else:
                value = 99
            if value != 5:
                result_html = br.find_element_by_xpath('//option[@value="%s"]' % value)
                result_html.click()

            value = data[15]
            if value.strip() == '否':
                result_html = br.find_element_by_xpath('//select[@id="ctl00_ContentPlaceHolder_DropDownList_LS"]')
                result_html.click()
                result_html = br.find_element_by_xpath('//option[@value="0"]')
                result_html.click()

            # 关爱情况说明
            if data[16]:
                checker_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_HCQKSM"]')
                checker_html.clear()
                checker_html.send_keys(data[16])


            checker_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_HCR"]')
            checker_html.clear()
            checker_html.send_keys(data[17])

            unit_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_SZDW"]')
            unit_html.clear()
            unit_html.send_keys(data[18])

            phnumber_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_SJHM"]')
            p = str(int(data[19])) if isinstance(data[19], float) else data[19]
            phnumber_html.clear()
            phnumber_html.send_keys(p)

            all_guardians = data[20: 29]

            if not data[20]:
                all_guardians[0] = str(data[23] or data[26])

            if not data[21]:
                all_guardians[1] = str(data[24] or data[27])

            if not data[22]:
                all_guardians[2] = data[25] or data[28]


            if not data[23]:
                all_guardians[3] = str(data[20] or data[26])

            if not data[24]:
                all_guardians[4] = str(data[21] or data[27])

            if not data[25]:
                all_guardians[5] = data[22] or data[28]


            if not data[26]:
                all_guardians[6] = str(data[20] or data[23])

            if not data[27]:
                all_guardians[7] = str(data[21] or data[24])

            if not data[28]:
                all_guardians[8] = data[22] or data[25]

            print(all_guardians)
            parent_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_FatherName"]')
            parent_html.clear()
            parent_html.send_keys(all_guardians[0])

            gx_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_FatherCompany"]')
            gx_html.clear()
            gx_html.send_keys(all_guardians[1])

            gxph_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_fatherTele"]')
            p = str(int(all_guardians[2])) if isinstance(all_guardians[2], float) else all_guardians[2]
            gxph_html.clear()
            gxph_html.send_keys(p)

            parent_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_motherName"]')
            parent_html.clear()
            parent_html.send_keys(all_guardians[3])

            gx_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_motherCompany"]')
            gx_html.clear()
            gx_html.send_keys(all_guardians[4])

            gxph_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_motherTele"]')
            p = str(int(all_guardians[5])) if isinstance(all_guardians[5], float) else all_guardians[5]
            gxph_html.clear()
            gxph_html.send_keys(p)

            parent_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_keeperName"]')
            parent_html.clear()
            parent_html.send_keys(all_guardians[6])

            gx_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_keeperAddress"]')
            gx_html.clear()
            gx_html.send_keys(all_guardians[7])

            gxph_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_LXDH"]')
            p = str(int(all_guardians[8])) if isinstance(all_guardians[8], float) else all_guardians[8]
            gxph_html.clear()
            gxph_html.send_keys(p)


            memo_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_TextBox_careText"]')
            p = data[29] if data[29] else '有'
            memo_html.clear()
            memo_html.send_keys(p)

            btn_html = br.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder_Button_Save"]')
            btn_html.click()

            # input('点击保存和弹出对话框。')
            time.sleep(0.3)
            alert = br.switch_to_alert()
            alert.accept()
            time.sleep(1)
            br.switch_to_default_content()
        except NoSuchElementException:
            print('失去响应，请重新登录！')
            br = get_explorer('https://xj.ahjygl.gov.cn/SMS.UI/Pages/Common/Login.aspx')
            # br.implicitly_wait(20)
            input('手工登录完成？')
        except NoSuchFrameException:
            print('失去响应，请重新登录！')
            br = get_explorer('https://xj.ahjygl.gov.cn/SMS.UI/Pages/Common/Login.aspx')
            # br.implicitly_wait(20)
            input('手工登录完成？')
        else:
            i += 1
    br.quit()

if __name__ == '__main__':
    main()
