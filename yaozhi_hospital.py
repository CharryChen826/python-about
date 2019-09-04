# python2
# -*- coding:utf-8 -*-

import requests
import xlwt
import json
from bs4 import BeautifulSoup


url = 'https://db.yaozh.com/hmap?yaozhVersion=1.5.60'

headers = {
	'Accept':'*/*',
	'Accept-Encoding': 'gzip, deflate',
	'Accept-Language': 'zh-CN,zh;q=0.9',
	'Connection': 'keep-alive',
	'Cookie': 'kztoken=nJail6zJp6iXaJqWl29pZWhwY5Wb; his=a%3A3%3A%7Bi%3A0%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwYpST%22%3Bi%3A1%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwYpSZ%22%3Bi%3A2%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwY5Wb%22%3B%7D; acw_tc=2f624a2115675781116266711e428d42960c7fdd4048dffee20988dd3c75fc; think_language=zh-CN; PHPSESSID=6io0fqajntt0k84nonuvk8mht1; _ga=GA1.2.1771196088.1567578117; _gat=1; Hm_lvt_65968db3ac154c3089d7f9a4cbb98c94=1567578117; Hm_lpvt_65968db3ac154c3089d7f9a4cbb98c94=1567578229; kztoken=nJail6zJp6iXaJqWl29pZWhwZJSS; his=a%3A3%3A%7Bi%3A0%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwYpST%22%3Bi%3A1%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwY5Wa%22%3Bi%3A2%3Bs%3A28%3A%22nJail6zJp6iXaJqWl29pZWhwZJSS%22%3B%7D',
	'Host': 'db.yaozh.com',
	# 'Referer': 'https://db.yaozh.com/hmap?p=1&pageSize=30',
	'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36'
}


# workbook = xlwt.Workbook(encoding='utf-8')

def new_workbook():
    workbook = xlwt.Workbook(encoding='utf-8')
    return workbook


def new_sheet(wb, i):
    sheet = wb.add_sheet(u'XX数据_' + str(i))
    head = [u'序号', u'医院名称', u'等级', u'类型', u'省', u'市', u'县', u'床位数', u'医院地址']

    for h in range(len(head)):
        sheet.write(0, h, head[h])

    return sheet


def grab_data(s, e, save_path):

    # 记录错误行数
    error_page = []

    workbook = new_workbook()

    sheet_index = 1
    sheet = new_sheet(workbook, sheet_index)

    row = 1
    for p in range(s, e):
        # 超过1500行，新建一页
        if row > 1500:
            sheet_index+=1
            sheet = new_sheet(workbook, sheet_index)
            row = 1

        print '*'*400
        _url = url + '&pageSize=30&p=' + str(p)
        print _url

        try:
            response = requests.get(_url, headers=headers)
            soup = BeautifulSoup(response.content, 'html.parser')

            body = soup.body
            table = body.table
            thead = table.thead
            # print thead
            tbody = table.tbody
            # print tbody

            for tr in tbody.find_all('tr'):
                # print '-'*50

                hospital = tr.th.a.text
                tds = []
                for td in tr.find_all('td'):
                    tds.append(td.text)
                # print tds

                level = tds[0]
                _type = tds[1]
                province = tds[2]
                city = tds[3]
                county = tds[4]
                bed = tds[5]
                location = tds[6]

                sheet.write(row, 0, row)
                sheet.write(row, 1, hospital)
                sheet.write(row, 2, level)
                sheet.write(row, 3, _type)
                sheet.write(row, 4, province)
                sheet.write(row, 5, city)
                sheet.write(row, 6, county)
                sheet.write(row, 7, bed)
                sheet.write(row, 8, location)

                row+=1

        except Exception as e:
            print 'Exception in page: ' + str(p)
            error_page.append(p)
            pass

    print 'row: ' + str(row)

    workbook.save(save_path)
    print 'save path: ' + save_path

    return error_page


if __name__ == '__main__':
    
    error_ = []

    for excel_i in range(2):
        _path = '/Users/chenchaoran/Tools/XX-yaozhi_data_%s.xls' % excel_i
        print _path
        s = 100 * excel_i + 1
        e = 100 * (excel_i + 1) + 1
        print s
        print e
        
        error_page = grab_data(s, e, _path)
        error_.extends(error_page)

    print error_



