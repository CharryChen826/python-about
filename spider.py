#!python 2

import requests
import xlwt
import json

def spider_data():
    params = {
        'subject': 320,
        'year':2017,
        'start':1,
        'end':100
    }

    result = requests.post('http://top100.imicams.ac.cn/assess_2018/public/ranking/rankingAction_searchRankByCode.action?d=0.6602740263232401', params)

    return result.content


def save_to_excel(datas, save_path):
    try:
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('Hospital_Analysis')
        head = ['Rank', 'Name', 'Input', 'Output', 'Infuluence', 'Sum', 'Province']
        for h in range(len(head)):
            sheet.write(0, h, head[h])

        i = 1
        for product in datas:
            sheet.write(i, 0, product['RANK'])
            sheet.write(i, 1, product['HOSPNAME'])
            sheet.write(i, 2, product['INPUT'])
            sheet.write(i, 3, product['OUTPUT'])
            sheet.write(i, 4, product['INFLUENCE'])
            sheet.write(i, 5, product['SUM'])
            sheet.write(i, 6, product['PROVINCE'])

            i+=1
        workbook.save(save_path)
        print('File path: ' + save_path)
    except Exception:
        print('failed!')
    pass


if __name__ == "__main__":
    result = spider_data()
    print result
    
    datas = json.loads(result, encoding='utf-8')
    print datas['rows']

    _path = '/Users/chenchaoran/Tools/hospital_data.xls'
    save_to_excel(datas['rows'], _path)
