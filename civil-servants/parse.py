# -*- coding: utf-8 -*-

import os
import csv
import xlrd
from lxml import etree

NS_SERVANTS_TITLE_0 = ["年別及性別", "Year & Sex", "總計 Grand Total", "性別比率 Rate of Sex(%)", "中央各機關 Central Government Agency",
                  "臺灣各縣市機關 Local County(City) Agency", "新北市各機關 New Taipei City Agency", "台北市各機關 Taipei City Agency", 
                  "臺中市各機關 Taichung City Agency", "臺南市各機關 Tainan City Agency",
                  "高雄市各機關 Kaohsiung City Agency", "金門縣,連江縣各機關 Kinmen & Lienchiang County Agency",
                  "行政機關 Administrative Agency", "公營事業機構 Public Enterprise Organization",
                  "衛生醫療機構 Hygiene & Medical Service Organization", "公立學校(職員) Public School(Staff)"]

NS_SERVANTS_TITLE_1 = ["官等別", "Rank", "總計 Grand Total", "中央各機關 Central Government Agency",
                  "臺灣各縣市機關 Local County(City) Agency", "新北市各機關 New Taipei City Agency", "台北市各機關 Taipei City Agency", 
                  "臺中市各機關 Taichung City Agency", "臺南市各機關 Tainan City Agency",
                  "高雄市各機關 Kaohsiung City Agency", "金門縣,連江縣各機關 Kinmen & Lienchiang County Agency"]

NS_SERVANTS_TITLE_2 = ['性別及機關別', 'Sex & Agency', '總計 Grand Total', '18-25歲 Years',
                    '25-30歲 Years', '30-35歲 Years', '35-40歲 Years', '40-45歲 Years',
                    '45-50歲 Years', '50-55歲 Years', '55-60歲 Years', '60-65歲 Years',
                    '65歲以上 65 Years and over', '平均年齡(歲) Average Age (Year)',
                    '男性 Male', '女性 Female']

NS_SERVANTS_TITLE_3 = ['性別及機關別', 'Sex & Agency', '總計 Grand Total', '計 Total',
                    '博士 Ph.D. Degree', '碩士 M.A. Degree', '大學 University',
                    '專科 College', '高中(高職) High (Vocational) School',
                    '國中初中初職以下 Junior High School & Under']

NS_SERVANTS_TITLE_4 = ['性別及機關別', 'Sex & Agency', '總計 Grand Total', '計 Total',
                    '高等考試 Senior Examination', '普通考試 Junior Examination',
                    '初等考試 Elementary Examination', '特種考試 Special Examination',
                    '升等考試 Rank Promotion Examination', '其他考試 Other Examination',
                    '依其他法令進用 Appointed by Other Decree']

NS_SERVANTS_TITLE_5 = ['性別及機關別', 'Sex & Agency', '總計 Grand Total', '0-5年 Years',
                    '5-10年 Years', '10-15年 Years', '15-20年 Years', '20-25年 Years',
                    '25-30年 Years', '30年以上 30 Years & Over', '平均年資(年) Average Seniority(Year)']

NS_SERVANTS_TITLE_6 = ['性別及年齡別', 'Sex & Age', '總計 Grand Total', '0-5年 Years',
                    '5-10年 Years', '10-15年 Years', '15年 Years', '16年 Years',
                    '17年 Years', '18年 Years', '19年 Years', '20年 Years',
                    '21年 Years', '22年 Years', '23年 Years', '24年 Years',
                    '25年 Years', '26年 Years', '27年 Years', '28年 Years',
                    '29年 Years', '30年以上 30 Years & Over']

NS_SERVANTS_TITLE_7 = ['機關別', 'Agency', '總計 Grand Total', '應銓敘(人) Qualification Should be Screened (person)',
                    '應銓敘(百分比) Qualification Should be Screened (%)',
                    '不必銓敘(人) Qualification Should not be Screened (person)',
                    '不必銓敘(百分比) Qualification Should not be Screened (%)']

NS_SERVANTS_TITLE_8 = ['職系別', 'Series', '統計 Grand Total', '男性 Male', '女性 Female',
                    '中央政府 Central Government Agency', '地方政府 Local Government Agency']

NS_SERVANTS_TITLE_9 = ['縣市別', 'County(City)', '總計 Grand Total', '男性 Male', '女性 Female',
                    '行政機關 Administrative Agency', '計 Total', '生產事業 Production Enterprise',
                    '交通事業 Transportation Enterprise', '金融事業 Financial Enterprise',
                    '衛生醫療機構 Hygien & Medical Service Org.', '公立學校(職員) Public School (Staff)']

NS_SERVANTS_TITLE_10 = ['縣市別', 'County(City)', '總計 Grand Total', '計 Total',
                    '博士 Ph.D. Degree', '碩士 M.A. Degree', '大學 University',
                    '專科 College', '高中(高職) High (Vocational) School',
                    '國中初中初職以下 Junior High School & Under']


NS_SERVANTS_TITLE_11 = ['縣市別', 'County(City)', '總計 Grand Total', '計 Total',
                    '高等考試 Senior Examination', '普通考試 Junior Examination',
                    '初等考試 Elementary Examination', '特種考試 Special Examination',
                    '升等考試 Rank Promotion Examination', '其他考試 Other Examination',
                    '依其他法令進用 Appointed by Other Decree']

NS_SERVANTS_TITLE_12 = ['縣市別', 'County(City)', '總計 Grand Total', '18-25歲 Years',
                    '25-30歲 Years', '30-35歲 Years', '35-40歲 Years', '40-45歲 Years',
                    '45-50歲 Years', '50-55歲 Years', '55-60歲 Years', '60-65歲 Years',
                    '65歲以上 65 Years and over', '平均年齡(歲) Average Age (Year)',
                    '平均年資(年) Average Seniority (Year)']

NS_SERVANTS_TITLE_13 = ['性別與機關別', 'Sex & Agency', '總計 Grand Total', '本機關內調動 Transfer in Original Agency 計 Total',
                     '調升 Promotion', '平調 General', '降調 Degrade', '新進人員(人)  Newly-appointed Staff (Person) 計 Total',
                     '考試及及格分發 Assignment of Passing Examination', '他機關調進 Transferred From Others',
                     '其他任用資格 Other Qualification', '調至他機關 Transferred To Other Agency',
                     '辭職(人) Resignation (Person)']

NS_SERVANTS_TITLE = [NS_SERVANTS_TITLE_0, NS_SERVANTS_TITLE_1, NS_SERVANTS_TITLE_2,
                     NS_SERVANTS_TITLE_3, NS_SERVANTS_TITLE_4, NS_SERVANTS_TITLE_5,
                     NS_SERVANTS_TITLE_6, NS_SERVANTS_TITLE_7, NS_SERVANTS_TITLE_8,
                     NS_SERVANTS_TITLE_9, NS_SERVANTS_TITLE_10, NS_SERVANTS_TITLE_11,
                     NS_SERVANTS_TITLE_12, NS_SERVANTS_TITLE_13]

NS_SERVANTS_SHEET_NAME =  ['全國公務人員人數按年別分', '全國公務人員人數按官等分',
                           '全國公務人員人數按年齡分', '全國公務人員人數按教育程度分',
                           '全國公務人員 人數按考試種類分', '全國公務人員人數按年資分',
                           '全國公務人員人數按年齡與年資分',  '全國公務人員人數按銓敘狀況分',
                           '全國公務人員人數按職系分', '全國公務人員人數按縣市別分',
                           '全國公務人員人數按縣市別及教育程度分', '全國公務人員人數按縣市別及考試種類分',
                           '全國公務人員人數按縣市別及年齡分', '全國公務人員人事異動狀況按機關別分']


def to_csv(file_path, output_path, by_index=None, by_name=None):
    if not by_index and not by_name:
        return ""
    
    if not file_path:
        wb = xlrd.open_workbook(file_path)
    else:
        wb = file_path
        
    if by_index:
        sh = wb.sheet_by_index(by_index)
        cv = open("%s/%s.csv" % (output_path, by_index), "w", newline="")
    elif by_name:
        sh = wb.sheet_by_name(by_name)
        cv = open("%s/%s.csv" % (output_path, by_name), "w", newline="")
        
    wr = csv.writer(cv, quoting=csv.QUOTE_ALL)

    for r in range(sh.nrows):
        wr.writerow(sh.row_values(r))
        
    cv.close()


def init_year(year):
    if not os.path.isdir(str(year)):
        os.mkdir(str(year))
    
    wb = xlrd.open_workbook("%s.xls" % year)
    for i in range(wb.nsheets):
        to_csv(wb, str(year), by_index=i)    


class NationalServants():
    def __init__(self, file_path):
        self.cont = open(file_path).read()
        self.root = etree.HTML(self.cont.encode('utf-8'))
        self.table = self.root.xpath("//table")
        
    def parse_by_html(self, index, NA=None):
        #table = [list(map(lambda x: x.replace("\xa0", "").replace(" ", "").replace("\u3000", ""), list(r.itertext())))
        #             for r in self.table[index]]
        table = []
        for r in self.table[index]:
            line = []
            for c in r:
                v = "".join(list(c.itertext())).replace("\xa0", "").replace("\u3000", "").replace(" ", "")
                if NA:
                    v = v.replace("－", NA)
    
                if v:
                    line.append(v)
                
            table.append(line)
            
        table = list(filter(lambda x: "".join(x), table))
            
        results = [NS_SERVANTS_TITLE[index]]
        if index == 0:
            # 全國公務人員人數按年別分
            for index, r in enumerate(table[5:35]):
                #print(index, r)
                results.append(r)
        elif index == 1:
            # 全國公務人員人數按官等分
            for index, r in enumerate(table[4:24] + table[30:]):
                #print(index, r)
                results.append(r)
        elif index == 2:
            # 全國公務人員人數按年齡分
            for index, r in enumerate(table[4:27] + table[32:]):
                #print(index, r)
                results.append(r)
        elif index == 3:
            # 全國公務人員人數按教育程度分
            for index, r in enumerate(table[5:]):
                #print(index, r)
                results.append(r)
        elif index == 4:
            # 全國公務人員人數按考試種類分
            for index, r in enumerate(table[5:]):
                #print(index, r)
                results.append(r)
        elif index == 5:
            # 全國公務人員人數按年資分
            for index, r in enumerate(table[4:27] + table[32:]):
                #print(index, r)
                results.append(r)
        elif index == 6:
            # 全國公務人員人數按年齡與年資分
            for index, r in enumerate(table[4:34]):
                #print(index, r)
                results.append(r)
        elif index == 7:
            # 全國公務人員人數按銓敘狀況分
            for index, r in enumerate(table[6:31] + table[38:-1]):
                #print(index, r)
                results.append(r)
        elif index == 8:
            # 全國公務人員人數按職系分
            for index, r in enumerate(table[5:31] + table[37:62] + table[68:93] +
                                      table[98:]):
                #print(index, r)
                results.append(r)
        elif index == 9:
            # 全國公務人員人數按縣市別分
            for index, r in enumerate(table[5:-1]):
                #print(index, r)
                results.append(r)
        elif index == 10:
            # 全國公務人員人數按縣市別及教育程度分
            for index, r in enumerate(table[5:-1]):
                #print(index, r)
                results.append(r)
        elif index == 11:
            # 全國公務人員人數按縣市別及考試種類分
            for index, r in enumerate(table[5:-1]):
                #print(index, r)
                results.append(r)
        elif index == 12:
            # 全國公務人員人數按縣市別及年齡分
            for index, r in enumerate(table[5:-1]):
                #print(index, r)
                results.append(r)
        elif index == 13:
            # 全國公務人員人事異動狀況按機關別分
            for index, r in enumerate(table[5:28] + table[34:]):
                #print(index, r)
                results.append(r)
                
        self.results = results
        return results
    
    
    def output(self, path, file_name, results):
        path = path.rstrip("/")
        if file_name.endswith("csv"):
            with open("%s/%s" % (path, file_name), "w", newline="") as dst:
                cv = csv.writer(dst, quoting=csv.QUOTE_ALL)
                for r in results:
                    cv.writerow(r)
                    
                
if __name__ == '__main__':
    for year in range(95, 102):
        if not os.path.isdir(str(year)):
            os.mkdir(str(year))

        ns = NationalServants("html/%s.html" % year)
    
        for i in range(14):
            ns.parse_by_html(i)
            ns.output("%s" % year, "%s_%s_%s.csv" % (year, i, NS_SERVANTS_SHEET_NAME[i]), ns.results)
        