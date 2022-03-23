# coding=utf-8
import csv

image_template = '<figure class="third"><img src="http://image.sinajs.cn/newchart/daily/n/{0}.gif" width="50%">' \
                 '<img src="http://image.sinajs.cn/newchart/min/n/{1}.gif" width="50%"></figure>'
link_template = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vCB_AllNewsStock/symbol/{0}.phtml'
link_temp2 = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vCI_CorpOtherInfo/stockid/{0}/menu_num/5.phtml'
link_taogu = 'https://www.taoguba.com.cn/quotes/{0}'

with open('reduced.csv', 'r') as f:
    reader = csv.reader(f)
    for row in reader:
        code = row[0]
        if code[0] != 's':
            num = code.split('.')[0]
            mkt = code.split('.')[1]
            code = mkt.lower() + num
        link1 = link_temp2.format(num)
        link2 = link_template.format(code)
        print('[基本面]({0})'.format(link1))
        print('[消息面]({0})'.format(link2))
        link3 = link_taogu.format(code)
        print('[淘股吧]({0})'.format(link3))
        print(image_template.format(code, code))
        print('\n')
