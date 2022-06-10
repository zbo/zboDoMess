# coding=utf-8
import comlib
import csv


image_template = '<figure class="third"><img src="http://image.sinajs.cn/newchart/daily/n/{0}.gif" width="50%">' \
                 '<img src="http://image.sinajs.cn/newchart/min/n/{1}.gif" width="50%"></figure>'
link_temp1 = '<a href="http://vip.stock.finance.sina.com.cn/corp/go.php/vCB_AllNewsStock/symbol/{0}.phtml">消息面</a>'
link_temp2 = '<p><a href="http://vip.stock.finance.sina.com.cn/corp/go.php/vCI_CorpOtherInfo/stockid/{0}/menu_num/5.phtml">基本面</a>'
link_taogu = '<a herf="https://www.taoguba.com.cn/quotes/{0}"></a>'
jquery = '<script src="./jquery.js"></script>'
chartjs = '<script src="./chart.js"></script>'
cate_dict = comlib.get_cate_dict_from_xl()
print(jquery)
print(chartjs)
print('<button id="btn_show1">Show me the money</button>')
print('<p id="result1"></p>')
with open('reduced.csv', 'r') as f:
    reader = csv.reader(f)
    for row in reader:
        code = row[0]
        if code[0] != 's':
            num = code.split('.')[0]
            mkt = code.split('.')[1]
            code = mkt.lower() + num
        link1 = link_temp2.format(num)
        link2 = link_temp1.format(code)
        print(link1)
        print(link2)
        link3 = link_taogu.format(code)
        print(link3)
        cate = 'Unknow'
        if row[0] in cate_dict:
            cate = cate_dict[row[0]]
        print('<a>{0}</a>'.format(cate))
        print('<input id ="{0}" type="checkbox"/></p>'.format(code))
        print(image_template.format(code, code))
print('<p id="result2"></p>')
print('<button id="btn_show2">Show me the money</button>')

