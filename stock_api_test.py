import json
import urllib.request

item_code = "373220"
url = "https://m.stock.naver.com/api/stock/%s/integration"%(item_code)

raw_data = urllib.request.urlopen(url).read()
json_data = json.loads(raw_data)

#종목명 가져오기
stock_name = json_data['stockName']
print("종목명 : %s"%(stock_name))

#가격 가져오기
current_price = json_data['dealTrendInfos'][0]['closePrice']
print("가격 : %s"%(current_price))

#시총 가져오기
for code in json_data['totalInfos']:
    if 'marketValue' == code['code']:
        marketSum_value = code['value']
        print("시총 : %s"%(marketSum_value))

#PER 가져오기
for i in json_data['totalInfos']:
    if 'per' == i['code']:
        per_value_str = i['value']
        print("PER : %s"%(per_value_str))


#PBR 가져오기
for v in json_data['totalInfos']:
    if 'pbr' == v['code']:
        pbr_value_str = v['value']
        print("PBR : %s"%(pbr_value_str))
