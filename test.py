# 简单demo：爬取网络数据，生成表格

import requests
from openpyxl import Workbook

cookies = {
  'D_S': 'mlb5orcjhc563dcmgqbhot5gds',
  'LGD_language': 'en',
  'LGD_firstVisit': '1',
  'LGD_USC': 'eff97eddcdca27b72bdd3246668acde5___ede3e99c-d83c-41db-affe-82506757fbcf',
  '_gid': 'GA1.2.1742586071.1689839907',
  '_ym_uid': '1689839917396746529',
  '_ym_d': '1689839917',
  '_ym_isad': '2',
  '_gcl_au': '1.1.1947098924.1689839922',
  'ln_or': 'eyIzMTUzOTg4IjoiZCJ9',
  'messagesUtk': '0154899ed9b34eeb829dbe53a4f0ce7c',
  '_ym_visorc': 'w',
  '_ga': 'GA1.2.350618632.1689839907',
  '_ga_52H0H4VTHM': 'GS1.1.1689845270.2.1.1689845407.49.0.0',
}

headers = {
  'authority': 'lgdeal.com',
  'accept': 'application/json, text/plain, */*',
  'accept-language': 'zh-CN,zh;q=0.9',
  'api-client-app': 'LGDEAL-WEB-APP',
  'api-client-device': 'Win32',
  'api-client-id': 'eff97eddcdca27b72bdd3246668acde5___ede3e99c-d83c-41db-affe-82506757fbcf',
  'api-client-user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
  'api-language': 'en',
  # 'cookie': 'D_S=mlb5orcjhc563dcmgqbhot5gds; LGD_language=en; LGD_firstVisit=1; LGD_USC=eff97eddcdca27b72bdd3246668acde5___ede3e99c-d83c-41db-affe-82506757fbcf; _gid=GA1.2.1742586071.1689839907; _ym_uid=1689839917396746529; _ym_d=1689839917; _ym_isad=2; _gcl_au=1.1.1947098924.1689839922; ln_or=eyIzMTUzOTg4IjoiZCJ9; messagesUtk=0154899ed9b34eeb829dbe53a4f0ce7c; _ym_visorc=w; _ga=GA1.2.350618632.1689839907; _ga_52H0H4VTHM=GS1.1.1689845270.2.1.1689845407.49.0.0',
  'referer': 'https://lgdeal.com/catalog?q=&sh%5B%5D=1&pfr=40.69&pto=140632.83&cfr=0.14&cto=28.42&rfr=70&rto=99.38&cot=1&covw%5B%5D=0&covw%5B%5D=1&covw%5B%5D=2&covw%5B%5D=3&covw%5B%5D=4&covw%5B%5D=5&covw%5B%5D=6&covf=&fin=&fov=&cl%5B%5D=0&cl%5B%5D=1&cl%5B%5D=2&cl%5B%5D=3&cl%5B%5D=4&cl%5B%5D=5&cl%5B%5D=6&cl%5B%5D=7&cu%5B%5D=0&cu%5B%5D=1&cu%5B%5D=2&cu%5B%5D=3&cu%5B%5D=4&cu%5B%5D=5&loc=0&ce%5B%5D=1&ce%5B%5D=2&ce%5B%5D=3&ce%5B%5D=4&ce%5B%5D=5&ce%5B%5D=6&po%5B%5D=0&po%5B%5D=1&po%5B%5D=2&po%5B%5D=3&po%5B%5D=4&po%5B%5D=5&sy%5B%5D=0&sy%5B%5D=1&sy%5B%5D=2&sy%5B%5D=3&sy%5B%5D=4&sy%5B%5D=5&mzfr=0&mzto=255&mxfr=0&mxto=255&myfr=0&myto=255&tbfr=0&tbto=100&dpfr=0&dpto=100&fl%5B%5D=0&fl%5B%5D=1&fl%5B%5D=2&fl%5B%5D=3&fl%5B%5D=4&gt%5B%5D=0&gt%5B%5D=1&gt%5B%5D=2&s=&sb=rapaport&p=6',
  'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"Windows"',
  'sec-fetch-dest': 'empty',
  'sec-fetch-mode': 'cors',
  'sec-fetch-site': 'same-origin',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
}

url = 'https://lgdeal.com/api/public/catalog/search?sh=1&pfr=40.69&pto=140632.83&cfr=0.14&cto=28.42&rfr=70&rto=99.38&cot=1&covw[]=0&covw[]=1&covw[]=2&covw[]=3&covw[]=4&covw[]=5&covw[]=6&cl[]=0&cl[]=1&cl[]=2&cl[]=3&cl[]=4&cl[]=5&cl[]=6&cl[]=7&cu[]=0&cu[]=1&cu[]=2&cu[]=3&cu[]=4&cu[]=5&loc=0&ce[]=1&ce[]=2&ce[]=3&ce[]=4&ce[]=5&ce[]=6&po[]=0&po[]=1&po[]=2&po[]=3&po[]=4&po[]=5&sy[]=0&sy[]=1&sy[]=2&sy[]=3&sy[]=4&sy[]=5&mzfr=0&mzto=255&mxfr=0&mxto=255&myfr=0&myto=255&tbfr=0&tbto=100&dpfr=0&dpto=100&fl[]=0&fl[]=1&fl[]=2&fl[]=3&fl[]=4&gt[]=0&gt[]=1&gt[]=2&sb=rapaport&p='


# 最终数据
datas = []
# 页码
page = 0

# 封装请求
def getData():
  print('当前在请求第几页：', page)
  _url = url + str(page)
  response = requests.get(
    url=_url,
    cookies=cookies,
    headers=headers,
  )
  response.encoding = 'utf-8'
  # 转为json格式
  json = response.json()
  # 当前20条数据
  list = json["data"]["products"]

  for item in list:
    # 处理每一条数据所需字段
    colData = []
    colData.append(item['shape']['text'])
    colData.append(item['carat'])
    colData.append(item['colorValue']['text'])
    colData.append(item['clarity']['text'])
    colData.append(item['cut']['text'])
    colData.append(item['certificate']['text'])
    colData.append(item['companyId']['countryId']['name'])
    # 将20条数据塞进总数据
    datas.append(colData)

# 每页20条，请求10000次
while page < 3:
  page += 1
  getData()

# 操作表格
# 新建工作簿
workbook = Workbook()
# 实例化一个工作簿(Workbook)后总会默认创建一个工作表(worksheet)，你可以使用 Workbook.active 属性来获取它（默认指向索引为0）
sheet = workbook.active
#工作表的标题
sheet.title = '默认title'
#工作表-表头数据（字段顺序与上述循环中顺序需保持一致）
columns = ['Shape', 'Carat', 'Color', 'Clarity', 'Cut', 'Cert.', 'Location']
sheet.append(columns)
# 工作表-内容数据
for data in datas:
    sheet.append(data)
# 工作簿保存及名称设置
workbook.save('数据列表.xlsx')