import gspread
from oauth2client.service_account import ServiceAccountCredentials
import matplotlib.pyplot as plt
import re
# 以下網址
# https://docs.google.com/spreadsheets/d/1xInDN1K8tb9ommJgID9z5-qmGgMKV_n0CyYLXYT1BfA/edit#gid=0
# https://script.google.com/home/projects/1K8Xsf8gJVI-QpqWMTYAw6va1cVI6c0_GbH8VNmpdbGsemE6Ldkql867k/edit
# https://console.cloud.google.com/apis/dashboard?project=proven-catcher-401401



# Google 試算表認證
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\User\Desktop\python\結訓報告_QRcode_BOT\proven-catcher-401401-38be34938054.json', scope)
gc = gspread.authorize(credentials)

# 開啟 Google 試算表
spreadsheet = gc.open_by_key("1xInDN1K8tb9ommJgID9z5-qmGgMKV_n0CyYLXYT1BfA")
worksheet = spreadsheet.worksheet("問券練習")

# 讀取資料
data = worksheet.get_all_values()
header = data[0]
responses = data[2:]

# 初始化每個選項的計數器
count_1 = count_2 = count_3 = 0

for response in responses:
    for i in range(1, len(header)):
        match = re.search(r'\d', response[i])
        if match:
            answer = int(match.group())
            if answer == 1:
                count_1 += 1
            elif answer == 2:
                count_2 += 1
            elif answer == 3:
                count_3 += 1

# 創建條形圖
plt.figure(num='Why so serious')
plt.rc('font', family='Microsoft JhengHei')
options = [1, 2, 3]
counts = [count_1, count_2, count_3]


plt.bar(options, counts)
plt.title('結訓後的我們...')
# plt.xlabel('選項')
plt.ylabel('數量')

colors = ['blue', 'green', 'orange']
plt.bar(options, counts, color=colors)

# 設定 x 軸刻度，避免小數點
plt.xticks(options, ['滿意清楚很快樂', '人生更上一層樓', '我不想努力了'])

# 設定 y 軸刻度，避免小數點
plt.yticks(range(0, max(counts) + 1))

# 顯示圖表
plt.show()
