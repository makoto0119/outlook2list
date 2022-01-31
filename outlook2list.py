import win32com.client
import datetime
import pyperclip
import win32timezone

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

calendar = outlook.GetDefaultFolder(9).items
calendar.Sort("[Start]")
calendar.IncludeRecurrences = "True"

select_items = [] # 指定した期間内の予定を入れるリスト

print("今日の予定は")

# 予定を抜き出したい期間を指定
today_date = datetime.date.today() # 今日だけ

# restrict appointments to specified range
calendar = calendar.Restrict("[Start] >= '" + str(today_date) +
                             "' AND [END] <= '" + str(today_date + datetime.timedelta(days=1)) + "'")
for item in calendar:
    if today_date == item.start.date():
        select_items.append(item)
    if today_date < item.start.date():
        break 

text =""
# 抜き出した予定の詳細を表示
for select_item in select_items:
    print("件名：", select_item.subject)
#    print("日時：", select_item.start.date())
#    print("場所：", select_item.location)
#    print("開始時刻：", select_item.start.time())
#    print("終了時刻：", select_item.end)
#    print("本文：", select_item.body)
#    print("----")

#    text = text + select_item.subject + '\r\n'
    text = text + select_item.subject + '\r'

pyperclip.copy(text)
print("クリップボードにコピーしました")
input("")
