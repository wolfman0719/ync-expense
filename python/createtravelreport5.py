import openpyxl
from openpyxl import Workbook
import pandas
import sys
import datetime

args = sys.argv

sys.path += ['/opt/iris/lib/python','/opt/iris/mgr/python']

import iris

inputFilename = args[1]
outputFilename = args[2]
reportmonth = args[3]
reportday = args[4]

year = datetime.datetime.today().year
reiwayear = year - 2018

status = iris.cls('YNC.Expense')._KillExtent()
wb = openpyxl.load_workbook(inputFilename)
ws = wb['input']
		
row_index = 0
for row in ws.iter_rows():
	row_index = row_index + 1
	if row_index < 5: continue
	month = ws.cell(row=row_index,column=2).value
	if (month == '' or month is None): continue
	day = ws.cell(row=row_index,column=3).value
	if (day == '' or day is None): continue
	paymentto = ws.cell(row=row_index,column=4).value
	accounts = ws.cell(row=row_index, column=5).value
	if accounts is None: accounts = '旅費交通費'
	amount = ws.cell(row=row_index, column=6).value
	if (amount == '' or amount is None): amount = 0
	description = ws.cell(row=row_index, column=7).value
	if description is None: description = 'no data'
	isjreimbursement = ws.cell(row=row_index,column=8).value
	onbehalf = ws.cell(row=row_index,column=9).value
	if isjreimbursement is None: isjreimbursement = ''
	if onbehalf is None: onbehalf = ''
	sql = iris.sql.prepare("insert into ync.expense(reportmonth,reportday,paymentto,accounts,amount,description,isjreimbursement,onbehalf) values(?,?,?,?,?,?,?,?)")
	sql.execute(month,day,paymentto,accounts,amount,description,isjreimbursement,onbehalf)

	sql = iris.sql.prepare("select description from ync.expenseitem where description = ?")
	rs = sql.execute(description)
	try:
		next(rs)
		exist = True
	except Exception:
		exist = False
	if exist is False:
		sql = iris.sql.prepare("insert into ync.expenseitem(paymentto,accounts,amount,description) values(?,?,?,?)")
		sql.execute(paymentto,accounts,amount,description)
	else:
		if not (amount == 0 or amount is None):
			sql = iris.sql.prepare("update ync.expenseitem set paymentto=?,accounts=?,amount=? where description= ?")
			sql.execute(paymentto,accounts,amount,description)

wb.close()

sql = None

wb = openpyxl.load_workbook(outputFilename)
		
ws = wb['印刷用']
		
minmonth = reportmonth
minday = reportday
minmd = int(minmonth) * 100 + int(minday)

# 出発日以降の旅費交通費の最終日を取得（月をまたぐケースに対応するため reportmd で比較）
itemline1 = iris.sql.exec("select max(+reportmd) from ync.expense where accounts = '旅費交通費' and +reportmd >= " + str(minmd)).dataframe()

for index,row in itemline1.iterrows():
	rowline = list(row)
	maxmd = rowline[0]
	if maxmd is None:
		maxmd = minmd
	maxmd = str(int(maxmd))
	if (len(maxmd) == 3): maxmd = '0' + maxmd
	maxmonth = maxmd[0:2]
	maxday = maxmd[2:4]

maxdt = datetime.datetime(year=year, month=int(maxmonth), day=int(maxday))
mindt = datetime.datetime(year=year, month=int(minmonth), day=int(minday))

td = maxdt - mindt
nights = td.days

ws.cell(row=2,column=2).value = reiwayear
ws.cell(row=2,column=4).value = maxmonth
ws.cell(row=2,column=6).value = maxday

ws.cell(row=7,column=14).value = reiwayear
ws.cell(row=7,column=16).value = minmonth
ws.cell(row=7,column=18).value = minday
ws.cell(row=8,column=14).value = reiwayear
ws.cell(row=8,column=16).value = maxmonth
ws.cell(row=8,column=18).value = maxday

ws.cell(row=13,column=4).value = str(reiwayear) + '年' + minmonth + '月' + minday + '日'
ws.cell(row=13,column=12).value = str(reiwayear) + '年' + maxmonth + '月' + maxday + '日'

maxmd_int = int(maxmonth) * 100 + int(maxday)
sql = iris.sql.prepare("Select reportmonth, reportday, amount, description, paymentto from ync.expense where (+reportmd >= ? and +reportmd <= ? and accounts = '旅費交通費') order by reportmonth, reportday")
itemline = sql.execute(minmd, maxmd_int).dataframe()

hotelamount = 0
hotelname = ''

# まず書き込み対象の行リストを組み立てる
# 各要素は (month, day, amount, description) のdict
items = []

for index,row in itemline.iterrows():
	rowline = list(row)
	month = rowline[0]
	day = rowline[1]
	amount = rowline[2]
	description = rowline[3]
	paymentto = rowline[4]

	sql = iris.sql.prepare("select accounts, amount from ync.expenseitem where description = ?")
	rs = sql.execute(description)
	for index,row in enumerate(rs):
		accounts = row[0]
		if (accounts is None or accounts == ''): accounts = '旅費交通費'
		if (amount == 0 or amount is None): amount = row[1]

	if description == '雲雀丘花屋敷>大阪梅田': continue
	if description == '大阪梅田>雲雀丘花屋敷': continue

	if description == 'HOTEL':
		hotelamount = hotelamount + amount
		hotelname = hotelname + paymentto + ' '
		continue

	if paymentto == 'JAL':
		transport_type = '飛行機'
	elif paymentto == 'PEACH':
		transport_type = '飛行機'
	elif paymentto == 'SKYMARK':
		transport_type = '飛行機'
	elif paymentto == 'ANA':
		transport_type = '飛行機'
	else:
		transport_type = '電車'

	full_description = transport_type + '　(' + description + ')'
	items.append({'month': month, 'day': day, 'end_month': month, 'end_day': day, 'amount': amount, 'description': full_description, 'transport_type': transport_type, 'route': description})


# description の最大文字数。これ以上になる場合はマージしない
MAX_DESC_LENGTH = 24


def _connect_routes(routes):
	"""隣接する区間で到着駅と次の出発駅が同じ場合は繋げて A>B>C 形式にする"""
	result = routes[0]
	for r in routes[1:]:
		prev_arrival = result.split('>')[-1]
		next_departure = r.split('>')[0]
		if prev_arrival == next_departure:
			result = result + '>' + '>'.join(r.split('>')[1:])
		else:
			result = result + ' ' + r
	return result


def _reorder_group_for_connection(group):
	"""グループ内を並び替えて、到着点と次の出発点が一致する項目を隣接させる。
	先頭項目は固定し、残りの項目から到着点に一致する出発点を持つものを貪欲に選ぶ。"""
	if len(group) <= 1:
		return group
	result = [group[0]]
	remaining = list(group[1:])
	while remaining:
		cur_arrival = result[-1]['route'].split('>')[-1]
		picked = -1
		for k, it in enumerate(remaining):
			if it['route'].split('>')[0] == cur_arrival:
				picked = k
				break
		if picked == -1:
			picked = 0
		result.append(remaining.pop(picked))
	return result


def merge_same_day(items, merge_count):
	"""同じ日付かつ同じ交通手段の行を merge_count 件ずつ1行にまとめる処理を1パス行う。
	マージ対象グループ内は、到着点と次の出発点が一致する項目が隣接するよう並び替える。"""
	merged = []
	i = 0
	while i < len(items):
		base = items[i]
		# 同じ日付・同じ交通手段の連続する行をmerge_count件集める
		group = [base]
		j = i + 1
		while j < len(items) and len(group) < merge_count:
			if (items[j]['month'] == base['month']
				and items[j]['day'] == base['day']
				and items[j]['transport_type'] == base['transport_type']):
				group.append(items[j])
				j += 1
			else:
				break
		if len(group) >= 2:
			# ルートが繋がるように並び替える（終点と起点が同じ項目を優先）
			group = _reorder_group_for_connection(group)
			merged_route = _connect_routes([item['route'] for item in group])
			merged_desc = base['transport_type'] + '　(' + merged_route + ')'
			if len(merged_desc) >= MAX_DESC_LENGTH:
				# マージ後の description が長すぎる場合はマージしない
				merged.append(base)
				i += 1
				continue
			merged_amount = sum(item['amount'] for item in group)
			last = group[-1]
			merged.append({
				'month': base['month'],
				'day': base['day'],
				'end_month': last.get('end_month', last['month']),
				'end_day': last.get('end_day', last['day']),
				'amount': merged_amount,
				'description': merged_desc,
				'transport_type': base['transport_type'],
				'route': merged_route
			})
			i = j
		else:
			merged.append(base)
			i += 1
	return merged


def merge_same_transport(items, merge_count):
	"""transport_typeが同じ行を merge_count 件ずつ1行にまとめる処理を1パス行う。
	日付が異なっても同一 transport_type であればマージ対象とする。
	項目選択時は、現在のグループの到着点と出発点が一致する項目を優先し、
	伊丹>羽田 と 羽田>伊丹 のような往復を繋いでまとめる。"""
	used = [False] * len(items)
	merged = []
	for i in range(len(items)):
		if used[i]:
			continue
		base = items[i]
		group = [base]
		used[i] = True
		added_indices = []  # この反復で used にしたインデックス (rollback 用)
		# 現在のグループに項目を追加する。
		# 1st 優先: 同一 transport_type で、現在の末尾ルートの到着点から出発する項目
		# 2nd 優先: 同一 transport_type で先頭に見つかる項目
		while len(group) < merge_count:
			cur_arrival = group[-1]['route'].split('>')[-1]
			chosen = -1
			for j in range(i + 1, len(items)):
				if used[j]:
					continue
				if items[j]['transport_type'] != base['transport_type']:
					continue
				if items[j]['route'].split('>')[0] == cur_arrival:
					chosen = j
					break
			if chosen == -1:
				for j in range(i + 1, len(items)):
					if used[j]:
						continue
					if items[j]['transport_type'] == base['transport_type']:
						chosen = j
						break
			if chosen == -1:
				break
			# 追加候補を入れた場合の merged_desc を事前に確認し、
			# MAX_DESC_LENGTH 以上になるなら追加せず打ち切る
			tentative = _reorder_group_for_connection(group + [items[chosen]])
			tentative_route = _connect_routes([it['route'] for it in tentative])
			tentative_desc = base['transport_type'] + '　(' + tentative_route + ')'
			if len(tentative_desc) >= MAX_DESC_LENGTH:
				break
			group.append(items[chosen])
			used[chosen] = True
			added_indices.append(chosen)
		if len(group) >= 2:
			# 選択後にもう一度並び替えて、可能な限り繋がるようにする
			group = _reorder_group_for_connection(group)
			merged_route = _connect_routes([item['route'] for item in group])
			merged_desc = base['transport_type'] + '　(' + merged_route + ')'
			if len(merged_desc) >= MAX_DESC_LENGTH:
				# 念のためもう一度チェック。長すぎる場合はマージせず、
				# 追加した項目を未使用に戻して base だけ出力する
				for idx in added_indices:
					used[idx] = False
				merged.append(base)
				continue
			merged_amount = sum(item['amount'] for item in group)
			# 最終日付は group 内で最も遅い日付を採用する
			def _md(it):
				return int(it.get('end_month', it['month'])) * 100 + int(it.get('end_day', it['day']))
			last = max(group, key=_md)
			merged.append({
				'month': base['month'],
				'day': base['day'],
				'end_month': last.get('end_month', last['month']),
				'end_day': last.get('end_day', last['day']),
				'amount': merged_amount,
				'description': merged_desc,
				'transport_type': base['transport_type'],
				'route': merged_route
			})
		else:
			merged.append(base)
	return merged


MAX_LINES = 10

# 10行を超える場合、同じ日付の2行→3行→...とマージを繰り返す
merge_count = 2
while len(items) > MAX_LINES:
	prev_len = len(items)
	items = merge_same_day(items, merge_count)
	# 同じmerge_countで変化がなくなったら merge_count を増やす
	if len(items) == prev_len:
		merge_count += 1
		if merge_count > 100:
			# 無限ループ防止: これ以上マージできない場合は打ち切り
			break

# それでも MAX_LINES を超える場合は、日付が異なっても同一 transport_type の行をマージ
merge_count = 2
while len(items) > MAX_LINES:
	prev_len = len(items)
	items = merge_same_transport(items, merge_count)
	if len(items) == prev_len:
		merge_count += 1
		if merge_count > 100:
			# 無限ループ防止: これ以上マージできない場合は打ち切り
			print('# of lineitems exceeded the max limit even after merging')
			break

# 行リストをシートに書き込む
linepos = 16
for item in items:
	linepos += 1
	ws.cell(row=linepos, column=1).value = str(item['month']) + '/' + str(item['day'])
	# マージ結果で開始日と最終日が異なる場合はカラム6に最終日付を設定する
	end_month = item.get('end_month', item['month'])
	end_day = item.get('end_day', item['day'])
	ws.cell(row=linepos, column=6).value = str(end_month) + '/' + str(end_day)
	ws.cell(row=linepos, column=10).value = item['description']
	ws.cell(row=linepos, column=16).value = item['amount']

sql = None
ws.cell(row=30,column=8).value = hotelamount
ws.cell(row=30,column=2).value = hotelname
ws.cell(row=28,column=5).value = nights
wb.save(outputFilename)
wb.close()
