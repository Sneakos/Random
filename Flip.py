import openpyxl
import operator

wb = openpyxl.load_workbook('Flips.xlsx', data_only=True)
sheet = wb.get_sheet_by_name('Values')

money = float(sheet["B25"].value)

items = {}
costs = {}
limits = {}

for i in range(2,24):
	cellName = "A" + str(i)
	cellCost = "C" + str(i)
	cellEfficiency = "E" + str(i)
	cellLimit = "F" + str(i)
	itemName = sheet[cellName].value
	items[itemName] = float(sheet[cellEfficiency].value)
	costs[itemName] = float(sheet[cellCost].value)
	limits[itemName] = int(sheet[cellLimit].value)

margins = sorted(items.items(), key=operator.itemgetter(1))
margins.reverse()

for item in margins:
	itemStr = item[0]
	limit = limits[itemStr]
	cost = costs[itemStr]
	if cost == 0:
		break
	canBuy = money / cost
	toBuy = 0
	if canBuy < limit:
		toBuy = canBuy
	else:
		canBuy = limit
	print ("%s (%d):" % (itemStr, cost), int(canBuy))
	money = money - (cost * toBuy)
	if canBuy < limit:
		break