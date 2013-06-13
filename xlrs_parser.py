from xlrd import open_workbook,cellname,empty_cell

book = open_workbook('test.xls')

sheet = book.sheet_by_index(0)

# section denotes the section of the spreasheet currently being parsed
#   it begins with START, moves to the purchase orders, then ends with
#   with the warehouse section
section = "START"

project_number = ""
project_name = ""
store_number = ""
purchase_order_count = 0

VENDOR_NAME_COL = 0
PURCHASE_ORDER_NUMBER_COL = 1

print "Sheet name: " + sheet.name
print "Sheet row count: " + str(sheet.nrows)
print "Sheet col count: " + str(sheet.ncols)

# for row_index in range(sheet.nrows):
#  for col_index in range(sheet.ncols):
#   print cellname(row_index,col_index),'-',
#   print sheet.cell(row_index,col_index).value

purchase_order_tuple_list = []

# determines the current section of the spreadsheet
def state_change(current_row):
  section = ""
  if current_row == "Purchase Orders":
    section = "PURCHASE_ORDERS"
  elif current_row == "Warehouse Orders":
    section = "WAREHOUSE_ORDERS"
  return section

# iterate through each row of the spreadsheet
for row_index in range(sheet.nrows):
#  print sheet.row_slice(row_index,0)
  state_changed = state_change(sheet.row_slice(row_index,0)[0].value)
  if (not sheet.row_slice(row_index,0)[0].value is empty_cell.value) and ("" == state_changed):
    if section == "START":
      project_number = sheet.row_slice(row_index,0)[2].value
      project_name = sheet.row_slice(row_index,0)[4].value
      store_number = sheet.row_slice(row_index,0)[6].value
    elif section == "PURCHASE_ORDERS":  
      if sheet.row_slice(row_index,0)[3].value is empty_cell.value:
        purchase_order_tuple_list.append(sheet.row_slice(row_index,0)[PURCHASE_ORDER_NUMBER_COL].value,sheet.row_slice(row_index,0)[VENDOR_NAME_COL].value,sheet.row_slice(row_index,0)[1],[])
      vendor_tuple_list.append(sheet.row_slice(row_index,0))
      print sheet.row_slice(row_index,0)[0] 
    elif section == "WAREHOUSE_ORDERS":
      print "in warehouse section"
  else:
     section = state_changed

print "Completed for Project " + project_name + ", # " + str(project_number) + ", Store # " + str(store_number)
