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

purchase_order_tuple_list = []

# determines the current section of the spreadsheet
def state_change(current_cell):
  section = ""
  if current_cell == "Purchase Orders":
    section = "PURCHASE_ORDERS"
  elif current_cell == "Warehouse Orders":
    section = "WAREHOUSE_ORDERS"
  return section

# iterate through each row of the spreadsheet
current_po_item_count = 0
for row_index in range(sheet.nrows):
#  print sheet.row_slice(row_index,0)
  state_changed = state_change(sheet.row_slice(row_index,0)[0].value)
  if (not sheet.row_slice(row_index,0)[0].value is empty_cell.value) and ("" == state_changed):
    if "START" == section:
      project_number = sheet.row_slice(row_index,0)[2].value
      project_name = sheet.row_slice(row_index,0)[4].value
      store_number = sheet.row_slice(row_index,0)[6].value
    elif "PURCHASE_ORDERS" == section:
      # if the row constitutes a new PO
      if sheet.row_slice(row_index,0)[3].value is empty_cell.value:
        # if purchase order list is empty
        if purchase_order_tuple_list:
          purchase_order_count += 1
        purchase_order_tuple_list.append((sheet.row_slice(row_index,0)[PURCHASE_ORDER_NUMBER_COL].value,sheet.row_slice(row_index,0)[VENDOR_NAME_COL].value,sheet.row_slice(row_index,0)[2],[]))
      # else the row is the seperate items for a PO
      else:
        purchase_order_tuple_list[purchase_order_count][3].append((sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5])) 
    elif "WAREHOUSE_ORDERS" == section:
      print "in warehouse section"
  else:
     section = state_changed

for po_tuple in purchase_order_tuple_list:
  print po_tuple[0]
  print "item count = " + str(po_tuple[3])

print "Completed for Project " + project_name + ", # " + str(project_number) + ", Store # " + str(store_number)
print "Purchase order count: " + str(purchase_order_count)
