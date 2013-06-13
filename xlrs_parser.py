from xlrd import open_workbook,cellname,empty_cell

book = open_workbook('test.xls')

sheet = book.sheet_by_index(0)

# section denotes the section of the spreasheet currently being parsed
#   it begins with START, moves to the purchase orders, then ends with
#   with the warehouse section
section = "START"

# Generic spreadsheet details
project_number = ""
project_name = ""
store_number = ""
purchase_order_count = 0

# Column name static variables
VENDOR_NAME_COL = 0
PURCHASE_ORDER_NUMBER_COL = 1

purchase_order_tuple_list = []

# determines the current section of the spreadsheet
def state_change(current_cell):
  state = ""
  if current_cell == "Purchase Orders":
    state = "PURCHASE_ORDERS"
  elif current_cell == "Warehouse Orders":
    state = "WAREHOUSE_ORDERS"
  return state

# print out a readable set of data for each purchase order and nested item
def po_print(po_list):
  for po_tuple in po_list:
    print "PO #:         " + str(po_tuple[0])
    print "Company Name: " + po_tuple[1]
    print "Ship Date:    " + po_tuple[2].value
    print
    for item_details in po_tuple[3]:
      print "    SKU Description: " + item_details[0]
      print "    Design ID:       " + item_details[1]
      print "    CSI Code:        " + item_details[2].value
      print "    CSI Description: " + str(item_details[3])
      print "    QTY Ordered:     " + str(item_details[4])
      print "    QTY UOM:         " + item_details[5].value
      print 

# iterate through each row of the spreadsheet
current_po_item_count = 0
for row_index in range(sheet.nrows):
#  print sheet.row_slice(row_index,0)
  state = state_change(sheet.row_slice(row_index,0)[0].value)
  if "" != state:
    section = state
  elif (not sheet.row_slice(row_index,0)[0].value is empty_cell.value):
    
    if "START" == section:
      project_number = sheet.row_slice(row_index,0)[2].value
      project_name = sheet.row_slice(row_index,0)[4].value
      store_number = sheet.row_slice(row_index,0)[6].value
    
    elif "PURCHASE_ORDERS" == section:
      # if the row constitutes a new PO
      if sheet.row_slice(row_index,0)[3].value is empty_cell.value:
        print "IN PO TITLE"
        # if purchase order list is empty
        if purchase_order_tuple_list:
          purchase_order_count += 1
        purchase_order_tuple_list.append((sheet.row_slice(row_index,0)[PURCHASE_ORDER_NUMBER_COL].value,sheet.row_slice(row_index,0)[VENDOR_NAME_COL].value,sheet.row_slice(row_index,0)[2],[]))
      
      # else the row is the seperate items for a PO
      else:
        print "IN PO NOT TITLE"
        purchase_order_tuple_list[purchase_order_count][3].append((sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5])) 
    
    elif "WAREHOUSE_ORDERS" == section:
      print "in warehouse section"

po_print(purchase_order_tuple_list)

print "Completed for Project " + project_name + ", # " + str(project_number) + ", Store # " + str(store_number)
print "Purchase order count: " + str(purchase_order_count)

