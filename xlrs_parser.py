from xlrd import open_workbook,cellname,empty_cell

def state_change(current_cell):
  """ determines the current section of the spreadsheet """
  state = ""
  if current_cell == "Purchase Orders":
    state = "PURCHASE_ORDERS"
  elif current_cell == "Warehouse Orders":
    state = "WAREHOUSE_ORDERS"
  return state

def po_print(po_list):
  """ prints out a readable set of data for each purchase order and nested item """
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

def warehouse_print(warehouse_list):
  """ prints out a readable set of data for each warehouse order and nested item """
  for warehouse_tuple in warehouse_list:
    print "Warehouse #:         " + str(warehouse_tuple[0])
    print
    for item_details in warehouse_tuple[1]:
      print "    WHO #:           " + item_details[0]
      print "    Stage:           " + item_details[1]
      print "    SKU #:           " + item_details[2].value
      print "    SKU Description: " + str(item_details[3])
      print "    Design ID:       " + str(item_details[4])
      print "    QTY:             " + item_details[5].value
      print "    Target Date:     " + item_details[6].value
      print


def retrieve_po_and_warehouse_lists(po_file_name):
  """ Parses a file with a given name (po_file_name) and returns a tuple with the PO list as the first element and the warehouse list as the second element """

  print "Retrieving PO and Warehouse orders and items lists" 
  book = open_workbook(po_file_name)
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
  warehouse_order_count = 0

  # Column name static variables
  VENDOR_NAME_COL = 0
  PURCHASE_ORDER_NUMBER_COL = 1

  purchase_order_tuple_list = []
  warehouse_order_tuple_list = []

  # iterate through each row of the spreadsheet
  current_po_item_count = 0
  for row_index in range(sheet.nrows):
    
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
          # if purchase order list is empty
          if purchase_order_tuple_list:
            purchase_order_count += 1
          purchase_order_tuple_list.append((sheet.row_slice(row_index,0)[PURCHASE_ORDER_NUMBER_COL].value,sheet.row_slice(row_index,0)[VENDOR_NAME_COL].value,sheet.row_slice(row_index,0)[2],[]))
      
        # else the row is the seperate items for a PO
        else:
          purchase_order_tuple_list[purchase_order_count][3].append((sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5])) 
    
      elif "WAREHOUSE_ORDERS" == section:
        # if the row constitutes a new warehouse
        if sheet.row_slice(row_index,0)[1].value is empty_cell.value:
          if warehouse_order_tuple_list:
            warehouse_order_count += 1
          warehouse_order_tuple_list.append(((sheet.row_slice(row_index,0)[0].value),[]))
        else: 
          warehouse_order_tuple_list[warehouse_order_count][1].append((sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5],sheet.row_slice(row_index,0)[6]))

  # Uncomment the following two lines to print out the stored PO and warehouse data
  #po_print(purchase_order_tuple_list)
  #warehouse_print(warehouse_order_tuple_list)

  print "Completed for Project [" + project_name + "], [# " + str(project_number) + "], Store #" + str(store_number)
  print "Purchase order count: " + str(purchase_order_count)
  print "Warehouse order count: " + str(warehouse_order_count)

  return purchase_order_tuple_list, warehouse_order_tuple_list

def retrieve_proposed_orders_lists(po_file_name):
  print "Retrieving proposed orders list from file: " + po_file_name + " ..."
  # TODO add implementation
