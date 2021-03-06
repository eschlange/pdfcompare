from xlrd import open_workbook,cellname,empty_cell
from datetime import datetime
import Levenshtein

def state_change(current_cell):
  """ determines the current section of the spreadsheet """
  state = ""
  if current_cell == "Purchase Orders":
    state = "PURCHASE_ORDERS"
  elif current_cell == "Warehouse Orders":
    state = "WAREHOUSE_ORDERS"
  elif current_cell == "DESIGN ID":
    state = "PROPOSED_ORDERS"
  return state

def po_item_print(po_item):
  print "    PO #:         " + str(item_details[6])
  print "    Company Name: " + item_details[7]
  print "    Ship Date:    " + item_details[8].value
  print "    Design ID:       " + item_details[1]
  print "    SKU Description: " + item_details[0]
  print "    CSI Code:        " + item_details[2].value
  print "    CSI Description: " + str(item_details[3])
  print "    QTY Ordered:     " + str(item_details[4])
  print "    QTY UOM:         " + item_details[5].value
  print

def po_item_list_print(po_item_list):
  for po_item in po_item_list:
    po_item_print(po_item)

def proposed_order_item_print(item_details):
  print "    Design ID:         " + item_details[0]
  print "    Mapping Status:    " + item_details[1]
  print "    Revit Description: " + item_details[2]
  print "    Category:          " + str(item_details[3])
  print "    Quantity:          " + str(item_details[4])
  print "    Coverage Unit:     " + item_details[5]
  print "    Responsibility:    " + item_details[6]
  print "    Comments:          " + item_details[7]
  print

def proposed_po_equality_check_print(proposed_item,po_item):
  print " Proposed Design ID  :" + proposed_item[0] + '\t\t\t'      + " PO Design ID  :" + po_item[1]
  print " Proposed Description:" + proposed_item[2] + '\t\t\t'      + " PO Description:" + po_item[0]
  print " Proposed Quantity   :" + str(proposed_item[4]) + '\t\t\t' + " PO Quantity   :" + str(po_item[4])
  print

def proposed_order_print(proposed_order_list):
  """ prints out a readable set of data for each warehouse order and nested item """
  for item_details in proposed_order_list:
    proposed_order_item_print(item_details)

def warehouse_item_print(item_details):
  print "    Warehouse ID:    " + str(item_details[8])
  print "    WHO #:           " + item_details[0]
  print "    Stage:           " + item_details[1]
  print "    SKU #:           " + item_details[2]
  print "    SKU Description: " + item_details[3].value
  print "    Design ID:       " + str(item_details[4])
  print "    QTY:             " + str(item_details[5])
  print "    Target Date:     " + item_details[6].value
  print

def warehouse_print(warehouse_list):
  """ prints out a readable set of data for each warehouse order and nested item """
  for warehouse_item in warehouse_list:
    warehouse_item_print(warehouse_item)

def retrieve_po_and_warehouse_lists(po_file_name):
  """ Parses a file with a given name (po_file_name) and returns a tuple with the PO list as the first element and the warehouse list as the second element """
  print "Parsing the specified PO and Warehouse orders and items lists from file \"" + po_file_name + "\"..."
  book = open_workbook(po_file_name)
  sheet = book.sheet_by_index(0)

  # section denotes the section of the spreasheet currently being parsed
  #   it begins with START, moves to the purchase orders, then ends with
  #   the warehouse section
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

  # Variables to hold current purchase order data
  current_po_id = 0
  current_vendor = ""
  current_target_del_date = ""  

  # Variables to hold current warehouse data
  current_warehouse_id = 0  
  current_warehouse_WHO = ""
  
  # iterate through each row of the spreadsheet
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
          current_po_id = sheet.row_slice(row_index,0)[PURCHASE_ORDER_NUMBER_COL].value
          current_vendor = sheet.row_slice(row_index,0)[0].value
          current_target_del_date = sheet.row_slice(row_index,0)[2].value 
        # else the row is the seperate items for a PO
        else:
          purchase_order_tuple_list.append((sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5],current_po_id,current_vendor,current_target_del_date))
    
      elif "WAREHOUSE_ORDERS" == section:
        # if the row constitutes a new warehouse
        if sheet.row_slice(row_index,0)[1].value is empty_cell.value:
          current_warehouse_id = sheet.row_slice(row_index,0)[0].value
          current_warehouse_WHO = sheet.row_slice(row_index,0)[1].value
          if warehouse_order_tuple_list:
            warehouse_order_count += 1
        else: 
          warehouse_order_tuple_list.append((current_warehouse_WHO,sheet.row_slice(row_index,0)[0].value,sheet.row_slice(row_index,0)[1].value,sheet.row_slice(row_index,0)[2],sheet.row_slice(row_index,0)[3].value,sheet.row_slice(row_index,0)[4].value,sheet.row_slice(row_index,0)[5],sheet.row_slice(row_index,0)[6],current_warehouse_id))

  # Uncomment the following two lines to print out the stored PO and warehouse data
  # po_print(purchase_order_tuple_list)
  # warehouse_print(warehouse_order_tuple_list)

  print "Completed for Project [" + project_name + "], [# " + str(project_number) + "], Store #" + str(store_number)
  print "Purchase order count: " + str(purchase_order_count)
  print "Warehouse order count: " + str(warehouse_order_count)
  print "*******************************"
  print

  return purchase_order_tuple_list, warehouse_order_tuple_list

def retrieve_proposed_orders_lists(po_file_name):
  """ Parses a file with a given name (po_file_name) and returns a tuple with the PO list as the first element and the warehouse list as the second element """
  print "*******************************"
  print "Parsing proposed orders list from file \"" + po_file_name + "\"..."

  print "Retrieving PO and Warehouse orders and items lists"
  book = open_workbook(po_file_name)
  sheet = book.sheet_by_index(0)

  # section denotes the section of the spreasheet currently being parsed
  #   it begins with START, then ends with the proposed orders section
  section = "START"

  # Generic spreadsheet details
  project_number = ""
  project_type = ""
  store_number = ""
  store_name = ""

  proposed_order_count = 0

  # Column name static variables
  DESIGN_ID_COL = 0
  MAPPING_STATUS_COL = 1
  REVIT_DESCRIPTION_COL = 2
  CATEGORY_COL = 3
  QUANTITY_COL = 4
  COVERAGE_UNIT_COL = 5
  RESPONSIBILITY_COL = 6
  COMMENTS_COL = 7

  proposed_order_tuple_list = []
 
  # Pull general details for the proposed orders
  project_number = sheet.row_slice(0,0)[1].value
  project_type = sheet.row_slice(3,0)[1].value
  store_number = sheet.row_slice(2,0)[1].value
  store_name = sheet.row_slice(1,0)[1].value

  for row_index in range(sheet.nrows):
    state = state_change(sheet.row_slice(row_index,0)[0].value)
    if "" != state:
      section = state

    elif (not sheet.row_slice(row_index,0)[0].value is empty_cell.value):
      if "PROPOSED_ORDERS" == section:
        # if proposed order list is empty
        if proposed_order_tuple_list:
          proposed_order_count += 1
        proposed_order_tuple_list.append((sheet.row_slice(row_index,0)[DESIGN_ID_COL].value,sheet.row_slice(row_index,0)[MAPPING_STATUS_COL].value,sheet.row_slice(row_index,0)[REVIT_DESCRIPTION_COL].value,sheet.row_slice(row_index,0)[CATEGORY_COL].value,sheet.row_slice(row_index,0)[QUANTITY_COL].value,sheet.row_slice(row_index,0)[COVERAGE_UNIT_COL].value,sheet.row_slice(row_index,0)[RESPONSIBILITY_COL].value,sheet.row_slice(row_index,0)[COMMENTS_COL].value,[]))

  # Uncomment the following two lines to print out the stored PO and warehouse data
  # proposed_order_print(proposed_order_tuple_list)

  print "Completed retrieval of proposed orders for Project type [" + project_type + "], [# " + str(project_number) + "], Store #" + str(store_number)
  print "Proposed order count: " + str(proposed_order_count)
  print "*******************************"
  print
  return proposed_order_tuple_list

def compare_po_description(proposed_item,po_item):
  ratio_to_return = 0
  ratio_found = Levenshtein.ratio(po_item[0].lower().replace(" ", ""),proposed_item[2].lower().replace(" ", ""))
  try:
    if ratio_found > 0.4:
      ratio_to_return = ratio_found
  except:
    print "error"
  return ratio_to_return

def compare_wh_description(proposed_item,wh_item):
  ratio_to_return = 0
  ratio_found = Levenshtein.ratio(wh_item[3].value.lower().replace(" ", ""),proposed_item[2].lower().replace(" ", ""))
  if ratio_found > 0.4:
    ratio_to_return = ratio_found
  return ratio_to_return


def compare(po_and_warehouse_file,proposed_order_file):
  print "***************************"
  startTime = datetime.now()
  po_and_warehouse_tuple_list = retrieve_po_and_warehouse_lists(po_and_warehouse_file)
  proposed_order_tuple_list = retrieve_proposed_orders_lists(proposed_order_file)
  found_po_pair_tuple_list = []
  found_warehouse_tuple_list = []
  not_found_po_warehouse_item_list = []
  found_map = {}

  for proposed_item_details in proposed_order_tuple_list:
    item_found = False
    winning_item_tuple = None

    # for each purchase order
    for po_item in po_and_warehouse_tuple_list[0]:
      if proposed_item_details[0] == po_item[1].replace('.0',''):
        found_po_pair_tuple_list.append((proposed_item_details,po_item))
        item_found = True
        break
      similarity_ratio = compare_po_description(proposed_item_details,po_item)
      if similarity_ratio > 0.55 and (winning_item_tuple is None or winning_item_tuple[0] < similarity_ratio):
        winning_item_tuple = similarity_ratio,proposed_item_details,po_item,"PO"

    # for each warehouse item
    for warehouse_order in po_and_warehouse_tuple_list[1]:
      if item_found:
        break
      if proposed_item_details[0] == warehouse_order[4]:
        found_warehouse_tuple_list.append((proposed_item_details,warehouse_order))
        item_found = True
        break
      similarity_ratio = compare_wh_description(proposed_item_details,warehouse_order)
      if similarity_ratio > 0.55 and (winning_item_tuple is None or winning_item_tuple[0] < similarity_ratio):
        winning_item_tuple = similarity_ratio,proposed_item_details,warehouse_order,"WH"

    # if design id match didn't occur and found a valid description match
    if not item_found and None is not winning_item_tuple:
      if "PO" == winning_item_tuple[3]:
         found_po_pair_tuple_list.append((winning_item_tuple[1],winning_item_tuple[2]))
      else:
         found_warehouse_tuple_list.append((winning_item_tuple[1],winning_item_tuple[2]))
      item_found = True

    if not item_found:
      not_found_po_warehouse_item_list.append(proposed_item_details)
    else: 
      found_map[str(proposed_item_details[0])]=True

  # Show all PO matches that don't have the same quantity
  for matched_po_tuple in found_po_pair_tuple_list:
    if matched_po_tuple[0][4] != matched_po_tuple[1][4]:
      proposed_po_equality_check_print(matched_po_tuple[0],matched_po_tuple[1])

  print "Found [" + str(len(found_po_pair_tuple_list)) + "] proposed orders that match with an item in the PO list."
  print "Found [" + str(len(found_warehouse_tuple_list)) + "] proposed orders that match with an item in the warehouse list."
  print "Found [" + str(len(not_found_po_warehouse_item_list)) + "] proposed orders that DO NOT have a design ID match in the PO and warehouse list."
  print "Total run time: " + str(datetime.now()-startTime)
  print
