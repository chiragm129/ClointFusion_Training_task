import ClointFusion as cf
cf.OFF_semi_automatic_mode()

excelfile= r"C:\Users\chirag mali\Downloads\Excel.xlsx"

### 1. Get All Sheet Names
cf.excel_get_all_sheet_names(excelfile)

### 2. Get All Header Columns
header_columns = cf.excel_get_all_header_columns(excel_path = excelfile,sheet_name="Sheet1")
print(header_columns)


### 3. Ger Row & Column Count
Row, Column = cf.excel_get_row_column_count(excel_path = excelfile)
print("Row Count is:",Row)
print("Column Count is:",Column)


###  4. Remove the duplicate data w.r.t ‘ID’ column
cf.excel_remove_duplicates(excel_path=excelfile,sheet_name='Sheet1',header=0,columnName='ID ')


# 5. Sort the data w.r.t ‘OrderDate’ column 
cf.excel_sort_columns(excel_path=excelfile,sheet_name="Sheet1",firstColumnToBeSorted="OrderDate")


### 6. Insert the data at the last row respectively:
#"ID":1027,"OrderDate":4/14/2020,"Region":"East","Rep":"Jones","Item":"Binder","Units":60,"UnitCost":4.99,"Total":449.1
cf.excel_set_single_cell(excel_path=excelfile,columnName="ID",setText=1027)
cf.excel_set_single_cell(excel_path=excelfile,columnName="OrderDate",setText=4/14/2020)
cf.excel_set_single_cell(excel_path=excelfile,columnName="Region",setText="East")
cf.excel_set_single_cell(excel_path=excelfile,columnName="Rep",setText="Jones")
cf.excel_set_single_cell(excel_path=excelfile,columnName="Item",setText="Binder")
cf.excel_set_single_cell(excel_path=excelfile,columnName="Units",setText=60)
cf.excel_set_single_cell(excel_path=excelfile,columnName="UnitCost",setText=4.99)
cf.excel_set_single_cell(excel_path=excelfile,columnName="Total",setText=449.1)



### 7. Split the excel on row count ‘12’. (This will create multiple excel files named ‘Split’. 
# Check the function arguments for further customization).
splitexcelfile = r"C:\Users\chirag mali\Desktop\clointfusion\task 1\split excelfile"
cf.excel_split_the_file_on_row_count(excel_path=excelfile,sheet_name='Sheet1', rowSplitLimit=12, outputFolderPath=splitexcelfile)


# 8. Create a python dictionary named ‘data’ 
# such that it stores the ‘ID’ and ‘Units’ of each row data in the excel file.
data = {}
row,column = cf.excel_get_row_column_count(excel_path=excelfile,sheet_name='Sheet1')
for i in range(row-1):
    key = cf.excel_get_single_cell(excel_path=excelfile,sheet_name='Sheet1',header=0,columnName='ID ',cellNumber=i)
    value = cf.excel_get_single_cell(excel_path=excelfile,sheet_name='Sheet1',columnName='Units',cellNumber=i)
    data[key] = value
print(data)