# RubyExcelDatatable
Ruby Excel datatable fundtions would help in managing the excel and also it has beautiful feature to retrive the column value with the Column name and Row numbers.


Below is the sampel code and usage of datatable functions.


# Exceldatatable function would work for both xls and xlsx.

# Attaching the Excel datatable functions.
require "C:/Azzi-Cdrive/Azzi/Datatable/ExcelDatatableGem/ExcelDataTable.rb"

# Excel path
path = "C:/Azzi-Cdrive/Azzi/Datatable/Input.xlsx"
# sheet index value
index=1

# Instantiate the exceldatatable class with path and index as aruments
@pDatatable = ExcelDataTable.new(path,index)

# Get the Rows count
iRow_Count=@pDatatable.rowCount()
puts iRow_Count

# Get the Columns count
icol_Count=@pDatatable.colCount()
puts icol_Count

# Get the value from passing the column name and row value

puts @pDatatable.getValue("Username",1)
puts @pDatatable.getValue("Password",1)


# Wrrite the value into excel columns name - Username and row - 2nd row
@pDatatable.writeValue("Username",2,"updateduserid")

# Close the datatable.
@pDatatable.close()
