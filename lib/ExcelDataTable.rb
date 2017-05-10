=begin license

Modified by : Azzi

Copyright (c) 2017, Qantom Software
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer. 
Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution. 
Neither the name of Qantom Software nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission. 
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

(based on BSD Open Source License)
=end

=begin
Class ExcelDataTable contains the functions to manage the excel( both .xls and .xlsx fomat) 
=end
require 'win32OLE'
    class ExcelDataTable      
         
=begin 
    Constructor that takes a excel file path and sheet indiex. The excel file path  is the full
    path to the input Excel datatable.
=end
            def initialize(sfilepath,isheetindex)
                @exlObj = WIN32OLE.new('Excel.Application')
                #~ @dt_file= nil                
                @colms = []
                @rows = []
                 load(sfilepath,isheetindex)
            end
=begin 
    Load the excel data table and specified sheet in memory 
=end    
	    def load(sfilepath,isheetindex)               
                @wb = @exlObj.Workbooks.open(sfilepath)
                setSheet(isheetindex)            
	    end
        
       def save(sfilepath)   
         @wb.SaveAs(sfilepath)
        end         
=begin 
    Set the active sheet to one of the Excel file's. 
=end
            def setSheet(isheetindex) 
                @ws = @wb.WorkSheets(isheetindex)
                @ws.activate            
                 readRowsCols()
              end
              
                 
=begin 
    Get the number of columns in the table. 
=end
            def colCount()
                return @colms.length
            end
            
=begin 
    Get the number of rows in the table. 
=end    
            def rowCount()
                return @rows.length
            end        
            
=begin 
    Get the text in the specified cell of the datatable. The row and 
    column are integers. 
=end
            def getcellValue(row, col)
                return @ws.Cells(row, col).text()
              end
=begin 
    Write the text in the specified cell of the datatable. The row and 
    column are integers. 
=end            
    def writeCell(row,col,value)
        @ws.Cells(row,col).Value= value
        @wb.Save
     
     end
=begin 
    Close the excel work book and also excel object.
=end     
        def close()
              @wb.close
              @exlObj.Quit
        end

            
=begin 
    Get the value of the cell based on colname and rownumber.
    Colname is string and rownum is integer
=end
            def getValue(colname, rownum)   
              
                 row_index=rownum.to_i+1
                
                col_index = @colms.index(colname)
                   return_value = nil
                 if row_index != nil || col_index != nil
                    col_index = col_index + 1
                   return_value = @ws.Cells(row_index, col_index).text.to_s() 
                end
      
  
end
               
=begin 
    Write the text in the specified cell of the datatable. Colname is string and row  
     is integer.
=end     
        def writeValue(colname, rownum, value)     
              
                ret_item = nil        
                row_index = @rows.index(rownum)
                col_index = @colms.index(colname)
                return_value = nil
                
                row_index = rownum.to_i
                col_index = @colms.index(colname)
                            
               if row_index != nil || col_index != nil
                    row_index = row_index + 1
                    col_index = col_index + 1                 
                     @ws.Cells(row_index,col_index).Value= value
                    @wb.Save
                end
            end
        
=begin 
    Private method read the rows and cols from the datatable and store in to @rows, @cols variable.
=end   
            private 
            def readRowsCols()
                @colms = []
                @rows = []
                cur_col = 1
                cur_row = 1
                while true do
                    val = @ws.Cells(cur_row, cur_col).value.to_s()                            
                    if val.strip() == "" 
                        break 
                    else
                        @colms << val
                    end
                    cur_col = cur_col + 1
                end         
                 cur_col = 1
                cur_row = 2            
                while true do
                    val = @ws.Cells(cur_row, cur_col).value
                    if val.to_s().strip() == "" 
                         break 
                    else
                        @rows << val                   
                     end
                    cur_row = cur_row + 1
                end     
                
            end
    end

