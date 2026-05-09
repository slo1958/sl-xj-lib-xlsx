#tag Module
Protected Module LibData_Interface
	#tag Method, Flags = &h0
		Function WorksheetToTable(Workbook as clWorkbook, SheetName as string, columnsHaveHeader as Boolean) As clDataTable
		  //
		  // Load a worksheet from the workbook into a clDataTable
		  //
		  // Workbook : the workbook from where the worksheet will be retrieved
		  // SheetName: name of the sheet to load
		  // columnsHaveHeader: use the first data row of the sheet as column name
		  //
		  
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  
		  var shouldPopulateHeader as Boolean =  columnsHaveHeader
		  
		  var sheet as clWorksheet =  Workbook.GetSheetFromName(SheetName)
		  
		  //
		  // Identify columns
		  //
		  
		  var rowIsEmpty as boolean
		  var columns() as clDataSerie
		  
		  for col as integer = 1 to sheet.lastColumn
		    var colName as string = clCellReference.GetColumnLabel(col)
		    columns.Add(new clDataSerie(colName))
		    
		    
		  next
		  
		  for each row as clWorkrow in sheet.rows
		    
		    if row <> nil then
		      var tempColumns() as variant
		      rowIsEmpty = true
		      
		      for col as integer = 1 to sheet.lastColumn
		        var rc as clCell = row.GetCell(col)
		        var tmp as variant
		        if rc <> nil then 
		          tmp = rc.GetValue( Workbook)
		          rowIsEmpty = False
		          
		        end if
		        
		        tempColumns.Add(tmp)
		        
		      next
		      
		      if rowIsEmpty then
		        
		      elseif shouldPopulateHeader then
		        for i as integer = 0 to columns.LastIndex
		          if tempColumns(i).StringValue.Trim.Length > 0 then
		            columns(i).Rename(tempColumns(i).StringValue.Trim)
		            
		          end if
		          
		        next
		        shouldPopulateHeader = false
		        
		      else
		        for i as integer = 0 to columns.LastIndex
		          columns(i).AddElement(tempColumns(i))
		          
		        next
		        
		      end if
		      
		    end if
		    
		  next
		  
		  var tbl as clDataTable = new clDataTable(SheetName, columns)
		  
		  return tbl
		  
		  
		End Function
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
