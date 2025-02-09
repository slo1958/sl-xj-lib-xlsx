#tag Module
Protected Module UI_Support_Methods
	#tag Method, Flags = &h0
		Sub SheetsToPopup(Workbook as clWorkbook, targetPopup as DesktopPopupMenu)
		  
		  targetPopup.RemoveAllRows
		  
		  var sheets() as string =  Workbook.GetSheetNames
		  
		  targetPopup.AddAllRows(sheets)
		  
		  targetPopup.SelectedRowIndex = 0
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SheetToListBox(Workbook as clWorkbook, SheetName as string, targetListbox as DesktopListBox)
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  
		  
		  var sheet as clWorksheet =  Workbook.GetSheet(SheetName)
		  
		  targetListbox.RemoveAllRows
		  
		  if sheet = nil then return
		  
		  targetListbox.ColumnCount = sheet.lastColumn + 1
		  
		  targetListbox.HeaderAt(0) = "#"
		  
		  for i as integer= 1 to sheet.lastColumn + 1
		    targetListbox.HeaderAt(i) = colBase.Middle(i-1,1)
		    
		  next
		  
		  targetListbox.ColumnWidths = "32,"
		  
		  for each row as clWorkrow in sheet.rows
		    
		    if row <> nil then
		      targetListbox.AddRow str(row.row)
		      
		      for col as integer = 1 to sheet.lastColumn
		        var rc as clCell = row.GetCell(col)
		        var tmp as string 
		        if rc <> nil then tmp = rc.GetValueAsString( Workbook)
		        
		        targetListbox.CellTextAt(targetListbox.LastAddedRowIndex, col) = tmp
		        
		      next
		      
		    end if
		    
		  next
		End Sub
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
