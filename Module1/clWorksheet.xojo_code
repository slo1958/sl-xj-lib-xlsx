#tag Class
Protected Class clWorksheet
	#tag Method, Flags = &h0
		Sub AddCell(row as integer, column as integer, cell as clCell)
		  
		  while rows.LastIndex <= row
		    rows.add(nil)
		    
		  wend
		  
		  if rows(row) = nil then rows(row) = new clWorkrow(row)
		  
		  if self.lastColumn < column then self.lastColumn = column
		  
		  rows(row).AddCell(column, cell)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(WorkFolder as folderItem, SheetName as string, SheetID as integer)
		  
		  self.Name = SheetName
		  self.Id = SheetID
		  
		  self.Filename = "sheet" + str(sheetID)+".xml"
		  
		  self.SourceFolder = WorkFolder
		  
		  self.LoadWorksheetInfo
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCell(row as integer, column as integer) As clCell
		  
		  if row > rows.LastIndex then return nil
		  
		  return rows(row).GetCell(column)
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSheetData(basenode as XMLNode)
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "row" then self.LoadSheetDataRow(x1)
		    
		    if True then
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSheetDataCell(basenode as XMLNode)
		  
		  var cellrange as string = basenode.GetAttribute("r")
		  var cellstyle as string = basenode.GetAttribute("s")
		  var celltype as String = basenode.GetAttribute("t") // if 't' == 's' => shares string, the value is the index
		  
		  var mycell as new clCell(cellrange, celltype, cellstyle)
		  
		  var p as pair = clCell.ExtractLocation(cellrange)
		  
		  self.AddCell(p.left, p.Right, mycell)
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "v" and x1.FirstChild <> nil then mycell.SetValue(x1.FirstChild.Value)
		    if x1.name = "f"  and x1.FirstChild <> nil then mycell.SetFormula(x1.FirstChild.Value)
		    if True then
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSheetDataRow(basenode as XMLNode)
		  
		  var rowId as string = basenode.GetAttribute("r")
		  var rowspan as string = basenode.GetAttribute("span")
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "c" then self.LoadSheetDataCell(x1)
		    
		    if True then
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorksheetInfo()
		  
		  
		  var tmp as FolderItem = self.SourceFolder
		  
		  if tmp = nil then return
		  
		  tmp = tmp.child("xl")
		  
		  tmp = tmp.child("worksheets")
		  
		  var worksheetxml as XMLDocument = new XMLDocument(tmp.Child(self.Filename))
		  
		  var x1 as xmlnode = worksheetxml.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "dimension" then System.DebugLog("REF:"+x1.GetAttribute("ref"))
		    if x1.name = "sheetData" then LoadSheetData(x1)
		    
		    if x1.name = "worksheet" then
		      x1 = x1.FirstChild
		      
		    else
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Filename As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Id As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		lastColumn As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As string
	#tag EndProperty

	#tag Property, Flags = &h0
		rows() As clWorkrow
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceFolder As FolderItem
	#tag EndProperty


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
		#tag ViewProperty
			Name="Filename"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Id"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
