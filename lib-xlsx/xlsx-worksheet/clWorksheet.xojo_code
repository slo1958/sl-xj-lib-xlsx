#tag Class
Protected Class clWorksheet
	#tag Method, Flags = &h0
		Sub AddCell(row as integer, column as integer, cell as clCell)
		  
		  var tmprow as integer = row - 1
		  
		  while rows.LastIndex <= tmprow
		    rows.add(nil)
		    
		  wend
		  
		  if rows(tmprow) = nil then rows(tmprow) = new clWorkrow(row)
		  
		  if self.lastColumn < column then self.lastColumn = column
		  
		  rows(tmprow).AddCell(column, cell)
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(WorkFolder as folderItem, SheetName as string, SheetFilePath as string)
		  
		  self.Name = SheetName
		  
		  self.FilePath = SheetFilePath.Split("/")
		  
		  self.SourceFolder = WorkFolder
		  
		  self.LoadWorksheetInfo
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCell(row as integer, column as integer) As clCell
		  
		  if row > rows.LastIndex then return nil
		  
		  return rows(row-1).GetCell(column)
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadSheetData(basenode as XMLNode)
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl, x1.name)
		    
		    if x1.name = "row" then self.LoadSheetDataRow(x1)
		    
		    if True then
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadSheetDataCell(basenode as XMLNode)
		  
		  const cFormula as string = "f"
		  const cCellValue as string = "v"
		  const cRichText as string = "is"
		  
		  var mycell as new clCell(basenode)
		  
		  var p as pair = clCell.ExtractLocation(mycell.CellLocation)
		  
		  self.AddCell(p.left, p.Right, mycell)
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl, x1.name)
		    
		    if x1.name = cRichText and x1.FirstChild <> nil then mycell.SetValueFromString(x1.FirstChild.Value)
		    if x1.name = cCellValue and x1.FirstChild <> nil then mycell.SetValueFromString(x1.FirstChild.Value)
		    if x1.name = cFormula and x1.FirstChild <> nil then mycell.SetFormula(x1.FirstChild.Value)
		    
		    if True then
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadSheetDataRow(basenode as XMLNode)
		  
		  var rowId as string = basenode.GetAttribute("r")
		  var rowspan as string = basenode.GetAttribute("span")
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl, x1.name)
		    
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
		  
		  for each child as string in self.FilePath
		    tmp = tmp.Child(child)
		    
		  next
		  
		  var worksheetxml as XMLDocument = new XMLDocument(tmp)
		  
		  var x1 as xmlnode = worksheetxml.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl, x1.name)
		    
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
		FilePath() As string
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
			Name="FilePath()"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="lastColumn"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
