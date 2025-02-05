#tag Class
Protected Class clWorksheet
	#tag Method, Flags = &h0
		Sub Constructor(WorkFolder as folderItem, SheetName as string, SheetID as integer)
		  
		  self.Name = SheetName
		  self.Id = SheetID
		  
		  self.Filename = SheetName+".xml"
		  
		  self.SourceFolder = WorkFolder
		  
		  self.LoadWorksheetInfo
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSheetData(basenode as XMLNode)
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "row" then System.DebugLog("ROW:"+x1.GetAttribute("r"))
		     
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
		Name As string
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
	#tag EndViewBehavior
End Class
#tag EndClass
