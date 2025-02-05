#tag Class
Protected Class clWorkbook
	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, workfolder as FolderItem = nil)
		  
		  self.SourceFile = file
		  self.TempFolder = workfolder
		  
		  if self.UnzipToTemporary <> 0 then
		    return
		    
		  end if
		  
		  
		  self.LoadWorkbookInfo()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorkbookInfo()
		  
		  var tmp as FolderItem = self.TempFolder
		  
		  if tmp = nil then return
		  
		  
		  var workbookxml as XMLDocument = new XMLDocument(tmp.Child("xl").child("workbook.xml"))
		  
		  var x1 as xmlnode = workbookxml.FirstChild
		  var lvl as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl)+":"+x1.name)
		    
		    if x1.name = "sheet" then
		      var name as string = x1.GetAttribute("name")
		      var sheetid as string = x1.GetAttribute("sheetId")
		      var rid as String = x1.GetAttribute("r:id")
		      sheets.Add(new clWorksheet( self.TempFolder, name, sheetid.ToInteger))
		      
		      lvl =lvl
		      
		    end if
		    
		    // Navigate the tree
		    if x1.name ="workbook" then 
		      x1 = x1.FirstChild
		      lvl = lvl+1
		      
		    elseif x1.name = "sheets" then
		      x1 = x1.FirstChild
		      lvl = lvl + 1
		      
		    else
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnzipToTemporary() As integer
		  
		  var file as FolderItem = self.SourceFile
		  
		  if self.TempFolder = nil then
		    
		    self.TempFolder  = SpecialFolder.Desktop //SpecialFolder.Temporary
		    
		    // prepare work area
		    
		    self.TempFolder = self.TempFolder.Child(file.name.ReplaceAll(".","-") + " folder")  
		    
		    if self.TempFolder.Exists then self.TempFolder.RemoveFolderAndContents
		    
		    self.TempFolder.CreateFolder
		    
		    if not file.Exists then Return -1
		    
		    if file.IsFolder then return -2
		    
		  end if
		  
		  file.Unzip(self.TempFolder )
		  
		  return 0
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Sheets() As clWorksheet
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		TempFolder As FolderItem
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
