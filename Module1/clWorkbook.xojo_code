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
		Function GetSheet(name as string) As clWorksheet
		  
		  for each sheet as clWorksheet in self.sheets
		    if sheet.Name = name then Return sheet
		    
		  next
		  
		  return nil
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetNames() As string()
		  var ret() as string
		  
		  for each sheet as clWorksheet in self.sheets
		    ret.Add(sheet.Name)
		    
		  next
		  
		  return ret
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorkbookInfo()
		  
		  var tmp as FolderItem = self.TempFolder
		  
		  if tmp = nil then return
		  
		  
		  var workbookxml as XMLDocument = new XMLDocument(tmp.Child("xl").child("workbook.xml"))
		  
		  var x1 as xmlnode = workbookxml.FirstChild
		  var lvl1 as integer = 0
		  
		  while x1 <> nil 
		    System.DebugLog(str(lvl1)+":"+x1.name)
		    
		    if x1.name = "sheet" then
		      var name as string = x1.GetAttribute("name")
		      var sheetid as string = x1.GetAttribute("sheetId")
		      var rid as String = x1.GetAttribute("r:id")
		      sheets.Add(new clWorksheet( self.TempFolder, name, sheetid.ToInteger))
		      
		       
		      
		    end if
		    
		    // Navigate the tree
		    if x1.name ="workbook" then 
		      x1 = x1.FirstChild
		      lvl1 = lvl1+1
		      
		    elseif x1.name = "sheets" then
		      x1 = x1.FirstChild
		      lvl1 = lvl1 + 1
		      
		    else
		      x1 = x1.NextSibling
		      
		    end if
		    
		  wend
		  
		  
		  var sharedstringxml as XMLDocument = new XMLDocument(tmp.Child("xl").child("sharedStrings.xml"))
		  
		  var x2 as xmlnode = sharedstringxml.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer
		  
		  while x2 <> nil 
		    System.DebugLog(str(lvl2)+":"+x2.name)
		    
		    if x2.name = "si" then 
		      
		      var x3 as XMLNode = x2.FirstChild
		       
		      
		      while sharedstrings.LastIndex < strCounter 
		        SharedStrings.Add("??")
		        
		      wend
		      
		      SharedStrings(strCounter) = x3.FirstChild.Value
		      strCounter = strCounter + 1
		      
		    end if
		    
		    if x2.name = "sst" then
		      self.ExpectedStringUniqueCount = x2.GetAttribute("uniqueCount").ToInteger
		      self.ExpectedStringCount = x2.GetAttribute("count").ToInteger
		      
		      x2 = x2.FirstChild
		      
		    else
		      x2 = x2.NextSibling
		      
		    end if
		    
		  wend
		  
		  return
		  
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
		ExpectedStringCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ExpectedStringUniqueCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedStrings() As String
	#tag EndProperty

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
