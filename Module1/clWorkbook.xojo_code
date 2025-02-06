#tag Class
Protected Class clWorkbook
	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, workfolder as FolderItem = nil, language as string = "")
		  
		  self.SourceFile = file
		  self.TempFolder = workfolder
		  
		  if self.UnzipToTemporary <> 0 then
		    return
		    
		  end if
		  
		  self.InitInternals(language)
		  
		  self.LoadSharedStrings()
		  self.LoadStyles()
		  self.LoadWorkbookInfo()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCellStyle(styleIndex as integer) As clCellXf
		  
		  return self.CellXf.Lookup(styleIndex, nil)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetFormat(formatIndex as integer) As string
		  
		  if self.NumberingFormat.HasKey(formatIndex) then
		    return self.NumberingFormat.lookup(formatIndex, "")
		    
		  else
		    Return self.CustomNumberingFormat.lookup(formatIndex, "")
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSharedString(stringIndex as integer) As String
		  
		  return self.SharedStrings.Lookup(stringIndex, "")
		  
		End Function
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
		Sub InitInternals(language as string)
		  //
		  //
		  // check language
		  //
		  // use as template to create language specific elements
		  //
		  
		  select case language
		    
		  case "zh-tw"
		    
		  case "zh-cn"
		    
		  case "ja-jp"
		    
		  case "ko-kr"
		    
		  case "th-th"
		    
		  case else
		    
		  end select
		  
		  
		  
		  // Create default numbering format
		  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
		  //
		  // numFmt (Number Format)
		  // This element specifies number format properties which indicate how to format and render the numeric value of a cell.
		  // Following is a listing of number formats whose formatCode value is implied rather than explicitly saved in the file. In this case a numFmtId value is written on the xf record, but no corresponding numFmt element is written. Some of these Ids can be interpreted differently, depending on the UI language of the implementing application.
		  // Ids not specified in the listing, such as 5, 6, 7, and 8, shall follow the number format specified by the formatCode attribute.
		  //
		  // The primary goal when a cell is using "General" formatting is to render the cell content without user-specified guidance
		  //  to the best ability of the application.
		  //
		  // Language specific formats are not created
		  //
		  NumberingFormat = new Dictionary
		  
		  //
		  // All languages
		  //
		  numberingformat.value(0) = "General"
		  numberingformat.value(1) = "0"
		  numberingformat.value(2) = "0.00"
		  numberingformat.value(3) = "#,##0"
		  numberingformat.value(4) = "#,##0.00"
		  
		  numberingformat.value(9) = "0%"
		  numberingformat.value(10) = "0.00%"
		  numberingformat.value(11) = "0.00E+00"
		  numberingformat.value(12) = "# ?/?"
		  numberingformat.value(13) = "# ??/??"
		  numberingformat.value(14) = "mm-dd-yy"
		  numberingformat.value(15) = "d-mmm-yy"
		  numberingformat.value(16) = "d-mmm"
		  numberingformat.value(17) = "mmm-yy"
		  numberingformat.value(18) = "h:mm AM/PM"
		  numberingformat.value(19) = "h:mm:ss AM/PM"
		  numberingformat.value(20) = "h:mm"
		  numberingformat.value(21) = "h:mm:ss"
		  numberingformat.value(22) = "m/d/yy h:mm"
		  
		  
		  numberingformat.value(37) = "#,##0 ;(#,##0)"
		  numberingformat.value(38) = "#,##0 ;[Red](#,##0)"
		  numberingformat.value(39) = "#,##0.00;(#,##0.00)"
		  numberingformat.value(40) = "#,##0.00;[Red](#,##0.00)"
		  
		  numberingformat.value(45) = "mm:ss"
		  numberingformat.value(46) = "[h]:mm:ss"
		  numberingformat.value(47) = "mmss.0"
		  numberingformat.value(48) = "##0.0E+0"
		  numberingformat.value(49) = "@"
		  
		  
		  //
		  // Aditional formats
		  //
		  numberingformat.value(5 ) = "$#,##0\-$#,##0"
		  numberingformat.value(6 ) = "$#,##0[Red]\-$#,##0"
		  numberingformat.value(7 ) = "$#,##0.00\-$#,##0.00"
		  numberingformat.value(8 ) = "$#,##0.00[Red]\-$#,##0.00"
		  
		  numberingformat.value(27 ) = "[$-404]e/m/d"
		  numberingformat.value(30 ) = "m/d/yy"
		  numberingformat.value(36 ) = "[$-404]e/m/d"
		  numberingformat.value(50 ) = "[$-404]e/m/d"
		  numberingformat.value(57 ) = "[$-404]e/m/d"
		  
		  numberingformat.value(59 ) = "t0"
		  numberingformat.value(60 ) = "t0.00"
		  numberingformat.value(61 ) = "t#,##0"
		  numberingformat.value(62 ) = "t#,##0.00"
		  numberingformat.value(67 ) = "t0%"
		  numberingformat.value(68 ) = "t0.00%"
		  numberingformat.value(69 ) = "t# ?/?"
		  numberingformat.value(70 ) = "t# ??/??"
		  
		  Return
		   
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadCellXfs(basenode as XMLNode)
		  
		  CellXf  = new Dictionary
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  
		  // Zero based array
		  var xfcount as integer = 0
		  
		  while x1 <> nil
		    if x1.name = "xf" then
		      CellXf.Value(xfcount) = new clCellXf( false, x1)
		      
		      xfcount = xfcount + 1
		      
		    end if
		    
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadNumFmts(baseNode as XMLNode)
		  
		  self.CustomNumberingFormat = new Dictionary
		  
		  var x1 as xmlnode = basenode.FirstChild
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName,-1, x1.name)
		    
		    if x1.name = "numFmt" then
		      var formatcCode as string = x1.GetAttribute("formatCode")
		      var id as integer = x1.GetAttribute("numFmtId").ToInteger
		      
		      self.CustomNumberingFormat.value(id) = formatcCode
		      
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSharedStrings()
		  
		  self.SharedStrings = new Dictionary
		  
		  var tmp as FolderItem = self.TempFolder
		  
		  if tmp = nil then return
		  
		  
		  var sharedstringxml as XMLDocument = new XMLDocument(tmp.Child("xl").child("sharedStrings.xml"))
		  
		  
		  var x2 as xmlnode = sharedstringxml.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer
		  
		  while x2 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
		    if x2.name = "si" then 
		      
		      var x3 as XMLNode = x2.FirstChild
		      
		      SharedStrings.value(strCounter) = x3.FirstChild.Value
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
		Sub LoadStyles()
		  
		  var tmp as FolderItem = self.TempFolder
		  
		  if tmp = nil then return
		  
		  var StyleXml as XMLDocument = new XMLDocument(tmp.Child("xl").child("styles.xml"))
		  
		  var x1 as xmlnode = StyleXml.FirstChild
		  var lvl1 as integer = 0
		  
		  
		  while x1 <> nil 
		    clWorkbook.Writelog(CurrentMethodName, lvl1, x1.name)
		    
		    if x1.name = "styleSheet" then x1 = x1.FirstChild
		    
		    if x1.name = "fonts" then
		      
		    elseif x1.name = "fills" then
		      
		    elseif x1.name = "borders" then
		      
		    elseif x1.name = "cellStyleXfs" then
		      self.LoadStyleXfs(x1)
		      
		    elseif x1.name = "cellXfs" then
		      self.LoadCellXfs(x1)
		      
		    elseif x1.name = "cellStyles" then
		      
		      
		      
		    elseif x1.name = "numFmts" then  
		      self.LoadNumFmts(x1)
		      
		    else
		      System.DebugLog(">>>" + CurrentMethodName+":" + x1.name)
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadStyleXfs(basenode as XMLNode)
		  
		  StyleXfs  = new Dictionary
		  
		  var x1 as xmlnode = basenode.FirstChild
		  var lvl as integer = 0
		  var xfcount as integer = 1
		  
		  while x1 <> nil
		    if x1.name = "xf" then
		      StyleXfs.Value(xfcount) = new clCellXf( true, x1)
		      
		      xfcount = xfcount + 1
		      
		    end if
		    
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorkbookInfo()
		  
		  var tmp as FolderItem = self.TempFolder
		  
		  if tmp = nil then return
		  
		  
		  var workbookxml as XMLDocument = new XMLDocument(tmp.Child("xl").child("workbook.xml"))
		  
		  var x1 as xmlnode = workbookxml.FirstChild
		  var lvl1 as integer = 0
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
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

	#tag Method, Flags = &h0
		Shared Sub Writelog(Source as string, level as integer, message as string)
		  var ignored() as string
		  
		  ignored.add("Module1.clWorksheet.")
		  ignored.add("Module1.clWorkbook.LoadSharedString")
		  ignored.add("Module1.clWorkbook.LoadWorkbookInfo")
		  
		  for each ignore as string in ignored
		    if source.IndexOf(ignore) >=0 then return
		    
		  next
		  
		  System.DebugLog(source+", " + "level " + str(level)+ ": " + message)
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CellXf As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		CustomNumberingFormat As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		ExpectedStringCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ExpectedStringUniqueCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberingFormat As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedStrings As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		Sheets() As clWorksheet
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		StyleXfs As Dictionary
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
		#tag ViewProperty
			Name="ExpectedStringCount"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ExpectedStringUniqueCount"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
