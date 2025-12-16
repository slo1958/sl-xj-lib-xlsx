#tag Class
Protected Class clWorkbookInformation
	#tag Method, Flags = &h0
		Sub Constructor(workbook as clWorkbook, DebuggingSet as clXLSX_Debugging)
		  self.WorkbookRef = workbook
		  self.DebuggingSettings = DebuggingSet
		  self.InCellLineBreak = " "
		  
		  self.Relations = new clFileRelations
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FindRelation(RelationId as string) As clFileRelationEntry
		  //
		  // Find an entry in relation table based on relation ID
		  //
		  // Used to find the path to sheet files
		  //
		  // Parameters:
		  // - relation id
		  //
		  // Returns:
		  // relation object or nil
		  //
		  
		  return self.Relations.GetEntryById(RelationId)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCellStyle(styleIndex as integer) As clCellXf
		  
		  //
		  // Find an entry in the style dictionary table based on style index
		  //
		  // Parameters:
		  // - style index
		  //
		  // Returns:
		  // style object or nil
		  //
		  
		  
		  return self.CellXf.Lookup(styleIndex, nil)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetFormat(formatIndex as integer) As clCellFormatter
		  
		  //
		  // Find an entry in the number format dictionaries table based on format index
		  //
		  // Parameters:
		  // - format index
		  //
		  // Returns:
		  // format object or nil
		  //
		  
		  
		  
		  if self.NumberingFormat.HasKey(formatIndex) then
		    return self.NumberingFormat.lookup(formatIndex, nil)
		    
		  else
		    Return self.CustomNumberingFormat.lookup(formatIndex, nil)
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSharedString(stringIndex as integer) As String
		  
		  //
		  // Find an entry in the shared string dictionary based on the string index
		  //
		  // Parameters:
		  // - string index
		  //
		  // Returns:
		  //  string (empty string if missing)
		  //
		  
		  
		  return self.SharedStrings.Lookup(stringIndex, "")
		  
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
		  
		  
		  var formatList(99) as string
		  
		  //
		  // All languages
		  //
		  
		  formatList(0) = "General"
		  formatList(1) = "0"
		  formatList(2) = "0.00"
		  formatList(3) = "#,##0"
		  formatList(4) = "#,##0.00"
		  
		  formatList(9) = "0%"
		  formatList(10) = "0.00%"
		  formatList(11) = "0.00E+00"
		  formatList(12) = "# ?/?"
		  formatList(13) = "# ??/??"
		  formatList(14) = "mm-dd-yy"
		  formatList(15) = "d-mmm-yy"
		  formatList(16) = "d-mmm"
		  formatList(17) = "mmm-yy"
		  formatList(18) = "h:mm AM/PM"
		  formatList(19) = "h:mm:ss AM/PM"
		  formatList(20) = "h:mm"
		  formatList(21) = "h:mm:ss"
		  formatList(22) = "m/d/yy h:mm"
		  
		  
		  formatList(37) = "#,##0 ;(#,##0)"
		  formatList(38) = "#,##0 ;[Red](#,##0)"
		  formatList(39) = "#,##0.00;(#,##0.00)"
		  formatList(40) = "#,##0.00;[Red](#,##0.00)"
		  
		  formatList(45) = "mm:ss"
		  formatList(46) = "[h]:mm:ss"
		  formatList(47) = "mmss.0"
		  formatList(48) = "##0.0E+0"
		  formatList(49) = "@"
		  
		  
		  //
		  // Aditional formats
		  //
		  formatList(5 ) = "$#,##0\-$#,##0"
		  formatList(6 ) = "$#,##0[Red]\-$#,##0"
		  formatList(7 ) = "$#,##0.00\-$#,##0.00"
		  formatList(8 ) = "$#,##0.00[Red]\-$#,##0.00"
		  
		  formatList(27 ) = "[$-404]e/m/d"
		  formatList(30 ) = "m/d/yy"
		  formatList(36 ) = "[$-404]e/m/d"
		  formatList(50 ) = "[$-404]e/m/d"
		  formatList(57 ) = "[$-404]e/m/d"
		  
		  formatList(59 ) = "t0"
		  formatList(60 ) = "t0.00"
		  formatList(61 ) = "t#,##0"
		  formatList(62 ) = "t#,##0.00"
		  formatList(67 ) = "t0%"
		  formatList(68 ) = "t0.00%"
		  formatList(69 ) = "t# ?/?"
		  formatList(70 ) = "t# ??/??"
		  
		  for index as integer = 0  to formatList.LastIndex
		    if formatList(index) = "" then
		      
		    else
		      NumberingFormat.value(index) = new clCellFormatter(formatList(index))
		      
		    end if
		    
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadDefinedNames(LoadMode as clWorkbook.LoadModes, XmlSheets as XMLNode, XmlLevel as integer)
		  
		  var x1 as xmlnode = XmlSheets
		  var lvl1 as integer = XmlLevel
		  
		  while x1 <> nil 
		    if self.WorkbookRef.TraceLoadWorkbook then  WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
		    if x1.name = "definedName" then
		      var name as string = x1.GetAttribute("name")
		      var range as string = x1.FirstChild.Value
		      var localID as string = x1.GetAttribute("localSheetId")
		      
		      var temp as new clWorkbookNamedRange(name, range, localID.ToInteger)
		      
		      NamedRanges.Add(temp)
		      
		      if self.TraceLoadNamedRanges then
		        if temp <> nil then Writelog(CurrentMethodName, 0, temp.SourceRange+" "+ temp.GetTargetSheetName )
		        Writelog(CurrentMethodName,0, "Loaded name range [" + name  + "] , loaclId " + localID + ", definition [" + range + "].")
		        
		      end if
		      
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSharedStrings(TemporaryFolder as FolderItem)
		  
		   
		  self.SharedStrings = new Dictionary
		  
		  var SharedStringXML as XMLDocument = self.WorkbookRef.OpenXMLDocument(TemporaryFolder, "xl","sharedStrings.xml")
		  
		  if SharedStringXML = nil then Return
		  
		  
		  var x2 as xmlnode = SharedStringXML.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer = 0
		  
		  while x2 <> nil 
		    if self.TraceLoadSharedStrings then WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
		    if x2.name = "si" then 
		      
		      var x3 as XMLNode = x2.FirstChild
		      var txt as string = ""
		      
		      while  x3 <> nil 
		        
		        if x3.name = "t" then
		          var d as integer
		          
		          if x3.FirstChild <> nil then txt = x3.FirstChild.value
		          
		        end if
		        
		        if x3.name = "r" then
		          var x4 as XmlNode = x3.FirstChild
		          
		          while x4 <> nil
		            if x4.name = "t" and x4.FirstChild <> nil then txt = txt + x4.FirstChild.Value
		            
		            x4 = x4.NextSibling
		            
		          wend
		          
		        end if
		        x3 = x3.NextSibling
		      wend
		      
		      
		      SharedStrings.value(strCounter) = txt.ReplaceAll(chr(10), self.InCellLineBreak)
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
		  
		  if self.TraceLoadSharedStrings then Writelog(CurrentMethodName,0, "Loaded " + str(strCounter)  + " items.")
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadStyleCellXfs(basenode as XMLNode)
		  
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

	#tag Method, Flags = &h1
		Protected Sub LoadStyleNumFmts(baseNode as XMLNode)
		  
		  self.CustomNumberingFormat = new Dictionary
		  
		  var x1 as xmlnode = basenode.FirstChild
		  
		  while x1 <> nil 
		    if self.TraceLoadStyles then WriteLog(CurrentMethodName,-1, x1.name)
		    
		    if x1.name = "numFmt" then
		      var formatcCode as string = x1.GetAttribute("formatCode")
		      var id as integer = x1.GetAttribute("numFmtId").ToInteger
		      
		      self.CustomNumberingFormat.value(id) = new clCellFormatter(formatcCode)
		      
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadStyles(TemporaryFolder as FolderItem)
		  
		  
		  var StyleXml as XMLDocument = self.WorkbookRef.OpenXMLDocument(TemporaryFolder, "xl","styles.xml")
		  
		  if StyleXml = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <styles.xml>")
		    Return
		  end if
		  
		  var x1 as xmlnode = StyleXml.FirstChild
		  var lvl1 as integer = 0
		  
		  
		  while x1 <> nil 
		    if self.TraceLoadStyles then Writelog(CurrentMethodName, lvl1, x1.name)
		    
		    if x1.name = "styleSheet" then x1 = x1.FirstChild
		    
		    if x1.name = "fonts" then
		      
		    elseif x1.name = "fills" then
		      
		    elseif x1.name = "borders" then
		      
		    elseif x1.name = "cellStyleXfs" then
		      self.LoadStyleXfs(x1)
		      
		    elseif x1.name = "cellXfs" then
		      self.LoadStyleCellXfs(x1)
		      
		    elseif x1.name = "cellStyles" then
		      
		      
		      
		    elseif x1.name = "numFmts" then  
		      self.LoadStyleNumFmts(x1)
		      
		    else
		      if self.TraceLoadStyles then Writelog(CurrentMethodName, lvl1, x1.name)
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadStyleXfs(basenode as XMLNode)
		  
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
		Sub LoadWorkBookRelations(TemporaryFolder as FolderItem)
		  
		  
		  
		  var WorkbookRelXML as XMLDocument = self.WorkbookRef.OpenXMLDocument(TemporaryFolder, "xl","_rels","workbook.xml.rels")
		  
		  if WorkbookRelXML = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <workbook.xml.rels>")
		    return
		    
		  end if
		  
		  var x2 as xmlnode = WorkbookRelXML.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer = 0
		  
		  while x2 <> nil 
		    if self.WorkbookRef.TraceLoadWorkbook then WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
		    if x2.name = "Relationship" then 
		      var relId as string = x2.GetAttribute("Id")
		      var reltype as string = x2.GetAttribute("Type")
		      var relTarget as string = x2.GetAttribute("Target")
		      
		      self.Relations.AddFileRelationEntry(relid, reltype, relTarget)
		      //self.Relations.add(new clWorkbookRelation(relId, reltype, relTarget))
		      
		    end if
		    
		    if x2.name = "Relationships" then
		      x2 = x2.FirstChild
		      
		    else 
		      x2 = x2.NextSibling
		      
		    end if
		    
		  wend
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TraceLoadNamedRanges() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadNamedRanges
		    
		    
		  end if
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TraceLoadSharedStrings() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadSharedStrings
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TraceLoadStyles() As boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadStyles
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Writelog(Source as string, level as integer, message as string)
		  self.WorkbookRef.Writelog(Source, level, message)
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CellXf As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		CustomNumberingFormat As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		DebuggingSettings As clXLSX_Debugging
	#tag EndProperty

	#tag Property, Flags = &h0
		ExpectedStringCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ExpectedStringUniqueCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		InCellLineBreak As string
	#tag EndProperty

	#tag Property, Flags = &h0
		NamedRanges() As clWorkbookNamedRange
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberingFormat As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Relations As clFileRelations
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedStrings As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		StyleXfs As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkbookRef As clWorkbook
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
			Name="WorkbookRef"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
