#tag Class
Protected Class clWorkbook
	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, workfolder as FolderItem = nil, DeferredLoad as boolean = False, language as string = "")
		  
		  self.InCellLineBreak = " "
		  
		  self.SourceFile = file
		  self.TempFolder = workfolder
		  
		  self.InitInternals(language)
		  
		  if not DeferredLoad then
		    self.load
		    
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCellStyle(styleIndex as integer) As clCellXf
		  
		  return self.CellXf.Lookup(styleIndex, nil)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetFormat(formatIndex as integer) As clCellFormatter
		  
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

	#tag Method, Flags = &h1
		Protected Sub InitInternals(language as string)
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
		Sub Load()
		  
		  
		  if self.UnzipToTemporary <> 0 then
		    return
		    
		  end if
		  
		  
		  
		  self.LoadSharedStrings()
		  self.LoadStyles()
		  self.LoadWorkbookInfo()
		  
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

	#tag Method, Flags = &h1
		Protected Sub LoadNumFmts(baseNode as XMLNode)
		  
		  self.CustomNumberingFormat = new Dictionary
		  
		  var x1 as xmlnode = basenode.FirstChild
		  
		  while x1 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName,-1, x1.name)
		    
		    if x1.name = "numFmt" then
		      var formatcCode as string = x1.GetAttribute("formatCode")
		      var id as integer = x1.GetAttribute("numFmtId").ToInteger
		      
		      self.CustomNumberingFormat.value(id) = new clCellFormatter(formatcCode)
		      
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadSharedStrings()
		  
		  self.SharedStrings = new Dictionary
		  
		  var SharedStringXML as XMLDocument = self.OpenDocument(self.TempFolder, "xl","sharedStrings.xml")
		  
		  if SharedStringXML = nil then Return
		  
		  
		  var x2 as xmlnode = SharedStringXML.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer = 0
		  
		  while x2 <> nil 
		    clWorkbook.WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
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
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadStyles()
		  
		  
		  var StyleXml as XMLDocument = self.OpenDocument(self.TempFolder, "xl","styles.xml")
		  
		  if StyleXml = nil then Return
		  
		  
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

	#tag Method, Flags = &h1
		Protected Sub LoadWorkbookInfo()
		  
		  
		  var WorkbookXML as XMLDocument = self.OpenDocument(self.TempFolder, "xl","workbook.xml")
		  
		  if WorkbookXML = nil then return 
		  
		  
		  var x1 as xmlnode = WorkbookXML.FirstChild
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
		Shared Function OpenDocument(BaseFolder as FolderItem, paramarray levels as string) As XMLDocument
		  
		  if BaseFolder = nil then return nil
		  
		  
		  var fld as FolderItem = BaseFolder
		  
		  for each level as string in levels
		    if not fld.Child(level).Exists then return nil
		    fld = fld.child(level)
		    
		  next
		  
		  return new XMLDocument(fld)
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnzipToTemporary() As integer
		  
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


	#tag Note, Name = Loading shared string
		
		Base (simple) case:
		
		
		Under the tag "sst", we have an array of "si", each <si> </si> entry is one string. Array is implicitely indexed, starting at zero
		
		<sst ....>
		   <si>
		      <t>simple text</t>
		   </si>
		</sst>
		
		
		A text may be formatted and contain characters from different fonts, for example the text "Company公司名:" is stored as:
		
		<si>
		        <r>
		            <rPr>
		                <b></b>
		                <sz val="12"></sz>
		                <color rgb="FF000000"></color>
		                <rFont val="Arial"></rFont>
		                <family val="2"></family>
		            </rPr>
		            <t>Company</t>
		        </r>
		        <r>
		            <rPr>
		                <b></b>
		                <sz val="12"></sz>
		                <color rgb="FF000000"></color>
		                <rFont val="宋体"></rFont>
		                <charset val="134"></charset>
		            </rPr>
		            <t>公司名</t>
		        </r>
		        <r>
		            <rPr>
		                <b></b>
		                <sz val="12"></sz>
		                <color rgb="FF000000"></color>
		                <rFont val="Arial"></rFont>
		                <family val="2"></family>
		            </rPr>
		            <t>:</t>
		        </r>
		    </si>
		
		
		We concatenate all elements of the <t>..</t> found in the <si> ..</si> without taking the font selection or the formatting information (color, size, ...)
		
		'In cell' line breakj (chr(10) are replaced by a single space. 
		
		
		
		
		
		
	#tag EndNote


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
		InCellLineBreak As string
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
