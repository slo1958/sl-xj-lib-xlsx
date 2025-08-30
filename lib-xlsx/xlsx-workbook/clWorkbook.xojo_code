#tag Class
Protected Class clWorkbook
	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, DebugOptions as clXLSX_Debugging, LoadMode as LoadModes = LoadModes.FullLoad, language as string = "", workfolder as FolderItem = nil)
		  //
		  // Initialize and  load information about the workbook
		  // - Relations
		  // - shared strings (assuming fixed path and file name, should take from relations)
		  // - Styles (assuming fixed path and file name, should take from relations)
		  // - Sheet references
		  //   Sheet data are loaded is load mode is not SheetOnDemand
		  //
		  // Parameters:
		  // - source file (expected to be a .xlsx file)
		  // - debug options
		  // - loadmode 
		  // - language info (not used)
		  // - path to workfolder (or nil)
		  //
		  // Returns
		  // (nothing)
		  //
		  
		  self.InCellLineBreak = " "
		  
		  self.DebuggingSettings = DebugOptions
		  self.SourceFile = file
		  self.TempFolder = workfolder
		  
		  self.InitInternals(language)
		  
		  if LoadMode = LoadModes.Manual then return
		  
		  self.load(LoadMode)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, LoadMode as LoadModes = LoadModes.FullLoad, language as string = "", workfolder as FolderItem = nil)
		  //
		  // Initialize and  load information about the workbook
		  // - Relations
		  // - shared strings (assuming fixed path and file name, should take from relations)
		  // - Styles (assuming fixed path and file name, should take from relations)
		  // - Sheet references
		  //   Sheet data are loaded is load mode is not SheetOnDemand
		  //
		  // Parameters:
		  // - source file (expected to be a .xlsx file)
		  // - loadmode 
		  // - language info (not used)
		  // - path to workfolder (or nil)
		  //
		  // Returns
		  // (nothing)
		  //
		  
		  self.InCellLineBreak = " "
		  
		  self.DebuggingSettings = new clXLSX_Debugging
		  
		  self.SourceFile = file
		  self.TempFolder = workfolder
		  
		  self.InitInternals(language)
		  
		  if LoadMode = LoadModes.Manual then return
		  
		  self.load(LoadMode)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DropSheetFromMemory(sheetName as string)
		  
		  //
		  // Find an entry in the table of worksheetref  based on sheet name
		  //
		  // Parameters:
		  // - sheet name
		  //
		  // Returns:
		  //  clWorksheet object or nil
		  //
		  
		  
		  var SheetRef as clWorkSheetRef
		  
		  for each sheet as clWorkSheetRef in self.SheetRefs
		    if sheet.Name = sheetName then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return
		  
		  SheetRef.DropData
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FindRelation(RelationId as string) As clWorkbookRelation
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
		  
		  for each rel as clWorkbookRelation in self.Relations
		    if rel.ID = RelationId then return rel
		    
		  next
		  
		  return nil
		  
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
		Function GetSheetFromID(sheetID as integer) As clWorksheet
		  
		  //
		  // Find an entry in the table of worksheetref  based on sheet name
		  //
		  // Parameters:
		  // - sheet ID
		  //
		  // Returns:
		  //  clWorksheet object or nil
		  //
		  
		  
		  var SheetRef as clWorkSheetRef
		  
		  for each sheet as clWorkSheetRef in self.SheetRefs
		    if sheet.Id = sheetID then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return nil
		  
		  return SheetRef.GetSheetData
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetFromLocalID(localID as integer) As clWorksheet
		  
		  if localID < 0 then Return nil
		  
		  if localID > self.SheetRefs.LastIndex then return nil
		  
		  var tmp as clWorkSheetRef = self.SheetRefs(localID)
		  
		  return tmp.GetSheetData
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetFromName(sheetName as string) As clWorksheet
		  
		  //
		  // Find an entry in the table of worksheetref  based on sheet name
		  //
		  // Parameters:
		  // - sheet name
		  //
		  // Returns:
		  //  clWorksheet object or nil
		  //
		  
		  
		  var SheetRef as clWorkSheetRef
		  
		  for each sheet as clWorkSheetRef in self.SheetRefs
		    if sheet.Name = sheetName then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return nil
		  
		  return SheetRef.GetSheetData
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetNames() As string()
		  //
		  // Get the list of known sheet names
		  //
		  // Parameters
		  // (nothing)
		  //
		  // Returns
		  // array of strings, sheet name
		  //
		  var ret() as string
		  
		  for each sheet as clWorkSheetRef in self.SheetRefs
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
		Sub Load(LoadMode as LoadModes)
		  //
		  // load information about the workbook
		  // - Relations
		  // - shared strings (assuming fixed path and file name, should take from relations)
		  // - Styles (assuming fixed path and file name, should take from relations)
		  // - Sheet references
		  //   Sheet data are loaded is load mode is not SheetOnDemand
		  //
		  // Parameters:
		  // - loadmode 
		  //
		  // Returns
		  // (nothing)
		  
		  if self.UnzipToTemporary <> 0 then
		    return
		    
		  end if
		  
		  self.LoadWorkBookRelations()
		  
		  self.LoadSharedStrings() 
		  self.LoadStyles()
		  
		  
		  self.LoadWorkbookInfo(LoadMode)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadDefinedNames(LoadMode as LoadModes, XmlSheets as XMLNode, XmlLevel as integer)
		  
		  var x1 as xmlnode = XmlSheets
		  var lvl1 as integer = XmlLevel
		  
		  while x1 <> nil 
		    if self.TraceLoadWorkbook then  clWorkbook.WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
		    if x1.name = "definedName" then
		      var name as string = x1.GetAttribute("name")
		      var range as string = x1.FirstChild.Value
		      var localID as string = x1.GetAttribute("localSheetId")
		      
		      NamedRanges.Add(new clWorkbookNamedRange(name, range, localID.ToInteger))
		      
		      
		      if self.TraceLoadNamedRanges then
		        Writelog(CurrentMethodName,0, "Loaded name range [" + name  + "] , loaclId " + localID + ", definition [" + range + "].")
		        
		      end if
		      
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadListWorksheets(LoadMode as LoadModes, XmlSheets as XMLNode, XmlLevel as integer)
		  
		  var x1 as xmlnode = XmlSheets
		  var lvl1 as integer = XmlLevel
		  
		  while x1 <> nil 
		    if self.TraceLoadWorkbook then  clWorkbook.WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
		    if x1.name = "sheet" then
		      var name as string = x1.GetAttribute("name")
		      var sheetid as string = x1.GetAttribute("sheetId")
		      var RelationId as String = x1.GetAttribute("r:id")
		      var RelationTarget as string = self.FindRelation(RelationId).Target
		      self.SheetRefs.add(new clWorkSheetRef(self.TempFolder, name, SheetId.ToInteger, RelationId, RelationTarget, self.TraceLoadSheetData))
		      
		      if self.TraceLoadSheetListInWorkbook then
		        Writelog(CurrentMethodName,0, "Loaded sheet [" + name  + "] , id " + str(sheetId) + ".")
		        
		      end if
		      
		      if LoadMode =LoadModes.FullLoad then
		        Self.SheetRefs(self.SheetRefs.LastIndex).LoadSheetData()
		        
		      end if
		      
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
		    if TraceLoadSharedStrings then clWorkbook.WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
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
		    if self.TraceLoadStyles then clWorkbook.WriteLog(CurrentMethodName,-1, x1.name)
		    
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

	#tag Method, Flags = &h1
		Protected Sub LoadStyles()
		  
		  
		  var StyleXml as XMLDocument = self.OpenDocument(self.TempFolder, "xl","styles.xml")
		  
		  if StyleXml = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <styles.xml>")
		    Return
		  end if
		  
		  var x1 as xmlnode = StyleXml.FirstChild
		  var lvl1 as integer = 0
		  
		  
		  while x1 <> nil 
		    if self.TraceLoadStyles then clWorkbook.Writelog(CurrentMethodName, lvl1, x1.name)
		    
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
		      if self.TraceLoadStyles then clWorkbook.Writelog(CurrentMethodName, lvl1, x1.name)
		      
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
		Protected Sub LoadWorkbookInfo(LoadMode as LoadModes)
		  //
		  // load information from workbook.xml
		  // - Sheet references
		  //   Sheet data are loaded is load mode is not SheetOnDemand
		  //
		  // Parameters:
		  // - loadmode 
		  //
		  // Returns
		  // (nothing)
		  //
		  
		  var WorkbookXML as XMLDocument = self.OpenDocument(self.TempFolder, "xl","workbook.xml")
		  
		  if WorkbookXML = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <workbook.xml>")
		    return
		    
		  end if
		  
		  var x1 as xmlnode = WorkbookXML.FirstChild
		  var lvl1 as integer = 0
		  
		  while x1 <> nil 
		    if self.TraceLoadWorkbook then  clWorkbook.WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
		    
		    // Navigate the tree
		    if x1.name ="workbook" then 
		      x1 = x1.FirstChild
		      lvl1 = lvl1+1
		      
		    elseif x1.name = "sheets" then
		      LoadListWorksheets(LoadMode, x1.FirstChild, lvl1+1)
		      x1 = x1.NextSibling
		      
		    elseif x1.name = "definedNames" then
		      LoadDefinedNames(LoadMode, x1.FirstChild, lvl1+1)
		      x1 = x1.NextSibling
		      
		    else
		      x1 = x1.NextSibling
		      
		    end if
		  wend
		  
		  
		  
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadWorkbookInfo_old(LoadMode as LoadModes)
		  //
		  // load information from workbook.xml
		  // - Sheet references
		  //   Sheet data are loaded is load mode is not SheetOnDemand
		  //
		  // Parameters:
		  // - loadmode 
		  //
		  // Returns
		  // (nothing)
		  //
		  
		  var WorkbookXML as XMLDocument = self.OpenDocument(self.TempFolder, "xl","workbook.xml")
		  
		  if WorkbookXML = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <workbook.xml>")
		    return
		    
		  end if
		  
		  var x1 as xmlnode = WorkbookXML.FirstChild
		  var lvl1 as integer = 0
		  
		  while x1 <> nil 
		    if self.TraceLoadWorkbook then  clWorkbook.WriteLog(CurrentMethodName ,lvl1, x1.name)
		    
		    if x1.name = "sheet" then
		      var name as string = x1.GetAttribute("name")
		      var sheetid as string = x1.GetAttribute("sheetId")
		      var RelationId as String = x1.GetAttribute("r:id")
		      var RelationTarget as string = self.FindRelation(RelationId).Target
		      self.SheetRefs.add(new clWorkSheetRef(self.TempFolder, name, SheetId.ToInteger, RelationId, RelationTarget, self.TraceLoadSheetData))
		      
		      if LoadMode =LoadModes.FullLoad then
		        Self.SheetRefs(self.SheetRefs.LastIndex).LoadSheetData()
		        
		      end if
		      
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

	#tag Method, Flags = &h1
		Protected Sub LoadWorkBookRelations()
		  
		  
		  
		  var WorkbookRelXML as XMLDocument = self.OpenDocument(self.TempFolder, "xl","_rels","workbook.xml.rels")
		  
		  if WorkbookRelXML = nil then 
		    self.Writelog(CurrentMethodName,-1,"Cannot find <workbook.xml.rels>")
		    return
		    
		  end if
		  
		  var x2 as xmlnode = WorkbookRelXML.FirstChild
		  var lvl2 as integer = 0
		  var strCounter as integer = 0
		  
		  while x2 <> nil 
		    if self.TraceLoadWorkbook then clWorkbook.WriteLog(CurrentMethodName ,lvl2, x2.name)
		    
		    if x2.name = "Relationship" then 
		      var relId as string = x2.GetAttribute("Id")
		      var reltype as string = x2.GetAttribute("Type")
		      var relTarget as string = x2.GetAttribute("Target")
		      
		      self.Relations.add(new clWorkbookRelation(relId, reltype, relTarget))
		      
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

	#tag Method, Flags = &h0
		Function TraceLoadNamedRanges() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadNamedRanges
		    
		    
		  end if
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceLoadSharedStrings() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadSharedStrings
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceLoadSheetData() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadSheetData
		    
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceLoadSheetListInWorkbook() As boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadSheetListInWorkbook
		     
		  end if
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceLoadStyles() As boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadStyles
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceLoadWorkbook() As Boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceLoadWorkbook
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TraceShowUnzippedWorkbool() As boolean
		  if self.DebuggingSettings = nil then
		    return false
		    
		  else
		    return self.DebuggingSettings.TraceShowUnzippedWorkbool
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnzipToTemporary() As integer
		  
		  var file as FolderItem = self.SourceFile
		  
		  if self.TempFolder = nil then
		    
		    self.TempFolder  =  SpecialFolder.Temporary
		    
		    // prepare work area
		    
		    self.TempFolder = self.TempFolder.Child(file.name.ReplaceAll(".","-") + " folder")  
		    
		    if self.TempFolder.Exists then self.TempFolder.RemoveFolderAndContents
		    
		    if not file.Exists then Return -1
		    
		    if file.IsFolder then return -2
		    
		  end if
		  
		  if not self.TempFolder.Exists then self.TempFolder.CreateFolder
		  
		  file.Unzip(self.TempFolder )
		  
		  if TraceShowUnzippedWorkbool then 
		    self.TempFolder.Open
		    
		  end if
		  
		  return 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub Writelog(Source as string, level as integer, message as string)
		  
		  
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
		Protected Relations() As clWorkbookRelation
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedStrings As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		SheetRefs() As clWorkSheetRef
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

	#tag Property, Flags = &h0
		zz As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		zzz As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		zzzzz As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		zzzzzz As Boolean
	#tag EndProperty


	#tag Enum, Name = LoadModes, Type = Integer, Flags = &h0
		Manual
		  FullLoad
		LoadSheetOnDemand
	#tag EndEnum


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
		#tag ViewProperty
			Name="InCellLineBreak"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="zzz"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="zz"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="zzzzz"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="zzzzzz"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
