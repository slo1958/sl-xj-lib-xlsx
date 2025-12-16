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
		  
		  self.Information = new clWorkbookInformation(self, self.DebuggingSettings)
		  
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
		   
		  self.DebuggingSettings = new clXLSX_Debugging
		  
		  self.Information = new clWorkbookInformation(self, self.DebuggingSettings)
		  
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
		  
		  
		  var SheetRef as clWorksheetRef
		  
		  for each sheet as clWorksheetRef in self.SheetRefs
		    if sheet.Name = sheetName then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return
		  
		  SheetRef.DropData
		  
		  
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
		  
		  if self.Information = nil then 
		    return nil
		    
		  else
		    return self.Information.FindRelation(RelationId)
		     
		  end if
		  
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
		  
		  
		  return self.Information.GetCellStyle(styleIndex)
		  
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
		  
		  return self.Information.GetFormat(formatIndex)
		   
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
		  
		  if self.Information = nil then
		    return "?no info"
		    
		  else
		    return self.Information.GetSharedString(stringIndex)
		    
		  end if
		  
		   
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
		  
		  
		  var SheetRef as clWorksheetRef
		  
		  for each sheet as clWorksheetRef in self.SheetRefs
		    if sheet.Id = sheetID then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return nil
		  
		  return SheetRef.GetSheetData(Information)
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetFromLocalID(localID as integer) As clWorksheet
		  
		  if localID < 0 then Return nil
		  
		  if localID > self.SheetRefs.LastIndex then return nil
		  
		  var tmp as clWorksheetRef = self.SheetRefs(localID)
		  
		  return tmp.GetSheetData(Information)
		  
		  
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
		  
		  
		  var SheetRef as clWorksheetRef
		  
		  for each sheet as clWorksheetRef in self.SheetRefs
		    if sheet.Name = sheetName then SheetRef = sheet
		    
		  next
		  
		  if SheetRef = nil then return nil
		  
		  return SheetRef.GetSheetData(Information)
		  
		  
		  
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
		  
		  for each sheet as clWorksheetRef in self.SheetRefs
		    ret.Add(sheet.Name)
		    
		  next
		  
		  return ret
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSourceFileName() As string
		  
		  return self.SourceFile.Name
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
		  
		  self.Information.InitInternals(language)
		  
		  return
		   
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
		  
		  
		  self.Information.LoadWorkBookRelations(self.TempFolder)
		  
		  self.Information.LoadSharedStrings(self.TempFolder)
		  
		  self.Information.LoadStyles(self.TempFolder)
		  
		  self.LoadWorkbookDetails(LoadMode)
		  
		  Return
		  
		  
		  
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
		      
		      self.SheetRefs.add(new clWorksheetRef(self.TempFolder, name, SheetId.ToInteger, RelationId, RelationTarget, self.TraceLoadSheetData))
		      
		      if self.TraceLoadSheetListInWorkbook then
		        Writelog(CurrentMethodName,0, "Loaded sheet [" + name  + "] , id " + str(sheetId) + ".")
		        
		      end if
		      
		      if LoadMode =LoadModes.FullLoad then
		        Self.SheetRefs(self.SheetRefs.LastIndex).LoadSheetData(Information)
		        
		      end if
		      
		    end if
		    
		    x1 = x1.NextSibling
		    
		  wend
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LoadWorkbookDetails(LoadMode as LoadModes)
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
		  
		  var WorkbookXML as XMLDocument = self.OpenXMLDocument(self.TempFolder, "xl","workbook.xml")
		  
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
		      self.Information.LoadDefinedNames(LoadMode, x1.FirstChild, lvl1+1)
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
		  
		  var WorkbookXML as XMLDocument = self.OpenXMLDocument(self.TempFolder, "xl","workbook.xml")
		  
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
		      self.SheetRefs.add(new clWorksheetRef(self.TempFolder, name, SheetId.ToInteger, RelationId, RelationTarget, self.TraceLoadSheetData))
		      
		      if LoadMode =LoadModes.FullLoad then
		        Self.SheetRefs(self.SheetRefs.LastIndex).LoadSheetData(Information)
		        
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

	#tag Method, Flags = &h0
		Shared Function OpenXMLDocument(BaseFolder as FolderItem, paramarray levels as string) As XMLDocument
		  
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


	#tag Note, Name = definedNames
		
		
		Examples
		
		<definedNames>  
		  <definedName name="NamedFormula"    comment="Comment text for defined name.">SUM(Sheet3!$B$2:$B$9)</definedName>  
		  <definedName name="NamedRange">Sheet3!$A$1:$C$12</definedName>  
		  <definedName name="NamedRangeFromExternalReference" localSheetId="2"     hidden="1">Sheet5!$A$1:$T$47</definedName>  
		</definedNames>
		
		<definedNames>
		  <definedName name="FMLA">Sheet1!$B$3</definedName>
		  <definedName name="SheetLevelName" comment="This name is scoped to Sheet1" localSheetId="0">Sheet1!$B$3</definedName>
		</definedNames>
		
		
		
	#tag EndNote

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
		DebuggingSettings As clXLSX_Debugging
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Information As clWorkbookInformation
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected SheetRefs() As clWorksheetRef
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected SourceFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected TempFolder As FolderItem
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
	#tag EndViewBehavior
End Class
#tag EndClass
