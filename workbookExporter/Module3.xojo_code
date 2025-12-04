#tag Module
Protected Module Module3
	#tag Method, Flags = &h0
		Sub addDefaultNodeToContent(topNode as xmlDocument, parentNode as XMLNode, extension as string, ContentType as string)
		  
		  var subNode as XMLNode
		  
		  subNode = parentNode.AppendChild(topNode.CreateElement("Default"))
		  subNode.SetAttribute("Extension",extension)
		  subNode.SetAttribute("ContentType",ContentType)
		  
		  Return
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub addOverrideNodeToContent(topNode as xmlDocument, parentNode as XMLNode, partName as string, ContentType as string)
		  
		  var subNode as XMLNode
		  
		  subNode = parentNode.AppendChild(topNode.CreateElement("Override"))
		  subNode.SetAttribute("PartName",partName)
		  subNode.SetAttribute("ContentType",ContentType)
		  
		  Return
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MakeFolder(fld as FolderItem) As FolderItem
		  
		  if not fld.Exists then fld.CreateFolder
		  
		  return fld
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub saveContentFile(targetFolder as FolderItem)
		  const filename as string = "[Content_Types].xml"
		  
		  var xmlDoc as new XMLDocument()
		  
		  Var typesNode As XmlNode
		  typesNode = xmlDoc.AppendChild(xmlDoc.CreateElement("Types"))
		  
		  addDefaultNodeToContent(xmlDoc, typesNode, "wmf","image/x-wmf")
		  addDefaultNodeToContent(xmlDoc, typesNode, "png", "image/png")
		  addDefaultNodeToContent(xmlDoc, typesNode, "xml", "application/xml")
		  addDefaultNodeToContent(xmlDoc, typesNode, "jpeg",  "image/jpeg")
		  addDefaultNodeToContent(xmlDoc, typesNode, "refs", "application/vnd.openxmlformats-package.relationships+xml")
		  addDefaultNodeToContent(xmlDoc, typesNode, "bin",  "application/vnd.openxmlformats-officedocument.oleObject")
		  
		  addOverrideNodeToContent(xmlDoc, typesNode, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
		  addOverrideNodeToContent(xmlDoc, typesNode, "/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
		  addOverrideNodeToContent(xmlDoc, typesNode, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
		  addOverrideNodeToContent(xmlDoc, typesNode, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
		  addOverrideNodeToContent(xmlDoc, typesNode, "/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
		  addOverrideNodeToContent(xmlDoc, typesNode, "/docProps/app.xml",  "application/vnd.openxmlformats-officedocument.extended-properties+xml")
		  
		  xmlDoc.SaveXML(targetFolder.Child(filename))
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub saveDocProps(targetFolder as FolderItem)
		  
		  saveTemplate(targetFolder, "app.xml", template_app)
		  
		  saveTemplate(targetFolder, "core.xml", template_core)
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub saveSharedStrings(targetFolder as FolderItem)
		  
		  const filename as string = "sharedStrings.xml"
		  
		  var dummyStrings() as string
		  
		  dummyStrings.Add("test1")
		  dummyStrings.Add("test2")
		  
		  //  
		  // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		  // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="101" uniqueCount="42">
		  // <si>
		  // <t>City</t>
		  // </si>
		  //
		  
		  var xmlDoc as new XMLDocument()
		  
		  Var typesNode As XmlNode
		  typesNode = xmlDoc.AppendChild(xmlDoc.CreateElement("sst"))
		  typesNode.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		  typesNode.SetAttribute("count", "100")
		  typesNode.SetAttribute("uniqueCount",str(dummyStrings.Count))
		  
		  for each s as string in dummyStrings
		    var level1 as XMLNode = typesNode.AppendChild(xmlDoc.CreateElement("si"))
		    
		    var level2 as XMLNode = level1.AppendChild(xmlDoc.CreateTextNode("t"))
		    
		    level2.Value = s
		    
		  next
		  
		  
		  xmlDoc.SaveXML(targetFolder.Child(filename))
		  
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub saveTemplate(targetFolder as FolderItem, filename as string, source_template as string)
		  
		  var file as FolderItem 
		  var tout as TextOutputStream
		  
		  
		  file = targetFolder.Child(filename)
		  tout = TextOutputStream.Create(file)
		  tout.Write(source_template)
		  tout.Close
		  
		  Return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SaveWorkbook(filename as string = "TestXLSXFile")
		  
		  var baseFolder as FolderItem = SpecialFolder.Desktop.child(filename)
		  
		  if baseFolder.Exists then baseFolder.RemoveFolderAndContents
		  
		  
		  
		  baseFolder = MakeFolder(baseFolder)
		  // 
		  // Create folder structure
		  //
		  
		  var workbookRelationFolder as FolderItem = MakeFolder(baseFolder.Child("_rels"))
		  
		  var workbookDocPropsFolder as FolderItem = MakeFolder(baseFolder.Child("docProps"))
		  
		  var worksheetBaseFolder as FolderItem = MakeFolder(baseFolder.Child("xl"))
		  
		  var worksheetRelationFolder as FolderItem = MakeFolder(worksheetBaseFolder.child("_rels"))
		  
		  var worksheetThemeFolder as FolderItem = MakeFolder(worksheetBaseFolder.child("theme"))
		  
		  var worksheetSheetsFolder as FolderItem = MakeFolder(worksheetBaseFolder.child("worksheets"))
		  
		  saveDocProps(workbookDocPropsFolder)
		  
		  saveTemplate(worksheetRelationFolder, "workbook.xml.rels", template_xl_rels)
		  
		  saveSharedStrings(worksheetBaseFolder)
		  
		  // Update folderTOC - called at the end
		  saveContentFile(baseFolder)
		  
		  
		  // Show in Finder
		  baseFolder.Open
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = Fold
		
		\_rels (done)
		- [Content_Types].xml (done)
		
		\docProps (done)
		- app.xml (done)
		- core.xml (done)
		- custom.xml
		
		\xl
		- sharedStrings.xml (done)
		- styles.xml
		- workbook.xml
		
		\xl\_rels
		- workbook.xml.rels
		
		\xl\theme
		- theme1.xml
		
		\xl\worksheets
		- sheetxx.xml
		
		
	#tag EndNote


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
End Module
#tag EndModule
