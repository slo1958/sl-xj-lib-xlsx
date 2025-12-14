#tag Class
Protected Class clFileTypeInfo
	#tag Method, Flags = &h0
		Shared Sub AddFileTypeEntry(RelationshipId as string, ContentType as ContentTypes, Filename as string, paramarray FileParentPath as string)
		  var c as new clFileTypeInfo
		  c.Filename = Filename
		  for each s as string in FileParentPath
		    c.FileParentPath.Add(s)
		    
		  next
		  
		  c.ShortContentType = ContentType
		  c.RelationShipID = RelationshipId
		  
		  FileTypeDir.Add(c)
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ContentType() As string
		  select case ShortContentType
		    
		  case ContentTypes.app
		    return "application/vnd.openxmlformats-officedocument.extended-properties+xml"
		    
		  case ContentTypes.core
		    Return "application/vnd.openxmlformats-package.core-properties+xml"
		    
		  case ContentTypes.sharedStrings
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
		    
		  case ContentTypes.style
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
		    
		  case ContentTypes.theme
		    Return "application/vnd.openxmlformats-officedocument.theme+xml"
		    
		  case ContentTypes.workbook
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
		    
		  case ContentTypes.worksheet
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
		    
		  case else
		    
		  end select
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FilePath(baseLevel as integer = 0) As string
		  
		  
		  
		  if baseLevel = 0 then
		    return string.FromArray(FileParentPath,"") + Filename
		    
		  else
		    var s as string
		    for i as integer = baseLevel to FileParentPath.LastIndex
		      s = s  + FileParentPath(i)
		      
		    next
		    
		    return s + Filename
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function GetAllEntries(ContentType as ContentTypes) As clFileTypeInfo()
		  
		  var r() as clFileTypeInfo
		  
		  for each c as clFileTypeInfo in FileTypeDir
		    if c.ShortContentType = ContentType then  r.add(c)
		    
		  next
		  
		  return r
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function GetFirstEntry(SelectedContentType as ContentTypes) As clFileTypeInfo
		  
		  for each c as clFileTypeInfo in FileTypeDir
		    if c.ShortContentType = SelectedContentType then return c
		    
		  next
		  
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub Initialize()
		  
		  
		  AddFileTypeEntry("xx", ContentTypes.app, "app.xml", "/docProps/")
		  AddFileTypeEntry("xx", ContentTypes.core, "core.xml", "/docProps/")
		  //AddFileTypeEntry("xx", ContentTypes.custom, "custom.xml", "/docProps/")
		  
		  AddFileTypeEntry("xx", ContentTypes.sharedStrings,"sharedStrings.xml", "/xl/")
		  
		  // to replace by dynamic setup
		  AddFileTypeEntry("xx", ContentTypes.style, "styles.xml", "/xl/")
		  AddFileTypeEntry("xx", ContentTypes.theme,"theme.xml", "/xl/","theme/")
		  AddFileTypeEntry("xx", ContentTypes.workbook,"workbook.xml", "/xl/")
		  
		  //to replace by dynamic setup
		  AddFileTypeEntry("xx", ContentTypes.worksheet,"sheet1.xml", "/xl/", "worksheets/")
		  
		  Return
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RelationType() As string
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Filename As string
	#tag EndProperty

	#tag Property, Flags = &h0
		FileParentPath() As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Shared FileTypeDir() As clFileTypeInfo
	#tag EndProperty

	#tag Property, Flags = &h0
		RelationShipID As string
	#tag EndProperty

	#tag Property, Flags = &h0
		ShortContentType As ContentTypes
	#tag EndProperty


	#tag Enum, Name = ContentTypes, Type = Integer, Flags = &h0
		app
		  core
		  sharedStrings
		  style
		  theme
		  workbook
		worksheet
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
			Name="Filename"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="RelationShipID"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ShortContentType"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="ContentTypes"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FileParentPath()"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
