#tag Class
Protected Class clFileRelationEntry
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
		Shared Function ContentTypeToShortContentType(s as string) As ContentTypes
		  select case s
		    
		  case "application/vnd.openxmlformats-officedocument.extended-properties+xml"
		    Return ContentTypes.app
		    
		  case  "application/vnd.openxmlformats-package.core-properties+xml"
		    Return ContentTypes.core
		    
		  case  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
		    Return ContentTypes.sharedStrings
		    
		  case  "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
		    Return ContentTypes.style
		    
		  case "application/vnd.openxmlformats-officedocument.theme+xml"
		    Return ContentTypes.theme
		    
		  case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
		    Return ContentTypes.workbook
		    
		  case "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
		    Return ContentTypes.worksheet
		    
		  case else
		    Return ContentTypes.undefined
		    
		  end select
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RelationPath() As string
		  var tmp() as string
		  
		  if self. RelationFolder <> "" then tmp.Add(self.RelationFolder)
		  tmp.add(Filename)
		  
		  return string.FromArray(tmp,"/")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RelationshipType() As string
		  // <Relationship Id="rId9" Type= Target="styles.xml"></Relationship>
		   
		  // <Relationship Id="rId7" Type= Target="theme/theme1.xml"></Relationship>
		  // <Relationship Id="rId6" Type= Target="worksheets/sheet6.xml"></Relationship>
		  
		  select case ShortContentType
		    
		  case ContentTypes.app
		    return "?app" 
		    
		  case ContentTypes.core
		    Return "?core"
		    
		  case ContentTypes.sharedStrings
		    Return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
		    
		  case ContentTypes.style
		    Return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
		    
		  case ContentTypes.theme
		    Return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
		    
		  case ContentTypes.workbook
		    Return "?wbook"
		    
		  case ContentTypes.worksheet
		    Return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
		    
		  case else
		    
		  end select
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function RelationshipTypeToShortContentType(s as string) As ContentTypes
		  
		  select case s
		    
		  case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
		    Return ContentTypes.sharedStrings
		    
		  case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
		    Return ContentTypes.style
		    
		  case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
		    Return ContentTypes.theme
		    
		  case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
		    Return ContentTypes.worksheet
		    
		  case else
		    System.DebugLog ("Unknow relationship type " + s)
		    return ContentTypes.undefined
		    
		  end Select
		   
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Target() As string
		  
		  Return self.RelationPath
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TypePartPath() As string
		  var tmp() as string
		  
		  if self.TypePartFolder <> "" then tmp.Add(self.TypePartFolder)
		  if self. RelationFolder <> "" then tmp.Add(self.RelationFolder)
		  tmp.add(Filename)
		  
		  return "/" + string.FromArray(tmp,"/")
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Filename As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RelationFolder As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RelationShipID As string
	#tag EndProperty

	#tag Property, Flags = &h0
		ShortContentType As ContentTypes
	#tag EndProperty

	#tag Property, Flags = &h0
		TypePartFolder As String
	#tag EndProperty


	#tag Enum, Name = ContentTypes, Flags = &h0
		app
		  core
		  sharedStrings
		  style
		  theme
		  workbook
		  worksheet
		undefined
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
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ShortContentType"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="ContentTypes"
			EditorType="Enum"
			#tag EnumValues
				"0 - app"
				"1 - core"
				"2 - sharedStrings"
				"3 - style"
				"4 - theme"
				"5 - workbook"
				"6 - worksheet"
				"7 - undefined"
			#tag EndEnumValues
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
