#tag Class
Protected Class clWorkbookNamedRange
	#tag Method, Flags = &h0
		Sub Constructor(pName as string, pRange as string, pSheetLimit as string)
		  self.Name = pName
		  self.rawRange = pRange
		  self.SheetLimit = pSheetLimit
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTargetSheetName() As string
		  
		  var idx as integer = Name.IndexOf("!")
		  
		  if idx <= 0 then return ""
		  
		  return name.left(idx-1).trim
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Name As string
	#tag EndProperty

	#tag Property, Flags = &h0
		rawRange As string
	#tag EndProperty

	#tag Property, Flags = &h0
		SheetLimit As String
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
			Name="Name"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
