#tag Class
Protected Class clWorkbookRelation
	#tag Method, Flags = &h0
		Sub Constructor(RelationId as string, RelationType as string, RelationTarget as string)
		  
		  self.rId = RelationId
		  self.rType = RelationType
		  self.rTarget = RelationTarget
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ID() As string
		  return rID
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Target() As string
		  
		  return rTarget
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected rID As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected rTarget As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected rType As String
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
