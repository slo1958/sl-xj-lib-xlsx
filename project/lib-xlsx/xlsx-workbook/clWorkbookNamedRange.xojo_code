#tag Class
Protected Class clWorkbookNamedRange
	#tag Method, Flags = &h0
		Sub Constructor(pName as string, pRange as string, pLocalSheetID as integer)
		  self.Name = pName
		  self.SourceRangeV = pRange
		  self.localSheetID = pLocalSheetID
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTargetSheetName() As string
		  // 
		  // var idx as integer = Name.IndexOf("!")
		  // 
		  // if idx <= 0 then return ""
		  // 
		  // return name.left(idx-1).trim
		  // 
		  
		  return self.Translate().SheetName
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SourceRange() As string
		  return self.SourceRangeV
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Translate() As clRangeReference
		  
		  if TranslatedRange = nil then
		    self.TranslatedRange = new clRangeReference(self.SourceRange)
		    
		    
		  end if
		  
		  return self.TranslatedRange
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected localSheetID As integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Name As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected SourceRangeV As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected TranslatedRange As clRangeReference
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
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="localSheetID"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="SourceRangeV"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
