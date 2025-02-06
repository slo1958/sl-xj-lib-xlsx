#tag Class
Protected Class clCellXf
	#tag Method, Flags = &h0
		Sub Constructor(StyleMode as boolean, xfNode as xmlNode)
		  
		  self.IsStyle = StyleMode
		  
		  self.FontId = xfNode.GetAttribute("fondId") .ToInteger
		  
		  self.FillId = xfNode.GetAttribute("fillId") .ToInteger
		  
		  self.BorderId = xfNode.GetAttribute("borderId").ToInteger
		  
		  self.NumberFormatId = xfNode.GetAttribute("numFmtId").ToInteger
		  
		  self.ApplyFont = xfNode.GetAttribute("applyFont").ToInteger = 1
		  
		  self.ApplyFill = xfNode.GetAttribute("applyFill").ToInteger = 1
		  
		  self.ApplyBorder = xfNode.GetAttribute("applyBorder").ToInteger = 1
		  
		  self.ApplyNumberFormat = xfNode.GetAttribute("applyNumberFormat").ToInteger = 1
		  
		  return
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		ApplyBorder As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ApplyFill As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ApplyFont As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ApplyNumberFormat As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		BorderId As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FillId As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FontId As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		IsStyle As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberFormatId As Integer
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
			Name="IsStyle"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
