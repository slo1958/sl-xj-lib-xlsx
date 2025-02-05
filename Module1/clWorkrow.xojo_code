#tag Class
Protected Class clWorkrow
	#tag Method, Flags = &h0
		Sub AddCell(column as integer, cell as clCell)
		  
		  while cells.LastIndex <= column
		    Cells.add nil
		    
		  wend
		  
		  cells(column) = cell
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(prmRow as integer)
		  
		  self.row = prmRow
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCell(column as integer) As clCell
		  
		  return cells(column)
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		cells() As clCell
	#tag EndProperty

	#tag Property, Flags = &h0
		row As Integer
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
			Name="cells()"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
