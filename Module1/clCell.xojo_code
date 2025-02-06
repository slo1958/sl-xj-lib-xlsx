#tag Class
Protected Class clCell
	#tag Method, Flags = &h0
		Shared Function CalculateAddress(row as integer, column as integer) As String
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  
		  var location as string 
		  
		  if column < 26 then
		    location = colBase.mid(column+1, 1) + str(row+1)
		    
		  end if
		  
		  return location
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(range as string, prmCelltype as string, cellFormat as string)
		  
		  self.CellLocation = range
		  self.CellSharedStringIndex = -1
		  
		  self.CellType = prmCelltype
		  
		  var p as pair = self.ExtractLocation(range)
		  
		  self.CellRow = p.Left
		  self.CellColumn = p.Right
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function ExtractLocation(CellAddress as string) As pair
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  
		  var tmpcol as integer
		  var tmprow as integer
		  
		  for i as integer = 0 to CellAddress.Length
		    var char as string = CellAddress.Middle(i,1)
		    
		    var a as integer = colBase.IndexOf(char)
		    
		    if "A" <= char and char <= "Z" then tmpcol = tmpcol * 26 + colBase.IndexOf(char)+1
		    if "0" <= char and char <= "9" then tmprow = tmprow*10 + char.ToInteger
		    
		  next
		  
		  return tmprow : tmpcol
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetValueAsString(wb as clWorkbook) As string
		  
		  if self.CellSharedStringIndex < 0 then
		    return CellValue
		    
		  else
		    return wb.GetSharedString(CellSharedStringIndex)
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetFormula(prmFormula as String)
		  
		  self.CellFormula = prmFormula
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetValue(prmValue as String)
		  
		  
		  if self.celltype = "s" then 
		    self.CellSharedStringIndex = val(prmValue)
		    
		  else
		    self.CellValue = prmValue
		    
		  end if
		  
		  
		  return
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CellColumn As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellFormula As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellLocation As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		CellSharedStringIndex As integer
	#tag EndProperty

	#tag Property, Flags = &h0
		CellType As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private CellValue As string
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
			Name="CellValue"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellColumn"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellFormula"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellLocation"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellRow"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellSharedStringIndex"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellType"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
