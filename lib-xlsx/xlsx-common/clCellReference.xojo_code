#tag Class
Protected Class clCellReference
	#tag Method, Flags = &h0
		Sub Constructor(cellRangeAsString as string)
		  
		  self.SourceValue = cellRangeAsString
		  self.ScanErrorAt = ParseCoordinate(cellRangeAsString)
		  
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
		Function Location() As pair
		  
		  return self.Row:self.Column
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParseCoordinate(CellAddress as string) As integer
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  const AbsoluteChar as string = "$"
		  
		  const modeStart as uint8 = 0
		  const modeColumn as uint8 = 1
		  const modeRow as uint8 = 2
		  
		  var currentMode as uint8 = modeStart
		  
		  
		  self.Column = -1
		  self.Row = -1
		  self.ColumnAbsolute = false
		  self.RowAbsolute = false
		  
		  var tempAbsoluteMark as Boolean = false
		  
		  for i as integer = 0 to CellAddress.Length
		    var char as string = CellAddress.Middle(i,1)
		    
		    select case currentMode
		      
		    case modeStart
		      if char = AbsoluteChar Then
		        tempAbsoluteMark = true
		        
		      elseif "A" <= char and char <= "Z" then
		        currentMode = modeColumn
		        
		        self.ColumnAbsolute = tempAbsoluteMark
		        self.Column =   colBase.IndexOf(char)+1
		        
		        tempAbsoluteMark = false
		        
		      elseif "0" <= char and char <= "9" then
		        currentMode = modeRow
		        
		        self.row =  char.ToInteger
		        self.RowAbsolute = tempAbsoluteMark
		        
		        tempAbsoluteMark = false
		        
		      else
		        
		      end if
		      
		    case modeColumn
		      if char = AbsoluteChar Then
		        CurrentMode = modeRow
		        self.RowAbsolute = true
		        self.row = 0
		        
		      elseif "A" <= char and char <= "Z" then
		        self.Column =   self.Column *26 + colBase.IndexOf(char)+1
		        
		      elseif "0" <= char and char <= "9" then
		        currentMode = modeRow
		        self.row =  char.ToInteger
		        
		      else
		        
		      end if
		      
		      
		    case modeRow
		      if "0" <= char and char <= "9" then
		        currentMode = modeRow
		        
		        self.row = self.row * 10 + char.ToInteger
		        
		      else
		        
		      end if
		      
		    case else
		      
		    end select
		    
		  next
		  
		  return 0
		  
		   
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As string
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  
		  var tmp as string
		  
		  if self.Column > -1 then
		    
		    
		    var tmpcol as integer = self.Column
		    
		    while tmpcol > 0
		      var rz as integer = tmpcol - (tmpcol \ 26) * 26
		      
		      tmp = colBase.Middle(rz-1,1) + tmp
		      
		      tmpcol = tmpcol  - rz 
		      
		      tmpcol = tmpcol / 26
		      
		    wend
		    
		    if self.ColumnAbsolute then tmp = "$" + tmp 
		    
		  end if
		  
		  if self.Row > -1 then
		    if self.RowAbsolute then tmp = tmp + "$"
		    
		    tmp = tmp + self.row.ToString
		    
		  end if
		  
		  return tmp
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Column As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ColumnAbsolute As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Row As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		RowAbsolute As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ScanErrorAt As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceValue As string
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
			Name="Row"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
