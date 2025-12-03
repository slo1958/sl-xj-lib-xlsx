#tag Class
Protected Class clRangeReference
	#tag Method, Flags = &h0
		Sub Constructor(rangeAsString as string)
		  
		  self.SourceValue = RangeAsString 
		  
		  self.ScanErrorAt = ParseCoordinate(RangeAsString)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Location() As pair
		  
		  if self.TopLeft = nil or self.BottomRight = nil then
		    return (0:0):(0:0)
		    
		  end if
		  
		  return self.TopLeft.Location : self.BottomRight.Location
		  
		   
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParseCoordinate(RangeAddress as string) As integer
		  
		  var tempRange as string = RangeAddress
		  
		  var idxSheetMark as integer = tempRange.IndexOf("!")
		  
		  self.SheetName = if(idxSheetMark<=0 ,"", tempRange.left(idxSheetMark).trim)
		  
		  tempRange =  if(idxSheetMark<=0 , tempRange, tempRange.Middle(idxSheetMark+1, 9999))
		  
		  var idxRowMark as integer = tempRange.IndexOf(":")
		  
		  if idxRowMark <= 0 then
		    return -1
		    
		  else
		    var TLPart as string = tempRange.Middle(0, idxRowMark)
		    var BRPart as string = tempRange.Middle(idxRowMark+1, 9999)
		    
		    self.TopLeft = new clCellReference(TLPart)
		    self.BottomRight = new clCellReference(BRPart)
		    
		    System.DebugLog(RangeAddress + " => Sheet: " + self.SheetName + "  range:" + tempRange + " split " + TLPart+" and "  + BRPart)
		    
		    return 0
		  end if
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As string
		  
		  return if(  self.SheetName.Length > 0, self.SheetName+"!","") + if(TopLeft = nil, "", TopLeft.ToString)  + ":" + if (BottomRight=nil, "", BottomRight.ToString)
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		BottomRight As clCellReference
	#tag EndProperty

	#tag Property, Flags = &h0
		ScanErrorAt As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SheetName As string
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceValue As string
	#tag EndProperty

	#tag Property, Flags = &h0
		TopLeft As clCellReference
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
