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
		Sub Constructor(baseNode as XMLNode)
		  
		  // Handle cell attributes
		  //
		  //
		  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
		  //
		  // An A1 style reference to the location of this cell
		  self.CellLocation =  basenode.GetAttribute("r")
		  
		  
		  // The index of this cell's style. Style records are stored in the Styles Part.
		  // The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
		  self.CellStyle = basenode.GetAttribute("s").ToInteger
		  
		  
		  // An enumeration representing the cell's data type.
		  // The possible values for this attribute are defined by the ST_CellType simple type
		  self.CellType =  basenode.GetAttribute("t")
		  
		  
		  // Cell meta data
		  // The zero-based index of the cell metadata record associated with this cell. 
		  // Metadata information is found in the Metadata Part. 
		  // Cell metadata is extra information stored at the cell level, and is attached to the cell (travels through moves, copy / paste, clear, etc). 
		  // Cell metadata is not accessible via formula reference.
		  //
		  //The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
		  // self.xxx = basenode.GetAttribute("cm")
		  
		  
		  // Value meta data
		  // The zero-based index of the value metadata record associated with this cell's value. 
		  // Metadata records are stored in the Metadata Part. Value metadata is extra information stored at the cell level, but associated with the value rather than the cell itself. 
		  // Value metadata is accessible via formula reference.
		  //
		  // The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype.
		  // self.xxx = basenode.GetAttribute("vm")
		  
		  
		  // Internals
		  self.CellSharedStringIndex = -1
		  
		  var p as pair = self.ExtractLocation(self.CellLocation)
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
		Function GetValueAsString(wb as clWorkbook) As String
		  
		  const CDefaultNumberFormat = "-#####0.00##"
		  
		  //
		  // String or shared string ?
		  //
		  select case self.CellType
		    
		  case  TypeSharedString  
		    return wb.GetSharedString(CellSharedStringIndex)
		    
		  case TypeFormulaString, TypeInlineString
		    return self.CellStringValue
		    
		  end select
		  
		  
		  //
		  // Could be date or number
		  //
		  if self.CellStyle > 0 then
		    var style as clCellXf = wb.GetCellStyle(self.CellStyle)
		    
		    if self.CellStyle = 2 then
		      var d  as integer = 1
		      var format as string = wb.GetFormat(style.NumberFormatId)
		      d=2
		    end if
		    
		    select case style.NumberFormatId
		    case 14, 15, 16, 17 , 58 // TODO REMOVE 58
		      var dateOffset as integer = self.CellValue
		      var d as new DateTime(new Date(1900,1,1))
		      
		      d = d.AddInterval(0,0, dateOffset-2)
		      
		      return d.SQLDate
		      
		    case else
		      var format as string = wb.GetFormat(style.NumberFormatId)
		      
		      if format = "General" then
		        return format(self.CellValue, CDefaultNumberFormat)
		        
		      else
		        return format(self.CellValue, format)
		        
		      end if
		      
		    end select
		    
		  end if
		  
		  
		  return format(self.CellValue, CDefaultNumberFormat)
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetFormula(prmFormula as String)
		  
		  self.CellFormula = prmFormula
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetValueFromString(prmValue as String)
		  
		  self.CellStringValue = prmValue
		  
		  select case self.CellType
		    
		  case  TypeSharedString  
		    self.CellSharedStringIndex = val(prmValue)
		    
		  case TypeFormulaString, TypeInlineString
		    self.CellValue = prmValue
		    
		  case TypeNumber
		    self.CellValue = self.ValueParser(prmValue)
		    
		  else
		    self.CellValue = self.ValueParser(prmValue)
		    
		  end select
		  
		  
		  return
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function ValueParser(s as String) As double
		  
		  if s.IndexOf("E") < 0 then return val(s)
		  
		  var tmps as string = s.left(s.IndexOf("E"))
		  var tmpx as string = s.right(s.Length - tmps.Length-1)
		  var tmpi as integer = val(tmpx)
		  
		  var tmpd as Double = val(tmps)
		  
		  if tmpi = 0 then return tmpd
		  
		  var tmpm as double = if(tmpi>0, 10.0, 0.1)
		  
		  tmpi = abs (tmpi)
		  
		  while tmpi > 0
		    tmpd = tmpd * tmpm
		    tmpi = tmpi - 1
		    
		  wend
		  
		  return tmpd
		  
		End Function
	#tag EndMethod


	#tag Note, Name = About cell type
		
		Source
		
		https://schemas.liquid-technologies.com/officeopenxml/2006/?page=st_celltype.html
		
	#tag EndNote


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
		CellStringValue As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellStyle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		CellType As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private CellValue As variant
	#tag EndProperty


	#tag Constant, Name = TypeBoolean, Type = String, Dynamic = False, Default = \"b", Scope = Public
	#tag EndConstant

	#tag Constant, Name = TypeError, Type = String, Dynamic = False, Default = \"e", Scope = Public
	#tag EndConstant

	#tag Constant, Name = TypeFormulaString, Type = String, Dynamic = False, Default = \"str", Scope = Public
	#tag EndConstant

	#tag Constant, Name = TypeInlineString, Type = String, Dynamic = False, Default = \"inlineStr", Scope = Public
	#tag EndConstant

	#tag Constant, Name = TypeNumber, Type = String, Dynamic = False, Default = \"n", Scope = Public
	#tag EndConstant

	#tag Constant, Name = TypeSharedString, Type = String, Dynamic = False, Default = \"s", Scope = Public
	#tag EndConstant


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
