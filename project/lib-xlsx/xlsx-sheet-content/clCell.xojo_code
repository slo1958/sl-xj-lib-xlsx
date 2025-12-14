#tag Class
Protected Class clCell
	#tag Method, Flags = &h0
		Function CellColumn() As integer
		  
		  if Reference = nil then
		    return -1
		    
		  else
		    Return Reference.Column
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CellLocationAsPair() As pair
		  
		  return Reference.Location
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CellRow() As integer
		  
		  if Reference = nil then
		    return -1
		    
		  else
		    Return Reference.Row
		    
		  end if
		  
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
		  
		  self.Reference = new clCellReference(self.CellLocation)
		  
		  // var p as pair = self.ExtractLocation(self.CellLocation)
		  // self.CellRow = p.Left
		  // self.CellColumn = p.Right
		  
		  return
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetValueAsDateTime(wb as clWorkbook) As DateTime
		  
		  //
		  // Return nil if the cell has no value
		  //
		  if self.CellSourceValue.trim.Length = 0 then return nil
		  
		  //
		  // Return nil if string or shared string 
		  //
		  select case self.CellType
		    
		  case  TypeSharedString  
		    return nil
		    
		  case TypeFormulaString, TypeInlineString
		    return nil
		    
		  end select
		  
		  //
		  // Could be date or number
		  //
		  if self.CellStyle > 0 then
		    var style as clCellXf = wb.GetCellStyle(self.CellStyle)
		    
		    var fmt as clCellFormatter = wb.GetFormat(style.NumberFormatId)
		    
		    if fmt <> nil and  fmt.IsDateFormat then
		      return fmt.MakeDate(self.CellValue)
		      
		    end if
		    
		  end if
		  
		  
		  return  nil
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetValueAsNumber(wb as clWorkbook) As double
		  
		  //
		  // Return empty string is the loaded cell value is empty
		  //
		  if self.CellSourceValue.trim.Length = 0 then return 0
		  //
		  // String or shared string ?
		  //
		  select case self.CellType
		    
		  case  TypeSharedString  
		    return 0
		    
		  case TypeFormulaString, TypeInlineString
		    return  0
		    
		  end select
		  
		  //
		  // Could be date or number
		  //
		  return self.CellValue
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetValueAsString(wb as clWorkbook) As String
		  
		  const CDefaultNumberFormat = "-#####0.00##"
		  
		  //
		  // Return empty string is the loaded cell value is empty
		  //
		  if self.CellSourceValue.trim.Length = 0 then return ""
		  
		  //
		  // String or shared string ?
		  //
		  select case self.CellType
		    
		  case  TypeSharedString  
		    return wb.GetSharedString(CellSharedStringIndex)
		    
		  case TypeFormulaString, TypeInlineString
		    return self.CellSourceValue
		    
		  end select
		  
		  //
		  // Could be date or number
		  //
		  if self.CellStyle > 0 then
		    var style as clCellXf = wb.GetCellStyle(self.CellStyle)
		    
		    
		    // retain for debugging
		    
		    if self.CellLocation = "E3" then
		      
		      var d  as integer 
		      d = 1
		      var format as clCellFormatter = wb.GetFormat(style.NumberFormatId)
		      d=2
		    end if
		    
		    
		    var fmt as clCellFormatter = wb.GetFormat(style.NumberFormatId)
		    
		    if fmt <> nil then
		      return fmt.FormatValue(self.CellValue)
		      
		      
		    end if
		  end if
		  
		  
		  return format(self.CellValue, CDefaultNumberFormat)
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GuessType(wb as clWorkbook) As GuessedType
		  
		  //
		  // Return general type if the loaded cell value is empty
		  //
		  if self.CellSourceValue.trim.Length = 0 then return GuessedType.General
		  
		  //
		  // String or shared string ?
		  //
		  select case self.CellType
		    
		  case  TypeSharedString  
		    return GuessedType.String
		    
		  case TypeFormulaString, TypeInlineString
		    return GuessedType.String
		    
		  end select
		  
		  //
		  // Could be date or number
		  //
		  if self.CellStyle > 0 then
		    var style as clCellXf = wb.GetCellStyle(self.CellStyle)
		    
		    var fmt as clCellFormatter = wb.GetFormat(style.NumberFormatId)
		    
		    if fmt = nil then return GuessedType.General
		    
		    
		    if fmt.IsDateFormat then
		      return GuessedType.Date
		      
		    else
		      return GuessedType.Number
		      
		      
		      
		    end if
		  end if
		  
		  return GuessedType.General
		  
		  
		  
		  
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
		  
		  self.CellSourceValue = prmValue
		  
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
		CellFormula As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellLocation As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CellSharedStringIndex As integer
	#tag EndProperty

	#tag Property, Flags = &h0
		CellSourceValue As string
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

	#tag Property, Flags = &h0
		Reference As clCellReference
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


	#tag Enum, Name = GuessedType, Type = Integer, Flags = &h0
		General
		  String
		  Date
		Number
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
			Name="CellFormula"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellLocation"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
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
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellSourceValue"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CellStyle"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
