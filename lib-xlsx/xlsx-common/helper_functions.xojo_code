#tag Module
Protected Module helper_functions
	#tag Method, Flags = &h0
		Function AddChildNode(xml as XMLDocument, name as string, attribute as string, value as string) As XmlNode
		  var aNode as XMLNode = xml.CreateElement(name)
		  
		  aNode.SetAttribute(attribute, value)
		  
		  return anode
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateBorderNode(xml as XMLDocument, leftStyle as string, leftColor as string, rightStyle as String, rightColor as String, topStyme as string, topColor as string, bottomStyle as string, bottomColor as string) As xmlNode
		  
		  
		  var borderNode as XmlNode = xml.CreateElement("border")
		  
		  var sides() as string = array("left", "right", "top","bottom","diagonal")
		  var sideStyle() as string = array(leftStyle, rightStyle, topStyme, bottomStyle, "")
		  var sideColor() as string = array(leftColor, rightColor, topColor, bottomColor, "")
		  
		  for i as integer = 0 to sides.LastIndex
		    var sideNode as xmlNode = xml.CreateElement(sides(i))
		    
		    if sideStyle(i).Length > 0 then
		      sideNode.SetAttribute("style", sideStyle(i))
		      
		    end if
		    
		    if sideColor(i).Length > 0 then 
		      var colorNode as xmlNode = xml.CreateElement("color")
		      colorNode.SetAttribute("rgb", sideColor(i))
		      
		      sideNode.AppendChild(colorNode)
		      
		    end if
		    
		    borderNode.AppendChild(sideNode)
		    
		  next
		  
		  return borderNode
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateStyleXml()
		  var xml as new XMLDocument
		  
		  var styleNode as XMLNode = xml.CreateElement("styleSheet")
		  
		  styleNode.SetAttribute("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		  styleNode.SetAttribute("xmlns:mc","http://schemas.openxmlformats.org/markup-compatibility/2006")
		  styleNode.SetAttribute("xmlns:x14","http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
		  styleNode.SetAttribute("xmlns:x14ac","http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
		  styleNode.SetAttribute("xmlns:x16r2","http://schemas.microsoft.com/office/spreadsheetml/2015/02/main")
		  styleNode.SetAttribute("mc:Ignorable","x14ac x16r2")
		  
		  xml.AppendChild(styleNode)
		  
		  
		  //
		  // Create fonts
		  //
		  var fontsNode as XMLNode = xml.CreateElement("fonts")
		  
		  var fontNode as XMLNode = xml.CreateElement("font")
		  fontNode.AppendChild(AddChildNode(xml, "sz","val","11.000000"))
		  fontNode.AppendChild(AddChildNode(xml, "color","theme","1"))
		  fontNode.AppendChild(AddChildNode(xml, "name","val","Calibri"))
		  fontNode.AppendChild(AddChildNode(xml, "scheme","val","minor"))
		  
		  fontsNode.AppendChild(fontNode)
		  fontsNode.SetAttribute("count","1")
		  
		  styleNode.AppendChild(fontsNode)
		  
		  
		  //
		  // Create default num formats
		  //
		  var numFmtsNode as XMLNode = xml.CreateElement("numFmts")
		  
		  var numFmtNode as  XmlNode = xml.CreateElement("numFmt")
		  numFmtNode.SetAttribute("numFmtId","176")
		  numFmtNode.SetAttribute("formatCode","yyyy-mm-dd")
		  
		  numFmtsNode.AppendChild(numFmtNode)
		  numFmtsNode.SetAttribute("count","1")
		  
		  styleNode.AppendChild(numFmtsNode)
		  
		  //
		  // Create fill
		  //
		  
		  var fillsNode as XMLNode = xml.CreateElement("fills")
		  
		  var fillNode as XMLNode = xml.CreateElement("fill")
		  fillNode.SetAttribute("patternType","none")
		  
		  fillsNode.AppendChild(fillNode)
		  fillsNode.SetAttribute("count","1")
		  
		  styleNode.AppendChild(fillsNode)
		  
		  
		  //
		  // Create Border style
		  //
		  
		  var bordersNode as XMLNode = xml.CreateElement("borders")
		  var borderCount as integer
		  var borderNode as XMLNode
		  
		  borderNode = CreateBorderNode(xml, "thin", "thin", "thin", "thin", "","","","")
		  
		  bordersNode.AppendChild(borderNode)
		  borderCount = borderCount + 1
		  
		  bordersNode.SetAttribute("count", str(borderCount))
		  styleNode.AppendChild(bordersNode)
		  
		  //
		  // Create cell style using that format
		  //
		  
		  var cellXfsNode as XmlNode = xml.CreateElement("cellXfs")
		  
		  var xfNode as XmlNode
		  var xfCount as integer
		  
		  xfNode = xml.CreateElement("xf")
		  xfNode.SetAttribute("numFmtId","0")
		  xfNode.SetAttribute("fontId","0")
		  xfNode.SetAttribute("fillId","0")
		  xfNode.SetAttribute("borderId","0")
		  xfNode.SetAttribute("applyNumberFormat","1")
		  
		  cellXfsNode.AppendChild(xfNode)
		  xfCount = xfCount + 1
		  
		  
		  xfNode = xml.CreateElement("xf")
		  xfNode.SetAttribute("numFmtId","176")
		  xfNode.SetAttribute("fontId","0")
		  xfNode.SetAttribute("fillId","0")
		  xfNode.SetAttribute("borderId","0")
		  xfNode.SetAttribute("applyNumberFormat","1")
		  
		  cellXfsNode.AppendChild(xfNode)
		  xfCount = xfCount + 1
		  
		  cellXfsNode.SetAttribute("count", str(xfCount))
		  
		  var tmp as string = xml.ToString
		  
		  return
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function findTestDataFolder() As FolderItem
		  var fld as FolderItem
		  
		  var upcount as integer = 15
		  
		  fld = App.ExecutableFile.parent
		  
		  while upcount > 0
		    
		    for each subfld as FolderItem in fld.Children
		      
		      if subfld.Name = "test_xlsx_data" and subfld.IsFolder then return subfld
		      
		    next
		    
		    fld = fld.Parent
		    upcount = upcount - 1
		    
		  wend
		  
		  return nil
		  
		  
		End Function
	#tag EndMethod


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
End Module
#tag EndModule
