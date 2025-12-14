#tag Class
Protected Class clXLSX_Debugging
	#tag Method, Flags = &h0
		Sub All_Off()
		  
		  self.TraceLoadSharedStrings = False
		  self.TraceLoadStyles = False
		  self.TraceLoadSheetData = False
		  self.TraceLoadWorkbook = False
		  self.TraceLoadNamedRanges = False
		  self.TraceLoadSheetListInWorkbook  = false
		  self.TraceShowUnzippedWorkbool = False
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub All_On(IncludeSheet_Details as Boolean)
		  
		  self.TraceLoadSharedStrings = True
		  self.TraceLoadStyles = True
		  self.TraceLoadSheetData = IncludeSheet_Details
		  self.TraceLoadWorkbook = True
		  self.TraceLoadNamedRanges = True
		  self.TraceLoadSheetListInWorkbook  = True
		  self.TraceShowUnzippedWorkbool = True
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  
		  self.TraceLoadSharedStrings = False
		  self.TraceLoadStyles = False
		  self.TraceLoadSheetData = False
		  self.TraceLoadWorkbook = False
		  
		  self.TraceLoadNamedRanges =True
		  self.TraceLoadSheetListInWorkbook  = True
		  
		  self.TraceShowUnzippedWorkbool = True
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		TraceLoadNamedRanges As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceLoadSharedStrings As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceLoadSheetData As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceLoadSheetListInWorkbook As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceLoadStyles As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceLoadWorkbook As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TraceShowUnzippedWorkbool As Boolean
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
			Name="TraceLoadSharedStrings"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TraceLoadSheetData"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TraceLoadStyles"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TraceLoadWorkbook"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TraceShowUnzippedWorkbool"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TraceLoadSheetListInWorkbook"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
