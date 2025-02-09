#tag Module
Protected Module Module1
	#tag Method, Flags = &h0
		Sub TestCellTypes()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TesTDateFormat()
		  
		  var FormatToTest() as string
		  
		  FormatToTest.add( "[$-130000]d/m/yyyy")
		  
		  FormatToTest.add( "mm-dd-yy")
		  FormatToTest.add( "d-mmm-yy")
		  FormatToTest.add( "d-mmm")
		  FormatToTest.add( "mmm-yy")
		  FormatToTest.add("[$-404]e/m/d")
		  FormatToTest.add( "m/d/yy")
		  FormatToTest.add( "[$-404]e/m/d")
		  
		  FormatToTest.Add("yyyy""年""m""月""")
		  
		  for Each f as string in FormatToTest
		    var vdate as new clCellFormatter(f)
		    System.DebugLog(vdate.FormatValue(DateTime.now))
		    
		  next
		  
		  
		  
		End Sub
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
