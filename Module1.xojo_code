#tag Module
Protected Module Module1
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

	#tag Method, Flags = &h0
		Function PrepareXLSX(file as FolderItem) As integer
		  
		  
		  var wb as new clWorkbook(file)
		  
		  return 0
		  
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
