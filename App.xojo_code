#tag Class
Protected Class App
Inherits DesktopApplication
	#tag Event
		Sub Opening()
		  // TestZip
		  
		  // IterateFolder (SpecialFolder.Desktop.Child("test_file_2-xlsx folder"))
		  //TesTDateFormat
		  //CreateStyleXml()
		  
		  
		  //ProcessXLSXFile("Archive5.xlsx")
		  
		  //
		  
		  ProcessXLSXFile( "test_file_1.xlsx")
		  //ProcessXLSXFile( "test_file_2.xlsx")
		  
		  return
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub IterateFolder(fld as FolderItem, level as integer = 0)
		  
		  var prefix as string
		  
		  for i as integer = 0 to level
		    prefix = prefix + "-----"
		    
		  next
		  
		  for each subfld as FolderItem in fld.Children
		    System.DebugLog(prefix + subfld.name)
		    
		    if subfld .name = ".DS_Store" then 
		      subfld.Remove
		    else
		      
		      if subfld.IsFolder then IterateFolder(subfld, level+1)
		    end if
		    
		  next
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessXLSXFile(filename as string)
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  self.loadedWorkbook = new Module1.clWorkbook(fld)
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RemoveDSStore(fld as FolderItem)
		  
		  
		  
		  for each subfld as FolderItem in fld.Children
		    //System.DebugLog(prefix + subfld.name)
		    
		    if subfld .name = ".DS_Store" then 
		      subfld.Remove
		    else
		      
		      if subfld.IsFolder then IterateFolder(subfld)
		    end if
		    
		  next
		  
		  return
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

	#tag Method, Flags = &h0
		Sub TestZip()
		  
		  var basefolder as FolderItem = SpecialFolder.Desktop.child("XLSX_Test_Folder")
		  
		  
		  // Unzip xlsx file
		  
		  var file1 as FolderItem = basefolder.Child("test_file_2.xlsx")
		  
		  var tempfolder1 as FolderItem = basefolder.child("test_folder_3")
		  
		  if TempFolder1.Exists then TempFolder1.RemoveFolderAndContents
		  
		  TempFolder1.CreateFolder
		  
		  file1.Unzip(TempFolder1 )
		  
		  
		  // Zip back
		  
		  var srcfld as FolderItem = basefolder.Child("test_folder_3")
		  var dst as FolderItem = basefolder.child("test_file_4.xlsx")
		  
		  if dst.Exists then dst.Remove
		  
		  RemoveDSStore(srcfld)
		  
		  call srcfld.zip(dst, True)
		  
		  // Unzip again xlsx file
		  
		  var file2 as FolderItem = basefolder.Child("test_file_4.xlsx")
		  
		  var tempfolder2 as FolderItem = basefolder.child("test_folder_5")
		  
		  if TempFolder2.Exists then TempFolder2.RemoveFolderAndContents
		  
		  TempFolder2.CreateFolder
		  
		  file2.Unzip(TempFolder2)
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		loadedWorkbook As Module1.clWorkbook
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoQuit"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowHiDPI"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="BugVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Copyright"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Description"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LastWindowIndex"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MajorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MinorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NonReleaseVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RegionCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="StageCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Version"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="_CurrentEventTime"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
