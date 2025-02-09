#tag Module
Protected Module Module2
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


End Module
#tag EndModule
