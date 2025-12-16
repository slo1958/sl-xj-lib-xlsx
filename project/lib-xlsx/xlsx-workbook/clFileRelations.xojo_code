#tag Class
Protected Class clFileRelations
	#tag Method, Flags = &h0
		Sub AddFileInfoEntry(RelationshipId as string, ContentType as clFileRelationEntry.ContentTypes, Filename as string, pTypePartFolder as string, pRelationFolder as string)
		  var c as new clFileRelationEntry
		  
		  c.Filename = Filename
		  c.RelationFolder = pRelationFolder
		  c.TypePartFolder = pTypePartFolder
		  c.ShortContentType = ContentType
		  c.RelationShipID = RelationshipId
		  
		  FileTypeDir.Add(c)
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddFilePartTypeEntry(RelationshipId as string, ContentType as clFileRelationEntry.ContentTypes, Filename as string, FileParentPath() as string)
		  var c as new clFileRelationEntry
		  
		  // c.Filename = Filename
		  // for each s as string in FileParentPath
		  // c.FileParentPath.Add(s)
		  // 
		  // next
		  // 
		  // c.ShortContentType = ContentType
		  // c.RelationShipID = RelationshipId
		  // 
		  // FileTypeDir.Add(c)
		  
		  System.DebugLog(CurrentMethodName + " called.")
		  
		  return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddFileRelationEntry(relId as string, reltype as string, relTarget as string)
		  
		  var c as new clFileRelationEntry
		  var tmp() as String
		  
		  
		  tmp = relTarget.Split("/")
		  
		  c.RelationShipID = relId
		  c.ShortContentType = clFileRelationEntry.RelationshipTypeToShortContentType(reltype)
		  c.Filename = tmp(tmp.LastIndex)
		  
		  if tmp.Count > 1 then
		    c.RelationFolder = tmp(0)
		    
		  end if
		  
		  c.TypePartFolder = "xl"
		  
		  FileTypeDir.Add(c)
		  
		  Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetAllEntries(ContentType as clFileRelationEntry.ContentTypes) As clFileRelationEntry()
		  
		  var r() as clFileRelationEntry
		  
		  for each c as clFileRelationEntry in FileTypeDir
		    if c.ShortContentType = ContentType then  r.add(c)
		    
		  next
		  
		  return r
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetEntryById(RelationId as string) As clFileRelationEntry
		  
		  for each c as clFileRelationEntry in FileTypeDir
		    if c.RelationShipID = RelationId then return c
		    
		  next
		  
		  return nil
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetFirstEntry(SelectedContentType as clFileRelationEntry.ContentTypes) As clFileRelationEntry
		  
		  for each c as clFileRelationEntry in FileTypeDir
		    if c.ShortContentType = SelectedContentType then return c
		    
		  next
		  
		  return nil
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Initialize()
		  
		  FileTypeDir.RemoveAll
		  
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.app, "app.xml", "docProps", "")
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.core, "core.xml", "docProps", "")
		  //AddFileTypeEntry("xx", ContentTypes.custom, "custom.xml", "/docProps/")
		  
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.sharedStrings,"sharedStrings.xml", "xl", "")
		  
		  // to replace by dynamic setup
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.style, "styles.xml", "xl", "")
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.theme,"theme.xml", "xl","theme")
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.workbook,"workbook.xml", "xl", "")
		  
		  //to replace by dynamic setup
		  AddFileInfoEntry("xx", clFileRelationEntry.ContentTypes.worksheet,"sheet1.xml", "xl", "worksheets")
		  
		  Return
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		FileTypeDir() As clFileRelationEntry
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
