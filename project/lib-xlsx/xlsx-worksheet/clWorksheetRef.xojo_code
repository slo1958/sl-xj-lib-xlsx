#tag Class
Protected Class clWorksheetRef
	#tag Method, Flags = &h0
		Sub Constructor(SourceFolder as folderitem, SheetName as string, SheetId as integer, SheetRelationId as string, SheetRelationTarget as string, TraceFlag as Boolean)
		  
		  self.TempFolder = SourceFolder
		  self.name = SheetName
		  self.id = SheetId
		  self.RelationId = SheetRelationId 
		  self.RelationTarget = SheetRelationTarget // Could be removed and calculated in LoadSheetData()
		  self.SheetData = nil
		  self.Trace = TraceFlag
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DropData()
		  
		  SheetData = nil
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSheetData(WorkbookInformation as clWorkbookInformation) As clWorksheet
		  
		  if not IsLoaded then
		    self.LoadSheetData(WorkbookInformation)
		    
		  end if
		  
		  Return SheetData
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsLoaded() As Boolean
		  
		  return SheetData <> nil
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadSheetData(WorkbookInformation as clWorkbookInformation)
		  
		  
		  self.SheetData = new clWorksheet(TempFolder, Name, RelationTarget, WorkbookInformation, self.Trace)
		  
		  return 
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Id As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RelationId As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RelationTarget As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected SheetData As clWorksheet
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected TempFolder As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Trace As Boolean
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
			Name="Name"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Id"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RelationId"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RelationTarget"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Trace"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
