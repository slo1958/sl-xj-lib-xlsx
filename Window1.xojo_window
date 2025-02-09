#tag DesktopWindow
Begin DesktopWindow Window1
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   Height          =   446
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   2092408831
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "Untitled"
   Type            =   0
   Visible         =   True
   Width           =   600
   Begin DesktopListBox ListBox1
      AllowAutoDeactivate=   True
      AllowAutoHideScrollbars=   True
      AllowExpandableRows=   False
      AllowFocusRing  =   True
      AllowResizableColumns=   False
      AllowRowDragging=   False
      AllowRowReordering=   False
      Bold            =   False
      ColumnCount     =   1
      ColumnWidths    =   ""
      DefaultRowHeight=   -1
      DropIndicatorVisible=   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      GridLineStyle   =   0
      HasBorder       =   True
      HasHeader       =   True
      HasHorizontalScrollbar=   False
      HasVerticalScrollbar=   True
      HeadingIndex    =   -1
      Height          =   279
      Index           =   -2147483648
      InitialValue    =   ""
      Italic          =   False
      Left            =   35
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      RowSelectionType=   0
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   109
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   545
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin DesktopPopupMenu PopupMenu1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialValue    =   ""
      Italic          =   False
      Left            =   35
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      SelectedRowIndex=   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   67
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   183
   End
   Begin DesktopLabel Label1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      Italic          =   False
      Left            =   35
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Untitled"
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   400
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   545
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  
		  var no as integer = 3
		  
		  select case no 
		    
		  case 0
		    ProcessXLSXFileAuto( "test_file_1.xlsx")
		    
		  case 1
		    ProcessXLSXFileOnDemand( "test_file_1.xlsx")
		    
		  case 2
		    ProcessXLSXFileAuto( "test_file_2.xlsx")
		    
		  case 3
		    ProcessXLSXFileOnDemand( "test_file_2.xlsx")
		    
		  end select
		  
		  return
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub ProcessXLSXFileAuto(filename as string)
		  //
		  // Example load workbook, 
		  // - auto mode 
		  // - use temporary folder as workarea
		  //
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  self.loadedWorkbook = new clWorkbook(fld )
		  
		  self.UpdateUI
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessXLSXFileOnDemand(filename as string)
		  //
		  // Example load workbook, 
		  // - SheetOnDemand mode 
		  // - use desktop folder as workarea
		  //
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  self.loadedWorkbook = new clWorkbook(fld _
		  , clWorkbook.LoadModes.LoadSheetOnDemand _
		  , "" _
		   , SpecialFolder.Desktop.child(filename.ReplaceAll(".","-") + " folder") _
		  )
		  
		  self.UpdateUI
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ShowSheet(sheetname as string)
		  
		  var sheet as clWorksheet
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateUI()
		  
		  
		  popupmenu1.RemoveAllRows
		  
		  var sheets() as string = self.loadedWorkbook.GetSheetNames
		  
		  popupmenu1.AddAllRows(sheets)
		  
		  popupmenu1.SelectedRowIndex = 0
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		loadedWorkbook As clWorkbook
	#tag EndProperty


#tag EndWindowCode

#tag Events ListBox1
	#tag Event
		Sub Opening()
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function CellPressed(row As Integer, column As Integer, x As Integer, y As Integer) As Boolean
		  
		  var tmp as string = me.CellTextAt(row, column)
		  
		  Label1.Text = tmp
		  
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events PopupMenu1
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  Const colBase as string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		  var name as string = me.SelectedRowText
		  
		  var sheet as clWorksheet = self.loadedWorkbook.GetSheet(name)
		  
		  ListBox1.RemoveAllRows
		  
		  if sheet = nil then return
		  
		  listbox1.ColumnCount = sheet.lastColumn + 1
		  
		  listbox1.HeaderAt(0) = "#"
		  
		  for i as integer= 1 to sheet.lastColumn + 1
		    listbox1.HeaderAt(i) = colBase.Middle(i-1,1)
		    
		  next
		  
		  
		  for each row as clWorkrow in sheet.rows
		    
		    if row <> nil then
		      Listbox1.AddRow str(row.row)
		      
		      for col as integer = 1 to sheet.lastColumn
		        var rc as clCell = row.GetCell(col)
		        var tmp as string 
		        if rc <> nil then tmp = rc.GetValueAsString(self.loadedWorkbook)
		        
		        Listbox1.CellTextAt(listbox1.LastAddedRowIndex, col) = tmp
		        
		      next
		      
		    end if
		    
		  next
		End Sub
	#tag EndEvent
#tag EndEvents
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
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
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
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
