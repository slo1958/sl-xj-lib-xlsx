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
   HasTitleBar     =   True
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
   Title           =   "Demo"
   Type            =   0
   Visible         =   True
   Width           =   600
   Begin DesktopTabPanel TabPanel1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   368
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Panels          =   ""
      Scope           =   0
      SmallTabs       =   False
      TabDefinition   =   "Extract cells\rTabular sheets\rForm"
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Value           =   0
      Visible         =   True
      Width           =   560
      Begin DesktopListBox ListBox2
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
         Height          =   218
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   ""
         Italic          =   False
         Left            =   40
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         RequiresSelection=   False
         RowSelectionType=   0
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   122
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   520
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
      Begin DesktopLabel Label2
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   44
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   40
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   True
         Scope           =   0
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Extract cells from example workbook. Show guessed type, source value and formatted value"
         TextAlignment   =   0
         TextColor       =   &c000000
         Tooltip         =   ""
         Top             =   66
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   520
      End
      Begin ccWorkbook ccWorkbook1
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowFocusRing  =   False
         AllowTabs       =   True
         Backdrop        =   0
         BackgroundColor =   &cFFFFFF
         Composited      =   False
         Enabled         =   True
         HasBackgroundColor=   False
         Height          =   308
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   25
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   2
         TabStop         =   True
         Tooltip         =   ""
         Top             =   70
         Transparent     =   True
         Visible         =   True
         Width           =   549
      End
      Begin ccWorkbook ccWorkbook2
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowFocusRing  =   False
         AllowTabs       =   True
         Backdrop        =   0
         BackgroundColor =   &cFFFFFF
         Composited      =   False
         Enabled         =   True
         HasBackgroundColor=   False
         Height          =   308
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   25
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   3
         TabStop         =   True
         Tooltip         =   ""
         Top             =   70
         Transparent     =   True
         Visible         =   True
         Width           =   549
      End
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Opening()
		  
		  // Extract sample data
		  ExtractCellSamples
		  
		  var baseFld as FolderItem = findTestDataFolder
		  
		  var fld as FolderItem
		  var debugOptions as  clXLSX_Debugging
		  
		  // Load a set of tabular sheets, use on-demand load
		  fld = baseFld.Child("test_file_2.xlsx")
		  if not fld.Exists then Return
		  
		  debugOptions = new clXLSX_Debugging
		  debugOptions.All_Off()
		  
		  ccWorkbook1.UseWorkbook(new clWorkbook(fld , debugOptions, clWorkbook.LoadModes.LoadSheetOnDemand))
		  
		  
		  
		  // Load a form sheet
		  fld = baseFld.Child("test_file_1.xlsx")
		  if not fld.Exists then Return
		  
		  debugOptions = new clXLSX_Debugging
		  debugOptions.All_On(False)
		  
		  ccWorkbook2.UseWorkbook( new clWorkbook(fld, debugOptions ))
		  
		  
		  return
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub ExampleLoadAuto(filename as string)
		  //
		  // Example load workbook, 
		  // - auto mode 
		  // - use temporary folder as workarea
		  //
		  
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  var myworkbook as clWorkbook =  new clWorkbook(fld)
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ExampleLoadOnDemand(filename as string)
		  //
		  // Example load workbook, 
		  // - SheetOnDemand mode 
		  // - use desktop folder as workarea
		  //
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  var myworkbook as clWorkbook = new clWorkbook(fld _
		  , clWorkbook.LoadModes.LoadSheetOnDemand _
		  , "" _
		  , SpecialFolder.Desktop.child(filename.ReplaceAll(".","-") + " folder") _
		  )
		  
		  
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ExtractCellSamples()
		  
		  var filename as string = "test_file_2.xlsx"
		  
		  var fld as FolderItem = findTestDataFolder
		  
		  fld = fld.Child(filename)
		  
		  if not fld.Exists then Return
		  
		  var Workbook as new clWorkbook(fld )
		  
		  var worksheet as clWorksheet = Workbook.GetSheetFromName("Sales Data 1")
		  
		  if worksheet = nil then return
		  
		  var collectedCells() as clCell
		  
		  for i as integer = 3 to 8
		    var cellLocation as string
		    
		    cellLocation = "A" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "B" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "C" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "D" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "E" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "F" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		    cellLocation = "G" + str(i)
		    collectedCells.Add(worksheet.GetCell(cellLocation))
		    
		  next
		  
		  
		  listbox2.RemoveAllRows
		  listbox2.ColumnCount = 5
		  
		  Listbox2.HeaderAt(0)="Location"
		  Listbox2.HeaderAt(1)="Loaded value"
		  Listbox2.HeaderAt(2)="Guessed type"
		  Listbox2.HeaderAt(3)="Value"
		  Listbox2.HeaderAt(4)="Formatted value"
		  
		  for each cell as clCell in collectedCells
		    if cell <> nil then
		      var cellLocation as string = cell.CellLocation
		      var cellDisplayStr as string
		      var cellType as string 
		      var cellLoadedStr as string =cell.CellSourceValue
		      
		      var cellFormattedStr as string = cell.GetValueAsString(workbook)
		      
		      select case cell.GuessType(workbook)
		        
		      case clCell.GuessedType.General
		        cellType = "General"
		        cellDisplayStr = "" 
		        
		      case clCell.GuessedType.Date
		        cellType = "Date"
		        cellDisplayStr = cell.GetValueAsDateTime(workbook).SQLDate 
		        
		      case clCell.GuessedType.Number
		        cellType = "Number"
		        cellDisplayStr =  format(cell.GetValueAsNumber(Workbook), "-#####0.000") 
		        
		      case  clCell.GuessedType.String
		        cellType = "String"
		        cellDisplayStr = ""
		        
		      end select
		      
		      
		      Listbox2.AddRow(cellLocation)
		      Listbox2.CellTextAt(Listbox2.LastAddedRowIndex,1 ) = cellLoadedStr
		      Listbox2.CellTextAt(Listbox2.LastAddedRowIndex, 2) = cellType
		      Listbox2.CellTextAt(Listbox2.LastAddedRowIndex, 3) = cellDisplayStr
		      Listbox2.CellTextAt(Listbox2.LastAddedRowIndex, 4) = cellFormattedStr
		      
		    end if
		    
		  next
		  
		  var k as integer = 1
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod


#tag EndWindowCode

#tag ViewBehavior
	#tag ViewProperty
		Name="HasTitleBar"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
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
