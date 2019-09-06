#tag Window
Begin Window Window1
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   668
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   668
   MaximizeButton  =   False
   MaxWidth        =   1134
   MenuBar         =   2013304831
   MenuBarVisible  =   True
   MinHeight       =   668
   MinimizeButton  =   True
   MinWidth        =   1134
   Placement       =   0
   Resizeable      =   False
   Title           =   "Looper Harness"
   Visible         =   True
   Width           =   1134
   Begin PushButton btnCreate
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Create"
      Default         =   False
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   66
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   145
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin PushButton btnMove
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Move"
      Default         =   False
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   476
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   145
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin Label Labels
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   113
      HelpTag         =   ""
      Index           =   0
      InitialParent   =   ""
      Italic          =   False
      Left            =   66
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   True
      Scope           =   2
      Selectable      =   False
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Press Create to add the number of specified rows to the Listbox on the left. Then press Move to move rows randomly to the right. Press Stop anytime to stop either process.\n\nUse the text area on the right to see that the UI is still responsive, but less so for Move since the ActiveTicks is set to 30. This prioritizes processing over the UI."
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   1029
   End
   Begin Label Labels
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   1
      InitialParent   =   ""
      Italic          =   False
      Left            =   291
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "rows"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   145
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   100
   End
   Begin TextField fldRowCount
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   178
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   2
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   144
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   80
   End
   Begin Looper_MTC CreateLooper
      ActiveTicks     =   3
      Index           =   -2147483648
      IsRunning       =   False
      LockedInPosition=   False
      PauseMilliseconds=   5
      Scope           =   2
      TabPanelIndex   =   0
   End
   Begin Listbox lbOriginal
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   2
      ColumnsResizable=   False
      ColumnWidths    =   "100"
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   False
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   2
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   428
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   66
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   2
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      ShowDropIndicator=   False
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   178
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   300
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin Listbox lbMoved
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   2
      ColumnsResizable=   False
      ColumnWidths    =   "100"
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   False
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   2
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   428
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   476
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   2
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      ShowDropIndicator=   False
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   178
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   300
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin ProgressBar ProgressBar1
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   66
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Maximum         =   0
      Scope           =   2
      TabIndex        =   7
      TabPanelIndex   =   0
      Top             =   618
      Transparent     =   False
      Value           =   0
      Visible         =   False
      Width           =   710
   End
   Begin Looper_MTC MoveLooper
      ActiveTicks     =   30
      Index           =   -2147483648
      IsRunning       =   False
      LockedInPosition=   False
      PauseMilliseconds=   5
      Scope           =   2
      TabPanelIndex   =   0
   End
   Begin TextArea TextArea1
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   True
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   429
      HelpTag         =   ""
      HideSelection   =   True
      Index           =   -2147483648
      Italic          =   False
      Left            =   917
      LimitText       =   0
      LineHeight      =   0.0
      LineSpacing     =   1.0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Mask            =   ""
      Multiline       =   True
      ReadOnly        =   False
      Scope           =   2
      ScrollbarHorizontal=   False
      ScrollbarVertical=   True
      Styled          =   True
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Just a place to type while things are going on."
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   177
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   178
   End
   Begin Label lblProgress
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   66
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Status"
      TextAlign       =   1
      TextColor       =   &c00000000
      TextFont        =   "SmallSystem"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   638
      Transparent     =   False
      Underline       =   False
      Visible         =   False
      Width           =   710
   End
   Begin Separator Separator1
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   488
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   847
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   150
      Transparent     =   False
      Visible         =   True
      Width           =   4
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  App.AutoQuit = true
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Sub UpdateControls()
		  if CreateLooper.IsRunning then
		    btnMove.Enabled = false
		    btnMove.Caption = "Move"
		    btnCreate.Caption = "Stop"
		    fldRowCount.Enabled = false
		    ProgressBar1.Visible = true
		    lblProgress.Visible = true
		    
		  elseif MoveLooper.IsRunning then
		    btnCreate.Enabled = false
		    btnCreate.Caption = "Create"
		    btnMove.Caption = "Stop"
		    fldRowCount.Enabled = false
		    ProgressBar1.Visible = true
		    lblProgress.Visible = true
		    
		  else
		    btnCreate.Enabled = true
		    btnCreate.Caption = "Create"
		    btnMove.Enabled = true
		    btnMove.Caption = "Move"
		    ProgressBar1.Visible = false
		    lblProgress.Visible = false
		    fldRowCount.Enabled = true
		    ProgressBar1.Visible = false
		    lblProgress.Visible = false
		    
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub UpdateProgress(newValue As Integer)
		  ProgressBar1.Value = newValue
		  lblProgress.Text = format( newValue, "#,0" ) + " of " + format( ProgressBar1.Maximum, "#,0" )
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private RowCount As Integer = 1000000
	#tag EndProperty


#tag EndWindowCode

#tag Events btnCreate
	#tag Event
		Sub Action()
		  if CreateLooper.IsRunning then
		    
		    CreateLooper.Stop
		    
		  else
		    
		    RowCount = fldRowCount.Text.Val
		    
		    ProgressBar1.Maximum = RowCount
		    UpdateProgress 0
		    
		    CreateLooper.Start
		    
		  end if
		  
		  UpdateControls
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnMove
	#tag Event
		Sub Action()
		  if MoveLooper.IsRunning then
		    
		    MoveLooper.Stop
		    
		  else
		    
		    ProgressBar1.Maximum = lbOriginal.ListCount
		    UpdateProgress 0
		    
		    MoveLooper.Start
		    
		  end if
		  
		  UpdateControls
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events fldRowCount
	#tag Event
		Sub Open()
		  me.Text = str( RowCount )
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events CreateLooper
	#tag Event , Description = 546865207461736B20697320636F6D706C6574652E
		Sub Finished()
		  UpdateControls
		End Sub
	#tag EndEvent
	#tag Event , Description = 50726F6365737320612073696E676C6520697465726174696F6E207573696E6720796F7572206C6F6F7020696E6465782C206574632E
		Sub HandleIteration()
		  static r as new Random
		  
		  lbOriginal.AddRow str( lbOriginal.ListCount + 1 )
		  lbOriginal.Cell( lbOriginal.LastIndex, 1 ) = str( r.InRange( 0, RowCount ) )
		End Sub
	#tag EndEvent
	#tag Event , Description = 496E6372656D656E7420746865206C6F6F7020696E6465782C206574632E20616E642072657475726E20547275652E20496620646F6E652C2072657475726E2046616C73652E2057696C6C2062652063616C6C6564206265666F726520746865207374617274206F66206561636820697465726174696F6E2C20696E636C7564696E67207468652066697273742E
		Function MoveNext() As Boolean
		  return lbOriginal.ListCount < RowCount
		  
		End Function
	#tag EndEvent
	#tag Event , Description = 546865207461736B206861732070617573656420697465726174696F6E7320746F20616C6C6F77206F746865722070726F63657373696E672E2049742077696C6C20726573756D652061667465722050617573654D696C6C697365636F6E6473206D696C6C697365636F6E64732E
		Sub Paused()
		  UpdateProgress lbOriginal.ListCount
		  
		End Sub
	#tag EndEvent
	#tag Event , Description = 546865207461736B2069732061626F757420746F2073746172742E20526573657420616E79206C6F6F7020636F756E746572732C206574632E2C2075736564207768696C652070726F63657373696E67206561636820697465726174696F6E206F6620746865206C6F6F702E
		Sub Reset()
		  lbOriginal.DeleteAllRows
		  lbMoved.DeleteAllRows
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events MoveLooper
	#tag Event , Description = 546865207461736B20697320636F6D706C6574652E
		Sub Finished()
		  UpdateControls
		  
		End Sub
	#tag EndEvent
	#tag Event , Description = 50726F6365737320612073696E676C6520697465726174696F6E207573696E6720796F7572206C6F6F7020696E6465782C206574632E
		Sub HandleIteration()
		  static r as new Random
		  
		  dim row as integer = r.InRange( 0, lbOriginal.ListCount - 1 )
		  dim rowText as string = lbOriginal.Cell( row, 0 )
		  dim data as string = lbOriginal.Cell( row, 1 )
		  
		  lbMoved.AddRow rowText
		  lbMoved.Cell( lbMoved.LastIndex, 1 ) = data
		  
		  lbOriginal.RemoveRow row
		  
		End Sub
	#tag EndEvent
	#tag Event , Description = 496E6372656D656E7420746865206C6F6F7020696E6465782C206574632E20616E642072657475726E20547275652E20496620646F6E652C2072657475726E2046616C73652E2057696C6C2062652063616C6C6564206265666F726520746865207374617274206F66206561636820697465726174696F6E2C20696E636C7564696E67207468652066697273742E
		Function MoveNext() As Boolean
		  return lbOriginal.ListCount <> 0
		  
		End Function
	#tag EndEvent
	#tag Event , Description = 546865207461736B206861732070617573656420697465726174696F6E7320746F20616C6C6F77206F746865722070726F63657373696E672E2049742077696C6C20726573756D652061667465722050617573654D696C6C697365636F6E6473206D696C6C697365636F6E64732E
		Sub Paused()
		  UpdateProgress ProgressBar1.Maximum - lbOriginal.ListCount
		End Sub
	#tag EndEvent
	#tag Event , Description = 546865207461736B2069732061626F757420746F2073746172742E20526573657420616E79206C6F6F7020636F756E746572732C206574632E2C2075736564207768696C652070726F63657373696E67206561636820697465726174696F6E206F6620746865206C6F6F702E
		Sub Reset()
		  //
		  // Nothing to do here since it's based on lbOriginal
		  //
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Integer"
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
			"9 - Metal Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Integer"
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
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Background"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
#tag EndViewBehavior
