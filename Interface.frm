VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form InterfaceWindow 
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6900
   ScaleHeight     =   17.813
   ScaleMode       =   4  'Character
   ScaleWidth      =   57.5
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox WindowInformationBox 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6495
   End
   Begin VB.PictureBox SearchOptionsBox 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   4.063
      ScaleMode       =   4  'Character
      ScaleWidth      =   55.125
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox IgnoreNullValuesBox 
         Alignment       =   1  'Right Justify
         Caption         =   "Ignore null values."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton SearchButton 
         Caption         =   "&Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         ToolTipText     =   "Gives the command to search any matching windows for properties."
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox WindowTextBox 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Enter the text contained by the windows to search for properties here."
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label WindowLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Window:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid PropertyTable 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Displays the search results."
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This class contains this window's main interface window.
Option Explicit
Private WindowIndexes() As Long   'Contains the displayed windows' indexes.
'This procedure displays the properties for the specified window.
Private Sub DisplayWindowProperties(WindowHandle As Long, WindowIndex As Long, Optional IsChild As Boolean = False)
On Error GoTo ErrorTrap
Dim PropertyIndex As Long

   ReDim Properties(0 To 0) As PropertyStr
   CheckForError EnumPropsExA(WindowHandle, AddressOf HandleWindowProperties, CLng(0))

   With PropertyTable
      For PropertyIndex = LBound(Properties()) To UBound(Properties())
         If Not (Properties(PropertyIndex).Value = vbNullString And IgnoreNullValuesBox.Value = vbChecked) Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            ReDim Preserve WindowIndexes(LBound(WindowIndexes()) To .Row) As Long
            WindowIndexes(UBound(WindowIndexes())) = WindowIndex

            .Col = 0
            If IsChild Then .CellAlignment = flexAlignRightCenter Else .CellAlignment = flexAlignLeftCenter
            .Text = CStr(WindowHandle)
            .Col = 1: .CellAlignment = flexAlignLeftCenter: .Text = Escape(Properties(PropertyIndex).NameV)
            .Col = 2: .CellAlignment = flexAlignLeftCenter: .Text = Escape(Properties(PropertyIndex).Value)
            DoEvents
         End If
      Next PropertyIndex
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure initializes the property table.
Private Sub InitializePropertyTable()
On Error GoTo ErrorTrap

   With PropertyTable
      .Rows = 1
      .Row = 0
      .Col = 0: .Text = "Window Handle:"
      .Col = 1: .Text = "Property Name:"
      .Col = 2:  .Text = "Property Value:"
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   With App
      Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & App.CompanyName
   End With
   
   InitializePropertyTable
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure adjusts the size and position of the objects to the new size of the window.
Private Sub Form_Resize()
On Error Resume Next
Dim Column As Long

   With PropertyTable
      .Width = (Me.ScaleWidth - 1) - PropertyTable.Left
      .Height = (Me.ScaleHeight - SearchOptionsBox.ScaleHeight) - 3.5
   
      For Column = 0 To .Cols - 1
         .ColWidth(Column) = Me.Width / 3.2
      Next Column
   End With
   
   SearchOptionsBox.Width = Me.ScaleWidth - 1
   IgnoreNullValuesBox.Left = SearchOptionsBox.ScaleWidth - (IgnoreNullValuesBox.Width + 1)
   SearchButton.Left = SearchOptionsBox.ScaleWidth - (SearchButton.Width + 1)
   WindowInformationBox.Width = Me.ScaleWidth - 2
   WindowInformationBox.Top = Me.ScaleHeight - 2
   WindowTextBox.Width = SearchOptionsBox.ScaleWidth - (WindowLabel.Width + SearchButton.Width + 4)
End Sub


'This procedure updates the window information box when the selection changes.
Private Sub PropertyTable_SelChange()
On Error GoTo ErrorTrap
Dim WindowIndex As Long

   WindowIndex = WindowIndexes(PropertyTable.Row)
   WindowInformationBox.Text = """" & Windows(WindowIndex).Text & """ [" & Windows(WindowIndex).ClassName & "]"
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to search the windows matching the specified criteria for any properties present.
Private Sub SearchButton_Click()
On Error GoTo ErrorTrap
Dim OtherWindowIndex As Long
Dim WindowIndex As Long

   SearchOptionsBox.Enabled = False

   CheckForError , ResetSuppression:=True
   InitializePropertyTable
   WindowInformationBox.Text = vbNullString
   
   ReDim WindowIndexes(0 To 0) As Long
   
   ReDim Windows(0 To 0) As WindowStr
   CheckForError EnumWindows(AddressOf HandleWindows, CLng(0))
     
   For WindowIndex = LBound(Windows()) To UBound(Windows())
      With Windows(WindowIndex)
         If InStr(LCase$(Trim$(.Text)), LCase$(Trim$(WindowTextBox.Text))) > 0 Then
            DisplayWindowProperties .Handle, WindowIndex
            For OtherWindowIndex = LBound(Windows()) To UBound(Windows())
               If .Handle = Windows(OtherWindowIndex).Parent Then
                  DisplayWindowProperties Windows(OtherWindowIndex).Handle, OtherWindowIndex, IsChild:=True
               End If
            Next OtherWindowIndex
         End If
      End With
   Next WindowIndex
   
EndRoutine:
   SearchOptionsBox.Enabled = True
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


