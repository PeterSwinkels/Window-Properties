VERSION 5.00
Begin VB.Form InterfaceWindow 
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5205
   ScaleHeight     =   13.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   43.375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PropertiesBox 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      ToolTipText     =   "Displays a list of any window properties found."
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox WindowTextBox 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Enter the text contained by the windows to search for properties here."
      Top             =   120
      Width           =   1815
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
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "Gives the command to search any matching windows for properties."
      Top             =   120
      Width           =   1215
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This class contains this window's main interface window.
Option Explicit

'This procedure displays the properties for the specified window.
Private Sub DisplayWindowProperties(WindowHandle As Long, Optional IsChild As Boolean = False)
On Error GoTo ErrorTrap
Dim PropertyIndex As Long

   ReDim Properties(0 To 0) As PropertyStr
   CheckForError EnumPropsExA(WindowHandle, AddressOf HandleWindowProperties, CLng(0))

   With PropertiesBox
      For PropertyIndex = LBound(Properties()) To UBound(Properties()) - 1
         If Properties(PropertyIndex).WindowHandle = WindowHandle Then
            If IsChild Then .Text = .Text & "  " Else .Text = .Text & " "
            .Text = .Text & "-" & Properties(PropertyIndex).NameV & " = """ & Escape(Properties(PropertyIndex).Value) & """" & vbCrLf
         End If
      Next PropertyIndex
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure formats the specified window text to be suitable for display in the search results.
Private Function FormatWindowText(WindowText As String) As String
On Error GoTo ErrorTrap
Dim FormattedText As String

   FormattedText = Trim$(Replace(WindowText, vbTab, " "))
   If InStr(FormattedText, vbCrLf) Then
      FormattedText = Left$(FormattedText, InStr(FormattedText, vbCrLf))
   ElseIf InStr(FormattedText, vbCr) Then
      FormattedText = Left$(FormattedText, InStr(FormattedText, vbCr))
   ElseIf InStr(FormattedText, vbLf) Then
      FormattedText = Left$(FormattedText, InStr(FormattedText, vbLf))
   End If
   
EndRoutine:
   FormatWindowText = FormattedText
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   With App
      Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & App.CompanyName
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure adjusts the size and position of the objects to the new size of the window.
Private Sub Form_Resize()
On Error Resume Next
   PropertiesBox.Width = (Me.ScaleWidth - 2) - PropertiesBox.Left
   PropertiesBox.Height = Me.ScaleHeight - 3
   SearchButton.Left = Me.ScaleWidth - (SearchButton.Width + 1)
   WindowTextBox.Width = Me.ScaleWidth - (WindowLabel.Width + SearchButton.Width + 4)
End Sub


'This procedure gives the command to search the specified for any properties present.
Private Sub SearchButton_Click()
On Error GoTo ErrorTrap
Dim OtherWindowIndex As Long
Dim WindowIndex As Long
 
   With PropertiesBox
      .Text = vbNullString
      
      ReDim Windows(0 To 0) As WindowStr
      CheckForError EnumWindows(AddressOf HandleWindows, CLng(0))

      For WindowIndex = LBound(Windows()) To UBound(Windows())
         If InStr(LCase$(Trim$(Windows(WindowIndex).Text)), LCase$(Trim$(WindowTextBox.Text))) > 0 Then
            .Text = .Text & "[" & FormatWindowText(Windows(WindowIndex).Text) & "] (" & Windows(WindowIndex).ClassName & ")" & vbCrLf
            DisplayWindowProperties Windows(WindowIndex).Handle
            For OtherWindowIndex = LBound(Windows()) To UBound(Windows())
               If Windows(OtherWindowIndex).Parent = Windows(WindowIndex).Handle Then
                  .Text = .Text & " [" & FormatWindowText(Windows(OtherWindowIndex).Text) & "] (" & Windows(OtherWindowIndex).ClassName & ")" & vbCrLf
                  DisplayWindowProperties Windows(OtherWindowIndex).Handle, IsChild:=True
               End If
            Next OtherWindowIndex
            .Text = .Text & vbCrLf
         End If
      Next WindowIndex
   End With
   
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


