Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions and structures used by this program.
Private Const EM_GETPASSWORDCHAR As Long = &HD2&
Private Const EM_SETPASSWORDCHAR As Long = &HCC&
Private Const ERROR_IO_PENDING As Long = 997
Private Const ERROR_NOT_ALL_ASSIGNED As Long = 1300
Private Const ERROR_SUCCESS As Long = 0
Private Const ES_PASSWORD As Long = &H20&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const GWL_STYLE As Long = -16
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_DISABLED As Long = &H0&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2&
Private Const TOKEN_ALL_ACCESS As Long = &HFF&
Private Const WM_GETTEXT As Long = &HD&
Private Const WM_GETTEXTLENGTH As Long = &HE&

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Public Declare Function EnumPropsExA Lib "User32.dll" (ByVal hwnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "User32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetClassNameA Lib "User32.dll" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function EnumChildWindows Lib "User32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long
Private Declare Function GetParent Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLongA Lib "User32.dll" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function LookupPrivilegeValueA Lib "Advapi32.dll" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function lstrcpyA Lib "Kernel32.dll" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "Kernel32.dll" (ByVal lpString As Any) As Long
Private Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessH As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function PostMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageW Lib "User32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WaitMessage Lib "User32.dll" () As Long
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'The constants used by this program.
Private Const MAX_STRING As Long = 65535   'Defines the maximum number of characters used for a string buffer.
Public Const NO_HANDLE As Long = 0         'Indicates no handle.

'This structure defines a window's property.
Public Type PropertyStr
   NameV As String        'Defines a property's name.
   Value As String        'Defines a property's value.
   WindowHandle As Long   'Defines a property's window handle.
End Type

'This structure defines the properties of a window.
Public Type WindowStr
   ClassName As String   'The window's class.
   Handle As Long        'The window handle.
   Parent As Long        'The window's parent.
   Text As String        'The window's text.
End Type

Public Properties() As PropertyStr   'Contains a window's properties.
Public Windows() As WindowStr        'Contains all open windows found.

'This procedure checks whether an error has occurred during the most recent Windows API call.
Public Function CheckForError(Optional ReturnValue As Long = 0, Optional ResetSuppression As Boolean = False, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String
Static SuppressAPIErrors As Boolean

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   
   If ResetSuppression Then SuppressAPIErrors = False
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCrLf
      Message = Message & "Continue displaying API error messages?"
      If Not SuppressAPIErrors Then SuppressAPIErrors = (MsgBox(Message, vbYesNo Or vbExclamation) = vbNo)
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure copies the string with the specified pointer to the specified target string.
Private Function CopyString(Target As String, SourceP As Long) As String
On Error GoTo ErrorTrap
   Target = String$(CheckForError(lstrlenA(SourceP)), vbNullChar)
   CheckForError lstrcpyA(Target, SourceP)
EndRoutine:
   CopyString = Target
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure converts non-displayable characters in the specified text to escape sequences.
Public Function Escape(Text As String, Optional EscapeCharacter As String = "/", Optional EscapeLineBreaks As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = vbNullString
   Index = 1
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         Escaped = Escaped & String$(2, EscapeCharacter)
      ElseIf Character = vbTab Or Character >= " " Then
         Escaped = Escaped & Character
      ElseIf Character & NextCharacter = vbCrLf And Not EscapeLineBreaks Then
         Escaped = Escaped & vbCrLf
         Index = Index + 1
      Else
         Escaped = Escaped & EscapeCharacter & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Escape = Escaped
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure returns the specified window's class.
Public Function GetWindowClass(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim WindowClass As String

   WindowClass = String$(MAX_STRING, vbNullChar)
   Length = CheckForError(GetClassNameA(WindowH, ByVal WindowClass, Len(WindowClass)))
   WindowClass = Left$(WindowClass, Length)
EndRoutine:
   GetWindowClass = WindowClass
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure collects the information of the most recently found active window.
Public Sub GetWindowInformation(WindowH As Long)
On Error GoTo ErrorTrap
   ReDim Preserve Windows(LBound(Windows()) To UBound(Windows()) + 1) As WindowStr
   
   With Windows(UBound(Windows()))
      .ClassName = GetWindowClass(WindowH)
      .Handle = WindowH
      .Parent = CheckForError(GetParent(WindowH))
      .Text = GetWindowText(WindowH)
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the text contained by the specified window.
Public Function GetWindowText(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim PasswordCharacter As Long
Dim WindowText As String

   If WindowHasStyle(WindowH, ES_PASSWORD) Then
      PasswordCharacter = CheckForError(SendMessageW(WindowH, EM_GETPASSWORDCHAR, CLng(0), CLng(0)))
      If Not PasswordCharacter = 0 Then
         CheckForError PostMessageA(WindowH, EM_SETPASSWORDCHAR, CLng(0), CLng(0))
         Sleep CLng(1000)
      End If
   End If
   
   WindowText = String$(CheckForError(SendMessageW(WindowH, WM_GETTEXTLENGTH, CLng(0), CLng(0))) + 1, vbNullChar)
   Length = CheckForError(SendMessageW(WindowH, WM_GETTEXT, Len(WindowText), StrPtr(WindowText)))
   
   If Not PasswordCharacter = 0 Then
      CheckForError PostMessageA(WindowH, EM_SETPASSWORDCHAR, PasswordCharacter, CLng(0))
   End If
   
   WindowText = Left$(WindowText, Length)
EndRoutine:
   GetWindowText = WindowText
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any child windows that are found.
Private Function HandleChildWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
   GetWindowInformation hwnd
EndRoutine:
   HandleChildWindows = CLng(True) 'Indicates to continue enumerating child windows.
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any errors that occur and notifies the user.
Public Sub HandleError()
   MsgBox "Error: " & CStr(Err.Number) & vbCr & Err.Description, vbExclamation
End Sub

'This procedure handles any windows that are found.
Public Function HandleWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
   GetWindowInformation hwnd
   CheckForError EnumChildWindows(hwnd, AddressOf HandleChildWindows, CLng(0))
EndRoutine:
   HandleWindows = CLng(True) 'Indicates to continue enumerating windows.
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure handles any window properties that are found.
Public Function HandleWindowProperties(ByVal hwnd As Long, ByVal lpszString As Long, ByVal hData As Long, ByVal dwData As Long) As Long
On Error GoTo ErrorTrap
   With Properties(UBound(Properties()))
      .NameV = CopyString(.NameV, lpszString)
      .Value = CopyString(.Value, GlobalLock(hData))
      .WindowHandle = hwnd
      CheckForError GlobalUnlock(hData)
   End With
   
   ReDim Preserve Properties(LBound(Properties()) To UBound(Properties()) + 1) As PropertyStr
EndRoutine:
   HandleWindowProperties = CLng(True) 'Indicates to continue enumerating window properties.
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   SetDebugPrivilege Disabled:=False
   
   InterfaceWindow.Show
   Do While DoEvents()
      CheckForError WaitMessage()
   Loop
   
   SetDebugPrivilege Disabled:=True
   End
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure enables/disables the debug privilege.
Private Sub SetDebugPrivilege(Disabled As Boolean)
On Error GoTo ErrorTrap
Dim Length As Long
Dim NewPrivileges As TOKEN_PRIVILEGES
Dim PreviousPrivileges As TOKEN_PRIVILEGES
Dim PrivilegeId As LUID
Dim ReturnValue As Long
Dim TokenH As Long

   ReturnValue = CheckForError(OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, TokenH))
   If Not ReturnValue = 0 Then
      ReturnValue = CheckForError(LookupPrivilegeValueA(vbNullString, SE_DEBUG_NAME, PrivilegeId), , Ignored:=ERROR_IO_PENDING)
      If Not ReturnValue = 0 Then
         NewPrivileges.Privileges(0).pLuid = PrivilegeId
         NewPrivileges.PrivilegeCount = CLng(1)
         
         If Disabled Then
            NewPrivileges.Privileges(0).Attributes = SE_PRIVILEGE_DISABLED
            CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length), , Ignored:=ERROR_NOT_ALL_ASSIGNED
         ElseIf Not Disabled Then
            NewPrivileges.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length)
         End If
      End If
      CheckForError CloseHandle(TokenH)
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure indicates whether a window has the specified style.
Public Function WindowHasStyle(WindowH As Long, Style As Long) As Boolean
On Error GoTo ErrorTrap
   WindowHasStyle = (CheckForError(GetWindowLongA(WindowH, GWL_STYLE) And Style) = Style)
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


