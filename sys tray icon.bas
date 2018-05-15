Attribute VB_Name = "Declare"

'---------------------------------------'

'   http://www.VisualBasic.Blogfa.com   '

'   AliMedia_vb@Yahoo.com               '

'---------------------------------------'

Public Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA

  cbSize As Long           ' size of the structure
  hWnd As Long             ' the handle of the window
  uID As Long              ' an unique ID for the icon
  uFlags As Long           ' flags(see below)
  uCallbackMessage As Long ' the Msg that call back when a user do something to the icon
  hIcon As Long            ' the memory location of the icon
  szTip As String * 64     ' tooltip max 64 characters

End Type

Public Const NIM_ADD = &H0      ' add an icon to the system tray
Public Const NIM_MODIFY = &H1   ' modify an icon in the system tray
Public Const NIM_DELETE = &H2   ' delete an icon in the system tray
Public Const NIF_MESSAGE = &H1  ' whether a message is sent to the window procedure for events
Public Const NIF_ICON = &H2     ' whether an icon is displayed
Public Const NIF_TIP = &H4      ' tooltip availibility

Public formloaded As Boolean
Public oldproc As Long

Public Function proc&(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)

  ' right button release on the icon
  ' so popup a menu
  ' change 517 to:
  ' 516 --- right button down
  ' 518 --- right button double click
  ' 513 -- left button down
  ' 514 -- left button up
  ' 515 -- left button double click
  ' 519 -- middle button down ( for some mouse only )
  ' 520 -- middle button up
  ' 521 -- middle button double click
  If Msg = 1400 And lParam = 517 And formloaded Then 'Form29.PopupMenu Form29.mnufile
  End If
  ' let VB handle the rest
  proc = CallWindowProcA(oldproc, hWnd, Msg, wParam, lParam)

End Function
