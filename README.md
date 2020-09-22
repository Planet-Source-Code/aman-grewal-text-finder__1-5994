<div align="center">

## Text Finder


</div>

### Description

This program uses API calls to find the text of almost any Window, including ones blocked with password characters. Just hover the mouse cursor over the window. Also shows Window handle. Makes use of SetWindowPos for "Always On Top" API call.

Note: This Visual Basic version makes use of the GetWindowText call, which does not always return Window text, while the SendMessage call with WM_GETTEXT seems to work more often. Contact me to obtain the VC++ version which uses SendMessage instead.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aman Grewal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aman-grewal.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aman-grewal-text-finder__1-5994/archive/master.zip)

### API Declarations

```
Private Type POINTAPI 'Simple point structure
    x As Long
    y As Long
End Type
'Returns mouse position as a point
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Returns HWND of Window that mouse cursor is currently over
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Returns length of Window text
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'Returns actual Window text
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Used to set Window to "Always On Top"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1     'Parameter to set Window to "Always On Top"
Const SWP_NOMOVE = &H2     'Don't move the Window
Const SWP_NOSIZE = &H1     'Don't resize the Window
```


### Source Code

```

Private Sub Form_Load() 'Set Window to "Always On Top"
  Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Private Sub tmrRefresh_Timer()
  Dim cursorPos As POINTAPI, textLength As Integer
  Dim hWnd As Long, winText As String
  Static prevHWnd As Long 'Store handle of previous Window
  Call GetCursorPos(cursorPos) 'Get current mouse position
  hWnd = WindowFromPoint(cursorPos.x, cursorPos.y) 'Get handle to Window mouse is over
  If prevHWnd <> hWnd Then 'If the Window mouse is the same as the previous Window that the mouse was over, don't refresh the information
    txtHWnd.Text = hWnd 'Show Window handle
    textLength = GetWindowTextLength(hWnd) + 1 'Get length of Window text
    winText = Space(textLength) 'Setup buffer to copy Window text
    Call GetWindowText(hWnd, winText, textLength) 'Get the actual text
    txtWinText.Text = winText
    prevHWnd = hWnd
  End If
End Sub
```

