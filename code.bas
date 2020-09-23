Attribute VB_Name = "code"
Option Explicit


Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As Pointapi) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As enHwnd, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As enSWP) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As enSw) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 
Public Type Pointapi
   x As Long
   Y As Long
End Type

Public Enum enSWP
    SWP_NOACTIVATE = &H10
    SWP_NOMOVE = &H2
    SWP_NOSIZE = &H1
End Enum

Public Enum enHwnd
    HWND_TOPMOST = -1
    HWND_TOP = 0
    HWND_BOTTOM = 1
    HWND_NOTOPMOST = -2
End Enum


Public Enum enSw
   SW_ERASE = &H4
   SW_FORCEMINIMIZE = 11
   SW_HIDE = 0
   SW_INVALIDATE = &H2
   SW_MAX = 10
   SW_MAXIMIZE = 3
   SW_MINIMIZE = 6
   SW_NORMAL = 1
   SW_RESTORE = 9
   SW_SHOW = 5
   SW_SHOWMINNOACTIVE = 7
   SW_SHOWNA = 8
   SW_SHOWNOACTIVATE = 4
   SW_SHOWNORMAL = 1
   SW_SMOOTHSCROLL = &H10
   SW_SHOWMAXIMIZED = 3
   SW_SHOWMINIMIZED = 2
   SW_SHOWDEFAULT = 10
   SW_SCROLLCHILDREN = &H1
   SW_PARENTOPENING = 3
   SW_PARENTCLOSING = 1
   SW_OTHERZOOM = 2
   SW_OTHERUNZOOM = 4
End Enum
 
Public hwnd_old                 As Long
Public m_arr_ctrls()            As Variant
Public m_arr_pallette()         As Object
Public wind_pt                  As Pointapi
Public mod_hide_tooltips        As Boolean
Public mod_m_hide_on_mouseover  As Boolean
'
'returns the hwnd of the window under the mouse
'
Function hwnd_under_mouse() As Long

  Dim x As Long
  
  GetCursorPos wind_pt
  x = WindowFromPoint(wind_pt.x, wind_pt.Y)
  hwnd_under_mouse = x
 
End Function
 
 
 
