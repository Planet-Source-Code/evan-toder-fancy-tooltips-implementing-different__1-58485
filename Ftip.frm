VERSION 5.00
Begin VB.Form Ftip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   7290
      Top             =   5175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7470
      Top             =   4950
   End
End
Attribute VB_Name = "Ftip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public calling_class    As cFancyTooltips
Public m_show_delay     As Long
Private timer_cnt       As Long
Private m_tip_arr_num   As Long
Private old_o           As Object

 
'
'this form will only get focus in one of two ways
'at form load (at such time we dont want to trigger
'the mouse down event) and when this form is clicked
'
Private Sub Form_GotFocus()
  Static m_b_loaded       As Boolean
  
  If m_b_loaded = False Then
     m_b_loaded = True
  Else
     calling_class.friend_mouse_down (code.m_arr_pallette(m_tip_arr_num).hwnd)
  End If
  
End Sub

Private Sub Form_Load()
   
   Timer1.Interval = 400
   Timer1.Enabled = True
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  If Visible = True Then Visible = False
  
End Sub

  
Private Sub paint_me()
  Dim o As Object
    
  On Error Resume Next
  
  old_o.Visible = False 'hide the last "pallette"
  Set o = m_arr_pallette(m_tip_arr_num) 'set reference
  Ftip.BackColor = o.BackColor 'this b color same as the new pallette
  o.Move 50, 50 'move and position the pallette
  Width = o.Width + 100 'resize this to match the new pallette
  Height = o.Height + 100
  o.Visible = True 'show it
  Set old_o = o 'store it as the old one for the next showing
  Me.Line (0, 0)-(Width - 20, Height - 20), RGB(100, 100, 120), B
  Set o = Nothing
  
End Sub
  

Public Sub Form_Unload(Cancel As Integer)

   Timer1.Interval = 0
   Timer1.Enabled = False
   Set old_o = Nothing
   
End Sub

'
'this function will ensure the tip stays within the
'boundries of the screens edges
'
Private Function tip_position(ByRef ret_x As Long, _
                       ByRef ret_y As Long) As Pointapi
                       
     Const SCREEN_LEFT = 0
     Const SCREEN_TOP = 0
     Const BUFFER = 12
    
     
    'start by attempting successful placement just
    'below the cursor and horizontally in the middle
    'of this tooltip
     Dim screen_right  As Long
     Dim screen_bottom As Long
     Dim half_tip_wid  As Long
     Dim tip_height    As Long
     Dim tip_wid       As Long
     
     'right edge of screen pixels
     screen_right = (Screen.Width / Screen.TwipsPerPixelX)
     'bottm edge screen pixels
     screen_bottom = (Screen.Height / Screen.TwipsPerPixelY)
     'half this tips (forms) width
     half_tip_wid = (Width * 0.5) / Screen.TwipsPerPixelX
     'the width of this tip
     tip_wid = (Width / Screen.TwipsPerPixelX)
     'the height of this tip
     tip_height = (Height / Screen.TwipsPerPixelY)
     
     
     '====== check the right and left edges of screen ========
     If (ret_x + half_tip_wid) > screen_right Then
         ret_x = (screen_right - tip_wid)
     'left side not ok
     ElseIf (ret_x - half_tip_wid) < SCREEN_LEFT Then
         ret_x = SCREEN_LEFT
     Else
          ret_x = (ret_x - half_tip_wid)
     End If
     
     '========= check the botom edges of screen ===========
     ' (top edge will never be factor unless app is upside down
     If (ret_y + tip_height + BUFFER) > screen_bottom Then
        ret_y = (ret_y - tip_height - (BUFFER * 2))
     Else
        ret_y = (ret_y + BUFFER)
     End If
     
  
End Function

 
Private Sub Timer1_Timer()
 
  If mod_hide_tooltips Then Exit Sub
 
  Dim bmatch        As Boolean
  Dim parent_hwnd   As Long
  Dim hwnd_und_mous As Long
  Dim lcnt          As Long
  

  hwnd_und_mous = code.hwnd_under_mouse
  parent_hwnd = GetParent(hwnd_und_mous)
  
  
  'see if the hwnd under mouse is in our array list
  'configured in the classes .Add method
  For lcnt = 0 To UBound(code.m_arr_ctrls)
         If hwnd_und_mous = code.m_arr_ctrls(lcnt) Then
            'if there isnt a hwnd match then this returns false
            'which, down below makes this tip hidden
            bmatch = True
            'this var holds the tip msg & is parsed in sub paint_me
            m_tip_arr_num = lcnt
            'start the timer that will show then hide the tip
            Timer2.Interval = 1000
            Timer2 = True
            Exit For
         End If
  Next lcnt
  
 
   
  'the mouse isnt over on of our controls but
  'maybe its over this form and the user wants it
  'that way (hide_on_mouseover = false)
  If Not (bmatch) Then
    If Not (mod_m_hide_on_mouseover) Then
      If parent_hwnd <> Me.hwnd Then
         If Visible Then Visible = False
      End If
    Else
         If Visible Then Visible = False
    End If
  End If
  
End Sub

Private Sub kill_timer2()
   
   Timer2.Enabled = False
   timer_cnt = 0
   
End Sub
Private Sub Timer2_Timer()
 
   If timer_cnt >= m_show_delay Then
     'pass the val of pt.x and pt.y to altered by the function
     'returning to us accurate placement of the tooltip
     Call tip_position(wind_pt.x, wind_pt.Y)
     'make this visible and put it on top
     SetWindowPos hwnd, HWND_TOPMOST, wind_pt.x, _
     wind_pt.Y, 0, 0, (SWP_NOSIZE Or SWP_NOACTIVATE)
     'without activating it
     ShowWindow hwnd, SW_SHOWNA
     Call paint_me
     timer_cnt = 0
     code.hwnd_old = code.hwnd_under_mouse
     Call kill_timer2
   End If
   
   timer_cnt = (timer_cnt + 1)
End Sub

