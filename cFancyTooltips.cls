VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFancyTooltips"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Enum enShowDelay
   [1 second] = 1
   [2 seconds] = 2
   [3 seconds] = 3
   [4 seconds] = 4
End Enum
 
 
Private m_tip_show_delay          As enShowDelay
Private m_hide_on_mouseover       As Boolean
Private btag_show_delay           As Boolean
Private m_hide_tooltips           As Boolean
Private m_your_form               As Form
Attribute m_your_form.VB_VarHelpID = -1
 
Private Const m_def_tip_show_delay = 2

Event MouseDown(pallette_hwnd As Long)


Sub add_ctrl(ctrl_hwnd As Long, pallette_ctrl As Object)
  Dim u   As Long
  
  If m_your_form Is Nothing Then
     Err.Raise 23232, , "You must set the value of "".your_form"""
     Exit Sub
  End If
 
  u = upper
  
  ReDim Preserve code.m_arr_ctrls(u)
  ReDim Preserve code.m_arr_pallette(u)
  code.m_arr_ctrls(u) = ctrl_hwnd
  Set code.m_arr_pallette(u) = pallette_ctrl
  SetParent pallette_ctrl.hwnd, Ftip.hwnd
  
End Sub

Private Function upper() As Long
 
  If IsArray(m_arr_ctrls) Then
    upper = UBound(m_arr_ctrls) + 1
  Else
    upper = 0
  End If
  
End Function
'-- check to see if item is array or array is initialized
Function IsArray(varArray As Variant) As Boolean
Dim upper As Integer
On Error Resume Next
 
  upper = UBound(varArray)
  
  If Err.Number Then
     If Err.Number = 9 Then
       IsArray = False
     End If
  Else
     IsArray = True
  End If

End Function
 
 
Private Sub your_form_Unload(Cancel As Integer)
 
    Unload Ftip
    Set Ftip = Nothing
    
End Sub
'
'this sub can only be called from Ftip and would only
'be called because the [hide_on_mouseover] = false AND
'the user has clicked the "tooltip" (Ftip)
'
Friend Sub friend_mouse_down(tip_hwnd As Long)
 
 RaiseEvent MouseDown(tip_hwnd)

End Sub
  
 
 
Public Property Get your_form() As Object

  Set your_form = m_your_form

End Property

Public Property Set your_form(ByVal vNewValue As Object)
   
   On Error Resume Next
    
   Set m_your_form = vNewValue
 
   With Ftip
      'show it in a manner that makes it owned by the caller
      'that way this form will unload with the calling form
      'and wont be left hanging in memory
      .Show vbModeless, m_your_form
      'keep it hidden to start
      .Visible = False
      Set Ftip.calling_class = Me
   End With
   
End Property
Public Property Get hide_tooltips() As Boolean
   
   hide_tooltips = m_hide_tooltips
   
End Property
Public Property Let hide_tooltips(ByVal vNewValue As Boolean)
 
  'set the value of a tag
  m_hide_tooltips = vNewValue
  'the value is passed on to the public var
  'in the module [code] so Ftip can read it
  'and decide to show itself or not
  mod_hide_tooltips = vNewValue
  'this val is also passed to module var
  'so ftip can read it as to whether to
  'hide itself on mouseover
  mod_m_hide_on_mouseover = m_hide_on_mouseover
   'check to see if user specified a
   'tim_show_time and tip_show_delay
   'if not we will use the m_def defaults
  tip_show_delay = IIf(btag_show_delay = True, m_tip_show_delay, m_def_tip_show_delay)
   
  
End Property
 
 
Public Property Get tip_show_delay() As enShowDelay
  
  tip_show_delay = m_tip_show_delay
  
End Property
Public Property Let tip_show_delay(ByVal vNewValue As enShowDelay)
  
  btag_show_delay = True
  m_tip_show_delay = vNewValue
  'set the value in ftip [form]
  Ftip.m_show_delay = vNewValue
 
End Property
Public Property Get hide_on_mouseover() As Boolean
  
  hide_on_mouseover = m_hide_on_mouseover
  
End Property
Public Property Let hide_on_mouseover(ByVal vNewValue As Boolean)
  
  m_hide_on_mouseover = vNewValue
  
End Property
Function help_about() As String

 Const R = vbCrLf
 
 
 help_about = "FancyTooltips is a way of implementing different ""pallettes""  for different controls on your form." & R & _
              "a ""pallette"" being any valid container control such as a picturebox or a frame." & R & _
              "You simple draw the controls/objects onto the different pallettes in any manner you choose" & R & _
              "using different labels, lines, ..whatever in any manner u see appropriate." & R & _
              "You simply assign the pallette object as the second parameter for the ""add_ctrl"" method" & R & _
              "For example, say on your form you want a fancyTooltip for ""command1"".  You simply place (in the form_load)" & R & _
              " ""fancyTooltips1.add_ctrl command1, frame1""  assuming that you want ""frame1"" to be the pallette for command1"
 Debug.Print help_about
 
End Function
