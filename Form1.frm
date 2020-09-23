VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   3
      Left            =   4815
      TabIndex        =   10
      Top             =   4815
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "Go"
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "perhaps even hold controls"
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   3
         Left            =   450
         TabIndex        =   11
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   1410
      Index           =   2
      Left            =   3825
      TabIndex        =   8
      Top             =   3195
      Width           =   2760
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2745
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "www.planetsourcecode.com/vb"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   14
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ARE YOU            CODE HUNGRY  ?"
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   45
         TabIndex        =   13
         Top             =   0
         Width           =   2715
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Or maybe we can have a tooltip that mimicks the ones on PlanetSourceCode the display detailed data and are actually hyperlinks"
         ForeColor       =   &H80000017&
         Height          =   780
         Index           =   2
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "click me!!"
         Top             =   270
         Width           =   2760
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   1035
      Width           =   2310
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "Form1.frx":0000
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "How about adding a pretty picture ??"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Top             =   0
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   4230
      TabIndex        =   4
      Top             =   2250
      Width           =   2535
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "we can do so much more with our tooltips than we really think we can"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   555
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   90
         Width           =   2220
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   735
         Index           =   0
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   2445
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1485
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1185
      Left            =   135
      TabIndex        =   2
      Top             =   1260
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1665
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   405
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "more info about this object and how to use it"
      Height          =   645
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents cTips As cFancyTooltips
Attribute cTips.VB_VarHelpID = -1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Command1_Click()
   cTips.help_about
End Sub

Private Sub cTips_MouseDown(pallette_hwnd As Long)
 

 If pallette_hwnd = Frame2(2).hwnd Then
    ShellExecute hwnd, "open", "http://www.planetsourcecode.com/vb", vbNullString, vbNullString, 1
 Else
    Debug.Print pallette_hwnd & " was clicked"
 End If

End Sub

Private Sub Form_Load()
  
  Set cTips = New cFancyTooltips
  Set cTips.your_form = Me
  cTips.add_ctrl Command1.hwnd, Frame2(0)
  cTips.add_ctrl Text1.hwnd, Frame2(1)
  cTips.add_ctrl Frame1.hwnd, Frame2(2)
  cTips.add_ctrl Combo1.hwnd, Frame2(3)
  cTips.hide_tooltips = False
 
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set cTips = Nothing
End Sub
