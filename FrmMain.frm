VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fun with Translucency!"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Push Me to E&xit"
      Height          =   405
      Left            =   2520
      TabIndex        =   6
      Top             =   1860
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "                                       "
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   4125
      Begin VB.CheckBox TransCheck 
         Caption         =   "Enable Translucency"
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   -90
         Width           =   1785
      End
      Begin MSComctlLib.Slider TransSlider 
         Height          =   630
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1111
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   20
         SmallChange     =   5
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   5
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "More Visible"
         Height          =   465
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Less Visible "
         Height          =   465
         Left            =   3390
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "The Roo Group / Andrew Saturn"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   1920
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "This is a simple demo to show off the super nifty cool translucency API in Microsoft® Windows 2000 and XP."
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   4005
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
FrmTrans.Show

'globals needed:
'DoEvents
'Dim NormalWindowStyle As Long

'turns transparency on:
'NormalWindowStyle = GetWindowLong(FrmTrans.hwnd, GWL_EXSTYLE)
'SetWindowLong FrmTrans.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

'turns transparency off:
'SetWindowLong FrmTrans.hwnd, GWL_EXSTYLE, GetWindowLong(FrmTrans.hwnd, GWL_EXSTYLE) And Not (WS_EX_LAYERED)

'sets the transparency level:
'the slider is 0-100 (like percent), but the API is 0-255 (like colors!).
'255 = 100% visible, 0 = 0% visible.

    'this example sets it to the value of the slider:
    'SetLayeredWindowAttributes FrmTrans.hwnd, 0, 255 * (1 - (Val(TransSlider.Value) / 100)), LWA_ALPHA

    'this one sets it to "155" (out of 255)
    'SetLayeredWindowAttributes Me.hwnd, 0, 155, LWA_ALPHA
End Sub

Private Sub TransCheck_Click()
'this code is for the checkbox. it turns the translucency off and on
'to save on processor usage, rather than just making it 0% and leaving
'it on.

DoEvents
Dim NormalWindowStyle As Long

If TransCheck.Value = 0 Then 'on -> off
    'turns the slider off
    TransSlider.Enabled = False
    'turns transparency off
    SetWindowLong FrmTrans.hwnd, GWL_EXSTYLE, GetWindowLong(FrmTrans.hwnd, GWL_EXSTYLE) And Not (WS_EX_LAYERED)

ElseIf TransCheck.Value = 1 Then 'off -> on
    'turns the slider on
    TransSlider.Enabled = True
    'turns transparency on
    NormalWindowStyle = GetWindowLong(FrmTrans.hwnd, GWL_EXSTYLE)
    SetWindowLong FrmTrans.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    'sets the transparency level
    SetLayeredWindowAttributes FrmTrans.hwnd, 0, 255 * (1 - (Val(TransSlider.Value) / 100)), LWA_ALPHA
End If

End Sub

Private Sub TransSlider_Scroll()
'this is the scroller thing. its set to _Scroll rather than _Change so
'that you see the effect instantly. if you don't want that (less processor
'usage) then just use _Change.

DoEvents
Dim NormalWindowStyle As Long
    
    'sets the transparency level
    SetLayeredWindowAttributes FrmTrans.hwnd, 0, 255 * (1 - (Val(TransSlider.Value) / 100)), LWA_ALPHA
End Sub
