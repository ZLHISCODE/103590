VERSION 5.00
Begin VB.Form frmIdentify泸州 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picKeyboard 
      BackColor       =   &H00008080&
      Height          =   3345
      Left            =   3810
      ScaleHeight     =   3285
      ScaleWidth      =   2940
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   3000
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   0
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2520
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "0"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   1
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1695
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "1"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   2
         Left            =   1035
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1695
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "2"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   3
         Left            =   2040
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1695
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "3"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   4
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   870
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "4"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   5
         Left            =   1035
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   870
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "5"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   6
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   870
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "6"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   7
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "7"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   8
         Left            =   1035
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "8"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlKey 
         Height          =   720
         Index           =   9
         Left            =   2040
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "9"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlClear 
         Height          =   720
         Left            =   1035
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2520
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "清除"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
      Begin zl9NewQuery.ctlButton ctlOK 
         Height          =   720
         Left            =   2040
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         Caption         =   "确定"
         AutoSize        =   0   'False
         ButtonHeight    =   600
      End
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   1965
      MouseIcon       =   "frmIdentify泸州.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1710
      Width           =   195
   End
   Begin VB.TextBox txtPsw 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1965
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   795
      Width           =   1620
   End
   Begin zl9NewQuery.ctlButton ctlCancel 
      Height          =   570
      Left            =   975
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1005
      Caption         =   "取消退出"
      AutoSize        =   0   'False
      ButtonHeight    =   450
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 离休人员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   240
      Left            =   2085
      MouseIcon       =   "frmIdentify泸州.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmIdentify泸州.frx":0614
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "请在读卡器的绿灯亮了之后，输入密码。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   525
      Left            =   765
      TabIndex        =   3
      Top             =   135
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "输入密码:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   315
      Left            =   795
      TabIndex        =   2
      Top             =   870
      Width           =   1440
   End
End
Attribute VB_Name = "frmIdentify泸州"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrPass As String
Private mstrType As String
Private mblnOK As Boolean

Public Function ShowForm(ByRef strPass As String, ByRef strType As String) As Boolean

    mstrPass = ""
    mblnOK = False
    
    ctlKey(0).ShowPicture = False
    ctlKey(1).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(2).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(3).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(4).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(5).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(6).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(7).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(8).ShowPicture = ctlKey(0).ShowPicture
    ctlKey(9).ShowPicture = ctlKey(0).ShowPicture
    ctlClear.ShowPicture = ctlKey(0).ShowPicture
    ctlOK.ShowPicture = ctlKey(0).ShowPicture
    
    
    Me.Show 1
    
    strPass = mstrPass
    strType = mstrType
    ShowForm = mblnOK
    
End Function

Private Sub DoEnter()
    mstrPass = txtPsw.Text
    mstrType = IIf(chk.Value = 1, "1", "0")
    
    mblnOK = True
    Unload Me
End Sub

Private Sub chk_Click()
    txtPsw.SetFocus
End Sub

Private Sub ctlCancel_CommandClick()
    mstrPass = ""
    Unload Me
End Sub

Private Sub ctlClear_CommandClick()
    txtPsw.Text = ""
    txtPsw.SetFocus
End Sub

Private Sub ctlKey_CommandClick(Index As Integer)
    txtPsw.Text = txtPsw.Text & Index
    txtPsw.SetFocus
    SendKeys "{END}"
End Sub



Private Sub ctlOK_CommandClick()
    
    Call DoEnter
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKey0, vbKeyNumpad0
        Call ctlKey_CommandClick(0)
    Case vbKey1, vbKeyNumpad1
        Call ctlKey_CommandClick(1)
    Case vbKey2, vbKeyNumpad2
        Call ctlKey_CommandClick(2)
    Case vbKey3, vbKeyNumpad3
        Call ctlKey_CommandClick(3)
    Case vbKey4, vbKeyNumpad4
        Call ctlKey_CommandClick(4)
    Case vbKey5, vbKeyNumpad5
        Call ctlKey_CommandClick(5)
    Case vbKey6, vbKeyNumpad6
        Call ctlKey_CommandClick(6)
    Case vbKey7, vbKeyNumpad7
        Call ctlKey_CommandClick(7)
    Case vbKey8, vbKeyNumpad8
        Call ctlKey_CommandClick(8)
    Case vbKey9, vbKeyNumpad9
        Call ctlKey_CommandClick(9)
    Case vbKeySeparator, vbKeyReturn
        Call ctlOK_CommandClick
    Case vbKeyDecimal, vbKeyDelete
        Call ctlClear_CommandClick
    End Select
End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C0, , True)
End Sub


Private Sub lblCheck_Click()
    If chk.Value = 1 Then
        chk.Value = 0
    Else
        chk.Value = 1
    End If
End Sub

Private Sub picKeyboard_Paint()
    Call RaisEffect(picKeyboard, -1)
    Call DrawColorToColor(picKeyboard, picKeyboard.BackColor, &HFFC0C0)
End Sub
