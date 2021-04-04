VERSION 5.00
Begin VB.Form frmCardPass 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3450
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Icon            =   "frmCardPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "取消并退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3690
      MouseIcon       =   "frmCardPass.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2925
      Width           =   1860
   End
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   3660
      MouseIcon       =   "frmCardPass.frx":0316
      MousePointer    =   99  'Custom
      ScaleHeight     =   2880
      ScaleWidth      =   1920
      TabIndex        =   2
      Top             =   30
      Width           =   1920
      Begin VB.CommandButton cmdBtn 
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   12
         Left            =   660
         TabIndex        =   14
         Top             =   1725
         Width           =   1245
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   10
         Left            =   30
         TabIndex        =   13
         Top             =   2295
         Width           =   1860
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   9
         Left            =   30
         TabIndex        =   12
         Top             =   1725
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   8
         Left            =   1290
         MouseIcon       =   "frmCardPass.frx":0620
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1155
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   7
         Left            =   660
         TabIndex        =   10
         Top             =   1155
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   6
         Left            =   30
         TabIndex        =   9
         Top             =   1155
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   5
         Left            =   1290
         TabIndex        =   8
         Top             =   585
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   4
         Left            =   660
         TabIndex        =   7
         Top             =   585
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   3
         Left            =   30
         TabIndex        =   6
         Top             =   585
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   2
         Left            =   1290
         TabIndex        =   5
         Top             =   15
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   15
         Width           =   615
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   15
         Width           =   615
      End
   End
   Begin VB.TextBox txt 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1350
      Locked          =   -1  'True
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1545
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   795
      TabIndex        =   15
      Top             =   1620
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmCardPass.frx":092A
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入您的就诊卡密码。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   825
      TabIndex        =   0
      Top             =   615
      Width           =   2310
   End
End
Attribute VB_Name = "frmCardPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mvarCount As Long
Private mvarPass As String
Private mvarOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdBtn_Click(Index As Integer)

    Dim strInputPassword As String
    
    Select Case cmdBtn(Index).Caption
    Case "确定"
       If mvarPass = "" Then mvarOK = True: mvarPass = Me.txt.Text: cmdClose_Click: Exit Sub
        strInputPassword = zlCommFun.zlStringEncode(txt.Text)
        If strInputPassword = mvarPass Then
'        If txt.Text = mvarPass Then
            mvarOK = True
            Call cmdClose_Click
        Else
            mvarCount = mvarCount + 1
'            If mvarCount = 3 Then
'                MsgBox "您输入就诊卡密码已经三次不成功！", vbInformation, gstrSysName
'                Call cmdClose_Click
'            Else
                MsgBox "您输入的密码不对，请重新输入！", vbInformation, gstrSysName
                txt.Text = ""
'            End If
        End If
    Case "清除"
        txt.Text = ""
        txt.SetFocus
    Case Else
        txt.Text = txt.Text & Trim(cmdBtn(Index).Caption)
        txt.SetFocus
        
    End Select
End Sub

Private Sub cmdBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Function ShowCardPass(ByVal v_Pass As String) As Boolean
    mvarOK = False
    mvarPass = v_Pass
    frmCardPass.Show 1
    ShowCardPass = mvarOK
End Function
Public Function GetCardPass(ByRef v_Pass As String) As String
    mvarOK = False
    mvarPass = ""
    frmCardPass.Show 1
    If mvarOK Then v_Pass = mvarPass
    GetCardPass = mvarOK
End Function

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey0, vbKeyNumpad0
            txt.Text = txt.Text & "0"
            SendKeys "{END}"
        Case vbKey1, vbKeyNumpad1
            txt.Text = txt.Text & "1"
            SendKeys "{END}"
        Case vbKey2, vbKeyNumpad2
            txt.Text = txt.Text & "2"
            SendKeys "{END}"
        Case vbKey3, vbKeyNumpad3
            txt.Text = txt.Text & "3"
            SendKeys "{END}"
        Case vbKey4, vbKeyNumpad4
            txt.Text = txt.Text & "4"
            SendKeys "{END}"
        Case vbKey5, vbKeyNumpad5
            txt.Text = txt.Text & "5"
            SendKeys "{END}"
        Case vbKey6, vbKeyNumpad6
            txt.Text = txt.Text & "6"
            SendKeys "{END}"
        Case vbKey7, vbKeyNumpad7
            txt.Text = txt.Text & "7"
            SendKeys "{END}"
        Case vbKey8, vbKeyNumpad8
            txt.Text = txt.Text & "8"
            SendKeys "{END}"
        Case vbKey9, vbKeyNumpad9
            txt.Text = txt.Text & "9"
            SendKeys "{END}"
        Case vbKeySeparator, vbKeyReturn
            Call cmdBtn_Click(10)
        Case vbKeyDecimal, vbKeyDelete
            Call cmdBtn_Click(12)
        End Select
End Sub

Private Sub Form_Load()
    mvarCount = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub txt_GotFocus()
    SelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub
