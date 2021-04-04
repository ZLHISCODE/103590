VERSION 5.00
Begin VB.Form frm修改密码 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改密码"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frm修改密码.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt卡号 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   255
      Width           =   1875
   End
   Begin VB.OptionButton opt卡类别 
      Caption         =   "IC卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   1
      Left            =   660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   945
   End
   Begin VB.OptionButton opt卡类别 
      Caption         =   "磁卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   0
      Left            =   645
      TabIndex        =   10
      Top             =   75
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.CommandButton cmd读卡 
      Caption         =   "启动"
      Height          =   350
      Left            =   4185
      TabIndex        =   9
      Top             =   240
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1665
      TabIndex        =   7
      Top             =   2580
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3195
      TabIndex        =   8
      Top             =   2580
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -225
      TabIndex        =   6
      Top             =   2400
      Width           =   5865
   End
   Begin VB.TextBox txt确认新密码 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1890
      Width           =   2265
   End
   Begin VB.TextBox txt新密码 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2265
   End
   Begin VB.TextBox txt原密码 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2355
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   990
      Width           =   2265
   End
   Begin VB.Label lbl卡号 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1755
      TabIndex        =   13
      Top             =   315
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frm修改密码.frx":000C
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lbl确认新密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "确认新密码(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   585
      TabIndex        =   4
      Top             =   1950
      Width           =   1680
   End
   Begin VB.Label lbl新密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新密码(&N)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1095
      TabIndex        =   2
      Top             =   1500
      Width           =   1170
   End
   Begin VB.Label lbl原密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "原密码(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1095
      TabIndex        =   0
      Top             =   1050
      Width           =   1170
   End
End
Attribute VB_Name = "frm修改密码"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr新密码 As String
Private mstr旧密码 As String
Private mintLen As Integer
Private mintType As Integer
Private mstrSNO As String, mstrIDNO As String
Private mblnGo As Boolean
Private mblnKey As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(TXT新密码.Text) = "" Then
        MsgBox "密码不能为空！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    
    If TXT新密码.Text <> txt确认新密码.Text Then
        MsgBox "两次输入的新密码不一致，请重输！", vbInformation, gstrSysName
        TXT新密码.SetFocus
        Exit Sub
    End If
    mstr新密码 = TXT新密码.Text
    mstr旧密码 = TXT原密码.Text
    Unload Me
    Exit Sub
End Sub

Private Sub cmd读卡_Click()
    mblnGo = True
    mblnGo = ReadCard(TXT原密码, 0)
End Sub

Private Function ReadCard(ByVal objTxt As TextBox, ByVal intType As Integer) As Boolean
    '
    Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String, strSNO As String
    Dim lngHandle As Long
    Dim blnNO As Boolean
    
    On Error GoTo errHand
    
    If Not mblnGo Then Exit Function
    
    strPort = GetSetting("ZLSOFT", "公共模块\贵阳市医保", "端口", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If

    '打开读卡器
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Function
    End If

    '读取PSAM芯片号码
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        blnNO = True
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)

    If mintType = 2 And intType = 0 Then
         '读取社保卡号
        STRERR = Space(2000)
        mstrSNO = Space(2000)
        strPin = "000000"
        strAddr = "MF|EF05|07|$MF|EF06|01|$"
        If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, mstrSNO, STRERR) <> 0 Then
            MsgBox STRERR, vbInformation, gstrSysName
            blnNO = True
            GoTo Exith
        End If
        mstrSNO = TruncZero(mstrSNO)
        mstrIDNO = Split(mstrSNO, "|")(5)
        mstrSNO = Split(mstrSNO, "|")(2)
        txt卡号.Text = mstrSNO
    End If
    If objTxt.Enabled = True Then objTxt.SetFocus
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        blnNO = True
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    objTxt.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    If blnNO = True Then Exit Function
    Select Case objTxt.Name
        Case "txt原密码"
            Call txt原密码_KeyDown(vbKeyReturn, 0)
        Case "txt新密码"
            Call TXT新密码_KeyDown(vbKeyReturn, 0)
        Case "txt确认新密码"
            Call txt确认新密码_KeyDown(vbKeyReturn, 0)
    End Select
    ReadCard = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Function

Private Sub Form_Activate()
    'Modified by ZYB 毕节
    If TXT原密码.Text <> "" Then Me.TXT新密码.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab): mblnKey = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    TXT原密码.Text = mstr新密码
    mstr新密码 = ""
    mblnGo = False
    mblnKey = False
    If gintType = 1 Then
        opt卡类别.Item(0).Value = True
        mintType = 1
    Else
        opt卡类别.Item(1).Value = True
        mintType = 2
    End If
    mstrSNO = "": mstrIDNO = ""
    Me.txt确认新密码.MaxLength = mintLen
    Me.TXT新密码.MaxLength = mintLen
    Me.TXT原密码.MaxLength = mintLen
End Sub

Private Sub opt卡类别_Click(Index As Integer)
    If Index = 0 Then
        lbl卡号.Caption = "卡号"
        cmd读卡.Caption = "启动"
    Else
        lbl卡号.Caption = "IC卡"
        cmd读卡.Caption = "读卡"
    End If
    mintType = Index + 1
End Sub

Private Sub txt确认新密码_GotFocus()
    txt确认新密码.SelStart = 0
    txt确认新密码.SelLength = mintLen
End Sub

Private Sub txt确认新密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
End Sub

Private Sub txt新密码_GotFocus()
    TXT新密码.SelStart = 0
    TXT新密码.SelLength = mintLen
End Sub

Private Sub TXT新密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
    Call ReadCard(txt确认新密码, 1)
End Sub

Private Sub txt原密码_Change()
    cmdOK.Enabled = (Len(TXT原密码.Text) <> 0)
End Sub

Private Sub txt原密码_GotFocus()
    TXT原密码.SelStart = 0
    TXT原密码.SelLength = mintLen
End Sub

Public Function ChangePassword(ByVal strPass As String, Optional ByRef strOldPassWord As String = "", Optional ByVal intLen As Integer = 8) As String
    mintLen = intLen
    mstr新密码 = strPass
    Me.Show 1
    strOldPassWord = mstr旧密码
    ChangePassword = mstr新密码
    
End Function

Private Sub txt原密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mblnKey = False Then zlCommFun.PressKey (vbKeyTab)
    mblnKey = False
    Call ReadCard(TXT新密码, 1)
End Sub
