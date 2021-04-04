VERSION 5.00
Begin VB.Form frmVerfyCodeInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "验证输入"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7230
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   450
      Left            =   5535
      TabIndex        =   8
      Top             =   3315
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   450
      Left            =   3960
      TabIndex        =   7
      Top             =   3315
      Width           =   1395
   End
   Begin VB.PictureBox picVerfy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   1785
      ScaleHeight     =   885
      ScaleWidth      =   4605
      TabIndex        =   5
      Top             =   1980
      Width           =   4635
   End
   Begin VB.TextBox txtEdit 
      Height          =   420
      Left            =   1770
      TabIndex        =   4
      Top             =   1440
      Width           =   4620
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   30
      TabIndex        =   2
      Top             =   3120
      Width           =   8010
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   -90
      TabIndex        =   1
      Top             =   1275
      Width           =   7845
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "换图"
      Height          =   855
      Left            =   930
      TabIndex        =   6
      Top             =   1995
      Width           =   750
   End
   Begin VB.Label cmdOptionNotes 
      Height          =   1005
      Left            =   90
      TabIndex        =   10
      Top             =   4005
      Width           =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   -15
      X2              =   7200
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Label lblInfor 
      Caption         =   "验证码输入错误"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   30
      TabIndex        =   9
      Top             =   3405
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Label lblInputMemo 
      AutoSize        =   -1  'True
      Caption         =   "验证码"
      Height          =   285
      Left            =   825
      TabIndex        =   3
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblNotes 
      Caption         =   "你输入的单据已经医保退费成功,但一卡通交易失败,为了保存数据的正确性,你必须继续调用结算交易!"
      Height          =   990
      Left            =   1170
      TabIndex        =   0
      Top             =   210
      Width           =   5700
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   210
      Picture         =   "frmVerfyCodeInput.frx":0000
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmVerfyCodeInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrVarData(0 To 36) As String
Private mblnOk As Boolean
Private mstrNotes As String
Private mstrCmdsCaption As String
Private mstrVerifyCode As String
Private Sub InitCommand()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化按钮
    '编制:刘兴洪
    '日期:2012-01-13 13:50:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    lblNotes.Caption = mstrNotes
    If mstrCmdsCaption = "" Then Exit Sub
    varData = Split(mstrCmdsCaption, ";")
    varTemp = Split(varData(0) & "|", "|")
    
    cmdOK.Caption = varTemp(0)
    cmdOK.ToolTipText = varTemp(1)
    cmdOK.Width = TextWidth(cmdOK.Caption) + 100
    varTemp = Split(varData(1) & "|", "|")
    cmdCancel.Caption = varTemp(0)
    cmdCancel.ToolTipText = varTemp(1)
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    cmdOptionNotes.Caption = IIf(cmdOK.ToolTipText = "", "", "[" & Split(cmdOK.Caption, "(")(0) & "]:" & cmdOK.ToolTipText)
    cmdOptionNotes.Caption = cmdOptionNotes.Caption & IIf(cmdCancel.ToolTipText = "", "", vbCrLf & "[" & Split(cmdCancel.Caption, "(")(0) & "]:" & cmdCancel.ToolTipText)
    If cmdOptionNotes.Caption = "" Then
        cmdOptionNotes.Visible = False: Line1.Visible = False: Me.Height = 4350
    Else
        Me.Height = 5535
    End If
End Sub
Private Sub InitVar()
    Dim i As Integer, j As Integer
    For i = Asc("a") To Asc("z")
        mstrVarData(j) = UCase(Chr(i))
        j = j + 1
    Next
    For i = Asc("0") To Asc("9")
        mstrVarData(j) = Chr(i)
        j = j + 1
    Next
End Sub
Private Sub SetVerifyCode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入验证码
    '编制:刘兴洪
    '日期:2012-01-13 11:59:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCode As String, strTmp As String
    Dim i As Long, j As Integer, sngX As Single, sngY As Single
    picVerfy.Cls
    picVerfy.CurrentX = 0
    picVerfy.CurrentY = 0
    sngX = 100
    sngY = 100
    Call Randomize
    Do While True
        i = Int(Rnd(Rnd * 100) * 100)
        If i < 36 And i >= 0 Then
            strTmp = mstrVarData(i)
            If j Mod 2 = 0 Then
                picVerfy.FontItalic = False
            Else
                picVerfy.FontItalic = True
            End If
            picVerfy.CurrentX = sngX
             picVerfy.CurrentY = sngY
            sngX = sngX + picVerfy.TextWidth(strTmp) + 100
            sngY = sngY + picVerfy.TextHeight(strTmp) \ 4
            picVerfy.Print strTmp
            strCode = strCode & strTmp
            j = j + 1
        End If
        If j > 3 Then Exit Do
    Loop
   picVerfy.Tag = strCode
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub cmdRead_Click()
    Call SetVerifyCode
    If txtEdit.Enabled Then txtEdit.SetFocus
End Sub

Private Sub cmdOK_Click()
    If picVerfy.Tag <> UCase(txtEdit.Text) Then
         lblInfor.Visible = True: Exit Sub
    End If
    mstrVerifyCode = picVerfy.Tag
    mblnOk = True: Unload Me
End Sub

Private Sub Form_Load()
    mblnOk = True
    Call InitVar
    Call SetVerifyCode
End Sub
Private Sub txtEdit_Change()
    lblInfor.Visible = False
End Sub

Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub
Public Function ShowMsg(ByVal frmMain As Object, ByVal strNotes As String, _
    ByVal strCmdsCaption As String, Optional strVerifyCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示Msgbox 信息
    '入参:frmMain-调用的主窗体
    '       strNotes-说明文字
    '       strCmdCaption-按钮文本,目前只有两个按钮,第一个为确定(按钮标题|按钮说明);第二个取消
    '
    '出参:strVerifyCode-返回输入的验证码
    '返回:确定或输入验证码成功后,返回ture,否则返回False
    '编制:刘兴洪
    '日期:2012-01-13 13:47:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False: mstrCmdsCaption = strCmdsCaption: mstrNotes = strNotes
    mstrVerifyCode = ""
    Call InitCommand
    Me.Show 1, frmMain
    strVerifyCode = mstrVerifyCode
    ShowMsg = mblnOk
End Function



