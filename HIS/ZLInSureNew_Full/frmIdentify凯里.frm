VERSION 5.00
Begin VB.Form frmIdentify凯里 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtCheckPass 
      Height          =   285
      Left            =   4695
      MaxLength       =   10
      TabIndex        =   36
      Top             =   225
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   34
      Top             =   240
      Width           =   1755
   End
   Begin VB.Frame fraInfo 
      Caption         =   "病人信息"
      Height          =   3255
      Left            =   68
      TabIndex        =   4
      Top             =   720
      Width           =   7035
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4170
         TabIndex        =   18
         Tag             =   "30"
         Top             =   2820
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4170
         TabIndex        =   17
         Tag             =   "20+22+24+26+28"
         Top             =   2385
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4170
         TabIndex        =   16
         Tag             =   "18"
         Top             =   1965
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4170
         TabIndex        =   15
         Tag             =   "11"
         Top             =   1530
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4170
         TabIndex        =   14
         Tag             =   "5"
         Top             =   1095
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4170
         TabIndex        =   13
         Tag             =   "3"
         Top             =   675
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4170
         TabIndex        =   12
         Tag             =   "1"
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1290
         TabIndex        =   11
         Tag             =   "19"
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1290
         TabIndex        =   10
         Tag             =   "21+23+25+27"
         Top             =   2385
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1290
         TabIndex        =   9
         Tag             =   "17"
         Top             =   1965
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1290
         TabIndex        =   8
         Tag             =   "6"
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1290
         TabIndex        =   7
         Tag             =   "4"
         Top             =   1095
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1290
         TabIndex        =   6
         Tag             =   "2"
         Top             =   675
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1290
         TabIndex        =   5
         Tag             =   "0"
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "个帐余额"
         Height          =   180
         Index           =   13
         Left            =   3375
         TabIndex        =   32
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "统筹报销累计"
         Height          =   180
         Index           =   11
         Left            =   3015
         TabIndex        =   31
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "个帐支出累计"
         Height          =   180
         Index           =   9
         Left            =   3015
         TabIndex        =   30
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "特殊病病种"
         Height          =   180
         Index           =   7
         Left            =   3195
         TabIndex        =   29
         Top             =   1605
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   4
         Left            =   3735
         TabIndex        =   28
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "人员类别"
         Height          =   180
         Index           =   3
         Left            =   3375
         TabIndex        =   27
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "社会保障号"
         Height          =   180
         Index           =   1
         Left            =   3195
         TabIndex        =   26
         Top             =   315
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "住院次数累计"
         Height          =   180
         Index           =   12
         Left            =   135
         TabIndex        =   25
         Top             =   2895
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "进入统筹累计"
         Height          =   180
         Index           =   10
         Left            =   135
         TabIndex        =   24
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "个帐增加累计"
         Height          =   180
         Index           =   8
         Left            =   135
         TabIndex        =   23
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "身份证号码"
         Height          =   180
         Index           =   6
         Left            =   315
         TabIndex        =   22
         Top             =   1605
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         Height          =   180
         Index           =   5
         Left            =   495
         TabIndex        =   21
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Index           =   2
         Left            =   855
         TabIndex        =   20
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC卡号"
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   19
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   400
      Left            =   83
      TabIndex        =   3
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   4853
      TabIndex        =   2
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5993
      TabIndex        =   1
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangPass 
      Caption         =   "改密码(&E)"
      Height          =   400
      Left            =   1238
      TabIndex        =   0
      Top             =   4065
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "确认密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   15
      Left            =   3615
      TabIndex        =   35
      Top             =   270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmIdentify凯里.frx":0000
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   14
      Left            =   795
      TabIndex        =   33
      Top             =   285
      Width           =   675
   End
End
Attribute VB_Name = "frmIdentify凯里"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPass As String
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    Me.Show vbModal
    gstr医保号 = txtInfo(1).Text
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrOther = "": mstrPatient = ""
    Me.Hide
End Sub

Private Sub cmdChangPass_Click()
    If txtCheckPass.Visible Then
        If txtCheckPass.Text = txtPassword.Text And txtPassword.Text <> "" Then
            glngReturn = ChangePass(mstrPass & "|" & txtCheckPass.Text)
            If glngReturn <> 0 Then
                MsgBox "修改密码失败", vbInformation, "修改密码"
            Else
                MsgBox "修改密码成功", vbInformation, "修改密码"
            End If
            cmdChangPass.Caption = "改密码(&E)"
            txtCheckPass.Visible = False
            txtPassword.Text = ""
            txtPassword.SetFocus
        ElseIf txtCheckPass.Text = txtPassword.Text Then
            MsgBox "输入的密码不能为空", vbInformation, "修改密码"
            txtPassword.SetFocus
        ElseIf txtCheckPass.Text <> txtPassword.Text Then
            MsgBox "两次输入的密码不同，请重新输入", vbInformation, "修改密码"
            txtPassword.SetFocus
        End If
    Else
        txtCheckPass.Text = ""
        txtCheckPass.Visible = True
        txtPassword.SetFocus
        cmdChangPass.Caption = "修改(&E)"
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cur支出累计 As Currency, cur增加累计 As Currency
    If txtInfo(0).Text = "" Then
        MsgBox "请先进行读卡操作", vbInformation, "身份验证"
        Exit Sub
    End If
    mstrOther = "": mstrPatient = ""

    mstrPatient = txtInfo(0).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & mstrPass & ";"                          '2 密码
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '5 出生日期
    mstrPatient = mstrPatient & txtInfo(6).Text & ";"                   '6 身份证
    mstrPatient = mstrPatient & ";"                                     '7 单位名称

    mstrOther = mstrOther & ";"                                         '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & txtInfo(13).Text & ";"                      '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & IIf(txtInfo(5).Text = "在职" Or txtInfo(5).Text = "在职二等乙", "1", "2") & ";"
    mstrOther = mstrOther & ";"                                         '14 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & txtInfo(8).Text & ";"                       '18 帐户增加累计
    mstrOther = mstrOther & txtInfo(9).Text & ";"                       '19 帐户支出累计
    mstrOther = mstrOther & txtInfo(10).Text & ";"                      '20 进入统筹累计
    mstrOther = mstrOther & txtInfo(11).Text & ";"                      '21 统筹报销累计
    mstrOther = mstrOther & txtInfo(12).Text & ";"                      '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & ";"                                         '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    Dim strPara() As String
    If txtPassword = "" Then
        MsgBox "请输入密码", vbInformation, "身份验证"
        txtPassword.SetFocus
    End If
    mstrPass = txtPassword
    gstrReturn = ""
    glngReturn = GetPersonInfo(mstrPass, gstrReturn)
    If glngReturn <> 0 Then
        mstrPass = ""
        txtPassword.SetFocus
        MsgBox "错误", vbInformation, "身份验证"
        Exit Sub
    End If
    strPara = Split(gstrReturn, "|")
    txtInfo(0).Text = strPara(0)
    txtInfo(1).Text = strPara(1)
    txtInfo(2).Text = strPara(2)
    txtInfo(3).Text = IIf(strPara(3) = "1", "男", "女")
    txtInfo(4).Text = strPara(4)
    '01.在职,02.退休,03.离休,04.老红军,05.在职二等乙,06.退休二等乙,07.离休二等乙
    Select Case strPara(5)
        Case "01"
            txtInfo(5).Text = "在职"
        Case "02"
            txtInfo(5).Text = "退休"
        Case "03"
            txtInfo(5).Text = "离休"
        Case "04"
            txtInfo(5).Text = "老红军"
        Case "05"
            txtInfo(5).Text = "在职二等乙"
        Case "06"
            txtInfo(5).Text = "退休二等乙"
        Case "07"
            txtInfo(5).Text = "离休二等乙"
    End Select
    txtInfo(6).Text = strPara(6)
    txtInfo(7).Text = strPara(10)
    txtInfo(8).Text = strPara(16)
    txtInfo(9).Text = strPara(17)
    txtInfo(10).Text = cNumber(strPara(19)) + cNumber(strPara(21)) + cNumber(strPara(23)) + cNumber(strPara(25)) + cNumber(strPara(27))
    txtInfo(11).Text = cNumber(strPara(20)) + cNumber(strPara(22)) + cNumber(strPara(24)) + cNumber(strPara(26)) + cNumber(strPara(28))
    txtInfo(12).Text = strPara(18)
    txtInfo(13).Text = strPara(29)
    txtPassword.Text = ""
    txtPassword.SetFocus
    cmdChangPass.Visible = True
End Sub

Private Sub txtCheckPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangPass_Click
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCheckPass.Visible = True Then
            txtCheckPass.SetFocus
        Else
            cmdRead_Click
        End If
    End If
End Sub
