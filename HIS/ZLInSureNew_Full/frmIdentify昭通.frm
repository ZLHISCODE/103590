VERSION 5.00
Begin VB.Form frmIdentify昭通 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRead 
      Caption         =   "刷卡(&R)"
      Height          =   400
      Left            =   3030
      TabIndex        =   17
      Top             =   2130
      Width           =   1100
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   1050
      TabIndex        =   15
      Top             =   1440
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1050
      TabIndex        =   8
      Top             =   560
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   4260
      TabIndex        =   7
      Top             =   560
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1050
      TabIndex        =   6
      Top             =   1000
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   4260
      TabIndex        =   5
      Top             =   1000
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -98
      TabIndex        =   4
      Top             =   1890
      Width           =   6810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   4140
      TabIndex        =   2
      Top             =   2130
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5245
      TabIndex        =   3
      Top             =   2130
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   16
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "卡    号"
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "医 保 号"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   650
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "姓    名"
      Height          =   180
      Index           =   3
      Left            =   3420
      TabIndex        =   12
      Top             =   650
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "性    别"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   11
      Top             =   1090
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   5
      Left            =   3420
      TabIndex        =   10
      Top             =   1090
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "密    码"
      Height          =   180
      Index           =   10
      Left            =   3420
      TabIndex        =   9
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify昭通"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte, blnEOF As Boolean
Public mstrPatient As String, mstrOther As String, mstr卡号 As String, mstr密码 As String

Public Function GetPatient(bytType As Byte) As String
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtInfo(2).Text = "" Or txtInfo(3).Text = "" Then Exit Sub
    mstrOther = "": mstrPatient = ""
    mstrPatient = txtInfo(0).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '2 密码
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & ";"                                     '5 出生日期
    mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 身份证
    mstrPatient = mstrPatient & ";"                                     '7 单位名称/编码
    
    mstrOther = mstrOther & ";"                                         '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & ";"                                         '14 在职状态
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & ";"                                         '20 进入统筹累计
    mstrOther = mstrOther & ";"                                         '21 统筹报销累计
    mstrOther = mstrOther & ";"                                         '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & ";"                                         '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    Dim intPort As Integer
    intPort = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", 1)
    Me.txtInfo(0).Text = frmConn昭通.readCard(intPort)
    If Me.txtInfo(0).Text <> "" Then
        Me.txtInfo(1).Text = frmConn昭通.readPassword(intPort)
        Call readInfo
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtInfo(0).SetFocus
    cmdOK.Enabled = False
End Sub

Private Sub Timer1_Timer()
    blnEOF = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = Len(txtInfo(Index).Text)
End Sub

Private Function readInfo() As Boolean
    Dim strPara As String, strReturn() As String
    strPara = txtInfo(0).Text & vbTab & IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text)
    If frmConn昭通.Execute("I200", 0, strPara, "正在读取病人医保信息......") = False Then
'        txtInfo(0).SetFocus
        Exit Function
    End If
    If frmConn昭通.Query(0, 1) = False Then Exit Function
    mstr卡号 = txtInfo(0).Text
    mstr密码 = IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text)
    cmdOK.Enabled = True
    strReturn = Split(Replace(frmConn昭通.strReturnInfo, " ", ""), vbTab)
    txtInfo(2).Text = strReturn(0)
    txtInfo(3).Text = strReturn(1)
    txtInfo(4).Text = IIf(strReturn(2) = 0, "女", "男")
    txtInfo(5).Text = strReturn(3)
    txtInfo(6).Text = strReturn(4)
    readInfo = True
    cmdOK.SetFocus
End Function
