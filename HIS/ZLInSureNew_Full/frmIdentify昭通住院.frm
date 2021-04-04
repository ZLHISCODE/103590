VERSION 5.00
Begin VB.Form frmIdentify昭通住院 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   15
      Left            =   4230
      TabIndex        =   36
      Top             =   3480
      Width           =   2085
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "刷卡(&R)"
      Height          =   400
      Left            =   3030
      TabIndex        =   35
      Top             =   4260
      Width           =   1100
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   14
      Left            =   4237
      TabIndex        =   31
      Top             =   3060
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   13
      Left            =   1027
      TabIndex        =   29
      Top             =   3135
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   12
      Left            =   4237
      TabIndex        =   27
      Top             =   2640
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   11
      Left            =   1027
      TabIndex        =   25
      Top             =   2700
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   4237
      TabIndex        =   23
      Top             =   2220
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   9
      Left            =   1027
      TabIndex        =   21
      Top             =   2265
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   4237
      TabIndex        =   19
      Top             =   1800
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   1027
      TabIndex        =   17
      Top             =   1830
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   1027
      TabIndex        =   15
      Top             =   3570
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1027
      TabIndex        =   0
      Top             =   540
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1027
      TabIndex        =   8
      Top             =   975
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   4237
      TabIndex        =   7
      Top             =   960
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1027
      TabIndex        =   6
      Top             =   1410
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   4237
      TabIndex        =   5
      Top             =   1380
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4237
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -98
      TabIndex        =   4
      Top             =   4020
      Width           =   6810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   4147
      TabIndex        =   2
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5257
      TabIndex        =   3
      Top             =   4260
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "缴费时期"
      Height          =   180
      Index           =   16
      Left            =   3390
      TabIndex        =   37
      Top             =   3570
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "入院状态"
      Height          =   180
      Index           =   15
      Left            =   180
      TabIndex        =   33
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "机构名称"
      Height          =   180
      Index           =   14
      Left            =   3397
      TabIndex        =   32
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "机构代码"
      Height          =   180
      Index           =   13
      Left            =   187
      TabIndex        =   30
      Top             =   3225
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   12
      Left            =   3397
      TabIndex        =   28
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "单位代码"
      Height          =   180
      Index           =   11
      Left            =   187
      TabIndex        =   26
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "起 付 金"
      Height          =   180
      Index           =   9
      Left            =   3397
      TabIndex        =   24
      Top             =   2310
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "住院次数"
      Height          =   180
      Index           =   8
      Left            =   187
      TabIndex        =   22
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "工作状态"
      Height          =   180
      Index           =   7
      Left            =   3397
      TabIndex        =   20
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "年    龄"
      Height          =   180
      Index           =   6
      Left            =   187
      TabIndex        =   18
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   2
      Left            =   187
      TabIndex        =   16
      Top             =   3660
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "卡    号"
      Height          =   180
      Index           =   1
      Left            =   187
      TabIndex        =   14
      Top             =   630
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "医 保 号"
      Height          =   180
      Index           =   0
      Left            =   187
      TabIndex        =   13
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "姓    名"
      Height          =   180
      Index           =   3
      Left            =   3397
      TabIndex        =   12
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "性    别"
      Height          =   180
      Index           =   4
      Left            =   187
      TabIndex        =   11
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   5
      Left            =   3397
      TabIndex        =   10
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "密    码"
      Height          =   180
      Index           =   10
      Left            =   3397
      TabIndex        =   9
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify昭通住院"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte, blnEOF As Boolean
Public mstrPatient As String, mstrOther As String, mstrState As String

Public Function GetPatient(bytType As Byte, strinState As String) As String
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    Combo1.AddItem "普通住院"
    Combo1.AddItem "转院"
    Combo1.AddItem "门诊抢救"
    Combo1.ListIndex = 0
    Me.Show vbModal
    strinState = mstrState
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mstrOther = "": mstrPatient = ""
    If txtInfo(2).Text = "" Or txtInfo(3).Text = "" Then Exit Sub
    mstrState = Combo1.ListIndex + 1
    mstrPatient = txtInfo(0).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '2 密码
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & txtInfo(0).Tag & ";"                    '5 出生日期
    mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 身份证
    mstrPatient = mstrPatient & txtInfo(11).Text & ";"                  '7 单位名称/编码
    
    mstrOther = mstrOther & txtInfo(13).Text & ";"                      '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & txtInfo(8).Text & ";"                       '14 在职状态
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & ";"                                         '20 进入统筹累计
    mstrOther = mstrOther & ";"                                         '21 统筹报销累计
    mstrOther = mstrOther & ";"                                         '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & txtInfo(10).Text & ";"                      '24 本次起付线
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
    strPara = txtInfo(0).Text & vbTab & IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text) & vbTab & Combo1.ListIndex + 1
    If frmConn昭通.Execute("I300", 0, strPara, "正在读取病人医保信息......") = False Then
'        txtInfo(0).SetFocus
        Exit Function
    End If
    If frmConn昭通.Query(0, 1) = False Then Exit Function
    cmdOK.Enabled = True
    strReturn = Split(Replace(frmConn昭通.strReturnInfo, " ", ""), vbTab)
    txtInfo(2).Text = strReturn(0)
    txtInfo(3).Text = strReturn(1)
    txtInfo(4).Text = IIf(strReturn(3) = 0, "女", "男")
    txtInfo(5).Text = strReturn(2)
    txtInfo(6).Text = strReturn(12)
    txtInfo(7).Text = strReturn(4)
    txtInfo(8).Text = IIf(strReturn(5) = 1, "在职", "退休")
    txtInfo(9).Text = strReturn(6)
    txtInfo(10).Text = strReturn(7)
    txtInfo(11).Text = strReturn(8)
    txtInfo(12).Text = strReturn(9)
    txtInfo(13).Text = strReturn(10)
    txtInfo(14).Text = strReturn(11)
    txtInfo(0).Tag = Left(strReturn(13), 4) & "-" & Mid(strReturn(13), 5, 2) & "-" & Right(strReturn(13), 2)
    txtInfo(15).Text = strReturn(14)
    cmdOK.SetFocus
    readInfo = True
End Function
