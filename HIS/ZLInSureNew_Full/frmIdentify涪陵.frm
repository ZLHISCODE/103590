VERSION 5.00
Begin VB.Form frmIdentify涪陵 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份标识"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdChangPass 
      Caption         =   "改密码(&E)"
      Height          =   400
      Left            =   1230
      TabIndex        =   37
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   6255
      TabIndex        =   17
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   5115
      TabIndex        =   16
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   400
      Left            =   75
      TabIndex        =   15
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Caption         =   "病人信息"
      Height          =   3675
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      Begin VB.CommandButton CmdSel 
         Caption         =   "…"
         Height          =   300
         Left            =   6930
         TabIndex        =   36
         Top             =   3255
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   15
         Left            =   4455
         TabIndex        =   35
         Top             =   3255
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1635
         TabIndex        =   25
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1635
         TabIndex        =   24
         Top             =   675
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1635
         TabIndex        =   23
         Top             =   1095
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1635
         TabIndex        =   22
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1635
         TabIndex        =   21
         Top             =   1965
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1635
         TabIndex        =   20
         Top             =   2385
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1635
         TabIndex        =   19
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   14
         Left            =   1635
         TabIndex        =   18
         Top             =   3255
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4455
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4455
         TabIndex        =   6
         Top             =   675
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4455
         TabIndex        =   5
         Top             =   1095
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4455
         TabIndex        =   4
         Top             =   1530
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4455
         TabIndex        =   3
         Top             =   1965
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4455
         TabIndex        =   2
         Top             =   2385
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4995
         TabIndex        =   1
         Top             =   2820
         Width           =   2235
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "病种选择"
         Height          =   180
         Index           =   15
         Left            =   3660
         TabIndex        =   34
         Top             =   3330
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "医保机构编码"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   33
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC卡号"
         Height          =   180
         Index           =   2
         Left            =   1020
         TabIndex        =   32
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   5
         Left            =   1200
         TabIndex        =   31
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "单位编码"
         Height          =   180
         Index           =   6
         Left            =   840
         TabIndex        =   30
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         Height          =   180
         Index           =   8
         Left            =   840
         TabIndex        =   29
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC卡余额"
         Height          =   180
         Index           =   10
         Left            =   840
         TabIndex        =   28
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "基本医疗最高限额"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   2895
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "大病医疗累计支出"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   26
         Top             =   3330
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "个人帐号"
         Height          =   180
         Index           =   1
         Left            =   3660
         TabIndex        =   14
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "身份证号码"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   13
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Index           =   4
         Left            =   4020
         TabIndex        =   12
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "单位名称"
         Height          =   180
         Index           =   7
         Left            =   3660
         TabIndex        =   11
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "人员类别"
         Height          =   180
         Index           =   9
         Left            =   3660
         TabIndex        =   10
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "起付标准"
         Height          =   180
         Index           =   11
         Left            =   3660
         TabIndex        =   9
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "基本医疗累计支出"
         Height          =   180
         Index           =   13
         Left            =   3480
         TabIndex        =   8
         Top             =   2895
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmIdentify涪陵"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean, mbytType As Byte
Public mstrPatient As String, mstrOther As String, mstr就诊编号 As String, mstr就诊次数 As String, mstr起付标准 As String
 
Public Function GetPatient(bytType As Byte, str就诊次数 As String) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    mstr就诊次数 = str就诊次数
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdChangPass_Click()
    initType
    mblnReturn = fl_changePassword(gstr医保机构编码, gstr医院编码, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
    Else
        MsgBox "密码修改成功", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cur支出累计 As Currency, cur增加累计 As Currency
    With gstrOutPara
        '支出累计 = 基本医疗累计支出 + 大病医疗累计支出
        cur支出累计 = IIf(txtInfo(13).Text = "", 0, CCur(txtInfo(13).Text)) + IIf(txtInfo(14).Text = "", 0, CCur(txtInfo(14).Text))
        '增加累计 = IC卡余额 + 支出累计
        cur增加累计 = IIf(txtInfo(10).Text = "", 0, CCur(txtInfo(10).Text)) + cur支出累计
        mstrOther = "": mstrPatient = ""
        
        mstrPatient = txtInfo(2).Text & ";"                                 '0 卡号
        mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '1 医保帐号
        mstrPatient = mstrPatient & ";"                                     '2 密码
        mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 姓名
        mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 性别
        mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '5 出生日期
        mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 身份证
        mstrPatient = mstrPatient & txtInfo(7).Text & "(" & txtInfo(6).Text & ");"                   '7 单位名称/编码
        
        mstrOther = mstrOther & txtInfo(0).Text & ";"                       '8 医保机构编码(中心)
        mstrOther = mstrOther & ";"                                         '9 顺序号
        mstrOther = mstrOther & ";"                                         '10 身份
        mstrOther = mstrOther & txtInfo(10).Text & ";"                      '11 余额
        mstrOther = mstrOther & ";"                                         '12 当前状态
        mstrOther = mstrOther & ";"                                         '13 病种ID
        mstrOther = mstrOther & IIf(txtInfo(9).Text = "在职", "1", IIf(txtInfo(9).Text = "退休", "2", "3")) & ";"
        mstrOther = mstrOther & CLng(mstr就诊次数) + 1 & ";"                '14 退休证号
        mstrOther = mstrOther & ";"                                         '16 年龄段
        mstrOther = mstrOther & ";"                                         '17 灰度级
        mstrOther = mstrOther & cur增加累计 & ";"                           '18 帐户增加累计
        mstrOther = mstrOther & cur支出累计 & ";"                           '19 帐户支出累计
        mstrOther = mstrOther & ";"                                         '20 进入统筹累计
        mstrOther = mstrOther & ";"                                         '21 统筹报销累计
        mstrOther = mstrOther & ";"                                         '22 住院次数累计
        mstrOther = mstrOther & ";"                                         '23 就诊类别
        mstrOther = mstrOther & txtInfo(11).Text & ";"                      '24 本次起付线
        mstrOther = mstrOther & ";"                                         '25 起付线累计
        mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    End With
    initType
    If mbytType = 0 Then      '若是门诊或住院则获取就诊编号
'    mblnReturn = fl_dall(gstr医保机构编码, gstr医院编码, "2003121000031", gstrOutPara)
        mblnReturn = fl_reg(gstr医保机构编码, gstr医院编码, 0, UserInfo.姓名, Format(zlDatabase.Currentdate, "yyyy-MM-dd"), "0", gstrOutPara)
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
            Exit Sub
        End If
        TrimType
        mstr就诊编号 = gstrOutPara.out1
    End If
    mstr起付标准 = txtInfo(11).Text
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    cmdOK.SetFocus
    initType
    mblnReturn = fl_getybjgbm(gstrOutPara)
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
        Exit Sub
    End If
    gstr医保机构编码 = Trim(gstrOutPara.out1)
    initType
    mblnReturn = fl_readicxx(gstr医保机构编码, gstr医院编码, "0", gstrOutPara)
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
        Exit Sub
    End If
    TrimType
    With gstrOutPara
        txtInfo(0).Text = .out1
        txtInfo(1).Text = .out2
        txtInfo(2).Text = .out3
        txtInfo(3).Text = .out5
        txtInfo(4).Text = IIf(.out6 = "0", "男", "女")
        txtInfo(5).Text = .out4
        txtInfo(6).Text = .out7
        txtInfo(7).Text = .out8
        txtInfo(8).Text = .out9
        'Modified by zyb 2004-10-09
        'txtInfo(9).Text = IIf(.out10 = "21", "在职", IIf(.out10 = "22", "退休", "下岗"))
        txtInfo(9).Text = IIf(.out10 = "11", "在职", IIf(.out10 = "21", "退休", "下岗"))
        txtInfo(10).Text = .out11
        txtInfo(11).Text = .out12
        txtInfo(12).Text = .out13
        txtInfo(13).Text = .out14
        txtInfo(14).Text = .out15
    End With
End Sub

