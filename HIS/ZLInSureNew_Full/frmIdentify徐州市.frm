VERSION 5.00
Begin VB.Form frmIdentify徐州市 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   4298
      TabIndex        =   13
      Top             =   225
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1208
      TabIndex        =   12
      Top             =   660
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   4298
      TabIndex        =   11
      Top             =   660
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1208
      TabIndex        =   10
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   4298
      TabIndex        =   9
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   1208
      TabIndex        =   8
      Top             =   1485
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   4298
      TabIndex        =   7
      Top             =   1485
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   1208
      TabIndex        =   6
      Top             =   1905
      Width           =   5175
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   1208
      TabIndex        =   5
      Top             =   2310
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   2790
      Width           =   6810
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   400
      Left            =   428
      TabIndex        =   3
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   4178
      TabIndex        =   2
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5283
      TabIndex        =   1
      Top             =   3000
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1208
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   225
      Width           =   2085
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   3540
      TabIndex        =   23
      Top             =   315
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "工作状态"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   22
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "姓    名"
      Height          =   180
      Index           =   3
      Left            =   3540
      TabIndex        =   21
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "性    别"
      Height          =   180
      Index           =   4
      Left            =   450
      TabIndex        =   20
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   5
      Left            =   3540
      TabIndex        =   19
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "联系电话"
      Height          =   180
      Index           =   6
      Left            =   450
      TabIndex        =   18
      Top             =   1575
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "邮政编码"
      Height          =   180
      Index           =   7
      Left            =   3540
      TabIndex        =   17
      Top             =   1575
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "家庭住址"
      Height          =   180
      Index           =   8
      Left            =   450
      TabIndex        =   16
      Top             =   1995
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "工作单位"
      Height          =   180
      Index           =   9
      Left            =   450
      TabIndex        =   15
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "IC卡密码"
      Height          =   180
      Index           =   10
      Left            =   450
      TabIndex        =   14
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify徐州市"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytType As Byte, mstrPatient As String, mstrOther As String, mint住院次数 As Integer
Private strTransNO As String, cur支出累计 As Currency, cur增加累计 As Currency, strPara As String, _
    strReturn As String, blnReadCard As Boolean
 
Public Function GetPatient(bytType As Byte) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '17-门诊起付线支付，18-住院起付线支付，19-本年住院次数，20-门诊费用，21-住院费用，22-帐户余额
    '23-参加统筹支付费用，24-统筹支付费用，25-参加大病支付费用，26-大病支付费用，27-是否特殊参保病人
    '28-参保年限，29-医保状态(0正常)
    Dim datCurr As Date
    If blnReadCard = False Then
        MsgBox "请先读卡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    If UCase(txtInfo(4).Text) = "YYYY-MM-DD" Then
        txtInfo(4).Enabled = True
        MsgBox "请输入正确的出生日期", vbInformation, gstrSysName
        txtInfo(4).SetFocus
        txtInfo(4).SelStart = 0
        txtInfo(4).SelLength = Len(txtInfo(4).Text)
        On Error GoTo 0
        Exit Sub
    Else
        datCurr = CDate(txtInfo(4).Text)
        If Err.Number <> 0 Then
            MsgBox "请按格式:yyyy-mm-dd输入正确的出生日期", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    mstrOther = "": mstrPatient = ""
    strReturn = Me.Tag
    mstrPatient = txtInfo(0).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & txtPass.Text & ";"                      '2 密码
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '5 出生日期
    mstrPatient = mstrPatient & ";"                                     '6 身份证
    mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '7 单位名称/编码
        
    mstrOther = mstrOther & ";"                                         '8 医保机构编码(中心)
    mstrOther = mstrOther & txtInfo(0).Tag & ";"                        '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & Split(strReturn, ",")(22) & ";"             '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & IIf(txtInfo(1).Text = "在职", "1", "3") & ";"
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & Split(strReturn, ",")(22) & ";"             '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & Split(strReturn, ",")(23) & ";"             '20 进入统筹累计
    mstrOther = mstrOther & Split(strReturn, ",")(24) & ";"             '21 统筹报销累计
    mstrOther = mstrOther & Split(strReturn, ",")(19) & ";"             '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & Split(strReturn, ",")(18) & ";"             '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    
    mint住院次数 = CInt(Split(strReturn, ",")(19))
    
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    Dim lngReturn As Long, strReturn As String, strErrInfo As String, strInfo() As String
    If Trim(txtPass.Text) = "" Then
        MsgBox "请输入IC卡密码", vbInformation, "读卡"
        Exit Sub
    End If
    lngReturn = init_com(intCOM徐州)
    If lngReturn <> 0 Then
        MsgBox "初始化端口错误", vbInformation, "读卡"
        Exit Sub
    End If
    
    lngReturn = sele_card(43)
    If lngReturn <> 0 Then
        MsgBox "定义卡类型错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "卡上电错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
    If lngReturn <> 0 Then
        MsgBox "读取卡信息错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    On Error GoTo powerOFF
    strInfo = Split(Trim(strReturn), "@")
    txtInfo(0).Text = strInfo(2)
    For lngReturn = 1 To 8
        If InStr(strInfo(lngReturn + 3), Chr(0)) > 0 Then
            strInfo(lngReturn + 3) = Left(strInfo(lngReturn + 3), InStr(strInfo(lngReturn + 3), Chr(0)) - 1)
        End If
        txtInfo(lngReturn).Text = IIf(lngReturn <> 3, IIf(lngReturn <> 1, strInfo(lngReturn + 3), IIf(strInfo(lngReturn + 3) = "0", "退休", "在职")), IIf(strInfo(lngReturn + 3) = "0", "男", "女"))
    Next
    txtInfo(0).Tag = strInfo(0)
    
    blnReadCard = True
    cmdOK.SetFocus

powerOFF:
    Call power_off
    Call close_com
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtPass.Text) = "" Then
            txtPass_GotFocus
            Exit Sub
        End If
        cmdRead_Click
        If blnReadCard = True Then cmdOK.SetFocus
    End If
End Sub

Private Sub clsText()
    Dim iLoop As Long
    For iLoop = 0 To 8
        txtInfo(iLoop).Text = ""
    Next
End Sub


