VERSION 5.00
Begin VB.Form frmIdentify冕宁 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmIdentify冕宁.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt单位名称 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   26
      Top             =   1470
      Width           =   4350
   End
   Begin VB.CommandButton cmd刷卡 
      Caption         =   "刷卡(&R)"
      Height          =   375
      Left            =   465
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txt起付线 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   24
      Top             =   2280
      Width           =   1185
   End
   Begin VB.TextBox txt人员状态 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   22
      Top             =   1065
      Width           =   1185
   End
   Begin VB.TextBox txt人员类别 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   20
      Top             =   1065
      Width           =   1185
   End
   Begin VB.TextBox txt帐户余额 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   18
      Top             =   2280
      Width           =   1185
   End
   Begin VB.TextBox txt网络状态 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4395
      TabIndex        =   16
      Top             =   255
      Width           =   1185
   End
   Begin VB.TextBox txt出生日期 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4395
      TabIndex        =   14
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txt消费类型 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   12
      Top             =   1875
      Width           =   1185
   End
   Begin VB.TextBox txt帐户状态 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1875
      Width           =   1185
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2940
      TabIndex        =   6
      Top             =   660
      Width           =   360
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1035
      TabIndex        =   5
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txt社保号 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1035
      TabIndex        =   3
      Top             =   255
      Width           =   2250
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4425
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   2985
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "起付线："
      Height          =   225
      Left            =   2580
      TabIndex        =   25
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label11 
      Caption         =   "人员状态："
      Height          =   225
      Left            =   2520
      TabIndex        =   23
      Top             =   1140
      Width           =   930
   End
   Begin VB.Label Label10 
      Caption         =   "人员类别："
      Height          =   225
      Left            =   300
      TabIndex        =   21
      Top             =   1140
      Width           =   930
   End
   Begin VB.Label Label9 
      Caption         =   "帐户余额："
      Height          =   225
      Left            =   300
      TabIndex        =   19
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label8 
      Caption         =   "网络状态："
      Height          =   255
      Left            =   3465
      TabIndex        =   17
      Top             =   270
      Width           =   930
   End
   Begin VB.Label Label7 
      Caption         =   "出生日期："
      Height          =   225
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label6 
      Caption         =   "消费类型："
      Height          =   225
      Left            =   2565
      TabIndex        =   13
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label5 
      Caption         =   "帐户状态："
      Height          =   225
      Left            =   300
      TabIndex        =   11
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "单位名称："
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "性别："
      Height          =   255
      Left            =   2355
      TabIndex        =   8
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "社保号："
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   780
   End
End
Attribute VB_Name = "frmIdentify冕宁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID As Long
Private mstrReturn As String

Function 身份标识(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = "99"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    身份标识 = mstrReturn
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmd刷卡_Click()
  Dim strSBH As String * 20 '社保号
 ' Dim strSbh1 As String '用于取得人员信息时传递的社保号,长度25位
  
  Dim net As String * 1 '网络状态
  Dim rylx As String * 1 '人员类型
  Dim zhzt As String * 3 '帐户状态
  Dim tmp As Double '返回值
  Dim Zhye As Double '帐户余额
  
  Dim fhz As String * 1000 '返回值

 Dim zffs As String * 1 '支付方式
 Dim tcqfx As Double '统筹起付线
 
net = Space(1)
strSBH = Space(20)
rylx = Space(1)
zhzt = Space(3)
zffs = Space(1)

If mbytType = 0 Then 'begin 门诊

Zhye = 0

'beging 1刷卡
  tmp = mzcsh(strSBH, net, rylx, zhzt, Zhye)
  Call WriteBusinessLOG("mzcsh", "sbh,net,rylx,zhzt,zhye", tmp & "," & strSBH & "," & net & "," & rylx & "," & zhzt & "," & Zhye)
  
  If tmp = 0 Then
    txt社保号.Text = Mid(strSBH, 1, 18)
    
    txt帐户状态.Tag = Mid(zhzt, 1, 3)
    If txt帐户状态.Tag = "002" Then
        txt帐户状态.Text = "帐户正常"
    Else
        txt帐户状态.Text = "帐户冻结"
    End If
    
    txt消费类型.Tag = Mid(rylx, 1, 1)
    Select Case txt消费类型.Tag
    Case 1
        txt消费类型.Text = "普通消费"
    Case 2
        txt消费类型.Text = "按月包干"
    Case 3
        txt消费类型.Text = "特殊消费"
    End Select
    
    txt网络状态.Tag = Mid(net, 1, 1)
    If txt网络状态.Tag = 1 Then
        txt网络状态.Text = "通"
    Else
        txt网络状态.Text = "不通"
    End If
    
    If Mid(rylx, 1, 1) <> 2 And Mid(zhzt, 1, 3) = "002" Then
        txt帐户余额.Text = Val(Zhye)
    Else
        txt帐户余额.Text = 0
    End If
    '2然后取得基本信息
    fhz = Space(1000)
    'strSbh1 = Trim(txt社保号.Text) & Space(25 - Len(Trim(txt社保号.Text)))
    tmp = getyhxx_vb(txt社保号.Text, fhz)
    Call WriteBusinessLOG("getyhxx", txt社保号.Text & "," & Trim(fhz), tmp)

    If Trim(fhz) <> "" Then
        txt姓名.Text = Split(fhz, ",")(0)
        txt性别.Text = Split(fhz, ",")(1)
        txt单位名称.Text = Split(fhz, ",")(2)
        txt人员类别.Text = Split(fhz, ",")(3)
        txt人员状态.Text = Split(fhz, ",")(4)
        txt出生日期.Text = Split(fhz, ",")(5)
    End If
    OKButton.Enabled = True
    SendKeys ("{Tab}")
  End If
  'end 刷卡
  If tmp = 1 Then MsgBox "个人信息无该社保号", vbInformation, gstrSysName
  If tmp = 2 Then MsgBox "本地基本信息库需要更新", vbInformation, gstrSysName
  If tmp = 99 Then MsgBox "错误", vbInformation, gstrSysName

End If 'end 门诊

If mbytType = 1 Then 'beging 住院
    tmp = rycsh(strSBH, zhzt, zffs, net, Zhye, tcqfx)
    Call WriteBusinessLOG("zycsh", "sbh, zhzt, zffs, net, Zhye, tcqfx", tmp & "," & strSBH & "," & zhzt & "," & zffs & "," & net & "," & Zhye & "," & tcqfx)
    If tmp = 0 Then
        txt社保号.Text = Mid(strSBH, 1, 18)
        
        txt帐户状态.Tag = Mid(zhzt, 1, 3)
        If txt帐户状态.Tag = "002" Then
            txt帐户状态.Text = "帐户正常"
        Else
            txt帐户状态.Text = "帐户冻结"
        End If
            
        txt消费类型.Tag = Mid(zffs, 1, 1)
        If txt消费类型.Tag = 1 Then
            txt消费类型.Text = "帐户不可用"
        Else
            txt消费类型.Text = "先用帐户"
        End If
        
        txt网络状态.Tag = Mid(net, 1, 1)
        If txt网络状态.Tag = 1 Then
            txt网络状态.Text = "通"
        Else
            txt网络状态.Text = "不通"
        End If
        
        txt起付线.Text = Val(tcqfx)
        
        If txt消费类型.Tag <> 1 And txt帐户状态.Tag = "002" Then
            txt帐户余额.Text = Val(Zhye)
        Else
            txt帐户余额.Text = 0
        End If
        
        '2然后取得基本信息
        fhz = Space(1000)
        'strSbh1 = Trim(txt社保号.Text) & Space(25 - Len(Trim(txt社保号.Text)))
        tmp = getyhxx_vb(Trim(txt社保号.Text), fhz)
        Call WriteBusinessLOG("getyhxx", Trim(txt社保号.Text) & "," & Trim(fhz), tmp)
        
        If Trim(fhz) <> "" Then
            txt姓名.Text = Split(fhz, ",")(0)
            txt性别.Text = Split(fhz, ",")(1)
            txt单位名称.Text = Split(fhz, ",")(2)
            txt人员类别.Text = Split(fhz, ",")(3)
            txt人员状态.Text = Split(fhz, ",")(4)
            txt出生日期.Text = Split(fhz, ",")(5)
        End If
        OKButton.Enabled = True
        SendKeys ("{Tab}")
    End If
  'end 刷卡
    If tmp = 1 Then MsgBox "个人信息无该社保号", vbInformation, gstrSysName
    If tmp = 2 Then MsgBox "本地基本信息库需要更新", vbInformation, gstrSysName
    If tmp = 99 Then MsgBox "错误", vbInformation, gstrSysName
End If 'end 住院
End Sub

Private Sub Form_Load()
    '初始化控件
    If mbytType = 0 Then
        txt起付线.Visible = False
        Label12.Visible = False
        Label6.Caption = "消费类型："
    End If
    
    If mbytType = 1 Then
        txt起付线.Visible = True
        Label12.Visible = True
        Label6.Caption = "支付方式："
    End If
    OKButton.Enabled = False
End Sub

Private Sub OKButton_Click()
    Dim strEmpInfo As String
    Dim straccinfo As String

    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    
    strEmpInfo = txt社保号.Text                               '0卡号
    strEmpInfo = strEmpInfo & ";" & txt社保号.Text              '1医保号
    strEmpInfo = strEmpInfo & ";" & txt网络状态.Tag             '2密码  本医保中保存的是网络状态
    strEmpInfo = strEmpInfo & ";" & txt姓名.Text                '3姓名
    strEmpInfo = strEmpInfo & ";" & txt性别.Text                '4性别
    strEmpInfo = strEmpInfo & ";" & txt出生日期.Text         '5出生日期
    strEmpInfo = strEmpInfo & ";"             '6身份证
    strEmpInfo = strEmpInfo & ";" & txt单位名称          '7.单位名称(编码)
    
    straccinfo = ";0"                                          '8.中心代码
    straccinfo = straccinfo & ";"                    '9.顺序号
    straccinfo = straccinfo & ";" & txt消费类型.Tag             '10人员身份
    straccinfo = straccinfo & ";" & Val(txt帐户余额.Text)        '11帐户余额
    straccinfo = straccinfo & ";" & txt人员状态.Tag   ' & g个人基本信息.在院状态16                             '12当前状态
    straccinfo = straccinfo & ";"                   '13病种ID
    straccinfo = straccinfo & ";1"                            '14在职(1,2,3)
    straccinfo = straccinfo & ";"                             '15退休证号
    straccinfo = straccinfo & ";"                             '16年龄段
    straccinfo = straccinfo & ";1"                            '17灰度级
    straccinfo = straccinfo & ";0"       '18帐户增加累计
    straccinfo = straccinfo & ";0"                              '19帐户支出累计
    straccinfo = straccinfo & ";0"                            '20上年工资总额
    straccinfo = straccinfo & ";"      '21
    straccinfo = straccinfo & ";"       '22住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strEmpInfo & straccinfo, mlng病人ID, TYPE_冕宁)
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_冕宁 & ",'就诊类别','''" & txt消费类型.Tag & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "应诊类别")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_冕宁 & ",'人员身份','''" & txt帐户状态.Tag & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "帐户状态")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_冕宁 & ",'帐户余额','''" & txt帐户余额.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "帐户余额")
    
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng病人ID & straccinfo
    End If
    Unload Me

End Sub

