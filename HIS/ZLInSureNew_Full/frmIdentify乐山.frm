VERSION 5.00
Begin VB.Form frmIdentify乐山 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frmIdentify乐山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Cmb中心 
      Height          =   300
      ItemData        =   "frmIdentify乐山.frx":000C
      Left            =   3240
      List            =   "frmIdentify乐山.frx":0034
      TabIndex        =   44
      Text            =   "市本级"
      Top             =   4650
      Width           =   1500
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   350
      Left            =   210
      TabIndex        =   41
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6720
      TabIndex        =   43
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5490
      TabIndex        =   42
      Top             =   4620
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      Enabled         =   0   'False
      Height          =   4335
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   19
         Left            =   5310
         TabIndex        =   40
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   18
         Left            =   5310
         TabIndex        =   38
         Top             =   3450
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   17
         Left            =   5310
         TabIndex        =   36
         Top             =   3060
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   16
         Left            =   5310
         TabIndex        =   34
         Top             =   2670
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   15
         Left            =   5310
         TabIndex        =   32
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   14
         Left            =   5310
         TabIndex        =   30
         Top             =   1890
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   13
         Left            =   5310
         TabIndex        =   28
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   5310
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   11
         Left            =   5310
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   10
         Left            =   5310
         TabIndex        =   22
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   9
         Left            =   1500
         TabIndex        =   20
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   8
         Left            =   1500
         TabIndex        =   18
         Top             =   3450
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   7
         Left            =   1500
         TabIndex        =   16
         Top             =   3060
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   6
         Left            =   1500
         TabIndex        =   14
         Top             =   2670
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   5
         Left            =   1500
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   4
         Left            =   1500
         TabIndex        =   10
         Top             =   1890
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   3
         Left            =   1500
         TabIndex        =   8
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   2
         Left            =   1500
         TabIndex        =   6
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   1
         Left            =   1500
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   0
         Left            =   1500
         TabIndex        =   2
         Top             =   330
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   4530
         TabIndex        =   39
         Top             =   3900
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "账户余额"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   18
         Left            =   4530
         TabIndex        =   37
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "补充住院累计"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   17
         Left            =   4170
         TabIndex        =   35
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "补充高额累计"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   4170
         TabIndex        =   33
         Top             =   2730
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付累计"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   15
         Left            =   4170
         TabIndex        =   31
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "门诊统筹支付额"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   3990
         TabIndex        =   29
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "门诊特病支付额"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   13
         Left            =   3990
         TabIndex        =   27
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "账户口令"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   12
         Left            =   4530
         TabIndex        =   25
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "账户状态"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   11
         Left            =   4530
         TabIndex        =   23
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "补充医保名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   10
         Left            =   4170
         TabIndex        =   21
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员类别"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   19
         Top             =   3900
         Width           =   900
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   720
         TabIndex        =   17
         Top             =   3510
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   7
         Left            =   720
         TabIndex        =   15
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员类别"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   13
         Top             =   2730
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员状态"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   720
         TabIndex        =   11
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   9
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参保ID号"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Label Lbl中心 
      AutoSize        =   -1  'True
      Caption         =   "请选择地区："
      Height          =   180
      Left            =   2040
      TabIndex        =   45
      Top             =   4725
      Width           =   1080
   End
End
Attribute VB_Name = "frmIdentify乐山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String
Private mbytType As Byte
Private mlng病人ID As Long

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0) As String
    mstrReturn = ""
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Const int医保号 As Integer = 0
    Const int卡号 As Integer = 1
    Const int姓名 As Integer = 2
    Const int性别 As Integer = 3
    Const int身份证号 As Integer = 4
    Const int单位编码 As Integer = 7
    Const int单位名称 As Integer = 8
    Const int密码 As Integer = 12
    Const int帐户余额 As Integer = 18
    Const int住院次数 As Integer = 19
    Dim str出生日期 As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str月份 As String, str日期 As String
    
    If Trim(txtInfo(int医保号).Text) = "" Then
        MsgBox "未得到参保病人的医保ID号，无法继续！", vbInformation, gstrSysName
        cmdRead.SetFocus
        Exit Sub
    End If
    
    '曾明春(2005-10-08)  检查病人状态,由于乐山市医保不同区县的参保ID可能相同，所以根据参保ID和所选择地区来进行唯一判断
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2] And 中心=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gintInsure, CStr(txtInfo(int医保号).Text), CInt(Cmb中心.ListIndex))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("状态") > 0 Then
            MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '做准备工作
    If Len(txtInfo(int身份证号).Text) = 15 Then
        str出生日期 = Mid(txtInfo(int身份证号).Text, 7, 6)
        If Mid(str出生日期, 1, 2) <= "05" Then
            str出生日期 = "20" & str出生日期
        Else
            str出生日期 = "19" & str出生日期
        End If
    Else
        str出生日期 = Mid(txtInfo(int身份证号).Text, 7, 8)
    End If
    
    If Mid(str出生日期, 5, 2) < 1 Or Mid(str出生日期, 5, 2) > 12 Then
       str月份 = Frm乐山_提示.出生日期更改_乐山(1, Mid(str出生日期, 5, 2))
    Else
       str月份 = Mid(str出生日期, 5, 2)
    End If
    If Mid(str出生日期, 7, 2) < 1 Or Mid(str出生日期, 7, 2) > 31 Then
       str日期 = Frm乐山_提示.出生日期更改_乐山(2, Mid(str出生日期, 7, 2))
    Else
       str日期 = Mid(str出生日期, 7, 2)
    End If
    If Mid(str出生日期, 5, 2) = 2 And Mid(str出生日期, 7, 2) > 28 Then
       str月份 = "2"
       str日期 = "28"
    End If
    str出生日期 = Mid(str出生日期, 1, 4) & "-" & str月份 & "-" & str日期
    
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = txtInfo(int卡号).Text                         '0卡号
    strIdentify = strIdentify & ";" & txtInfo(int医保号).Text   '1医保号（个人编号）
    strIdentify = strIdentify & ";" & txtInfo(int密码).Text     '2密码
    strIdentify = strIdentify & ";" & txtInfo(int姓名).Text     '3姓名
    strIdentify = strIdentify & ";" & txtInfo(int性别).Text     '4性别
    strIdentify = strIdentify & ";" & str出生日期               '5出生日期
    strIdentify = strIdentify & ";" & txtInfo(int身份证号).Text '6身份证
    strIdentify = strIdentify & ";" & txtInfo(int单位名称).Text & "(" & txtInfo(int单位编码).Text & ")"          '7.单位名称(编码)
    '曾明春(2005-10-08)修改
    strAddition = ";" & Cmb中心.ListIndex                       '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";"                             '10人员身份
    strAddition = strAddition & ";" & Val(txtInfo(int帐户余额).Text)      '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";"                             '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txtInfo(int帐户余额).Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";" & Val(txtInfo(int住院次数).Text)  '21住院次数累计
    
    Call DebugTool(strIdentify & strAddition)
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_乐山)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    gbytReturn_乐山 = LS_GetPersonInfo(gPersonInfo_乐山)
    If GetErrInfo_乐山 Then Exit Sub
    
    With gPersonInfo_乐山
        txtInfo(0).Text = .PSN_ID              ' As Integer      '医疗参保ID号
'        txtInfo(1).Text = .PSN_No             ' As Integer      '参保人编码
        txtInfo(2).Text = .PSN_NAME            ' As String * 100 '参保人姓名
        txtInfo(3).Text = .Sex                 ' As String * 100 '性别
        txtInfo(4).Text = .IDCARD              ' As String * 100 '身份证号码
        txtInfo(5).Text = .PSN_STS             ' As String * 100 '参保人状态
        txtInfo(6).Text = .PSN_TYP             ' As String * 100 '人员类别
        txtInfo(7).Text = .UNIT_CODE           ' As String * 100 '单位编码
        txtInfo(8).Text = .UNIT_NAME           ' As String * 100 '单位名称
        txtInfo(9).Text = .OFFICAL_TYP         ' As String * 100 '公务员类别
        txtInfo(10).Text = .HAI_TYP            ' As String * 100 '补充医保名称
        txtInfo(11).Text = .ACCT_STS           ' As String * 100 '医保账户状态
        txtInfo(12).Text = .HI_ACCT_PWD        ' As String * 100 '医保帐户口令
        txtInfo(13).Text = .SILL_PAY_AMT_TOTAL ' As Single       '年内进入门诊特殊疾病支付金额
        txtInfo(14).Text = .SILL_YR_FUND_AMT   ' As Single       '年内门诊统筹基金支付金额
        txtInfo(15).Text = .YR_FUND_AMT        ' As Single       '年内统筹基金支付金额
        txtInfo(16).Text = .HAI_YR_HIGH_AMT    ' As Single       '年内补充高额支付金额
        txtInfo(17).Text = .HAI_YR_INBED_AMT   ' As Single       '年内补充住院补助支付金额
        txtInfo(18).Text = .GZ_CUR_AMT         ' As Single       '个人账户余额
        txtInfo(19).Text = .YR_INBED_CNT       ' As Integer      '年内住院次数
        txtInfo(1).Text = .CARD_NO
    End With
    cmdOK.Enabled = True
    cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    If mlng病人ID <> 0 Then Call ReadPatient
End Sub

Private Sub ClearCons()
    Dim intClear As Integer, intCOUNT As Integer
    '清除垃圾数据
    
    intCOUNT = txtInfo.UBound - 1
    For intClear = 0 To intCOUNT
        txtInfo(intClear).Text = ""
    Next
End Sub

Private Sub ReadPatient()
    '
End Sub

Private Sub WriteFace()
    '
End Sub
