VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmSock铜仁 
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "frmSock铜仁.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   5445
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer timUnload 
      Interval        =   1000
      Left            =   630
      Top             =   180
   End
   Begin MSWinsockLib.Winsock sckCenter 
      Left            =   210
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在与医保中心进行数据交换……"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Top             =   780
      Width           =   4500
   End
End
Attribute VB_Name = "frmSock铜仁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mint秒数 As Integer

Private mstr中心地址 As String
Private mbln读卡   As Boolean
Private mint场合   As Integer
Private mstr数据体 As String   '写卡的参数太多，传起来比较麻烦，所以干脆采用直接传递数据体的方式

Public Function CommIC(ByVal 中心地址 As String, ByVal 读卡 As Boolean, ByVal 场合 As Integer, ByVal 数据体 As String) As Boolean
'功能：与中心进行数据交流，模仿IC卡操作
'参数：读卡   读卡时为True,写卡则为False
'      场合   0 表示门诊;1 表示住院

    '###############确保之前不能使用任何控件，避免调用Load事件
    mbln读卡 = 读卡
    mint场合 = False
    mblnOK = False
    mint秒数 = 0
    mstr中心地址 = 中心地址
    mstr数据体 = 数据体
    '###############确保之前不能使用任何控件，避免调用Load事件
    
    frmSock铜仁.Show vbModal
    CommIC = mblnOK
End Function

Private Sub Form_Load()
    On Error Resume Next
    sckCenter.RemoteHost = mstr中心地址
    sckCenter.RemotePort = 1800
    Me.Visible = False
    sckCenter.Connect
End Sub

Private Sub sckCenter_Connect()
    Dim str操作定义 As String
    Dim strInput As String
    
    On Error Resume Next
    '构成发送数据
    If mbln读卡 = True Then
        '查询信息
        str操作定义 = "000"
    Else
        '更新
        str操作定义 = "100"
    End If
    
    strInput = "*|" & mint场合 & "|" & str操作定义 & "|" & LenB(StrConv(mstr数据体, vbFromUnicode)) & "|" & mstr数据体 & "|*"
    '连接已经建立，可以发送数据
    sckCenter.SendData strInput
    If Err <> 0 Then
        '出现错误，肯定有问题
        MsgBox "数据发送失败。", vbInformation, gstrSysName
        Unload Me
    End If
End Sub

Private Sub sckCenter_DataArrival(ByVal bytesTotal As Long)
    Dim str数据体 As String, str操作定义 As String
    Dim strData As String, var返回 As Variant
    
    On Error Resume Next
    '分析连接数据
    timUnload.Enabled = False '不再设置超时了
    sckCenter.GetData strData, vbString
    If Err <> 0 Then
        '出现错误，肯定有问题
        MsgBox "数据接收失败。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    var返回 = Split(strData, "|")
    If UBound(var返回) < 6 Then
        MsgBox "数据接收格式错误。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If var返回(0) <> "*" Or var返回(UBound(var返回)) <> "*" Then
        MsgBox "数据接收格式错误。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    If var返回(1) <> "10" Then
        If var返回(1) = "11" Then
            If mint场合 <> 1 Then
                MsgBox "病人正在住院，不能继续。", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        Else
            MsgBox "数据接收格式错误。", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    
    If mbln读卡 = True Then
        '中心代码|职工编码|姓名|性别|出生日期|单位代码|职工身份|属地代码|是否公务员|是否参加补充|帐户累计注入|帐户累计支出|统筹支付费用累积|统筹支付金额累积|有效住院次数
        With gIC铜仁Temp
            .CenterCode = var返回(5)
            .Cardno = var返回(6)          ' 卡号
            .IDCardno = Split(mstr数据体, "|")(0)       ' 身份证号 长度不足后补#0
            .MediAccountNo = var返回(6)  ' 医保号
            .Name = var返回(7)           ' 姓名
            .Sex = var返回(8)            ' 性别 1-男  0-女
            .Birthday = var返回(9)       ' 出生日期 YYYYMMDD
            .UnitCode = var返回(10)       ' 用人单位编码
            .ClassCode = var返回(11)      ' 职工身份：0x：在职1x：退休, 05和11为一次性缴费
            .DomainCode = var返回(12)     ' 职工属地 0-正常 1-常驻外地 2-异地安置
            .MediYear = Year(zldatabase.Currentdate())       ' 医保年度
            .InNo = 0           ' 装钱期次
            .OutSerialNo = 0    ' 支付顺序号
            .InPerAcc = var返回(15)       ' 个人帐户累计注入金额
            .OutPerAcc = var返回(16)      ' 个人帐户累计支出金额
            .PlanPaidFee = var返回(17)    ' 统筹基金支付费用累计（基本+补充）
            .PlanPaidAmt = var返回(18)    ' 统筹基金支付金额累计（基本+补充）
            .ChronicPaidFee = 0 ' 慢性病支付费用累计
            .ChronicPaidAmt = 0 ' 慢性病支付金额累计
            .InHosPaidAmt = 0   ' 住院个人帐户支付金额
            .ClinicPaidAmt = 0  ' 门诊个人帐户支付金额
            .Password = Split(mstr数据体, "|")(1)      ' 个人密码
            .InHosTimes = var返回(19)     ' 本年有效住院次数
            .IsOffical = var返回(13)      ' 公务员 0-否；其他-是
            .IsAttend = 0       ' 医疗照顾对象 0-否；1-是
            .InpatientFlag = 0  ' 住院标志 0-不住院 1-住院
            .QuotaPaidAmt = 0   ' 慢性病额度已支付金额
            .ChronicSillPaidAmt = 0  ' 慢性病起付金已支付金额
        End With
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub sckCenter_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "连接出现错误：" & Description, vbInformation, gstrSysName
    Unload Me
End Sub

Private Sub timUnload_Timer()
    mint秒数 = mint秒数 + 1
    If mint秒数 > 30 Then
        If mbln读卡 = True Then
            '只有读卡才可以超时退出，写卡只能不停地等待
            If MsgBox("连接超时，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Unload Me
                Exit Sub
            End If
        End If
        mint秒数 = 0
    End If
End Sub
