VERSION 5.00
Begin VB.Form frmIdentify江苏 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo医疗方式 
      Height          =   300
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   150
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5535
      TabIndex        =   3
      Top             =   4695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4410
      TabIndex        =   2
      Top             =   4695
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "病人信息"
      Height          =   4035
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   6495
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   1455
         TabIndex        =   33
         Top             =   3510
         Width           =   4740
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   4125
         TabIndex        =   31
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   1095
         TabIndex        =   29
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   4125
         TabIndex        =   27
         Top             =   2595
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   1095
         TabIndex        =   25
         Top             =   2595
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   4125
         TabIndex        =   23
         Top             =   2145
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   1095
         TabIndex        =   21
         Top             =   2145
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1095
         TabIndex        =   19
         Top             =   1695
         Width           =   5100
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4125
         TabIndex        =   17
         Top             =   1230
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1095
         TabIndex        =   15
         Top             =   1230
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4125
         TabIndex        =   13
         Top             =   780
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1095
         TabIndex        =   11
         Top             =   780
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4125
         TabIndex        =   9
         Top             =   330
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1095
         TabIndex        =   7
         Top             =   330
         Width           =   2070
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "已申请特殊病"
         Height          =   180
         Index           =   14
         Left            =   315
         TabIndex        =   32
         Top             =   3585
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "行政职务"
         Height          =   180
         Index           =   13
         Left            =   3330
         TabIndex        =   30
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "本人身份"
         Height          =   180
         Index           =   12
         Left            =   315
         TabIndex        =   28
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "民族"
         Height          =   180
         Index           =   11
         Left            =   3690
         TabIndex        =   26
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         Height          =   180
         Index           =   10
         Left            =   315
         TabIndex        =   24
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "人员类别"
         Height          =   180
         Index           =   9
         Left            =   3330
         TabIndex        =   22
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   8
         Left            =   675
         TabIndex        =   20
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "单位名称"
         Height          =   180
         Index           =   7
         Left            =   315
         TabIndex        =   18
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "单位ID"
         Height          =   180
         Index           =   6
         Left            =   3510
         TabIndex        =   16
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "身份证号"
         Height          =   180
         Index           =   5
         Left            =   315
         TabIndex        =   14
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "区属代码"
         Height          =   180
         Index           =   4
         Left            =   3330
         TabIndex        =   12
         Top             =   855
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Index           =   3
         Left            =   675
         TabIndex        =   10
         Top             =   855
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   8
         Top             =   405
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "个人ID"
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   6
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdComp 
      Caption         =   "验证(&S)"
      Height          =   350
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   1100
   End
   Begin VB.TextBox txtNo 
      Height          =   300
      Left            =   1005
      MaxLength       =   10
      TabIndex        =   1
      Top             =   150
      Width           =   2070
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "个人证号"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify江苏"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strVoucherID As String, mbytType As Byte, mlng病人ID As Long, intReturn As Long
Private strArrInfo(20) As String, sngArrInfo(20) As Single, iLoop As Long
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, lng病人ID As Long) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show vbModal
    gint医疗方式 = cbo医疗方式.ListIndex + 1
    
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cbo医疗方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdComp_Click()
    Dim int有效天数 As Integer
    Dim lng病人ID As Long
    Dim datCurr As Date
    Dim strReadCard As String, intRecType As Long
    Dim intInsID As Long, i As Long, rsTemp As New ADODB.Recordset
    strReadCard = GetSetting(appName:="ZLSOFT", Section:="医保信息", Key:="ReadCard", Default:="0")
    gblnReadCard = Not strReadCard = "0"
    
    If mbytType = 1 Then
        intRecType = 1
    Else
        intRecType = 0
    End If
    
    If strReadCard = "0" Then
        If IsNumeric(txtNO.Text) = False Or Len(txtNO.Text) <> 10 Then
            MsgBox "请输入10位个人证号（个人证号中的“-”无需输入）。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Sub
        Else
            strVoucherID = Left(txtNO.Text, 5) & "-" & Right(txtNO.Text, 5)
        End If
    Else
        strVoucherID = ""
    End If
    
    If mbytType = 1 Or mbytType = 3 Then
        gstrRecCode = String(12, " ")
        intReturn = FGetRecCode(intRecType, gstrRecCode)
        If intReturn <> 0 Then
            MsgBox "在获取收费流水号时发生错误，未获得错误信息。", vbExclamation, gstrSysName
            cmdOK.Enabled = False
            Exit Sub
        End If
    ElseIf mbytType = 0 Then                    '无刷卡器读卡时可以这样操作，如果改用读卡方式就需要进行修改
        gstrSQL = "Select * From 保险帐户 where 险类=[1] And 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_江苏, strVoucherID)
        If rsTemp.EOF Then
            MsgBox "找不到病人的挂号信息，请确认该病人是否挂号或卡号输入是否正确。", vbInformation, "错误"
            Exit Sub
        End If
        gstrRecCode = Nvl(rsTemp!退休证号)
    Else
        gstrRecCode = String(12, " ")
    End If
    
    If InStr(gstrRecCode, Chr(0)) > 0 Then gstrRecCode = Left(gstrRecCode, InStr(gstrRecCode, Chr(0)) - 1)
'===============================================================================================================
'功能：读取参保人的基本信息和帐户、支出信息
'入口参数：类型（0门诊/1住院）,收费流水号,个人证号
'出口参数：0个人ID,1卡号,2姓名,3区属代码,4身份证号码,5单位ID,6单位名称,7性别(男/女),8人员类别,9出生日期,10民族,
'          11本人身份,12行政职务,13门诊特殊病种(脱机返回'未知'/联机未申请时返回'未申请'/其他返回已申请的特殊病种),
'          14其它(保留),15本年累计住院次数,16帐户总收,17帐户总支,18支出版本号,19本年统筹支付累计,20本年大病基金支付累计,
'          21本年公务员补充/企业补充支付累计,22本年普通门诊费用累计,23本年普通门诊三个范围内费用累计,
'          24本年特殊门诊三个范围内费用累计,25本年比照住院三个范围内费用累计,26本年普通住院费用累计,
'          27本年普通住院三个范围内费用累计,28本年家庭病床住院三个范围内费用累计,29其他1,30其他2,
'          31本年储蓄帐户支付累计,32本年其它基金支付累计,33本年现金支付累计,34帐户余额
'===============================================================================================================
'            psCardID        : pChar;     0         //O卡号[C16]
'        psName          : pChar;    1         //O姓名[C8]
'        psAreaCode      : pChar;     2        //O区属代码[C3]
'        psQueryID       : pChar;     3        //O身份证号码[C18]
'        psUnitID        : pChar;     4         //O单位ID[C8]
'        psUnitName      : pChar;    5         //O单位名称[C50]
'        psSex           : pChar;    6         //O性别[C2](男/女)
'        psKind          : pChar;    7         //O人员类别[C4]
'        psBirthday      : pChar;     8         //O出生日期[C10](YYYY-MM-DD)
'        psNational      : pChar;     9         //O民族[C20]
'        psIndustry      : pChar;     10         //O本人身份[C20]
'        psDuty          : pChar;    11         //O行政职务[C30]
'        psChronic       :pChar;     12     // O门诊特殊病种[C200](脱机返回'未知'/联机未申请时返回'未申请'/其他返回已申请的特殊病种)
'        psOthers1       :pChar;     13     // O其它(保留)[C200]
    strArrInfo(0) = String(16, " ")
    strArrInfo(1) = String(8, " ")
    strArrInfo(2) = String(3, " ")
    strArrInfo(3) = String(18, " ")
    strArrInfo(4) = String(8, " ")
    strArrInfo(5) = String(50, " ")
    strArrInfo(6) = String(2, " ")
    strArrInfo(7) = String(4, " ")
    strArrInfo(8) = String(10, " ")
    strArrInfo(9) = String(20, " ")
    strArrInfo(10) = String(20, " ")
    strArrInfo(11) = String(30, " ")
    strArrInfo(12) = String(200, " ")
    strArrInfo(13) = String(200, " ")
    intReturn = FGetCardInfo(intRecType, gstrRecCode, strVoucherID, intInsID, strArrInfo(0), strArrInfo(1), _
        strArrInfo(2), strArrInfo(3), strArrInfo(4), strArrInfo(5), strArrInfo(6), strArrInfo(11), strArrInfo(8), _
        strArrInfo(9), strArrInfo(10), strArrInfo(7), strArrInfo(12), strArrInfo(13), sngArrInfo(19), sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14), _
        sngArrInfo(15), sngArrInfo(16), sngArrInfo(17), sngArrInfo(18))
    If intReturn <> 0 Then
        MsgBox "在获取收病人信息时发生错误，未获得错误信息。", vbExclamation, gstrSysName
        cmdOK.Enabled = False
        txtNO.SetFocus
        Exit Sub
    End If
    
    '如果是挂号则进行检查，保险参数中指定天数内不允许再次挂号
    If mbytType = 3 Or mbytType = 0 Then
        '提取挂号天数
        int有效天数 = 2
        datCurr = zlDatabase.Currentdate()
        '取病人ID
        gstrSQL = "select 病人ID from 保险帐户 where 险类=[1] and 卡号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人id", TYPE_江苏, Substr((strArrInfo(0)), 1, 11))
        If rsTemp.RecordCount <> 0 Then
            lng病人ID = rsTemp!病人ID
            gstrSQL = " Select MAX(结帐ID) AS 结帐ID From 门诊费用记录" & _
                      " Where 记录性质=1 and 收费类别 in('5','6','7') And 记录状态=1 And 病人ID=[1]" & _
                      " And 登记时间 Between to_date('" & Format(DateAdd("d", -1 * int有效天数, datCurr), "yyyy-MM-dd") & " 00:00:00" & "','yyyy-MM-dd hh24:mi:ss')" & _
                      " And to_date('" & Format(datCurr, "yyyy-MM-dd") & " 23:59:59" & "','yyyy-MM-dd hh24:mi:ss')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取有效期内最后一次有效的挂号数据", lng病人ID)
            If Nvl(rsTemp!结帐ID, 0) <> 0 Then
                If MsgBox("3天内已挂号就诊过一次，是否允许再次就诊？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    
    For iLoop = 0 To 20
        If InStr(strArrInfo(iLoop), Chr(0)) > 0 Then
            strArrInfo(iLoop) = Left(strArrInfo(iLoop), InStr(strArrInfo(iLoop), Chr(0)) - 1)
        End If
    Next
    txtInfo(0).Text = intInsID
    For i = 1 To txtInfo.UBound
        txtInfo(i).Text = strArrInfo(i - 1)
    Next
    
    cmdOK.Enabled = True
    cbo医疗方式.SetFocus
End Sub

Private Sub cmdOK_Click()
    If Me.txtInfo(0).Text = "" Then
        MsgBox "请先确定病人身份！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
    '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
    
'出口参数：0个人ID,1卡号,2姓名,3区属代码,4身份证号码,5单位ID,6单位名称,7性别(男/女),8人员类别,9出生日期,10民族,
'          11本人身份,12行政职务,13门诊特殊病种(脱机返回'未知'/联机未申请时返回'未申请'/其他返回已申请的特殊病种),
'          14其它(保留),0本年累计住院次数,1帐户总收,2帐户总支,3支出版本号,4本年统筹支付累计,5本年大病基金支付累计,
'          6本年公务员补充/企业补充支付累计,7本年普通门诊费用累计,8本年普通门诊三个范围内费用累计,
'          9本年特殊门诊三个范围内费用累计,10本年比照住院三个范围内费用累计,11本年普通住院费用累计,
'          12本年普通住院三个范围内费用累计,13本年家庭病床住院三个范围内费用累计,14其他1,15其他2,
'          16本年储蓄帐户支付累计,17本年其它基金支付累计,18本年现金支付累计,19帐户余额
    
    mstrPatient = "": mstrOther = ""
    mstrPatient = mstrPatient & strArrInfo(0) & ";"             '卡号
    mstrPatient = mstrPatient & strVoucherID & ";"              '医保号
    mstrPatient = mstrPatient & ";"                             '密码
    mstrPatient = mstrPatient & strArrInfo(1) & ";"             '姓名
    mstrPatient = mstrPatient & strArrInfo(6) & ";"             '性别
    mstrPatient = mstrPatient & strArrInfo(8) & ";"             '出生日期
    mstrPatient = mstrPatient & strArrInfo(3) & ";"             '身份证号
    mstrPatient = mstrPatient & strArrInfo(5) & ";"             '单位名称
 
    mstrOther = mstrOther & ";"                                 '中心
    mstrOther = mstrOther & ";"                                 '顺序号
    mstrOther = mstrOther & strArrInfo(11) & ";"                '10人员身份
    mstrOther = mstrOther & sngArrInfo(18) & ";"                '11帐户余额
    mstrOther = mstrOther & ";"                                 '12当前状态
    mstrOther = mstrOther & ";"                                 '13病种ID
    mstrOther = mstrOther & strArrInfo(7) & ";"                 '14在职
    mstrOther = mstrOther & ";"                                 '15退休证号
    mstrOther = mstrOther & ";"                                 '16年龄段
    mstrOther = mstrOther & ";"                                 '17灰度级
    mstrOther = mstrOther & sngArrInfo(0) & ";"                '18帐户增加累计
    mstrOther = mstrOther & sngArrInfo(1) & ";"                '19帐户支出累计
    mstrOther = mstrOther & ";"                                 '20进入统筹累计
    mstrOther = mstrOther & sngArrInfo(4) & ";"                '21统筹报销累计
    mstrOther = mstrOther & sngArrInfo(19) & ";"                '22住院次数累计
    mstrOther = mstrOther & ";"                                 '23就诊类型
    
    Me.Hide
End Sub

Private Sub Form_Load()
    cbo医疗方式.AddItem "普通门诊"
    cbo医疗方式.AddItem "普通住院"
    cbo医疗方式.AddItem "特殊病"
    cbo医疗方式.AddItem "紧急抢救"
    cbo医疗方式.AddItem "急诊"
    cbo医疗方式.ListIndex = 0
End Sub

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO.Text)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdComp_Click
    End If
End Sub
