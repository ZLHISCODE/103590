VERSION 5.00
Begin VB.Form frm生育保险申报单 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生育保险申报单"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frm生育保险申报单.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt产前检查费 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1200
      TabIndex        =   28
      Top             =   2400
      Width           =   1155
   End
   Begin VB.ComboBox cbo保险类别 
      Height          =   300
      Left            =   3300
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "计生住院"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   2
      Left            =   5430
      TabIndex        =   20
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   22
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   24
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt统筹基金 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   26
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl费用总额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "分娩住院非包干"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   1
      Left            =   2820
      TabIndex        =   13
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt统筹基金 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   19
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   15
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl费用总额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "分娩住院包干"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt费用总额 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   10
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt统筹基金 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl费用总额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "费用总额"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.ComboBox cbo期号 
      Height          =   300
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1665
   End
   Begin VB.CommandButton cmd取数 
      Caption         =   "取数(&D)"
      Height          =   350
      Left            =   5310
      TabIndex        =   4
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmd申报 
      Caption         =   "申报(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   5
      Top             =   210
      Width           =   1100
   End
   Begin VB.Label lbl产前检查费 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "产前检查费"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label lbl保险类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "保险类别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2520
      TabIndex        =   2
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lbl期号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "期号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frm生育保险申报单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long              '0-新增;非零表示查阅
Private mblnOK As Boolean           '编辑成功

Private Enum 分类
    分娩住院包干
    分娩住院非包干
    计生住院
End Enum
'2、生育申报清算中，
'   a、分娩住院包干（保险类别为生育，入院方式不是计划生育的，清算=5）
'   b、计生（入院方式为计划生育）
'   c、非包干（保险类别等于生育的-分娩包干-计生）

Public Function ShowME(ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngID = lngID
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmd取数_Click()
    Dim str期号 As String, str开始日期 As String, str结束日期 As String, str上期结束日期 As String
    Dim int就诊人次 As Integer, dbl费用总额 As Double, dbl统筹基金 As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mlngID <> 0 Then
        '查阅模式
        Unload Me
        Exit Sub
    End If
    
    '清空
    Call ClearCons
    
    str期号 = Me.cbo期号.Text
    str开始日期 = Mid(str期号, 1, 4) & "-" & Mid(str期号, 5, 2) & "-01 00:00:00"
    gstrSQL = " SELECT last_day(to_date('" & Mid(str开始日期, 1, 10) & "','yyyy-MM-dd')) from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取月度最后一天")
    str结束日期 = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd") & " 23:59:59"
    str上期结束日期 = Format(DateAdd("d", -1, str开始日期), "yyyy-MM-dd")
    
    '根据设定的条件取数
    '1、分娩住院包干（保险类别为生育，入院方式不是计划生育的，清算=5）
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 就诊人次, " & _
             "        NVL(SUM(NVL(B.医保总费用,0)),0) AS 医保总费用, " & _
             "        nvl(sum(c.统筹基金),0) As 统筹基金 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B," & _
             "      (Select 结帐id,Nvl(Sum(Decode(结算方式, '医保基金', Nvl(冲预交, 0), 0)), 0) As 统筹基金 " & _
             "      From 病人预交记录 " & _
             "      Where 收款时间 BETWEEN [3] AND [4]" & _
             "      Group By 结帐id) C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID And A.医疗类别<>'32'" & _
             " AND B.清算方式=5 And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "控制线住院", CInt(cbo保险类别.ItemData(cbo保险类别.ListIndex)), TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(分娩住院包干).Text = Format(rsTemp!就诊人次, "#0;-#0; ;")
    Me.txt费用总额(分娩住院包干).Text = Format(rsTemp!医保总费用, "#0.00;-#0.00; ;")
    Me.txt统筹基金(分娩住院包干).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    
    '2、计生住院
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 就诊人次, " & _
             "        NVL(SUM(NVL(B.医保总费用,0)),0) AS 医保总费用, " & _
             "        nvl(sum(c.统筹基金),0) As 统筹基金 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B," & _
             "      (Select 结帐id,Nvl(Sum(Decode(结算方式, '医保基金', Nvl(冲预交, 0), 0)), 0) As 统筹基金 " & _
             "      From 病人预交记录 " & _
             "      Where 收款时间 BETWEEN [3] AND [4]" & _
             "      Group By 结帐id) C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID And A.医疗类别='32'" & _
             " AND A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重症住院", CInt(cbo保险类别.ItemData(cbo保险类别.ListIndex)), TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(计生住院).Text = Format(rsTemp!就诊人次, "#0;-#0; ;")
    Me.txt费用总额(计生住院).Text = Format(rsTemp!医保总费用, "#0.00;-#0.00; ;")
    Me.txt统筹基金(计生住院).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    
    '3、统计保险类别等于生育的所有数据
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 就诊人次, " & _
             "        NVL(SUM(NVL(B.医保总费用,0)),0) AS 医保总费用, " & _
             "        nvl(sum(c.统筹基金),0) As 统筹基金 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B," & _
             "      (Select 结帐id,Nvl(Sum(Decode(结算方式, '医保基金', Nvl(冲预交, 0), 0)), 0) As 统筹基金 " & _
             "      From 病人预交记录 " & _
             "      Where 收款时间 BETWEEN [3] AND [4]" & _
             "      Group By 结帐id) C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID " & _
             " AND A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重症住院", CInt(cbo保险类别.ItemData(cbo保险类别.ListIndex)), TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    int就诊人次 = rsTemp!就诊人次 - Val(txt就诊人次(分娩住院包干).Text) - Val(txt就诊人次(计生住院).Text)
    dbl费用总额 = rsTemp!医保总费用 - Val(txt费用总额(分娩住院包干).Text) - Val(txt费用总额(计生住院).Text)
    dbl统筹基金 = rsTemp!统筹基金 - Val(txt统筹基金(分娩住院包干).Text) - Val(txt统筹基金(计生住院).Text)
    
    '4、提取产前检查费合计
    gstrSQL = "SELECT Nvl(Sum(Decode(结算方式, '产前检查费', Nvl(冲预交, 0), 0)), 0) As 产前检查费 " & _
             "      From 病人预交记录 " & _
             "      Where 收款时间 BETWEEN [1] AND [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "产前检查费", CDate(str开始日期), CDate(str结束日期))
    Me.txt产前检查费.Text = Format(rsTemp!产前检查费, "#0.00;-#0.00; ;")
    
    Me.txt就诊人次(分娩住院非包干).Text = Format(int就诊人次, "#0;-#0; ;")
    Me.txt费用总额(分娩住院非包干).Text = Format(dbl费用总额, "#0.00;-#0.00; ;")
    Me.txt统筹基金(分娩住院非包干).Text = Format(dbl统筹基金, "#0.00;-#0.00; ;")
    
    Me.Tag = 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ClearCons
End Sub

Private Sub cmd申报_Click()
    Dim str流水号 As String
    On Error GoTo errHand
    
    If Val(Me.Tag) = 0 Then
        MsgBox "请指定条件后点“取数”按钮！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnGYYB.BeginTrans
    '对XML DomDocument对象进行初始化
    If InitXML = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    '住院虚拟结算只要求传入个人编码，正式结算时才要求传入磁卡数据及密码
    Call InsertChild(mdomInput.documentElement, "PERIOD", cbo期号.Text)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo保险类别.ItemData(cbo保险类别.ListIndex))
    Call InsertChild(mdomInput.documentElement, "FMBGPSNS", Val(txt就诊人次(分娩住院包干).Text))                 ' 门诊就诊人次
    Call InsertChild(mdomInput.documentElement, "FMBGFEEALL", Val(txt费用总额(分娩住院包干).Text))
    Call InsertChild(mdomInput.documentElement, "FMBGFUND", Val(txt统筹基金(分娩住院包干).Text))
    Call InsertChild(mdomInput.documentElement, "FMPSNS", Val(txt就诊人次(分娩住院非包干).Text))
    Call InsertChild(mdomInput.documentElement, "FMFEEALL", Val(txt费用总额(分娩住院非包干).Text))
    Call InsertChild(mdomInput.documentElement, "FMFUND", Val(txt统筹基金(分娩住院非包干).Text))
    Call InsertChild(mdomInput.documentElement, "JSPSNS", Val(txt就诊人次(计生住院).Text))
    Call InsertChild(mdomInput.documentElement, "JSFEEALL", Val(txt费用总额(计生住院).Text))
    Call InsertChild(mdomInput.documentElement, "JSFUND", Val(txt统筹基金(计生住院).Text))
    Call InsertChild(mdomInput.documentElement, "JCF", Val(txt产前检查费.Text))
    '调用接口
    If CommRecServer("APPRECB") = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    str流水号 = GetElemnetValue("APPNO")
    
    '产生数据
    mlngID = GetNextID("清算单", gcnGYYB)
    gstrSQL = "ZL_清算单_INSERT(" & mlngID & ",1,'" & Me.cbo期号.Text & "'," & cbo保险类别.ItemData(cbo保险类别.ListIndex) & "," & _
        "'" & cbo保险类别.Text & "','" & gstrUserName & "',sysdate,'" & str流水号 & "',NULL)"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "ZL_生育清算明细_INSERT(" & mlngID & "," & Val(txt就诊人次(分娩住院包干).Text) & "," & Val(txt费用总额(分娩住院包干).Text) & "," & Val(txt统筹基金(分娩住院包干).Text) & "," & _
            Val(txt就诊人次(分娩住院非包干).Text) & "," & Val(txt费用总额(分娩住院非包干).Text) & "," & Val(txt统筹基金(分娩住院非包干).Text) & "," & _
            Val(txt就诊人次(计生住院).Text) & "," & Val(txt费用总额(计生住院).Text) & "," & Val(txt统筹基金(计生住院).Text) & "," & Val(txt产前检查费.Text) & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    gcnGYYB.CommitTrans
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnGYYB.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim str上月 As String, str本月 As String
    Dim rsData As New ADODB.Recordset
    
    If mlngID = 0 Then
        '缺省只装入上月、本月供申报
        curDate = zlDatabase.Currentdate()
        str上月 = Format(DateAdd("m", -1, curDate), "yyyyMM")
        str本月 = Format(curDate, "yyyyMM")
        With cbo期号
            .Clear
            .AddItem str上月
            .AddItem str本月
            .ListIndex = 0
        End With
        With cbo保险类别
            .Clear
            .AddItem "企业生育保险"
            .ItemData(.NewIndex) = 4
            .AddItem "机关事业单位生育保险"
            .ItemData(.NewIndex) = 5
            .ListIndex = 0
        End With
        Exit Sub
    End If
    
    '读取申报单数据
    gstrSQL = "SELECT  " & _
             "        A.ID, A.期号, A.保险类别, A.操作员, A.日期 ,B.分娩包干人次, B.分娩包干费用总额, B.分娩包干统筹支付, B.分娩非包干人次, " & _
             "        B.分娩非包干费用总额, B.分娩非包干统筹支付, B.计生人次, B.计生费用总额, B.计生统筹支付, A.清算流水号, A.处理情况 " & _
             " FROM 清算单 A, 生育清算明细 B " & _
             " WHERE A.ID=B.清算单ID AND A.ID= [1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取申报单数据", mlngID)
    
    '填数
    With rsData
        Me.cbo期号.AddItem !期号
        Me.cbo期号.ListIndex = 0
        
        Me.txt就诊人次(分娩住院包干).Text = Format(Nvl(!分娩包干人次, 0), "#0;-#0; ;")
        Me.txt费用总额(分娩住院包干).Text = Format(Nvl(!分娩包干费用总额, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(分娩住院包干).Text = Format(Nvl(!分娩包干统筹支付, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(分娩住院非包干).Text = Format(Nvl(!分娩非包干人次, 0), "#0;-#0; ;")
        Me.txt费用总额(分娩住院非包干).Text = Format(Nvl(!分娩非包干费用总额, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(分娩住院非包干).Text = Format(Nvl(!分娩非包干统筹支付, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(计生住院).Text = Format(Nvl(!计生人次, 0), "#0;-#0; ;")
        Me.txt费用总额(计生住院).Text = Format(Nvl(!计生费用总额, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(计生住院).Text = Format(Nvl(!计生统筹支付, 0), "#0.00;-#0.00; ;")
    End With
    
    '设置控件状态
    Me.cbo期号.Enabled = False
    
    cmd申报.Visible = False
    cmd取数.Caption = "退出(&X)"
End Sub

Private Sub ClearCons()
    Me.Tag = ""
    Me.txt就诊人次(分娩住院包干).Text = ""
    Me.txt费用总额(分娩住院包干).Text = ""
    Me.txt统筹基金(分娩住院包干).Text = ""
    
    Me.txt就诊人次(分娩住院非包干).Text = ""
    Me.txt费用总额(分娩住院非包干).Text = ""
    Me.txt统筹基金(分娩住院非包干).Text = ""
    
    Me.txt就诊人次(计生住院).Text = ""
    Me.txt费用总额(计生住院).Text = ""
    Me.txt统筹基金(计生住院).Text = ""
End Sub
