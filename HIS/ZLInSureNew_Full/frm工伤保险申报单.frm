VERSION 5.00
Begin VB.Form frm工伤保险申报单 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "工伤保险申报单"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frm工伤保险申报单.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "住院"
      Enabled         =   0   'False
      Height          =   1245
      Index           =   1
      Left            =   2790
      TabIndex        =   9
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   11
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt统筹基金 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   13
         Top             =   720
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
         TabIndex        =   10
         Top             =   360
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
         TabIndex        =   12
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "门诊"
      Enabled         =   0   'False
      Height          =   1245
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   720
      Width           =   2445
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   6
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt统筹基金 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   8
         Top             =   720
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
         TabIndex        =   5
         Top             =   360
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
         TabIndex        =   7
         Top             =   780
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
      Left            =   2580
      TabIndex        =   2
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmd申报 
      Caption         =   "申报(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   3
      Top             =   210
      Width           =   1100
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
Attribute VB_Name = "frm工伤保险申报单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long              '0-新增;非零表示查阅
Private mblnOK As Boolean           '编辑成功

Const int工伤保险 As Integer = 7
Const str工伤保险 As String = "工伤保险"

Private Enum 分类
    门诊
    住院
End Enum

Public Function ShowME(ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngID = lngID
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmd取数_Click()
    Dim str期号 As String, str开始日期 As String, str结束日期 As String, str上期结束日期 As String
    Dim int就诊人次 As Integer, dbl统筹基金 As Double
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
    '1、门诊（保险类别为工伤，入院方式不是计划工伤的，清算=5）
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 就诊人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.性质=1 And A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID " & _
             " AND A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "控制线住院", int工伤保险, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(门诊).Text = Format(rsTemp!就诊人次, "#0;-#0; ;")
    Me.txt统筹基金(门诊).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    
    '2、住院
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 就诊人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.性质=2 And A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID " & _
             " AND A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重症住院", int工伤保险, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(住院).Text = Format(rsTemp!就诊人次, "#0;-#0; ;")
    Me.txt统筹基金(住院).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    
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
    Call InsertChild(mdomInput.documentElement, "MZPSNS", Val(txt就诊人次(门诊).Text))                 ' 门诊就诊人次
    Call InsertChild(mdomInput.documentElement, "MZFUND", Val(txt统筹基金(门诊).Text))
    Call InsertChild(mdomInput.documentElement, "ZYPSNS", Val(txt就诊人次(住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZYFUND", Val(txt统筹基金(住院).Text))
    '调用接口
    If CommRecServer("APPRECG") = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    str流水号 = GetElemnetValue("APPNO")
    
    '产生数据
    mlngID = GetNextID("清算单", gcnGYYB)
    gstrSQL = "ZL_清算单_INSERT(" & mlngID & ",2,'" & Me.cbo期号.Text & "'," & int工伤保险 & "," & _
        "'" & str工伤保险 & "','" & gstrUserName & "',sysdate,'" & str流水号 & "',NULL)"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "ZL_工伤清算明细_INSERT(" & mlngID & "," & Val(txt就诊人次(门诊).Text) & "," & Val(txt统筹基金(门诊).Text) & "," & _
            Val(txt就诊人次(住院).Text) & "," & Val(txt统筹基金(住院).Text) & ")"
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
        Exit Sub
    End If
    
    '读取申报单数据
    gstrSQL = "SELECT  " & _
             "        A.ID, A.期号, A.保险类别, A.操作员, A.日期 ,B.门诊人次,  B.门诊统筹支付, B.住院人次, B.住院统筹支付, A.清算流水号, A.处理情况 " & _
             " FROM 清算单 A, 工伤清算明细 B " & _
             " WHERE A.ID=B.清算单ID AND A.ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取申报单数据", mlngID)
    
    '填数
    With rsData
        Me.cbo期号.AddItem !期号
        Me.cbo期号.ListIndex = 0
        
        Me.txt就诊人次(门诊).Text = Format(Nvl(!门诊人次, 0), "#0;-#0; ;")
        Me.txt统筹基金(门诊).Text = Format(Nvl(!门诊统筹支付, 0), "#0.00;-#0.00; ;")
        Me.txt就诊人次(住院).Text = Format(Nvl(!住院人次, 0), "#0;-#0; ;")
        Me.txt统筹基金(住院).Text = Format(Nvl(!住院统筹支付, 0), "#0.00;-#0.00; ;")
    End With
    
    '设置控件状态
    Me.cbo期号.Enabled = False
    
    cmd申报.Visible = False
    cmd取数.Caption = "退出(&X)"
End Sub

Private Sub ClearCons()
    Me.Tag = ""
    Me.txt就诊人次(门诊).Text = ""
    Me.txt统筹基金(门诊).Text = ""
    Me.txt就诊人次(住院).Text = ""
    Me.txt统筹基金(住院).Text = ""
End Sub
