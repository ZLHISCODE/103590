VERSION 5.00
Begin VB.Form frmDistOver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "完成就诊"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5415
      TabIndex        =   8
      Top             =   3855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4065
      TabIndex        =   7
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   -15
      TabIndex        =   10
      Top             =   3600
      Width           =   6675
   End
   Begin VB.CommandButton cmdDoct 
      Caption         =   "…"
      Height          =   240
      Left            =   6195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3225
      Width           =   270
   End
   Begin VB.TextBox txtDoct 
      Height          =   300
      Left            =   4500
      TabIndex        =   5
      Top             =   3195
      Width           =   1995
   End
   Begin VB.TextBox txtMain 
      Height          =   2670
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   6420
   End
   Begin VB.ComboBox CboRoom 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3210
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "医生(&D)"
      Height          =   180
      Left            =   3840
      TabIndex        =   4
      Top             =   3270
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "就诊诊室(&R)"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   3270
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "就诊摘要"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmDistOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mobjQueue As Object, mbytQueueType As Byte, mstrPrivs As String, mstrQueuePrivs As String
Private mlngModule As Long, mlng病人ID As Long, mstrNo As String, mblnOk As Boolean
Private mlng挂号ID As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long

Public Function zlShowEdit(ByVal frmMain As Form, ByVal strPrivs As String, ByVal strQueuePrivs As String, _
        ByVal objQueue As Object, ByVal lngModule As Long, _
        strNO As String, lng病人ID As Long, str缺省诊室 As String, str缺省医生 As String, bytQueueType As Byte, _
        Optional lng挂号ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：程序入口
    '入参：frmMain-调用的主窗体
    '         objQueue-排队叫号对象
    '         lngModule-模块号
    '         bytQueueType-排队号号模式
    '出参：
    '返回：成功,返回true,否则反回False
    '编制：刘兴洪
    '日期：2010-06-03 17:22:29
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    mlng病人ID = lng病人ID: mlngModule = lngModule: mstrNo = strNO: mblnOk = False
    mlng挂号ID = lng挂号ID
    Set mobjQueue = objQueue: mbytQueueType = bytQueueType: mstrPrivs = strPrivs: mstrQueuePrivs = strQueuePrivs
    Err = 0: On Error GoTo Errhand:

    '  读出所有的该号类的诊室供选择
    If gbytRegistMode = 0 Then
        strSQL = _
            " Select b.编码,b.名称,b.位置" & vbCrLf & _
            " From 挂号安排诊室 a,门诊诊室 b,挂号安排 c,病人挂号记录 d" & vbCrLf & _
            " Where a.门诊诊室=b.名称 And a.号表ID=c.id And c.号码=d.号别 And d.NO=[1] AND d.记录性质=1 and d.记录状态=1 "
    Else
        If Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            strSQL = _
                " Select b.编码,b.名称,b.位置" & vbCrLf & _
                " From 挂号安排诊室 a,门诊诊室 b,挂号安排 c,病人挂号记录 d" & vbCrLf & _
                " Where a.门诊诊室=b.名称 And a.号表ID=c.id And c.号码=d.号别 And d.NO=[1] AND d.记录性质=1 and d.记录状态=1 "
        Else
            strSQL = _
                " Select b.编码,b.名称,b.位置" & vbCrLf & _
                " From 临床出诊诊室记录 a,门诊诊室 b,临床出诊记录 c,病人挂号记录 d" & vbCrLf & _
                " Where a.诊室id=b.id And a.记录ID=c.id And c.id=d.出诊记录id And d.NO=[1] AND d.记录性质=1 and d.记录状态=1 "
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    With rsTemp
        CboRoom.Clear
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                CboRoom.AddItem zlCommFun.Nvl(!名称)
                If CboRoom.ListIndex < 0 Then CboRoom.ListIndex = CboRoom.NewIndex
                 If zlCommFun.Nvl(!名称) = str缺省诊室 Then CboRoom.ListIndex = CboRoom.NewIndex
                .MoveNext
        Loop
    End With
    Me.Show 1, frmMain
    zlShowEdit = mblnOk
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function QueueStauteUpdate() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：排队的执行状态更改
    '编制：刘兴洪
    '日期：2010-06-03 18:11:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte

    If mbytQueueType = 0 Then
        QueueStauteUpdate = True
        Exit Function
    End If
    If mobjQueue Is Nothing Then Exit Function
    If Not (InStr(mstrQueuePrivs, ";基本;") > 0) Then Exit Function

    strSQL = "SELECT ID,执行部门ID,诊室,执行人 From 病人挂号记录  where NO=[1]  AND 记录性质=1 and 记录状态=1"
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo)
    If rsTemp.EOF Then Exit Function
    
    strQueueName = Nvl(rsTemp!执行部门id)
    lngID = Val(Nvl(rsTemp!ID))
    '完成就诊
    mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    QueueStauteUpdate = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Sub CboRoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDoct_Click()
    On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "Select 执行部门ID From 病人挂号记录 Where NO=[1]  AND 记录性质=1 And 记录状态=1"
    mstrSQL = _
        " Select c.编号,c.姓名,c.简码,c.id From 人员性质说明 a, 部门人员 b ,人员表 c" & vbCrLf & _
        " Where b.人员id=c.id And b.人员id=a.人员id  And  a.人员性质='医生' And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) " & vbCrLf & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
        " And b.部门id in (" & mstrSQL & ") "
    '医生
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrNo)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mstrSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "编号,1200,0,2;姓名,1500,0,1;简码,1500,0,2;id,1,0,2", 1, "医生选择", , Me.txtDoct.Text, 1, , 6000)
        If mstrSQL <> "" Then
            Me.txtDoct.Text = mstrSQL
        End If
    Else
        MsgBox "无任何医生可以选择！", vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    If Trim(Me.txtDoct.Text) = "" Then
        MsgBox "请输入接诊医生！", vbInformation, gstrSysName
        Me.txtDoct.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.CboRoom.Text) = "" Then
        MsgBox "请选择接诊诊室！", vbInformation, gstrSysName
        Me.CboRoom.SetFocus
        Exit Sub
    End If
    
    If ExcPlugInFun(2, mlng挂号ID, Me.txtDoct.Text, Me.CboRoom.Text) = False Then Exit Sub
    
    mstrSQL = Replace(Me.txtMain.Text, "'", "''")
    '病人ID_IN 病人信息.病人ID%TYPE,
    'NO_IN     病人挂号记录.NO%TYPE,
    '诊室_IN   病人挂号记录.诊室%TYPE:=NULL,
    '执行人_IN 病人挂号记录.执行人%TYPE:=NULL,
    '摘要_IN   病人挂号记录.摘要%TYPE:=NULL
    mstrSQL = "ZL_病人接诊完成(" & mlng病人ID & ",'" & mstrNo & "','" & Me.CboRoom.Text & "','" & Me.txtDoct.Text & "','" & mstrSQL & "',1)"
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(mstrSQL, Me.Caption)
    If QueueStauteUpdate = False Then
            gcnOracle.RollbackTrans: Exit Sub
    End If
    gcnOracle.CommitTrans
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDoct_GotFocus()
    zlControl.TxtSelAll Me.txtDoct
End Sub

Private Sub txtDoct_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    Dim CurPoint As POINTAPI
    Dim rsTmp As ADODB.Recordset
    Dim strWidth As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        mstrSQL = _
            " (c.编号 like '" & UCase(txtDoct.Text) & "%' or " & _
            "  c.姓名 like '" & gstrLike & UCase(txtDoct.Text) & "%' or " & _
            "  c.简码 like '" & gstrLike & UCase(txtDoct.Text) & "%' ) "
        
        mstrSQL = _
            "Select c.编号,c.姓名,c.简码,c.id From 人员性质说明 a, 部门人员 b ,人员表 c" & vbCrLf & _
            " Where b.人员id=c.id And b.人员id=a.人员id  And  a.人员性质='医生' And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And " & mstrSQL & vbCrLf & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            " And b.部门id in (Select 执行部门ID From 病人挂号记录 Where NO=[1] And 记录性质=1 And 记录状态=1) "
        '医生
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mstrNo)
        If rsTmp.RecordCount > 1 Then
            rsTmp.MoveFirst
            '定位选择器
            CurPoint.X = (txtDoct.Left) / Screen.TwipsPerPixelX
            CurPoint.Y = (txtDoct.Top + txtDoct.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Me.Hwnd, CurPoint
            '初始选择器
            strWidth = "1000;1200;1200;0"
            strWidth = frmSelectChild.ShowSelectChild(Me, CurPoint.X * Screen.TwipsPerPixelX, CurPoint.Y * Screen.TwipsPerPixelY, 3400 + 30 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            '求出返回的参数
            txtDoct.Text = Split(strWidth, ";")(1)
            zlCommFun.PressKey vbKeyTab
        ElseIf rsTmp.RecordCount = 1 Then
            txtDoct.Text = zlCommFun.Nvl(rsTmp!姓名)
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    mstrSQL = "Select 摘要,执行人 From 病人挂号记录 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, mstrSQL, Me.Caption)
    
    txtMain.MaxLength = rsTmp.Fields("摘要").DefinedSize
    txtDoct.MaxLength = rsTmp.Fields("执行人").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
