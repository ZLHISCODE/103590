VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm门诊结算数据核对_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊结算数据核对_贵阳"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   Icon            =   "frm门诊结算数据核对_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk本地 
      Caption         =   "仅查询本机的数据"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2835
   End
   Begin VB.CheckBox chk历史数据 
      Caption         =   "查询最近两个月的历史数据"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   5700
      Width           =   2835
   End
   Begin VB.CommandButton cmd查询 
      Caption         =   "查询(&R)"
      Height          =   350
      Left            =   210
      TabIndex        =   3
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmd冲正 
      Caption         =   "冲正(&O)"
      Height          =   350
      Left            =   9180
      TabIndex        =   4
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   9180
      TabIndex        =   5
      Top             =   5730
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   13275520
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm门诊结算数据核对_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mint性质 As Integer

Public Sub ShowME(ByVal int性质 As Integer, ByVal intinsure As Integer)
    mintInsure = intinsure
    mint性质 = int性质
    Me.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmd查询_Click()
    Call LoadData
End Sub

Private Sub cmd冲正_Click()
    Dim blnOK As Boolean
    Dim intRow As Integer, intRows As Integer
    
    If MsgBox("你确定要对所选择的数据发起结算作废交易吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intRows = mshDetail.Rows - 1
    For intRow = 1 To intRows
        If Val(mshDetail.TextMatrix(intRow, 0)) <> 0 And mshDetail.TextMatrix(intRow, mshDetail.Cols - 1) = "√" Then
            If mint性质 = 1 Then
                If blnOK = False Then blnOK = 门诊冲正(intRow)
            Else
                'Call 住院冲正(intRow)
            End If
        End If
    Next
    
    If blnOK Then Call LoadData
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub LoadData()
    Dim strStart As String
    Dim strStation As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '提取当前操作员的当天的异常记录
    
    mshDetail.Rows = 2
    mshDetail.Cols = 2
    mshDetail.Clear
    
    strStation = AnalyseComputer
    strStation = " And 工作站='" & strStation & "'"
    If chk本地.Value = 0 Then strStation = ""
    If chk历史数据.Value = 0 Then
        strStart = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    Else
        strStart = Format(DateAdd("m", -2, zlDatabase.Currentdate), "yyyy-MM-dd 00:00:00")
    End If
    
    gstrSQL = "" & _
              "        (Select 结帐ID From 结算日志_贵阳 " & _
              "         Where Nvl(已冲正,0)=0 And 性质=" & mint性质 & _
                        strStation & " And 本地时间 >= to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
              "         MINUS" & _
              "         Select 记录ID From 保险结算记录" & _
              "         Where 性质=" & mint性质 & _
                        strStation & " And 结算时间 >= to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')) B"
    gstrSQL = " Select A.结帐ID,A.病人ID,C.医保号,D.姓名,A.就诊顺序号,A.结算编号," & _
              "        A.支付类别,DECODE(A.支付类别,'11','普通','特殊') AS 支付类别名称,A.操作员,A.工作站,A.本地时间 AS 结算时间,'√' AS 标志" & _
              " From 结算日志_贵阳 A," & gstrSQL & ",保险帐户 C,病人信息 D" & _
              " Where A.结帐ID=B.结帐ID And A.病人ID=C.病人ID And A.病人ID=D.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取未结算的异常数据")
    If rsTemp.RecordCount = 0 Then
        MsgBox "未发现异常数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mshDetail.DataSource = rsTemp
    mshDetail.ColWidth(10) = 2000
    mshDetail.ColWidth(11) = 600
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mshDetail_DblClick()
    Call mshDetail_KeyDown(vbKeySpace, 0)
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Val(mshDetail.TextMatrix(mshDetail.Row, 0)) = 0 Then Exit Sub
        If mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = "" Then
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = "√"
        Else
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = ""
        End If
    End If
End Sub

Private Function 门诊冲正(ByVal intRow As Integer) As Boolean
    Dim bln离休 As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '完成中心数据冲正，以及本地数据的更改
    
    gstrSQL = " Select 1 From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在HIS的结算数据", CLng(Val(mshDetail.TextMatrix(intRow, 0))))
    If rsTemp.RecordCount <> 0 Then
        MsgBox "该记录是正常结算记录,不允许作废！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", mshDetail.TextMatrix(intRow, 4))    ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", mshDetail.TextMatrix(intRow, 5))    ' 结算编号
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", mshDetail.TextMatrix(intRow, 6))   ' 支付类别
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))   ' 办理日期
    
    '调用接口
    bln离休 = IS离休(Val(mshDetail.TextMatrix(intRow, 1)))
    If CommServer("RETBALANCE", IIf(bln离休, 1, 0)) = False Then Exit Function
    
    '更新
    gcnOracle.Execute "ZL_结算日志_贵阳_冲正(" & Val(mshDetail.TextMatrix(intRow, 0)) & ")", , adCmdStoredProc
    门诊冲正 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


