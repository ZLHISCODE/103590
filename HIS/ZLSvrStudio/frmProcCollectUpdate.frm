VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcCollectUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "搜集更新/过程差异检查"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmProcCollectUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgTrueFalse 
      Left            =   9360
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcCollectUpdate.frx":6852
            Key             =   "T"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcCollectUpdate.frx":6DEC
            Key             =   "F"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10455
      TabIndex        =   8
      Top             =   6015
      Width           =   10455
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   5880
         TabIndex        =   13
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&X)"
         Default         =   -1  'True
         Height          =   350
         Left            =   9000
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "开始(&S)"
         Height          =   350
         Left            =   7800
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pbrCollect 
         Height          =   105
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Visible         =   0   'False
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "开始收集"
         Height          =   180
         Left            =   135
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   10455
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdConnet 
         Caption         =   "连接配置(&L)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3090
         TabIndex        =   5
         Top             =   1125
         Width           =   1290
      End
      Begin VB.OptionButton optDB 
         Caption         =   "其他数据库"
         Height          =   255
         Index           =   1
         Left            =   1635
         TabIndex        =   4
         Top             =   1170
         Width           =   1380
      End
      Begin VB.OptionButton optDB 
         Caption         =   "当前数据库"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   1170
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.PictureBox picFunCap 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   120
         Picture         =   "frmProcCollectUpdate.frx":7386
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblFunNote 
         Caption         =   $"frmProcCollectUpdate.frx":8250
         Height          =   450
         Left            =   1020
         TabIndex        =   7
         Top             =   630
         Width           =   9180
      End
      Begin VB.Label lblFunCap 
         AutoSize        =   -1  'True
         Caption         =   "搜集登记过程/函数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         TabIndex        =   6
         Top             =   150
         Width           =   2820
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8760
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9975
      _cx             =   17595
      _cy             =   7329
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProcCollectUpdate.frx":82B2
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmProcCollectUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
Private mobjMain As Object
Private mblnOk As Boolean
Private mcnOracle As ADODB.Connection
Private mrsUpgradeFiles As ADODB.Recordset
Private mintType As Integer
Private WithEvents mfrmPageConfigure As frmProcConfigure
Attribute mfrmPageConfigure.VB_VarHelpID = -1
Private Enum SysInfoCol
    SC_序号 = 0
    SC_版本号 = 1
    SC_系统名称 = 2
    SC_安装脚本 = 3
    SC_配置版本 = 4
End Enum

Private Enum DBType
    CurDB = 0
    OtherDB = 1
End Enum
'==============================================================
'==公共接口
'==============================================================
Public Function ShowMe(ByVal objMain As Object, Optional ByVal intType As Integer) As Boolean
'参数：intType： 0-存储过程收集,1-差异校验
    Dim strSQL As String, rsData As ADODB.Recordset
        
    On Error GoTo errHand
    mblnOk = False
    Set mobjMain = objMain
    mintType = intType
    '差异对比必须数据库存在自定义存储过程与空白存储过程
    If mintType <> 0 Then
        strSQL = "Select ID,名称,所有者 From zlprocedure Where 类型 In (1,2)"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程列表")
        If rsData.EOF Then
            MsgBox "当前管理工具没有标准过程和空白过程！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
    Else
        strSQL = "Select 1 From Zlsystems Where Upper(所有者) = User"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "判断是否系统所有者")
        If rsData.RecordCount = 0 Then
            MsgBox "请以系统所有者登录进行存储过程收集！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
    End If
    Me.Show 1, mobjMain
    ShowMe = mblnOk
    Exit Function
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub cmdDel_Click()
    Call vsfMain_KeyDown(vbKeyDelete, 0)
End Sub

'==============================================================
'==控件事件
'==============================================================
Private Sub cmdExit_Click()
    '正在收集或者检查过程
    If Not cmdStart.Enabled Then
        If MsgBox("正在进行" & IIf(mintType = 0, "过程收集", "过程检查") & ",确认要退出吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdConnet_Click()
    If mfrmPageConfigure Is Nothing Then Set mfrmPageConfigure = New frmProcConfigure
    Call mfrmPageConfigure.ShowConfigure(Me)
End Sub

Private Sub cmdStart_Click()
    Dim arrTmp As Variant, lngLoop As Long, lngEnd As Long
    Dim strUpgrade As String, rsFile As ADODB.Recordset
    Dim arrSQL() As Variant
    
    mblnOk = False
    Call ShowState("(1/8)正在检查必要数据.")
    DoEvents
    If Not ValidData Then GoTo errEnd
    Call ShowState("(2/8)正在建立临时目录..")
    Call DealWithTmpFolder(True)
    Call ShowState(IIf(mintType = 0, "(3/8)正在准备数据库过程..", "(3/8)正在准备上次的标准过程.."))
    '将数据库过程生成单个脚本文件
    If Not LoadBaseProcs(App.Path & "\BaseProcedure") Then GoTo errEnd
    Call ShowState("(4/8)正在准备标准脚本过程..")
    If Not LoadComProcs(App.Path & "\ComProcedure") Then GoTo errEnd
    '将数据库中的过程与脚本进行比对，生成html报告
    Call ShowState("(5/8)正在比较..")
    If Not CompareFolder(App.Path & "\BaseProcedure", App.Path & "\ComProcedure", App.Path & "\Reports") Then
        GoTo errEnd
    End If
    Call ShowState("(6/8)正在生成数据...")
    Call CreateResultSQL(App.Path & "\BaseProcedure", App.Path & "\ComProcedure", App.Path & "\Reports", arrSQL)
    On Error GoTo errHand
    Call ShowState("(7/8)正在提交数据...")
    If Not gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, "保存结果") Then GoTo errHand
    Call ShowState("(8/8)正在清除临时数据...")
    Call DealWithTmpFolder
    Call ShowState("", True)
    mblnOk = True
    If mintType = 0 Then
        MsgBox "存储过程和函数的收集操作完成！", vbInformation, gstrSysName
    Else
        MsgBox "存储过程和函数的检查操作完成！", vbInformation, gstrSysName
    End If
    Unload Me
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errEnd:
    Call ShowState("", True)
    Exit Sub
errHand:
    Call ShowState("", True)
    MsgBox IIf(mintType = 0, "收集操作失败！", "检查操作失败！") & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Call InitFace
    Call optDB_Click(CurDB)
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    If mintType = 0 Then
        picTop.Height = cmdConnet.Top + cmdConnet.Height + 60
    Else
        picTop.Height = picFunCap.Top + picFunCap.Height + 30
    End If
    vsfMain.Move 120, picTop.Top + picTop.Height + 30
    vsfMain.Height = picBottom.Top - 30 - vsfMain.Top
    vsfMain.Width = Me.ScaleWidth - 3 * vsfMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mfrmPageConfigure Is Nothing) Then Unload mfrmPageConfigure
End Sub

Private Sub mfrmPageConfigure_AfterConn(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    Call LoadData
End Sub

Private Sub optDB_Click(Index As Integer)
    cmdConnet.Enabled = optDB(OtherDB).value
    Select Case Index
        Case 0
            Set mcnOracle = gcnOracle
            Call LoadData
        Case 1
            vsfMain.Rows = vsfMain.FixedRows
    End Select
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfMain
        .Redraw = False
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, SC_序号, .Rows - 1, SC_序号) = Color.深灰色
            .Cell(flexcpFontBold, .FixedRows, SC_序号, .Rows - 1, SC_序号) = False
        End If
        .Cell(flexcpFontBold, .Row, SC_序号, .Row, SC_序号) = True
        .Cell(flexcpForeColor, .Row, SC_序号, .Row, SC_序号) = Color.兰色
        .Redraw = True
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    Call SetSerial
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= SC_系统名称 Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = SC_安装脚本 Then
        With dlg
            .DialogTitle = "选择应用安装配置文件"
            .Filter = "(应用安装配置文件)|zlSetup.ini"
            .ShowOpen
            If .FileName = "" Then
                Exit Sub
            Else
                vsfMain.TextMatrix(Row, Col) = .FileName
            End If
        End With
    End If
End Sub

Private Sub vsfMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMain.Row >= vsfMain.FixedRows Then
            If Shift <> vbCtrlMask Then
                vsfMain.RemoveItem vsfMain.Row
                Call SetSerial
            ElseIf vsfMain.Col = SC_安装脚本 Then
                vsfMain.TextMatrix(vsfMain.Row, SC_安装脚本) = ""
            End If
        End If
    End If
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> SC_安装脚本
End Sub
'==============================================================
'==私有方法
'==============================================================
Private Sub InitFace()
'功能：初始化界面
    If mintType = 0 Then
        lblFunCap.Caption = "收集登记过程/函数"
        lblFunNote.Caption = "根据安装版本与系统版本比较，从安装脚本与升级中获取最新的标准过程。最新标准过程与数据库过程进行对比，获得经过用户调整的标准过程与用户新增的自定义过程。"
        Me.Caption = "收集登记"
    Else
         lblFunCap.Caption = "检查过程/函数差异"
         lblFunNote.Caption = "根据安装版本与系统版本比较，从安装脚本与升级中获取最新的标准过程。最新标准过程与升级前的收集到的标准过程进行对比，获得需要因标准脚本发生改变而需要调整的用户过程。"
         Me.Caption = "差异检查"
    End If
End Sub

Private Sub LoadData()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errH
    strSQL = "Select a.编号, a.版本号, a.名称 As 系统名称, b.文件名" & vbNewLine & _
                    "From Zlsystems a, Zlsysfiles b" & vbNewLine & _
                    "Where a.编号 = b.系统(+) And b.操作 = 1" & vbNewLine & _
                    "Order By Nvl(a.共享号, 0), a.编号"
                    
    Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "获取安装配置文件")
    With vsfMain
        .Redraw = flexRDNone
        .Rows = vsfMain.FixedRows
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, SC_序号) = lngRow
            .TextMatrix(lngRow, SC_版本号) = rsTmp!版本号 & ""
            .TextMatrix(lngRow, SC_系统名称) = rsTmp!系统名称 & ""
            .TextMatrix(lngRow, SC_安装脚本) = rsTmp!文件名 & ""
            .RowData(lngRow) = Val(rsTmp!编号 & "")
            rsTmp.MoveNext
        Loop
        If .Rows <> vsfMain.FixedRows Then
            vsfMain.Row = vsfMain.FixedRows
        End If
        Call vsfMain_AfterRowColChange(-1, -1, 1, 1)
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub SetSerial()
'功能：生成行序号
    Dim i As Long
    With vsfMain
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, SC_序号) = i
        Next
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, SC_序号, .Rows - 1, SC_序号) = Color.深灰色
            .Cell(flexcpFontBold, .FixedRows, SC_序号, .Rows - 1, SC_序号) = False
        End If
        If .Row > 0 Then
            .Cell(flexcpFontBold, .Row, SC_序号, .Row, SC_序号) = True
            .Cell(flexcpForeColor, .Row, SC_序号, .Row, SC_序号) = Color.兰色
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Function ValidData() As Boolean
    Dim i As Long, rsInit As ADODB.Recordset
    Dim strPath As String, strCurMax As String
    
    On Error GoTo errH
    Set mrsUpgradeFiles = Nothing
    If mcnOracle Is Nothing Then
        MsgBox "请先进行连接配置，已确认收集来源！", vbInformation + vbOKOnly, "中联软件"
        Exit Function
    End If
    pbrCollect.value = 5
    With vsfMain
        If .Rows = .FixedRows Then
            MsgBox "当前数据库没有安装任何系统。", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        pbrCollect.value = 10
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, SC_安装脚本) = "" Then
                MsgBox "请选择" & .TextMatrix(i, SC_系统名称) & "没有选择安装配置文件", vbInformation, gstrSysName
                .Row = i: .Col = SC_安装脚本: Exit Function
            End If
            If Not CheckInitFile(.RowData(i), .TextMatrix(i, SC_安装脚本), False, rsInit, False) Then
                .Row = i: .Col = SC_安装脚本: Exit Function
            End If
            rsInit.Filter = "项目='版本号'"
            .TextMatrix(i, SC_配置版本) = rsInit!内容 & ""
        Next
        pbrCollect.value = 20
        '验证升迁脚本的能否支持系统升迁
        For i = .FixedRows To .Rows - 1
            '配置文件版本比应用系统高，则无法
            If VerFull(GetPrimaryVer(.TextMatrix(i, SC_配置版本))) > VerFull(.TextMatrix(i, SC_版本号)) Then
                If MsgBox(.TextMatrix(i, SC_系统名称) & "的安装脚本版本为" & GetPrimaryVer(.TextMatrix(i, SC_配置版本)) & "，高于系统当前版本" & .TextMatrix(i, SC_版本号) & "，无法获取标准脚本。是否继续？", vbInformation + vbYesNo + vbDefaultButton2, "中联软件") = vbNo Then
                    Exit Function
                End If
            Else
                '安装脚本对应的大版本，因此使用GetPrimaryVer(.TextMatrix(i, SC_配置版本))，传入当前系统版本.TextMatrix(i, SC_版本号)，获取从安装脚本版本可升级到当前系统版本的所有升级脚本，以及这些脚本支持的最大版本（若不缺失脚本，应与当前系统版本相同）。
                Set mrsUpgradeFiles = GetUpgradeFiles(mrsUpgradeFiles, .RowData(i), GetPrimaryVer(.TextMatrix(i, SC_配置版本)), .TextMatrix(i, SC_安装脚本), , , .TextMatrix(i, SC_版本号), strCurMax, , True)
                '没有获取到脚本，且安装脚本版本与系统当前脚本不相同，则提示脚本缺失
                If strCurMax = "" And GetPrimaryVer(.TextMatrix(i, SC_配置版本)) <> .TextMatrix(i, SC_版本号) Then
                    MsgBox .TextMatrix(i, SC_系统名称) & "缺失从" & GetPrimaryVer(.TextMatrix(i, SC_配置版本)) & "升迁到" & .TextMatrix(i, SC_版本号) & "的升级脚本。", vbInformation + vbOKOnly, "中联软件"
                    Exit Function
                '获取到脚本，但是不小于当前系统版本的升级脚本无法支持系统升级到当前系统版本，则提示脚本缺失
                ElseIf strCurMax <> "" And strCurMax <> .TextMatrix(i, SC_版本号) Then
                    MsgBox .TextMatrix(i, SC_系统名称) & "缺失从" & GetPrimaryVer(.TextMatrix(i, SC_配置版本)) & "升迁到" & .TextMatrix(i, SC_版本号) & "的升级脚本。", vbInformation + vbOKOnly, "中联软件"
                    Exit Function
                End If
            End If
            pbrCollect.value = 20 + (i / (.Rows - 1)) * 70
        Next
        If mrsUpgradeFiles Is Nothing Then
            MsgBox "未收集到必要脚本，请更新正确的安装脚本与升级脚本。", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        '只保留应用系统的标准升级脚本
        Call RecDelete(mrsUpgradeFiles, "(SysType<>" & ST_App & ") OR (SysType = " & ST_App & " And FileType<>" & FT_Standard & ")")
    End With
    pbrCollect.value = 100
    ValidData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub DealWithTmpFolder(Optional ByVal blnCreate As Boolean)
'功能：处理临时目录
    '转换为大写的脚本
    If gobjFile.FolderExists(App.Path & "\BaseProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\BaseProcedure", True)
    pbrCollect.value = 16 * IIf(blnCreate, 1, 2)
    If gobjFile.FolderExists(App.Path & "\ComProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\ComProcedure")
    pbrCollect.value = 16 * IIf(blnCreate, 1, 2) * 2
    '对比报告
    If gobjFile.FolderExists(App.Path & "\Reports") And blnCreate Then gobjFile.DeleteFolder (App.Path & "\Reports")
    pbrCollect.value = 50 * IIf(blnCreate, 1, 2)
    If blnCreate Then
        Call gobjFile.CreateFolder(App.Path & "\BaseProcedure")
        pbrCollect.value = 16.5 * 4
        Call gobjFile.CreateFolder(App.Path & "\ComProcedure")
        pbrCollect.value = 16.5 * 5
        Call gobjFile.CreateFolder(App.Path & "\Reports")
        pbrCollect.value = 100
    End If
End Sub

Private Sub ShowState(Optional ByVal strInfo As String, Optional ByVal blnEnd As Boolean)
    lblTitle.Caption = strInfo
    lblTitle.Tag = strInfo
    lblTitle.Visible = strInfo <> ""
    pbrCollect.Visible = strInfo <> ""
    pbrCollect.value = 0
    cmdStart.Enabled = blnEnd
    cmdDel.Enabled = blnEnd
End Sub

Private Function LoadBaseProcs(ByVal strPath As String) As Boolean
    '功能：加载数据库存储过程
    Dim rsSource As ADODB.Recordset, strSQL As String
    Dim objText As TextStream, strProcName As String, strProcText As String
    Dim objPercent As New clsPercent
    
    On Error GoTo errH
    '存储过程收集，收集数据库作为基本存储过程
    If mintType = 0 Then
        strSQL = "Select Name, Type, Text, Line 序号 From User_Source Where Type In ('PROCEDURE', 'FUNCTION') Order By Name, Line"
        Set rsSource = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "获取数据库过程源码")
    '差异对比，以上次标准存储过程为基本存储过程
    Else
        strSQL = "Select a.Id, Upper(a.名称) Name, b.序号, b.内容 Text" & vbNewLine & _
                        "From Zlprocedure a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.过程id And b.性质 = " & ProcTextType.本次标准过程 & " And a.类型 In (1, 2)" & vbNewLine & _
                        "Order By a.Id, b.序号"
        Set rsSource = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "获取数据库过程源码")
    End If
    If Not rsSource.EOF Then
        pbrCollect.Visible = True
        Call objPercent.InitPercent(pbrCollect, rsSource.RecordCount)
        Do While Not rsSource.EOF
            If strProcName <> rsSource!name & "" Then
                If strProcName <> "" Then
                    '数据库源码没有CREATE OR REPLACE
                    If mintType = 0 Then
                        strProcText = "CREATE OR REPLACE " & strProcText
                    End If
                    '创建单个过程脚本文件
                    Set objText = gobjFile.CreateTextFile(strPath & "\" & strProcName & ".sql", True)
                    objText.Write strProcText
                End If
                strProcName = rsSource!name & ""
                strProcText = ""
            End If
            If rsSource!序号 = 1 Then
                '名称带双引号，则去掉
                If UCase(rsSource!Text) Like "*" & """" & UCase(strProcName) & """" & "*" Then
                    strProcText = strProcText & Replace(UCase(rsSource!Text), """" & UCase(strProcName) & """", strProcName)
                Else
                    strProcText = strProcText & rsSource!Text
                End If
            Else
                strProcText = strProcText & rsSource!Text
            End If
            rsSource.MoveNext
            Call objPercent.LoopPercent
        Loop
        If strProcName <> "" Then
            '创建单个过程脚本文件
            Set objText = gobjFile.CreateTextFile(strPath & "\" & strProcName & ".sql", True)
            objText.Write strProcText
        End If
        objText.Close
        pbrCollect.Visible = False
    End If
    LoadBaseProcs = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function LoadComProcs(ByVal strPath As String) As Boolean
'功能：加载数据库存储过程
    Dim i As Long, strFile As String
    Dim objPercent As New clsPercent
    On Error GoTo errH
    With vsfMain
        mrsUpgradeFiles.Filter = ""
        mrsUpgradeFiles.Sort = "系统编号,FullSPVer"
        Call objPercent.InitPercent(pbrCollect, mrsUpgradeFiles.RecordCount + .Rows - 1)
        For i = .FixedRows To .Rows - 1
            lblTitle.Caption = lblTitle.Tag & "    正在提取“" & .TextMatrix(i, SC_系统名称) & "”安装脚本.."
            strFile = gobjFile.GetParentFolderName(.TextMatrix(i, SC_安装脚本)) & "\zlProgram.sql"
            Call LoadProcFile(strFile, strPath)
            objPercent.LoopPercent
            mrsUpgradeFiles.Filter = "系统编号=" & .RowData(i)
            mrsUpgradeFiles.Sort = "FullSPVer"
            lblTitle.Caption = lblTitle.Tag & "    正在提取" & .TextMatrix(i, SC_系统名称) & "升级脚本.."
            Do While Not mrsUpgradeFiles.EOF
                Call LoadProcFile(mrsUpgradeFiles!FilePath, strPath)
                objPercent.LoopPercent
                mrsUpgradeFiles.MoveNext
            Loop
        Next
    End With
    LoadComProcs = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function LoadProcFile(ByVal strFile As String, ByVal strFilePath As String) As Boolean
'功能：讲脚本中存在的存储过程与函数分别拆分为文件存储。
    Dim objScript As New clsRunScript
    Dim objText As TextStream
    Dim objPercent As New clsPercent
    
    With objScript
         If .OpenFile(strFile) And Not .EOF Then
            Do While Not .EOF
                If .SQLInfo.Block = True Then
                    If .SQLInfo.BlockType = "PROCEDURE" Or .SQLInfo.BlockType = "FUNCTION" Then
                        Set objText = gobjFile.CreateTextFile(strFilePath & "\" & .SQLInfo.BlockName & ".sql", True)
                        objText.Write .SQLInfo.SQL
                        objText.Close
                    End If
                End If
                Call .ReadNextSQL
            Loop
         End If
    End With
    LoadProcFile = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub CreateResultSQL(ByVal strBasePath As String, strComPath As String, ByVal strRportPth As String, ByRef arrSQL As Variant)
    Dim objFolder As Folder, objFile As File
    Dim strFileName As String
    Dim rsObjectInfo As ADODB.Recordset, strSQL As String
    Dim lngKey As Long, pt As ProcType, strNote As String, strOwner As String
    Dim objPercent As New clsPercent
    Dim rsSouce As ADODB.Recordset
    Dim strTMp As String
    Dim objLog As TextStream
    
    On Error GoTo errH
    If mintType = 0 Then
        '报告中存在的即为需要调整的过程
        lblTitle.Caption = lblTitle.Tag & "    正在产生变动过程.."
        Call objPercent.InitPercent(pbrCollect, gobjFile.GetFolder(strRportPth).Files.Count + gobjFile.GetFolder(strBasePath).Files.Count)
        Set objFolder = gobjFile.GetFolder(strRportPth)
        strSQL = "Select b.Owner, b.Object_Name, c.Id, c.类型, c.名称, c.状态, c.说明, c.修改人员, c.修改时间, c.上次修改人员, c.上次修改时间" & vbNewLine & _
                        "From (Select a.Owner, a.Object_Name" & vbNewLine & _
                        "       From All_Objects a" & vbNewLine & _
                        "       Where a.Object_Type In ('PROCEDURE', 'FUNCTION') And a.Owner In (Select Distinct 所有者 From Zlsystems)) b," & vbNewLine & _
                        "     Zlprocedure c" & vbNewLine & _
                        "Where b.Object_Name = Upper(c.名称(+)) " & vbNewLine & _
                        "Order by Object_Name,Owner,Id"
        Set rsObjectInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取对象信息")
        
        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Manage(0)")
        For Each objFile In objFolder.Files
            strFileName = Split(objFile.name, ".")(0): strOwner = "": lngKey = 0: pt = ProcType.变动过程
            rsObjectInfo.Filter = "Object_Name='" & UCase(strFileName) & "'"
            If Not rsObjectInfo.EOF Then
                lblTitle.Caption = lblTitle.Tag & "    正在产生变动过程：" & strFileName
                strOwner = rsObjectInfo!Owner & "": lngKey = Val(rsObjectInfo!Id & "")
                '过程不存在，自动添加为变动过程或空白过程
                If lngKey = 0 Then
                    lngKey = gclsBase.GetNextId("zlProcedure")
                ElseIf Val(rsObjectInfo!类型 & "") = 2 Then
                    pt = ProcType.空白过程
                End If
                Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & pt & ",'" & strFileName & "'," & ProcState.待检查 & ",'" & rsObjectInfo!说明 & "','" & strOwner & "')")
                '保存本次自定义过程
                Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次自定过程, strBasePath & "\" & strFileName & ".sql", arrSQL, True)
                '保存本次标准过程
                Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次标准过程, strComPath & "\" & strFileName & ".sql", arrSQL, True)
            End If
            objPercent.LoopPercent
        Next
        
        lblTitle.Caption = lblTitle.Tag & "     正在产生用户过程.."
        Set objFolder = gobjFile.GetFolder(strBasePath)
        For Each objFile In objFolder.Files
            '数据库中的过程在脚本中没有，说明是用户过程
            If Not gobjFile.FileExists(strComPath & "\" & objFile.name) Then
                strFileName = Split(objFile.name, ".")(0)
                If Not UCase(strFileName) Like "ZL*_UPGRADECHECK" Then '升迁检查函数排除
                    strOwner = "": lngKey = 0: pt = ProcType.用户过程
                    rsObjectInfo.Filter = "Object_Name='" & UCase(strFileName) & "'"
                    If Not rsObjectInfo.EOF Then
                        lblTitle.Caption = lblTitle.Tag & "    正在产生用户过程：" & strFileName
                        strOwner = rsObjectInfo!Owner & "": lngKey = Val(rsObjectInfo!Id & "")
                        If lngKey = 0 Then '过程不存在，自动添加为用户过程
                            lngKey = gclsBase.GetNextId("zlProcedure")
                        End If
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & pt & ",'" & strFileName & "'," & ProcState.待检查 & ",'" & rsObjectInfo!说明 & "','" & strOwner & "')")
                        '保存本次自定义过程
                        Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次自定过程, strBasePath & "\" & objFile.name, arrSQL, True)
                    End If
                End If
            End If
            Call objPercent.LoopPercent
        Next
        '测试环境，收集缺失的函数
        If gblnInIDE Then
            If Not gobjFile.FolderExists("C:\AppSoft\Log\过程管理\") Then
                gobjFile.CreateFolder ("C:\AppSoft\Log\过程管理\")
            End If
            Set objLog = gobjFile.CreateTextFile("C:\AppSoft\Log\过程管理\Proc_缺失过程.ini", True)
            Set objFolder = gobjFile.GetFolder(strComPath)
            For Each objFile In objFolder.Files
                If Not gobjFile.FileExists(strBasePath & "\" & objFile.name) Then
                    objLog.WriteLine Split(objFile.name, ".")(0)
                End If
            Next
        End If
    Else
        lblTitle.Caption = lblTitle.Tag & "     正在调整过程状态.."
        strSQL = "Select ID,名称,类型,说明,所有者 From zlprocedure"
        Set rsObjectInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程列表")
        Call objPercent.InitPercent(pbrCollect, rsObjectInfo.RecordCount)
        Do While Not rsObjectInfo.EOF
            strFileName = rsObjectInfo!名称 & ""
            If Val(rsObjectInfo!类型 & "") <> ProcType.用户过程 Then
                pt = Val(rsObjectInfo!类型 & "")
                If gobjFile.FileExists(strComPath & "\" & strFileName & ".sql") Then
                    '标准过程在升级前后有变化
                    If gobjFile.FileExists(strRportPth & "\" & strFileName & ".sql.htm") Then
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & rsObjectInfo!Id & "," & pt & ",'" & strFileName & "'," & ProcState.待调整 & ",'" & rsObjectInfo!说明 & "','" & rsObjectInfo!所有者 & "')")
                    '标准过程在升级前后无变化
                    Else
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & rsObjectInfo!Id & "," & pt & ",'" & strFileName & "'," & ProcState.无变化 & ",'" & rsObjectInfo!说明 & "','" & rsObjectInfo!所有者 & "')")
                    End If
                    '保存本次标准过程
                    Call gclsBase.GetProcSQL(Val(rsObjectInfo!Id & ""), ProcTextType.本次标准过程, strComPath & "\" & strFileName & ".sql", arrSQL, True)
                End If
            Else '用户过程，调整为变动过程
                If gobjFile.FileExists(strComPath & "\" & strFileName & ".sql") Then '用户过程存在于标准脚本中则应自动调整为变动过程
                    pt = ProcType.变动过程: strOwner = ""
                    strTMp = gclsBase.GetProgram(strFileName, strOwner)
                    If strTMp = "" Then
                        strOwner = rsObjectInfo!所有者 & ""
                    End If
                    Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & Val(rsObjectInfo!Id & "") & "," & pt & ",'" & strFileName & "'," & ProcState.调整中 & ",'" & rsObjectInfo!说明 & "','" & strOwner & "')")
                    If strTMp <> "" Then
                        Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次自定过程, strTMp, arrSQL)
                    End If
                    '保存本次标准过程
                    Call gclsBase.GetProcSQL(Val(rsObjectInfo!Id & ""), ProcTextType.本次标准过程, strComPath & "\" & strFileName & ".sql", arrSQL, True)
                End If
            End If
            Call objPercent.LoopPercent
            rsObjectInfo.MoveNext
        Loop
        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Manage(1)")
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

