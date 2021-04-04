VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmEPRAuditOutline 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3810
      Index           =   2
      Left            =   525
      ScaleHeight     =   3810
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   1140
      Width           =   4500
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   270
         Width           =   3405
         _cx             =   6006
         _cy             =   5265
         Appearance      =   2
         BorderStyle     =   1
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
         WallPaper       =   "frmEPRAuditOutline.frx":0000
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRAuditOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mlngMoual As Long
Private mblnShowAll As Boolean

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object, ByVal lngMoual As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngMoual = lngMoual
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    
End Function


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call RptPrint(2)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call RptPrint(1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call RptPrint(3)
        
    End Select
    
End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With vfgThis
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               '预览,打印,输出到Excel
        
            Control.Enabled = (.Rows > .FixedRows + 1)
        
        End Select
        
    End With
    
End Sub

Public Function zlRefreshData(ByVal intKind As Integer, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    '******************************************************************************************************************
    '功能:根据审查范围组织显示审查数据
    '******************************************************************************************************************
    Dim lngTotal As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    On Error GoTo errHand
    
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '门诊病历
        
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, P.门诊人次, W.门诊完成, P.急诊人次, W.急诊完成" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select 执行部门id, Sum(Decode(急诊, 1, 0, 1)) As 门诊人次, Sum(Decode(急诊, 1, 1, 0)) As 急诊人次" & vbNewLine & _
                "        From 病人挂号记录" & vbNewLine & _
                "        Where Nvl(执行状态, 0) <> 0 And 记录性质=1 And 记录状态=1 And 登记时间 Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "              To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 执行部门id) P," & vbNewLine & _
                "      (Select W.科室id, Sum(W.已完成) As 已完成, Sum(W.在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.事件, '门诊', W.已完成, Null)) As 门诊完成," & vbNewLine & _
                "               Sum(Decode(F.事件, '急诊', W.已完成, Null)) As 急诊完成" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 1) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 1 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "                     To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '临床' And M.服务对象 In (1, 3) And D.ID = P.执行部门id(+) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) And" & vbNewLine & _
                "       D.ID = W.科室id(+)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "科室": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "门诊": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "急诊": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case 2  '住院病历
        
        
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, I.入院人次, W.入院病历, E.转入人次, W.转入病历, O.出院人次," & vbNewLine & _
                "        W.出院病历, O.死亡人次, W.死亡病历, G.转出人次, W.转出病历, S.手术人次, W.手术病历" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select W.科室id, Sum(已完成) As 已完成, Sum(在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '入院', 1, '首次入院', 1, '再次入院', 1, 0), 0) * 已完成) As 入院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 0, 1), 0), 0) * 已完成) As 转入病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '出院', 1, '24小时出院', 1, 0), 0) * 已完成) As 出院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '死亡', 1, '24小时死亡', 1, 0), 0) * 已完成) As 死亡病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 1, 0), 0), 0) * 已完成) As 转出病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '手术', 1, 0), 0) * 已完成) As 手术病历" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 2) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 2 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W," & vbNewLine
        strSQL = strSQL & "      (Select 入院科室id, Count(*) As 入院人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 入院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 入院科室id) I," & vbNewLine & _
                "      (Select 科室id, Count(*) As 转入人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 开始原因 = 3 And 开始时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 科室id) E," & vbNewLine & _
                "      (Select 出院科室id, Sum(Decode(出院方式, '死亡', 0, 1)) As 出院人次, Sum(Decode(出院方式, '死亡', 1, 0)) As 死亡人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 出院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 出院科室id) O," & vbNewLine & _
                "      (Select 科室id, Count(*) As 转出人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 终止原因 = 3 And 终止时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 科室id) G," & vbNewLine & _
                "      (Select R.执行科室id, Count(*) As 手术人次" & vbNewLine & _
                "        From 病人医嘱记录 R" & vbNewLine & _
                "        Where R.诊疗类别 = 'F' And R.手术时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By R.执行科室id) S" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '临床' And 服务对象 In (2, 3) And D.ID = W.科室id(+) And D.ID = I.入院科室id(+) And" & vbNewLine & _
                "       D.ID = E.科室id(+) And D.ID = O.出院科室id(+) And D.ID = G.科室id(+) And D.ID = S.执行科室id(+) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "科室": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "入院": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "转入": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "出院": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "死亡": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "转出": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            .TextMatrix(0, 15) = "手术": .TextMatrix(0, 16) = .TextMatrix(0, 15)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
            .TextMatrix(1, 9) = "人次": .TextMatrix(1, 10) = "完成病历"
            .TextMatrix(1, 11) = "人次": .TextMatrix(1, 12) = "完成病历"
            .TextMatrix(1, 13) = "人次": .TextMatrix(1, 14) = "完成病历"
            .TextMatrix(1, 15) = "人次": .TextMatrix(1, 16) = "完成病历"
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '护理病历
        strSQL = "Select D.ID, D.编码, D.名称, W.已完成, W.在书写, I.入院人次, W.入院病历, E.转入人次, W.转入病历, O.出院人次," & vbNewLine & _
                "        W.出院病历, O.死亡人次, W.死亡病历, G.转出人次, W.转出病历" & vbNewLine & _
                " From 部门表 D, 部门性质说明 M," & vbNewLine & _
                "      (Select W.科室id, Sum(已完成) As 已完成, Sum(在书写) As 在书写," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '入院', 1, '首次入院', 1, '再次入院', 1, 0), 0) * 已完成) As 入院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 0, 1), 0), 0) * 已完成) As 转入病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '出院', 1, '24小时出院', 1, 0), 0) * 已完成) As 出院病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '死亡', 1, '24小时死亡', 1, 0), 0) * 已完成) As 死亡病历," & vbNewLine & _
                "               Sum(Decode(F.唯一, 1, Decode(F.事件, '转科', Decode(Sign(F.书写时限), -1, 1, 0), 0), 0) * 已完成) As 转出病历" & vbNewLine & _
                "        From (Select F.ID, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "               From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "               Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 4) F," & vbNewLine & _
                "             (Select 科室id, 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成," & vbNewLine & _
                "                      Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "               From 电子病历记录" & vbNewLine & _
                "               Where 病历种类 = 4 And 创建时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By 科室id, 文件id) W" & vbNewLine & _
                "        Where F.ID = W.文件id And (F.通用 = 1 Or F.通用 = 2 And F.科室id = W.科室id)" & vbNewLine & _
                "        Group By W.科室id) W," & vbNewLine
        strSQL = strSQL & "      (Select 入院病区id, Count(*) As 入院人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 入院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 入院病区id) I," & vbNewLine & _
                "      (Select 病区id, Count(*) As 转入人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 开始原因 = 3 And 开始时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 病区id) E," & vbNewLine & _
                "      (Select 当前病区id, Sum(Decode(出院方式, '死亡', 0, 1)) As 出院人次, Sum(Decode(出院方式, '死亡', 1, 0)) As 死亡人次" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 出院日期 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 当前病区id) O," & vbNewLine & _
                "      (Select 病区id, Count(*) As 转出人次" & vbNewLine & _
                "        From 病人变动记录" & vbNewLine & _
                "        Where 终止原因 = 3 And 终止时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By 病区id) G" & vbNewLine & _
                " Where D.ID = M.部门id And M.工作性质 = '护理' And 服务对象 In (2, 3) And D.ID = W.科室id(+) And D.ID = I.入院病区id(+) And" & vbNewLine & _
                "       D.ID = E.病区id(+) And D.ID = O.当前病区id(+) And D.ID = G.病区id(+) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null)" & vbNewLine & _
                " Order By D.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "病区": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "病历书写情况": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "入院": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "转入": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "出院": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "死亡": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "转出": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            
            .TextMatrix(1, 1) = "编码": .TextMatrix(1, 2) = "名称"
            .TextMatrix(1, 3) = "已完成": .TextMatrix(1, 4) = "在书写"
            .TextMatrix(1, 5) = "人次": .TextMatrix(1, 6) = "完成病历"
            .TextMatrix(1, 7) = "人次": .TextMatrix(1, 8) = "完成病历"
            .TextMatrix(1, 9) = "人次": .TextMatrix(1, 10) = "完成病历"
            .TextMatrix(1, 11) = "人次": .TextMatrix(1, 12) = "完成病历"
            .TextMatrix(1, 13) = "人次": .TextMatrix(1, 14) = "完成病历"
        End With
    End Select
    
    
    '求合计
    '------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnData As Boolean
    
    With Me.vfgThis
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "合计"
        For lngCol = 3 To .Cols - 1
            lngTotal = 0
            For lngRow = .FixedRows To .Rows - 2
                lngTotal = lngTotal + Val(.TextMatrix(lngRow, lngCol))
            Next
            .TextMatrix(.Rows - 1, lngCol) = lngTotal
        Next
        .Row = .FixedRows: .Col = 1
        Call .AutoSize(1, .Cols - 1)
        
        .Redraw = False
        If mblnShowAll Then
            For lngRow = .FixedRows To .Rows - 2
                .RowHeight(lngRow) = .RowHeightMin
                .RowHidden(lngRow) = False
            Next
        Else
            For lngRow = .FixedRows To .Rows - 2
                blnData = False
                For lngCol = 3 To .Cols - 1
                    If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                Next
                If blnData = False Then
                    .RowHeight(lngRow) = 0
                    .RowHidden(lngRow) = True
                End If
            Next
        End If
        .Redraw = True
    End With
    
    '显示或隐藏空行
'    Call chkNoData_Click
'    Me.stbThis.Panels(2).Text = "点击“展开(Ctrl+O)”详细审查当前科室病人病历情况或病历分类书写情况…"
    
    If Me.Visible Then Me.vfgThis.SetFocus
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsThis)
    Set cbsThis.Icons = frmPubResource.imgApp.Icons
    cbsThis.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = False
    
            
    '部门工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 10, "显示无业务科室", , , xtpButtonIconAndCaption)
    objControl.Checked = True
        
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = vfgThis
    
    Select Case mintKind
    Case 1
        objPrint.Title.Text = "门诊病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
    Case 2
        objPrint.Title.Text = "住院病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
    Case 4
        objPrint.Title.Text = "护理病历(" & mstrDateFrom & "至" & mstrDateTo & ")书写情况"
    End Select
    
    Set objPrint.Title.Font = vfgThis.Font

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
                
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                        
        
        If Val(zlDatabase.GetPara("显示无业务科室", glngSys, mlngMoual, "1")) = 1 Then
            mblnShowAll = True
        Else
            mblnShowAll = False
        End If


    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        

    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 10
        
        mblnShowAll = Not mblnShowAll
        Control.Checked = mblnShowAll
        
    
        Dim blnData As Boolean
        Dim lngRow As Long
        Dim lngCol As Long
        
        
        With vfgThis
            .Redraw = False
            If mblnShowAll Then
                For lngRow = .FixedRows To .Rows - 2
                    .RowHeight(lngRow) = .RowHeightMin
                    .RowHidden(lngRow) = False
                Next
            Else
                For lngRow = .FixedRows To .Rows - 2
                    blnData = False
                    For lngCol = 3 To .Cols - 1
                        If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                    Next
                    If blnData = False Then
                        .RowHeight(lngRow) = 0
                        .RowHidden(lngRow) = True
                    End If
                Next
            End If
            .Redraw = True
        End With


    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long
    Dim lngScaleTop  As Long
    Dim lngScaleRight  As Long
    Dim lngScaleBottom  As Long
    
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    picPane(2).Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 10
        
        Control.IconId = IIf(mblnShowAll, 12, 10)
        
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SetPara("显示无业务科室", IIf(mblnShowAll, 1, 0), mlngMoual)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        vfgThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub
