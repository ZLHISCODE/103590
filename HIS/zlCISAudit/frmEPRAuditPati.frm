VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmEPRAuditPati 
   BorderStyle     =   0  'None
   Caption         =   "病人病历书写审查"
   ClientHeight    =   6930
   ClientLeft      =   -60
   ClientTop       =   15
   ClientWidth     =   10455
   Icon            =   "frmEPRAuditPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   2745
      ScaleHeight     =   3135
      ScaleWidth      =   5580
      TabIndex        =   4
      Top             =   225
      Width           =   5580
      Begin VSFlex8Ctl.VSFlexGrid vfgPati 
         Height          =   2640
         Left            =   300
         TabIndex        =   5
         Top             =   270
         Width           =   4320
         _cx             =   7620
         _cy             =   4657
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   5175
      Width           =   1905
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   -45
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1320
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   285
      ScaleHeight     =   240
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   1830
      Width           =   1350
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   -30
         Width           =   930
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditPati.frx":6852
      Left            =   525
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'常量
'----------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    标志 = 0: 事件缘由: 应写病历: 监测点: 基点时间: 要求时间: 完成时间: 完成记录id: 当前时间: 备注说明
End Enum

Private Const conPane_Pati = 1
Private Const conPane_Audit = 2
Private Const conPane_Word = 3

'变量
'----------------------------------------------------------------------------------------------------------------------
Private mlngDeptId As Long      '科室id
Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mstrEvent As String     '病人事件范围
Private WithEvents mclsDockAduit As zlRichEPR.clsDockAduits
Attribute mclsDockAduit.VB_VarHelpID = -1
Private mfrmMain As Object
Private mblnReading As Boolean
Private mclsPati As clsVsf

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
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
    With vfgPati
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               '预览,打印,输出到Excel
            If ActiveControl Is Nothing Then Exit Sub
            If Me.ActiveControl.Name = .Name Then
                Control.Enabled = (.Rows > .FixedRows)
            Else
                Control.Enabled = (.Rows > .FixedRows)
            End If
        
        End Select
        
    End With
    
End Sub

Public Function zlRefreshData(ByVal intKind As Integer, ByVal strDateFrom As String, ByVal strDateTo As String, Optional blnShow As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldValue As String
    On Error GoTo errH
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    
    Call ExecuteCommand("初始数据")

    Select Case mintKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '门诊病历
        
        strSQL = "Select D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 = '临床' And M.服务对象 In (1, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2  '住院病历
        
        strSQL = "Select D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 = '临床' And M.服务对象 In (2, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '护理病历
    
        strSQL = "Select D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 = '临床' And M.服务对象 In (2, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"

    End Select
    
    strOldValue = cboDept.Text
    cboDept.Clear
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Call cboDept.AddItem(rs("名称").Value)
            cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
            rs.MoveNext
        Loop
    End If
    
    mblnReading = True
    If Len(strOldValue) > 0 Then
        cboDept.Text = strOldValue
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
    Else
        cboDept.ListIndex = 0
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
    End If
    mblnReading = False
    
    If blnShow = False Then
        zlRefreshData = RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    Exit Function
errH:
    If Err.Number = 383 Then
        Err.Clear
        cboDept.ListIndex = 0
        Resume Next
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

'######################################################################################################################

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
    '------------------------------------------------------------------------------------------------------------------
            
            
    '部门工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlLabel, 7, "科室：", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(3).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "类型：", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(2).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 9, "查阅病历...", True, , xtpButtonIconAndCaption)
    
End Function


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
    
    mblnReading = True
    
    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
                
        Call InitGrid
        Call InitCommandBar
        
        '划分停靠区域
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMan.CreatePane(conPane_Pati, 300, 400, DockLeftOf, Nothing): objPane.Title = "病人列表": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Audit, 700, 100, DockRightOf, objPane): objPane.Title = "时限审查": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Word, 700, 300, DockBottomOf, objPane): objPane.Title = "內容监测": objPane.Options = PaneNoCaption
        
        dkpMan.SetCommandBars cbsThis
        Call DockPannelInit(dkpMan)
        
        
                            
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                        
        With cboType
            .Clear
            .AddItem "所有"
            
            If mintKind = 1 Then
                .AddItem "门诊"
                .AddItem "急诊"
            Else
                .AddItem "入院"
                .AddItem "转入"
                .AddItem "出院"
                .AddItem "死亡"
                .AddItem "转出"
                .AddItem "手术"
            End If
            
            .ListIndex = 0
            mstrEvent = "所有"
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "时限监测"
        
        strSQL = "Zl_病历时限监测_Neaten(" & Val(varParam(0)) & "," & Val(varParam(1)) & "," & mintKind & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, mfrmMain.Caption)
        
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
    mblnReading = False
End Function

Private Function InitGrid(Optional ByVal strType As String = "2") As Boolean
    Set mclsPati = New clsVsf
    With mclsPati
        Call .Initialize(Me.Controls, vfgPati, False, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[路径]", False)
        Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        If InStr(strType, "1") > 0 Then
            Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Else
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        End If
        If strType = "1" Then
            Call .AppendColumn("门诊号", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Else
            Call .AppendColumn("住院号", 900, flexAlignLeftCenter, flexDTString, "", , True)
        End If
        Call .AppendColumn("姓名", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("性别", 990, flexAlignLeftCenter, flexDTString, "", , True)
        Select Case strType
            Case "2"
                Call .AppendColumn("入院时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "入院时间", True)
            Case "1"
                Call .AppendColumn("就诊时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "就诊时间", True)
            Case "21"
                Call .AppendColumn("转入时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "转入时间", True)
            Case "22"
                Call .AppendColumn("出院日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "出院日期", True)
            Case "23"
                Call .AppendColumn("死亡日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "死亡日期", True)
            Case "24"
                Call .AppendColumn("转出时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "转出时间", True)
            Case "25"
                Call .AppendColumn("手术时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "手术时间", True)
            Case Else
                Call .AppendColumn("入院时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "入院时间", True)
        End Select
        If strType = "1" Then
            Call .AppendColumn("医生", 900, flexAlignLeftCenter, flexDTString, "", , True)
        End If
    End With
End Function
Private Function RefreshData(ByVal intKind As Integer, ByVal strEvent As String, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    Select Case intKind
    Case 1
        Call InitGrid("1")
        Select Case strEvent
        Case "门诊"
            strSQL = "Select Null as 路径, 病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And Nvl(急诊, 0) <> 1 And 记录性质=1 And 记录状态=1 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        Case "急诊"
            strSQL = "Select Null as 路径, 病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And Nvl(急诊, 0) = 1 And 记录性质=1 And 记录状态=1 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        Case Else
            strSQL = "Select Null as 路径,病人id, ID, 门诊号, 姓名, 性别, To_Char(执行时间, 'yyyy-mm-dd hh24:mi') As 就诊时间, 执行人 As 医生" & vbNewLine & _
                    "From 病人挂号记录" & vbNewLine & _
                    "Where 执行部门id + 0 = [1] And Nvl(执行状态, 0) <> 0 And 记录性质=1 And 记录状态=1 And" & vbNewLine & _
                    "      登记时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By 执行时间"
        End Select
    Case 2
        Select Case strEvent
        Case "入院"
            Call InitGrid("2")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.入院时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select 病人id, 主页id, To_Char(Max(开始时间), 'yyyy-mm-dd hh24:mi') As 入院时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 开始原因 In (1, 2, 9) And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By 病人id, 主页id) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.入院时间"
        Case "转入"
            Call InitGrid("21")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转入时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(开始时间, 'yyyy-mm-dd hh24:mi') As 转入时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 开始原因 = 3 And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转入时间"
        Case "出院"
            Call InitGrid("22")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 出院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.出院科室id + 0 = [1] And P.出院方式 <> '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "死亡"
            Call InitGrid("23")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 死亡日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.出院科室id + 0 = [1] And P.出院方式 = '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "转出"
            Call InitGrid("24")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转出时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(终止时间, 'yyyy-mm-dd hh24:mi') As 转出时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id + 0 = [1] And 终止原因 = 3 And 终止时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转出时间"
        Case "手术"
            Call InitGrid("25")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, 手术时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select R.病人id, R.主页id, To_Char(S.首次时间, 'yyyy-mm-dd hh24:mi') As 手术时间" & vbNewLine & _
                    "       From 病人医嘱记录 R, 病人医嘱发送 S" & vbNewLine & _
                    "       Where R.ID = S.医嘱id And R.诊疗类别 = 'F' And R.相关id Is Null And R.医嘱期效 = 1 And" & vbNewLine & _
                    "             (R.医嘱状态 = 8 Or R.医嘱状态 = 9) And R.病人科室id + 0 = [1] And" & vbNewLine & _
                    "             S.首次时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.手术时间"
        Case Else
            Call InitGrid
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, P.入院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 科室id = [1] And" & vbNewLine & _
                    "             (开始原因 In (1, 2, 3, 9) And 开始时间 between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             终止原因 In (1, 3, 10) And (终止时间 between To_Date([2], 'yyyy-mm-dd') and To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or 终止时间 Is Null))) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By P.入院日期"
                    
        End Select
    Case 4
        Select Case strEvent
        Case "入院"
            Call InitGrid("2")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.入院时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select 病人id, 主页id, To_Char(Max(开始时间), 'yyyy-mm-dd hh24:mi') As 入院时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 开始原因 In (1, 2, 9) And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By 病人id, 主页id) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.入院时间"
        Case "转入"
            Call InitGrid("21")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转入时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(开始时间, 'yyyy-mm-dd hh24:mi') As 转入时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 开始原因 = 3 And 开始时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转入时间"
        Case "出院"
            Call InitGrid("22")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 出院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.当前病区id + 0 = [1] And P.出院方式 <> '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "死亡"
            Call InitGrid("23")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, To_Char(P.出院日期, 'yyyy-mm-dd hh24:mi') As 死亡日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.当前病区id + 0 = [1] And P.出院方式 = '死亡' And" & vbNewLine & _
                    "      P.出院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.出院日期"
        Case "转出"
            Call InitGrid("24")
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, L.转出时间" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id, To_Char(终止时间, 'yyyy-mm-dd hh24:mi') As 转出时间" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id + 0 = [1] And 终止原因 = 3 And 终止时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By L.转出时间"
        Case Else
            Call InitGrid
            strSQL = "Select Decode(P.路径状态,Null,'',0,'','lujin') as 路径,P.病人id, P.主页id, P.住院号, I.姓名, I.性别, P.入院日期" & vbNewLine & _
                    "From 病人信息 I, 病案主页 P," & vbNewLine & _
                    "     (Select Distinct 病人id, 主页id" & vbNewLine & _
                    "       From 病人变动记录" & vbNewLine & _
                    "       Where 病区id = [1] And" & vbNewLine & _
                    "             (开始原因 In (1, 2, 3, 9) And 开始时间 between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             终止原因 In (1, 3, 10) And (终止时间 between To_Date([2], 'yyyy-mm-dd') and To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or 终止时间 Is Null))) L" & vbNewLine & _
                    "Where I.病人id = P.病人id And P.病人id = L.病人id And P.主页id = L.主页id" & vbNewLine & _
                    "Order By P.入院日期"
        End Select
    End Select
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptKey, strDateFrom, strDateTo)
    
    Dim lngCount As Long
    mclsPati.ClearGrid
    Call mclsDockAduit.zlClearTime
    With vfgPati
        If rs.RecordCount > 0 Then
            Call mclsPati.LoadDataSource(rs)
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .ColWidth(1) = 0
            .ColHidden(1) = True
            .Row = 0
        End If
        
    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strDeptName As String
    
    If cboDept.ListIndex < 0 Then Exit Sub
    
    strDeptName = cboDept.List(cboDept.ListIndex)
    
    If Me.ActiveControl Is Nothing Then
        Set objPrint.Body = vfgPati
        objPrint.Title.Text = strDeptName & mstrEvent & "病人清单"
    Else
        If Me.ActiveControl.Name = vfgPati.Name Then
            Set objPrint.Body = vfgPati
            objPrint.Title.Text = strDeptName & mstrEvent & "病人清单"
        Else
                        
            Set objPrint.Body = mclsDockAduit.zlGetFormAuditTimeLimit.vfgAudit
            objPrint.Title.Text = "病人病历时限报告"
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(Me.vfgPati.TextMatrix(Me.vfgPati.FixedRows - 1, 2) & ":" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 2))
            Call objAppRow.Add("姓名:" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 3))
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    End If
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub


Private Sub cboDept_Click()
    If mblnReading Then Exit Sub
    If cboDept.ListIndex < 0 Then Exit Sub
    
    If mlngDeptId <> cboDept.ItemData(cboDept.ListIndex) Then
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
        Call RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
End Sub

Private Sub cboType_Click()
    If mblnReading Then Exit Sub
    If cboType.ListIndex < 0 Then Exit Sub
    
    If mstrEvent <> cboType.List(cboType.ListIndex) Then
        mstrEvent = cboType.List(cboType.ListIndex)
        Call RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRecordId As Long
    
    Select Case Control.ID
    Case 9
        
        lngRecordId = mclsDockAduit.zlGetFormAuditTimeLimit.GetCurrentEPRKey

        If lngRecordId > 0 Then
            If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
            Call mclsDockAduit.zlOpenEPRDocument(lngRecordId, mfrmMain)
        End If
        
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case 9
        Control.Enabled = (mclsDockAduit.zlGetFormAuditTimeLimit.GetCurrentEPRKey > 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Pati
        Item.Handle = picPane(0).hWnd
    Case conPane_Audit
        
        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitTime(Me)
        
        Item.Handle = mclsDockAduit.zlGetFormAuditTimeLimit.hWnd
    Case conPane_Word
        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitMonitor(Me)
    
        Item.Handle = mclsDockAduit.zlGetFormAuditMonitor.hWnd

    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMan, 1, 100, 15, 350, Me.ScaleHeight)
    
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mclsDockAduit Is Nothing) Then Set mclsDockAduit = Nothing
        
End Sub

Private Sub mclsDockAduit_AfterDocumentChanged(ByVal lngEPRKey As Long)
    Call mclsDockAduit.zlRefreshMonitor(lngEPRKey)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgPati.Move 0, 0, picPane(Index).Width, picPane(Index).Height

        cboDept.Move -30, -30, picPane(3).Width + 45
        cboType.Move -30, -30, picPane(2).Width + 45
    End Select
End Sub

Private Sub vfgPati_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    With vfgPati
        If OldRow <> NewRow And NewRow > 0 Then
                                    
            Call mclsDockAduit.zlRefreshTime(Val(.TextMatrix(NewRow, 1)), Val(.TextMatrix(NewRow, 2)), mintKind)
                       
        End If
    End With
    
End Sub