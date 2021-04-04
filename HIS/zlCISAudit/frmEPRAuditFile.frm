VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEPRAuditFile 
   BorderStyle     =   0  'None
   Caption         =   "病历文件分类审查"
   ClientHeight    =   6210
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   9630
   Icon            =   "frmEPRAuditFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2280
      Index           =   0
      Left            =   735
      ScaleHeight     =   2280
      ScaleWidth      =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   4440
      Begin VSFlex8Ctl.VSFlexGrid vfgFile 
         Height          =   1200
         Left            =   735
         TabIndex        =   4
         Top             =   360
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   1
      Left            =   945
      ScaleHeight     =   3135
      ScaleWidth      =   5625
      TabIndex        =   2
      Top             =   3435
      Width           =   5625
      Begin VSFlex8Ctl.VSFlexGrid vfgEPRs 
         Height          =   1200
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   9
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
      Left            =   5655
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   540
      Width           =   1905
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   -45
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   -30
         Width           =   1320
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
      Bindings        =   "frmEPRAuditFile.frx":058A
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'常量
'----------------------------------------------------------------------------------------------------------------------
Private Const conPane_File = 1
Private Const conPane_EPRs = 2
Private Const conPane_Word = 3


'变量
'----------------------------------------------------------------------------------------------------------------------
Private mlngDeptId As Long      '科室id
Private mstrDeptName As String  '科室名
Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mclsDockAduit As zlRichEPR.clsDockAduits
Private mfrmMain As Object
Private mblnReading As Boolean
Private mclsEPRs As clsVsf
Private mclsFile As clsVsf

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
    With vfgFile
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
    '参数：blnShow 可以只刷新cboDept
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldValue As String
    On Error GoTo ErrH
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    
    Select Case mintKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '门诊病历
        
        strSQL = "Select distinct D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 in ('临床','手术','治疗') And M.服务对象 In (1, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2  '住院病历
        
        strSQL = "Select distinct D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 in ('临床','手术','治疗') And M.服务对象 In (2, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '护理病历
    
        strSQL = "Select distinct D.ID, D.编码, D.名称 From 部门表 D, 部门性质说明 M Where D.ID = M.部门id And M.工作性质 in ('临床','手术','治疗') And M.服务对象 In (2, 3) And ( TO_CHAR (D.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or D.撤档时间 is null) Order By D.编码"

    End Select
    
    '------------------------------------------------------------------------------------------------------------------
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
    
    '------------------------------------------------------------------------------------------------------------------
    
    If blnShow = False Then
        zlRefreshData = RefreshData(mintKind, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    Exit Function
ErrH:
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
    
    
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = False
                
            
    '部门工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlLabel, 7, "科室：", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(3).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 9, "查阅病历...", True, , xtpButtonIconAndCaption)
    
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    '------------------------------------------------------------------------------------------------------------------
    Set mclsFile = New clsVsf
    With mclsFile
        
        Call .Initialize(Me.Controls, vfgFile, True, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("编号", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("名称", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("应写数", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("已完成", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("在书写", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("要求", 1200, flexAlignLeftCenter, flexDTString, "", , True)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsEPRs = New clsVsf
    With mclsEPRs
        
        Call .Initialize(Me.Controls, vfgEPRs, False, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                
        Call .AppendColumn("门诊号", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("姓名", 1200, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("性别", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("就诊日期", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
        Call .AppendColumn("创建人", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("书写人", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("完成时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
        
    End With
                    
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
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
        
        '划分停靠区域
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMan.CreatePane(conPane_File, 100, 100, DockLeftOf, Nothing): objPane.Title = "文件清单": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_EPRs, 400, 300, DockRightOf, objPane): objPane.Title = "书写记录": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Word, 600, 400, DockBottomOf, objPane): objPane.Title = "内容监测": objPane.Options = PaneNoCaption
        
        dkpMan.SetCommandBars cbsThis
        Call DockPannelInit(dkpMan)
                                
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                        
        
                
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

Private Function RefreshData(ByVal intKind As Integer, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim intOut24h As Byte    '是否区分24小时出院或死亡：0-不区分,1-区分；根据是否定义24小时事件对应病历确定
    
    On Error GoTo errHand
    
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        strSQL = "Select F.ID, F.编号, F.名称, F.事件 || '时书写' As 要求, P.人次 As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 1) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine & _
                "     (Select E.事件, Decode(E.事件, '门诊', 门诊, '初诊', 初诊, '复诊', 复诊, '急诊', 急诊) As 人次" & vbNewLine & _
                "       From (Select Sum(Decode(急诊, 1, 0, 1)) As 门诊, Sum(Decode(急诊, 1, 0, Decode(复诊, 1, 0, 1))) As 初诊," & vbNewLine & _
                "                     Sum(Decode(急诊, 1, 0, Decode(复诊, 1, 1, 0))) As 复诊, Sum(Decode(急诊, 1, 1, 0)) As 急诊" & vbNewLine & _
                "              From 病人挂号记录" & vbNewLine & _
                "              Where 执行部门id = [1] And Nvl(执行状态, 0) <> 0 And 记录性质=1 And 记录状态=1 And 登记时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '门诊', 2, '初诊', 3, '复诊', 4, '急诊') As 事件 From 病历书写事件 Where Rownum < 5) E) P," & vbNewLine & _
                "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 1 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where F.事件 = P.事件 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        strSQL = "Select Sign(Nvl(Count(*), 0))" & vbNewLine & _
                "From (Select F.ID, F.通用, A.科室id" & vbNewLine & _
                "       From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "       Where F.ID = A.文件id(+) And F.ID = Q.文件id And Q.事件 In ('24小时出院', '24小时死亡') And F.种类 = 2) F" & vbNewLine & _
                "Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, mlngDeptId)
        If rs.RecordCount <= 0 Then
            intOut24h = 0
        Else
            intOut24h = rs.Fields(0).Value
        End If
        strSQL = "Select  F.ID, F.编号, F.名称, F.事件 || Decode(Sign(F.书写时限), -1, '前', '后') || '书写' As 要求," & vbNewLine & _
                "       Decode(F.唯一, 1, To_Char(P.人次), '<循环>') As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 2) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine
        If intOut24h = 1 Then
            strSQL = strSQL & "     (Select E.事件, '后' As 时机, Decode(E.事件, '入院', 入院, '首次入院', 首次入院, '再次入院', 再次入院) As 人次" & vbNewLine & _
                    "       From (Select Count(*) As 入院, Sum(Decode(再入院, 1, 0, 1)) As 首次入院," & vbNewLine & _
                    "                     Sum(Decode(再入院, 1, 1, 0)) As 再次入院" & vbNewLine & _
                    "              From 病案主页" & vbNewLine & _
                    "              Where 入院科室id + 0 = [1] And Nvl(出院日期, Sysdate + 1) - 入院日期 > 1 And" & vbNewLine & _
                    "                    入院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '入院', 2, '首次入院', 3, '再次入院') As 事件 From 病历书写事件 Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sign(出院日期 - 入院日期 - 1), -1, Decode(出院方式, '死亡', '24小时死亡', '24小时出院')," & vbNewLine & _
                    "                      Decode(出院方式, '死亡', '死亡', '出院')) As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                    "       From 病案主页" & vbNewLine & _
                    "       Where 出院科室id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(Sign(出院日期 - 入院日期 - 1), -1, Decode(出院方式, '死亡', '24小时死亡', '24小时出院')," & vbNewLine & _
                    "                        Decode(出院方式, '死亡', '死亡', '出院'))" & vbNewLine
        Else
            strSQL = strSQL & "     (Select E.事件, '后' As 时机, Decode(E.事件, '入院', 入院, '首次入院', 首次入院, '再次入院', 再次入院) As 人次" & vbNewLine & _
                    "       From (Select Count(*) As 入院, Sum(Decode(再入院, 1, 0, 1)) As 首次入院," & vbNewLine & _
                    "                     Sum(Decode(再入院, 1, 1, 0)) As 再次入院" & vbNewLine & _
                    "              From 病案主页" & vbNewLine & _
                    "              Where 入院科室id + 0 = [1] And" & vbNewLine & _
                    "                    入院日期 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '入院', 2, '首次入院', 3, '再次入院') As 事件 From 病历书写事件 Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(出院方式, '死亡', '死亡', '出院') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                    "       From 病案主页" & vbNewLine & _
                    "       Where 出院科室id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(出院方式, '死亡', '死亡', '出院')" & vbNewLine
        End If
        strSQL = strSQL & "       Union All" & vbNewLine & _
                "       Select Decode(开始原因, 3, '转科', 7, '交班') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 科室id + 0 = [1] And 开始原因 In (3, 7) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             开始时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(开始原因, 3, '转科', 7, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(终止原因, 3, '转科', 7, '交班') As 事件, '前' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 科室id + 0 = [1] And 终止原因 In (3, 7) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             终止时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(终止原因, 3, '转科', 7, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select R.事件, E.时机, Decode(E.时机, '前', R.前人次, '后', R.后人次) As 人次" & vbNewLine & _
                "       From (Select Decode(R.诊疗类别,'F', '手术', Decode(I.操作类型, '7', '会诊', '抢救')) As 事件," & vbNewLine & _
                "                     Sum(Decode(R.病人科室id, [1], 1, 0)) As 前人次, Sum(Decode(R.执行科室id, [1], 1, 0)) As 后人次" & vbNewLine & _
                "              From 病人医嘱记录 R, 诊疗项目目录 I, 病人医嘱发送 S" & vbNewLine & _
                "              Where R.ID = S.医嘱id And R.诊疗项目id = I.ID And" & vbNewLine & _
                "                    (R.诊疗类别 = 'F' Or R.诊疗类别 = 'Z' And I.操作类型 In ('7', '8')) And R.相关id Is Null And" & vbNewLine & _
                "                    R.医嘱期效 = 1 And (R.医嘱状态 = 8 Or R.医嘱状态 = 9) And" & vbNewLine & _
                "                    (R.病人科室id + 0 = [1] Or R.执行科室id + 0 = [1]) And S.发送时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "              Group By Decode(R.诊疗类别,'F', '手术', Decode(I.操作类型, '7', '会诊', '抢救'))) R," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '前', 2, '后') As 时机 From 病历书写事件 Where Rownum < 3) E) P," & vbNewLine
        
        strSQL = strSQL & "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 2 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where  F.事件 = P.事件 and Decode(Sign(F.书写时限), -1, '前', '后') = P.时机 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
                
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "Select F.ID, F.编号, F.名称, F.事件 || Decode(Sign(F.书写时限), -1, '前', '后') || '书写' As 要求," & vbNewLine & _
                "       Decode(F.唯一, 1, To_Char(P.人次), '<循环>') As 应写数, W.已完成, W.在书写" & vbNewLine & _
                "From (Select F.ID, F.编号, F.名称, F.事件, F.唯一, F.书写时限" & vbNewLine & _
                "       From (Select F.ID, F.编号, F.名称, F.通用, A.科室id, Q.事件, Q.唯一, Q.书写时限" & vbNewLine & _
                "              From 病历文件列表 F, 病历应用科室 A, 病历时限要求 Q" & vbNewLine & _
                "              Where F.ID = A.文件id(+) And F.ID = Q.文件id And F.种类 = 4) F" & vbNewLine & _
                "       Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [1]) F," & vbNewLine
        strSQL = strSQL & "     (Select '入院' As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病案主页" & vbNewLine & _
                "       Where 入院病区id + 0 = [1] And 入院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(出院方式, '死亡', '死亡', '出院') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病案主页" & vbNewLine & _
                "       Where 当前病区id + 0 = [1] And 出院日期 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(出院方式, '死亡', '死亡', '出院')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(开始原因, 3, '转科', 8, '交班') As 事件, '后' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 病区id + 0 = [1] And 开始原因 In (3, 8) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             开始时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(开始原因, 3, '转科', 8, '交班')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(终止原因, 3, '转科', 8, '交班') As 事件, '前' As 时机, Count(*) As 人次" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 病区id + 0 = [1] And 终止原因 In (3, 8) And Nvl(附加床位, 0) = 0 And" & vbNewLine & _
                "             终止时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(终止原因, 3, '转科', 8, '交班')) P," & vbNewLine
        strSQL = strSQL & "     (Select 文件id, Sum(Decode(完成时间, Null, 0, 1)) As 已完成, Sum(Decode(完成时间, Null, 1, 0)) As 在书写" & vbNewLine & _
                "       From 电子病历记录" & vbNewLine & _
                "       Where 病历种类 = 4 And 科室id + 0 = [1] And 创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By 文件id) W" & vbNewLine & _
                "Where F.事件 = P.事件 And Decode(Sign(F.书写时限), -1, '前', '后') = P.时机 And P.人次 > 0 And F.ID = W.文件id(+)" & vbNewLine & _
                "Order By F.编号"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
    End Select
    
    Dim lngCount As Long
    
    mclsFile.ClearGrid
    mclsEPRs.ClearGrid
'    If Not (mfrmEPRAuditMonitor Is Nothing) Then mfrmEPRAuditMonitor.zlClearData
    If Not (mclsDockAduit Is Nothing) Then mclsDockAduit.zlClearMonitor
    
    With vfgFile
        If rs.RecordCount > 0 Then
            .Clear
            Set .DataSource = rs
            .ColWidth(0) = 0: .ColHidden(0) = True
            .ColAlignment(4) = flexAlignRightCenter
            For lngCount = 1 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
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
    
    If Me.ActiveControl Is Nothing Then
        Set objPrint.Body = vfgFile
        objPrint.Title.Text = mstrDeptName & "病历文件列表"
    Else
        If Me.ActiveControl.Name = vfgFile.Name Then
            Set objPrint.Body = vfgFile
            objPrint.Title.Text = mstrDeptName & "病历文件列表"
        Else
            Set objPrint.Body = vfgEPRs
            objPrint.Title.Text = mstrDeptName & vfgFile.TextMatrix(vfgFile.Row, 2) & "清单"
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
        Call RefreshData(mintKind, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRecordId As Long
    
    Select Case Control.ID
    Case 9
        
        With Me.vfgEPRs
            lngRecordId = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
        
        If lngRecordId > 0 Then
            If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
            Call mclsDockAduit.zlOpenEPRDocument(lngRecordId, mfrmMain)
        End If
        
    End Select
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 9
        With vfgEPRs
            Control.Enabled = (Val(.TextMatrix(.Row, 0)) > 0)
        End With
    End Select
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_File
        Item.Handle = picPane(0).hWnd
    Case conPane_EPRs
        Item.Handle = picPane(1).hWnd
    Case conPane_Word

        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitMonitor(Me)
        Item.Handle = mclsDockAduit.zlGetFormAuditMonitor.hWnd

    End Select
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMan, 1, 100, 15, 360, Me.ScaleHeight)
    
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mclsDockAduit Is Nothing) Then Set mclsDockAduit = Nothing
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgFile.Move 0, 0, picPane(Index).Width, picPane(Index).Height
        cboDept.Move -30, -30, picPane(3).Width + 45
    Case 1
        vfgEPRs.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub vfgEPRs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgEPRs
        If OldRow <> NewRow And NewRow > 0 Then
            
            If Not (mclsDockAduit Is Nothing) Then
                
                Call mclsDockAduit.zlRefreshMonitor(Val(vfgFile.TextMatrix(vfgFile.Row, 0)))
                
            End If
                        
        End If
    End With
End Sub

Private Function RefreshPatient(ByVal intKind As Integer, ByVal lngFileID As Long, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        strSQL = "Select W.ID, P.病人id, P.门诊号, P.姓名, P.性别, To_Char(P.执行时间, 'mm-dd hh24:mi') As 就诊日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病人挂号记录 P" & vbNewLine & _
                "Where W.主页id = P.ID And W.病历种类 = 1 And W.科室id + 0 = [1] And W.文件id + 0 = [4] And" & vbNewLine & _
                "      W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By P.执行时间"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo, lngFileID)
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        strSQL = "Select W.ID, I.病人id, P.住院号, I.姓名, I.性别, To_Char(P.入院日期, 'mm-dd hh24:mi') As 入院日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病案主页 P, 病人信息 I" & vbNewLine & _
                "Where I.病人id = P.病人id And P.病人id = W.病人id And P.主页id = W.主页id And W.病历种类 = 2 And W.科室id + 0 = [1] And" & vbNewLine & _
                "      W.文件id + 0 = [4] And W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By 入院日期"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "Select W.ID, I.病人id, P.住院号, I.姓名, I.性别, To_Char(P.入院日期, 'mm-dd hh24:mi') As 入院日期, W.创建人 As 书写人," & vbNewLine & _
                "       To_Char(W.完成时间, 'mm-dd hh24:mi') As 完成时间" & vbNewLine & _
                "From 电子病历记录 W, 病案主页 P, 病人信息 I" & vbNewLine & _
                "Where I.病人id = P.病人id And P.病人id = W.病人id And P.主页id = W.主页id And W.病历种类 = 4 And W.科室id + 0 = [1] And" & vbNewLine & _
                "      W.文件id + 0 = [4] And W.创建时间 Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By 入院日期"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo, lngFileID)
    End Select
    
    Dim lngCount As Long
    
    mclsEPRs.ClearGrid
    
    With Me.vfgEPRs
        
        If rs.RecordCount > 0 Then
            .Clear
            Set .DataSource = rs
            .ColWidth(0) = 0: .ColHidden(0) = True
            For lngCount = 1 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
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

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgFile
        If OldRow <> NewRow And NewRow > 0 Then
            
            Call RefreshPatient(mintKind, Val(.TextMatrix(NewRow, 0)), mlngDeptId, mstrDateFrom, mstrDateTo)
            
        End If
    End With
End Sub

