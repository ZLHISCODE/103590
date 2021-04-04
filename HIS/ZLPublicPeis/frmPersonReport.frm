VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonReport 
   Caption         =   "体检报告"
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13245
   Icon            =   "frmPersonReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   13245
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   0
      Left            =   270
      ScaleHeight     =   5850
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   1200
      Width           =   3825
      Begin MSComctlLib.TreeView tvw 
         Height          =   3465
         Left            =   525
         TabIndex        =   2
         Top             =   420
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   1
      Left            =   4590
      ScaleHeight     =   5850
      ScaleWidth      =   8250
      TabIndex        =   0
      Top             =   1440
      Width           =   8250
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   5010
         Left            =   390
         TabIndex        =   3
         Top             =   345
         Width           =   7800
         _cx             =   13758
         _cy             =   8837
         Appearance      =   0
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
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
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
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   6885
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonReport.frx":6852
            Key             =   "报告"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   105
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPersonReport.frx":D0B4
      Left            =   210
      Top             =   660
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPersonReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPersonKey As Long
Private mlngTaskKey As Long

Private mclsVsf As clsVsf

'######################################################################################################################

Public Function ShowReportForm(ByVal frmMain As Object, ByVal lngPersonKey As Long) As Boolean
    mlngPersonKey = lngPersonKey
    
    Call InitData
    Call LoadTaskReport(mlngTaskKey)
    Me.Show 1, frmMain
    ShowReportForm = True
End Function

'******************************************************************************************************************
'功能：初始化数据
'参数：
'返回：
'******************************************************************************************************************
Private Sub InitData()

    Call InitVsf
    Call InitCommandBar
    Call InitDockPannel
     
     Set tvw.ImageList = ils16
End Sub

'******************************************************************************************************************
'功能：加载任务报告
'参数：
'返回：
'******************************************************************************************************************
Private Sub LoadTaskReport(ByVal lngTaskKey As Long)
    
    On Error GoTo errHand
    If lngTaskKey = 0 Or mlngPersonKey = 0 Then Exit Sub
    
    Call LoadItem(lngTaskKey)
    Call LoadResult(lngTaskKey)
    
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub

'******************************************************************************************************************
'功能：加载项目
'参数：
'返回：
'******************************************************************************************************************
Private Sub LoadItem(ByVal lngTaskKey As Long)

    Dim rsData As ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo errHand
    
    tvw.Nodes.Clear
    tvw.Style = tvwPlusPictureText
    
    Set rsData = gclsPackage.Get_PeisPersonItem(lngTaskKey, mlngPersonKey)
    
    Do While Not rsData.EOF
         Set objNode = tvw.Nodes.Add(, , "K" & NVL(rsData("清单ID").Value), NVL(rsData("项目").Value), "报告", "报告")
         
         rsData.MoveNext
    Loop
    
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub


'******************************************************************************************************************
'功能：加载结果
'参数：
'返回：
'******************************************************************************************************************
Private Sub LoadResult(ByVal lngTaskKey As Long)
    Dim rsConclusion As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    mclsVsf.ClearGrid
    
    
     With vsf
        
        
        '获取总检结论
        Set rsConclusion = gclsPackage.Get_PeisPersonConclusion(2, lngTaskKey, mlngPersonKey)
        If rsConclusion.BOF = False Then
            .Row = .Rows - 1
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .MergeRow(.Row) = True
            .TextMatrix(.Row, .ColIndex("ID")) = NVL(rsConclusion("ID").Value)
            .TextMatrix(.Row, .ColIndex("项目")) = "总检结论"
            .Cell(flexcpData, .Row, .ColIndex("项目"), .Row, .Cols - 1) = "总检结论"
            .Cell(flexcpText, .Row, .ColIndex("项目"), .Row, .Cols - 1) = "总检结论"
            
            Do While Not rsConclusion.EOF
                If Trim(.TextMatrix(.Rows - 1, .ColIndex("项目"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                If rsConclusion.AbsolutePosition = 1 Then .TextMatrix(.Row, .ColIndex("排序")) = 1
                .TextMatrix(.Row, .ColIndex("指标")) = NVL(rsConclusion("结论描述").Value)
                rsConclusion.MoveNext
            Loop
            
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, .ColIndex("项目")) = " "
        End If
        
        '获取体检项目
        Set rsItem = gclsPackage.Get_PeisPersonItem(lngTaskKey, mlngPersonKey)
        Do While Not rsItem.EOF
           
           If .TextMatrix(.Rows - 1, .ColIndex("项目")) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
           .Row = .Rows - 1
           
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .MergeRow(.Row) = True
           .TextMatrix(.Row, .ColIndex("ID")) = NVL(rsItem("清单ID").Value)
           
            .TextMatrix(.Row, .ColIndex("项目")) = NVL(rsItem("项目").Value)
            .Cell(flexcpData, .Row, .ColIndex("项目"), .Row, .Cols - 1) = NVL(rsItem("项目").Value)
            .Cell(flexcpText, .Row, .ColIndex("项目"), .Row, .Cols - 1) = NVL(rsItem("项目").Value)
           
           '获取指标结果
           Set rsResult = gclsPackage.get_PeisPersonResult(lngTaskKey, mlngPersonKey, Val(NVL(rsItem("清单ID").Value)))
           If rsResult.BOF = False Then
                
                If .TextMatrix(.Rows - 1, .ColIndex("项目")) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = Color.浅灰色
                
                .TextMatrix(.Row, .ColIndex("排序")) = 1
                .TextMatrix(.Row, .ColIndex("指标")) = "指标名称"
                .TextMatrix(.Row, .ColIndex("结果")) = "指标结果"
                .TextMatrix(.Row, .ColIndex("提示")) = "提示"
                .TextMatrix(.Row, .ColIndex("参考")) = "参考范围"
                
                Do While Not rsResult.EOF
                    
                    If Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .TextMatrix(.Row, .ColIndex("指标")) = NVL(rsResult("指标").Value)
                    .TextMatrix(.Row, .ColIndex("结果")) = NVL(rsResult("结果").Value)
                    .TextMatrix(.Row, .ColIndex("提示")) = NVL(rsResult("提示").Value)
                    .TextMatrix(.Row, .ColIndex("参考")) = NVL(rsResult("参考").Value)
                     Call ApplyResultColor(.Row, NVL(rsResult("提示").Value))
                    rsResult.MoveNext
                Loop
                
           End If
           '检验项目则增加检验备注、标本形态
           If Val(NVL(rsItem("采集方式id").Value)) > 0 Then
                 If Trim(.TextMatrix(.Rows - 1, .ColIndex("项目"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                .TextMatrix(.Row, .ColIndex("指标")) = "检验备注"
                .TextMatrix(.Row, .ColIndex("结果")) = NVL(rsItem("备注说明").Value)
                .Cell(flexcpData, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsItem("备注说明").Value)
                .Cell(flexcpText, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsItem("备注说明").Value)
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                .TextMatrix(.Row, .ColIndex("指标")) = "标本形态"
                .TextMatrix(.Row, .ColIndex("结果")) = NVL(rsItem("标本形态").Value)
                .Cell(flexcpData, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsItem("标本形态").Value)
                .Cell(flexcpText, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsItem("标本形态").Value)
                
           End If
           
           '获取项目小结
           Set rsConclusion = gclsPackage.Get_PeisPersonConclusion(1, lngTaskKey, mlngPersonKey, Val(NVL(rsItem("清单id").Value)))
           
           Do While Not rsConclusion.EOF
                If Trim(.TextMatrix(.Rows - 1, .ColIndex("项目"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                If rsConclusion.AbsolutePosition = 1 Then
                        
                    .TextMatrix(.Row, .ColIndex("指标")) = "【小结】"
                    .TextMatrix(.Row, .ColIndex("结果")) = NVL(rsConclusion("结论描述").Value)
                    .Cell(flexcpData, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsConclusion("结论描述").Value)
                    .Cell(flexcpText, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsConclusion("结论描述").Value)
                Else
                    
                    .TextMatrix(.Row, .ColIndex("指标")) = ""
                    .TextMatrix(.Row, .ColIndex("结果")) = NVL(rsConclusion("结论描述").Value)
                    .Cell(flexcpData, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsConclusion("结论描述").Value)
                    .Cell(flexcpText, .Row, .ColIndex("结果"), .Row, .Cols - 1) = NVL(rsConclusion("结论描述").Value)
                End If
                rsConclusion.MoveNext
           Loop
           
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("项目"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("指标"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, .ColIndex("项目")) = " "
            
           rsItem.MoveNext
        Loop
        .AutoSize 0, .ColIndex("结果")
     End With
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub


'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Private Function ApplyResultColor(ByVal lngRow As Long, ByVal str报警 As String) As Boolean
    Dim lngColor As Long
    Dim lngForeColor As Long
    Dim lngCol As Long
    Dim strSign As String
    
    If lngRow = 0 Then Exit Function
    
    strSign = str报警
    Select Case str报警
    Case "偏低"
        lngColor = Color.低标背景色
        lngForeColor = Color.超标前景色
        strSign = "↓"
    Case "偏高"
        lngColor = Color.超标背景色
        lngForeColor = Color.超标前景色
        strSign = "↑"
    Case "异常"
        lngColor = Color.超标背景色
        lngForeColor = Color.超标前景色
    Case "警戒上限"
        lngColor = Color.警戒偏高背景色
        lngForeColor = Color.超标前景色
    Case "警戒下限"
        lngColor = Color.警戒偏低背景色
        lngForeColor = Color.超标前景色
    Case "复查上限"
        lngColor = Color.复查偏高背景色
        lngForeColor = Color.超标前景色
    Case "复查下限"
        lngColor = Color.复查偏低背景色
        lngForeColor = Color.超标前景色
    Case Else
        lngColor = &H80000005
        lngForeColor = Color.默认前景色
    End Select
    
    lngCol = vsf.ColIndex("结果")
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
    vsf.TextMatrix(lngRow, vsf.ColIndex("提示")) = strSign
    
    ApplyResultColor = True
    
    
End Function


'******************************************************************************************************************
'功能：初始化表格
'参数：
'返回：
'******************************************************************************************************************
Private Sub InitVsf()
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf, True, False)
        Call .ClearColumn
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, False, False, True)
        Call .AppendColumn("项目", 255, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("指标", 2100, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("结果", 4030, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("提示", 450, flexAlignLeftTop, flexDTString, "", , True)
        Call .AppendColumn("排序", 0, flexAlignLeftTop, flexDTString, "", , True, , , True)
        Call .AppendColumn("类型", 0, flexAlignLeftTop, flexDTString, "", , True, , , True)
        Call .AppendColumn("参考", 900, flexAlignLeftTop, flexDTString, , , False)
        
        
        .AppendRows = False
        .AutoRowHeight = True
    End With
    vsf.RowHidden(0) = True

End Sub

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Private Function InitCommandBar() As Boolean
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim rsData As ADODB.Recordset

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    
    cbsMain.Options.LargeIcons = True
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "体检历史", True)
    objControl.IconId = 9023
    
    Set rsData = gclsPackage.Get_PeisPersonReportBill(mlngPersonKey)
    
    Do While Not rsData.EOF
        
       
        
        If rsData.AbsolutePosition < 3 Then
             Set objControl = NewToolBar(objBar, xtpControlButton, 1, Format(NVL(rsData("开始时间").Value), "yyyy-MM-dd"), True)
             objControl.Parameter = Val(NVL(rsData("任务ID").Value))
             If rsData.AbsolutePosition = 1 Then
'                objControl.Checked = True
                mlngTaskKey = Val(NVL(rsData("任务ID").Value))
             End If
        ElseIf rsData.AbsolutePosition = 3 Then
              Set objPopup = NewToolBar(objBar, xtpControlPopup, 0, "更多", True)
              
              Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 1, Format(NVL(rsData("开始时间").Value), "yyyy-MM-dd"))
              objControl.Parameter = Val(NVL(rsData("任务ID").Value))
              
        ElseIf rsData.AbsolutePosition > 3 Then
              
              Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, 1, Format(NVL(rsData("开始时间").Value), "yyyy-MM-dd"))
              objControl.Parameter = Val(NVL(rsData("任务ID").Value))
        End If
        
        rsData.MoveNext
    Loop
    
     
    Set objControl = NewToolBar(objBar, xtpControlButton, 2, "退出", True)
    objControl.IconId = 191
            
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Private Sub InitDockPannel()
    Dim objPane As Pane

    Call DockPannelInit(dkpMain)
    Set objPane = DockPannelCreate(dkpMain, 1, 200, 100, DockLeftOf, Nothing, "人员", PaneNoCaption)
    Set objPane = DockPannelCreate(dkpMain, 2, 600, 100, DockRightOf, objPane, "详细资料", PaneNoCaption)

    dkpMain.SetCommandBars cbsMain
End Sub



'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case 1  '任务
            mlngTaskKey = Val(Control.Parameter)
            Call LoadTaskReport(Val(Control.Parameter))
        Case 2
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case 1
            Control.Checked = (Val(Control.Parameter) = mlngTaskKey)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picPane(0).hWnd
        Case 2
            Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 200, 100, 300, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lngRow As Long
    
    lngRow = mclsVsf.FindRow(Mid(Node.Key, 2), vsf.ColIndex("ID"))

    If lngRow > 0 Then
        vsf.TopRow = lngRow
    End If
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngSvrBkColor As Long
    Dim rc As RECT
    Dim rc1 As RECT
    Dim r1%, g1%, b1%
    Dim r2%, g2%, b2%
    Dim rg%, gg%, bg%
    Dim lngLoop As Long
    
    On Error Resume Next
    
    With vsf
        
        If Val(.TextMatrix(Row, .ColIndex("排序"))) <> 1 Then Exit Sub

'        'flexODOver
'        '--------------------------------------------------------------------------------------------------------------
        rc.Left = Left
        rc.Top = Top
        rc.Right = Right
        rc.Bottom = Top + 1


        'Draw Frame
        '--------------------------------------------------------------------------------------------------------------
        lngSvrBkColor = SetBkColor(hDC, 0)

        Call ExtTextOut(hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, " ", 1, lngLoop)
        Call InflateRect(rc, -1, -1)

'        Call SetBkColor(hDC, RGB(255, 255, 255))
        Call ExtTextOut(hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, " ", 1, lngLoop)
        
    End With
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0
            tvw.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        Case 1
            vsf.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub
