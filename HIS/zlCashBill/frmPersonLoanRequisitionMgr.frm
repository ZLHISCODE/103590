VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmPersonLoanRequisitionMgr 
   BorderStyle     =   0  'None
   Caption         =   "借款列表"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   4710
      Left            =   75
      TabIndex        =   0
      Top             =   1620
      Width           =   7980
      _Version        =   589884
      _ExtentX        =   14076
      _ExtentY        =   8308
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrint 
      Height          =   1365
      Left            =   8970
      TabIndex        =   1
      Top             =   2550
      Visible         =   0   'False
      Width           =   540
      _cx             =   952
      _cy             =   2408
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   1320
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":0000
            Key             =   "等待借款"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":059A
            Key             =   "拒绝申请"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":0B34
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":10CE
            Key             =   "取消借出"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonLoanRequisitionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mstrPrivs As String, mlngModule As Long, mArrFilter As Variant  '过滤条件
Private mcbsThis As Object
Private Type rptColIndexType    '列索引
    ColID  As Integer
    Col图标 As Integer
    Col状态 As Integer
    Col申请人 As Integer
    Col申请时间 As Integer
    Col备注 As Integer
    Col借款金额  As Integer
    Col借出人 As Integer
    Col借出时间 As Integer
    Col取消时间 As Integer
    Col取消原因 As Integer
End Type
Private mRptCol As rptColIndexType

Public Function zlReLoadData(ByVal mcllFilter As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-09-07 14:43:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mArrFilter = mcllFilter
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitReportColumn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化的取表列
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-07 11:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCol As ReportColumn, i As Long
    
    With mRptCol
        .ColID = 0: i = i + 1
        .Col图标 = i: i = i + 1
        .Col状态 = i: i = i + 1
        .Col申请人 = i: i = i + 1
        .Col申请时间 = i: i = i + 1
        .Col备注 = i: i = i + 1
        
        .Col借款金额 = i: i = i + 1
        .Col借出人 = i: i = i + 1
        .Col借出时间 = i: i = i + 1
        .Col取消时间 = i: i = i + 1
        .Col取消原因 = i: i = i + 1
    End With
    With rptList
        '当前顺序:ID,Col图标,Col状态,申请人,申请时间,借款金额,借出人,借出时间, 取消时间,取消原因
        
       ' 当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(mRptCol.ColID, "ID", 0, False)
            objCol.Sortable = False: objCol.Visible = False
            
        Set objCol = .Columns.Add(mRptCol.Col图标, "", 25, False)
        
        Set objCol = .Columns.Add(mRptCol.Col状态, "状态", 55, False): objCol.Visible = False
        Set objCol = .Columns.Add(mRptCol.Col申请人, "申请人", 55, True): objCol.Visible = False
            'objCol.TreeColumn = True: 'objCol.Visible = False
            'objCol.Sortable = False: objCol.AllowDrag = False
        Set objCol = .Columns.Add(mRptCol.Col申请时间, "申请时间", 136, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(mRptCol.Col备注, "备注", 300, True)
        Set objCol = .Columns.Add(mRptCol.Col借款金额, "借款金额", 65, True)
        objCol.Alignment = xtpAlignmentRight
        Set objCol = .Columns.Add(mRptCol.Col借出人, "借出人", 65, True)
        Set objCol = .Columns.Add(mRptCol.Col借出时间, "借出时间", 136, True)
        Set objCol = .Columns.Add(mRptCol.Col取消时间, "取消时间", 136, True)
        Set objCol = .Columns.Add(mRptCol.Col取消原因, "取消原因", 200, True)
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = mRptCol.Col状态
            objCol.Groupable = objCol.Index = mRptCol.Col借出人
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有申请借款的人员..."
        End With
        .PreviewMode = True: .AllowColumnRemove = False: .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False: .SetImageList Me.imgList
        .GroupsOrder.Add .Columns(mRptCol.Col状态): .GroupsOrder(0).SortAscending = True       '分组之后,如果分组列不显示,分组列的排序是不变的
         
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        '.SortOrder.Add .Columns(mRptCol.Col借出人): .SortOrder(0).SortAscending = True
    End With
End Sub

Private Sub Form_Load()
    '初始化权限串
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    '初始化列
    Call InitReportColumn
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With rptList
        .Left = ScaleLeft: .Top = ScaleTop
        .Width = ScaleWidth: .Height = ScaleHeight
    End With
End Sub
Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给报表
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, str借出人 As String, j As Long, i As Long
    Dim objParent As ReportRecord, objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strTemp As String
    Err = 0: On Error GoTo ErrHand:
    
    If CStr(mArrFilter("申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 借出时间 between [3] and [4] or 取消时间 between [5] and [6])  "
    ElseIf CStr(mArrFilter("申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("取消时间")(0)) = "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 借出时间 between [3] and [4]   )   "
    ElseIf CStr(mArrFilter("申请时间")(0)) <> "1901-01-01" And CStr(mArrFilter("借出时间")(0)) = "1901-01-01" And CStr(mArrFilter("取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2] or 取消时间 between [5] and [6])   "
    ElseIf CStr(mArrFilter("申请时间")(0)) = "1901-01-01" And CStr(mArrFilter("借出时间")(0)) <> "1901-01-01" And CStr(mArrFilter("取消时间")(0)) <> "1901-01-01" Then
        strFilter = "   ( 借出时间 between [3] and [4] or 取消时间 between [5] and [6])  "
    ElseIf CStr(mArrFilter("申请时间")(0)) <> "1901-01-01" Then
        strFilter = "   (申请时间 between [1] and [2]   ) and 借出时间 is  Null "
    ElseIf CStr(mArrFilter("借出时间")(0)) <> "1901-01-01" Then
        strFilter = "   (借出时间 between [3] and [4])"
    Else
        strFilter = "   (取消时间 between [5] and [6] )"
    End If
 
    strFilter = strFilter & " and 借款人 = [7]"
    If CStr(mArrFilter("借出人")) <> "" Then strFilter = strFilter & " and 借出人 like [8]"
    
    zlCommFun.ShowFlash "正在装载借款数据,请稍后..."

    gstrSQL = " " & _
    "    Select Id, 借款金额, 备注, 借款人, to_char(申请时间,'yyyy-mm-dd hh24:mi:ss') as 申请时间 ,  " & _
    "           借出人, to_char(借出时间,'yyyy-mm-dd hh24:mi:ss') as 借出时间, " & _
    "           to_char(取消时间,'yyyy-mm-dd hh24:mi:ss') as 取消时间, 取消原因, " & _
    "           decode(借出时间,NULL,'等待借出',decode(取消时间,NULL,'已经借出','借出取消')) as 状态, " & _
    "           decode(借出时间,NULL,1,decode(取消时间,NULL,2,3)) as 状态标志 " & _
    "    From 人员借款记录 " & _
    "    Where " & strFilter & _
    "    Order by 状态,借出人,申请时间"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("申请时间")(0)), CDate(mArrFilter("申请时间")(1)), _
        CDate(mArrFilter("借出时间")(0)), CDate(mArrFilter("借出时间")(1)), _
        CDate(mArrFilter("取消时间")(0)), CDate(mArrFilter("取消时间")(1)), _
        UserInfo.姓名, GetMatchingSting(CStr(mArrFilter("借出人")), False))
    
    rptList.Records.DeleteAll
    rptList.Columns(mRptCol.ColID).Visible = False
    With rsTemp
        Do While Not .EOF
            Set objRecord = Me.rptList.Records.Add()
            objRecord.Tag = CStr(Nvl(!ID))  '用于定位
            
            'ID,Col图标,Col状态,申请人,申请时间,借款金额,借出人,借出时间, 取消时间,取消原因
            Set objItem = objRecord.AddItem(Val(Nvl(!ID)))  '分组以Value进行排序
            objItem.Caption = CStr(Nvl(!ID))
        
            '图标:注意用在这里是从0开始编号。
            '     图标Value用于存放是否已提交审查，点击才读取
            Set objItem = objRecord.AddItem(-1)
            objItem.Caption = " "
            objItem.Icon = Decode(Nvl(!状态), "已经借出", 2, "借出取消", 3, 1)
        
        
            Set objItem = objRecord.AddItem(Val(Nvl(!状态标志)))
            objItem.Caption = Nvl(!状态)
'            If Nvl(!状态) = "借出取消" Then
'                objRecord.PreviewText = "  取消原因:" & Nvl(!取消原因)
'            End If
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!借款人)))
            objItem.Caption = CStr(Nvl(!借款人))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!申请时间)))
            objItem.Caption = CStr(Nvl(!申请时间))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!备注)))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!借款金额)))
            objItem.Caption = Format(Val(Nvl(!借款金额)), "###0.00;-###0.00")
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!借出人)))
            objItem.Caption = CStr(Nvl(!借出人))
            Set objItem = objRecord.AddItem(CStr(Nvl(!借出时间)))
            objItem.Caption = CStr(Nvl(!借出时间))
        
            Set objItem = objRecord.AddItem(CStr(Nvl(!取消时间)))
            objItem.Caption = CStr(Nvl(!取消时间))
            Set objItem = objRecord.AddItem(CStr(Nvl(!取消原因)))
            objItem.Caption = CStr(Nvl(!取消原因))
       
            '显示颜色
            For j = 0 To rptList.Columns.Count - 1
                If j = mRptCol.Col取消时间 Then
                    objRecord.Item(j).ForeColor = vbRed
                End If
            Next
           .MoveNext
        Loop
    End With
    rptList.Populate
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
End Sub
Private Function GetCurrRecordFun(Optional ByRef lngID As Long = 0) As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取操作功能
    '出参:返回当前选择的ID
    '返回:0-当前选择的分组,不进行任何处理;1-等待借出,2-已经借出,但未取消审核;3-已经被取消借出
    '编制:刘兴洪
    '日期:2009-09-09 09:26:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    With rptList.SelectedRows(0)
        If .GroupRow Then Exit Function
        lngID = Val(.Record(mRptCol.ColID).Value)
        GetCurrRecordFun = Val(.Record(mRptCol.Col状态).Value)
        
    End With
    If lngID = 0 Then Exit Function
End Function
Private Function DeleteLoanRequisition() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除借款操作
    '返回:删除成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-09 09:42:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, str借出人 As String
    With rptList.SelectedRows(0)
        If .GroupRow Then Exit Function
        lngID = Val(.Record(mRptCol.ColID).Value)
        str借出人 = .Record(mRptCol.Col借出人).Value
        If MsgBox("你真的要删除向“" & str借出人 & "”的借款申请吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
    End With
    If lngID = 0 Then Exit Function
    
    'Zl_人员借款记录_Delete(Id_In In 人员借款记录.ID%Type) Is
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "Zl_人员借款记录_Delete(" & lngID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    DeleteLoanRequisition = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlDefCommandBars(ByVal cbsThis As Object) As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
      
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrintSet, "借款单打印设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "申请借款(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改借款(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除借款(&D)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        mcbrControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): mcbrControl.BeginGroup = True
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "申请借款"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, objRow As ReportRow
    Dim lngID  As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                rptList.SelectedRows(0).Expanded = False
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    rptList.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '因折叠定位到分组上,不会自动激活该事件
        'Call rptList_SelectionChanged
    Case conMenu_View_Expend_CurExpend '展开当前组
        If rptList.SelectedRows.Count > 0 Then
            rptList.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '折叠所有组
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '因折叠定位到分组上,不会自动激活该事件
        'Call rptList_SelectionChanged
    Case conMenu_View_Expend_AllExpend '展开所有组
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
        
    Case conMenu_Edit_NewItem   '申请
        If frmPersonLoanRequisitionEdit.ShowEdit(Me, FN_申请, mstrPrivs, mlngModule) = False Then Exit Sub
        '重新刷新数据
        Call LoadDataToRpt
    Case conMenu_Edit_Modify    '修改
        With rptList.SelectedRows(0)
            If .GroupRow Then Exit Sub
            lngID = Val(.Record(mRptCol.ColID).Value)
        End With
        If lngID = 0 Then Exit Sub
            
        If frmPersonLoanRequisitionEdit.ShowEdit(Me, FN_修改, mstrPrivs, mlngModule, lngID) = False Then Exit Sub
        '重新刷新数据
        Call LoadDataToRpt
    Case conMenu_Edit_Delete '删除操作
        If DeleteLoanRequisition = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call LoadDataToRpt
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With rptList.SelectedRows(0)
                If .GroupRow = False Then
                    lngID = Val(.Record(mRptCol.ColID).Value)
                End If
            End With
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ID=" & lngID)
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count >= 1)
    Case conMenu_Edit_NewItem '申请
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "申请借款")
            Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改借款")
        Control.Enabled = Control.Visible And GetCurrRecordFun(lngID) = 1 '0-当前选择的分组,不进行任何处理;1-等待借出,2-已经借出,但未取消审核;3-已经被取消借出
        Control.Enabled = Control.Enabled And lngID <> 0
        
    Case conMenu_Edit_Delete
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除借款")
        Control.Enabled = Control.Visible And GetCurrRecordFun(lngID) = 1 '0-当前选择的分组,不进行任何处理;1-等待借出,2-已经借出,但未取消审核;3-已经被取消借出
        Control.Enabled = Control.Enabled And lngID <> 0
    Case conMenu_View_Expend_CurExpend '展开当前组
        blnEnabled = False
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptList.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        blnEnabled = False
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                blnEnabled = rptList.SelectedRows(0).Expanded
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptList.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '折叠/展开组
        Control.Enabled = rptList.GroupsOrder.Count > 0 And rptList.Rows.Count > 0
    Case conMenu_View_Refresh
        
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, rptRow As ReportRow, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
     
    With rptList
        vsPrint.Redraw = flexRDNone
        vsPrint.Cols = .Columns.Count + 1
        For i = 0 To .Columns.Count - 1
            vsPrint.TextMatrix(0, i) = .Columns(i).Caption
            vsPrint.ColWidth(i) = .Columns(i).Width * Screen.TwipsPerPixelX
        Next
        vsPrint.Clear 1
        vsPrint.Rows = 2: lngRow = 1
        For r = 0 To .Rows.Count - 1
            Set rptRow = .Rows(r)
            If rptRow.GroupRow = False Then
                For i = 0 To .Columns.Count - 1
                    vsPrint.TextMatrix(lngRow, i) = rptRow.Record(i).Caption
                Next
                lngRow = lngRow + 1
                vsPrint.Rows = vsPrint.Rows + 1
            End If
        Next
        vsPrint.Redraw = flexRDBuffered
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & "借款清单"
    
    If CStr(mArrFilter("申请时间")(0)) <> "1901-01-01" Then
        objRow.Add "申请时间：" & CStr(mArrFilter("申请时间")(0)) & "至" & CStr(mArrFilter("申请时间")(1))
    End If
    If CStr(mArrFilter("借出时间")(0)) <> "1901-01-01" Then
        objRow.Add "借出时间：" & CStr(mArrFilter("借出时间")(0)) & "至" & CStr(mArrFilter("借出时间")(1))
    End If
    If CStr(mArrFilter("取消时间")(0)) <> "1901-01-01" Then
        objRow.Add "取消时间：" & CStr(mArrFilter("取消时间")(0)) & "至" & CStr(mArrFilter("取消时间")(1))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "借款人：" & UserInfo.姓名
    If CStr(mArrFilter("借出人")) <> "" Then objRow.Add "借出人：" & mArrFilter("借出人")
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsPrint
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
        
    If Button = 2 Then
        Set objHitTest = rptList.HitTest(x, y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = mcbsThis.FindControl(, conMenu_View_Expend, , True)
            ElseIf objHitTest.Row.Childs.Count = 0 Then
                Set objPopup = mcbsThis.ActiveMenuBar.Controls(2)
            End If
        Else
            Set objPopup = mcbsThis.ActiveMenuBar.Controls(2)
        End If
        rptList.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
