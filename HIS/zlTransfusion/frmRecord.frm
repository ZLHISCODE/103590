VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecord 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList img16 
      Left            =   2610
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox pictbcKernel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   -15
      ScaleHeight     =   3735
      ScaleWidth      =   6450
      TabIndex        =   0
      Top             =   2505
      Width           =   6450
      Begin XtremeSuiteControls.TabControl tbcKernel 
         Height          =   3495
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   6210
         _Version        =   589884
         _ExtentX        =   10954
         _ExtentY        =   6165
         _StockProps     =   64
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1740
      Left            =   60
      TabIndex        =   2
      Top             =   315
      Width           =   6435
      _cx             =   11351
      _cy             =   3069
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   33023
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRecord.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.Label lbl说明 
      Alignment       =   1  'Right Justify
      Caption         =   "状态说明:  … 正在执行 √ 完成 × 拒绝"
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   2115
      Width           =   3420
   End
   Begin VB.Image img完成 
      Height          =   240
      Left            =   4905
      Picture         =   "frmRecord.frx":009B
      Top             =   15
      Width           =   240
   End
   Begin VB.Image img在执行 
      Height          =   240
      Left            =   4620
      Picture         =   "frmRecord.frx":68ED
      Top             =   15
      Width           =   240
   End
   Begin VB.Image img拒绝 
      Height          =   240
      Left            =   4365
      Picture         =   "frmRecord.frx":D13F
      Top             =   15
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   2175
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpSub 
      Bindings        =   "frmRecord.frx":13991
      Left            =   495
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum rptCOL
    rptCOL_执行分类 = 0
    rptCOL_接单时间 = 1
    rptCOL_配药人 = 2
    rptCOL_组数 = 3
    rptCOL_耗时 = 4
    rptCOL_滴系数 = 5
    rptCOL_接单人 = 6
    rptCOL_流水号 = 7
End Enum


Private Enum vsListCol
    col_完成情况 = 0
    col_序号 = 1
    col_顺序 = 2
    col_医嘱内容 = 3
    col_剂量 = 4
    col_单位 = 5
    col_金额 = 6
    col_执行频率 = 7
    col_用法 = 8
    col_皮试结果 = 9
    col_收费金额 = 10
    col_滴速 = 11
    col_容量 = 12
    col_时间 = 13
    col_进度 = 14
    col_执行人 = 15
    col_核对人 = 16
    col_剩余次数 = 17
    col_说明 = 18
    col_BillKey = 19
    col_groupkey = 20
    col_修改标志 = 21
End Enum

Private Const conMenu_File_BillPrintEx As Long = 3554        '输液瓶签、腕带

Private mlngType As Integer '显示类型 0-治疗 1-输液 2-注射 3-皮试

'项目附费
'--------
'Public WithEvents mclsExpenses As zlCISKernel.clsDockExpense
Public WithEvents mclsExpenses As zlPublicExpense.clsDockExpense
Attribute mclsExpenses.VB_VarHelpID = -1
'Private mclsPubExpense As zlPublicExpense.clsPublicExpense
Private mcolSubForm As Collection
Private mfrmActive As Form
'--------
Private mstrGroupKey As String      '当前选定的执行项目
Private mlng流水号 As Long          '当前选定的流水号
Private mlngModi As Long '是否修改状态
Private mstr执行人 As String

Private mblnUpdate As Boolean '是否修改过

Private mfrmMain As frmTransfusion
Private mcbsMain As CommandBars
Private mstrPatiStat As String  '主窗体当前人员的状态
Private marrRecord As Variant   '当前执行项目缓存

'Private mobjExecRecord As ExecRecord '病人项目类

Public Property Get 流水号() As Long
    流水号 = mlng流水号
End Property

Public Property Let 流水号(ByVal vData As Long)
    mlng流水号 = vData
End Property

Public Property Get 执行人() As String
     执行人 = mstr执行人
End Property

Public Property Let 执行人(ByVal vData As String)
    mstr执行人 = vData
End Property


Public Property Let 编辑(ByVal vData As Long)
    mlngModi = vData
End Property

Public Property Get 编辑() As Long
    编辑 = mlngModi
End Property

Public Property Let 组Key(ByVal vData As String)
    mstrGroupKey = vData
End Property

Public Property Get 组Key() As String
    组Key = mstrGroupKey
End Property

Public Property Let 修改过(ByVal vData As Boolean)
    mblnUpdate = vData
End Property

Public Property Get 修改过() As Boolean
    修改过 = mblnUpdate
End Property

Private Sub cbsSub_Resize()
    On Error Resume Next
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsList.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - lbl说明.Height
    lbl说明.Move lngLeft, vsList.Top + vsList.Height, vsList.Width - 45
End Sub

Private Sub dkpSub_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = pictbcKernel.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Call DockSubInit
    pictbcKernel.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
    
    marrRecord = Array()
    
    'TabControl
    '-------------
'    Set mclsExpenses = New zlCISKernel.clsDockExpense
'    Set mclsPubExpense = New zlPublicExpense.clsPublicExpense
    Set mclsExpenses = New zlPublicExpense.clsDockExpense
    Set mcolSubForm = New Collection

    '初始化
'    mclsPubExpense.zlInitCommon glngSys, gcnOracle
    mclsExpenses.zlInitCommon glngSys, gcnOracle

    mcolSubForm.Add mclsExpenses.zlGetForm, "_项目附费"
    With Me.tbcKernel
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        '1257 -项目附费管理
        If GetInsidePrivs(1257, True) <> "" Then
            .InsertItem(0, "项目附费", mcolSubForm("_项目附费").hwnd, 0).Tag = "项目附费"
            .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
           ' Call zlDefCommandBars(.Selected) '初始刷新定义一次菜单及按钮
        End If
    End With

End Sub

Private Sub DockSubInit()
    Dim objPaneA As Pane, objPaneB As Pane, ojbPaneC As Pane
    Dim lngX As Long
    Dim lngY As Long
    
    'DockingPane 初始化
    '-----------------------------------------------------
    Me.dkpSub.SetCommandBars Me.cbsSub
    
    If GetInsidePrivs(1257, True) <> "" Then
        Set objPaneA = Me.dkpSub.CreatePane(1, 280, 235, DockBottomOf)
        objPaneA.Title = "项目附费"
        objPaneA.Options = PaneNoCloseable Or PaneNoFloatable
    Else
        Me.pictbcKernel.Visible = False
    End If
    
    Me.dkpSub.Options.UseSplitterTracker = False '实时拖动
    Me.dkpSub.Options.ThemedFloatingFrames = True
    Me.dkpSub.Options.AlphaDockingContext = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mlng流水号 = 0
    mstrGroupKey = ""
    mlngModi = 0
    mstr执行人 = ""
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    
    Set mclsExpenses = Nothing
    Set marrRecord = Nothing
End Sub

Private Sub mclsExpenses_RequestRefresh()
    Call mfrmMain.刷新
End Sub

'Private Sub mclsExpenses_StatusTextUpdate(ByVal Text As String)
'    Call mfrmMain.更新状态栏(Text)
'End Sub

Private Sub pictbcKernel_Resize()
    On Error Resume Next
    With pictbcKernel
        tbcKernel.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    '主窗体调用本窗体的执行功能
    
    Dim lng流水号 As Long, lngDeptID As Long, strGroupKey As String
    Dim lngErrNo As Long
    
    Select Case Control.ID
        Case conMenu_Manage_ThingDel
            '撤消接单
            lng流水号 = mfrmMain.Get流水号
            If lng流水号 <> 0 Then
                If MsgBox("是否撤消流水号为" & lng流水号 & "的执行记录？", vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                    Call mfrmMain.撤消接单(lng流水号)
                End If
            End If
        Case conMenu_Edit_Transf_Modify
            '修改
            lng流水号 = mfrmMain.Get流水号
            If lng流水号 <> 0 Then
                Call RecordBuffer(1, Me.vsList)
                If mlngModi = 0 Then
                    mlngModi = lng流水号                '修改状态
                    Call mfrmMain.rptRecord_SelectionChanged
                Else
                    MsgBox "已有一张单据正在修改，不能同时修改多张单据！", vbInformation, gstrSysName
                End If
            End If
        Case conMenu_Edit_Transf_Save
            '保存
            'lng流水号 = Val(vsMain.TextMatrix(vsMain.Row, col_流水号))
            If mlngModi <> 0 Then
                '###
                Call SaveToExecRecord(mlngModi, lngErrNo)
                mlngModi = 0
                If lngErrNo = 0 Then
                    'Call mfrmMain.rptRecord_SelectionChanged
                    Call mfrmMain.rptPati_SelectionChanged
                Else
                    Call mfrmMain.rptRecord_SelectionChanged
                    Call RecordBuffer(2, Me.vsList)
                End If
            End If
        Case conMenu_File_BillPrint, conMenu_File_BillPrintEx
            '单据打印
            lng流水号 = mfrmMain.Get流水号
            If lng流水号 <> 0 Then
                If Control.ID = conMenu_File_BillPrintEx Then
                    Select Case Val(Control.Caption)
                    Case 1    '输液瓶签
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_4", mfrmMain, "接单流水号=" & lng流水号)
                    Case 2    '腕带
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_5", mfrmMain, "接单流水号=" & lng流水号)
                    End Select
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_" & mlngType, mfrmMain, "接单流水号=" & lng流水号, 2)
                End If
            End If
        Case conMenu_File_BillPrintView
            '单据预览
            lng流水号 = mfrmMain.Get流水号
            If lng流水号 <> 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_" & mlngType, mfrmMain, "接单流水号=" & lng流水号, 1)
            End If
'        Case conMenu_Edit_Transf_Positive
'            '阴性(-)
'            If mlngModi = 0 Then
'                lng流水号 = mfrmMain.Get流水号
'                If lng流水号 <> 0 And mstrGroupKey <> "" Then
'                    strGroupKey = mstrGroupKey
'                    Call mfrmMain.ExecuteTest(CStr(lng流水号), strGroupKey, "(-)")
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
'        Case conMenu_Edit_Transf_Negative
'            '阳性(+)
'            If mlngModi = 0 Then
'                lng流水号 = mfrmMain.Get流水号
'                If lng流水号 <> 0 And mstrGroupKey <> "" Then
'                    strGroupKey = mstrGroupKey
'                    Call mfrmMain.ExecuteTest(CStr(lng流水号), strGroupKey, "(+)")
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
        Case conMenu_Edit_Test
            '皮试结果
            If mlngModi = 0 Then
                lng流水号 = mfrmMain.Get流水号
                If lng流水号 <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    Call mfrmMain.ExecuteTest(CStr(lng流水号), strGroupKey)
                    Call mfrmMain.刷新
                End If
            End If
        Case conMenu_Edit_Transf_Cancle
            '取消
            If mlngModi <> 0 Then
                If MsgBox("是否取消已修改的内容？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    mlngModi = 0
                    Call mfrmMain.rptRecord_SelectionChanged
                End If
            End If
         Case conMenu_Manage_Undone
            '撤消完成
            If mlngModi = 0 Then
                lng流水号 = mfrmMain.Get流水号
                lngDeptID = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
                If lng流水号 <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    
                    If mfrmMain.mobjRecord.Item(CStr(lng流水号)).Item(strGroupKey).执行状态 = 1 Then
                        '已完成的，先改状态为执行中
                        If mfrmMain.mobjRecord.Item(CStr(lng流水号)).ExecCanle(strGroupKey, mfrmMain.mbln皮试验证, lngDeptID, Me) = False Then
                            Exit Sub
                        End If
                    End If
                    If (mfrmMain.mobjRecord.Item(CStr(lng流水号)).Item(strGroupKey).核对人 = "" And Mid(gstr医嘱核对, 2, 1) <> "1") Or (mfrmMain.mobjRecord.Item(CStr(lng流水号)).Item(strGroupKey).核对人 = "" And Mid(gstr医嘱核对, 2, 1) = "1" And mfrmMain.mobjRecord.Item(CStr(lng流水号)).Item(strGroupKey).执行状态 <> 1) Then
                        Call mfrmMain.ExecStart(CStr(lng流水号), strGroupKey, True)
                    End If
                    'Call mfrmMain.rptRecord_SelectionChanged
                    Call mfrmMain.刷新
                End If
            End If
        Case conMenu_Manage_Complete
            '完成 2012－09－10 改为开始功能，填写每条输液医嘱的开始时间，开始人,如是最后一次，则还要调原来的完成功能
            '剩余次数不为0，不能执行完成功能 ,objGroup.发送数次 - objGroup.已执行数次
            '
            If mlngModi = 0 Then
                lng流水号 = mfrmMain.Get流水号
                If lng流水号 <> 0 And mstrGroupKey <> "" Then
                    strGroupKey = mstrGroupKey
                    If mfrmMain.ExecStart(CStr(lng流水号), strGroupKey) Then
                    
                        If Val(vsList.TextMatrix(vsList.Row, col_剩余次数)) = 0 Then
                            
                            Call mfrmMain.ExecComplt(CStr(lng流水号), strGroupKey)
                            Call mfrmMain.刷新
                        Else
                            '----- 改成修改 病人医嘱执行 表的开始时间
                            'MsgBox "该项目还有剩余执行次数，还不能标记为完成！", vbInformation, gstrSysName
    
                            Call mfrmMain.刷新
                        End If
                    End If
                End If
            End If
'拒绝和取消拒绝,放到接单中处理
'        Case conMenu_Manage_Refuse
'            '拒绝执行
'            If mlngModi = 0 Then
'                lng流水号 = mfrmMain.Get流水号
'                If lng流水号 <> 0 And mstrGroupKey <> "" Then
'                    Call mfrmMain.mobjRecord.Item(CStr(流水号)).FuncExecRefuse(mstrGroupKey)
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
'
'        Case conMenu_Manage_ReGet
'            '取消拒绝
'            If mlngModi = 0 Then
'                lng流水号 = mfrmMain.Get流水号
'                If lng流水号 <> 0 And mstrGroupKey <> "" Then
'                    Call mfrmMain.mobjRecord.Item(CStr(流水号)).FuncExecRestore(mstrGroupKey)
'                    Call mfrmMain.rptRecord_SelectionChanged
'                End If
'            End If
        Case conMenu_Manage_ThingAudit '核对
            If mlngModi = 0 Then
                lng流水号 = mfrmMain.Get流水号
                strGroupKey = mstrGroupKey
                If lng流水号 <> 0 And strGroupKey <> "" Then
                    If mfrmMain.FuncThingAudit(CStr(lng流水号), strGroupKey) Then
                        mfrmMain.刷新
                    End If
                End If
            End If
        Case conMenu_Manage_ThingDelAudit '取消核对
            If mlngModi = 0 Then
                lng流水号 = mfrmMain.Get流水号
                strGroupKey = mstrGroupKey
                If lng流水号 <> 0 And strGroupKey <> "" Then
                    If mfrmMain.FuncThingDelAudit(CStr(lng流水号), strGroupKey) Then
                        mfrmMain.刷新
                    End If
                End If
            End If
        Case Else
            If Not tbcKernel.Selected Is Nothing Then
                If tbcKernel.Selected.Tag = "项目附费" Then
                    Call mclsExpenses.zlExecuteCommandBars(Control)
                End If
            End If
    End Select
    
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '主窗框体调用本窗体的弹出菜单
    Call mclsExpenses.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '主窗体调用本窗体的状态更新
    Dim blnVisable As Boolean, blnAllEnable As Boolean, blnOneEnable As Boolean, bln状态 As Boolean
    Dim intItem As Integer
    
    blnVisable = (mlngType = 3)
    'blnAllEnable = (mlng流水号 > 0 And InStr(mstrGroupKey, "_") <= 0 And mlngModi = 0) '按流水号
    blnAllEnable = (mlng流水号 > 0 And mlngModi = 0)  '按流水号
    blnOneEnable = (mlng流水号 > 0 And InStr(mstrGroupKey, "_") <> 0 And mlngModi = 0) '按执行项目
    
    If mfrmMain.mobjRecord Is Nothing Then
        blnAllEnable = False
        blnOneEnable = False
    End If
    
    Select Case Control.ID
        Case conMenu_Manage_ThingDel, conMenu_File_BillPrint, conMenu_File_BillPrintView, conMenu_File_BillPrintEx
            '撤消接单
            Control.Enabled = blnAllEnable

            If Control.Enabled Then
                 
                For intItem = 1 To mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Count
                    If mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(intItem).执行状态 = 1 Then
                        Control.Enabled = False
                        Exit For
                    End If
                    If blnVisable And Mid(gstr医嘱核对, 2, 1) = "1" And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 <> "" Then
                        Control.Enabled = False
                        Exit For
                    End If
                Next
     
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
        Case conMenu_Edit_Transf_Modify
            '修改
            Control.Enabled = mlngModi = 0 And mlng流水号 > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                    If blnVisable And Mid(gstr医嘱核对, 2, 1) = "1" Then
                        If Control.Enabled Then Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 = ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
        Case conMenu_Edit_Transf_Save
            '保存
            Control.Enabled = mlngModi <> 0 And mlng流水号 > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
        Case conMenu_Edit_Transf_Cancle
            '取消
            Control.Enabled = mlngModi <> 0 And mlng流水号 > 0
            If InStr(mstrGroupKey, "_") <= 0 Then Control.Enabled = False
            If Control.Enabled Then
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
'        Case conMenu_Manage_Refuse
'            '拒绝
'
'        Case conMenu_Manage_ReGet
'            '取消拒绝
        Case conMenu_Manage_Complete
            '完成--  改成了开始
            Control.Visible = Not blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                    
                    If Control.Enabled Then
                        Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行人 = ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
        Case conMenu_Manage_Undone
            '取消完成
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And (mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行人 <> "" _
                      Or mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 = 1)
                    If blnVisable And Mid(gstr医嘱核对, 2, 1) = "1" Then
                         Control.Enabled = (mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 <> "" And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 = 1) Or (mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 = "" And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 = 3)
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"

'        Case conMenu_Edit_Transf_Negative
'            '阳性
'            Control.Visible = blnVisable
'            If InStr(mstrGroupKey, "_") <= 0 Then
'                Control.Enabled = False
'            Else
'                If Not mfrmMain.mobjRecord Is Nothing Then
'                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
'                Else
'                    Control.Enabled = False
'                End If
'            End If
'            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
'        Case conMenu_Edit_Transf_Positive
'            '阴性
'            Control.Visible = blnVisable
'            If InStr(mstrGroupKey, "_") <= 0 Then
'                Control.Enabled = False
'            Else
'                If Not mfrmMain.mobjRecord Is Nothing Then
'                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
'                Else
'                    Control.Enabled = False
'                End If
'            End If
'            If Control.Enabled Then Control.Enabled = mstrPatiStat <> "3-退号"
        Case conMenu_Edit_Test
            '皮试结果
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                    If blnVisable And Mid(gstr医嘱核对, 2, 1) = "1" Then
                        If Control.Enabled Then Control.Enabled = mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 <> ""
                    End If
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Manage_ThingAudit
            '医嘱核对
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1 And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 = ""
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Manage_ThingDelAudit
            '医嘱取消核对
            Control.Visible = blnVisable
            If InStr(mstrGroupKey, "_") <= 0 Then
                Control.Enabled = False
            Else
                If Not mfrmMain.mobjRecord Is Nothing Then
                    Control.Enabled = blnOneEnable And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).核对人 <> "" And mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(mstrGroupKey).执行状态 <> 1
                Else
                    Control.Enabled = False
                End If
            End If
        Case Else
            If Not tbcKernel.Selected Is Nothing Then
                If tbcKernel.Selected.Tag = "项目附费" Then
                    Call mclsExpenses.zlUpdateCommandBars(Control)
                End If
            End If
    End Select

End Sub

Public Sub zlRefresh(ByVal objRecord As ExecRecord, ByVal objPati As cPatient)
    '主窗体要求本窗体刷新
    Call mfrmMain.ShowReport
    Call ShowVsList(mfrmMain.Get流水号)
    Call KernalRefresh
    
    '执行人
    mstr执行人 = ""
    
    Dim ObjOutNurse As New OutNurses, objNurs As OutNurse
    
    If mstr执行人 = "" And mfrmMain.cboDept.ListIndex >= 0 Then
        ObjOutNurse.getOutNurse (mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex))
        For Each objNurs In ObjOutNurse
            mstr执行人 = mstr执行人 & "|" & objNurs.姓名
        Next
        If Mid(mstr执行人, 1, 1) = "|" Then mstr执行人 = Mid(mstr执行人, 2)
    End If
    '当前病人的状态
    If Not objPati Is Nothing Then
        mstrPatiStat = objPati.排队状态
    Else
        mstrPatiStat = ""
    End If
End Sub

Public Sub KernalRefresh()
    Dim strInfo As String

    If mstrGroupKey = "" Then
        Call mclsExpenses.zlRefresh(0, "")
    Else
        '科室ID,医嘱ID,发送号
        strInfo = Trim(Split(mstrGroupKey, "_")(0)) & ":" & Split(mstrGroupKey, "_")(1)
'        On Error Resume Next
'        Call mclsExpenses.zlRefresh(mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex), _
                                    Val(Split(mstrGroupKey, "_")(0)), Val(Split(mstrGroupKey, "_")(1)), False)
        Call mclsExpenses.zlRefresh(mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex), _
                                    strInfo)
    End If
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    '主窗体要求初始化主窗体上的菜单
    
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '病人项目的菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "执行(&E)", objMenu.Index + 1, False)
    End If
    
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "撤消接单(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Modify, "修改(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintEx, "1-打印输液瓶签"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintEx, "2-打印腕带标签")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrint, "重打单据(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_BillPrintView, "单据预览(&V)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒绝执行(&R)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "取消拒绝(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "开始执行(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "撤消执行(&U)")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Negative, "阳性(&+)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Positive, "阴性(&-)")
        If Mid(gstr医嘱核对, 2, 1) = "1" Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "取消核对")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "皮试结果(&H)"): objControl.BeginGroup = True

    End With
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
'    If objMenu Is Nothing Then
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
'        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
'    End If
'    With objMenu.CommandBar.Controls
'        '子项放在最前面,反序加入
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Reprint, "重打单据(&R)", 1)
'    End With
    
    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Modify, "修改", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "执行", objControl.Index + 1): objControl.BeginGroup = True
        objControl.ToolTipText = "开始执行"
    End With
    Set mfrmMain = frmParent
    cbsSub.ActiveMenuBar.Visible = False
    '--项目附费
    Call mclsExpenses.zlDefCommandBars(frmParent, cbsMain)
    Set mcbsMain.Icons = zlCommFun.GetPubIcons


End Sub

Public Sub ShowVsList(ByVal lng流水号 As Long)
    
    Dim objGroup As Group, objBIll As Bill, lng序号 As Long, strListHead As String, str状态 As String
    Dim date开始时间 As Date, lng已开始 As Long
    
    If lng流水号 = 0 Then
        mlngType = 0
    Else
        mlngType = Mid(mfrmMain.mobjRecord.Item(CStr(lng流水号)).执行分类, 1, 1)
    End If
    
    Select Case mlngType
    Case 1
        '输液
        strListHead = "状态,450,4;序号,0,4;顺序,450,4;内容,2300,1;单量,600,7;单位,0,4;金额,700,7;执行频率,900,1;用法,900,1;皮试结果,0,1;输液费,700,7;滴速,450,7;容量(ml),450,7;时间(分),450,4;进度,1300,1;执行人,675,1;核对人,0,1;剩余次数,450,7;备注,500,1;billKey,0,1;GroupKey,0,1;修改标志,0,1"
    Case 2
        '注射
        strListHead = "状态,450,4;序号,0,4;顺序,0,4;内容,2300,1;剂量,600,7;单位,0,4;金额,700,7;执行频率,900,1;用法,900,1;皮试结果,0,1;注射费,700,7;滴速,0,7;容量(ml),0,7;时间(分),0,4;进度,0,1;执行人,675,1;核对人,0,1;剩余次数,450,7;备注,500,1;billKey,0,1;GroupKey,0,1;修改标志,0,1"
    Case 3
        '皮试
        If Mid(gstr医嘱核对, 2, 1) = "1" Then
            strListHead = "状态,450,4;序号,0,4;顺序,0,4;内容,2300,1;剂量,0,7;单位,0,4;金额,0,7;执行频率,900,1;用法,1800,1;皮试结果,450,1;皮试费,700,7;滴速,0,7;容量(ml),0,7;时间(分),0,4;进度,1300,1;执行人,675,1;核对人,675,1;剩余次数,450,7;备注,500,1;billKey,0,1;GroupKey,0,1;修改标志,0,1"
        Else
            strListHead = "状态,450,4;序号,0,4;顺序,0,4;内容,2300,1;剂量,0,7;单位,0,4;金额,0,7;执行频率,900,1;用法,1800,1;皮试结果,450,1;皮试费,700,7;滴速,0,7;容量(ml),0,7;时间(分),0,4;进度,1300,1;执行人,675,1;核对人,0,1;剩余次数,450,7;备注,500,1;billKey,0,1;GroupKey,0,1;修改标志,0,1"
        End If
    Case Else
        '治疗
        '1 左对齐 4 居中 7 右对齐
        strListHead = "状态,450,4;序号,0,4;顺序,0,4;内容,2300,1;单量,600,7;单位,0,4;金额,0,7;执行频率,900,1;用法,900,1;皮试结果,0,1;治疗费,700,7;滴速,0,7;容量(ml),0,7;时间(分),0,4;进度,0,1;执行人,675,1;核对人,0,1;剩余次数,450,7;备注,500,1;billKey,0,1;GroupKey,0,1;修改标志,0,1"
    End Select
    Call SetVsFlexGridHead(strListHead, vsList)
    If lng流水号 = 0 Then Exit Sub
    If mfrmMain.mobjRecord Is Nothing Then Exit Sub
    'vsList.Redraw = False
    date开始时间 = mfrmMain.mobjRecord.Item(CStr(lng流水号)).执行时间
    lng已开始 = DateDiff("n", date开始时间, zlDatabase.Currentdate)
    
    For Each objGroup In mfrmMain.mobjRecord.Item(CStr(lng流水号))
        lng序号 = 0
        date开始时间 = date开始时间 + (objGroup.耗时 / 24 / 60)
        
        
        With vsList
            For Each objBIll In objGroup.BillsItem(objGroup.执行医嘱ID & "_" & objGroup.发送号)
                lng序号 = lng序号 + 1
                '0-未执行;1-完全执行;2-拒绝执行;3-正在执行
                Select Case objGroup.执行状态
                    Case 0
                        str状态 = ""
                    Case 1
                        str状态 = "完成"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_完成情况) = img完成.Picture
                        .TextMatrix(.Rows - 1, col_完成情况) = "√"
                    Case 2
                        str状态 = "拒绝"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_完成情况) = img拒绝.Picture
                        .TextMatrix(.Rows - 1, col_完成情况) = "×"
                    Case Else
                        str状态 = "执行中"
                        .TextMatrix(.Rows - 1, col_完成情况) = "…"
                        'Set .Cell(flexcpPicture, .Rows - 1, col_完成情况) = img在执行.Picture
                End Select
                '.TextMatrix(.Rows - 1, col_完成情况) = str状态
                
                .TextMatrix(.Rows - 1, col_序号) = lng序号
                .TextMatrix(.Rows - 1, col_顺序) = Val(objGroup.组次)
                .TextMatrix(.Rows - 1, col_医嘱内容) = objBIll.医嘱内容
                .TextMatrix(.Rows - 1, col_剂量) = IIf(Mid(objBIll.单量, 1, 1) = ".", "0" & objBIll.单量 & objBIll.单位, objBIll.单量 & objBIll.单位)
                .TextMatrix(.Rows - 1, col_单位) = objBIll.单位
                .TextMatrix(.Rows - 1, col_金额) = IIf(Format(objBIll.金额, "0.00") = 0, "", Format(objBIll.金额, "0.00"))
                If objBIll.明细计费状态 = -1 Then
                    .TextMatrix(.Rows - 1, col_金额) = "不计费"
                ElseIf objBIll.明细计费状态 = -2 Then
                    If objBIll.金额 = 0 Then .TextMatrix(.Rows - 1, col_金额) = "零费用"
                ElseIf objBIll.明细计费状态 = -3 Then
                    .TextMatrix(.Rows - 1, col_金额) = "已退费"
                End If
                .TextMatrix(.Rows - 1, col_执行频率) = objGroup.执行频次
                .TextMatrix(.Rows - 1, col_用法) = objGroup.用法
                .TextMatrix(.Rows - 1, col_皮试结果) = objGroup.皮试结果
                .TextMatrix(.Rows - 1, col_收费金额) = IIf(Format(objGroup.收费金额, "0.00") = 0, "", Format(objGroup.收费金额, "0.00"))
                If objGroup.计费状态 = -1 Then
                    .TextMatrix(.Rows - 1, col_收费金额) = "不计费"
                ElseIf objGroup.计费状态 = -2 Then
                    If objGroup.收费金额 = 0 Then .TextMatrix(.Rows - 1, col_收费金额) = "零费用"
                ElseIf objGroup.计费状态 = -3 Then
                    .TextMatrix(.Rows - 1, col_收费金额) = "已退费"
                End If
                .TextMatrix(.Rows - 1, col_滴速) = objGroup.滴速
                .TextMatrix(.Rows - 1, col_容量) = objGroup.液体量
                .TextMatrix(.Rows - 1, col_时间) = objGroup.耗时
                
                '.TextMatrix(.Rows - 1, col_进度) = Format(date开始时间, "MM-dd hh:mm")
                
                If lng已开始 >= objGroup.耗时 Then
                    lng已开始 = lng已开始 - objGroup.耗时
                    .Cell(flexcpData, .Rows - 1, col_进度) = 100
                    '.Cell(flexcpFloodPercent, .Rows - 1, col_进度) = 100
                    '.Cell(flexcpFloodColor, .Rows - 1, col_进度) = RGB(215, 215, 235)
                    
                Else
                    If lng已开始 >= 0 Then
                        .Cell(flexcpData, .Rows - 1, col_进度) = (lng已开始 / objGroup.耗时) * 100
                        '.Cell(flexcpFloodPercent, .Rows - 1, col_进度) = (lng已开始 / objGroup.耗时) * 100
                        '.Cell(flexcpFloodColor, .Rows - 1, col_进度) = RGB(215, 215, 235)
                        lng已开始 = lng已开始 - objGroup.耗时
                    End If
                End If
                .TextMatrix(.Rows - 1, col_说明) = IIf(objBIll.医生嘱托 = "", "　", objBIll.医生嘱托)
                .TextMatrix(.Rows - 1, col_执行人) = IIf(objGroup.执行人 = "", "　", objGroup.执行人)
                .TextMatrix(.Rows - 1, col_核对人) = IIf(objGroup.核对人 = "", "　", objGroup.核对人)
                .TextMatrix(.Rows - 1, col_剩余次数) = objGroup.发送数次 - objGroup.已执行数次 - objGroup.本次数次
                .TextMatrix(.Rows - 1, col_BillKey) = objGroup.执行医嘱ID & "_" & objBIll.医嘱ID
                .TextMatrix(.Rows - 1, col_groupkey) = objGroup.执行医嘱ID & "_" & objGroup.发送号
                
                '字段颜色；皮试类（阳性：红色；阴性：蓝色）；非皮试类黑色；
                If InStr(objGroup.皮试结果, "(-)") > 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HFF0000
                ElseIf InStr(objGroup.皮试结果, "(+)") > 0 Or InStr(objGroup.皮试结果, "(++)") > 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0&
                End If
                                
                .Rows = .Rows + 1
                '.TextMatrix(.Rows - 1, col_groupkey) = "要隐藏"
            Next
            '每组药之间加一空行,避免相同内容被合并为一个单元格
            '.Rows = .Rows + 1
'            If .TextMatrix(.Rows - 2, col_groupkey) = "要隐藏" Then
'                .RowHidden(.Rows - 2) = True
'            End If
        End With
    Next
    If vsList.Rows > 2 Then
        vsList.RemoveItem vsList.Rows - 1
    End If
    'vsMain.Redraw = True
    With vsList
'        .MergeCells = flexMergeRestrictColumns
'        .MergeCol(col_完成情况) = True
'        .MergeCol(col_执行频率) = True
'        .MergeCol(col_用法) = True
'        .MergeCol(col_滴速) = True
'        .MergeCol(col_容量) = True
'        .MergeCol(col_时间) = True
'        .MergeCol(col_剩余次数) = True
'        .MergeCol(col_执行人) = True
'        .MergeCol(col_进度) = True
'        .MergeCol(col_说明) = True
        .AutoSize 1, col_groupkey
        .RowHeight(0) = 500
        '.BackColorSel = .BackColor
        '.ForeColorSel = .ForeColor
        '.SelectionMode = flexSelectionFree
    End With
    
    vsList.ColDataType(col_滴速) = flexDTLong
    vsList.ColDataType(col_容量) = flexDTLong
    
    If mlngModi <> 0 Then
        '修改状态
        vsList.Cell(flexcpBackColor, 1, col_容量, vsList.Rows - 1, col_容量) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_滴速, vsList.Rows - 1, col_滴速) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_说明, vsList.Rows - 1, col_说明) = VsModiBackColor
        vsList.Cell(flexcpBackColor, 1, col_执行人, vsList.Rows - 1, col_执行人) = VsModiBackColor
        
        If mstr执行人 <> "" Then
            vsList.ColComboList(col_执行人) = mstr执行人
        End If
        vsList.Editable = flexEDKbdMouse
        
        'vsMain.Cell(flexcpForeColor, vsMain.Row, 0, vsMain.Row, vsMain.Cols - 1) = vbRed
        mblnUpdate = True
    Else
        '查看状态
        
        vsList.Cell(flexcpBackColor, 1, col_容量, vsList.Rows - 1, col_容量) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_滴速, vsList.Rows - 1, col_滴速) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_说明, vsList.Rows - 1, col_说明) = vsList.BackColor
        vsList.Cell(flexcpBackColor, 1, col_执行人, vsList.Rows - 1, col_执行人) = vsList.BackColor
        vsList.Editable = flexEDNone
        'vsMain.Cell(flexcpForeColor, vsMain.Row, 0, vsMain.Row, vsMain.Cols - 1) = vsMain.ForeColor
        mblnUpdate = False
    End If
    vsList_RowColChange
End Sub


Private Sub tbcKernel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsList.AutoSize 1, col_groupkey
    vsList.RowHeight(0) = 500
End Sub

Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If col_groupkey < vsList.Cols Then
        If mfrmMain.mobjRecord.Item(CStr(mlng流水号)).Item(vsList.TextMatrix(Row, col_groupkey)).执行状态 = 1 Then
            '完成的记录不能修改
            Cancel = True
            Exit Sub
        End If
    End If
    If InStr("," & col_容量 & "," & col_滴速 & "," & col_说明 & "," & col_执行人 & ",", _
             "," & Col & ",") <= 0 Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub vsList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim LeftCol As Long, RightCol As Long, topRow As Long, BottomRow As Long
    
    If Col = col_进度 Then
        If Val(vsList.TextMatrix(Row, col_序号)) = 1 And Trim(vsList.TextMatrix(Row, col_执行人)) <> "" Then
            With vsList
                Call vfgDrawProgress(vsList, Row, Col, hDC, Left, Top, Right, Bottom, .Cell(flexcpData, Row, Col))
            End With
        End If
    End If
    
    
    If Not MergeRow(Row, topRow, BottomRow) Then Exit Sub '非合并行,退出
    If topRow = BottomRow Then Exit Sub
    
    LeftCol = col_顺序: RightCol = col_顺序
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_完成情况: RightCol = col_完成情况
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_执行频率: RightCol = col_用法
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)

    LeftCol = col_滴速: RightCol = col_剩余次数
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_说明: RightCol = col_说明
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
    LeftCol = col_收费金额: RightCol = col_收费金额
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsList)
    
End Sub

Private Function MergeRow(ByVal Row As Long, topRow, BottomRow As Long) As Boolean
    '是否合并行
    Dim strGroupKey As String, lngRow As Long
    With vsList
        If .Cols < col_groupkey Then Exit Function
        strGroupKey = .TextMatrix(Row, col_groupkey)
        topRow = Row: BottomRow = Row
        For lngRow = Row To 0 Step -1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                topRow = lngRow + 1
                Exit For
            Else
                topRow = lngRow
            End If
        Next
        
        For lngRow = Row To .Rows - 1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                BottomRow = lngRow - 1
                Exit For
            Else
                BottomRow = lngRow
            End If
        Next
    End With

    If topRow > 0 And BottomRow > 0 Then MergeRow = True
End Function

Private Sub vsList_EnterCell()
    If mlngModi = 1 Then
        If vsList.Col = col_容量 Or vsList.Col = col_滴速 Or vsList.Col = col_说明 Or vsList.Col = col_执行人 Then
            Call vsList.CellBorder(vsList.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End If
End Sub

Private Sub vsList_GotFocus()
    If col_groupkey < vsList.Cols Then
        mstrGroupKey = Trim(vsList.TextMatrix(vsList.Row, col_groupkey))
    Else
        mstrGroupKey = ""
    End If
End Sub

Private Sub vsList_LeaveCell()
    If mlngModi = 1 Then
        If vsList.Col = col_容量 Or vsList.Col = col_滴速 Or vsList.Col = col_说明 Or vsList.Col = col_执行人 Then
            On Error Resume Next
            Call vsList.CellBorder(vsList.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End If
End Sub

Private Sub vsList_LostFocus()
    mstrGroupKey = ""
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsList_RowColChange()
    If col_groupkey < vsList.Cols Then
        
        If mstrGroupKey <> Trim(vsList.TextMatrix(vsList.Row, col_groupkey)) Then
            mstrGroupKey = Trim(vsList.TextMatrix(vsList.Row, col_groupkey))
            Call KernalRefresh
        End If
    Else
        mstrGroupKey = ""
    End If
    
    '保持选中的当前行字体前景色与原字体前景色相同
    If vsList.Row > 0 Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, vsList.Row, 0)
    End If
End Sub

Private Sub SaveToExecRecord(ByVal str流水号 As String, Optional ByRef lngErrNo_Out As Long)
    '保存修改的信息
    Dim iRow As Integer, strGroupKey As String, strBillKey As String, int执行状态 As Integer, blnLoad完成 As Boolean
    Dim lngDeptID As Long
    Dim cnNew As ADODB.Connection, strUserName As String
    Dim lngErrNo As Long
    
    If str流水号 = "" Then Exit Sub
    '11471 解决修改时，没有按回车,则不保存修改的内容。
    If vsList.Col < vsList.Cols - 1 Then
        vsList.Select vsList.Row, vsList.Col + 1
    Else
        vsList.Select vsList.Row, vsList.Col - 1
    End If
    
    For iRow = 1 To vsList.Rows - 1
        
        If vsList.TextMatrix(iRow, col_修改标志) = "Update" Then
            strGroupKey = vsList.TextMatrix(iRow, col_groupkey)
            
            If vsList.TextMatrix(iRow, col_序号) = 1 Then
                strBillKey = vsList.TextMatrix(iRow, col_BillKey)
                mfrmMain.mobjRecord.Item(str流水号).Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).容量 = Val(vsList.TextMatrix(iRow, col_容量))
                mfrmMain.mobjRecord.Item(str流水号).Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).时间 = Val(vsList.TextMatrix(iRow, col_时间))
            End If
            
            mfrmMain.mobjRecord.Item(str流水号).Item(strGroupKey).滴速 = Val(vsList.TextMatrix(iRow, col_滴速))
            mfrmMain.mobjRecord.Item(str流水号).Item(strGroupKey).说明 = Replace(vsList.TextMatrix(iRow, col_说明), "　", "")
            mfrmMain.mobjRecord.Item(str流水号).Item(strGroupKey).执行人 = Replace(vsList.TextMatrix(iRow, col_执行人), "　", "")
            
            lngDeptID = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
            Call mfrmMain.mobjRecord.Item(str流水号).Update(str流水号, strGroupKey, lngDeptID, lngErrNo)
            
            If lngErrNo <> 0 Then
                lngErrNo_Out = lngErrNo
                Exit Sub
            End If
            
            If str流水号 = mfrmMain.Get流水号 Then
                mfrmMain.rptRecord.SelectedRows(0).Record(rptCOL_耗时).Value = mfrmMain.mobjRecord.Item(str流水号).总耗时
                mfrmMain.rptRecord.SelectedRows(0).Record(rptCOL_耗时).Caption = mfrmMain.mobjRecord.Item(str流水号).总耗时
                mfrmMain.rptRecord.Populate
            End If
        End If
    Next

End Sub

Private Sub vsList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnExit As Boolean, i As Integer
    
    Select Case Col
    Case col_完成情况
        If Val(vsList.TextMatrix(Row, col_剩余次数)) <> 0 And vsList.EditText = "完成" Then
            Cancel = True
            MsgBox "剩余次数不为0，现在还不能完成！", vbInformation, gstrSysName
            Exit Sub
        End If
    Case col_容量
        i = 0: blnExit = False
        If Val(Trim(vsList.TextMatrix(Row + i, col_序号))) = 0 Then Exit Sub
        If Val(vsList.EditText) < 0 Or Val(vsList.EditText) > 10000 Then
            Cancel = True: Exit Sub
        End If
        Do While blnExit = False
            If blnExit = False Then
                If Row + i >= vsList.Rows Then
                    blnExit = True
            
                ElseIf (Val(vsList.TextMatrix(Row + i, col_序号)) = 1 Or Val(vsList.TextMatrix(Row + i, col_序号)) = 0) And i > 0 Then
                    blnExit = True
                Else
                    vsList.TextMatrix(Row + i, col_时间) = CacleTransTime(Val(vsList.EditText), _
                                                                     mfrmMain.mobjRecord.Item(CStr(mlng流水号)).滴系数, _
                                                                     Val(vsList.TextMatrix(Row, col_滴速)))
                    i = i + 1
                End If
            End If
        Loop
    Case col_滴速
        If Val(vsList.EditText) < 10 Or Val(vsList.EditText) > 100 Then
             Cancel = True: Exit Sub
        End If
        i = 0: blnExit = False
        If Val(Trim(vsList.TextMatrix(Row + i, col_序号))) = 0 Then Exit Sub
        Do While blnExit = False
            If blnExit = False Then
                If Row + i >= vsList.Rows Then
                    blnExit = True
                ElseIf (Val(vsList.TextMatrix(Row + i, col_序号)) = 1 Or Val(vsList.TextMatrix(Row + i, col_序号)) = 0) And i > 0 Then
                    blnExit = True
                Else
                    
                    vsList.TextMatrix(Row + i, col_时间) = CacleTransTime(Val(vsList.TextMatrix(Row, col_容量)), _
                                                                      mfrmMain.mobjRecord.Item(CStr(mlng流水号)).滴系数, _
                                                                     Val(vsList.EditText))
                    i = i + 1
                End If
            End If
        Loop
    Case col_皮试结果
        If InStr(",(+),(-),免试", vsList.TextMatrix(Row, col_皮试结果)) > 0 Then
            Cancel = True
            MsgBox "已填写结果的记录不允许修改!", vbInformation, gstrSysName
        End If
    End Select
    vsList.TextMatrix(Row, col_修改标志) = "Update"
End Sub

Private Sub RecordBuffer(ByVal bytMode As Byte, ByVal vsfVal As VSFlexGrid)
'功能：缓存当前执行项目记录，或恢复修改前的执行项目记录信息
    Dim i As Integer
    
    If bytMode = 2 Then
        '恢复
        For i = 0 To UBound(marrRecord)
            vsfVal.TextMatrix(vsfVal.Row, i) = marrRecord(i)
        Next
    Else
        '缓存
        If UBound(marrRecord) < 0 Then ReDim Preserve marrRecord(vsfVal.Cols - 1)
        For i = 0 To UBound(marrRecord)
            marrRecord(i) = vsfVal.TextMatrix(vsfVal.Row, i)
        Next
    End If
End Sub




