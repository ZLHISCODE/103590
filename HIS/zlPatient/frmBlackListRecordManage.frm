VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "ZLIDKIND.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmBlackListRecordManage 
   BorderStyle     =   0  'None
   Caption         =   "病人不良记录"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgList16 
      Left            =   6120
      Top             =   7740
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
            Picture         =   "frmBlackListRecordManage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackListRecordManage.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5310
      Left            =   465
      ScaleHeight     =   5310
      ScaleWidth      =   9060
      TabIndex        =   0
      Top             =   1725
      Width           =   9060
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3705
         _Version        =   589884
         _ExtentX        =   6535
         _ExtentY        =   2514
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTitle 
         Height          =   360
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   7905
         _Version        =   589884
         _ExtentX        =   13944
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "病人不良记录"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   5040
         Top             =   720
         Width           =   405
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGridPrint 
      Height          =   555
      Left            =   12990
      TabIndex        =   3
      Top             =   2055
      Visible         =   0   'False
      Width           =   645
      _cx             =   1961559154
      _cy             =   1961558995
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin zlIDKind.PatiIdentify patiFind 
      Height          =   345
      Left            =   10620
      TabIndex        =   4
      Top             =   375
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBlackListRecordManage.frx":0B34
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
End
Attribute VB_Name = "frmBlackListRecordManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String
Private mstr行为类别 As String
 
Private Enum mEnm_RptHeadCol
    COL_ID = 0
    COL_图标
    COL_行为类别
    COL_病人姓名
    COL_年龄
    COL_性别
    COL_出生日期
    COL_门诊号
    COL_发生时间
    COL_加入原因
    COL_加入详细说明
    COL_加入时间
    COL_附加信息
    COL_登记人
    COL_撤消人
    COL_撤消时间
    COL_撤消原因
    COL_是否固定
End Enum
Private mblnShowCancelRecord As Boolean '是否显示已撤消的不良记录
Private mlngPreSelID As Long
Private mintFindType As Integer
Private mrs不良记录 As Recordset
Private mcllFilter As Collection    '过滤条件

Public Event zlActivate(ByVal frmSubForm As Form) '事件触发
Public Event zlShowStatusText(ByVal bytPancel As Byte, ByVal strText As String)  '显示状态栏文本
Public Sub zlCancelBands()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:控件解绑
    '编制:刘兴洪
    '日期:2018-11-15 15:48:53
    '主要是在重建前，删除控件后，可能存在绑定的控件还在工具栏这个容器中，造成删除时，会儿控件一并删除
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrToolBar As CommandBar
    On Error GoTo errHandle
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "工具栏")
    If cbrToolBar Is Nothing Then Exit Sub
    cbrToolBar.Controls.DeleteAll
    Set patiFind.Container = Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Sub zlInitComm(frmMain As Form, cbsThis As Object, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口
    '入参:objPati-调用主窗口
    '     cbsThis-菜单对象
    '     strPrivs-权限串
    '     lngModule-模块号
    '编制:刘兴洪
    '日期:2018-11-08 11:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain: Set mcbsMain = cbsThis
    mstrPrivs = strPrivs: mlngModule = lngModule
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlLoadData(ByVal str行为类别 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-13 15:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    On Error GoTo errHandle
    mstr行为类别 = str行为类别
    zlLoadData = LoadRecordDataToGrid(str行为类别, mcllFilter)
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHandle
 
 
 
    '     '文件菜单
    '    '-----------------------------------------------------
    '    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    '    With cbrMenuBar.CommandBar.Controls
    '        '放在输出到Excel之后
    '        Set cbrControl = .Find(, conMenu_File_Excel)
    ''        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
    '    End With
    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加不良记录(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改不良记录(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除不良记录(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "撤消不良记录(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消撤消不良记录(&T)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤(&F)", cbrControl.Index): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示已撤消的不良记录(&S)", cbrControl.Index)
        cbrControl.Checked = mblnShowCancelRecord
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "工具栏")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = mcbsMain.Add("工具栏", xtpBarTop)
    End If
    
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup And cbrControl.Index > 1 Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "撤消", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消撤消", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤", cbrControl.Index + 1): cbrControl.BeginGroup = True
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
 
    '被绑定的控件必须动态加载，因为工具栏一但被删除，被绑定的控件的句柄就会变成0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    objCustom.Handle = patiFind.hwnd
    objCustom.flags = xtpFlagRightAlign
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsAllowOperation(ByVal intOperationType As Byte, Optional ByRef strID_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前的记录是否允许操作
    '入参:intOperationType-0-修改;1-删除;2-撤消;3-取消撤消;
    '返回:true-允许;False-不允许
    '编制:刘兴洪
    '日期:2018-11-09 11:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln是否固定 As Boolean, bln撤消 As Boolean, bln其他类型 As Boolean
    On Error GoTo errHandle
    strID_Out = ""
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    strID_Out = rptData.SelectedRows(0).Record(COL_ID).Value
    If strID_Out = "" Then Exit Function
    
    '143491:李南春,2019/8/7，允许对系统固定的【其他】行为类别增删改
    bln是否固定 = rptData.SelectedRows(0).Record(COL_是否固定).Value = "√"
    bln其他类型 = bln是否固定 And rptData.SelectedRows(0).Record(COL_行为类别).Value = "其他"
    bln撤消 = rptData.SelectedRows(0).Record(COL_撤消时间).Value <> ""
    Select Case intOperationType
    Case 0 '修改
        If bln是否固定 And Not bln其他类型 Then Exit Function
        IsAllowOperation = Not bln撤消
    Case 1 '删除
        If bln是否固定 And Not bln其他类型 Then Exit Function
         IsAllowOperation = Not bln撤消
    Case 2 '撤消
         IsAllowOperation = Not bln撤消
    Case 3 '取消撤消
         IsAllowOperation = bln撤消
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnable As Boolean
    Dim blnStop As Boolean '是否已停用
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next

    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "编辑不良记录")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(0)   '修改
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(1)   '删除
    Case conMenu_Edit_Stop
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(2)   '撤消
    Case conMenu_Edit_Reuse
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(3)   '取消撤消：0-修改;1-删除;2-撤消;3-取消撤消;
    Case conMenu_View_ShowStoped '显示已撤消的记录
        Control.Checked = mblnShowCancelRecord
    Case conMenu_View_Filter '过虑
    End Select
End Sub

Private Function ExecuteAddItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行增加不良记录操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    On Error GoTo errHandle
    'bytEditType-编辑类别:0-新增;1-修改;2-撤消;3-取消撤消;4-查看
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_增加, mstr行为类别) Then Exit Function
    
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    ExecuteAddItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteModifyItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行修改不良记录操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    If IsAllowOperation(0, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_撤消时间).Value) <> "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录已经撤档,不允许修改!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
        If Trim(rptData.SelectedRows(0).Record(COL_是否固定).Value) <> "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录是系统自动生成的，你只能撤消不良记录,但不允许修改!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    If strID = "" Then
        MsgBox "当前未选中要删除的不良记录，不能进行修改操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'bytEditType-编辑类别:0-新增;1-修改;2-撤消;3-取消撤消;4-查看
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_修改, mstr行为类别, strID) Then Exit Function

    
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    ExecuteModifyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行过滤操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmFilter As New frmBlackListRecordFilter
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    If Not frmFilter.zlShowEdit(mfrmMain, mlngModule, mcllFilter) Then Exit Function
    
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    ExecuteFilter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteStopItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行撤消不良记录操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    '0-修改;1-删除;2-撤消;3-取消撤消;
    If IsAllowOperation(2, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_撤消时间).Value) <> "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录已经撤档,不允许再次撤消!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "当前未选中要撤消的不良记录，不能进行撤消操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'bytEditType-编辑类别:0-新增;1-修改;2-撤消;3-取消撤消;4-查看
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_撤消, mstr行为类别, strID) Then Exit Function
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    
    ExecuteStopItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteCancelStopItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行取消撤消不良记录操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    '0-修改;1-删除;2-撤消;3-取消撤消;
    If IsAllowOperation(2, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_撤消时间).Value) = "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录未被撤档,不允许取消撤消操作!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "当前未选中要取消撤消的不良记录，不能进行取消撤消操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'bytEditType-编辑类别:0-新增;1-修改;2-撤消;3-取消撤消;4-查看
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_取消撤消, mstr行为类别, strID) Then Exit Function

    
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    ExecuteCancelStopItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteView() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行取消撤消不良记录操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    strID = rptData.SelectedRows(0).Record(COL_ID).Value
 
    If strID = "" Then
        MsgBox "当前未选中要取消撤消的不良记录，不能进行取消撤消操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'bytEditType-编辑类别:0-新增;1-修改;2-撤消;3-取消撤消;4-查看
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_查看, mstr行为类别, strID) Then Exit Function
    ExecuteView = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行删除操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-09 11:23:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strID As String, strSQL As String
    
    On Error GoTo errHandle
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    If IsAllowOperation(1, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_撤消时间).Value) <> "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录已经撤档,不允许删除!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
        If Trim(rptData.SelectedRows(0).Record(COL_是否固定).Value) <> "" Then
            strInfor = "病人“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
            strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录是系统自动生成的，你只能撤消不良记录,但不允许删除!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "当前未选中要删除的不良记录，不能进行删除操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If


    strInfor = "“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "”"
    strInfor = strInfor & "的发生时间在“" & Trim(rptData.SelectedRows(0).Record(COL_病人姓名).Value) & "的不良记录吗?"
    
    If MsgBox("你确定要删除" & strInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strSQL = "Zl_病人不良记录_Delete(" & strID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function



Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
     Err = 0: On Error GoTo errHandle
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem: Call ExecuteAddItem
    Case conMenu_Edit_Modify: Call ExecuteModifyItem
    Case conMenu_Edit_Delete: Call ExcuteDelete
    Case conMenu_Edit_Stop: Call ExecuteStopItem
    Case conMenu_Edit_Reuse: Call ExecuteCancelStopItem
    
    Case conMenu_View_ShowStoped '显示已撤消的不良记录
        mblnShowCancelRecord = Not mblnShowCancelRecord
        Control.Checked = mblnShowCancelRecord
        Call zlDatabase.SetPara("显示撤消记录", IIf(mblnShowCancelRecord, "1", "0"), glngSys, mlngModule)
        Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    Case conMenu_View_Refresh
        Call LoadRecordDataToGrid(mstr行为类别, mcllFilter)
    Case conMenu_View_Find
         If patiFind.Visible And patiFind.Enabled Then patiFind.SetFocus
    Case conMenu_View_Filter: Call ExecuteFilter
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitRptColHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化列表
    '编制:刘兴洪
    '日期:2018-11-09 11:59:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    
    Err = 0: On Error GoTo errHandle
    
    With rptData
        .AutoColumnSizing = False '不使用自动列宽
        .AllowColumnRemove = False '不允许拖动删除列
        .ShowGroupBox = True '显示分组框
        .ShowItemsInGroups = False '不显示已分组的列
        .MultipleSelection = False '不允许多行选择
        .SetImageList Me.imgList16
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid '竖向表格线格式
            .HorizontalGridStyle = xtpGridSolid '横向表格线格式
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的内容..."
            .ShadeSortColor = .BackColor
            Set .CaptionFont = Me.Font
            Set .TextFont = Me.Font
        End With
    End With

    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_图标, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        
        Set objCol = .Add(COL_行为类别, "行为类别", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_病人姓名, "病人姓名", 80, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_年龄, "年龄", 50, True): objCol.Alignment = xtpAlignmentLeft
        Set objCol = .Add(COL_性别, "性别", 50, True): objCol.Alignment = xtpAlignmentLeft
        Set objCol = .Add(COL_出生日期, "出生日期", 80, True)
        Set objCol = .Add(COL_门诊号, "门诊号", 80, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_发生时间, "发生时间", 130, True)
        
        Set objCol = .Add(COL_加入原因, "加入原因", 80, True)
        Set objCol = .Add(COL_加入详细说明, "加入详细说明", 200, True)
        
        
        Set objCol = .Add(COL_加入时间, "加入时间", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_附加信息, "附加信息", 50, True)
        Set objCol = .Add(COL_登记人, "登记人", 80, True)
        Set objCol = .Add(COL_撤消人, "撤消人", 80, True)
        Set objCol = .Add(COL_撤消时间, "撤消时间", 130, True)
        Set objCol = .Add(COL_撤消原因, "撤消原因", 100, True)
        Set objCol = .Add(COL_是否固定, "是否删除", 50, True): objCol.Visible = False
    End With
    
    With rptData
    '        '将行为类别缺省升序排列
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_发生时间)
        .SortOrder(0).SortAscending = True
        
        '将行为类别分组且升序排列
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_行为类别)
'        .GroupsOrder(0).SortAscending = True
        .Columns(COL_行为类别).Visible = False
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBlackListRecords(ByVal strType As String, ByVal cllFilter As Collection, ByRef rsBlackLists_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取黑名单记录数据
    '入参:cllFilter-条件(Array("条件",值 )
    '    strType-当前类别
    '出参:rsBlackLists_Out-返回黑名单记录数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-09 12:06:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSQL As String, varData As Variant, i As Long
    Dim dt加入时间_Start  As Date, dt加入时间_End As Date
    Dim dt撤消时间_Start  As Date, dt撤消时间_End As Date
    Dim dt发生时间_Start  As Date, dt发生时间_End As Date
    Dim str加入原因 As String, str登记人 As String, str撤消人 As String
    Dim lng病人ID As Long
    Dim dtCurdate As Date
    
    strWhere = ""
    If strType <> "" Then strWhere = " And  A.行为类别=[1]"
    If cllFilter Is Nothing Then
        Set cllFilter = New Collection
        dtCurdate = zlDatabase.Currentdate
        dt加入时间_End = Format(dtCurdate, "yyyy-mm-dd 23:59:59")
        dt加入时间_Start = Format(DateAdd("m", -6, dtCurdate), "yyyy-mm-dd 00:00:00") '缺省半年
        strWhere = strWhere & " And 加入时间 Between [3] and [4]"
    Else
     
        For i = 1 To cllFilter.Count
            varData = cllFilter(i)
            
            Select Case varData(0)
            Case "病人ID"
                lng病人ID = Val(varData(1))
                strWhere = strWhere & " And A.病人ID=[2]"
            Case "加入时间"
                dt加入时间_End = CDate(varData(2))
                dt加入时间_Start = CDate(varData(1))
                strWhere = strWhere & " And A.加入时间 Between [3] and [4]"
            Case "撤消时间"
                dt撤消时间_End = CDate(varData(2))
                dt撤消时间_Start = CDate(varData(1))
                strWhere = strWhere & " And A.撤消时间 Between [5] and [6]"
            Case "发生时间"
                dt发生时间_End = CDate(varData(2))
                dt发生时间_Start = CDate(varData(1))
                strWhere = strWhere & " And A.发生时间 Between [7] and [8]"
            Case "加入原因"
                str加入原因 = varData(1)
                strWhere = strWhere & " And A.加入原因=[9]"
            Case "登记人"
                str登记人 = varData(1)
                strWhere = strWhere & " And A.登记人=[10]"
            Case "撤消人"
                str撤消人 = varData(1)
                strWhere = strWhere & " And A.撤消人=[11]"
            End Select
        Next
    End If
    If Not mblnShowCancelRecord Then strWhere = strWhere & "  And A.撤消时间 is NULL"
    
    strSQL = "" & _
    " Select a.Id,a.行为类别, a.病人ID,b.姓名 as 病人姓名,b.性别,b.年龄,to_char(b.出生日期,'yyyy-mm-dd') as 出生日期," & _
    "       b.门诊号, to_char(a.发生时间,'yyyy-mm-dd hh24:mi:ss') as 发生时间, a.加入原因  , a.加入说明 , " & _
    "       to_char(a.加入时间,'yyyy-mm-dd hh24:mi:ss') as 加入时间  ," & vbNewLine & _
    "       a.附加信息, a.登记人 ,a.撤消原因, a.撤消人, to_char(a.撤消时间,'yyyy-mm-dd hh24:mi:ss') as 撤消时间,nvl(C.是否固定,0) as 是否固定" & vbNewLine & _
    " From 病人不良记录 A, 病人信息 B,不良行为分类 C" & vbNewLine & _
    " Where a.病人ID+0 = b.病人Id(+)  and a.行为类别=C.名称(+) " & vbNewLine & strWhere & _
    " Order by a.行为类别,a.发生时间"
    
    Set mrs不良记录 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strType, lng病人ID, dt加入时间_Start, dt加入时间_End, dt撤消时间_Start, _
        dt撤消时间_End, dt发生时间_Start, dt发生时间_End, str加入原因, str登记人, str撤消人)
        
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InsertRowData(ByVal strID As String, ByVal str行为类别 As String, ByVal str病人姓名 As String, ByVal str年龄 As String, _
    ByVal str性别 As String, ByVal str出生日期 As String, ByVal str门诊号 As String, ByVal str发生时间 As String, ByVal str加入原因 As String, _
    ByVal str加入详细说明 As String, ByVal str加入时间 As String, ByVal str附加信息 As String, _
    ByVal str登记人 As String, ByVal str撤消人 As String, str撤消时间 As String, str撤消原因 As String, ByVal bln是否固定 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向列表中插入行数据
    '编制:刘兴洪
    '日期:2018-11-09 13:43:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strTemp As String
    Dim i As Long
    
    Err = 0: On Error GoTo errHandle
    With rptData
        
        Set objRecord = .Records.Add()
        
        Set objItem = objRecord.AddItem(strID)
        Set objItem = objRecord.AddItem("")
        objItem.Icon = IIf(bln是否固定, 1, 0) '图标设置
        
        Set objItem = objRecord.AddItem(str行为类别)
        Set objItem = objRecord.AddItem(str病人姓名)
        Set objItem = objRecord.AddItem(str年龄)
        Set objItem = objRecord.AddItem(str性别)
        Set objItem = objRecord.AddItem(str出生日期)
        Set objItem = objRecord.AddItem(str门诊号)
        Set objItem = objRecord.AddItem(str发生时间)
        Set objItem = objRecord.AddItem(str加入原因)
        Set objItem = objRecord.AddItem(str加入详细说明)
        Set objItem = objRecord.AddItem(str加入时间)
        Set objItem = objRecord.AddItem(str附加信息)
        Set objItem = objRecord.AddItem(str登记人)
        Set objItem = objRecord.AddItem(str撤消人)
        Set objItem = objRecord.AddItem(str撤消时间)
        Set objItem = objRecord.AddItem(str撤消原因)
        Set objItem = objRecord.AddItem(IIf(bln是否固定, "√", ""))
 
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function LoadRecordDataToGrid(ByVal str行为类别 As String, ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给报表控件
    '入参:cllFilter-需要过滤的条件
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-09 13:40:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim i As Long, j As Long, lngSelectRow As Long, strNewItem As String
    
    Err = 0: On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).Index
    
    rptData.Records.DeleteAll
    
    If GetBlackListRecords(str行为类别, cllFilter, mrs不良记录) Then Exit Function
    With mrs不良记录
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Call InsertRowData(Nvl(!ID), Nvl(!行为类别), Nvl(!病人姓名), Nvl(!年龄), _
                 Nvl(!性别), Nvl(!出生日期), Nvl(!门诊号), Nvl(!发生时间), Nvl(!加入原因), Nvl(!加入说明), _
                Nvl(!加入时间), Nvl(!附加信息), Nvl(!登记人), _
                Nvl(!撤消人), Nvl(!撤消时间), Nvl(!撤消原因), Val(Nvl(!是否固定)) = 1)
            .MoveNext
        Loop
    End With
    With rptData
        For i = 0 To .Records.Count - 1
            If i > .Records.Count - 1 Then Exit For
            If .Records(i).Item(COL_撤消时间).Value <> "" Then
                For j = 0 To .Columns.Count - 1
                    .Records(i).Item(j).ForeColor = vbRed ' &H8000000C
                Next
            End If
        Next
    End With
    
    Call rptData.Populate '发布数据以更新界面
    If rptData.Rows.Count > 0 Then '该行选中且显示在可见区域
        If strNewItem <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_病人姓名).Caption = strNewItem Then
                        rptData.FocusedRow = rptData.Rows(i)
                        Exit For
                    End If
                End If
            Next
        Else
            If lngSelectRow = 0 Then
                rptData.FocusedRow = rptData.Rows(0)
            ElseIf lngSelectRow > rptData.Rows.Count - 1 Then
                rptData.FocusedRow = rptData.Rows(rptData.Rows.Count - 1)
            Else
                rptData.FocusedRow = rptData.Rows(lngSelectRow)
            End If
        End If
    End If
    Call SetReportControlBackColorAlternate(rptData)
    RaiseEvent zlShowStatusText(2, "当前共有" & mrs不良记录.RecordCount & "条病人不良记录")
    Screen.MousePointer = vbDefault
    Exit Function
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picBack
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
End Sub

Private Sub patiFind_FindPatiArfter(ByVal objCard As zlOneCardComLib.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlOneCardComLib.clsPatientInfo, objCardData As zlOneCardComLib.clsPatientInfo, strErrMsg As String, blnCancel As Boolean)
    Dim cllFilter As Collection, lngPatiID As Long
    If objHisPati Is Nothing Then
        If patiFind.GetCurCard.名称 Like "*姓*名*" Then
            lngPatiID = GetPatient(ShowName)
        Else
            lngPatiID = 0
        End If
    Else
        lngPatiID = objHisPati.病人ID
    End If
    
    Set cllFilter = New Collection
    cllFilter.Add Array("病人ID", lngPatiID), "病人ID"
    Call LoadRecordDataToGrid(mstr行为类别, cllFilter)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then
        rptData.SetFocus
    ElseIf Not Me.ActiveControl Is patiFind Then
        rptData.SetFocus
    End If
    RaiseEvent zlActivate(Me)
End Sub

Private Sub InitFace()

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面信息
    '编制:刘兴洪
    '日期:2018-11-09 15:32:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFind As String, objCard As zlOneCardComLib.Card, i As Long
    Dim objCards As zlOneCardComLib.Cards, strKindstr As String, dtCurdate As Date
    Dim dt加入时间_End As Date, dt加入时间_Start As Date
    On Error GoTo errHandle
    
    mblnShowCancelRecord = Val(zlDatabase.GetPara("显示撤消记录", glngSys, mlngModule, "0")) = 1
    Call InitRptColHead
    
    Set mcllFilter = New Collection
    dtCurdate = zlDatabase.Currentdate
    dt加入时间_End = Format(dtCurdate, "yyyy-mm-dd 23:59:59")
    dt加入时间_Start = Format(DateAdd("m", -6, dtCurdate), "yyyy-mm-dd 00:00:00") '缺省半年
    mcllFilter.Add Array("加入时间", dt加入时间_Start, dt加入时间_End), "加入时间"
    
    strKindstr = "姓|姓名或就诊卡|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;手|手机号|0"
    Call patiFind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindstr, gstrProductName)
    patiFind.objIDKind.AllowAutoICCard = True
    patiFind.objIDKind.AllowAutoIDCard = True
    
    Set objCards = patiFind.objIDKind.Cards
    If Not objCards Is Nothing Then
        strFind = Val(zlDatabase.GetPara("上次查找类别", glngSys, mlngModule, ""))  '查找缺省项
        If strFind <> "" Then
            For i = 1 To objCards.Count
                Set objCard = objCards(i)
                If objCard.名称 = strFind Then
                    If patiFind.GetKindIndex(objCard.接口序号) >= 0 Then
                        patiFind.IDKindIDX = i + 1
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Load()
    mlngPreSelID = -1: Call InitFace
    RestoreWinState Me, App.ProductName
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName
    If Not patiFind.GetCurCard Is Nothing Then
        Call zlDatabase.SetPara("上次查找类别", patiFind.GetCurCard.名称, glngSys, mlngModule)
    End If
    If Not mrs不良记录 Is Nothing Then Set mrs不良记录 = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '读卡
    If patiFind.Visible And patiFind.Enabled Then patiFind.ActiveFastKey
End Sub

Private Sub picBack_Resize()
    Err = 0: On Error Resume Next
    With picBack
        shpBorder.Move 0, 0, .ScaleWidth - 6, .ScaleHeight - 6
        stcTitle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        rptData.Left = .ScaleLeft + 10
        rptData.Top = stcTitle.Top + stcTitle.Height
        rptData.Width = .ScaleWidth - 30
        rptData.Height = .ScaleHeight - stcTitle.Height - 30
    End With
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandle
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    
    Me.SetFocus: RaiseEvent zlActivate(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim blnStop As Boolean, bln是否固定 As Boolean, lngID As Long
    
    Err = 0: On Error GoTo errHandle
    If rptData.SelectedRows.Count = 0 Then Exit Sub
    If rptData.SelectedRows(0).GroupRow Then Exit Sub
    
    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    blnStop = rptData.SelectedRows(0).Record(COL_撤消时间).Value <> ""
    bln是否固定 = rptData.SelectedRows(0).Record(COL_是否固定).Value <> ""
    
    If lngID = 0 Then Exit Sub
    
    If zlStr.IsHavePrivs(mstrPrivs, "编辑不良记录") And Not blnStop And Not bln是否固定 Then
        Call ExecuteModifyItem  '编辑
    Else
        Call ExecuteView '查看
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim lngID As Long
    
    Err = 0: On Error GoTo errHandle
    lngID = 0
    If rptData.SelectedRows.Count <> 0 Then
        With rptData.SelectedRows(0)
            If Not .GroupRow Then
                lngID = Val(.Record(COL_ID).Value)
            End If
        End With
    End If
    If mlngPreSelID = lngID Then Exit Sub
    mlngPreSelID = lngID
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub
Private Sub zlDataPrint(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2018-11-09 15:57:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandle
    
    If UserInfo.姓名 = "" Then Call GetUserInfo
    
    '将ReportControl转换为VSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(COL_图标) & "," & CStr(COL_是否固定)
    If zlGetVsfGrid(rptData, vsfGridPrint, strHiddenCols) = False Then Exit Sub
    
    objOut.Title.Text = "病人不良记录清单"
    Set objOut.Body = vsfGridPrint
    
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "不良行为类别：" & IIf(mstr行为类别 = "", "所有类别", mstr行为类别)
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub stcTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub
    
Private Function GetPatient(ByVal str姓名 As String) As Long
    Dim strCard As String, strPati As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If gblnShowCard Then
            strCard = "A.就诊卡号 as 就诊卡,A.就诊卡号 as 就诊卡号,"
        Else
            strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡,A.就诊卡号 as 就诊卡号,"
        End If
        
        '通过姓名模糊查找病人(允许输入病人标识时)
        strPati = _
            " Select A.病人ID ID,A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别," & _
            "   B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
            "   To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期," & _
            "   A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
            "   Nvl(a.病人类型,Decode(a.险类,Null,'普通病人','医保病人')) 病人类型" & _
            " From 病人信息 A,部门表 B,部门表 C" & _
            " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.停用时间 is NULL And A.姓名 Like [1]" & _
            " Order by A.姓名,A.登记时间 Desc"
        
        vRect = zlControl.GetControlRect(patiFind.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, patiFind.Height, blnCancel, False, True, str姓名 & "%")
        
        If rsTmp Is Nothing Then GetPatient = 0: Exit Function
        If blnCancel Then GetPatient = 0: Exit Function
        
        GetPatient = Val(Nvl(rsTmp!病人ID))

        Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

