VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEventMsgEdit 
   Caption         =   "#"
   ClientHeight    =   10185
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   16290
   Icon            =   "frmEventMsgEdit.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   16290
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4950
      Index           =   3
      Left            =   255
      ScaleHeight     =   4950
      ScaleWidth      =   8625
      TabIndex        =   4
      Top             =   990
      Width           =   8625
      Begin VB.PictureBox picPane 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   4
         Left            =   135
         ScaleHeight     =   435
         ScaleWidth      =   8310
         TabIndex        =   5
         Top             =   195
         Width           =   8310
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   3
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   75
            Width           =   570
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   0
            Left            =   1815
            Picture         =   "frmEventMsgEdit.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   75
            Width           =   300
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   1
            Left            =   6075
            Picture         =   "frmEventMsgEdit.frx":D0A4
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   75
            Width           =   315
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   75
            Width           =   975
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   4305
            TabIndex        =   7
            Top             =   75
            Width           =   1755
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   6885
            TabIndex        =   6
            Top             =   75
            Width           =   6975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "消息版本"
            Height          =   180
            Index           =   3
            Left            =   2175
            TabIndex        =   14
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "使用消息"
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "入口信息"
            Height          =   180
            Index           =   1
            Left            =   3570
            TabIndex        =   12
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明"
            Height          =   180
            Index           =   2
            Left            =   6465
            TabIndex        =   11
            Top             =   120
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3645
         Index           =   1
         Left            =   630
         TabIndex        =   17
         Top             =   945
         Width           =   5925
         _cx             =   10451
         _cy             =   6429
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
         GridColor       =   -2147483638
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
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   1
      Left            =   9300
      ScaleHeight     =   4890
      ScaleWidth      =   3180
      TabIndex        =   0
      Top             =   1155
      Width           =   3180
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   150
         ScaleHeight     =   255
         ScaleWidth      =   4095
         TabIndex        =   1
         Top             =   60
         Width           =   4095
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   -45
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   -45
            Width           =   1905
         End
      End
      Begin MSComctlLib.TreeView tvw 
         Height          =   1890
         Left            =   165
         TabIndex        =   3
         Top             =   435
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3334
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   0
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4725
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   9825
      Width           =   16290
      _ExtentX        =   28734
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEventMsgEdit.frx":138F6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23865
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEventMsgEdit.frx":1418A
      Left            =   525
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEventMsgEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private Type Items
    入口信息 As String
End Type
Private usrSaveItem As Items
Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjToolBar As Object
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mlngModualCode As Long
Private mblnContiune As Boolean
Private mrsInfoTree As ADODB.Recordset
Private mstrEventDataKey As String
Private mblnStartUp As Boolean
Private mintMaxOutlineLevel As Integer
Private mintAryOutlineLevel(0 To 11) As Integer

'Private WithEvents mfrmEventMsgEditGroup As frmEventMsgEditGroup
Private WithEvents mfrmEventMsgEditSegment As frmEventMsgEditSegment
Attribute mfrmEventMsgEditSegment.VB_VarHelpID = -1
'Private WithEvents mfrmEventMsgEditNode As frmEventMsgEditNode

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    
    
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strEventDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    mstrEventDataKey = strEventDataKey
    mstrDataKey = ""
    
    Me.Caption = "新增事件消息"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    Call InitDockPannel
    
'    Call mclsVsf.LoadDataSource(gclsBusiness.EventMsgServerRead("配置"))
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub CopyNewData(ByVal strEventDataKey As String, ByVal strReferMsgDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 4
    If strEventDataKey = "" Then Exit Sub
        
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "业务事件id", strEventDataKey)
    Call zlCommFun.SetParameter(mrsPara, "参考消息id", strReferMsgDataKey)
        
    If gclsBusiness.EventMsgEdit("Copy", mrsPara) Then
        ShowSimpleMsg "已完成复制消息操作！"
        RaiseEvent AfterNewData("")
    End If
    
End Sub

Public Sub ModifyData(ByVal strEventDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mstrEventDataKey = strEventDataKey
    mbytMode = 2
    mstrDataKey = strDataKey
    
    Me.Caption = "修改事件消息"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    Call InitDockPannel
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "事件消息id", mstrDataKey)
        
    If gclsBusiness.EventMsgEdit("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################
Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    Set tvw.ImageList = gfrmPubResource.GetImageCtl
    mblnContiune = False
        
    cbo.AddItem "1 - 固定信息"
    cbo.ItemData(cbo.NewIndex) = 1
    cbo.AddItem "2 - 业务信息"
    cbo.ItemData(cbo.NewIndex) = 2
    cbo.ListIndex = 0
    
    InitData = True
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '初始网格控件
            
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(1), True, False, GetImageList(16))
        Call .ClearColumn
        
'        Call .AppendColumn("", 120, flexAlignLeftCenter, flexDTString, , "", False, False, False)
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[序号]", False, False, False)
        
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("parent_id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("节点类型", 2100, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("节点标题", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据类型", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("重复频率", 810, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据重复_Key", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("数据重复", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据赋值_Key", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("数据赋值", 4500, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("节点说明", 1800, flexAlignLeftCenter, flexDTString, , "", True)
        
        vsf(1).OutlineCol = .ColIndex("节点类型")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("数据重复"), True, vbVsfEditCombox)
        Call .InitializeEditColumn(.ColIndex("数据赋值"), True, vbVsfEditText)
        
'        .AppendRows = True
        
    End With
    
    
    InitGrid = True
    
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "事件消息id", strDataKey)
    
    mblnReading = True
        
    Set rsTmp = gclsBusiness.EventMsgRead("事件消息", rsCondition)
    If rsTmp.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rsTmp("消息").Value)
        txt(1).Text = zlCommFun.NVL(rsTmp("入口").Value)
        cmd(1).Tag = zlCommFun.NVL(rsTmp("tab_id").Value)
        txt(2).Text = zlCommFun.NVL(rsTmp("说明").Value)
        txt(3).Text = zlCommFun.NVL(rsTmp("版本").Value)
        
        usrSaveItem.入口信息 = txt(1).Text
        
        Call GetRelationInfomation(cmd(1).Tag)
        
    End If
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "事件消息id", mstrDataKey)
    
    With mclsVsf
        Call .LoadGrid(gclsBusiness.EventMsgConfigRead("事件消息", rsCondition))
        Call vsf(1).AutoSize(.ColIndex("数据赋值"), .ColIndex("数据赋值"))
        mintMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"))
        mobjToolBar.Visible = (mintMaxOutlineLevel > 0)
        For intLoop = mintMaxOutlineLevel To 1 Step -1
            If intLoop < 12 Then mintAryOutlineLevel(intLoop) = 1
            Call mclsVsf.Outline(intLoop)
        Next
    End With
        
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "事件"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "详细"
    objPane.Options = PaneNoCaption
            
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

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
    Dim objFindKey As CommandBarControl
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    
    Set mobjToolBar = cbsMain.Add("工具栏", xtpBarTop)
    mobjToolBar.ContextMenuPresent = False
    mobjToolBar.ShowTextBelowIcons = False
    mobjToolBar.EnableDocking xtpFlagStretched
            
'    Set objPopup = zlCommFun.NewToolBar(mobjToolBar, xtpControlPopup, conMenu_Edit_NewItem, "新增")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "&1.组"): objControl.Parameter = "组"
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "&2.段"): objControl.Parameter = "段"
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "&3.节点"): objControl.Parameter = "节点"
'
'    Set objPopup = zlCommFun.NewToolBar(mobjToolBar, xtpControlPopup, conMenu_Edit_Insert, "插入")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "&1.组"): objControl.Parameter = "组"
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "&2.段"): objControl.Parameter = "段"
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "&3.节点"): objControl.Parameter = "节点"
'
'    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_Edit_Modify, "修改")
'    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True)
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "保存后继续新增", "保存后继续修改"))
    objControl.IconId = conMenu_View_UnCheck
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlLabel, 0, "层次:", True, , xtpButtonCaption)
    For intLoop = 1 To 10
        Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, " " & intLoop & " ", , , xtpButtonCaption, "选中展开当前层，不选中则收拢当前层")
        objControl.Parameter = intLoop
    Next
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, "更多", , , xtpButtonCaption, "选中展开当前层，不选中则收拢当前层")

    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Forward, "上一条", True)
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Backward, "下一条")
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_File_Exit, "退出", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub UpdateLevelCommandBar()
    Dim intLoop As Integer
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

'    Set objBar = cbsMain.Item(1)
'
'    objBar.Controls.DeleteAll
'
'    If mintMaxOutlineLevel > 0 Then
'        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 0, "层次:", , , xtpButtonCaption)
'        For intLoop = 1 To mintMaxOutlineLevel
'            Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, " " & intLoop & " ", , , xtpButtonCaption, "选中展开当前层，不选中则收拢当前层")
'            objControl.Parameter = intLoop
'        Next
'        objBar.Visible = True
'    Else
'        objBar.Visible = False
'    End If
    
End Sub

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCurrentOrder As Integer
    Dim lngLoop As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim intLoop As Integer
    Dim strKey As String
    Dim strParentKey As String
    Dim varElement As Variant
    Dim varElementKey As Variant
    Dim intCount As Integer
    Dim strName As String
    
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "必须选择一个消息！"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
'    If Trim(cmd(1).Tag) = "" Then
'        ShowSimpleMsg "必须选择一个入口信息！"
'        Call LocationObj(txt(1))
'        Exit Function
'    End If
    
    
    '同一树型路径中不能存在两个相同的“信息表.Count"配置
    With vsf(1)
        For lngLoop = 1 To .Rows - 1
            If InStr(UCase(.TextMatrix(lngLoop, .ColIndex("数据重复"))), UCase(".Count")) > 0 And .TextMatrix(lngLoop, .ColIndex("parent_id")) <> "" Then
                If CheckParentRepeat(.TextMatrix(lngLoop, .ColIndex("数据重复")), .TextMatrix(lngLoop, .ColIndex("parent_id"))) = False Then
                    Exit Function
                End If
            End If
        Next
            
'        '同一树型路径中若存在两个不同的“信息表.Count"配置，则其顺序必须和入口信息产生的信息表树一致
'        '原始的在mrsInfoTree记录集中（类型=0的，按序号排升序）
'        For lngLoop = 1 To .Rows - 1
'            If InStr(UCase(.TextMatrix(lngLoop, .ColIndex("数据重复"))), UCase(".Count")) > 0 And .TextMatrix(lngLoop, .ColIndex("parent_id")) <> "" Then
'
'                mrsInfoTree.Filter = ""
'                mrsInfoTree.Filter = "名称='" & Replace(.TextMatrix(lngLoop, .ColIndex("数据重复")), ".Count", "") & "'"
'                If mrsInfoTree.RecordCount > 0 Then
'                    intCurrentOrder = mrsInfoTree("序号").Value
'                    If CheckTableOrder(intCurrentOrder, .TextMatrix(lngLoop, .ColIndex("数据重复")), .TextMatrix(lngLoop, .ColIndex("parent_id"))) = False Then
'                        Exit Function
'                    End If
'                End If
'            End If
'        Next
'
'        mrsInfoTree.Filter = ""
        
    End With
    
    
    '根据“数据重复”和“数据赋值”产生对应的Key值，用于产生消息内容
    With vsf(1)
        
        
        .Cell(flexcpText, 1, .ColIndex("数据重复_Key"), .Rows - 1, .ColIndex("数据重复_Key")) = ""
        .Cell(flexcpText, 1, .ColIndex("数据赋值_Key"), .Rows - 1, .ColIndex("数据赋值_Key")) = ""
        
        For lngLoop = 1 To .Rows - 1
            
            '数据重复
            If InStr(UCase(.TextMatrix(lngLoop, .ColIndex("数据重复"))), UCase(".Count")) > 0 Then
                
                strTemp = .TextMatrix(lngLoop, .ColIndex("数据重复"))
                strTemp = Replace(strTemp, "[S.", "[" & txt(1).Text & ".")
                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                strTemp = Replace(strTemp, ".Count", "")
                varTemp = Split(strTemp, ".")
                strParentKey = ""
                strKey = ""
                For intLoop = 0 To UBound(varTemp)
                    mrsInfoTree.Filter = ""
                    
'                    If intLoop = UBound(varTemp) Then
'                        mrsInfoTree.Filter = "上级id='" & strParentKey & "' And 名称='" & varTemp(intLoop) & "' And 类型=1"
'                    Else
                    mrsInfoTree.Filter = "上级id='" & strParentKey & "' And 名称='" & varTemp(intLoop) & "' And 类型=0"
'                    End If
                                        
                    If mrsInfoTree.RecordCount > 0 Then
                        strKey = mrsInfoTree("id").Value
                        strParentKey = mrsInfoTree("id").Value
                    Else
                        ShowSimpleMsg "信息配置有误！"
                        .Row = lngLoop
                        .Col = .ColIndex("数据重复")
                        .ShowCell lngLoop, .ColIndex("数据重复")
                        .SetFocus
                        Exit Function
                    End If
                Next
                
                If InStr(strKey, "R_") > 0 Then
                    strKey = Mid(strKey, InStr(strKey, "R_") + 2, 32)
                Else
                    If InStr(strKey, "T_") > 0 Then
                        strKey = Mid(strKey, InStr(strKey, "T_") + 2, 32)
                    Else
                        strKey = ""
                    End If
                End If
                If strKey <> "" Then
                    .TextMatrix(lngLoop, .ColIndex("数据重复_Key")) = "[" & strKey & ".Count]"
                End If
                
            End If
            
            
            '数据赋值
            If .TextMatrix(lngLoop, .ColIndex("数据赋值")) <> "" Then
                
                strTemp = .TextMatrix(lngLoop, .ColIndex("数据赋值"))
                
'                If strTemp = "[S.开始时间]^[S.病案主页.入院日期]" Then
'                    strTemp = "[S.开始时间]^[S.病案主页.入院日期]"
'                End If
                
                strTemp = Replace(strTemp, "[S.", "[" & txt(1).Text & ".")
                
                varElement = GetElement(strTemp)
                varElementKey = varElement
                If IsEmpty(varElement) = False Then
                    For intCount = 0 To UBound(varElement)
                        
                        varTemp = Split(varElement(intCount), ".")
                        strParentKey = ""
                        strKey = ""
                        strName = ""
                        For intLoop = 0 To UBound(varTemp) - 1
                            mrsInfoTree.Filter = ""
                            If intLoop = UBound(varTemp) Then
                                mrsInfoTree.Filter = "上级id='" & strParentKey & "' And 名称='" & varTemp(intLoop) & "' And 类型=1"
                            Else
                                mrsInfoTree.Filter = "上级id='" & strParentKey & "' And 名称='" & varTemp(intLoop) & "' And 类型=0"
                            End If
                    
                            If mrsInfoTree.RecordCount > 0 Then
                                strKey = mrsInfoTree("id").Value
                                strParentKey = mrsInfoTree("id").Value
                                strName = strName & "." & varTemp(intLoop)
                            Else
                                ShowSimpleMsg "信息配置有误！"
                                .Row = lngLoop
                                .Col = .ColIndex("数据赋值")
                                .ShowCell lngLoop, .ColIndex("数据赋值")
                                .SetFocus
                                Exit Function
                            End If
                        Next
                        If strName <> "" Then strName = Mid(strName, 2)
                        
                        If InStr(strKey, "R_") > 0 Then
                            strKey = Mid(strKey, InStr(strKey, "R_") + 2, 32)
                        Else
                            If InStr(strKey, "T_") > 0 Then
                                strKey = Mid(strKey, InStr(strKey, "T_") + 2, 32)
                            Else
                                strKey = ""
                            End If
                        End If
                        
                        If strKey <> "" Then varElementKey(intCount) = Replace(varElementKey(intCount), strName & ".", strKey & ".")
'                        If strKey <> "" Then strTemp = Replace(strTemp, strName & ".", strKey & ".")
                        
                    Next
                    
                    For intCount = 0 To UBound(varElement)
                        strTemp = Replace(strTemp, varElement(intCount), varElementKey(intCount))
                    Next
                    
                End If
                
                .TextMatrix(lngLoop, .ColIndex("数据赋值_Key")) = strTemp
                
            End If
            
        Next
        If Not (mrsInfoTree Is Nothing) Then mrsInfoTree.Filter = ""
    End With
    
    
    ValidData = True
    
End Function

Private Function GetElement(ByVal strExpress As String) As Variant
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim lngBeginVar As Long
    Dim lngEndVar As Long
    Dim strVar As String
    Dim strTemp As String
    Dim strChar As String
    
    lngCount = Len(strExpress)
    For lngLoop = 1 To lngCount
        strChar = Mid(strExpress, lngLoop, 1)
        Select Case strChar
        Case "["
            lngBeginVar = lngLoop
        Case "]"
            If lngBeginVar > 0 Then
                lngEndVar = lngLoop
                strTemp = Mid(strExpress, lngBeginVar + 1, lngEndVar - lngBeginVar - 1)
                
                If InStr("'" & strVar & "'", "'" & strTemp & "'") = 0 And InStr(strTemp, ".") > 0 Then
                    strVar = strVar & "'" & strTemp
                End If
                                
                lngBeginVar = 0
                lngEndVar = 0
            End If
        End Select
    Next
    If strVar <> "" Then strVar = Mid(strVar, 2)
    If strVar <> "" Then GetElement = Split(strVar, "'")
End Function

Private Function CheckTableOrder(ByVal intChildOrder As Integer, ByVal strCurrentConfig As String, ByVal strParentKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strParentConfig As String
    Dim intColIndex As Integer
    Dim lngRow As Long
    Dim intCurrentOrder As Integer
    
    With vsf(1)
        lngRow = .FindRow(strParentKey)
        
        intColIndex = .ColIndex("数据重复")
        
        If InStr(UCase(.TextMatrix(lngRow, intColIndex)), UCase(".Count")) > 0 Then
            strParentConfig = .TextMatrix(lngRow, intColIndex)
            
            mrsInfoTree.Filter = ""
            mrsInfoTree.Filter = "名称='" & Replace(strParentConfig, ".Count", "") & "'"
            If mrsInfoTree.RecordCount > 0 Then
                intCurrentOrder = mrsInfoTree("序号").Value
                If intCurrentOrder >= intChildOrder Then
                    ShowSimpleMsg "同一树型分支路径中如果存在两个及以上不同的“" & strCurrentConfig & "”配置，则其顺序必须和入口信息产生的信息表树一致！"
                    .ShowCell lngRow, intColIndex
                    .Row = lngRow
                    .Col = intColIndex
                    .SetFocus
                    Exit Function
                End If
            End If
        End If
        strParentKey = .TextMatrix(lngRow, .ColIndex("parent_id"))
                
        If CheckTableOrder(intCurrentOrder, strCurrentConfig, strParentKey) = False Then
            Exit Function
        End If
        
    End With
                    
    CheckTableOrder = True
    
End Function

Private Function CheckParentRepeat(ByVal strCurrentConfig As String, ByVal strParentKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strParentConfig As String
    Dim intColIndex As Integer
    Dim lngRow As Long
    
    With vsf(1)
'        lngRow = .FindRow(strParentKey)
        lngRow = mclsVsf.FindRow(strParentKey, .ColIndex("ID"), 1)
        
        intColIndex = .ColIndex("数据重复")
        
        If InStr(UCase(.TextMatrix(lngRow, intColIndex)), UCase(".Count")) > 0 Then
            strParentConfig = .TextMatrix(lngRow, intColIndex)
            If strCurrentConfig = strParentConfig Then
                '提示
                ShowSimpleMsg "同一树型分支路径中不能存在两个相同的“" & strCurrentConfig & "”配置！"
                .ShowCell lngRow, intColIndex
                .Row = lngRow
                .Col = intColIndex
                .SetFocus
                Exit Function
            End If
        End If
        
        strParentKey = .TextMatrix(lngRow, .ColIndex("parent_id"))
        If strParentKey <> "" Then
            If CheckParentRepeat(strCurrentConfig, strParentKey) = False Then
                Exit Function
            End If
        End If
        
    End With
                    
    CheckParentRepeat = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim strTemp As String
    Dim strLine As String
    Dim lngLoop As Long
    Dim aryTemp As Variant
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "消息标识", Trim(txt(0).Text))
    Call zlCommFun.SetParameter(rsPara, "消息版本", Trim(txt(3).Text))
    Call zlCommFun.SetParameter(rsPara, "业务事件id", mstrEventDataKey)
    Call zlCommFun.SetParameter(rsPara, "入口信息", cmd(1).Tag)
    Call zlCommFun.SetParameter(rsPara, "失败处理", 0)
    Call zlCommFun.SetParameter(rsPara, "说明", Trim(txt(2).Text))
        
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    lngCount = 0
    With vsf(1)
        For lngLoop = 1 To .Rows - 1
            
            strLine = ""
            strLine = Trim(.TextMatrix(lngLoop, .ColIndex("id")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("parent_id")))
            strLine = strLine & "," & lngLoop
            
            '1-Segment;2-Data;3-Composite;4-Group
            Select Case Trim(.TextMatrix(lngLoop, .ColIndex("节点类型")))
            Case "Segment"
                strLine = strLine & ",1"
            Case "Data"
                strLine = strLine & ",2"
            Case "Composite"
                strLine = strLine & ",3"
            Case "Group"
                strLine = strLine & ",4"
            End Select
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("节点标题")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("数据类型")))
            
            If .TextMatrix(lngLoop, .ColIndex("重复频率")) <> "" Then
                aryTemp = Split(.TextMatrix(lngLoop, .ColIndex("重复频率")), "～")
                strLine = strLine & "," & Trim(aryTemp(0))
                strLine = strLine & "," & Trim(aryTemp(1))
            Else
                strLine = strLine & ",0"
                strLine = strLine & ",0"
            End If
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("数据重复")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("数据赋值")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("数据重复_Key")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("数据赋值_Key")))
            strLine = strLine & "," & Replace(Trim(.TextMatrix(lngLoop, .ColIndex("节点说明"))), "'", "''")
            
            If LenB(strTemp & ";" & strLine) > 3500 Then
                If strTemp <> "" Then
                    lngCount = lngCount + 1
                    strTemp = Mid(strTemp, 2)
                    Call zlCommFun.SetParameter(rsPara, "消息配置_" & lngCount, strTemp)
                    strTemp = ""
                End If
            End If
            strTemp = strTemp & ";" & strLine
        Next
    End With
    
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "消息配置_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "消息配置段数", lngCount)
    
'    '------------------------------------------------------------------------------------------------------------------
'    strTemp = ""
'    With vsf(0)
'        For lngLoop = 1 To .Rows - 1
'            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 Then
'                strTemp = strTemp & ";" & Trim(.TextMatrix(lngLoop, .ColIndex("id")))
'            End If
'        Next
'    End With
'    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
'    Call zlCommFun.SetParameter(rsPara, "投递目标", strTemp)
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.EventMsgEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsBusiness.EventMsgEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click()
    Dim objNode As Node
    
    tvw.Nodes.Clear
    
    Select Case cbo.ItemData(cbo.ListIndex)
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        
        tvw.Nodes.Add , , "K6", "字段符号", "file", "file"
        tvw.Nodes.Add , , "K7", "分隔符号", "file", "file"
        tvw.Nodes.Add , , "K1", "发送程序", "file", "file"
        tvw.Nodes.Add , , "K2", "发送设备", "file", "file"
        tvw.Nodes.Add , , "K3", "接收程序", "file", "file"
        tvw.Nodes.Add , , "K4", "接收设备", "file", "file"
        tvw.Nodes.Add , , "K5", "当前时间", "file", "file"
        tvw.Nodes.Add , , "K8", "消息类型", "file", "file"
        tvw.Nodes.Add , , "K9", "事件类型", "file", "file"
        tvw.Nodes.Add , , "K10", "消息结构", "file", "file"
        tvw.Nodes.Add , , "K12", "消息标识", "file", "file"
        tvw.Nodes.Add , , "K13", "事件时间", "file", "file"
        tvw.Nodes.Add , , "K11", "消息版本", "file", "file"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        If Not (mrsInfoTree Is Nothing) Then
            
            mrsInfoTree.Filter = ""
            If mrsInfoTree.RecordCount > 0 Then
                mrsInfoTree.MoveFirst
                Do While Not mrsInfoTree.EOF
                    
                    If zlCommFun.NVL(mrsInfoTree("上级id").Value) = "" Then
                        Set objNode = tvw.Nodes.Add(, , "K_" & mrsInfoTree("id").Value, mrsInfoTree("名称").Value)
                        objNode.Expanded = True
                    Else
                        Set objNode = tvw.Nodes.Add("K_" & mrsInfoTree("上级id").Value, tvwChild, "K_" & mrsInfoTree("id").Value, mrsInfoTree("名称").Value)
                        objNode.Expanded = False
                    End If
                    
                    'constitute
                    
                    If Val(mrsInfoTree("关系").Value) = 2 Then
                        objNode.Image = "constitute"
                    Else
                        objNode.Image = IIf(Val(mrsInfoTree("类型").Value) = 0, "folder", "file")
                    End If
                    
                    mrsInfoTree.MoveNext
                Loop
            End If
            
        End If
        
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim intIndex As Integer
    Dim intMaxIndex As Integer
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_NewItem
'
'        Select Case Control.Parameter
'        Case "组"
'
'        Case "段"
'            If mfrmEventMsgEditSegment Is Nothing Then
'                Set mfrmEventMsgEditSegment = New frmEventMsgEditSegment
'                Call mfrmEventMsgEditSegment.InitDialog(Me)
'            End If
'            Call mfrmEventMsgEditSegment.NewData
'        End Select
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Modify
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Delete
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
        Call Save
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit
        Unload Me
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '上一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '下一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    '--------------------------------------------------------------------------------------------------------------
    Case 1
        intIndex = Val(Control.Parameter)
        If mintAryOutlineLevel(intIndex) = 1 Then
            '展开,前面的自动展开
            
            With vsf(1)
                If intIndex > 10 Then
                    intMaxIndex = mintMaxOutlineLevel
                Else
                    intMaxIndex = intIndex
                End If
                
                For lngRow = 1 To .Rows - 1
                    If .IsSubtotal(lngRow) = True And .RowOutlineLevel(lngRow) <= intMaxIndex Then
                        .IsCollapsed(lngRow) = flexOutlineExpanded
                    End If
                Next
            End With
            
            For intLoop = intIndex To 1 Step -1
                mintAryOutlineLevel(intLoop) = -1
            Next
            
        Else
            If intIndex > 10 Then
                For intLoop = 11 To mintMaxOutlineLevel
                    mclsVsf.Outline intLoop
                Next
            Else
                mclsVsf.Outline intIndex
            End If
            
            mintAryOutlineLevel(intIndex) = 1
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append
        Call Fill(True)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        Call Fill(False)
    End Select
    
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '窗体其它控件Resize处理
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward
        Control.Visible = (mbytMode = 2)
        Control.Enabled = (Control.Visible And mblnDataChanged = False)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    '--------------------------------------------------------------------------------------------------------------
    Case 1
        
        Control.Checked = (mintAryOutlineLevel(Val(Control.Parameter)) = -1)
        Control.Visible = (Val(Control.Parameter) > 0 And Val(Control.Parameter) <= mintMaxOutlineLevel)
        
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim strFile As String
    Dim objclsHL7 As clsHL7V2EDI
    Dim rsFormat As ADODB.Recordset
    Dim strMessageType As String
    Dim strMessageVer As String
    Dim rsData As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim strTemp As String
    Dim intLoop As Integer
    
    Select Case Index
'    '------------------------------------------------------------------------------------------------------------------
'    Case 0
'
'        'cmdlg
'        strFile = OpenHLDialog
'        If strFile <> "" Then
'
'            Set objclsHL7 = New clsHL7V2EDI
'
'            If objclsHL7.GetMessageFormat(strFile, strMessageType, strMessageVer, rsFormat) Then
'
'                txt(0).Text = strMessageType
'                txt(3).Text = strMessageVer
'
'                If Not (rsFormat Is Nothing) Then
'                    With mclsVsf
'                        Call .LoadGrid(rsFormat)
'                        mintMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"))
'                        mobjToolBar.Visible = (mintMaxOutlineLevel > 0)
'                        For intLoop = mintMaxOutlineLevel To 1 Step -1
'                            If intLoop < 12 Then mintAryOutlineLevel(intLoop) = 1
'                            Call mclsVsf.Outline(intLoop)
'                        Next
'                    End With
'                End If
'
'            End If
'
'            Set objclsHL7 = Nothing
'        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        
        
        Set rsData = gclsBusiness.TableRead("SelectData")
        
        If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "编码,900,0,1;名称,2400,0,0;说明,1200,0,0", mfrmParent.Name & "\入口信息选择", "请从下表中选择一个入口信息", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
            
            If Trim(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                
                With vsf(1)
                    .Cell(flexcpText, 1, .ColIndex("数据重复"), .Rows - 1, .ColIndex("数据重复")) = ""
                    .Cell(flexcpText, 1, .ColIndex("数据重复"), .Rows - 1, .ColIndex("数据赋值")) = ""
                End With
                
                txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                txt(Index).Tag = ""
                usrSaveItem.入口信息 = txt(Index).Text
                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                
                mblnDataChanged = True
                
                Call GetRelationInfomation(zlCommFun.NVL(rs("ID").Value))
                
            End If
'            Call ReEnabled
            Call LocationObj(txt(Index), True)
        Else
            txt(Index).Text = usrSaveItem.入口信息
            txt(Index).Tag = ""
'            Call ReEnabled
            Call LocationObj(txt(Index), True)
            Exit Sub
        End If
        
    End Select
    
End Sub

Private Sub GetRelationInfomation(ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    Dim varTemp As Variant
    Dim intLoop As Integer
    Dim strParentKey As String
    Dim strElement As String
    
'    strTemp = "0|1"
    Set mrsInfoTree = gclsBusiness.GetTableTree(strDataKey)
    mrsInfoTree.Filter = "关系=2"
    If mrsInfoTree.RecordCount > 0 Then
        mrsInfoTree.MoveFirst
        Do While Not mrsInfoTree.EOF
            strTemp = strTemp & "|" & mrsInfoTree("名称").Value & "." & mrsInfoTree("上级id").Value
            mrsInfoTree.MoveNext
        Loop
    End If
    
    
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '
    varTemp = Split(strTemp, "|")
    
    strTemp = ""
    For intLoop = 0 To UBound(varTemp)
        
        strElement = ""
        strParentKey = Mid(varTemp(intLoop), InStr(varTemp(intLoop), ".") + 1)
        strElement = Mid(varTemp(intLoop), 1, InStr(varTemp(intLoop), ".") - 1)
        If strParentKey = "" Then
            strElement = "[" & Mid(varTemp(intLoop), 1, InStr(varTemp(intLoop), ".") - 1) & ".Count]"
        Else
            Do While strParentKey <> ""
                mrsInfoTree.Filter = ""
                mrsInfoTree.Filter = "id='" & strParentKey & "'"
                If mrsInfoTree.RecordCount > 0 Then
                    '
                    strElement = mrsInfoTree("名称").Value & "." & strElement
                    strParentKey = mrsInfoTree("上级id").Value
                Else
                    strParentKey = ""
                End If
            Loop
'            If strElement <> "" Then strElement = Mid(strElement, 1, Len(strElement) - 1)
            If strElement <> "" Then strElement = "[" & strElement & ".Count]"
        End If
        
        strTemp = strTemp & "|" & strElement
    Next
    
    strTemp = Replace(strTemp, "[" & txt(1).Text & ".", "[S.")
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then
        strTemp = "0|1|" & strTemp
    Else
         strTemp = "0|1"
    End If
    
    mclsVsf.DropColData(mclsVsf.ColIndex("数据重复")) = strTemp
                
End Sub

Public Function OpenHLDialog() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim strTmp As String
            
    With cmdlg
        .DialogTitle = "请选择HL7消息标准"
        .Filter = "消息标准(*.config)|*.config"
    
        On Error Resume Next
    
        .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenHLDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    ShowSimpleMsg "不能打开文件(" & strTmp & "),该文件可能正在使用或文件不存在!"
End Function

'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub

Private Sub Save()
    Dim strOldDataKey As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnDataChanged = True Then
        If ValidData = True Then
    
            If SaveData(mstrDataKey) = True Then
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(0).Text = ""
                        txt(1).Text = ""
                        txt(2).Text = ""
                        txt(3).Text = ""
                        cmd(1).Tag = ""
                        mclsVsf.ClearGrid
                        Call LocationObj(txt(0))
                    Else
                        vsf(1).SetFocus
                    End If
                    
                    mblnDataChanged = False
                End If
                
            End If
        End If
    Else
        If mblnContiune = False Then Unload Me
    End If
    
End Sub

Private Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    
    Select Case bytPlace
    Case 1  '
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Append, "追加(&A)")
        cbrPopupItem.DefaultItem = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "替换(&U)")
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(3).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call zlComLib.RestoreWinState(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 2, 200, 15, 200, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    
    Call zlComLib.SaveWinState(Me)
    
    
    If Not (mclsVsf Is Nothing) Then Set mclsVsf = Nothing
    If Not (mrsInfoTree Is Nothing) Then Set mrsInfoTree = Nothing
'    If Not (mfrmEventMsgEditGroup Is Nothing) Then
'        Unload mfrmEventMsgEditGroup
'        Set mfrmEventMsgEditGroup = Nothing
'    End If
    
    If Not (mfrmEventMsgEditSegment Is Nothing) Then
        Unload mfrmEventMsgEditSegment
        Set mfrmEventMsgEditSegment = Nothing
    End If
'
'    If Not (mfrmEventMsgEditNode Is Nothing) Then
'        Unload mfrmEventMsgEditNode
'        Set mfrmEventMsgEditNode = Nothing
'    End If
    
End Sub

Private Sub mclsVsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    With vsf(1)
        .TextMatrix(Row, .ColIndex("数据重复")) = ""
        .TextMatrix(Row, .ColIndex("数据赋值")) = ""
        mblnDataChanged = True
    End With
End Sub

Private Sub mfrmEventMsgEditSegment_AfterNewData(ByVal DataKey As String)
    '
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        picPane(2).Move 15, 15, picPane(Index).Width - 30
        tvw.Move 15, picPane(2).Top + picPane(2).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picPane(2).Top + picPane(2).Height + 15) - 15
    Case 2
        cbo.Move -30, -30, picPane(Index).Width + 60
    Case 3
        picPane(4).Move 15, 15, picPane(Index).Width - 30
        
        
        vsf(1).Move 15, picPane(4).Top + picPane(4).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picPane(4).Top + picPane(4).Height + 15) - 15
        
    Case 4
        txt(2).Move txt(2).Left, txt(2).Top, picPane(Index).Width - txt(2).Left - 60
        
    End Select
    
End Sub

Private Sub tvw_DblClick()
    
    Call Fill(True)
    
End Sub

Private Sub Fill(Optional ByVal blnAppend As Boolean = True)
    
    Dim strSelectText As String
    Dim strCellText As String
    Dim objNode As Node
        
    If tvw.SelectedItem.Child Is Nothing Then
        
        If mclsVsf.AllowColEdit(mclsVsf.ColIndex("数据赋值")) = False Then GoTo EndPoint
        
        '先取出选择的内容
        Set objNode = tvw.SelectedItem
        Select Case cbo.ItemData(cbo.ListIndex)
        Case 1
            strSelectText = "[" & objNode.Text & "]"
        Case 2
            If Not (objNode.Parent Is Nothing) Then
                strSelectText = objNode.Text
                Do While Not (objNode.Parent Is Nothing)
                    strSelectText = objNode.Parent.Text & "." & strSelectText
                    Set objNode = objNode.Parent
                Loop
                If strSelectText <> "" Then strSelectText = "[" & strSelectText & "]"
            End If
            strSelectText = Replace(strSelectText, "[" & txt(1).Text & ".", "[S.")
        End Select
        If strSelectText = "" Then GoTo EndPoint
                
        With vsf(1)
            
                                
            If blnAppend Then
                strCellText = .TextMatrix(.Row, .ColIndex("数据赋值"))
                If Trim(strCellText) = "" Then
                    strCellText = strSelectText
                Else
                    strCellText = strCellText & "^" & strSelectText
                End If
                
            Else
                strCellText = strSelectText
            End If
'
'            If Trim(strCellText) <> "" Then
'                If cbo.ItemData(cbo.ListIndex) = 1 Then
'                    strCellText = strSelectText
'                Else
'                    strCellText = strCellText & "^" & strSelectText
'                End If
'            Else
'                strCellText = strSelectText
'            End If
            
            .TextMatrix(.Row, .ColIndex("数据赋值")) = strCellText
            
            .AutoSize .ColIndex("数据赋值"), .ColIndex("数据赋值")
            If .TextMatrix(.Row, .ColIndex("数据赋值")) <> "" And (.TextMatrix(.Row, .ColIndex("数据重复")) = "" Or .TextMatrix(.Row, .ColIndex("数据重复")) = "0") Then
                .TextMatrix(.Row, .ColIndex("数据重复")) = "1"
            End If
            mblnDataChanged = True
            
        End With
    End If
    
EndPoint:
    vsf(1).Col = vsf(1).ColIndex("数据赋值")
    vsf(1).ShowCell vsf(1).Row, vsf(1).Col
    vsf(1).SetFocus
    
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call tvw_DblClick
    End If
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call zlCommFun.SendLMouseButton(tvw.hWnd, X, Y)

        Call ShowConetneMenu(1).ShowPopup

        
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
    
    Select Case Index
    Case 1
    
        txt(Index).Tag = "Changed"
            
    End Select
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            cmd(Index).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.入口信息 = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
'            If cmd(index).Visible And cmd(index).Enabled Then Call cmd_Click(index)
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""

                Set rsCondition = zlCommFun.CreateCondition
                Call zlCommFun.SetCondition(rsCondition, "FilterText", txt(Index).Text)
                
                Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
                If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "编码,900,0,1;名称,2400,0,0;说明,1200,0,0", mfrmParent.Name & "\入口信息过滤", "请从下表中选择一个入口信", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
                    
                    If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                        txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                        txt(Index).Tag = ""
                        usrSaveItem.入口信息 = txt(Index).Text
                        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                        mblnDataChanged = True
                    End If
                Else
                    txt(Index).Text = usrSaveItem.入口信息
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            End If
        End Select
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Cancel Then Exit Sub

    Select Case Index
    Case 1
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.入口信息
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    mblnDataChanged = True
    
    With vsf(Index)
        Select Case Col
        Case .ColIndex("数据赋值")
            If .TextMatrix(Row, .ColIndex("数据赋值")) <> "" And (.TextMatrix(Row, .ColIndex("数据重复")) = "" Or .TextMatrix(Row, .ColIndex("数据重复")) = "0") Then
                .TextMatrix(Row, .ColIndex("数据重复")) = "1"
            End If
            .AutoSize .ColIndex("数据赋值"), .ColIndex("数据赋值")
        End Select
        
    End With
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnStartUp Then Exit Sub

    With vsf(Index)
        
        Select Case UCase(.TextMatrix(NewRow, .ColIndex("节点类型")))
        Case UCase("Segment"), UCase("Group")
            mclsVsf.AllowColEdit(mclsVsf.ColIndex("数据赋值")) = False
        Case Else
            mclsVsf.AllowColEdit(mclsVsf.ColIndex("数据赋值")) = True
        End Select

    End With
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    Dim rsData As ADODB.Recordset
'    Dim strDataInfo As String
'
'    With vsf(Index)
'        Select Case Col
'        Case .ColIndex("数据赋值")
'
'            strDataInfo = ShowPubSelect(Me, vsf(Index), mrsInfoTree, .TextMatrix(Row, Col), 7500, 6000)
'            .Cell(flexcpData, Row, Col, Row, Col) = strDataInfo
'            .TextMatrix(Row, Col) = strDataInfo
'            mblnDataChanged = True
'        End Select
'    End With
    
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim lngRow As Long
    
    Call mclsVsf.DbClick
    
    With vsf(Index)
        Select Case .Col
        Case .ColIndex("数据重复"), .ColIndex("数据赋值")
        
        Case Else
            
            lngRow = .Row
            
            If .IsSubtotal(lngRow) = True Then
                .IsCollapsed(lngRow) = IIf(.IsCollapsed(lngRow) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
            End If
            
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngRow As Long
    
    With vsf(Index)
        If KeyAscii = vbKeySpace Then
            Select Case .Col
            Case .ColIndex("数据重复"), .ColIndex("数据赋值")
            
            Case Else
                Call vsf_DblClick(Index)
            End Select
        End If
    End With
    
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub
