VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppforBillDesign 
   Caption         =   "申请单设计"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmAppforBillDesign.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   15105
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboDeptSel 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   870
      Width           =   2025
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6345
      Left            =   6480
      ScaleHeight     =   6345
      ScaleWidth      =   3225
      TabIndex        =   1
      Top             =   450
      Width           =   3225
      Begin VSFlex8Ctl.VSFlexGrid VSFListDept 
         Height          =   2835
         Left            =   60
         TabIndex        =   3
         Top             =   3150
         Width           =   2895
         _cx             =   5106
         _cy             =   5001
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFList 
         Height          =   1995
         Left            =   150
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1111111"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   930
            TabIndex        =   10
            Top             =   1500
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionDept 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "申请单执行小组"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionItem 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "申请单项目(点击""调整顺序""之后可拖动改变顺序)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   60
      ScaleHeight     =   5235
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   840
      Width           =   3225
      Begin VSFlex8Ctl.VSFlexGrid VSFType 
         Height          =   1995
         Left            =   30
         TabIndex        =   9
         Top             =   510
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
      Begin XtremeSuiteControls.ShortcutCaption ShortCaptionType 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2745
         _Version        =   589884
         _ExtentX        =   4842
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "申请单"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   14737632
         GradientColorDark=   14737632
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8700
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23945
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   390
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppforBillDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmIfShow As Boolean                                        '窗体是否显示完成
Private mlngkeyID As Long                                               '分类ID
Private mblnAllSite As Boolean                                          '查看所有站点
Private mblnItemSort As Boolean                                         '是否处于顺序调整状态

'实现拖动效果需要的变量
Private mlngMouseRow As Long            '新增的行
Private mlngMouseDownRow As Long        '鼠标按下的行

Public Sub ShowMe(objFrm As Object, ByVal blnAddSite As Boolean)
    mblnAllSite = blnAddSite
    Me.Show 1, objFrm
End Sub

Private Sub cboDeptSel_Click()
    Call ReadTypeData
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill                     '增加申请单
            frmAppforBillDesignEditBill.ShowMe Me, 0, Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex), "", "", 0, False
            Call ReadTypeData
        Case ConMenu_Appfro_ModifyBill                  '修改分类
            Call ModifyType
        Case ConMenu_Appfro_DelBill                     '删除申请单
            Call DelType
        Case ConMenu_Appfro_ModifyItem                  '修改申请项目
            Call ModifyItem
        Case ConMenu_Appfro_Group                       '选择分组
            Call SelectGroup
        Case ConMenu_Appfro_ModifyDept                  '执行小组
            Call ModifyGroup
        Case ConMenu_Appfor_ItemSort                    '调整顺序
            If Control.Caption = "保存" Then
                Call SaveItemSort
                Control.Caption = "调整顺序"
                cbrthis.RecalcLayout
            Else
                mblnItemSort = True
                Control.Caption = "保存"
                cbrthis.RecalcLayout
            End If
        Case ConMenu_Appfro_Refresh                     '刷新
            Call ReadTypeData
        Case ConMenu_Appfro_Exit                        '退出
            Unload Me
    End Select
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    With Me.picLeft
        .Top = Top
        .Left = Left + 10
        .Width = (Right - Left) / 3
        .Height = Bottom - Top - stbThis.Height + 25
    End With
    With Me.picRight
        .Top = Top + 6
        .Left = picLeft.Left + picLeft.Width + 25
        .Width = (Right - Left) - .Left - 25
        .Height = Me.picLeft.Height
    End With
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill                     '增加申请单
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyBill                  '修改分类
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_DelBill                     '删除申请单
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyItem                  '修改申请项目
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_Group                       '选择分组
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_ModifyDept                  '执行小组
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Appfro_Refresh                     '刷新
            Control.Enabled = Not mblnItemSort
    End Select
    picLeft.Enabled = Not mblnItemSort
End Sub

Private Sub Form_Activate()
    If mblnfrmIfShow = False Then
        Call InitVSF
        Call LoadDept
        
        mblnfrmIfShow = True
    End If
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_AddBill, "增加")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyBill, "修改")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_DelBill, "删除")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyItem, "修改项目")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Group, "选择分组")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyDept, "执行小组")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ItemSort, "调整顺序")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "退出")
        
        
    End With
    
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlLabel, 0, "应用科室")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = cbrToolBar.Controls.Add(xtpControlCustom, ConMenu_Appfro_DeptSel, "应用科室")
    cbrCustom.ShortcutText = "应用科室"
    cbrCustom.Handle = Me.cboDeptSel.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmIfShow = False
    mblnAllSite = False
    mblnItemSort = False
    mlngMouseRow = 0
    mlngMouseDownRow = 0
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With ShortCaptionType
        .Top = 6
        .Left = 6
        .Width = Me.picLeft.ScaleWidth
    End With
    With VSFType
        .Top = ShortCaptionType.Height + 12
        .Left = 6
        .Width = picLeft.ScaleWidth - .Left * 2
        .Height = picLeft.ScaleHeight - .Top
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With ShortCaptionItem
        .Top = 6
        .Left = 6
        .Width = Me.picRight.ScaleWidth
    End With
    With VSFList
        .Top = ShortCaptionItem.Top + ShortCaptionItem.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth - .Left * 2
        .Height = picRight.ScaleHeight - VSFListDept.Height - ShortCaptionDept.Height - 48
    End With
    With ShortCaptionDept
        .Top = VSFList.Top + VSFList.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth
    End With
    With VSFListDept
        .Top = ShortCaptionDept.Top + ShortCaptionDept.Height + 6
        .Left = 6
        .Width = picRight.ScaleWidth
        .Height = Me.picRight.ScaleHeight - .Top
    End With
    
End Sub
Private Sub InitVSF()
      '初始化列表
1         On Error GoTo InitVSF_Error

2         With Me.VSFList
3             .Rows = 2
4             .Cols = 6
5             .FixedRows = 1
6             .ColKey(0) = "编码": .ColWidth(.ColIndex("编码")) = 2000: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
7             .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
8             .ColKey(1) = "组合项目": .ColWidth(.ColIndex("组合项目")) = 3000: .ColAlignment(.ColIndex("组合项目")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("组合项目")) = "组合项目"
9             .Cell(flexcpAlignment, 0, .ColIndex("组合项目"), 0, .ColIndex("组合项目")) = flexAlignCenterCenter
10            .ColKey(2) = "分类": .ColWidth(.ColIndex("分类")) = 2000: .ColAlignment(.ColIndex("分类")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("分类")) = "分类"
11            .Cell(flexcpAlignment, 0, .ColIndex("分类"), 0, .ColIndex("分类")) = flexAlignCenterCenter
12            .ColKey(3) = "排列顺序": .ColAlignment(.ColIndex("排列顺序")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("排列顺序")) = "排列顺序"
13            .Cell(flexcpAlignment, 0, .ColIndex("排列顺序"), 0, .ColIndex("排列顺序")) = flexAlignCenterCenter: .ColHidden(.ColIndex("排列顺序")) = True
14            .ColKey(4) = "ID": .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID"
15            .Cell(flexcpAlignment, 0, .ColIndex("ID"), 0, .ColIndex("ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("ID")) = True
16            .ColKey(5) = "分组": .ColWidth(.ColIndex("分组")) = 2000: .ColAlignment(.ColIndex("分组")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("分组")) = "分组"
17            .Cell(flexcpAlignment, 0, .ColIndex("分组"), 0, .ColIndex("分组")) = flexAlignCenterCenter
18        End With

19        With Me.VSFType
20            .Rows = 2
21            .Cols = 5
22            .FixedRows = 1
23            .ColKey(0) = "ID": .ColWidth(.ColIndex("ID")) = 1000: .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID": .ColHidden(.ColIndex("ID")) = True
24            .ColKey(1) = "编码": .ColWidth(.ColIndex("编码")) = 1000: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
25            .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
26            .ColKey(2) = "分类": .ColWidth(.ColIndex("分类")) = 2000: .ColAlignment(.ColIndex("分类")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("分类")) = "分类"
27            .Cell(flexcpAlignment, 0, .ColIndex("分类"), 0, .ColIndex("分类")) = flexAlignCenterCenter
28            .ColKey(3) = "颜色": .ColWidth(.ColIndex("颜色")) = 700: .ColAlignment(.ColIndex("颜色")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("颜色")) = "颜色"
29            .Cell(flexcpAlignment, 0, .ColIndex("颜色"), 0, .ColIndex("颜色")) = flexAlignCenterCenter
30            .ColKey(4) = "耐受试验": .ColWidth(.ColIndex("耐受试验")) = 700: .ColAlignment(.ColIndex("耐受试验")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("耐受试验")) = "耐受试验"
31            .Cell(flexcpAlignment, 0, .ColIndex("耐受试验"), 0, .ColIndex("耐受试验")) = flexAlignCenterCenter
32        End With

33        With Me.VSFListDept
34            .Rows = 2
35            .Cols = 4
36            .FixedRows = 1
37            .ColKey(0) = "编码": .ColWidth(.ColIndex("编码")) = 2000: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
38            .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
39            .ColKey(1) = "小组名称": .ColWidth(.ColIndex("小组名称")) = 2500: .ColAlignment(.ColIndex("小组名称")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("小组名称")) = "小组名称"
40            .Cell(flexcpAlignment, 0, .ColIndex("小组名称"), 0, .ColIndex("小组名称")) = flexAlignCenterCenter
41            .ColKey(2) = "HIS部门编码": .ColWidth(.ColIndex("HIS部门编码")) = 2500: .ColAlignment(.ColIndex("HIS部门编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("HIS部门编码")) = "HIS部门编码"
42            .Cell(flexcpAlignment, 0, .ColIndex("HIS部门编码"), 0, .ColIndex("HIS部门编码")) = flexAlignCenterCenter
43            .ColKey(3) = "默认": .ColWidth(.ColIndex("默认")) = 700: .ColAlignment(.ColIndex("默认")) = flexAlignCenterCenter: .TextMatrix(0, .ColIndex("默认")) = "默认"
44        End With


45        Exit Sub
InitVSF_Error:
46        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(InitVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
47        Err.Clear
End Sub
Private Sub ReadTypeData()
          '功能   读入分类数据到列表中
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim intloop As Integer
          
1         On Error GoTo ReadTypeData_Error

2         VSFList.Rows = 1: VSFList.Rows = 2
3         VSFListDept.Rows = 1: VSFListDept.Rows = 2
          
4         strSQL = " select ID,编码,名称,颜色,是否耐受申请单 from 检验申请单 where nvl(科室ID,0) = [1] order by 编码 "
5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入申请单", Val(Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex)))
6         With Me.VSFType
7             .Rows = 1
8             Do Until rsTmp.EOF
9                 .Rows = .Rows + 1
10                .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
11                .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
12                .TextMatrix(.Rows - 1, .ColIndex("分类")) = rsTmp("名称") & ""
13                .Cell(flexcpBackColor, .Rows - 1, .ColIndex("颜色"), .Rows - 1, .ColIndex("颜色")) = Val(rsTmp("颜色") & "")
14                .TextMatrix(.Rows - 1, .ColIndex("耐受试验")) = IIf(Val(rsTmp("是否耐受申请单") & "") = 1, "是", "")
15                rsTmp.MoveNext
16            Loop
17            If .Rows = 1 Then
18                .Rows = .Rows + 1
19                .Row = 1
20                mlngkeyID = 0
21                Exit Sub
22            End If
23            If mlngkeyID = 0 Then
24                .Row = 1
25                Exit Sub
26            End If
27            For intloop = 1 To .Rows - 1
28                If .TextMatrix(intloop, .ColIndex("ID")) = mlngkeyID Then
29                    .Row = intloop
30                    Exit For
31                End If
32            Next
'33            mlngkeyID = 0
34        End With


35        Exit Sub
ReadTypeData_Error:
36        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(ReadTypeData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
37        Err.Clear
End Sub

Private Sub ReadItemData()
    '功能       读入分类明细
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strItem As String
    Dim strDefault As String
    
    On Error GoTo ReadItemData_Error
    
    '项目信息
    strSQL = "Select a.id, b.分类, b.编码, b.名称,a.排列顺序,d.名称 分组" & vbNewLine & _
             " From 检验申请单明细 A, 检验组合项目 B,检验申请单 c,检验申请单分组 d Where a.申请单id =c.id and A.组合id = B.Id and a.分组id=d.id(+) and b.停用日期 is null and a.申请单ID = [1] order by a.分组ID,a.排列顺序, b.编码 "
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入分类明细", mlngkeyID)
    With Me.VSFList
        .Rows = 1
        
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
            .TextMatrix(.Rows - 1, .ColIndex("组合项目")) = rsTmp("名称") & ""
            .TextMatrix(.Rows - 1, .ColIndex("分类")) = rsTmp("分类") & ""
            .TextMatrix(.Rows - 1, .ColIndex("排列顺序")) = rsTmp("排列顺序") & ""
            .TextMatrix(.Rows - 1, .ColIndex("分组")) = rsTmp("分组") & ""
            
            rsTmp.MoveNext
        Loop
        If .Rows = 1 Then .Rows = 2
    End With
    
    '执行小组
    strSQL = "Select 执行小组,默认执行小组 From 检验申请单 Where ID = [1]"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "申请单执行小组", mlngkeyID)
    If Not rsTmp.EOF Then
        strItem = rsTmp("执行小组") & ""
        strDefault = rsTmp("默认执行小组") & ""
    End If
    
    If gUserInfo.NodeNo = "-" Or mblnAllSite Then
        strSQL = "select 编码,名称 小组名称,HIS部门编码 from 检验小组记录" & vbNewLine & _
                "where 编码 in (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)))"
    Else
        strSQL = "select 编码,名称 小组名称,HIS部门编码 from 检验小组记录" & vbNewLine & _
                "where 编码 in (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) and (站点=[2] or 站点 is null)"
    End If
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入检验小组", strItem, gUserInfo.NodeNo)
    With Me.VSFListDept
        .Rows = 1
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
            .TextMatrix(.Rows - 1, .ColIndex("小组名称")) = rsTmp("小组名称") & ""
            .TextMatrix(.Rows - 1, .ColIndex("HIS部门编码")) = rsTmp("HIS部门编码") & ""
            
            .Cell(flexcpChecked, .Rows - 1, .ColIndex("默认"), .Rows - 1, .ColIndex("默认")) = 2
            .Cell(flexcpPictureAlignment, .Rows - 1, .ColIndex("默认"), .Rows - 1, .ColIndex("默认")) = flexAlignCenterCenter
            If InStr("," & strDefault & ",", "," & rsTmp("编码") & ",") > 0 Then
                .Cell(flexcpChecked, .Rows - 1, .ColIndex("默认"), .Rows - 1, .ColIndex("默认")) = 1
            End If
            rsTmp.MoveNext
        Loop
    End With


    Exit Sub
ReadItemData_Error:
    Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(ReadItemData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

End Sub

Private Sub DelType()
          '功能   删除分类
          Dim strSQL As String
          
1         On Error GoTo DelType_Error

2         If mlngkeyID = 0 Then
3             MsgBox "请选择一个分类!", vbInformation, "删除分类"
4             Exit Sub
5         End If
6         With Me.VSFType
7             If MsgBox("您确定要删除<" & .TextMatrix(.Row, .ColIndex("分类")) & ">分类?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
8                 Exit Sub
9             End If
              
10            strSQL = "Zl_检验申请单_Edit('" & 3 & "','" & mlngkeyID & "','','')"
11            Call ComExecuteProc(Sel_Lis_DB, strSQL, "删除分类")
12            SaveDBLog 18, 6, 0, "删除", "删除申请单:" & .TextMatrix(.Row, .ColIndex("分类")), 1012, "申请单设置"
13            mlngkeyID = 0
14            Call ReadTypeData
15        End With
          


16        Exit Sub
DelType_Error:
17        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(DelType)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear
         
End Sub

Private Sub ModifyType()
          '功能   修改分类
1         On Error GoTo ModifyType_Error

2         With Me.VSFType
3             If mlngkeyID = 0 Then
4                 MsgBox "请选择一个分类!", vbInformation, "修改分类"
5                 Exit Sub
6             End If
7             frmAppforBillDesignEditBill.ShowMe Me, mlngkeyID, Me.cboDeptSel.ItemData(Me.cboDeptSel.ListIndex), _
                              .TextMatrix(.Row, .ColIndex("编码")), .TextMatrix(.Row, .ColIndex("分类")), _
                              Val(.Cell(flexcpBackColor, .Row, .ColIndex("颜色"), .Row, .ColIndex("颜色"))), _
                              IIf(Trim(.TextMatrix(.Row, .ColIndex("耐受试验"))) = "是", True, False)
8             Call ReadTypeData
              
9         End With


10        Exit Sub
ModifyType_Error:
11        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(ModifyType)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
12        Err.Clear
End Sub

Private Sub ModifyItem()
    '功能   修改项目
    If mlngkeyID = 0 Then
        MsgBox "请选择一个分类!", vbInformation, "修改项目"
        Exit Sub
    End If
    frmAppforBillDesignEditItem.ShowMe Me, mlngkeyID, mblnAllSite, IIf(VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("耐受试验")) <> "", True, False)
    Call ReadTypeData
End Sub




Private Sub VSFType_RowColChange()
    With Me.VSFType
        If .Rows = 0 Then Exit Sub
        If .ColIndex("ID") = -1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ID"))) <= 0 Then Exit Sub
        
        'If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> mlngkeyID Then
            mlngkeyID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            Call ReadItemData
        'End If
    End With
End Sub

Private Sub LoadDept()
          '功能   读入科室
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
              
1         On Error GoTo LoadDept_Error

2         strSQL = "Select distinct ID, 编码, 名称 From 部门表 A, 部门性质说明 B Where A.Id = B.部门id And B.工作性质 In ('护理', '临床')"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入部门表")
4         With cboDeptSel
5             .Clear
6             .AddItem "所有科室"
7             .ItemData(.NewIndex) = 0
8             Do Until rsTmp.EOF
9                 .AddItem rsTmp("编码") & "-" & rsTmp("名称")
10                .ItemData(.NewIndex) = rsTmp("ID")
11                rsTmp.MoveNext
12            Loop
13            .ListIndex = 0
14        End With


15        Exit Sub
LoadDept_Error:
16        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(LoadDept)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear
          
End Sub
Private Sub ModifyGroup()
    '功能   修改单据执行小组
    With Me.VSFType
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 Then
            frmAppforBillDesignDept.ShowMe Me, .TextMatrix(.Row, .ColIndex("ID"))
            ReadItemData
        End If
    End With
End Sub

Private Sub SelectGroup()
    '功能   修改项目
    If mlngkeyID = 0 Then
        MsgBox "请选择一个分类!", vbInformation, "选择分组"
        Exit Sub
    End If
    frmAppforBillGroup.ShowMe Me, mlngkeyID, VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("分类"))
    Call ReadTypeData
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/6/7
'功    能:保存排序
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SaveItemSort()
          Dim lngRow As Long
          Dim strSQL As String
          Dim lngCount As Long
          
1         On Error GoTo SaveItemSort_Error

2         ReDim strArrSQL(0)
3         With Me.VSFList
4             For lngRow = 1 To .Rows - 1
5                 If .RowHidden(lngRow) = False Then
6                     lngCount = lngCount + 1
7                     strSQL = "Zl_申请单明细_Sort(" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & "," & lngCount & ")"
8                     Call ComExecuteProc(Sel_Lis_DB, strSQL, "申请单排序")
9                 End If
10            Next
11        End With
12        mblnItemSort = False


13        Exit Sub
SaveItemSort_Error:
14        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesign", "执行(SaveItemSort)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear
End Sub


'==========以下代码功能为:将右侧VSF列表中的数据拖动到左侧VSF中=============
'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:模拟拖动，点击列表时，将标签定位到点击的位置，方便跟随鼠标移动
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         On Error GoTo VSFList_MouseDown_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         With Me.VSFList
5             If .MouseRow <= 0 Or .MouseCol < 0 Then Exit Sub
6             Me.lblShow.Caption = .TextMatrix(.MouseRow, .ColIndex("组合项目"))
7             Me.lblShow.Tag = .TextMatrix(.MouseRow, .ColIndex("id")) & "|" & .TextMatrix(.MouseRow, .ColIndex("编码")) & "|" & .TextMatrix(.MouseRow, .ColIndex("分类")) & "|" & .TextMatrix(.MouseRow, .ColIndex("组合项目")) & "|" & .TextMatrix(.MouseRow, .ColIndex("分组"))
8             mlngMouseDownRow = .MouseRow
9         End With


10        Exit Sub
VSFList_MouseDown_Error:
11        MsgBox "执行(VSFList_MouseDown)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
12        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:标签跟随鼠标移动
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim lngRow As Long
          Dim lngCol As Long
          
1         On Error GoTo VSFList_MouseMove_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         If Me.lblShow.Caption = "" Then Exit Sub
5         With Me.lblShow
6             If .Visible = False Then .Visible = True
7             .Left = X - (.Width / 2)
8             .Top = Y - (.Height / 2)
9         End With
          
          '设置右侧列表在拖动鼠标时的效果
10        With Me.VSFList
11            lngRow = .MouseRow
12            lngCol = .MouseCol
13            If lngRow > -1 And lngCol > -1 Then
14                If mlngMouseRow <> lngRow And mlngMouseRow > 0 And lngRow > 0 Then
                      '移动到某一行上之后新增一个空行
15                    If mlngMouseRow <= .Rows - 1 Then
16                        If Trim(.TextMatrix(mlngMouseRow, .ColIndex("组合项目"))) = "" Then .RemoveItem mlngMouseRow    '先移除之前的空行
17                    End If
18                    Debug.Print 1
19                    .AddItem "", lngRow
20                    mlngMouseRow = lngRow
21                    .Row = mlngMouseRow
22                ElseIf mlngMouseRow = 0 And lngRow > 0 Then
23                    Debug.Print 2
24                    .AddItem "", lngRow
25                    mlngMouseRow = lngRow
26                ElseIf lngRow = .Rows - 1 And Trim(.TextMatrix(.Rows - 1, .ColIndex("组合项目"))) <> "" Then
                      '如果移动到最后一行,则在最后新增一行
27                    Debug.Print 3
28                    .AddItem "", .Rows
29                    mlngMouseRow = .Rows
30                End If
31            ElseIf lngRow = -1 And .Rows < 2 Then
32                Debug.Print 4
33                .Rows = .Rows + 1
34                mlngMouseRow = .Rows - 1
35            ElseIf lngRow = -1 And lngCol = -1 And mlngMouseRow <= .Rows - 1 Then
36                If Trim(.TextMatrix(mlngMouseRow, .ColIndex("组合项目"))) = "" Then
37                    .RemoveItem mlngMouseRow
38                End If
39            End If
40        End With
          

41        Exit Sub
VSFList_MouseMove_Error:
42        MsgBox "执行(VSFList_MouseMove)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
43        Err.Clear

End Sub


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:松开鼠标时,将拖动的值复制到右边的VSF中
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
1         On Error GoTo VSFList_MouseUp_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
4         With Me.VSFList
              '仅在右侧列表中拖动数据时
5             If .MouseCol > -1 And mlngMouseRow > 0 And mlngMouseRow <= .Rows - 1 Then
6                 .TextMatrix(mlngMouseRow, .ColIndex("id")) = Split(Me.lblShow.Tag, "|")(0): .ColAlignment(.ColIndex("id")) = flexAlignLeftCenter
7                 .TextMatrix(mlngMouseRow, .ColIndex("编码")) = Split(Me.lblShow.Tag, "|")(1): .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter
8                 .TextMatrix(mlngMouseRow, .ColIndex("分类")) = Split(Me.lblShow.Tag, "|")(2): .ColAlignment(.ColIndex("分类")) = flexAlignLeftCenter
9                 .TextMatrix(mlngMouseRow, .ColIndex("组合项目")) = Split(Me.lblShow.Tag, "|")(3): .ColAlignment(.ColIndex("组合项目")) = flexAlignLeftCenter
10                .TextMatrix(mlngMouseRow, .ColIndex("分组")) = Split(Me.lblShow.Tag, "|")(4): .ColAlignment(.ColIndex("分组")) = flexAlignLeftCenter
11                If mlngMouseDownRow > 0 And Me.lblShow.Visible = True Then
12                    If mlngMouseRow > mlngMouseDownRow Then
13                        .RemoveItem mlngMouseDownRow
14                    ElseIf mlngMouseDownRow + 1 <= .Rows - 1 Then
15                        If .MouseCol > -1 Then .RemoveItem mlngMouseDownRow + 1
16                    End If
17                End If
18            End If
              
19        End With
20        mlngMouseRow = 0
21        Me.lblShow.Caption = ""
22        If Me.lblShow.Visible = True Then Me.lblShow.Visible = False

23        Exit Sub
VSFList_MouseUp_Error:
24        MsgBox "执行(VSFList_MouseUp)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
          'WriteLog "执行(VSFList_MouseUp)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl
25        Err.Clear
End Sub
'=============================================================
