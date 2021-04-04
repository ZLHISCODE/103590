VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFinaceSuperviseCollectList 
   BorderStyle     =   0  'None
   Caption         =   "财务收款列表"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   5325
      ScaleHeight     =   2670
      ScaleWidth      =   2880
      TabIndex        =   8
      Top             =   6375
      Width           =   2910
      Begin VB.PictureBox picImgPlanGroup 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   60
         Width           =   210
         Begin VB.Image imgColPlanGroup 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseCollectList.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsGroup 
         Height          =   1800
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseCollectList.frx":054E
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
   Begin VB.PictureBox picGroupPerson 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   5235
      ScaleHeight     =   2670
      ScaleWidth      =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   2910
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   60
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseCollectList.frx":0603
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsGroupPerson 
         Height          =   1800
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseCollectList.frx":0B51
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
   Begin VB.PictureBox picRollingCurtain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   150
      ScaleHeight     =   1515
      ScaleWidth      =   10170
      TabIndex        =   2
      Top             =   420
      Width           =   10170
      Begin VB.PictureBox picImgPlanRC 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   210
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   11
         Top             =   570
         Width           =   210
         Begin VB.Image imgColPlanRC 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseCollectList.frx":0C06
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.TextBox txtNO 
         Height          =   345
         Left            =   2835
         TabIndex        =   6
         Top             =   113
         Width           =   2820
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   930
         Left            =   150
         TabIndex        =   3
         Top             =   525
         Width           =   10740
         _cx             =   18944
         _cy             =   1640
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseCollectList.frx":1154
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
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         Caption         =   "收费轧帐信息     轧帐单号"
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   285
      ScaleHeight     =   2520
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   3810
      Width           =   3435
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   -15
         TabIndex        =   1
         Top             =   -15
         Width           =   2865
         _Version        =   589884
         _ExtentX        =   5054
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFinaceSuperviseCollectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------
'程序入口相关
Public Enum CollectType  '收款类型
    EM_TY_收费员 = 1 '针对收费员收款
    EM_TY_小组 = 2 '针对小组收款
End Enum
Private mCollectType As CollectType
Private mstrPerson As String, mlngGroupID As Long
Private mlngModule As Long, mstrPrivs As String
'--------------------------------------------------------------

Private Enum mPgIndex
    EM_PG_收款及票据 = 250101
    EM_PG_财务组收款 = 250102
    EM_PG_收费员轧帐 = 250103
End Enum
Private Enum mPaneIndex
    EM_PN_轧帐信息 = 1
    EM_PN_明细列表 = 2
End Enum
Private mobjChargeBill As clsChargeBill
Private mblnNotBrush As Boolean
Private mblnLoadChargeAndBilled As Boolean '加载了收费及票据汇总
Private mblnLoadGrouped As Boolean '加载了组收款信息
Private mblnLoadGroupPersoned As Boolean '加载了组轧帐信息
 'PersonChange:事件,当通过单据号查找,发生人员姓名变化时触发该事件
Public Event PersonChange(ByVal strPerson As String, ByRef Cancel As Boolean)

Private Sub Form_Load()
    Set mobjChargeBill = New clsChargeBill
    Call InitGrid: Call ClearData
    Call InitPage: Call InitPanel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjChargeBill Is Nothing Then Set mobjChargeBill = Nothing
End Sub
Private Sub picGroup_Resize()
    Err = 0: On Error Resume Next
    With picGroup
        vsGroup.Left = .ScaleLeft
        vsGroup.Top = .ScaleTop
        vsGroup.Width = .ScaleWidth
        vsGroup.Height = .ScaleHeight
    End With
End Sub
Private Sub picGroupPerson_Resize()
    Err = 0: On Error Resume Next
    With picGroupPerson
        vsGroupPerson.Left = .ScaleLeft
        vsGroupPerson.Top = .ScaleTop
        vsGroupPerson.Width = .ScaleWidth
        vsGroupPerson.Height = .ScaleHeight
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub picRollingCurtain_Resize()
    Err = 0: On Error Resume Next
    With picRollingCurtain
        vsRollingCurtain.Left = .ScaleLeft
        vsRollingCurtain.Top = txtNO.Top + txtNO.Height + 50
        vsRollingCurtain.Width = .ScaleWidth
        vsRollingCurtain.Height = .ScaleHeight - vsRollingCurtain.Top - 100
    End With
    With picImgPlanRC
        .Left = vsRollingCurtain.Left + 60
        .Top = vsRollingCurtain.Top + 45
    End With
End Sub
Private Sub InitPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2013-09-22 17:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(EM_PG_收款及票据, "收款及票据汇总", mobjChargeBill.GetChargeAndBillTotalForm.hWnd, 0)
    objItem.Tag = EM_PG_收款及票据
    Set objItem = tbPage.InsertItem(EM_PG_财务组收款, "财务组收款信息", picGroup.hWnd, 0)
    objItem.Tag = EM_PG_财务组收款
    Set objItem = tbPage.InsertItem(EM_PG_收费员轧帐, "收费员轧帐信息", picGroupPerson.hWnd, 0)
    objItem.Tag = EM_PG_收费员轧帐
    mblnNotBrush = True
     With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    mblnNotBrush = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Sub
Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化区域控件
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-22 17:13:23
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, lngHeight As Long
    lngHeight = 1515 \ Screen.TwipsPerPixelY
    With dkpMan
        Set objPane = .CreatePane(mPaneIndex.EM_PN_轧帐信息, 400, lngHeight, DockLeftOf, Nothing)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picRollingCurtain.hWnd
        objPane.MinTrackSize.Height = lngHeight * 0.5
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_明细列表, 100, 200, DockBottomOf, objPane)
        objPane.Title = "":
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '编制:刘兴洪
    '日期:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHead As String, varData As Variant
    Dim lngWidth As Long
    strHead = "过滤,选择,ID,轧帐单号,轧帐类别,开始时间,终止时间,轧帐人,轧帐时间,收款员,收款部门,冲预交款,借入合计,借出合计,小组收款人,小组收款时间,轧帐说明"
    varData = Split(strHead, ",")
    
    With vsRollingCurtain
        .Clear: .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "过滤" Then .TextMatrix(0, i) = ""
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" _
                Or .ColKey(i) = "终止时间" Or .ColKey(i) = "轧帐时间" _
                Or .ColKey(i) = "轧帐类别" Then .ColData(i) = "1|0"
            If .ColKey(i) = "轧帐人" Or .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) = "选择" Or .ColKey(i) = "过滤" Or .ColKey(i) = "ID" Or .ColKey(i) = "收款部门" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "选择" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
        
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        .ColWidth(.ColIndex("过滤")) = 300
        lngWidth = .ColWidth(.ColIndex("选择"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
        .ColWidth(.ColIndex("选择")) = lngWidth
        .Editable = IIf(zlStr.IsHavePrivs(mstrPrivs, "轧帐收款"), flexEDKbdMouse, flexEDNone)
        .Tag = ""
    End With
    
    '组收款信息
    With vsGroup
        strHead = "过滤,ID,收款单号,收款部门,冲预交款,借入合计,借出合计,收款人,收款时间,收款说明"
        varData = Split(strHead, ",")
        .Clear: .Rows = 2: .Cols = UBound(varData) + 1
        .ExtendLastCol = True
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "过滤" Then .TextMatrix(0, i) = ""
            If .ColKey(i) = "过滤" Or .ColKey(i) = "ID" Or .ColKey(i) = "收款部门" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "收款单号" Or .ColKey(i) = "收款时间" Then .ColData(i) = "1|0"
            If .ColKey(i) = "收款部门" Or .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) Like "*单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsGroup, Me.Name, "组收款信息列表", False
        .Tag = ""
    End With
    
    '组人员轧帐记录
    strHead = "过滤,ID,轧帐单号,开始时间,终止时间,轧帐人,轧帐时间,收款部门,冲预交款,借入合计,借出合计,小组收款人,小组收款时间,轧帐说明"
    varData = Split(strHead, ",")
    With vsGroupPerson
        .Clear: .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "过滤" Then .TextMatrix(0, i) = ""
            If .ColKey(i) = "过滤" Or .ColKey(i) = "ID" Or .ColKey(i) = "收款部门" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "轧帐时间" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColData(i) = "1|0"
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) = "收款部门" Or .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsGroupPerson, Me.Name, "小组收费员轧帐信息列表", False
        .Tag = ""
    End With
End Sub
Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面信息
    '编制:刘兴洪
    '日期:2013-09-26 11:42:00
    '--------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsRollingCurtain
        .Clear 1: .Rows = 2
    End With
    Call mobjChargeBill.ClearChargeAndBillTotalForm
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlInitVar(ByVal bytType As CollectType, ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '        bytType-1-收费员;2-小组
    '       strPrivs-权限串
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mCollectType = bytType
    Call InitGrid
    Call SetCtrlVisible
End Sub
Public Function zlLoadCollectData(ByVal bytType As CollectType, ByVal strPerson As String, Optional ByVal lngGroupID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '入参:strPerson-指定的人员
    '        bytType-1-收费员;2-小组
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-26 11:38:15
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng轧帐ID As Long, strNO As String, lngWidth As Long
    Dim strWhere As String
    mCollectType = bytType: mstrPerson = strPerson
    mlngGroupID = lngGroupID
    On Error GoTo errHandle
    Call SetCtrlVisible
    zlLoadCollectData = LoadCollectData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetCtrlVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Visible属性
    '编制:刘兴洪
    '日期:2013-09-27 16:16:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With tbPage
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = EM_PG_财务组收款 Or Val(tbPage.Item(i).Tag) = EM_PG_收费员轧帐 Then
                tbPage.Item(i).Visible = mCollectType = EM_TY_小组
            End If
        Next
    End With
End Sub

Public Function LoadCollectData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '入参:strPerson-指定的人员
    '        bytType-1-收费员;2-小组
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-26 11:38:15
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng轧帐ID As Long, strNO As String, lngWidth As Long
    Dim strWhere As String, bytType As Byte, i As Long
    Dim blnCancel As Boolean
    On Error GoTo errHandle
    mblnLoadChargeAndBilled = False: mblnLoadGroupPersoned = False
    strNO = txtNO.Text
    If mstrPerson = "" And strNO = "" Then Call ClearData: LoadCollectData = True: Exit Function
    If strNO <> "" Then
        strWhere = " And A.NO=[3]"
    Else
        strWhere = " And Instr(',' || [2] || ',',',' || A.收款员 || ',') > 0"
    End If
    If mlngGroupID <> 0 And mCollectType = EM_TY_小组 Then
        strWhere = strWhere & " And A.缴款组ID = [4]"
    End If
    If mCollectType = EM_TY_收费员 Then
        strSQL = "" & _
        "   Select " & IIf(zlStr.IsHavePrivs(mstrPrivs, "轧帐收款"), "-1", "0") & " as 选择,a.Id,a.No As 轧帐单号, " & _
        "         Substr(Decode(是否挂号,1,',挂号','') || Decode(是否就诊卡,1,',就诊卡','') || Decode(是否消费卡,1,',消费卡','') || Decode(是否收费,1,',收费','') || Decode(是否结帐,1,',结帐','') || Decode(预交类别,1,',预交',2,',门诊预交',3,',住院预交',''),2) As 轧帐类别, " & _
        "        a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
        "         a.收款员 ,b.名称 As 收款部门, " & _
        "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
        "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
        "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计," & _
        "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间, " & _
        "         a.摘要 As 轧帐说明" & _
        "  From 人员收缴记录 A, 部门表 B " & _
        "  Where a.收款部门id = b.Id(+) And a.记录性质 = [1] " & strWhere & _
        "        And A.作废时间 is Null and A.财务收款ID is null And A.小组收款时间 Is Null " & _
        "  Order by 登记时间 desc,轧帐单号 desc"
    ElseIf mCollectType = EM_TY_小组 Then
        strSQL = "" & _
        "   Select  " & IIf(zlStr.IsHavePrivs(mstrPrivs, "轧帐收款"), "-1", "0") & " as 选择,a.Id,a.No As 轧帐单号,'' as 轧帐类别, a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
        "         a.收款员 ,b.名称 As 收款部门, " & _
        "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
        "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
        "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计," & _
        "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间, " & _
        "         a.摘要 As 轧帐说明" & _
        "  From 人员收缴记录 A, 部门表 B " & _
        "  Where a.收款部门id = b.Id(+) And a.记录性质 = [1] " & strWhere & _
        "               And A.作废时间 is Null and A.财务收款ID is null  " & _
        "  Order by 登记时间 desc,轧帐单号 desc,小组收款时间 desc"
    End If
    bytType = IIf(mCollectType = EM_TY_小组, 3, 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytType, mstrPerson, strNO, mlngGroupID)
    
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        If rsTemp.RecordCount <> 0 Then
            Do While Not rsTemp.EOF
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTemp!ID)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = Nvl(rsTemp!轧帐单号)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐类别")) = Nvl(rsTemp!轧帐类别)
                .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Nvl(rsTemp!开始时间)
                .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Nvl(rsTemp!终止时间)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐人")) = Nvl(rsTemp!轧帐人)
                .TextMatrix(.Rows - 1, .ColIndex("收款员")) = Nvl(rsTemp!收款员)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = Nvl(rsTemp!轧帐时间)
                .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTemp!冲预交款)
                .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTemp!借入合计)
                .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTemp!借出合计)
                .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = Nvl(rsTemp!小组收款人)
                .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = Nvl(rsTemp!小组收款时间)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐说明")) = Nvl(rsTemp!轧帐说明)
                .Rows = .Rows + 1
                rsTemp.MoveNext
            Loop
            rsTemp.MoveFirst
            'Set .DataSource = rsTemp
            If strNO <> "" Then
                '财务组时，mstrPerson 是所有成员
                If InStr("," & mstrPerson & ",", "," & Nvl(rsTemp!收款员) & ",") = 0 Then
                    '找到单据的人员不对,需要定位到该人员上
                    blnCancel = False
                    RaiseEvent PersonChange(Nvl(rsTemp!收款员), blnCancel)
                    If blnCancel Then
                        Call ClearData: LoadCollectData = True: Exit Function
                    End If
                    mstrPerson = Nvl(rsTemp!收款员)
                End If
            End If
        End If
        
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "选择" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        If .Rows > 2 Then .Rows = .Rows - 1
        Call .AutoSize(1, .Cols - 1)
        lngWidth = .ColWidth(.ColIndex("选择"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
         .ColWidth(.ColIndex("选择")) = lngWidth
        If .Enabled And .Visible Then .SetFocus
    End With
    '加载明细数据
    LoadCollectData = LoadDetail
    mblnNotBrush = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Function

Private Function LoadGroupCollectData(ByVal lng轧帐ID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '编制:刘兴洪
    '日期:2013-09-27 11:08:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select /*+ rule */a.Id,a.No As 收款单号,  " & _
    "         b.名称 As 收款部门, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计, " & _
    "         a.摘要 As 收款说明, " & _
    "         a.登记人 as 收款人,To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 收款时间  " & _
    "  From 人员收缴记录 A, 部门表 B " & _
    "  Where a.收款部门id = b.Id(+) And a.记录性质 = 3 And A.小组轧账ID= [1]" & _
    "  Order by 登记时间 desc,收款单号 desc,小组收款时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng轧帐ID)
    With vsGroup
        .Clear 1: .Rows = 2
        .FixedRows = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("收款单号")) = Nvl(rsTemp!收款单号)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("收款人")) = Nvl(rsTemp!收款人)
            .TextMatrix(.Rows - 1, .ColIndex("收款时间")) = Nvl(rsTemp!收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("收款说明")) = Nvl(rsTemp!收款说明)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "收款单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        
        If .Rows > 2 Then .Rows = .Rows - 1
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsGroup, Me.Name, "组收款信息列表", False
        If .Enabled And .Visible Then .SetFocus
    End With
    LoadGroupCollectData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function LoadGroupPersonData(ByVal lng轧帐ID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '入参:strPerson-指定的人员
    '        bytType-1-收费员;2-小组
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-26 11:38:15
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNO As String, lngWidth As Long
    Dim strWhere As String, bytType As Byte, i As Long
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */-1 as 选择,a.Id,a.No As 轧帐单号, a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
    "         b.名称 As 收款部门, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计," & _
    "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间, " & _
    "         a.摘要 As 轧帐说明" & _
    "  From 人员收缴记录 A, 部门表 B " & _
    "  Where a.收款部门id = b.Id(+) And a.记录性质 = 1 " & _
    "               And A.作废时间 is Null and A.财务收款ID is null And  A.小组轧账ID= [1]   " & _
    "  Order by 登记时间 desc,轧帐单号 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng轧帐ID)
    With vsGroupPerson
        .Clear 1: .Rows = 2
        .FixedRows = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = Nvl(rsTemp!轧帐单号)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Nvl(rsTemp!开始时间)
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Nvl(rsTemp!终止时间)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐人")) = Nvl(rsTemp!轧帐人)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = Nvl(rsTemp!轧帐时间)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐说明")) = Nvl(rsTemp!轧帐说明)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = Nvl(rsTemp!小组收款人)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = Nvl(rsTemp!小组收款时间)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "选择" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        If .Rows > 2 Then .Rows = .Rows - 1
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsGroupPerson, Me.Name, "小组收费员轧帐信息列表", False
        If .Enabled And .Visible Then .SetFocus
    End With
    LoadGroupPersonData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotBrush Then Exit Sub
    Call LoadDetail
End Sub

Private Sub txtNO_Change()
    txtNO.Tag = ""
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO: zlCommFun.OpenIme False
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mCollectType = EM_TY_小组 Then
        txtNO.Text = GetFullNO(Trim(txtNO.Text), 139)
    Else
        txtNO.Text = GetFullNO(Trim(txtNO.Text), 137)
    End If
    Call LoadCollectData
End Sub

Private Sub vsGroup_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsGroupPerson_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
    If OldRow = NewRow Or mblnNotBrush Then Exit Sub
    '将所有都设置为False
    mblnLoadChargeAndBilled = False: mblnLoadGroupPersoned = False
    mblnLoadGrouped = False
   Call LoadDetail
End Sub
Private Sub vsRollingCurtain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("选择")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("选择")
            Cancel = True
        Case Else
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
    vsRollingCurtain.Tag = "1"
End Sub
Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub
Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsGroupPerson_GotFocus()
    Call zl_VsGridGotFocus(vsGroupPerson)
    vsGroupPerson.Tag = "1"
End Sub
Private Sub vsGroupPerson_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsGroupPerson)
    vsGroupPerson.Tag = "0"
End Sub
Private Sub vsGroupPerson_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGroupPerson, Me.Name, "小组收费员轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsGroupPerson_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsGroupPerson, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsGroupPerson_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGroupPerson, Me.Name, "小组收费员轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsGroup_GotFocus()
    Call zl_VsGridGotFocus(vsGroup)
    vsGroup.Tag = "1"
End Sub
Private Sub vsGroup_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsGroup)
    vsGroup.Tag = "0"
End Sub
Private Sub vsGroup_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGroup, Me.Name, "组收款信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsGroup_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsGroup, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsGroup_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGroup, Me.Name, "组收款信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Function LoadDetail() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-27 16:37:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng轧帐ID As Long, strNO As String
    On Error GoTo errHandle
    With vsRollingCurtain
        If .Row > 0 Then
            lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        End If
    End With
    
    If lng轧帐ID = 0 Then
        vsGroup.Clear 1: vsGroup.Rows = 2
        vsGroupPerson.Clear 1: vsGroupPerson.Rows = 2
        mobjChargeBill.ClearChargeAndBillTotalForm
        LoadDetail = True
        Exit Function
    End If
    '加载明细数据
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收款及票据
        If mblnLoadChargeAndBilled = True Then LoadDetail = True: Exit Function
        LoadDetail = mobjChargeBill.LoadChargeAndBillTotalData(Me, _
            mlngModule, mstrPrivs, IIf(mCollectType = EM_TY_收费员, 1, 3), lng轧帐ID)
            mblnLoadChargeAndBilled = True
    Case EM_PG_财务组收款
        If mblnLoadGrouped = True Then LoadDetail = True: Exit Function
        LoadDetail = LoadGroupCollectData(lng轧帐ID)
        mblnLoadGrouped = True
    Case EM_PG_收费员轧帐
        If mblnLoadGroupPersoned = True Then LoadDetail = True: Exit Function
        LoadDetail = LoadGroupPersonData(lng轧帐ID)
        mblnLoadGroupPersoned = True
    Case Else
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlClearData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面信息或某些变量信息
    '编制:刘兴洪
    '日期:2013-09-29 11:22:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearData
    txtNO.Text = "": mstrPerson = ""
End Sub

Public Property Get GetSelRollingCurtainIds() As String
    '获取选中的轧帐ID
    Dim i As Long, strIDs As String, lngID As Long
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            lngID = Val(.TextMatrix(i, .ColIndex("ID")))
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And lngID <> 0 Then
                strIDs = strIDs & "," & lngID
            End If
        Next
        lngID = 0
        If strIDs <> "" Then
            GetSelRollingCurtainIds = Mid(strIDs, 2): Exit Property
        End If
        '没有时,以当前选中的行作为缺省的轧帐ID
        If .Row > 0 And .Row <= .Rows - 1 Then
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            If lngID <> 0 Then
                GetSelRollingCurtainIds = lngID: Exit Property
            End If
        End If
        GetSelRollingCurtainIds = ""
    End With
End Property
Public Property Get IsSelRollingCurtainRecord() As Boolean
    '是否选中了轧账记录
    Dim i As Long, strIDs As String, lngID As Long
    With vsRollingCurtain
        '以当前先是否选中为依据
        If .Row > 0 And .Row <= .Rows - 1 Then
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            If lngID <> 0 Then
                IsSelRollingCurtainRecord = True: Exit Property
            End If
        End If
        For i = 1 To .Rows - 1
            lngID = Val(.TextMatrix(i, .ColIndex("ID")))
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And lngID <> 0 Then
                  IsSelRollingCurtainRecord = True: Exit Property
            End If
        Next
    End With
End Property
Public Property Get GetRollingCurtainID() As Long
    '是否选中了轧账记录
    With vsRollingCurtain
        '以当前先是否选中为依据
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Property
        GetRollingCurtainID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property
Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng轧帐ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = False
    End With
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
    '1-收费员轧帐；2-小组收款;3-小组轧帐;4-财务收款。
   Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, IIf(mCollectType = EM_TY_收费员, 1, 3), CStr(lng轧帐ID), dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
End Sub
Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng轧帐ID As Long, blnDel As Boolean
    Dim strNO As String
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("轧帐单号")))
        lng轧帐ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = False
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "类型=" & IIf(mCollectType = EM_TY_收费员, 1, 2), _
        "轧帐ID=" & lng轧帐ID, _
        "NO=" & strNO, _
        "收款员=" & mstrPerson)
End Sub

Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    
    If Val(vsRollingCurtain.Tag) = 0 Then
        '打印收款及票据汇总
        With vsRollingCurtain
            If .Row < 1 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("轧帐单号")) = "" Then Exit Sub
        End With
        Call mobjChargeBill.zlPrint(bytMode, "", ""): Exit Sub
    End If
    
    '输出轧帐信息
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr单位名称 & "未收款轧帐清册"
    Set objRow = New zlTabAppRow
    objRow.Add "收费员:" & mstrPerson
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsRollingCurtain
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
              zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Property Get GetCashMoney() As Double
    '获取现金金额
   GetCashMoney = mobjChargeBill.GetChargeAndBillTotalForm.GetCashMoney
End Property

Private Sub imgColPlanRC_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanRC.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanRC.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlanRC.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlanRC_Click()
    Call imgColPlanRC_Click
End Sub

Private Sub imgColPlanGroup_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanGroup.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanGroup.Height
    Call frmVsColSel.ShowColSet(Me, Me.Name, vsGroup, lngLeft, lngTop, imgColPlanGroup.Height)
    zl_vsGrid_Para_Save mlngModule, vsGroup, Me.Name, "组收款信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlanGroup_Click()
    Call imgColPlanGroup_Click
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGroupPerson, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsGroupPerson, Me.Caption, "小组收费员轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub

