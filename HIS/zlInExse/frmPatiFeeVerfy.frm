VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiFeeVerfy 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   270
      ScaleHeight     =   2955
      ScaleWidth      =   6240
      TabIndex        =   11
      Top             =   5355
      Width           =   6240
      Begin VB.PictureBox picImgList 
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   90
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   30
         Width           =   210
         Begin VB.Image imgColList 
            Height          =   195
            Index           =   0
            Left            =   0
            Picture         =   "frmPatiFeeVerfy.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Bindings        =   "frmPatiFeeVerfy.frx":054E
         Height          =   1395
         Left            =   15
         TabIndex        =   12
         Top             =   0
         Width           =   7110
         _cx             =   12541
         _cy             =   2461
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
         BackColorSel    =   16444122
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFeeVerfy.frx":0562
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
   End
   Begin VB.PictureBox pic医嘱 
      BorderStyle     =   0  'None
      Height          =   3765
      Left            =   330
      ScaleHeight     =   3765
      ScaleWidth      =   7245
      TabIndex        =   9
      Top             =   1170
      Width           =   7245
      Begin VB.PictureBox picImgList 
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   120
         Width           =   210
         Begin VB.Image imgColList 
            Height          =   195
            Index           =   1
            Left            =   0
            Picture         =   "frmPatiFeeVerfy.frx":059E
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3555
         Left            =   45
         TabIndex        =   10
         Top             =   90
         Width           =   5925
         _cx             =   10451
         _cy             =   6271
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
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
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
         FixedCols       =   2
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFeeVerfy.frx":0AEC
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin MSComctlLib.ImageList img16 
            Left            =   1920
            Top             =   600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":0B87
                  Key             =   "签名"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":0ED9
                  Key             =   "屏蔽打印"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":1473
                  Key             =   ""
                  Object.Tag             =   "3"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img16dbl 
            Left            =   2535
            Top             =   615
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPatiFeeVerfy.frx":1A0D
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1230
      ScaleHeight     =   600
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   240
      Width           =   11880
      Begin VB.CheckBox chk记帐 
         Caption         =   "只含记帐费用的医嘱"
         Height          =   180
         Left            =   7695
         TabIndex        =   1
         Top             =   165
         Width           =   1935
      End
      Begin VB.CheckBox chkType 
         Caption         =   "临嘱(&T)"
         Height          =   300
         Index           =   1
         Left            =   6615
         TabIndex        =   2
         Top             =   120
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.CheckBox chkType 
         Caption         =   "长嘱(&L)"
         Height          =   300
         Index           =   0
         Left            =   5640
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboCons 
         Height          =   300
         Index           =   1
         Left            =   2895
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   165
         Width           =   1875
      End
      Begin VB.ComboBox cboCons 
         Height          =   300
         Index           =   0
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "执行科室"
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "诊疗类别"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblCons 
         AutoSize        =   -1  'True
         Caption         =   "期效"
         Height          =   180
         Index           =   2
         Left            =   5115
         TabIndex        =   6
         Top             =   195
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imgFlag 
      Left            =   0
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":20A7
            Key             =   "紧急"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":22C1
            Key             =   "补录"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":27DB
            Key             =   "未用"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":29F5
            Key             =   "报告已阅"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":2F0F
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3429
            Key             =   "自由"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3643
            Key             =   "待审核"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":3BDD
            Key             =   "审核通过"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4177
            Key             =   "审核未通过"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPass 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4711
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":4FFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiFeeVerfy.frx":52F9
            Key             =   ""
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmPatiFeeVerfy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long, mlng主页ID As Long
Private mlngModule As Long
'------------------------------------------------------------------
'局部变量
Private Enum mEM_Pancel
    EM_过滤 = 1
    EM_医嘱 = 2
    EM_费用 = 3
End Enum
Private mrsSkinTest As ADODB.Recordset
Private mrsDefine As ADODB.Recordset    '医嘱内容定义
Private mblnUnload As Boolean
Private mblnDataMove As Boolean '是否历史表中数据
Private mlngFontSize As Long '字号大小
Private mstr诊疗类别 As String
Private mstr执行科室ID As String
Private mrs诊疗类别 As ADODB.Recordset
Private Enum CboIdx
    EM_IDX诊疗类别 = 0
    EM_IDX执行科室 = 1
End Enum
Private mblnNotClick As Boolean
Private mrs医嘱 As ADODB.Recordset
Private mblnChangeData As Boolean
Private mstrPrivs As String
Private mbytFontSize As Byte
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置字体大小
    '编制:刘兴洪
    '日期:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("刘") + 20
        Case UCase("VsFlexGrid")
            Call zlControl.VSFSetFontSize(objCtrl, mbytFontSize)
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("刘兴" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("刘") * 1.5
        Case UCase("textBox")
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = CtlFont.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    Call picFilter_Resize
End Sub

Public Function ShowData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnDataMove As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示数据
    '编制:刘兴洪
    '日期:2012-05-31 11:01:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    mblnDataMove = blnDataMove
    Call Load医嘱
    ShowData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域设置
    '编制:刘兴洪
    '日期:2012-05-30 13:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_过滤, 200, 580, DockTopOf, Nothing)
    panThis.Title = "过滤条件": panThis.Handle = picFilter.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.MaxTrackSize.Height = picFilter.Height \ Screen.TwipsPerPixelY
    panThis.MinTrackSize.Height = picFilter.Height \ Screen.TwipsPerPixelY
    panThis.Tag = mEM_Pancel.EM_过滤
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_医嘱, 250, 580, DockBottomOf, panThis)
    panThis.Title = "": panThis.Handle = pic医嘱.hWnd
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = mEM_Pancel.EM_医嘱
    
    Set panThis = dkpMan.CreatePane(mEM_Pancel.EM_费用, 250, 580, DockBottomOf, panThis)
    panThis.Title = "": panThis.Handle = picFeeList.hWnd
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = mEM_Pancel.EM_费用
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

Private Sub cboCons_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
     Call Load医嘱(True)
End Sub

Private Sub chkType_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        mblnNotClick = True
        chkType(Index).Value = 1
        mblnNotClick = False
    End If
    Load医嘱 True
End Sub

Private Sub chk记帐_Click()
    If mblnNotClick Then Exit Sub
    Load医嘱 True
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mEM_Pancel.EM_过滤
        Item.Handle = picFilter.hWnd
    Case mEM_Pancel.EM_医嘱
        Item.Handle = pic医嘱.hWnd
    Case mEM_Pancel.EM_费用
        Item.Handle = picFeeList.hWnd
    End Select
End Sub
 
Private Function Load医嘱(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载医嘱数据
    '编制:刘兴洪
    '日期:2012-05-30 14:20:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    Dim strSQL As String, strWhere As String, lng医嘱ID As Long, lng相关ID As Long
    Dim byt医嘱期效 As Byte, str重整 As String, dt重整 As Date
    Dim str记帐 As String, strFeeTable As String
    Dim strFilter As String
    
    Call InitAdvice
    If mlng病人ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
    strFilter = ""
    On Error GoTo ErrHand:
    With vsAdvice
        If .Row > 0 Then
            i = .ColIndex("医嘱ID")
            If i > 0 Then
                lng医嘱ID = Val(.TextMatrix(.Row, i))  '记录当前行
            End If
        End If
    End With
    If Not (chkType(0).Value = 1 And chkType(1).Value = 1) Then
        strFilter = strFilter & " And 期效='" & IIf(chkType(0).Value = 1, "长嘱", "临嘱") & "'"
    End If
    
    '只显示包含未记帐费用的医嘱
    If chk记帐.Value = 1 Then
         strFilter = strFilter & " And 费用医嘱ID<>0"
      '  str记帐 = _
            " And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From (Select Nvl(C.相关id, C.ID) As 医嘱id" & vbNewLine & _
            "              From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C" & vbNewLine & _
            "              Where A.医嘱id = C.ID And A.NO = B.NO And A.记录性质 = B.记录性质 And A.记录性质 = 2 And B.记录状态 = 0 And" & vbNewLine & _
            "                    C.病人id = [1] And C.主页id = [2]" & ")" & vbNewLine & _
            "       Where A.ID = 医嘱id Or A.相关id = 医嘱id )"
    End If
    With cboCons(CboIdx.EM_IDX诊疗类别)
        If .ListIndex >= 0 Then
            If Chr(.ItemData(.ListIndex)) <> 0 Then
                strFilter = strFilter & " and 诊疗类别='" & Chr(.ItemData(.ListIndex)) & "'"
            End If
        End If
    End With
    
    With cboCons(CboIdx.EM_IDX执行科室)
        If .ListIndex >= 0 Then
            If .ItemData(.ListIndex) <> 0 Then
                strFilter = strFilter & " And 执行科室ID=" & .ItemData(.ListIndex) & ""
            End If
        End If
    End With
    
    strFeeTable = "" & _
        "   Select nvl(B.相关ID,B.ID) as 医嘱序号,Sum(nvl(应收金额,0)) as 应收金额,Sum(nvl(实收金额,0)) as 实收金额  " & _
        "   From 住院费用记录 A,病人医嘱记录 B " & _
        "   Where A.医嘱序号=B.ID and  A.病人ID=[1] and A.主页ID=[2] " & _
        "   Group by nvl(B.相关ID,B.ID)"
    strFeeTable = strFeeTable & " Union All " & Replace(strFeeTable, "住院费用记录", "门诊费用记录")
    
    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法'总量及用法计算
    strSQL = _
    "   Select /*+ RULE */ A.ID as 医嘱ID,A.相关ID,A.序号,Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态," & _
    "               Nvl(A.诊疗类别,'*') as 诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志,nvl(是否费用审核,0) as 审核," & _
    "               A.审查结果,Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
    "               To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,A.医嘱内容,Null as 内容,A.皮试结果 as 皮试," & _
    "               Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位)," & _
    "               '4',A.总给予量||G.计算单位,'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
    "               Decode(A.单次用量,NULL,NULL,A.单次用量||Decode(A.诊疗类别,'4',G.计算单位,B.计算单位)) as 单量,A.天数," & _
    "               A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
    "               A.医生嘱托,A.执行时间方案 as 执行时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间," & _
    "               nvl(E.ID,decode(nvl(A.执行性质,0),0,-1,5,-2,NULL)) as 执行科室ID," & _
    "               Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
    "               Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质," & _
    "               To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & _
    "               Decode(A.医嘱状态,1,'新开',2,'疑问',3,'校对',4,'作废',5,'重整',6,'暂停',7,'启用',8,'停止',9,'确认停止') as 状态," & _
    "               A.开嘱医生,To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间,A.校对护士,To_Char(A.校对时间,'YYYY-MM-DD HH24:MI') as 校对时间," & _
    "               A.停嘱医生,To_Char(A.停嘱时间,'YYYY-MM-DD HH24:MI') as 停嘱时间,F.操作人员 as 停嘱护士," & _
    "               To_Char(A.确认停嘱时间,'YYYY-MM-DD HH24:MI') as 确认停嘱时间,A.诊疗项目ID,B.试管编码,A.执行标记,A.屏蔽打印,A.前提ID,Decode(S.签名ID,NULL,0,1) as 签名否," & _
    "               M.病历文件ID as 文件ID,Nvl(N.通用,0) as 报告项,Y.病历ID as 报告ID,Y.查阅状态,A.收费细目ID,B.计算单位 as 单量单位,A.开嘱科室ID,A.审核状态, " & _
    "               A.申请序号,A1.应收金额,A1.实收金额, nvl(A1.医嘱序号,0) as 费用医嘱ID"
    strSQL = strSQL & _
    " From 病人医嘱记录 A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,收费项目目录 G," & _
    "       病人医嘱状态 F,病人医嘱状态 S,病人医嘱报告 Y,病历单据应用 M,病历文件列表 N," & _
    "      (" & strFeeTable & ") A1" & _
    " Where A.诊疗项目ID=B.ID(+) And nvl(a.相关ID,a.Id) =A1.医嘱序号(+) And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
    "       And A.收费细目ID=D.药品ID(+) And A.收费细目ID=G.ID(+) And A.ID=Y.医嘱ID(+)" & _
    "       And (Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL) Or A.诊疗类别='E' And B.操作类型='8')" & _
    "       And A.ID=F.医嘱ID(+) And F.操作类型(+)=9 And A.ID=S.医嘱ID And S.操作类型=1" & _
    "       And A.诊疗项目ID=M.诊疗项目ID(+) And M.应用场合(+)=2 And M.病历文件ID=N.ID(+) And N.种类(+)=7" & _
    "       And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1"
    
    '重整显示格式处理
    strSQL = strSQL & " Order by 婴儿ID,序号"
    
    '访问历史空间处理
    If mblnDataMove Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "住院费用记录", "H住院费用记录")
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
    End If
    If Not blnFilter Or mrs医嘱 Is Nothing Or mblnChangeData Then
        Set mrs医嘱 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID, byt医嘱期效, dt重整)
        With mrs医嘱
            mstr诊疗类别 = "": mstr执行科室ID = ""
            Do While Not .EOF
                If InStr(1, mstr诊疗类别 & ",", "," & Nvl(!诊疗类别) & ",") = 0 And Nvl(!诊疗类别) <> "" Then
                    mstr诊疗类别 = mstr诊疗类别 & "," & Nvl(!诊疗类别)
                End If
                If InStr(1, mstr执行科室ID & ",", "," & Val(Nvl(!执行科室ID)) & ",") = 0 Then
                    mstr执行科室ID = mstr执行科室ID & "," & Val(Nvl(!执行科室ID))
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
        Call InitCons
    End If
    
    mrs医嘱.Filter = 0
    If strFilter <> "" Then
        strFilter = Mid(strFilter, 5)
        mrs医嘱.Filter = strFilter
    End If
    With vsAdvice
            .Redraw = flexRDNone: .MergeCells = flexMergeNever
            '绑定时按设计时的FormatString恢复一些缺省值(固定行列数，固定行列文字及行列对齐,尺寸,可见)
            'FormatString在运行时赋值无效
            '如果AutoResize=True,则所有列宽或行高被自动调整(根据AutoSizeMode)
            '如果WordWrap=True,则行高会被自动调整
            .WordWrap = False
            Set .DataSource = mrs医嘱
            If mrs医嘱.RecordCount = 0 Then .Rows = .FixedRows + 1

            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            
           .WordWrap = True
            For i = 0 To .Cols - 1
                .ColKey(i) = Switch(i = 0, "医嘱图标", i = 1, "报告标志", True, Trim(.TextMatrix(0, i)))
                .FixedAlignment(i) = flexAlignCenterCenter
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                Select Case .ColKey(i)
                Case "序号", "诊疗类别", "操作类型", "毒理分类", "标志", "审查结果", _
                        "医嘱状态", "执行标记", "屏蔽打印", "签名否", "报告项", "查阅状态", "审核状态", _
                        "申请序号"
                    .ColHidden(i) = True: .ColData(i) = "-1|1"
                Case "皮试", "总量", "单量", "天数", "频率"
                    .ColHidden(i) = True
                Case "执行时间", "终止时间", "执行性质", "上次执行"
                    .ColHidden(i) = True
                Case "状态", "开嘱时间", "校对护士", "校对时间"
                    .ColHidden(i) = True
                Case "校对时间", "停嘱时间", "停嘱护士", "确认停嘱时间"
                    .ColHidden(i) = True
                Case "审核"
                    .ColData(i) = "1|0": .ColAlignment(i) = flexAlignCenterCenter
                    .ColDataType(i) = flexDTBoolean
                Case "医嘱内容", "期效"
                    .ColData(i) = "1|0": .ColAlignment(i) = flexAlignLeftCenter
                Case Else
                    If .ColKey(i) Like "*ID" Then
                        .ColHidden(i) = True: .ColData(i) = "-1|1"
                    End If
                End Select
            Next
            For i = 1 To .Rows - 1
                .Cell(flexcpData, i, .ColIndex("应收金额")) = .TextMatrix(i, .ColIndex("应收金额"))
                .Cell(flexcpData, i, .ColIndex("实收金额")) = .TextMatrix(i, .ColIndex("实收金额"))
            Next
    End With
    If Not mrs医嘱.EOF Then
        Call ReModifyData(dt重整)
    End If
    
    With vsAdvice
'            '计算费用合计
'            For i = 1 To .Rows - 1
'                lng医嘱ID = Val(.TextMatrix(i, .ColIndex("医嘱ID")))
'                lng相关ID = Val(.TextMatrix(i, .ColIndex("相关ID")))
'                If lng医嘱ID = 0 Then lng医嘱ID = -1
'                If lng相关ID = 0 Then lng相关ID = -1
'
'            Next
                
        '自动调整行高
        If InStr("2505,3345,1005,1335", .ColWidth(.ColIndex("用法"))) > 0 Then .ColWidth(.ColIndex("用法")) = IIf(mlngFontSize = 9, 2505, 3345)   '用户未改该列宽时才设置
        .AutoSize .ColIndex("内容"), .ColIndex("用法")
        .ColWidth(.ColIndex("开始时间")) = IIf(mlngFontSize = 9, 1130, 1510)
        '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
        '电子签名图标对齐
        .Cell(flexcpPictureAlignment, .FixedRows, .ColIndex("医嘱内容"), .Rows - 1, .ColIndex("医嘱内容")) = 0
        i = 0
         If lng医嘱ID <> 0 Then i = vsAdvice.FindRow(CStr(lng医嘱ID), , .ColIndex("医嘱ID"))
        If i < .FixedRows Then i = .FixedRows
        .Row = i
        If .RowHidden(.Row) Then
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            For i = .Row - 1 To .FixedRows Step -1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
             .AddItem "":  .Row = .Rows - 1
        End If
        .Col = .FixedCols
        Call vsAdvice.ShowCell(.Row, .Col)
        zl_vsGrid_Para_Restore mlngModule, vsAdvice, Me.Caption, "医嘱审核列头信息"
        If mrs医嘱.RecordCount <> 0 And InStr(";" & mstrPrivs, ";审核病人;") > 0 Then
            vsAdvice.Editable = flexEDKbd
        Else
            vsAdvice.Editable = flexEDNone
        End If
        .Redraw = flexRDDirect
    End With
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    mblnChangeData = False
    Load医嘱 = True
    Exit Function
ErrHand:
    vsAdvice.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitCons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 初始化条件
    '编制:刘兴洪
    '日期:2012-05-31 16:11:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPreKey As String, rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    mblnNotClick = True
    With cboCons(CboIdx.EM_IDX诊疗类别)
        If .ListIndex >= 0 And .ListCount > 0 Then strPreKey = Chr(.ItemData(.ListIndex))
        .Clear
        .AddItem "所有类别"
        .ItemData(.NewIndex) = Asc("0")
        If mrs诊疗类别.RecordCount <> 0 Then mrs诊疗类别.MoveFirst
        Do While Not mrs诊疗类别.EOF
            If InStr(mstr诊疗类别 & ",", "," & Nvl(mrs诊疗类别!编码) & ",") > 0 Then
                .AddItem Nvl(mrs诊疗类别!名称)
                .ItemData(.NewIndex) = Asc(Nvl(mrs诊疗类别!编码))
                If strPreKey = Nvl(mrs诊疗类别!编码) Then
                    .ListIndex = .NewIndex
                End If
            End If
            mrs诊疗类别.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
    End With
   If mstr执行科室ID = "" Then Exit Sub
   strSQL = "" & _
    "   Select /*+ RULE */A.ID,A.编码,A.名称" & _
    "   From 部门表 A, (Select Column_Value From Table(f_num2list([1]))) J " & _
    "   Where A.ID=J. Column_Value " & _
    "   Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "0" & mstr执行科室ID)
    
   With cboCons(CboIdx.EM_IDX执行科室)
        If .ListIndex >= 0 And .ListCount > 0 Then strPreKey = .ItemData(.ListIndex)
        .Clear
        .AddItem "所有执行科室"
        If InStr(1, mstr执行科室ID & ",", ",-1,") > 0 Then
            .AddItem "<叮嘱>": .ItemData(.NewIndex) = -1
            If strPreKey = "-1" Then .ListIndex = .NewIndex
        End If
        If InStr(1, mstr执行科室ID & ",", ",-2,") > 0 Then
            .AddItem "院外执行": .ItemData(.NewIndex) = -2
            If strPreKey = "-2" Then .ListIndex = .NewIndex
        End If
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
                If strPreKey = Nvl(rsTemp!ID) Then
                    .ListIndex = .NewIndex
                End If
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub ReModifyData(ByVal dat重整 As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新修正数据
    '入参:blnFilter-true,表示是条件过滤
    '编制:刘兴洪
    '日期:2012-05-30 15:05:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln给药途径  As Boolean, bln中药用法  As Boolean, bln采集方法 As Boolean, bln输血途径  As Boolean
    Dim i As Long, j As Long, strTemp  As String, strFormat As String
    Dim dtCurdate As Date, dtDate1 As Date, dtDate2 As Date
    Dim strCurDate As String, strDate1 As String, strDate2 As String
    Dim blnDo As Boolean, lngTop As Long, strTime As String, blnFirst As Boolean
    Dim strType As String '诊疗类别
    
    dtCurdate = zlDatabase.Currentdate: strCurDate = Format(dtCurdate, "yyyy-MM-DD")
    dtDate1 = DateAdd("D", -1, dtCurdate): strDate1 = Format(dtDate1, "yyyy-MM-DD")
    dtDate2 = DateAdd("D", -2, dtCurdate): strDate2 = Format(dtDate2, "yyyy-MM-DD")
    
    On Error GoTo errHandle
    With vsAdvice
        i = .FixedRows:
        Do While i <= .Rows - 1
            .Cell(flexcpData, i, .ColIndex("开始时间")) = CStr(.TextMatrix(i, .ColIndex("开始时间"))) '合理用药接口调用时取数
            strTemp = Format(.TextMatrix(i, .ColIndex("开始时间")), "yyyy-MM-dd")
            Select Case strTemp
            Case strCurDate '当天
                    .TextMatrix(i, .ColIndex("开始时间")) = "今 天 " & Format(.TextMatrix(i, .ColIndex("开始时间")), "HH:mm")
            Case strDate1   '昨天
                    .TextMatrix(i, .ColIndex("开始时间")) = "昨 天 " & Format(.TextMatrix(i, .ColIndex("开始时间")), "HH:mm")
            Case strDate2  '前天
                    .TextMatrix(i, .ColIndex("开始时间")) = "前 天 " & Format(.TextMatrix(i, .ColIndex("开始时间")), "HH:mm")
            Case Else
                    .TextMatrix(i, .ColIndex("开始时间")) = Format(.TextMatrix(i, .ColIndex("开始时间")), "MM-dd HH:mm")
            End Select
            '应收金额,实收金额
            .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            
            bln给药途径 = False: bln中药用法 = False: bln采集方法 = False: bln输血途径 = False
            If Trim(.TextMatrix(i, .ColIndex("诊疗类别"))) = "E" Then   '治疗
                     If Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) = Val(.TextMatrix(i, .ColIndex("医嘱ID"))) Then
                        Select Case .TextMatrix(i - 1, .ColIndex("诊疗类别"))
                        Case "5", "6"   '西药和中成药
                                bln给药途径 = True
                                For j = i - 1 To .FixedRows Step -1
                                      If Val(.TextMatrix(j, .ColIndex("相关ID"))) <> Val(.TextMatrix(i, .ColIndex("医嘱ID"))) Then Exit For
                                     '显示成药的给药途径
                                    .TextMatrix(j, .ColIndex("用法")) = .TextMatrix(i, .ColIndex("用法"))
                                    '合并用法列:用法 频率 天数
                                    strFormat = .TextMatrix(j, .ColIndex("用法"))
                                    strTemp = .TextMatrix(j, .ColIndex("频率"))
                                    If strTemp <> "" Then strFormat = strFormat & IIf(strFormat <> "", ",", "") & strTemp
                                    strTemp = .TextMatrix(j, .ColIndex("天数"))
                                    If strTemp <> "" Then
                                        strFormat = strFormat & IIf(strFormat <> "", ",", "") & "共" & strTemp & "天"
                                    End If
                                     .TextMatrix(j, .ColIndex("用法")) = strFormat
                                     
                                     ''显示成药的执行性质
                                    If Val(.TextMatrix(j, .ColIndex("执行性质"))) = 5 And Val(.TextMatrix(i, .ColIndex("执行性质"))) <> 5 Then
                                        .TextMatrix(j, .ColIndex("执行性质")) = "自备药"
                                    ElseIf Val(.TextMatrix(j, .ColIndex("执行性质"))) <> 5 And Val(.TextMatrix(i, .ColIndex("执行性质"))) = 5 Then
                                        .TextMatrix(j, .ColIndex("执行性质")) = "离院带药"
                                    Else
                                        .TextMatrix(j, .ColIndex("执行性质")) = IIf(Val(.TextMatrix(j, .ColIndex("执行标记"))) = 1, "自取药", "")
                                    End If
                                     .TextMatrix(j, .ColIndex("皮试")) = .TextMatrix(i, .ColIndex("医生嘱托"))
                                    If .TextMatrix(j, .ColIndex("皮试")) <> "" Then
                                        .TextMatrix(j, .ColIndex("内容")) = .TextMatrix(j, .ColIndex("内容")) & "," & .TextMatrix(j, .ColIndex("皮试"))
                                    End If
                                Next
                            Case "7", "C" '中草药/检验
                                    bln中药用法 = .TextMatrix(i - 1, .ColIndex("诊疗类别")) = "7" '中药用法行
                                    bln采集方法 = .TextMatrix(i - 1, .ColIndex("诊疗类别")) = "C" '采集方法行
                                    If bln采集方法 Then
                                        '采集方式的管码与一并的第一个检验相同
                                          j = .FindRow(.TextMatrix(i, .ColIndex("医嘱ID")), .FixedRows, .ColIndex("相关ID"))
                                        .TextMatrix(i, .ColIndex("试管编码")) = .TextMatrix(j, .ColIndex("试管编码"))
                                     End If
                                    '显示中药配方或检验组合的执行科室
                                    .TextMatrix(i, .ColIndex("执行科室")) = .TextMatrix(i - 1, .ColIndex("执行科室"))
                                     .TextMatrix(i, .ColIndex("执行性质")) = ""
                                     If bln中药用法 Then
                                        '显示中药配方执行性质
                                        If Val(.TextMatrix(i - 1, .ColIndex("执行性质"))) = 5 And Val(.TextMatrix(i, .ColIndex("执行性质"))) <> 5 Then
                                            .TextMatrix(i, .ColIndex("执行性质")) = "自备药"
                                        ElseIf Val(.TextMatrix(i - 1, .ColIndex("执行性质"))) <> 5 And Val(.TextMatrix(i, .ColIndex("执行性质"))) = 5 Then
                                            .TextMatrix(i, .ColIndex("执行性质")) = "离院带药"
                                        Else
                                            .TextMatrix(i, .ColIndex("执行性质")) = IIf(Val(.TextMatrix(i - 1, .ColIndex("执行标记"))) = 1, "自取药", "")
                                        End If
                                     End If
                                    '删除单味中药行,以及检验组合中的检验项目
                                    For j = i - 1 To .FixedRows Step -1
                                        If Val(.TextMatrix(j, .ColIndex("相关ID"))) <> Val(.TextMatrix(i, .ColIndex("医嘱ID"))) Then Exit For
                                        .TextMatrix(i, .ColIndex("报告项")) = .TextMatrix(j, .ColIndex("报告项")) '检验、配方以首行医嘱为准
                                        .TextMatrix(i, .ColIndex("文件ID")) = .TextMatrix(j, .ColIndex("文件ID"))
                                        .RemoveItem j: i = i - 1
                                    Next
                            End Select
                     ElseIf .TextMatrix(i - 1, .ColIndex("诊疗类别")) = "K" And Val(.TextMatrix(i - 1, .ColIndex("医嘱ID"))) = Val(.TextMatrix(i, .ColIndex("相关ID"))) Then
                            bln输血途径 = True
                            '显示输血途径
                            .TextMatrix(i - 1, .ColIndex("用法")) = .TextMatrix(i, .ColIndex("用法"))
                      Else
                         .TextMatrix(i, .ColIndex("执行性质")) = ""
                     End If
            End If
            
           '处理可见行的的一些标识:排开不可见但暂时未删除的行
            If Not (bln给药途径 Or bln输血途径) And .TextMatrix(i, .ColIndex("诊疗类别")) <> "7" Then
                If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                '重整医嘱分隔
                If Val(.TextMatrix(i, .ColIndex("医嘱ID"))) = 0 And .Rows > .FixedRows + 1 Then
                    .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = "━━━━ 重整医嘱(" & Format(dat重整, "yyyy-MM-dd HH:mm") & ") ━━━━"
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
                    .Cell(flexcpAlignment, i, .FixedCols, i, .Cols - 1) = 4
                    .MergeRow(i) = True
                    .MergeCells = flexMergeFree
                End If
                If Left(.TextMatrix(i, .ColIndex("总量")), 1) = "." Then
                    .TextMatrix(i, .ColIndex("总量")) = "0" & .TextMatrix(i, .ColIndex("总量"))
                End If
                If Left(.TextMatrix(i, .ColIndex("单量")), 1) = "." Then
                    .TextMatrix(i, .ColIndex("单量")) = "0" & .TextMatrix(i, .ColIndex("单量"))
                End If
            
                If Val(.TextMatrix(i, .ColIndex("报告ID"))) <> 0 Then
                        If Val(.TextMatrix(i, .ColIndex("查阅状态"))) = 0 Then
                            Set .Cell(flexcpPicture, i, .ColIndex("报告标志")) = imgFlag.ListImages("报告").Picture
                        ElseIf Val(.TextMatrix(i, .ColIndex("查阅状态"))) = 1 Then
                            Set .Cell(flexcpPicture, i, .ColIndex("报告标志")) = imgFlag.ListImages("报告已阅").Picture
                        End If
                End If
                
                '医嘱颜色
                blnDo = False
                If Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 2 Then
                    '校对疑问
                    If lngTop = 0 Then lngTop = i '有删除行也不会影响取值
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80& '深红
                    blnDo = True
                ElseIf Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 4 Then
                    '已作废
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                    .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                    blnDo = True
                ElseIf InStr(",8,9,", Val(.TextMatrix(i, .ColIndex("医嘱状态")))) > 0 Then
                    '已停止,已确认停止:长嘱都以终止时间进行判断
                    If strCurDate >= .TextMatrix(i, .ColIndex("终止时间")) Or .TextMatrix(i, .ColIndex("期效")) = "临嘱" Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 8 And strCurDate < .TextMatrix(i, .ColIndex("终止时间")) Then
                        '长嘱,停止后,停止时间未到这一种情况
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080 '浅蓝
                        blnDo = True
                    End If
                ElseIf Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 6 Then
                    '已暂停
                    strTime = Format(GetAdviceTime(Val(.TextMatrix(i, .ColIndex("医嘱ID"))), 6), "yyyy-MM-dd HH:mm")
                    If strCurDate >= strTime Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '深绿
                        blnDo = True
                    Else
                        '长嘱,暂停后,暂停时间未到这一种情况
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080 '浅蓝
                        blnDo = True
                    End If
                ElseIf Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 7 Then
                    '已启用
                    strTime = Format(GetAdviceTime(Val(.TextMatrix(i, .ColIndex("医嘱ID"))), 7), "yyyy-MM-dd HH:mm")
                    If strCurDate < strTime Then
                        '长嘱,启用后,启用时间未到这一种情况
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H4AAD00 '浅绿
                        blnDo = True
                    End If
                End If
                If Not blnDo Then
                    If lngTop = 0 Then lngTop = i
                    If Val(.TextMatrix(i, .ColIndex("医嘱状态"))) <> 1 And Val(.TextMatrix(i, .ColIndex("医嘱ID"))) <> 0 Then
                        '已通过校对(也包含后续的多个状态)
                        If Format(.TextMatrix(i, .ColIndex("上次执行")), "YYYY-MM-DD") >= Format(strCurDate, "YYYY-MM-DD") Then   '当天已发送的(长嘱可能发送到将来)
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HA08000               '海蓝
                        Else
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '深蓝
                        End If
                    End If
                End If
                '校对后术前术后医嘱红色显示
                If .TextMatrix(i, .ColIndex("诊疗类别")) = "Z" And (Val(.TextMatrix(i, .ColIndex("操作类型"))) = 4 Or Val(.TextMatrix(i, .ColIndex("操作类型"))) = 14) _
                    And InStr(",-1,1,2,4,", Val(.TextMatrix(i, .ColIndex("医嘱状态")))) = 0 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '红色
                End If
                
                '发送后转科医嘱红色显示
                If .TextMatrix(i, .ColIndex("诊疗类别")) = "Z" And Val(.TextMatrix(i, .ColIndex("操作类型"))) = 3 And Val(.TextMatrix(i, .ColIndex("医嘱状态"))) = 8 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '红色
                End If
            
                '毒麻精药品标识:中药配方及组成味中药不处理
                If .TextMatrix(i, .ColIndex("毒理分类")) <> "" Then
                    If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(i, .ColIndex("毒理分类"))) > 0 Then
                        .Cell(flexcpFontBold, i, .ColIndex("医嘱内容")) = True
                    End If
                End If
                '皮试结果标识
                If .TextMatrix(i, .ColIndex("诊疗类别")) = "E" And .TextMatrix(i, .ColIndex("操作类型")) = "1" And .TextMatrix(i, .ColIndex("皮试")) <> "" Then
                    j = zl获取皮试结果(Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))), .TextMatrix(i, .ColIndex("皮试")))
                    .Cell(flexcpForeColor, i, .ColIndex("皮试")) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, .ColIndex("皮试")))
                End If
                '自由录入
                If Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))) = 0 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("自由").Picture
                End If
                '紧急标志:一并给药只显示在第一行
                blnFirst = True
                If InStr(",5,6,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Val(.TextMatrix(i, .ColIndex("标志"))) = 1 Then
                        Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("紧急").Picture
                    ElseIf Val(.TextMatrix(i, .ColIndex("标志"))) = 2 Then
                        Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("补录").Picture
                    End If
                    
                    If Val(.TextMatrix(i, .ColIndex("医嘱状态"))) < 2 Then   '新开或暂存的医嘱
                        Select Case Val(.TextMatrix(i, .ColIndex("审核状态")))
                        '0-无需审核，1-待审核，2-审核通过，3-审核未通过
                            Case 1
                                Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("待审核").Picture
                            Case 2
                                Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("审核通过").Picture
                            Case 3
                                Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("审核未通过").Picture
                            Case Else
                        End Select
                        .Cell(flexcpPictureAlignment, i, .ColIndex("医嘱图标")) = 4
                    End If
                End If
                '未用医嘱标识
                If Val(.TextMatrix(i, .ColIndex("执行标记"))) = -1 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("未用").Picture
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                End If
                
                '未用医嘱标识
                If Val(.TextMatrix(i, .ColIndex("执行标记"))) = -1 Then
                    Set .Cell(flexcpPicture, i, .ColIndex("医嘱图标")) = imgFlag.ListImages("未用").Picture
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                End If
                'Pass:根据审查结果显示警示灯
                If .TextMatrix(i, .ColIndex("审查结果")) <> "" Then
                    Set .Cell(flexcpPicture, i, .ColIndex("审查结果")) = imgPass.ListImages(Val(.TextMatrix(i, .ColIndex("审查结果"))) + 1).Picture
                    .TextMatrix(i, .ColIndex("审查结果")) = ""
                End If
                '电子签名标识，屏蔽打印标识
                Call SetAdviceIcon(i)
            End If
            If bln给药途径 Or bln输血途径 Then
                 .RemoveItem i
            Else
                '组合医嘱内容
                strFormat = .TextMatrix(i, .ColIndex("医嘱内容"))
                If .TextMatrix(i, .ColIndex("诊疗类别")) <> "Z" And Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))) <> 0 And InStr(strFormat, "重整医嘱") = 0 Then
                    '医嘱内容定义中包含了相关项时，不再重复组合
                    mrsDefine.Filter = "诊疗类别='" & .TextMatrix(i, .ColIndex("诊疗类别")) & "'"
                
                    strTemp = .TextMatrix(i, .ColIndex("皮试"))
                    If strTemp <> "" Then strFormat = strFormat & strTemp
                    
                    If Not (InStr("5,6,7", .TextMatrix(i, .ColIndex("诊疗类别"))) = 0 And .TextMatrix(i, .ColIndex("频率")) = "一次性") Then
                        blnDo = True
                        If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                        If blnDo Then
                            strTemp = .TextMatrix(i, .ColIndex("总量"))
                            If strTemp <> "" Then strFormat = strFormat & ",共" & strTemp
                        End If
                        
                        blnDo = True
                        If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[单量]") = 0
                        If blnDo Then
                            strTemp = .TextMatrix(i, .ColIndex("单量"))
                            If strTemp <> "" Then strFormat = strFormat & ",每次" & strTemp
                        End If
                    End If
                End If
                
                .TextMatrix(i, .ColIndex("内容")) = strFormat
                '合并用法列:用法 频率 天数(一并给药的在前面已处理)
                If .TextMatrix(i, .ColIndex("诊疗类别")) <> "Z" And Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))) <> 0 And InStr(strFormat, "重整医嘱") = 0 Then
                    strFormat = .TextMatrix(i, .ColIndex("用法"))
                    strTemp = .TextMatrix(i, .ColIndex("频率"))
                    If strTemp <> "" Then strFormat = strFormat & IIf(strFormat <> "", ",", "") & strTemp
                    
                    strTemp = .TextMatrix(i, .ColIndex("天数"))
                    If strTemp <> "" Then
                        strFormat = strFormat & IIf(strFormat <> "", ",", "") & "共" & strTemp & "天"
                    End If
                    .TextMatrix(i, .ColIndex("用法")) = strFormat
                End If
                i = i + 1
            End If
        Loop
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ModiyStartDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改开始时间
    '编制:刘兴洪
    '日期:2012-05-30 15:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Sub
Private Sub InitAdvice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除网格行
    '编制:刘兴洪
    '日期:2012-05-30 14:54:18
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsAdvice
        .Clear 1
        .Rows = vsAdvice.FixedRows + 1
        .Editable = flexEDNone
        For i = .FixedRows To .Rows - 1
            .RowHidden(i) = False
        Next
    End With
End Sub
Private Function GetAdviceTime(ByVal lng医嘱ID As Long, ByVal int类型 As Integer) As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取医嘱指定操作的时间
    '编制:刘兴洪
    '日期:2012-05-30 16:44:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(操作时间) as 时间 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, int类型)
    If rsTemp.EOF Then Exit Function
    If Not IsNull(rsTemp!时间) Then GetAdviceTime = rsTemp!时间
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zl获取皮试结果(ByVal lng项目id As Long, ByVal str结果 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据皮试结果标注，返回阴阳性
    '入参:str结果=皮试结果标注符号,如"(+)"
    '返回:-1-阴性,1-阳性,0-无结果
    '编制:刘兴洪
    '日期:2012-05-30 16:50:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim var阳性 As Variant, var阴性 As Variant
    Dim strSQL As String, i As Integer
    On Error GoTo errH
    If mrsSkinTest Is Nothing Then
        strSQL = "Select ID,Nvl(标本部位,'阳性(+);阴性(-)') as 标注 From 诊疗项目目录 Where 类别='E' And 操作类型='1'"
        Set mrsSkinTest = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    mrsSkinTest.Filter = "ID=" & lng项目id
    If mrsSkinTest.EOF Then Exit Function
    var阳性 = Split(Split(mrsSkinTest!标注, ";")(0), ",")
    var阴性 = Split(Split(mrsSkinTest!标注, ";")(1), ",")
    
    For i = 0 To UBound(var阳性)
        If Right(var阳性(i), Len(str结果)) = str结果 Then
            zl获取皮试结果 = 1: Exit Function
        End If
    Next
    For i = 0 To UBound(var阴性)
        If Right(var阴性(i), Len(str结果)) = str结果 Then
            zl获取皮试结果 = -1: Exit Function
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    '(518716, 1)
    Call InitData
    Call InitPancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mrsSkinTest Is Nothing Then
        If mrsSkinTest.State <> 1 Then mrsSkinTest.Close
    End If
    Set mrsSkinTest = Nothing
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "医嘱审核列头信息"
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
End Sub

Private Sub SetAdviceIcon(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前行的内容设置医嘱内容的图标标识
    '编制:刘兴洪
    '日期:2012-05-30 17:02:56
    '说明:注意是单行设置，不是一组设置
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsAdvice
        If Val(.TextMatrix(lngRow, .ColIndex("签名否"))) = 1 And Val(.TextMatrix(lngRow, .ColIndex("屏蔽打印"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("医嘱内容")) = img16dbl.ListImages(1).Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("内容")) = img16dbl.ListImages(1).Picture
        ElseIf Val(.TextMatrix(lngRow, .ColIndex("签名否"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("医嘱内容")) = img16.ListImages("签名").Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("内容")) = img16.ListImages("签名").Picture
        ElseIf Val(.TextMatrix(lngRow, .ColIndex("屏蔽打印"))) = 1 Then
            Set .Cell(flexcpPicture, lngRow, .ColIndex("医嘱内容")) = img16.ListImages("屏蔽打印").Picture
            Set .Cell(flexcpPicture, lngRow, .ColIndex("内容")) = img16.ListImages("屏蔽打印").Picture
        Else
            Set .Cell(flexcpPicture, lngRow, .ColIndex("医嘱内容")) = Nothing
            Set .Cell(flexcpPicture, lngRow, .ColIndex("内容")) = Nothing
        End If
    End With
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2012-05-30 17:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    mlngFontSize = 9: mlngModule = glngModul: mblnChangeData = False
    strSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
    Set mrsDefine = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    strSQL = "Select 编码,名称 From 诊疗项目类别"
    Set mrs诊疗类别 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mstrPrivs = gstrPrivs
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckPatiDataMoved(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定病人的数据是否已转出
    '编制:刘兴洪
    '日期:2012-05-30 17:30:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    strSQL = "Select 数据转出 From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查转出", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!数据转出) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

 
Private Sub imgColList_Click(Index As Integer)
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgList(Index).hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList(Index).Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, IIf(Index = 1, vsAdvice, vsFeeList), lngLeft, lngTop, imgColList(Index).Height)
    If Index = 1 Then
        zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "医嘱审核列头信息"
    Else
        zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
    End If
End Sub

Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    cboCons(0).Top = (picFilter.ScaleHeight - cboCons(0).Height) \ 2
    cboCons(1).Top = cboCons(0).Top
    lblCons(0).Top = cboCons(0).Top + (cboCons(0).Height - lblCons(0).Height) \ 2
    lblCons(1).Top = lblCons(0).Top
    lblCons(2).Top = lblCons(0).Top
    chkType(0).Top = cboCons(0).Top + (cboCons(0).Height - chkType(0).Height) \ 2
    chkType(1).Top = chkType(0).Top
    chk记帐.Top = chkType(0).Top
    
    cboCons(0).Left = lblCons(0).Left + lblCons(0).Width + 10
    lblCons(1).Left = cboCons(0).Left + cboCons(0).Width + 50
    cboCons(1).Left = lblCons(1).Left + lblCons(1).Width + 10
    lblCons(2).Left = cboCons(1).Left + cboCons(1).Width + 50
    chkType(0).Left = lblCons(2).Left + lblCons(2).Width + 10
    chkType(1).Left = chkType(0).Left + chkType(0).Width + 20
    chk记帐.Left = chkType(1).Left + chkType(1).Width + 50
    
End Sub

Private Sub picImgList_Click(Index As Integer)
    Call imgColList_Click(Index)
End Sub
Private Sub pic医嘱_Resize()
    Err = 0: On Error Resume Next
    With pic医嘱
        
        vsAdvice.Left = .ScaleLeft + 10
        vsAdvice.Top = .ScaleTop
        vsAdvice.Height = .ScaleHeight - 20
        vsAdvice.Width = .ScaleWidth
        picImgList(1).Top = vsAdvice.Top + 30
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft + 10
        vsFeeList.Top = .ScaleTop
        vsFeeList.Height = .ScaleHeight
        vsFeeList.Width = .ScaleWidth - 20
    End With
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lng医嘱ID As Long, bln审核 As Boolean
    With vsAdvice
        If Row <= 0 Then Exit Sub
        If Col <> .ColIndex("审核") Then Exit Sub
        lng医嘱ID = Val(.TextMatrix(Row, .ColIndex("医嘱ID")))
        If lng医嘱ID = 0 Then Exit Sub
        bln审核 = Val(.TextMatrix(Row, Col)) <> 0
        bln审核 = SaveData(lng医嘱ID, bln审核)
        If bln审核 = False Then
            .TextMatrix(Row, Col) = IIf(bln审核, 0, 1)
        End If
    End With
End Sub

Private Sub vsAdvice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng医嘱ID As Long, lngCol As Long
    Dim lng相关ID As Long
    
    If NewRow = OldRow And vsAdvice.Visible = False Then Exit Sub
    On Error GoTo errHandle
    With vsAdvice
        lngCol = .ColIndex("医嘱ID")
        If NewRow > 0 And lngCol >= 0 Then lng医嘱ID = Val(.TextMatrix(NewRow, lngCol))
        lngCol = .ColIndex("相关ID")
        If NewRow > 0 And lngCol >= 0 Then lng相关ID = Val(.TextMatrix(NewRow, lngCol))
    End With
    Call LoadFeeList(lng医嘱ID, lng相关ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    With vsAdvice
        Select Case Col
        Case .ColIndex("医嘱内容"), .ColIndex("内容")
                .AutoSize Col, .ColIndex("用法")
        Case .ColIndex("皮试")
            If .ColWidth(Col) > 1200 Then .ColWidth(Col) = 1200
        Case Else
            If Row = -1 Then
                lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
                If vsAdvice.ColWidth(Col) < lngW Then
                    vsAdvice.ColWidth(Col) = lngW
                ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
                    vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
                End If
            End If
        End Select
    End With
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Me.Caption, "医嘱审核列头信息"
End Sub



Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        Select Case Col
        Case .ColIndex("审核")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then Cancel = True: Exit Sub
        Select Case Col
        Case .ColIndex("审查结果"), .ColIndex("审核")
             Cancel = True: Exit Sub
        Case Else
        End Select
    End With
    If Row = -1 Then
        With vsAdvice
            If Col <= .FixedCols - 1 Then
                Cancel = True
            ElseIf Col = .ColIndex("审查结果") Then
                Cancel = True
            End If
        End With
    End If
End Sub
Private Sub LoadFeeList(ByVal lng医嘱ID As Long, Optional lng相关ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费用明细
    '编制:刘兴洪
    '日期:2012-05-30 17:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dbl应收 As Double, dbl实收 As Double
    Dim i As Long
    On Error GoTo errHandle
    strSQL = _
        " Select a.No, a.价格父号, a.序号, a.收费细目id, a.执行部门id, a.记录状态, a.执行状态," & _
        "        a.登记时间, a.付数, a.数次, a.标准单价, a.应收金额, a.实收金额, a.医嘱序号" & _
        " From 住院费用记录 A" & _
        " Where a.病人id= [1] And (a.主页id = [2] Or a.主页id Is Null)"
    strSQL = strSQL & " Union All " & Replace(strSQL, "住院费用记录", "门诊费用记录")
    
    strSQL = _
    " Select b.No as 单据号, b.序号, b.收费细目id,B.记录状态,Q.名称 As 收费名称," & _
    "        q.规格, b.数量, Decode(r.住院单位, Null, q.计算单位, r.住院单位) As 单位, " & _
    "        to_char(b.单价 * Nvl(r.住院包装, 1),'999999999999990.99') As 单价, " & _
    "        b.应收金额, b.实收金额, j.名称 As 执行科室, " & _
    "        Decode(b.记录状态, 2, -1 * b.执行状态 || '次退费', Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '部份执行', '异常收费')) As 执行状态, " & _
    "        To_Char(b.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间 " & _
    " From (Select a.No, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, a.执行部门id, a.记录状态, a.执行状态, a. 登记时间, " & _
    "              Avg(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量, Avg(a.标准单价) As 单价," & _
    "              Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额" & _
    "       From (" & strSQL & ") A,病人医嘱记录 B" & _
    "       Where A.医嘱序号=B.ID And (B.ID in ([3],[4]) or nvl(B.相关ID,-2) in ([3],[4] )) " & _
    "       Group By a.No, Nvl(a.价格父号, a.序号), a.收费细目id, a.执行部门id, a.记录状态, a.执行状态, a.登记时间) B," & _
    "      部门表 J, 收费项目目录 Q, 药品规格 R " & _
    " Where b.执行部门id = j.Id(+) And b.收费细目id = q.Id And b.收费细目id = r.药品id(+) " & _
    " Order By 登记时间, 单据号,序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng医嘱ID, lng相关ID)
    With vsFeeList
        .Clear 1
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 0 To .Cols - 1
            .ColKey(i) = IIf(i = 0, "固定标志", .TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "记录状态" Then
                .ColHidden(i) = True:  .ColData(i) = "-1|1"
            End If
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "固定标志"
                    .ColData(i) = "-1|1"
            Case "单据号", "单位"
                .ColData(i) = "1|0"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "收费名称"
                .ColData(i) = "1|0"
            Case "数量", "应收金额", "实收金额", "单价"
                .ColData(i) = "1|0"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
        .ColWidth(.ColIndex("固定标志")) = 300
        dbl应收 = 0: dbl实收 = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(Val(.TextMatrix(i, .ColIndex("数量"))), 5)
            .TextMatrix(i, .ColIndex("单价")) = Format(Val(.TextMatrix(i, .ColIndex("单价"))), "######" & gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), "######" & gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), "######" & gstrDec)
            dbl应收 = dbl应收 + Val(.TextMatrix(i, .ColIndex("应收金额")))
            dbl实收 = dbl实收 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            Select Case Val(.TextMatrix(i, .ColIndex("记录状态")))
            Case 2
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
            Case 3
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbBlue
            Case Else
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
            End Select
        Next
        If rsTemp.RecordCount <> 0 Then
            .Rows = .Rows + 1: i = .Rows - 1
            .TextMatrix(i, .ColIndex("单据号")) = "合计"
            .TextMatrix(i, .ColIndex("应收金额")) = Format(dbl应收, "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .TextMatrix(i, .ColIndex("实收金额")) = Format(dbl实收, "######" & gstrDec & ";-#####" & gstrDec & "; ;")
            .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
    '      2.Cell的GridLine从上下左右向内都是从第1根线开始
    '      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)
            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '擦除一并给药相关行列的边线及内容
            lngLeft = vsAdvice.ColIndex("期效"): lngRight = vsAdvice.ColIndex("开始时间")
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = vsAdvice.ColIndex("天数"): lngRight = vsAdvice.ColIndex("用法")
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = vsAdvice.ColIndex("皮试"): lngRight = vsAdvice.ColIndex("皮试")
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub
Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, .ColIndex("诊疗类别")) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, .ColIndex("诊疗类别"))) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function


Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, strPrompt As String
    strPrompt = ""
    With vsAdvice
            lngRow = vsAdvice.MouseRow
            If Not (Button = 0 And lngRow > 0) Then Exit Sub
            Select Case .MouseCol
            Case .ColIndex("内容")
            Case .ColIndex("医嘱图标")
                If Val(.TextMatrix(lngRow, .ColIndex("诊疗项目ID"))) = 0 Then
                    strPrompt = "自由录入的医嘱"
                ElseIf Val(.TextMatrix(lngRow, .ColIndex("标志"))) = 1 Then
                    strPrompt = "紧急医嘱"
                ElseIf Val(.TextMatrix(lngRow, .ColIndex("标志"))) = 2 Then
                    strPrompt = "补录医嘱"
                End If
                 '如果有抗菌用药审核信息，优先显示
                If Val(.TextMatrix(lngRow, .ColIndex("医嘱状态"))) = 1 Then
                    Select Case Val(.TextMatrix(lngRow, .ColIndex("审核状态")))
                    Case 1
                        strPrompt = "抗菌用药待审核"
                    Case 2
                        strPrompt = "抗菌用药审核通过"
                    Case 3
                       strPrompt = "抗菌用药审核未通过:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, .ColIndex("医嘱ID"))))
                    End Select
                End If
            End Select
            If strPrompt <> "" Then
               Call zlCommFun.ShowTipInfo(vsAdvice.hWnd, strPrompt)
            Else
                Call zlCommFun.ShowTipInfo(vsAdvice.hWnd, "")
            End If
    End With
End Sub

Private Function GetKSSAuditQuestion(ByVal lng医嘱ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取抗菌用药审核未通过的反馈信息
    '编制:刘兴洪
    '日期:2012-05-31 14:40:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(操作说明,'无') as 操作说明 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=12 Order by 操作时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
    If Not rsTmp.EOF Then GetKSSAuditQuestion = rsTmp!操作说明
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, Me.Caption, "费用审核列头信息"
End Sub

Private Sub vsFeeList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsFeeList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long

    With vsFeeList
        If Col > .FixedCols - 1 Then Done = True: Exit Sub
    '擦除固定列中的表格线
    SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)
    '仅左边表格线
    vRect.Left = Left
    vRect.Top = Top
    vRect.Right = Left + 1
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '仅上边表格线
    vRect.Left = Left
    vRect.Top = Top
    vRect.Right = Right
    vRect.Bottom = Top + 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '仅下边表格线
    vRect.Left = Left
    vRect.Top = Bottom - 1
    vRect.Right = Right
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    '仅右边表格线
    vRect.Left = Right - 1
    vRect.Top = Top
    vRect.Right = Right
    vRect.Bottom = Bottom
    If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
    If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

    End With
    Done = True
End Sub
Private Function SaveData(ByVal lng医嘱ID As Long, ByVal bln审核 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:费用审核
    '编制:刘兴洪
    '日期:2012-05-31 17:58:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    'Zl_病人医嘱记录_费用审核
    strSQL = "Zl_病人医嘱记录_费用审核("
    'Id_In           病人医嘱记录.Id%Type,
    strSQL = strSQL & "" & lng医嘱ID & ","
    '是否费用审核_In 病人医嘱记录.是否费用审核%Type
    strSQL = strSQL & "" & IIf(bln审核, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mblnChangeData = True
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

