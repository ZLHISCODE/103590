VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBillIn 
   Caption         =   "票据入库管理"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9210
   Icon            =   "frmBillIn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPage 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   3600
      ScaleHeight     =   1710
      ScaleWidth      =   3225
      TabIndex        =   12
      Top             =   2940
      Width           =   3225
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2850
         Left            =   -1050
         TabIndex        =   13
         Top             =   315
         Width           =   3615
         _Version        =   589884
         _ExtentX        =   6376
         _ExtentY        =   5027
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDrawBill 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   5880
      ScaleHeight     =   2250
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   4155
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsDrawBill 
         Height          =   2145
         Left            =   0
         TabIndex        =   10
         Top             =   420
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":0442
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picColDrawBill 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   11
            Top             =   45
            Width           =   210
            Begin VB.Image imgColDrawBill 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":05F2
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picDamnifyBill 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   2025
      ScaleHeight     =   2250
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   4740
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsDamnifyBill 
         Height          =   2145
         Left            =   0
         TabIndex        =   7
         Top             =   255
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":0B40
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picColDamnifyBill 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   8
            Top             =   45
            Width           =   210
            Begin VB.Image imgColDamnifyBill 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":0C44
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picBillIn 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   1665
      ScaleHeight     =   2250
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   240
      Width           =   6795
      Begin VSFlex8Ctl.VSFlexGrid vsBillIn 
         Height          =   1185
         Left            =   0
         TabIndex        =   4
         Top             =   855
         Width           =   5550
         _cx             =   9790
         _cy             =   2090
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":1192
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgIn 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgColIn 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":1335
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   135
      ScaleHeight     =   1665
      ScaleWidth      =   2340
      TabIndex        =   1
      Top             =   3630
      Width           =   2340
      Begin MSComctlLib.ListView lvwMain 
         Height          =   3345
         Left            =   -15
         TabIndex        =   2
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   5900
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "票据种类"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   635
      SimpleText      =   $"frmBillIn.frx":1883
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillIn.frx":18CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":215E
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":25B0
            Key             =   "C2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":28CA
            Key             =   "C3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":2BE4
            Key             =   "C5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":2EFE
            Key             =   "C1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":3218
            Key             =   "C4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":3532
            Key             =   "C7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":384C
            Key             =   "C6"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   390
      Top             =   1605
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmBillIn.frx":3B66
      Left            =   930
      Top             =   1815
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBillIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************************************************************
'功能:票据入库管理
'编制:刘兴洪
'日期:2010-11-15 11:38:31
'说明:
'     33725
'********************************************************************************************************************************************
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrCmb As CommandBarComboBox

Private WithEvents mfrmFilter As frmBillInFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mlngModul As Long, mstrPrivs As String
Private mblnFirst As Boolean  '第一次加载窗体
Private mstrKey As String    '上一次的记录
Private mArrFilter As Variant   '条件
Private mobjPaneList As Pane
Private mPanSearch As Pane
Private mstr票据 As String   '上一次的票据种类
Private mstr票据长度  As String  '票据长度

Private Enum mPgIndex
    Pg_报损情况 = 250101
    Pg_领用记录 = 250102
End Enum
Private Enum mPaneID
    Pane_List = 1    '搜索条件
    Pane_Search = 2    '入库列表
    Pane_BillIn = 3  '页面
    Pane_BillDetails = 4    '详细列表
End Enum
Private mblnItem As Boolean  '为真表示单击到ListView某一项上
Private mintColumn As Integer '
Private mblnDateMoved As Boolean '当前时间范围是否在转出之前
 
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitFace = False Then Unload Me: Exit Sub
    Call LoadInDataToRpt
End Sub
Private Sub Form_Load()
    mblnFirst = True: mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitPanel: Call InitPage: Call InitVsGrid
     Set mArrFilter = mfrmFilter.GetFilterCon
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.编号, UserInfo.姓名)
    End If
    Err.Clear: On Error GoTo 0
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem     '增加入库
        Call BillInAdd
    Case conMenu_Edit_Modify     '修改入库
        Call BillInModify
    Case conMenu_Edit_Delete      '删除入库
        Call DeleteBillIn
    Case conMenu_Edit_DamnifyAdd      '增加报损
        Call ExcuteDamnifyFun(1)
    Case conMenu_Edit_DamnifyDelete      '删除报损
        Call ExcuteDamnifyFun(2)
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call LoadInDataToRpt
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
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
        Control.Enabled = zlIsHaveData
        
    Case conMenu_Edit_NewItem     '增加入库
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增加票据")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify     '修改入库
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改票据")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_Delete      '删除入库
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除票据")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_DamnifyAdd      '增加报损
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增加报损")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_DamnifyDelete      '删除报损
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除报损")
        If tbPage.Selected.Tag = mPgIndex.Pg_报损情况 Then
            With vsDamnifyBill
                If .Row > 0 Then
                    Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
                Else
                    Control.Enabled = Control.Visible And False
                End If
            End With
        Else
                Control.Enabled = Control.Visible And False
        End If
    Case conMenu_View_Refresh   '刷新
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
           ' Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1504" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '参数调用
        Case conMenu_Edit_UserType  '使用类别选择
            Call SetDefaultUserType: Call LoadInDataToRpt
        Case Else   '其他操作功能调用
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, int票种 As gBillType
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_Edit_UserType
        If lvwMain.SelectedItem Is Nothing Then
             int票种 = 0
         Else
             int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡,6-消费卡
         End If
        Select Case int票种
        Case gBillType.收费收据, gBillType.结帐收据
            Control.Visible = True
        Case gBillType.预交收据
            Control.Visible = True
        Case gBillType.就诊卡, gBillType.消费卡
            Control.Visible = True
        Case Else
            Control.Visible = False
        End Select
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub
Private Sub Form_Initialize()
  Call InitCommonControls
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    Err = 0: On Error Resume Next
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    Set mfrmFilter = Nothing
    
   SaveWinState Me, App.ProductName
   zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "入库列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
   zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "领用列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
   zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "报损列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
   zlSaveDockPanceToReg Me, dkpMan, "区域"
   Set mcbrCmb = Nothing
   mstr票据 = ""
   
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = False
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub imgColDamnifyBill_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picColDamnifyBill.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColDamnifyBill.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsDamnifyBill, lngLeft, lngTop, imgColDamnifyBill.Height)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Caption, "报损列表", True
End Sub

Private Sub mfrmFilter_WindowsHeight(lngHeght As Long)
    Dim lngHeightY As Long
    lngHeightY = lngHeght / Screen.TwipsPerPixelY
    If dkpMan.FindPane(mPaneID.Pane_Search) Is Nothing Then Exit Sub
    dkpMan.FindPane(mPaneID.Pane_Search).MinTrackSize.Height = lngHeightY
    dkpMan.FindPane(mPaneID.Pane_Search).MaxTrackSize.Height = lngHeightY
    dkpMan.RecalcLayout
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '重新加载数据
    Call LoadInDataToRpt
End Sub

Private Sub picColDamnifyBill_Click()
    Call imgColDamnifyBill_Click
End Sub

 
Private Sub imgColDrawBill_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picColDrawBill.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picColDrawBill.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsDrawBill, lngLeft, lngTop, imgColDrawBill.Height)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Caption, "领用列表", True
End Sub

Private Sub picColDrawBill_Click()
    Call imgColDrawBill_Click
End Sub

Private Sub imgColIn_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgIn.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgIn.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsBillIn, lngLeft, lngTop, imgColIn.Height)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Caption, "入库列表", True
End Sub

Private Sub picImgIn_Click()
    Call imgColIn_Click
End Sub

Private Sub lvwMain_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwMain.Drag 0
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstr票据 = Item.Key Then Exit Sub
    mstr票据 = Item.Key
    
    '调整列标题显示名称
    Call UpdateGridColumnCaption(CurrentIsBill(Val(Mid(Item.Key, 2))))
    
    Call LoadCombox(mcbrCmb)
    Call LoadInDataToRpt '重新加载数据
End Sub

Private Sub UpdateGridColumnCaption(ByVal blnBill As Boolean)
    '更新显示列名
    On Error GoTo ErrHandler
    With vsBillIn
        .TextMatrix(0, .ColIndex("开始票号")) = IIf(blnBill, "开始票号", "开始卡号")
        .TextMatrix(0, .ColIndex("终止票号")) = IIf(blnBill, "终止票号", "终止卡号")
        .TextMatrix(0, .ColIndex("票据张数")) = IIf(blnBill, "票据张数", "卡片张数")
    End With
    With vsDamnifyBill
        .TextMatrix(0, .ColIndex("开始票号")) = IIf(blnBill, "开始票号", "开始卡号")
        .TextMatrix(0, .ColIndex("终止票号")) = IIf(blnBill, "终止票号", "终止卡号")
    End With
    With vsDrawBill
        .TextMatrix(0, .ColIndex("开始票号")) = IIf(blnBill, "开始票号", "开始卡号")
        .TextMatrix(0, .ColIndex("终止票号")) = IIf(blnBill, "终止票号", "终止卡号")
        .TextMatrix(0, .ColIndex("当前号码")) = IIf(blnBill, "当前号码", "当前卡号")
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwMain.HitTest(x, y) Is Nothing Then Exit Sub
        lvwMain.Drag 1
    End If
End Sub


Private Sub picBillIn_Resize()
    Err = 0: On Error Resume Next
    With picBillIn
        vsBillIn.Left = .ScaleLeft
        vsBillIn.Top = .ScaleTop
        vsBillIn.Width = .ScaleWidth
        vsBillIn.Height = .ScaleHeight
    End With
End Sub
Private Sub picDamnifyBill_Resize()
    Err = 0: On Error Resume Next
    With picDamnifyBill
        vsDamnifyBill.Left = .ScaleLeft
        vsDamnifyBill.Top = .ScaleTop
        vsDamnifyBill.Width = .ScaleWidth
        vsDamnifyBill.Height = .ScaleHeight
    End With
End Sub
Private Sub picDrawBill_Resize()
    Err = 0: On Error Resume Next
    With picDrawBill
        vsDrawBill.Left = .ScaleLeft
        vsDrawBill.Top = .ScaleTop
        vsDrawBill.Width = .ScaleWidth
        vsDrawBill.Height = .ScaleHeight
    End With
End Sub

 
Private Sub picPage_Resize()
    Err = 0: On Error Resume Next
    With picPage
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picType_Resize()
    Err = 0: On Error Resume Next
    With picType
        lvwMain.Left = .ScaleLeft
        lvwMain.Top = .ScaleTop
        lvwMain.Width = .ScaleWidth
        lvwMain.Height = .ScaleHeight
    End With
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-15 11:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加入库(&N)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改入库(&M)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除入库(&D)"):
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyAdd, "增加报损(&A)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyDelete, "删除报损(&R)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("N"), conMenu_Edit_DamnifyAdd
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加入库"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改入库")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除入库")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyAdd, "增加报损"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyDelete, "删除报损")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        Set mcbrCmb = .Add(xtpControlComboBox, conMenu_Edit_UserType, "使用类别")
        mcbrCmb.Flags = xtpFlagRightAlign
        mcbrCmb.HideFlags = xtpNoHide
        With mcbrCmb
            .Clear
            .Width = (TextWidth("使用类别") * 4) / Screen.TwipsPerPixelX
        End With
         mcbrCmb.Style = xtpComboLabel
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
            mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function LoadCombox(ByVal objCombox As CommandBarComboBox) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载Combox数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-27 10:22:29
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As gBillType, strSQL As String, rsTemp As ADODB.Recordset
    Dim str类别 As String
    
    On Error GoTo errHandle
    
    If objCombox Is Nothing Then Exit Function
    If lvwMain.SelectedItem Is Nothing Then
        int票种 = 0
    Else
        int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡,6-消费卡
        str类别 = lvwMain.SelectedItem.Tag
    End If
    
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        strSQL = "Select 编码,名称,简码,缺省标志 From 票据使用类别 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            .AddItem "所有类别"
            If str类别 = "所有类别" Then .ListIndex = .ListCount
            .AddItem " "
            If str类别 = " " Then .ListIndex = .ListCount
            
            .ItemData(.ListCount) = -1
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!名称)
                .ItemData(.ListCount) = 1
                If Val(NVL(rsTemp!缺省标志)) = 1 And .ListIndex <= 0 Then .ListIndex = .ListCount
                If str类别 = NVL(rsTemp!名称) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "使用类别"
   Case gBillType.预交收据
        With objCombox
            .Clear
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 _
                And InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem "所有预交"
                .ItemData(.ListCount) = 0
                If Val(str类别) = 0 Then .ListIndex = 0
            End If
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 Then
                .AddItem "门诊预交": .ItemData(.ListCount) = 2
                If Val(str类别) = 2 Then .ListIndex = .ListCount
            End If
            If InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem "住院预交": .ItemData(.ListCount) = 3
                If Val(str类别) = 3 Then .ListIndex = .ListCount
            End If
            '58071
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 _
                And InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem " ": .ItemData(.ListCount) = 1
                If Val(str类别) = 1 Then .ListIndex = .ListCount
            End If
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "使用类别"
    Case gBillType.就诊卡
        strSQL = "Select ID,编码,名称,缺省标志 From 医疗卡类别 where nvl(是否启用,0) >=1 Order by 编码 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
                .ItemData(.ListCount) = Val(NVL(rsTemp!ID))
                If Val(NVL(rsTemp!缺省标志)) = 1 And .ListIndex <= 0 Then .ListIndex = .ListCount
                If Val(str类别) = Val(NVL(rsTemp!ID)) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = .ListCount
        End With
        mcbrCmb.Caption = "卡类别"
    Case gBillType.消费卡
        strSQL = "Select 编号,名称 From 消费卡类别目录 where nvl(启用,0) >=1 Order by 编号 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!编号) & "-" & NVL(rsTemp!名称)
                .ItemData(.ListCount) = Val(NVL(rsTemp!编号))
                If Val(str类别) = Val(NVL(rsTemp!编号)) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "卡类别"
    Case Else
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSearchWindowsHeight()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置查找窗体的高度
    '编制:刘兴洪
    '日期:2010-11-15 14:50:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngSearchHeight As Long
    lngSearchHeight = mfrmFilter.Height / Screen.TwipsPerPixelY
    mPanSearch.MinTrackSize.Height = lngSearchHeight: mPanSearch.MaxTrackSize.Height = lngSearchHeight
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2010-11-15 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, lngWidth As Long
    Dim lngHeight As Long
    
    If mfrmFilter Is Nothing Then
        Set mfrmFilter = New frmBillInFilter
        Load mfrmFilter
    End If
    lngWidth = lvwMain.Width / Screen.TwipsPerPixelX
    With dkpMan
        Set mobjPaneList = .CreatePane(mPaneID.Pane_List, lngWidth, 400, DockLeftOf, Nothing)
        mobjPaneList.Title = "票种信息": mobjPaneList.Options = PaneNoCloseable
        mobjPaneList.MinTrackSize.Width = lngWidth: mobjPaneList.MaxTrackSize.Width = lngWidth
        mobjPaneList.Handle = picType.hWnd
        mobjPaneList.Tag = mPaneID.Pane_List
        
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockRightOf, mobjPaneList)
        mPanSearch.Title = "条件过滤"
        mPanSearch.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        mPanSearch.Handle = mfrmFilter.hWnd
        mPanSearch.Tag = mPaneID.Pane_Search
        
        Set objPane = .CreatePane(mPaneID.Pane_BillIn, 400, 400, DockBottomOf, mPanSearch)
        objPane.Title = "票据入库信息"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBillIn.hWnd
        objPane.Tag = mPaneID.Pane_BillIn
        Set objPane = .CreatePane(mPaneID.Pane_BillDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPage.hWnd
        objPane.Tag = mPaneID.Pane_BillDetails
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
         
    End With
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Function
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPaneID.Pane_Search    '搜索条件窗体
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_BillIn    '入库信息
        Item.Handle = picBillIn.hWnd
    Case mPaneID.Pane_BillDetails  '详细列表
        Item.Handle = picPage.hWnd
    Case mPaneID.Pane_List
        Item.Handle = picType.hWnd
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2010-11-15 13:56:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_报损情况, "报损情况", picDamnifyBill.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_报损情况
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_领用记录, "领用信息", picDrawBill.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_领用记录
     With tbPage
        tbPage.Item(i).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2010-11-15 13:58:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBillIn
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("入库批次")) = "1|0"
        .ColData(.ColIndex("开始票号")) = "1|1"
        .ColData(.ColIndex("终止票号")) = "1|0"
    End With
    With vsDamnifyBill
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("报损日期")) = "1|0"
        .ColData(.ColIndex("开始票号")) = "1|1"
        .ColData(.ColIndex("终止票号")) = "1|0"
    End With
    With vsDrawBill
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("开始票号")) = "1|1"
        .ColData(.ColIndex("终止票号")) = "1|0"
    End With
End Sub
Private Function LoadInDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载入库数据给网格
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-15 14:54:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, int票种 As gBillType, lngPreID As Long
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim str类别 As String
    Err = 0: On Error GoTo ErrHand:
    
    If lvwMain.SelectedItem Is Nothing Then
        int票种 = 0
    Else
        int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡,6-消费卡
    End If
    strWhere = ""
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        str类别 = mcbrCmb.Text
        If str类别 = " " Then
              strWhere = strWhere & " And nvl(A.使用类别,'LXH')=[6]"
              str类别 = "LXH"
        ElseIf str类别 <> "所有类别" Then
              strWhere = strWhere & " And nvl(A.使用类别,'LXH')=[6]"
        End If
    Case gBillType.预交收据
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str类别 = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str类别) <> 0 Then
            str类别 = Val(str类别) - 1
            strWhere = strWhere & " And nvl(A.使用类别,'0')=[6]"
        End If
    Case gBillType.就诊卡
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str类别 = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str类别) <> 0 Then
            str类别 = Val(str类别)
            strWhere = strWhere & " And nvl(A.使用类别,'0')=[6]"
        End If
    Case gBillType.消费卡
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str类别 = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str类别) <> 0 Then
            str类别 = Val(str类别)
            strWhere = strWhere & " And nvl(A.接口编号,0)=[6]"
        End If
    Case Else
    End Select
    
    
    If mArrFilter("入库时间")(0) <> "1901-01-01" And mArrFilter("入库时间")(0) <> "1901-01-01" Then
        strWhere = strWhere & " And (A.登记时间 Between [1] And [2] )"
    End If
    
    If mArrFilter("登记人") <> "" Then
        strWhere = strWhere & " And A.登记人=[3]"
    End If
    If Val(mArrFilter("仅显示有库存发票")) = 1 Then
        If int票种 = gBillType.消费卡 Then
            strWhere = strWhere & " And nvl(A.剩余数量,0)>0 and A.是否存在卡 =1"
        Else
            strWhere = strWhere & " And nvl(A.剩余数量,0)>0 and A.有无票据 =1"
        End If
    End If
    If int票种 <> gBillType.消费卡 Then
        strWhere = strWhere & " And nvl(A.票种,0)=[4]"
    End If
    If InStr(1, mstrPrivs, ";允许操作他人登记票据;") = 0 Then
        strWhere = strWhere & " And A.登记人 =[5]"
    End If
    If int票种 = gBillType.消费卡 Then
        gstrSQL = "" & _
        "  Select A.Id,nvl(B.名称,'就诊卡') as 使用类别, A.前缀文本, A.开始卡号 As 开始号码,A.终止卡号 As 终止号码, A.入库数量," & _
        "       A.剩余数量, A.备注, A.登记人, A.登记时间, A.批次  " & vbCrLf & _
        "  From 消费卡入库记录 A,消费卡类别目录 B" & vbCrLf & _
        "  Where nvl(A.接口编号,0)=B.编号(+) And " & Mid(strWhere, 5)
    ElseIf int票种 = gBillType.就诊卡 Then
        gstrSQL = "" & _
        "  Select A.Id,nvl(B.名称,'就诊卡') as 使用类别, A.前缀文本, A.开始号码,A.终止号码, A.入库数量," & _
        "       A.剩余数量, A.备注, A.登记人, A.登记时间, A.批次  " & vbCrLf & _
        "  From 票据入库记录 A,医疗卡类别 B" & vbCrLf & _
        "  where   to_number(nvl(A.使用类别,'0'))=B.ID(+) And " & Mid(strWhere, 5)
    ElseIf int票种 = gBillType.预交收据 Then
        gstrSQL = "" & _
        "  Select A.Id,decode(nvl(A.使用类别,0),0,'','1','门诊','住院') as 使用类别, A.票种, A.前缀文本, A.开始号码,A.终止号码, A.入库数量," & _
        "       A.剩余数量, A.备注, A.登记人, A.登记时间, A.批次  " & vbCrLf & _
        "  From 票据入库记录 A" & vbCrLf & _
        "  where   " & Mid(strWhere, 5)
    Else
        gstrSQL = "" & _
        "  Select A.Id,A.使用类别, A.前缀文本, A.开始号码,A.终止号码, A.入库数量," & _
        "       A.剩余数量, A.备注, A.登记人, A.登记时间, A.批次  " & vbCrLf & _
        "  From 票据入库记录 A" & vbCrLf & _
        "  where   " & Mid(strWhere, 5)
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("入库时间")(0)), CDate(mArrFilter("入库时间")(1)), _
        CStr(mArrFilter("登记人")), int票种, UserInfo.姓名, str类别)
    
    With vsBillIn
        .Clear 1
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .RowData(1) = 0
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("使用类别")) = NVL(rsTemp!使用类别)
            .TextMatrix(lngRow, .ColIndex("入库批次")) = NVL(rsTemp!批次)
            .TextMatrix(lngRow, .ColIndex("前缀文本")) = NVL(rsTemp!前缀文本)
            .TextMatrix(lngRow, .ColIndex("开始票号")) = NVL(rsTemp!开始号码)
            .TextMatrix(lngRow, .ColIndex("终止票号")) = NVL(rsTemp!终止号码)
            .TextMatrix(lngRow, .ColIndex("票据张数")) = NVL(rsTemp!入库数量)
            .TextMatrix(lngRow, .ColIndex("剩余张数")) = NVL(rsTemp!剩余数量)
            .TextMatrix(lngRow, .ColIndex("登记人")) = NVL(rsTemp!登记人)
            .TextMatrix(lngRow, .ColIndex("登记时间")) = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("备注")) = NVL(rsTemp!备注)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .FixedCols = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBillIn, Me.Name, "入库列表", True, True
    Call vsBillIn_AfterRowColChange(-1, 0, vsBillIn.Row, 0)
    vsBillIn.ColHidden(vsBillIn.ColIndex("标志")) = False
    vsBillIn.ColWidth(vsBillIn.ColIndex("标志")) = 300
    LoadInDataToRpt = True
    Exit Function
ErrHand:
    vsBillIn.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub LoadDamnifyData(ByVal lng入库批次 As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载报损数据
    '编制:刘兴洪
    '日期:2010-11-15 17:42:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngPreID As Long, lngRow As Long
    Dim int票种 As gBillType
    
    On Error GoTo ErrHand:
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If int票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select ID, 入库id, 开始卡号 As 开始号码, 终止卡号 As 终止号码, 数量, 报损原因, 报损人, 报损时间" & vbNewLine & _
            "From 消费卡报损记录" & vbNewLine & _
            "Where 入库id = [1]"
    Else
        gstrSQL = _
            "Select ID, 入库id, 开始号码, 终止号码, 数量, 报损原因, 报损人, 报损时间" & vbNewLine & _
            "From 票据报损记录" & vbNewLine & _
            "Where 入库id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng入库批次)
    
    With vsDamnifyBill
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .RowData(1) = ""
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("报损时间")) = Format(rsTemp!报损时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("开始票号")) = NVL(rsTemp!开始号码)
            .TextMatrix(lngRow, .ColIndex("终止票号")) = NVL(rsTemp!终止号码)
            .TextMatrix(lngRow, .ColIndex("报损数量")) = NVL(rsTemp!数量)
            .TextMatrix(lngRow, .ColIndex("报损人")) = NVL(rsTemp!报损人)
            .TextMatrix(lngRow, .ColIndex("报损原因")) = NVL(rsTemp!报损原因)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDamnifyBill, Me.Name, "报损列表", True, True
    Exit Sub
ErrHand:
    vsDamnifyBill.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub LoadDrawData(ByVal lng入库批次 As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:lng入库批次-入库ID
    '编制:刘兴洪
    '日期:2010-11-15 17:31:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngPreID As Long, lngRow As Long
    Dim int票种 As gBillType
    
    On Error GoTo ErrHand:
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If int票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select ID,领用人,前缀文本,开始卡号 As 开始号码,终止卡号 As 终止号码," & _
            "       decode(使用方式,1,'自用','共用') as 使用方式,to_char(登记时间,'yyyy-MM-dd') as 登记时间," & _
            "       登记人,当前卡号 As 当前号码,剩余数量,批次,核对人 " & _
            "From 消费卡领用记录 " & _
            "Where 批次=[1]"
    Else
        gstrSQL = "" & _
        "   Select ID,领用人,前缀文本,开始号码,终止号码," & _
        "           decode(使用方式,1,'自用','共用') as 使用方式,to_char(登记时间,'yyyy-MM-dd') as 登记时间," & _
        "           登记人,当前号码,剩余数量,批次,核对人 " & _
        "   From 票据领用记录 " & _
        "   Where 批次=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng入库批次)
    
   With vsDrawBill
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            
            .TextMatrix(lngRow, .ColIndex("开始票号")) = NVL(rsTemp!开始号码)
            .TextMatrix(lngRow, .ColIndex("终止票号")) = NVL(rsTemp!终止号码)
            .TextMatrix(lngRow, .ColIndex("领用人")) = NVL(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("当前号码")) = NVL(rsTemp!当前号码)
            .TextMatrix(lngRow, .ColIndex("剩余数量")) = NVL(rsTemp!剩余数量)
            .TextMatrix(lngRow, .ColIndex("入库批次")) = NVL(rsTemp!批次)
            .TextMatrix(lngRow, .ColIndex("使用方式")) = NVL(rsTemp!使用方式)
            .TextMatrix(lngRow, .ColIndex("登记人")) = NVL(rsTemp!登记人)
            .TextMatrix(lngRow, .ColIndex("登记时间")) = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("核对人")) = NVL(rsTemp!核对人)
            .TextMatrix(lngRow, .ColIndex("前缀文本")) = NVL(rsTemp!前缀文本)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDrawBill, Me.Name, "领用列表", True, True
    Exit Sub
ErrHand:
    vsDrawBill.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function InitFace() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面数据
    '编制:刘兴洪
    '日期:2010-11-15 15:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim i As Integer, strTmp As String
    
    mstr票据长度 = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    If zlStr.IsHavePrivs(mstrPrivs, "收费收据") Then
        strTmp = strTmp & "|" & "收费收据,1"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "预交收据") _
        And (zlStr.IsHavePrivs(mstrPrivs, "预交门诊票据") _
            Or zlStr.IsHavePrivs(mstrPrivs, "预交住院票据")) Then
        strTmp = strTmp & "|" & "预交收据,2"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "结帐收据") Then
        strTmp = strTmp & "|" & "结帐收据,3"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "挂号收据") Then
        strTmp = strTmp & "|" & "挂号收据,4"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "医疗卡") Then
        strTmp = strTmp & "|" & "医疗卡,5"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "消费卡") Then
        strTmp = strTmp & "|" & "消费卡,6"
    End If
    If strTmp = "" Then
        MsgBox "你没有操作任何票据的权限!", vbInformation, App.ProductName
        Exit Function
    Else
        strTmp = Mid(strTmp, 2)
    End If
    arrTemp1 = Split(strTmp, "|")
    For i = 0 To UBound(arrTemp1)
        arrTemp2 = Split(arrTemp1(i), ",")
        lvwMain.ListItems.Add , "C" & arrTemp2(1), arrTemp2(0), "C" & arrTemp2(1)
    Next
    lvwMain.ListItems(1).Selected = True
    Call LoadCombox(mcbrCmb)
    Call SetDefaultUserType
    InitFace = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetDefaultUserType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的使用类别
    '编制:刘兴洪
    '日期:2011-04-27 14:23:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As gBillType
    
    If mcbrCmb Is Nothing Then Exit Sub
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡,6-消费卡
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        lvwMain.SelectedItem.Tag = mcbrCmb.Text
    Case gBillType.预交收据, gBillType.就诊卡, gBillType.消费卡
        If mcbrCmb.ListIndex <= 0 Then
            lvwMain.SelectedItem.Tag = ""
        Else
            lvwMain.SelectedItem.Tag = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        End If
    Case Else
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-15 15:34:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String, bln明细 As Boolean
    bln明细 = False
    If Me.ActiveControl Is vsDrawBill Then
        Set vsBill = vsDrawBill: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "领用明细": bln明细 = True
    ElseIf Me.ActiveControl Is vsDamnifyBill Then
        Set vsBill = vsDamnifyBill: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "报损明细": bln明细 = True
    Else
        Set vsBill = vsBillIn: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "入库清册"
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    If bln明细 Then
        With vsBillIn
            objRow.Add "入库批次:" & Val(.TextMatrix(.Row, .ColIndex("入库批次")))
            objRow.Add "票据范围:" & Trim(.TextMatrix(.Row, .ColIndex("开始票号"))) & "至" & Trim(.TextMatrix(.Row, .ColIndex("终止票号")))
            objRow.Add "登记人:" & Trim(.TextMatrix(.Row, .ColIndex("登记人")))
        End With
    Else
        If CStr(mArrFilter("入库时间")(0)) <> "1901-01-01" Then
            objRow.Add "入库时间：" & CStr(mArrFilter("入库时间")(0)) & "至" & CStr(mArrFilter("入库时间")(1))
        End If
        If Val(mArrFilter("仅显示有库存发票")) = 1 Then
            objRow.Add "仅显示有库存发票"
        End If
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    '由于打印控件不能识别列隐藏属性
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    '恢复
    With vsBill
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub ExcuteDamnifyFun(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:报损功能
    ' 入参:bytFun-1-新增;2-删除
    '编制:刘兴洪
    '日期:2010-11-15 15:36:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, lngRow As Long
    Dim rsTemp As ADODB.Recordset
    Dim lng入库ID As Long, int票种 As gBillType
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    With vsBillIn
        lngID = Val(.RowData(.Row))
        lng入库ID = lngID
        If lngID < 0 Then Exit Sub
    End With
    If bytFun = 1 Then '新增
        If frmBillInDamnifyEdit.zlBillEdit(Me, EdS_报损, mstrPrivs, mlngModul, int票种, lngID) = False Then Exit Sub
        gstrSQL = "Select 剩余数量 From 票据入库记录 where ID=[1]"
        If int票种 = gBillType.消费卡 Then
            gstrSQL = "Select 剩余数量 From 消费卡入库记录 where ID=[1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng入库ID)
        If Not rsTemp.EOF Then
            vsBillIn.TextMatrix(vsBillIn.Row, vsBillIn.ColIndex("剩余张数")) = NVL(rsTemp!剩余数量)
        End If
        Call LoadDamnifyData(lngID)
        Exit Sub
    End If
    
    With vsDamnifyBill
        lngID = Val(.RowData(.Row))
        If lngID = 0 Then Exit Sub
        If Not (.TextMatrix(.Row, .ColIndex("终止票号")) = "" Or .TextMatrix(.Row, .ColIndex("开始票号")) = .TextMatrix(.Row, .ColIndex("终止票号"))) Then
            strTemp = "-" & .TextMatrix(.Row, .ColIndex("终止票号"))
        End If
        If MsgBox("你确认要删除报损号为“" & .TextMatrix(.Row, .ColIndex("开始票号")) & strTemp & "”的" & _
             "报损记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End With
    Err = 0: On Error GoTo errHandle:
    
    Me.MousePointer = 11
    If int票种 = gBillType.消费卡 Then
        'Zl_消费卡报损记录_Delete(Id_In In 消费卡报损记录.ID%Type) Is
        gstrSQL = "Zl_消费卡报损记录_Delete(" & lngID & ")"
    Else
        'Zl_票据报损记录_Delete(Id_In In 票据报损记录.ID%Type) Is
        gstrSQL = "Zl_票据报损记录_Delete(" & lngID & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Me.MousePointer = 0
    With vsDamnifyBill
        lngRow = .Row
        If .Rows - 1 > .Row Then
            .Row = .Row + 1
            .RemoveItem lngRow
        ElseIf .Row = 1 Then
            .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            .RowData(.Row) = ""
        Else
              .Row = .Row - 1
              .RemoveItem lngRow
        End If
    End With
    gstrSQL = "Select 剩余数量 From 票据入库记录 where ID=[1]"
    If int票种 = gBillType.消费卡 Then
        gstrSQL = "Select 剩余数量 From 消费卡入库记录 where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng入库ID)
    If Not rsTemp.EOF Then
        vsBillIn.TextMatrix(vsBillIn.Row, vsBillIn.ColIndex("剩余张数")) = NVL(rsTemp!剩余数量)
    End If
    
    'If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub DeleteBillIn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除入库票据
    '编制:刘兴洪
    '日期:2010-11-15 16:14:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngID As Long
    Dim strExpended As String, int票种 As gBillType, blnTrans As Boolean
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    With vsBillIn
        If .Row <= 0 Then Exit Sub
        lngID = Val(.RowData(.Row))
        If lngID < 0 Then Exit Sub
        If MsgBox("你确认要删除" & lvwMain.SelectedItem.Text & "的批次为『" & .TextMatrix(.Row, .ColIndex("入库批次")) & "』;开始" & _
            IIf(CurrentIsBill(int票种), "号码", "卡号") & "为“" & .TextMatrix(.Row, .ColIndex("开始票号")) & "”的" & _
            "入库记录吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End With
    '102996:李南春,2016/11/23,医疗发票电子化管理
    If gblnBillPrint Then
        On Error Resume Next
        If gobjBillPrint.zlBillInCheckValied(3, int票种, "", lngID, "", "", strExpended) = False Then Exit Sub
    End If
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    If int票种 = gBillType.消费卡 Then
        'Zl_消费卡入库记录_Delete(Id_In In 消费卡入库记录.ID%Type) Is
        gstrSQL = "Zl_消费卡入库记录_Delete(" & lngID & ")"
    Else
        'Zl_票据入库记录_Delete(Id_In In 票据入库记录.ID%Type) Is
        gstrSQL = "Zl_票据入库记录_Delete(" & lngID & ")"
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If gblnBillPrint Then
        On Error Resume Next
        strExpended = ""
        If gobjBillPrint.zlBillIn(3, int票种, "", lngID, strExpended) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            Me.MousePointer = 0
            Exit Sub
        End If
        Err = 0: On Error GoTo errHandle
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = 0
    With vsBillIn
        lngRow = .Row
        If .Rows - 1 > .Row Then
            .Row = .Row + 1
            .RemoveItem lngRow
        ElseIf .Row = 1 Then
            .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            .RowData(.Row) = ""
        Else
              .Row = .Row - 1
              .RemoveItem lngRow
        End If
    End With
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
    Me.MousePointer = 0
End Sub
Private Sub BillInModify()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改入库票据
    '编制:刘兴洪
    '日期:2010-11-15 17:01:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID  As Long, strNO As String, rsTemp As ADODB.Recordset, lngLen As Long
    Dim int票种 As gBillType
    
    With vsBillIn
        If .Row <= 0 Then Exit Sub
        lngID = Val(.RowData(.Row))
        If lngID < 0 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("开始票号")))
    End With
    '102181:李南春,2016/11/10,医疗卡票据长度
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If Not (int票种 = gBillType.就诊卡 Or int票种 = gBillType.消费卡) Then
        lngLen = Val(Split(mstr票据长度, "|")(Mid(lvwMain.SelectedItem.Key, 2) - 1))
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "修改票据") = False Then
        If frmBillInEdit.zlBillEdit(Me, int票种, Ed_查看, mstrPrivs, mlngModul, lngID) = False Then Exit Sub
        Exit Sub
    End If
 
    If frmBillInEdit.zlBillEdit(Me, int票种, Ed_修改, mstrPrivs, mlngModul, lngID) = False Then Exit Sub
    '    If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadInDataToRpt
End Sub
Private Sub BillInAdd()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加入库票据
    '编制:刘兴洪
    '日期:2010-11-15 17:01:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As gBillType
    Dim str类别 As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If zlStr.IsHavePrivs(mstrPrivs, "增加票据") = False Then Exit Sub
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If int票种 = gBillType.收费收据 Or int票种 = gBillType.结帐收据 Then
        If Not mcbrCmb Is Nothing Then str类别 = Trim(mcbrCmb.Text)
    ElseIf int票种 = gBillType.预交收据 Or int票种 = gBillType.就诊卡 Or int票种 = gBillType.消费卡 Then
        If Not mcbrCmb Is Nothing Then
            If mcbrCmb.ListIndex > 0 Then
                str类别 = Trim(mcbrCmb.ItemData(mcbrCmb.ListIndex))
            End If
        End If
    End If
    
    If frmBillInEdit.zlBillEdit(Me, int票种, Ed_增加, mstrPrivs, mlngModul, 0, str类别) = False Then Exit Sub
    '    If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadInDataToRpt
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    '编制:刘兴洪
    '日期:2010-11-15 17:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str登记人 As String, str入库ID As String, int票据 As Integer
    
    With vsBillIn
        If .Row > 0 Then
            str入库ID = Val(.RowData(.Row))
            str登记人 = Trim(.TextMatrix(.Row, .ColIndex("登记人")))
        End If
    End With
    int票据 = Val(Mid(lvwMain.SelectedItem.Key, 2))
 
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
        "票种=" & int票据, "登记人=" & str登记人, "入库ID=" & str入库ID)
End Sub

Private Sub vsBillIn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngID As Long
    If OldRow = NewRow Then Exit Sub
    With vsBillIn
        lngID = Val(.RowData(NewRow))
        If lngID = 0 Then lngID = -1 '不能显示无入库流程的领用记录
        Call LoadDrawData(lngID)
        Call LoadDamnifyData(lngID)
    End With
End Sub

Private Sub vsBillIn_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillIn
        If .ColIndex("标志") = Col Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBillIn_DblClick()
    Call BillInModify
End Sub
Private Sub vsBillIn_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Call BillInModify
End Sub
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否存在数据
    '编制:刘兴洪
    '日期:2010-11-15 17:54:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is vsDrawBill Then
        With vsDrawBill
            If .Rows < 1 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    ElseIf Me.ActiveControl Is vsDamnifyBill Then
        With vsDrawBill
            If .Rows < 1 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    Else
        With vsBillIn
            If .Rows < 0 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    End If
End Function

Private Sub vsBillIn_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "入库列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBillIn_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "入库列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsDamnifyBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDamnifyBill
        If .ColIndex("标志") = Col Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsDrawBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "领用列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsDrawBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "领用列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsDamnifyBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "报损列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsDamnifyBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "报损列表", True, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
 
Private Sub vsDrawBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDrawBill
        If .ColIndex("标志") = Col Then Cancel = True: Exit Sub
    End With
End Sub
