VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathTableOut 
   BorderStyle     =   0  'None
   Caption         =   "门诊临床路径表"
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraPath 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   380
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox cboPath 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "路径名称"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblInDiag 
         BackColor       =   &H8000000E&
         Caption         =   "导入诊断："
         Height          =   255
         Left            =   10080
         TabIndex        =   11
         Top             =   120
         Width           =   4995
      End
      Begin VB.Label lblOutDate 
         BackColor       =   &H8000000E&
         Caption         =   "结束时间：3000-01-01 00:01"
         Height          =   255
         Left            =   7620
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblInDate 
         BackColor       =   &H8000000E&
         Caption         =   "导入时间：3000-01-01 00:00"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblInPep 
         BackColor       =   &H8000000E&
         Caption         =   "导入人：***"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
   End
   Begin VB.Frame fraline 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   5640
      Width           =   8175
   End
   Begin MSComctlLib.ImageList imgCharacter 
      Left            =   8280
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":0000
            Key             =   "已经执行"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":059A
            Key             =   "尚未执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":0B34
            Key             =   "取消执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":10CE
            Key             =   "部分执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":1668
            Key             =   "提前执行"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":1C02
            Key             =   "延后执行"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgMore 
         Height          =   225
         Left            =   0
         Picture         =   "frmPathTableOut.frx":219C
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":259D
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":26E4
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2833
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":29B5
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2B7B
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2FFE
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":3479
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":3899
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "双击查看路径项目定义"
      Top             =   2400
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3CBD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFlow 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "双击查看路径阶段定义"
      Top             =   390
      Width           =   8175
      _cx             =   14420
      _cy             =   3387
      Appearance      =   2
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1800
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3DF8
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   101
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPathPrint 
      Height          =   3105
      Index           =   0
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "双击查看路径项目定义"
      Top             =   -99999
      Visible         =   0   'False
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3E69
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   8880
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTableOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)     '要求查看报告
Public Event Activate()                                                         '自已激活时
Public Event RequestRefresh(ByVal lngPathState As Long)                         '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String)                             '要求更新主窗体状态栏文字
Private Const C_Exe = "√"                                                      '■
Private Const CON_SmallFontSize As Long = 9                                     '小字体
Private Const CON_BigFontSize As Long = 12                                      '大字体
Private Const CON_PathOutItemColor As Long = &HC0FFFF                           '路径外项目，浅黄色
Private Const CON_PathOutItemColorBlue As Long = &HFAEADA                       '暂存路径外项目,浅蓝色标识
Private Const C_UnExe = "□"

Private Enum EFixedRow
    R0阶段名 = 0
    R1天数 = 1
    R2日期 = 2
End Enum

Private Enum PatiType
    pt候诊 = 0
    pt就诊 = 1
    pt已诊 = 2
    pt转诊 = 3
    pt预约 = 4
    pt回诊 = 5
    pt排队叫号 = 6
End Enum

Private mfrmParent          As Object
Private mcbsMain            As Object
Private mobjPublicPACS      As Object

Private mPP                 As TYPE_PATH_Pati
Private mPati               As TYPE_Pati
Private mcolReason          As Collection

Private mbln启用执行环节    As Boolean                  '是否启用路径执行环节
Private mblnUnChange        As Boolean                  '不调用单元格变化事件，刷新单元格内容
Private mblnInOverScope     As Boolean                  '病人当前执行天数是否在标准治疗时间范围（允许结束路径）
Private mlngFontSize        As Long                     '界面字体大小
Private mlngPathCount As Long   '当次住院的路径数

Private Sub SetUnImport()
'功能：设置未导入时的状态和信息
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .ForeColorSel = vbBlack
        .TextMatrix(0, 0) = "  该病人未导入门诊临床路径。"
    End With
    Call ClearPathItem
End Sub

Private Sub SetImportFalse()
'功能：设置当病人导入门诊临床路径失败时的状态和信息
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .TextMatrix(0, 0) = "  该病人不符合路径导入条件。" & vbCrLf & "  原因：" & mPP.未导入原因
        .AutoSize 0
        .ForeColorSel = &HC0&
        If .Visible And .Enabled Then .SetFocus
    End With
    Call ClearPathItem
End Sub

Private Sub ClearPathItem(Optional blnImported As Boolean)
'功能：当病人没有可用的门诊临床路径时清除路径表项目
    With vsPath
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = 0
        If blnImported Then
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = vbCrLf & "  该病人还没有生成路径项目。"
            .Select 0, 0
            .CellAlignment = flexAlignLeftTop
        End If
    End With
End Sub

Private Sub cboPath_Click()
    If cboPath.ListIndex >= 0 Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态, False, , Val(cboPath.ItemData(cboPath.ListIndex)))
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If fraPath.Visible Then
        fraPath.Top = lngTop
        fraPath.Left = lngLeft
        fraPath.Width = lngRight - lngLeft
        lngTop = fraPath.Top + fraPath.Height
    End If
    vsFlow.Left = lngLeft
    vsFlow.Top = lngTop
    vsFlow.Width = lngRight - lngLeft
    vsFlow.Height = 1140

    If vsPath.FixedRows = 0 And vsPath.Rows = 0 Then  '没有导入路径时
        vsFlow.Height = Me.Height
        vsPath.Visible = False
        UCAdvice.Visible = False
        fraline.Visible = False
    Else
        If vsPath.Visible = False Then vsPath.Visible = True
        If UCAdvice.Visible = False Then UCAdvice.Visible = True
        If fraline.Visible = False Then fraline.Visible = True
        
        With vsPath
            .Top = lngTop + vsFlow.Height
            .Width = lngRight - lngLeft
            If lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30 > 0 Then
                .Height = lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30
            Else
                .Height = lngBottom - lngTop - vsFlow.Height
            End If

            If .FixedRows = 0 And .Rows = 1 Then             '没有生成项目
                .ColWidth(0) = .Width - 30
                .RowHeight(0) = .Height
            End If
            fraline.Top = .Top + .Height
            fraline.Width = .Width

            UCAdvice.Top = fraline.Top + fraline.Height
            UCAdvice.Width = .Width
        End With
    End If

    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If vsPath.Height + Y < 1000 Or vsPath.Height - Y < 500 Then Exit Sub
        If UCAdvice.Height + Y < 250 Or UCAdvice.Height - Y < 500 Then Exit Sub

        If fraMore.Visible Then fraMore.Visible = False

        fraline.Top = fraline.Top + Y
        vsPath.Height = vsPath.Height + Y
        UCAdvice.Top = UCAdvice.Top + Y
        UCAdvice.Height = UCAdvice.Height - Y
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Call InitCbsSubBar
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
End Sub

Private Sub LoadPathFlow()
'功能：根据病人导入的路径表加载路径基本信息和流程
    Dim strSql As String, i As Long, j As Long, lngCurCol As Long
    Dim rsTmp As ADODB.Recordset, lngDayMin As Long, lngDayMax As Long
    Dim lng理论天数 As Long
    Dim lng序号 As Long
    Dim str标准治疗时间 As String
    
    On Error GoTo errH
    
    With vsFlow
        .Clear
        .Rows = 1: .Cols = 1
        .ForeColorSel = vbBlack
        mblnInOverScope = False

        strSql = " Select a.ID,a.名称 阶段名,Decode(a.结束天数, Null, 0, 1) 多天,b.分类,b.名称 路径名,b.最新版本,c.标准治疗时间" & _
                 " From 门诊路径阶段 a,门诊路径目录 b,门诊路径版本 c " & _
                 " Where a.路径id = [1] And a.版本号 = [2] And a.路径id=b.id And a.父ID is null And b.id = c.路径id And a.版本号 = c.版本号 " & _
                 " Order by a.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID, mPP.版本号)
        str标准治疗时间 = NVL(rsTmp!标准治疗时间)
        If rsTmp.RecordCount > 0 Then
            .Rows = 1
            .Cols = rsTmp.RecordCount * 2                         '第一列为路径名，箭头为阶段数-1
            .Select 0, 0
            .RowHeight(0) = 1100

            '第一列显示路径名称
            .ColWidth(0) = 2800
            If mPP.病人路径状态 > 0 Then
                .TextMatrix(0, 0) = rsTmp!路径名 & ""

                If mPP.病人路径状态 = 3 Then
                    .Cell(flexcpForeColor, 0, 0) = vbRed
                End If
            Else
                .TextMatrix(0, 0) = rsTmp!路径名 & ""
            End If
            
            If mPP.当前天数 > 0 And mPP.病人路径状态 = 1 Then
            
                '获取标准治疗时间
                If InStr(str标准治疗时间, "-") > 0 Then
                    j = Split(str标准治疗时间, "-")(1)
                    lngDayMin = Val(Split(str标准治疗时间, "-")(0))
                    lngDayMax = j
                Else
                    j = Val(str标准治疗时间)                                '小于等于n天的情况
                    lngDayMin = 1
                    lngDayMax = j
                End If

                lng理论天数 = GetMustDayOut(mPP.病人路径ID, mPP.当前天数)

                i = Format(lng理论天数 / j * 100, "0")
                If i = 100 And lng理论天数 <> j Then
                    i = 99
                End If
                
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "进度：" & i & "%"

                If lng理论天数 > lngDayMax Then
                    mblnInOverScope = True
                Else
                    mblnInOverScope = Between(lng理论天数, lngDayMin, lngDayMax)
                End If
            End If
            
            If mPP.病人路径状态 > 0 Then
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "状态：" & IIf(mPP.病人路径状态 = 1, "执行中", IIf(mPP.病人路径状态 = 2, "完成", "变异退出"))
            End If
            .Cell(flexcpTextStyle, 0, 0) = 3

            For i = 1 To .Cols Step 2
                .TextMatrix(0, i) = " " & rsTmp!阶段名 & " "            '设置边距
                .ColAlignment(i) = flexAlignCenterCenter

                .ColWidth(i) = 1750
                .Col = i
                .PicturesOver = True
                .CellPictureAlignment = flexPicAlignLeftCenter
                If mPP.当前阶段ID = rsTmp!ID Or mPP.阶段父ID = rsTmp!ID Or (mPP.当前阶段ID = 0 And i = 1 And mPP.病人路径状态 = 1) Then
                    lngCurCol = i
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!多天 = 1, "currmultnode", "currnode")).Picture
                    Call .ShowCell(0, i)
                Else
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!多天 = 1, "multnode", "node")).Picture
                End If
                .ColData(i) = Val(rsTmp!ID)

                rsTmp.MoveNext

                '箭头
                If i < .Cols - 1 Then
                    .ColWidth(i + 1) = 550
                    .Col = i + 1
                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .CellPicture = imgFlow.ListImages(IIf(i + 1 > lngCurCol And lngCurCol <> 0 Or mPP.病人路径状态 > 1, "arrowlate", "arrow")).Picture
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get执行结果性质图标(ByVal lng执行结果性质 As Long) As Long
'功能：根据执行结果性质返回对应的图标序号
'1-已经执行，2-尚未执行，3-取消执行，4-部分执行，5-提前执行，6-延后执行
    Dim lngIdx As Long
    Select Case lng执行结果性质
        Case 1
            lngIdx = imgCharacter.ListImages("已经执行").Index
        Case 2
            lngIdx = imgCharacter.ListImages("尚未执行").Index
        Case 3
            lngIdx = imgCharacter.ListImages("取消执行").Index
        Case 4
            lngIdx = imgCharacter.ListImages("部分执行").Index
        Case 5
            lngIdx = imgCharacter.ListImages("提前执行").Index
        Case 6
            lngIdx = imgCharacter.ListImages("延后执行").Index
    End Select
    Get执行结果性质图标 = lngIdx
End Function

Private Sub LoadPathItem()
'功能：加载病人已生成的路径项目
    Dim strSql As String, strOldType As String, str评估结果 As String
    Dim lngRow As Long, lngCol As Long, i As Long, j As Long, arrtmp As Variant, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim CPos As New Collection  '每个分类的起始行
    Dim lngPreRow As Long, lngPreCol As Long, lngDayRow As Long
    Dim rsSort As Recordset
    Dim str变异原因 As String

    With vsPath
        lngPreRow = -1
        lngPreCol = -1
        If .Row >= .FixedRows Then lngPreRow = .Row
        If .Col >= .FixedCols Then lngPreCol = .Col

        '1)分类部分
        .Redraw = flexRDNone
        mblnUnChange = True
        .Clear
        .Rows = 3: .FixedRows = 3
        .Cols = 1: .FixedCols = 1
        mblnUnChange = False
        .MergeCol(0) = True
        .MergeRow(0) = True

        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 0, .FixedRows - 1, 0) = "时间阶段"
        On Error GoTo errH
        
        '读取分类的名称和个数
        strSql = " Select 分类, Max(个数) As 个数,100 as 序号" & vbNewLine & _
                 " From (Select Count(a.Id) As 个数, a.分类, a.阶段id, a.日期" & vbNewLine & _
                 "       From 病人门诊路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.分类, a.日期, a.阶段id)" & vbNewLine & _
                 " Group By 分类"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp)

        '读取序号
        strSql = " Select 分类, 序号" & vbNewLine & _
                 " From (Select 分类, 序号, Row_Number() Over(Partition By 分类 Order By 1) As Top" & vbNewLine & _
                 "       From (Select a.序号, a.名称 As 分类" & vbNewLine & _
                 "              From 门诊路径分类 A, 病人门诊路径执行 B, 门诊路径项目 C" & vbNewLine & _
                 "              Where a.名称 = c.分类 And b.路径记录id = [1] And b.项目id = c.Id And c.路径id = a.路径id And c.版本号 = a.版本号" & vbNewLine & _
                 "              Union" & vbNewLine & _
                 "              Select a.序号, a.名称 As 分类" & vbNewLine & _
                 "              From 门诊路径分类 A, 病人门诊路径执行 B, 门诊路径阶段 C" & vbNewLine & _
                 "              Where a.名称 = b.分类 And b.阶段id + 0 = c.Id And b.路径记录id = [1] And b.项目id Is Null And a.路径id = c.路径id And" & vbNewLine & _
                 "                    a.版本号 = c.版本号))" & vbNewLine & _
                 " Where Top = 1" & vbNewLine & _
                 " Order By 序号"
        Set rsSort = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        '排序
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                rsSort.Filter = "分类='" & rsTmp!分类 & "'"
                If rsSort.RecordCount > 0 Then
                    rsTmp!序号 = Val(rsSort!序号 & "")
                    rsTmp.Update
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Sort = "序号"
            rsTmp.MoveFirst
        End If
        
        For i = 1 To rsTmp.RecordCount
            CPos.Add .Rows, "T" & rsTmp!分类
            .Rows = .Rows + rsTmp!个数
            For j = 1 To rsTmp!个数
                .TextMatrix(.Rows - j, .FixedCols - 1) = rsTmp!分类
            Next
            rsTmp.MoveNext
        Next

        '2)时间阶段部分
        '阶段排序时用 NVL(c.序号,b.序号) 是为了处理备用分支序列排序的问题，取值b.序号 是因为界面上需要显示是第几个分支。
        strSql = "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,Decode(g.路径id,b.路径id,1,0) as 排序" & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id" & vbNewLine & _
                 "       From 病人门诊路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 门诊路径阶段 B,门诊路径阶段 C,病人门诊路径 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And g.id=A.路径记录ID " & vbNewLine & _
                 "Order By 日期,排序, NVL(c.序号,b.序号)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .Cols = .Cols + rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            .ColWidth(i) = 2800
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColData(i) = Val("" & rsTmp!阶段ID)
            If IsNull(rsTmp!父ID) Then
                .TextMatrix(EFixedRow.R0阶段名, i) = Replace(rsTmp!阶段名, vbLf, vbCrLf)                                '为了打印时正常换行(vbLf时换行显示有问题)
            Else
                .TextMatrix(EFixedRow.R0阶段名, i) = Replace(rsTmp!阶段名, vbLf, vbCrLf) & ",分支:" & NVL(rsTmp!说明, rsTmp!序号)
            End If
            .TextMatrix(EFixedRow.R1天数, i) = "第" & rsTmp!天数 & "天"
            .Cell(flexcpData, EFixedRow.R1天数, i) = rsTmp!天数
            .TextMatrix(EFixedRow.R2日期, i) = rsTmp!日期 & "(" & rsTmp!星期 & ")"
            .Cell(flexcpData, EFixedRow.R2日期, i) = rsTmp!日期 & ""
            
            If rsTmp!天数 = mPP.当前天数 Then
                mPP.当前日期 = rsTmp!日期
            End If
            rsTmp.MoveNext
        Next

        For i = 1 To mcolReason.count
            mcolReason.Remove 1                 '删除局部变量数据(上下移动后,重新加载时需要清空变异原因)
        Next i
        
        '3)路径项目部分
        strSql = " Select a.Id, Nvl(b.图标id, a.图标id) 图标id, a.分类, To_Char(a.日期, 'yyyy-mm-dd') 日期, a.天数, a.阶段id, Nvl(a.项目序号, b.项目序号) As 项目序号," & vbNewLine & _
                 " Nvl(b.项目内容, a.项目内容) 项目内容, a.项目id, Decode(a.执行人, Null, 0, 1) 执行状态, Nvl(b.执行方式, 1) 执行方式, a.添加原因, c.名称 As 变异原因," & vbNewLine & _
                 " Nvl(b.项目结果, a.项目结果) As 项目结果, a.执行结果, d.路径id " & vbNewLine & _
                 " From 病人门诊路径执行 A, 门诊路径项目 B, 门诊变异常见原因 C, 门诊路径阶段 D" & vbNewLine & _
                 " Where a.路径记录id = [1] And a.项目id = b.Id(+) And a.变异原因 = c.编码(+) And a.阶段id + 0 = d.Id" & vbNewLine & _
                 " Order By a.日期,分类,项目序号"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        
        For lngCol = .FixedCols To .Cols - 1
            rsTmp.Filter = "阶段ID='" & .ColData(lngCol) & "' And 天数=" & Val(Replace(.TextMatrix(EFixedRow.R1天数, lngCol), "第", ""))
            strOldType = ""

            Do While Not rsTmp.EOF
                If strOldType <> rsTmp!分类 Then
                    lngRow = CPos("T" & rsTmp!分类)
                    strOldType = rsTmp!分类
                End If

                If mbln启用执行环节 Then
                    .TextMatrix(lngRow, lngCol) = IIf(rsTmp!执行方式 = 0, "", IIf(rsTmp!执行状态 = 0, C_UnExe, C_Exe)) & rsTmp!项目内容
                Else
                    .TextMatrix(lngRow, lngCol) = "" & rsTmp!项目内容   '医嘱界面添加后，还未弹出路径外项目前刷新，项目内容为空
                End If
                '附加数据组织形式 ID|项目ID|项目序号
                '路径外项目项目id为空
                .Cell(flexcpData, lngRow, lngCol) = Val(rsTmp!ID) & "|" & Val("" & rsTmp!项目ID) & "|" & Val("" & rsTmp!项目序号)

                If IsNull(rsTmp!项目ID) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColor               '路径外项目，浅黄色
                    mcolReason.Add "变异说明：" & rsTmp!添加原因 & vbCrLf & "变异原因：" & rsTmp!变异原因, "C" & rsTmp!ID
                    If rsTmp!变异原因 & "" <> "" Or rsTmp!添加原因 & "" <> "" Then
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因 & vbCrLf & "变异说明：" & rsTmp!添加原因
                    End If
                ElseIf Val(NVL(rsTmp!执行方式)) = 1 Then                                        '必须生成的，未生成
                    If Not IsNull(rsTmp!变异原因) Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED                       '浅灰色
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因
                    End If
                ElseIf rsTmp!执行方式 = 3 Then                                                  '可选项，深蓝色
                    .Cell(flexcpForeColor, lngRow, lngCol) = &HC00000
                    If Not IsNull(rsTmp!变异原因) Then                                          '中药路径项目的变异原因
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED                       '浅灰色
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因
                    End If
                End If

                If InStr(rsTmp!项目结果, "|") > 0 And Not IsNull(rsTmp!执行结果) Then
                    i = Val(Mid(rsTmp!项目结果, InStr(rsTmp!项目结果, rsTmp!执行结果) + Len(rsTmp!执行结果) + 1, 1))
                    If i > 0 Then i = Get执行结果性质图标(i)
                Else
                    i = 0
                End If

                If Not IsNull(rsTmp!图标ID) Or i > 0 Then
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = flexPicAlignRightCenter    ' flexPicAlignLeftCenter
                    If i > 0 Then
                        .Cell(flexcpPicture, lngRow, lngCol) = imgCharacter.ListImages(i).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(rsTmp!图标ID)
                    End If
                End If

                lngRow = lngRow + 1
                rsTmp.MoveNext
            Loop
        Next

        '4)显示评估信息
        If .Rows = .FixedRows And .Cols = .FixedCols Then
            Call ClearPathItem(True)
            .BackColorSel = vbWhite
            .ForeColorSel = vbBlack
        Else
            .BackColorSel = &H8000000D
            .ForeColorSel = &H8000000E
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            lngDayRow = .FixedRows - 2
            .TextMatrix(lngRow, .FixedCols - 1) = "评估情况"
            .Cell(flexcpBackColor, lngRow, 0) = .BackColorFixed         '&HEFF0E0      '&HD0EFFF
            Call .CellBorderRange(.Rows - 1, 0, .Rows - 1, .Cols - 1, vbBlack, 0, 1, 0, 0, 0, 0)

            strSql = " Select a.阶段id, a.天数, a.评估结果, a.评估说明, a.评估人,a.评估时间, c.名称 As 变异原因, a.变异审核人, Nvl(a.时间进度, 0) 时间进度" & vbNewLine & _
                     " From 病人门诊路径评估 A, 病人门诊路径变异 B, 门诊变异常见原因 C" & vbNewLine & _
                     " Where a.路径记录id = b.路径记录id(+) And a.阶段ID=B.阶段ID(+) And a.日期=b.日期(+) And a.路径记录id = [1] And b.变异原因 = c.编码(+)"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
            For lngCol = .FixedCols To .Cols - 1
                .Cell(flexcpBackColor, lngRow, lngCol) = &HEDF8FF

                rsTmp.Filter = "阶段ID='" & .ColData(lngCol) & "' And 天数=" & Val(Replace(.TextMatrix(EFixedRow.R1天数, lngCol), "第", ""))
                str变异原因 = ""
                For j = 1 To rsTmp.RecordCount
                    '读取多个变异原因
                    str变异原因 = str变异原因 & rsTmp!变异原因 & "、"
                    If j = rsTmp.RecordCount Then
                        str变异原因 = Mid(str变异原因, 1, Len(str变异原因) - 1)
                        If InStr(rsTmp!评估说明, vbCrLf) = 0 Or IsNull(rsTmp!评估说明) Then
                            strTmp = "" & rsTmp!评估说明
                        Else
                            arrtmp = Split(rsTmp!评估说明, vbCrLf)
                            strTmp = ""
                            For i = 0 To UBound(arrtmp)
                                strTmp = strTmp & vbCrLf & Space(4) & (i + 1) & "." & arrtmp(i)
                            Next
                        End If
                        strTmp = strTmp & vbCrLf & "评 估 人：" & rsTmp!评估人
                        If rsTmp!评估结果 = 1 Then
                            str评估结果 = "正常"
                        ElseIf mPP.病人路径状态 = 3 And lngCol = .Cols - 1 Then
                            str评估结果 = "变异后退出" & vbCrLf & "变异原因：" & str变异原因 & vbCrLf & "审 核 人：" & rsTmp!变异审核人

                        ElseIf mPP.病人路径状态 = 2 And lngCol = .Cols - 1 Then
                            str评估结果 = "变异后完成" & vbCrLf & "变异原因：" & str变异原因
                        Else
                            str评估结果 = "变异后继续" & vbCrLf & "变异原因：" & str变异原因
                            If Not IsNull(rsTmp!变异审核人) Then str评估结果 = str评估结果 & vbCrLf & "审 核 人：" & rsTmp!变异审核人
                        End If

                        .TextMatrix(lngRow, lngCol) = "评估结果：" & str评估结果 & vbCrLf & "评估说明：" & strTmp
                        If rsTmp!评估结果 = -1 Then
                            .Cell(flexcpForeColor, lngRow, lngCol) = vbRed     '变异用红色表示
                        End If

                        If rsTmp!时间进度 = 1 Or rsTmp!时间进度 = 2 Then
                            '提前
                            .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "←"
                            .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                        ElseIf rsTmp!时间进度 = -1 Then    '延后
                            .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "→"
                            .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                        End If
                    End If
                    rsTmp.MoveNext
                Next
                If rsTmp.RecordCount = 0 Then
                    .TextMatrix(lngRow, lngCol) = ""
                End If
            Next
        End If
        
        '5)生成情况部分
        
        
    
        .Redraw = True
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效

        If lngPreRow <> -1 And lngPreCol <> -1 And lngPreRow <= .Rows - 1 And lngPreCol <= .Cols - 1 Then
            .Select lngPreRow, lngPreCol
        Else
            .Select .FixedRows, .FixedCols
        End If
    End With

    Exit Sub
errH:
    vsPath.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Get病人路径信息(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal lng科室ID As Long, Optional ByVal lng路径记录ID As Long)
'功能：获取病人的门诊临床路径信息
'参数：lng路径记录ID=当一个病人有多条路径时，刷新指定路径记录ID的路径表
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    '一次就诊只支持一个路径，不管科室
    '当前阶段为0表示还未生成过路径
    strSql = " Select a.ID,a.路径ID,c.路径ID as 原路径ID,a.版本号,a.状态,a.当前阶段ID,a.当前天数," & _
             " b.名称 as 未导入原因,c.父ID,e.名称 as 路径名称,a.导入人,a.导入时间,a.结束时间" & _
             " From 病人门诊路径 A,门诊变异常见原因 B,门诊路径阶段 C,门诊路径阶段 D,门诊路径目录 E" & _
             " Where a.病人ID = [1] And a.科室ID = [2] And a.路径ID=e.id And a.未导入原因 = b.编码(+) And a.当前阶段ID = c.ID(+) And a.前一阶段ID=d.id(+)" & _
             IIf(lng路径记录ID <> 0, " And a.ID=[3] ", "") & _
             " Order By a.导入时间 Desc"                                                                         '取最后一次导入的路径（支持一次就诊多个路径）
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng科室ID, lng路径记录ID)
    If rsTmp.RecordCount > 0 Then
        mPP.原路径ID = Val("" & rsTmp!原路径ID)
        mPP.路径ID = rsTmp!路径ID
        mPP.版本号 = rsTmp!版本号
        mPP.病人路径ID = rsTmp!ID
        mPP.病人路径状态 = rsTmp!状态
        mPP.当前阶段ID = Val("" & rsTmp!当前阶段ID)
        mPP.阶段父ID = Val("" & rsTmp!父ID)
        mPP.当前天数 = Val("" & rsTmp!当前天数)
        mPP.当前日期 = "0"                                      '在LoadPathItem中赋值
        mPP.未导入原因 = "" & rsTmp!未导入原因
        
        If lng路径记录ID = 0 Then mlngPathCount = rsTmp.RecordCount
    Else
        mPP.原路径ID = 0
        mPP.路径ID = 0
        mPP.版本号 = 0
        mPP.病人路径ID = 0
        mPP.病人路径状态 = -1
        mPP.当前阶段ID = 0
        mPP.阶段父ID = 0
        mPP.当前天数 = 0
        mPP.当前日期 = "0"
        mPP.未导入原因 = ""
        mlngPathCount = 0
    End If

    If mlngPathCount > 1 Then
        fraPath.Visible = True
        lblInDiag.Caption = "导入诊断:" & Get导入诊断(mPP.病人路径ID)
        lblInPep.Caption = "导入人:" & rsTmp!导入人
        lblInDate.Caption = "导入时间:" & Format(rsTmp!导入时间, "YYYY-MM-DD HH:mm")
        lblOutDate.Caption = "结束时间:" & Format(rsTmp!结束时间, "YYYY-MM-DD HH:mm")
        If lng路径记录ID = 0 Then
            cboPath.Clear
            Do While Not rsTmp.EOF
                cboPath.AddItem rsTmp!路径名称 & ""
                cboPath.ItemData(cboPath.NewIndex) = rsTmp!ID & ""
                rsTmp.MoveNext
            Loop
            zlControl.CboSetIndex cboPath.Hwnd, 0
        End If
    Else
        fraPath.Visible = False
        cboPath.Clear
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get导入诊断(ByVal lng路径记录ID As Long) As String
'功能：取导入诊断的名称
    Dim strSql As String, rsTmp As Recordset

    On Error GoTo errH
    strSql = " Select B.诊断描述 From 病人门诊路径 A,病人诊断记录 B Where " & _
             " a.病人id = b.病人id And a.挂号id = b.主页id  and a.诊断类型 = b.诊断类型 And " & _
             " a.诊断来源 = b.记录来源 And NVL(a.疾病id,0) = NVL(b.疾病id,0) And NVL(a.诊断id,0) = NVL(b.诊断id,0) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get导入诊断", lng路径记录ID)
    If rsTmp.RecordCount > 0 Then Get导入诊断 = rsTmp!诊断描述 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlPrintOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'功能：门诊临床路径表单打印
'参数：bytStyle=1-打印,2-预览,3-输出到Excel,4-输出到PDF
'      blnIsSetup-表示批量打印，不进行打印前设置
'      当bytStyle=4时，需要传入strPDFFile=PDF输出默认路径,包含文件名、后缀
    Call FuncPathTableOutput(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal str挂号NO As String, ByVal lng科室ID As Long, _
                          ByVal int病人状态 As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal blnForceRefresh As Boolean = True, _
                          Optional ByVal lng路径记录ID As Long) As Long
'参数：lng路径记录ID=当一个病人有多条路径时，刷新指定路径记录ID的路径表
'      blnForceRefresh=True 未切换病人刷新时也进行刷新，否则不刷新
    Dim objControl As CommandBarControl
    Dim strPrePati As String

    strPrePati = mPati.病人ID & "_" & mPati.挂号ID
    If strPrePati = lng病人ID & "_" & lng挂号ID And lng病人ID <> 0 And Not blnForceRefresh Then Exit Function       '保持之前单元格位置不变

    mPati.病人ID = lng病人ID
    mPati.挂号ID = lng挂号ID
    mPati.挂号NO = str挂号NO
    mPati.科室ID = lng科室ID
    mPati.病人状态 = int病人状态

    Set mcolReason = New Collection
    Call Get病人路径信息(lng病人ID, lng挂号ID, lng科室ID, lng路径记录ID)

    If mPP.病人路径ID = 0 Then
        Call SetUnImport
    Else
        If mPP.病人路径状态 = 0 Then
            Call SetImportFalse
        Else
            Call LoadPathFlow
            Call LoadPathItem
        End If
    End If
    Call Form_Resize                                '根据路径流程表是否有滚动条来调整高度
End Function

Private Sub InitCbsSubBar()
    Dim objBar As CommandBar

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True         '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False     'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False

    Set objBar = cbsSub.Add("内部工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False              '只有内部调用时才显示(zlDefCommandBars)
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim lngStart As Long, i As Long

    mbln启用执行环节 = Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, P门诊路径应用, 1))

    Set mfrmParent = frmParent

    If cbsMain Is Nothing Then Exit Sub

    Set mcbsMain = cbsMain
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '文件菜单
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Find(, conMenu_File_Excel)
        objControl.Caption = "输出到&Excel(医师版)…"
        Set objControl = .Find(, conMenu_File_Print)
        objControl.Caption = "打印路径表(医师版)(&P)"
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_PatiPath, "打印路径表(患者版)(&Q)", objControl.Index + 1)
        objControl.IconId = conMenu_File_Print
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "路径(&E)", objPopup.Index + 1, False)
    objPopup.ID = conMenu_EditPopup
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "导入路径(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消导入")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "生成路径项目(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "补充生成项目(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "重新生成医嘱")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "添加路径外项目")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改路径外项目")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "取消本次生成(&X)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "取消当前项目(&V)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive, "项目执行登记(&E)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "取消执行登记(&Z)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "批量执行登记(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "批量取消执行(&F)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "评估(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "修改评估")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "取消评估")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "完成路径(&O)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消完成")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogModi, "修改出径登记表")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "上移")
        objControl.IconId = conMenu_Manage_Up
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "下移")
        objControl.IconId = conMenu_Manage_Down
    End With

    '查看菜单
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "标准路径参考")
        objControl.BeginGroup = True
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Select, "查看导入评估")
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "查看出径登记表")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)", objControl.Index + 1)
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        '
        Set objControl = .Add(xtpControlButton, conMenu_Edit_View, "查看项目定义(&A)")
    End With

    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objPopup.Index, False)
        objPopup.ID = conMenu_ToolPopup
    End If

    '工具栏定义
    '-----------------------------------------------------
    lngStart = 0
    Set cbrToolBar = cbsMain(2)
    For Each objControl In cbrToolBar.Controls    '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = cbrToolBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    lngStart = objControl.Index + 1

    If lngStart <> 0 Then
        With cbrToolBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "导入", lngStart)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "生成", objControl.Index + 1)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "补充", objControl.Index + 1)
            objControl.ToolTipText = "补充生成可选生成的路径项目"

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "执行", objControl.Index + 1)
            objControl.BeginGroup = True
            objControl.ToolTipText = "批量执行路径项目"
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "评估", objControl.Index + 1)
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "完成", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "上移", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Up
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "下移", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Down

            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "报告", objControl.Index + 1)
            objPopup.BeginGroup = True
            objPopup.IconId = conMenu_Manage_Report
            objPopup.ToolTipText = "查阅报告"
        End With
    End If
End Sub

Private Sub FuncPatiPathPrint()
'功能：输出患者版临床路径
    Dim WordApp As Object       'Word.Application
    Dim WordDoc As Object       'Word.Document
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strFileName As String, strFilePath As String
    Dim lngRetu As Long, strInfo As String

    If vsPath.FixedRows < 3 Then
        MsgBox "该病人还未生成门诊路径项目。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    '获得路径
    Screen.MousePointer = 11
    strSql = "Select 文件名 from 门诊路径文件 where 路径ID=[1] And 类别=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID)
    If rsTmp.RecordCount > 0 Then
        strFileName = rsTmp!文件名 & ""
        strFilePath = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & strFileName
        If gobjFile.FileExists(strFilePath) Then gobjFile.DeleteFile strFilePath, True
        '将数据库中BLOB数据读到本地临时文件目录下
        strFilePath = Sys.ReadLob(glngSys, 26, mPP.路径ID & "," & strFileName, strFilePath)
        If Not gobjFile.FileExists(strFilePath) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Sub
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "该路径表没有设置对应的门诊临床路径表(患者版),请到门诊临床路径管理中设置。", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordApp = CreateObject("Word.Application")
    If WordApp Is Nothing Then
        MsgBox "请安装Microsoft Office Word。", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordDoc = WordApp.Documents.Open(strFilePath)      '打开RTF文档
    WordDoc.PrintPreview
    WordApp.Visible = True
    WordApp.ScreenUpdating = True
    WordApp.Activate
    Screen.MousePointer = 0

    '记录打印信息
    Call zlDatabase.ExecuteProcedure("Zl_电子病历打印_Insert(" & mPP.病人路径ID & ",12," & mPati.病人ID & "," & mPati.挂号ID & ",'" & UserInfo.姓名 & "')", "打印患者版路径表")
    '打印后强制重新加载提示信息，更新提示信息
    Call LoadPathFlow
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'功能：输出门诊临床路径表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel,4-输出到PDF
'      blnIsSetup-批量打印不进行打印前设置
'      strPDFFile=PDF输出默认路径
'      strDeviceName=指定打印机名称
    Dim rsTmp As ADODB.Recordset
    Dim vsBody As VSFlexGrid
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngColor As Long, bytR As Byte
    Dim strSql As String
    Dim rsSQLTmp As ADODB.Recordset
    Dim strDisease As String            '诊断描述
    Dim strStandardDate As String       '标准治疗时间
    Dim i As Long, j As Long
    Dim strTitle As String
    Dim strTmp As String
    Dim lngDefDay As Long

    strSql = " Select a.病人id, a.挂号id, b.疾病id, b.诊断id, b.诊断描述, c.标准治疗时间" & vbNewLine & _
             " From 病人门诊路径 A, 病人诊断记录 B, 门诊路径版本 C" & vbNewLine & _
             " Where a.病人id = b.病人id And a.挂号id = b.主页id And a.诊断类型 = b.诊断类型" & vbNewLine & _
             " And a.诊断来源 = b.记录来源 And c.路径id = a.路径id And c.版本号 = a.版本号 And" & vbNewLine & _
             " b.诊断次序 = 1 And a.病人id = [1] And a.挂号id = [2] And a.ID=[3]"
    mblnUnChange = True
    If vsPath.FixedRows < 3 Then
        '输出PDF，如果不是路径病人，则直接退出不提示
        If bytStyle = 4 Then Exit Sub
        '批量打印不提示
        If blnIsSetup Then Exit Sub
        MsgBox "该病人还未生成门诊路径项目。", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error GoTo errH
    Set rsTmp = GetPatiInfoOut(mPati.病人ID, mPati.挂号ID)
    Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.挂号ID, mPP.病人路径ID)

    If rsSQLTmp.RecordCount > 0 Then
        strDisease = rsSQLTmp!诊断描述 & ""
        strStandardDate = rsSQLTmp!标准治疗时间 & ""
    Else
        strDisease = ""
        strStandardDate = ""
    End If
    '表头
    If InStr(vsFlow.TextMatrix(0, 0), vbCrLf) > 0 Then
        strTitle = Mid(vsFlow.TextMatrix(0, 0), 1, InStr(vsFlow.TextMatrix(0, 0), vbCrLf) - 1)
    Else
        strTitle = vsFlow.TextMatrix(0, 0)
    End If
    objOut.Title.Text = strTitle & vbCrLf & "门诊临床路径表"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 20
    objOut.Title.Font.Bold = True

    '表上
    strSql = " Select a.诊断描述 From 病人诊断记录 A" & vbNewLine & _
             " Where a.病人id = [1] And a.主页id = [2] And a.记录来源 = 3 And a.诊断类型 In (1, 11) Order By a.诊断次序"
    Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.挂号ID)
    If rsSQLTmp.RecordCount > 0 Then
        strTmp = rsSQLTmp!诊断描述 & ""
        strTmp = Mid(strTmp, InStr(strTmp, ")") + 1) & Mid(strTmp, 1, InStr(strTmp, ")"))
    Else
        strTmp = ""
    End If
   
    Set objRow = New zlTabAppRow
    objRow.Add "适用对象：第一诊断为 " & strTmp
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "患者姓名：" & rsTmp!姓名 & " 性别：" & rsTmp!性别 & " 年龄：" & rsTmp!年龄 & " 门诊号：" & rsTmp!门诊号
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "就诊日期:" & Format(rsTmp!接收时间, "yyyy年MM月dd日")
    objRow.Add "完成就诊日期:" & Format(rsTmp!完成时间, "yyyy年MM月dd日")
    objRow.Add "标准治疗时间：" & IIf(InStr(strStandardDate, "-") > 0, "", "≤") & strStandardDate & "天"
    objOut.UnderAppRows.Add objRow
    objOut.AppFont.Size = 12
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow

    '页脚
    objOut.Footer = ";第[页码]页，共[页数]页;"
    objOut.PageFooter = 5

    '表体
    strTmp = zlDatabase.GetPara("路径表单打印规则", glngSys, P门诊路径应用, "0")
    If strTmp = "1" Then
        Set vsBody = FuncConvertPathTable
    Else
        Set vsBody = vsPath
    End If

    '输出
    With vsBody
        .Redraw = flexRDNone
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "医生签名"
        .RowHeight(.Rows - 1) = 440

        '默认打印天数
        lngDefDay = Val(zlDatabase.GetPara("路径表单每页打印的天数", glngSys, P门诊路径应用, "2"))
        objOut.PageCols = lngDefDay + .FixedCols
        '表格列数不够时，补充空列
        If (.Cols - 1) Mod lngDefDay <> 0 Then
           .Cols = .Cols + (lngDefDay - ((.Cols - 1) Mod lngDefDay))
        End If
        '打印表格转换
        Call FuncPathTableChange(vsBody, lngDefDay)

        '破坏合并列首行,打印部件中对合并列有单独处理
        For i = .FixedCols To .Cols - 1
            If i Mod 2 = 0 Then
                .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, i) & vbTab
            End If
        Next
        .Redraw = flexRDDirect
        '行宽自适应
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效

        objOut.FixCol = vsBody.FixedCols
        objOut.FixRow = vsBody.FixedRows
        Set objOut.Body = vsBody

        '指定打印机
        If strDeviceName <> "" Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", strDeviceName
        If bytStyle = 1 Or bytStyle = 4 Then
            If bytStyle = 4 Then
                bytR = 4
                objOut.Privileged = True '电子病案查阅 用于公共部件Zl9PrintMode内部跳过打印权限检查
            Else
                If Not blnIsSetup Then
                    bytR = zlPrintAsk(objOut)
                Else
                    bytR = 1
                End If
            End If
            Me.Refresh

            If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR, strPDFFile
            '打印完了产生打印记录
            strSql = "zl_电子病历打印_insert(" & mPP.病人路径ID & ",11," & mPati.病人ID & "," & mPati.挂号ID & ",'" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Else
            zlPrintOrView1Grd objOut, bytStyle
        End If
        mblnUnChange = False
        '恢复到初始状态
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)    'vsPath变动后重新加载

        If vsPathPrint.UBound = 1 Then Unload vsPathPrint(1)
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Import, conMenu_Edit_Untread, conMenu_Edit_ImportMerge, conMenu_Edit_UnImportMerge
            If InStr(GetInsidePrivs(P门诊路径应用), ";导入路径;") = 0 Then blnVisible = False
        Case conMenu_Edit_Send, conMenu_Edit_Append, conMenu_Edit_Delete, conMenu_Edit_Blankoff, conMenu_Edit_SendBack
            If InStr(GetInsidePrivs(P门诊路径应用), ";生成路径;") = 0 Then blnVisible = False
            If Control.ID = conMenu_Edit_SendBack And blnVisible Then
                blnVisible = Not InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0
            End If
        Case conMenu_Edit_Surplus, conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down
            If InStr(GetInsidePrivs(P门诊路径应用), ";路径外项目;") = 0 Then blnVisible = False
        Case conMenu_Edit_Archive, conMenu_Edit_UnArchive, conMenu_Edit_Merge, conMenu_Edit_DeleteParent
            If InStr(GetInsidePrivs(P门诊路径应用), ";执行路径;") = 0 Or mbln启用执行环节 = False Then blnVisible = False
            '启用路径执行环节时，启用场合和当前场合不一致时,隐藏菜单按钮
        Case conMenu_Edit_Audit, conMenu_Edit_Reuse, conMenu_Edit_Clear
            If InStr(GetInsidePrivs(P门诊路径应用), ";阶段评估;") = 0 Then blnVisible = False
        Case conMenu_Edit_Stop, conMenu_Edit_ClearUp
            If InStr(GetInsidePrivs(P门诊路径应用), ";结束路径;") = 0 Then blnVisible = False
        Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView
            If Control.ID = conMenu_Edit_OutLogModi Then
                If InStr(GetInsidePrivs(P门诊路径应用), ";结束路径;") = 0 Then blnVisible = False
            End If
            If blnVisible Then blnVisible = CheckPathOutLog
        Case conMenu_Edit_Compend
            '报告弹出(含打印),查阅报告
            If InStr(GetInsidePrivs(p住院医嘱下达), ";报告查阅;") = 0 Then blnVisible = False
    End Select
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveItem As Boolean
    Dim lng项目ID As Long

    If vsPath.Redraw = flexRDNone Then Exit Sub

    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub

    With vsPath
        blnHaveItem = .Row > .FixedRows - 1 And .FixedRows <> 0 And .Col > .FixedCols - 1   '.FixedRows=0时，只有一行提示信息
    End With
    Select Case Control.ID
        '0.输出
    Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel, conMenu_File_Print_PatiPath
        Control.Enabled = mPP.病人路径ID <> 0
        '1.导入
        '-----------------------------------------
    Case conMenu_Edit_Import    '导入路径
        Control.Enabled = (mPati.病人状态 = pt回诊 Or mPati.病人状态 = pt就诊) And mPati.病人ID <> 0 And cboPath.ListIndex <= 0
    Case conMenu_Edit_Untread   '取消导入(仅在第一次生成时可取消导入)
        Control.Enabled = (mPati.病人状态 = pt回诊 Or mPati.病人状态 = pt就诊) And mPP.病人路径ID <> 0 And (mPP.病人路径状态 = 0 Or mPP.病人路径状态 = 1) And vsPath.Cols <= vsPath.FixedCols + 1
    Case conMenu_Edit_Select      '查看导入评估
        Control.Enabled = mPP.病人路径ID <> 0
        '2.生成
        '-----------------------------------------
    Case conMenu_Edit_Send      '生成路径
        Control.Enabled = (mPati.病人状态 = pt回诊 Or mPati.病人状态 = pt就诊) And mPP.病人路径ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Append    '补充生成
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Blankoff  '取消本次生成
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Delete, conMenu_Edit_SendBack   '取消路径项目,重新生成医嘱
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Col > 0 Then
                    Control.Enabled = (.ColData(.Col) = mPP.当前阶段ID And .Col = .Cols - 1)
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Surplus   '添加路径外项目
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down     '修改路径外项目
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" And vsPath.Row <> vsPath.Rows - 1 Then
                lng项目ID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '路径外项目为0
                Control.Enabled = lng项目ID = 0
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_View      '查看项目定义
        Control.Enabled = blnHaveItem
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" Then
                lng项目ID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '路径外项目为0
                Control.Enabled = lng项目ID <> 0
            End If
        End If
        '3.执行
        '-----------------------------------------
    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive   '单个项目执行(仅最后一次的列才能) '取消执行
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Row >= .FixedRows Then
                    Control.Enabled = (.ColData(.Col) = mPP.当前阶段ID And .Col = .Cols - 1)
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Merge, conMenu_Edit_DeleteParent    '批量执行,批量取消执行
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        '4.评估
        '-----------------------------------------
    Case conMenu_Edit_Audit     '阶段评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Reuse     '修改评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Clear     '取消评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        '5.完成
        '-----------------------------------------
    Case conMenu_Edit_Stop      '完成路径
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then    '当前天数到达标准治疗时间范围且已评估才允许完成
            Control.Enabled = mblnInOverScope And vsPath.TextMatrix(vsPath.Rows - 1, vsPath.Cols - 1) <> ""
        End If
    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView   '出径登记表
        Control.Enabled = (mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3)     '2-正常完成，3-变异完成
    Case conMenu_Edit_ClearUp   '取消完成
        If mPP.病人路径状态 = 3 Then
            Control.Caption = "取消退出"
        Else
            Control.Caption = "取消完成"
        End If
        Control.Enabled = (mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3) And cboPath.ListIndex <= 0    '2-正常完成，3-变异完成
    Case conMenu_Edit_Compend    '查看报告
        With vsPath
            Control.Enabled = blnHaveItem
            If Control.Enabled Then Control.Enabled = .Cell(flexcpData, .Row, .Col) <> ""
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As ADODB.Recordset, str疾病编码 As String
    Dim blnDo As Boolean
    Dim strTmp As String

    Select Case Control.ID
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
    Case conMenu_File_Print_PatiPath
        '打印患者版路径表
        Call FuncPatiPathPrint
        '1.导入
        '-----------------------------------------
    Case conMenu_Edit_Import                '导入路径
        Call FuncImport(, True)
    Case conMenu_Edit_Untread               '取消导入
        Call FuncUnImport
    Case conMenu_Edit_Select                '查看导入评估
        Call frmEvaluateOut.ShowMe(mfrmParent, 0, 0, mPati, mPP)
        '2.生成
        '-----------------------------------------
    Case conMenu_Edit_Send                  '生成路径
        Call FuncSendItem
    Case conMenu_Edit_Append                '补充生成
        Call FuncSendItemApend
    Case conMenu_Edit_Delete                '取消已生成的项目
        Call FuncDelItem
    Case conMenu_Edit_Blankoff              '取消本次生成
        Call FuncDelAllItem
    Case conMenu_Edit_SendBack              '重新生成医嘱
        Call FuncReSendItem
    Case conMenu_Edit_Surplus               '添加路径外项目
        Call FuncAppendItem(0)
    Case conMenu_Edit_Modify                '修改路径外项目
        Call FuncAppendItemModify
        '3.执行
        '-----------------------------------------
    Case conMenu_Edit_Archive               '执行路径
        Call FuncExecuteItem
    Case conMenu_Edit_Merge                 '批量执行
        Call FuncExecuteAll
    Case conMenu_Edit_UnArchive             '取消执行
        Call FuncExecuteItemCancel
    Case conMenu_Edit_DeleteParent          '批量取消执行
        Call FuncExecuteAllCancel
        '4.评估
        '-----------------------------------------
    Case conMenu_Edit_Audit                 '评估
        Call FuncEvaluate
    Case conMenu_Edit_Reuse                 '修改评估
        Call FuncReEvaluate
    Case conMenu_Edit_Clear                 '取消评估
        Call FuncEvaluateCancel
        '5.完成
        '-----------------------------------------
    Case conMenu_Edit_Stop                  '完成路径
        Call FuncOver
    Case conMenu_Edit_ClearUp               '取消完成
        Call FuncOverCancel
    Case conMenu_Edit_OutLogModi    '修改出径登记表
        Call OutLogModi
    Case conMenu_Edit_OutLogView   '查看出径登记表
        Call frmPathOutLogOut.ShowMe(mfrmParent, mPati.病人ID, mPati.挂号ID, 1, Nothing, mPP.路径ID, mPP.病人路径ID)
        '6.上移，下移
        '-----------------------------------------
    Case conMenu_Edit_Up                    '1-上移
        Call MovePathItem(1)
    Case conMenu_Edit_Down                  '-1-下移
        Call MovePathItem(-1)
        '7.其它
        '-----------------------------------------
    Case conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 10  '限制最多10个报告
        If InStr(Control.Parameter, ":") > 0 Then
            Call FuncViewReport(Split(Control.Parameter, ":")(0), Split(Control.Parameter, ":")(1))
        End If
    Case conMenu_Edit_View                  '显示路径项目定义的信息
        Call vsPath_DblClick
    Case conMenu_View_StPath                '查看标准路径参考
        Set rsTmp = GetPatiDiagnose(mPati.病人ID, mPati.挂号ID, 2)  '获取首要诊断
        If rsTmp.RecordCount <> 0 Then
            str疾病编码 = rsTmp!编码
        End If
        Call frmStPathList.ShowMe(mfrmParent, str疾病编码, 1)
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
'功能：定义菜单项的弹出菜单
    Dim objControl As CommandBarControl
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long
    Dim rsTmpPacs As Recordset

    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Edit_Compend
             With CommandBar.Controls
                .DeleteAll
                With vsPath
                    Set rsTmp = GetReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                    Set rsTmpPacs = GetPACSReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                End With

                If rsTmp.RecordCount = 0 And rsTmpPacs.RecordCount = 0 Then
                     .Add xtpControlButton, conMenu_Edit_Compend * 10 + 1, "无报告或未书写"
                Else
                    For i = 1 To rsTmp.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i, rsTmp!病历名称 & "(&" & i & ")")
                        objControl.Parameter = rsTmp!ID & ":" & rsTmp!医嘱id
                        rsTmp.MoveNext
                    Next
                    i = rsTmp.RecordCount
                    For j = 1 To rsTmpPacs.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i + j, rsTmpPacs!文档标题 & "(&" & i + j & ")")
                        objControl.Parameter = rsTmpPacs!报告ID & ":" & rsTmpPacs!医嘱id
                        rsTmpPacs.MoveNext
                    Next
                End If
            End With
    End Select
End Sub

Private Function GetReportOfPath(ByVal lng路径执行ID As Long) As ADODB.Recordset
'功能：获取路径对应的报告名称
    Dim strSql As String

    strSql = " Select d.id, d.病历名称,c.医嘱Id" & vbNewLine & _
             " From 病人门诊路径执行 A, 病人门诊路径医嘱 B, 病人医嘱报告 C, 电子病历记录 D" & vbNewLine & _
             " Where a.Id = [1] And a.Id = b.路径执行id And b.病人医嘱id = c.医嘱Id And c.病历id = d.Id"
    On Error GoTo errH
    Set GetReportOfPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径执行ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPACSReportOfPath(ByVal lng路径执行ID As Long) As ADODB.Recordset
'功能：获取路径对应的报告名称
    Dim strSql As String
    Dim strIDs As String
    Dim rsTmp As Recordset

    strSql = " Select b.病人医嘱id" & vbNewLine & _
             " From 病人门诊路径执行 A, 病人门诊路径医嘱 B" & vbNewLine & _
             " Where a.Id = [1] And a.Id = b.路径执行id "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径执行ID)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!病人医嘱id
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Set GetPACSReportOfPath = mobjPublicPACS.zlDocGetListWithAdvice(strIDs)
    Else
        Set rsTmp = New Recordset
        rsTmp.Fields.Append "ID", adInteger, 1
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        Set GetPACSReportOfPath = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.姓名)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS公共部件未创建成功！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Private Sub FuncOver()
'功能：完成路径
    Dim strSql As String, blnOK As Boolean, lngValue As Long
    Dim colSQL As New Collection, blnTrans As Boolean, i As Long
    Dim str审核人 As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long

    On Error GoTo errH

    str审核人 = UserInfo.姓名

    If MsgBox("你确定要完成当前病人的门诊临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If

    lngPPStatus = mPP.病人路径状态
    
    If CheckPathOutLogOut Then
        blnOK = frmPathOutLogOut.ShowMe(mfrmParent, mPati.病人ID, mPati.挂号ID, 0, colSQL, mPP.路径ID, mPP.病人路径ID)
        If blnOK = False Then
            lngValue = Val(zlDatabase.GetPara("必须填写出径登记表", glngSys, P门诊路径应用, "0"))
            If lngValue = 1 Then
                MsgBox "由于完成路径前必须填写出径登记表，你取消了填写，路径完成操作未执行。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If

    strSql = "Zl_病人门诊路径结束_Update(" & mPP.病人路径ID & ")"
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSql, "取消路径完成")
    gcnOracle.CommitTrans: blnTrans = False

    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)

    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncOverCancel()
'功能：取消路径的完成
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long

    On Error GoTo errH
    '完成后没有新增可用的医嘱才允许取消
    strSql = " Select 1 From 病人门诊路径 A, 病人医嘱记录 B, 病人挂号记录 C, 病人门诊路径记录 D " & vbNewLine & _
             " Where a.ID = d.路径记录ID And d.挂号ID = C.ID And  B.挂号单 = c.No And" & vbNewLine & _
             "       b.开嘱时间 > Trunc(a.结束时间, 'MI') And b.医嘱状态 Not In (-1, 4) And Nvl(b.婴儿, 0) = 0 And a.Id = [1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "路径完成后已产生了新的医嘱，请删除或作废后再进行取消操作。", vbInformation, gstrSysName
        Exit Sub
    End If

    If mPP.病人路径状态 = 3 Then
        If MsgBox("当前路径是变异后自动完成的，取消后变异评估结果将同时删除，并且取消评估，你确定要继续吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    lngPPStatus = mPP.病人路径状态

    strSql = "Zl_病人门诊路径结束_Delete(" & mPP.病人路径ID & "," & mPP.病人路径状态 & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消路径完成")
    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)

    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItem()
'功能：执行路径项目
    Dim lng执行ID As Long, lng项目ID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String

    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)    '路径外项目为0
    End With

    strSql = "Select 1 From 病人门诊路径执行 Where ID = [1] And 执行时间 is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目已执行。", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmPathExecute.ShowMe(mfrmParent, 1, mPati, mPP, lng执行ID, 0, , 1) Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteAll()
'功能：批量执行
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, 0, , 1) Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItemCancel()
'功能：取消路径项目的执行
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng执行ID As Long
    Dim blnTip As Boolean

    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With

    strSql = "Select 1 From 病人门诊路径执行 Where ID = [1] And 执行时间 is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目还未执行。", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, Val(vsPath.ColData(vsPath.Col)), CDate(vsPath.Cell(flexcpData, EFixedRow.R2日期, vsPath.Col)))
    If rsTmp.RecordCount > 0 Then
        '强制取消评估，不检查权限
        If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能取消执行。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call FuncEvaluateCancel(False, False)
        Else
            Exit Sub
        End If
    Else
        blnTip = True
    End If

    If blnTip Then
        If MsgBox("你确定要取消[" & vsPath.TextMatrix(vsPath.Row, vsPath.Col) & "]的执行吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    strSql = "Zl_病人门诊路径执行_Delete(" & lng执行ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FuncExecuteAllCancel(Optional blnRefresh As Boolean = True) As Boolean
'功能：批量取消路径项目的执行
'说明：医生站评估的时候会检查医生生成者的执行登记情况
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnDo As Boolean

    On Error GoTo errH

    If blnRefresh = True Then
        strSql = " Select 1 From 病人门诊路径执行 A,门诊路径项目 B Where A.路径记录ID = [1] And A.阶段ID = [2] And A.天数 = [3] And A.项目ID=B.ID(+) " & _
                 " And A.执行时间 is Not Null And Rownum<2 "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
        If rsTmp.RecordCount = 0 Then
            MsgBox "当前不存在由医生执行登记的任何项目。", vbInformation, gstrSysName
            FuncExecuteAllCancel = True
            Exit Function
        End If
    End If

    '评估环节检查
    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        '强制取消评估，不检查权限
        If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能取消执行。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call FuncEvaluateCancel(False, True)
        Else
            Exit Function
        End If
    End If

    blnDo = frmPathExecute.ShowMe(mfrmParent, 2, mPati, mPP, 0, 0, False, 1)
    If blnDo And blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    FuncExecuteAllCancel = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncSendItem(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'功能：执行生成路径
'参数：blnIsCancel，没有路径可生成时，用户是否取消了评估。true=取消
'     lngType:1-医嘱编辑界面调用，则评估后不继续生成，因为医嘱编辑界面不能再调用医嘱编辑。
    Dim rsTmp As ADODB.Recordset
    Dim lng天数 As Long, lng时间进度 As Long, lng理论天数 As Long
    Dim lng阶段ID As Long
    Dim lngPPStatus As Long
    Dim strTmp As String
    Dim strSql As String
    Dim strDate As String
    Dim strPhase As String
    Dim strMsg As String
    Dim blnDo As Boolean
    Dim blnIsNext As Boolean
    Dim blnEvaluate As Boolean
    Dim blnRefresh As Boolean

    On Error GoTo errH

    If mPP.当前天数 = 0 Then '第一天
        strSql = " Select To_number(Trunc(Sysdate)-Trunc(a.开始时间)+1) as 就诊天数 " & _
                 " From 病人门诊路径 a,门诊路径目录 b Where a.ID = [1] And a.路径id = b.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        lng天数 = rsTmp!就诊天数
        lng时间进度 = 2
    Else
        '2.当前未评估，不允许生成新的
        strSql = "Select 时间进度 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))

        If rsTmp.RecordCount = 0 Then
            If InStr(GetInsidePrivs(P门诊路径应用), ";阶段评估;") = 0 Then
                MsgBox "该病人在" & mPP.当前日期 & "还没有进行评估，不能进行后续操作。", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("该病人在" & mPP.当前日期 & "还没有进行评估，必须先评估。" & vbCrLf & "你现在要进行评估操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '评估前需先检查执行登记情况
                    If Not CheckPathIsExecuted() Then
                        Exit Function
                    End If

                    If frmEvaluateOut.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                        Exit Function
                    Else
                        lngPPStatus = mPP.病人路径状态
                        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)

                        '评估后，可能结束或退出路径，所以根据评估有的状态进行判断是否要继续生成,退出或完成则不继续生成
                        If mPP.病人路径状态 <> 1 Or lngType = 1 Then
                            Exit Function
                        End If

                        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
                        strSql = "Select 时间进度 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
                        If rsTmp.RecordCount <> 0 Then
                            lng时间进度 = Val("" & rsTmp!时间进度): blnEvaluate = True
                        Else
                            Exit Function
                        End If
                        blnIsNext = True
                    End If
                Else
                    Exit Function
                End If
            End If
        Else
            lng时间进度 = Val("" & rsTmp!时间进度): blnEvaluate = True
        End If

        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        If lng时间进度 = 0 Then
            If mPP.当前日期 = strDate Then
                lng理论天数 = GetMustDayOut(mPP.病人路径ID, mPP.当前天数)
                'a.如果当天还有其它阶段，允许生成其他阶段，但天数仍是当天
                If CheckSameDayOfPhaseOut(mPP.当前阶段ID, lng理论天数) Then
                    lng天数 = mPP.当前天数
                Else
                    MsgBox "该病人当天没有其他可用的阶段可以生成。", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mPP.当前日期 < strDate Then          '
                lng天数 = DateDiff("D", mPP.当前日期, strDate) + IIf(mPP.当前天数 = 0, 0, mPP.当前天数)
            Else                                        'c.提前生成后续阶段
                Exit Function
            End If
        ElseIf lng时间进度 = 1 Then                     '下一阶段提前至今天(时间不变，同一天生成多个阶段的内容)
            lng天数 = mPP.当前天数
        ElseIf lng时间进度 = 2 Then                     '下一阶段提前至明天
            If mPP.当前日期 = strDate Then
                MsgBox "上一阶段评估为“下一阶段提前至明天”,当天没有其他可用的阶段可以生成。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            lng天数 = DateDiff("D", mPP.当前日期, strDate) + 1
        Else                                            '下一阶段延后(继续当前阶段)
            If mPP.当前日期 = strDate Then
                MsgBox "该病人在今天的路径已生成。", vbInformation, gstrSysName
                Exit Function
            End If
            lng天数 = DateDiff("D", mPP.当前日期, strDate) + 1
        End If
    End If

    If frmPathSendOut.ShowMe(mfrmParent, 0, mPati, mPP, mPP.当前阶段ID, lng天数, 0, 0, lng时间进度, blnDo) Then
        FuncSendItem = True
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncSendItemApend()
'功能：补充生成路径
'      当天的路径已生成时才允许补充
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strTmp As String
    Dim strDate As String

    On Error GoTo errH
    strSql = "Select Max(ID) as ID From 病人门诊路径执行 Where 路径记录ID = [1] And 阶段ID = [2] And 天数 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
    If IsNull(rsTmp!ID) Then
        MsgBox "该病人在今天的路径还没有生成。", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P门诊路径应用), ";阶段评估;") = 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再补充生成项目。", vbInformation, gstrSysName
            Exit Sub
        Else
            '取消评估
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能补充生成项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If
    
    If frmPathSendOut.ShowMe(mfrmParent, 1, mPati, mPP, mPP.当前阶段ID, mPP.当前天数) Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncReSendItem()
'功能：重新生成路径项目的医嘱
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng执行ID As Long, lng项目ID As Long, blnMust As Boolean, lng天数 As Long

    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Val(Split(.Cell(flexcpData, .Row, .Col), "|")(1))
    End With
    If lng项目ID = 0 Then
        MsgBox "要重新生成路径外项目，请取消该项目的生成后重新添加。", vbInformation, gstrSysName
        Exit Sub
    End If

    '1.已经执行的不允许重新生成;已经发送的医嘱不能重新生成
    strSql = "Select a.执行时间, c.医嘱状态" & vbNewLine & _
            "From 病人门诊路径医嘱 B, 病人医嘱记录 C, 病人门诊路径执行 A" & vbNewLine & _
            "Where a.Id = b.路径执行id And b.病人医嘱id = c.Id And a.Id = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!执行时间) And mbln启用执行环节 Then
            If rsTmp.RecordCount > 0 Then
                MsgBox "该项目已执行，不能重新生成。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        rsTmp.Filter = "医嘱状态=8"
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目对应的医嘱已经发送生效，请作废后再执行此操作。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        MsgBox "该项目不是医嘱类项目，不能重新生成。", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmPathSendOut.ShowMe(mfrmParent, 3, mPati, mPP, mPP.当前阶段ID, mPP.当前天数, lng项目ID, lng执行ID) Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDelPhaseItem()
'功能：强制删除最后一天所有的执行项目(用于测试时清除数据)
    Dim strSql As String
    Dim lng执行ID As Long
    Dim i As Long

    On Error GoTo errH
    With vsPath
        For i = .FixedRows To .Rows - 2     '最后一行是评估
            If .TextMatrix(i, .Cols - 1) <> "" Then
                lng执行ID = Split(.Cell(flexcpData, i, .Cols - 1), "|")(0)
                strSql = "Zl_病人门诊路径生成_Delete(" & lng执行ID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
            End If
        Next
    End With
    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    Exit Sub
 Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FuncDelAllItem(Optional ByVal blnRefresh As Boolean = True, Optional ByVal blnPrompt As Boolean = True) As Boolean
'功能：整体取消本次生成的所有路径项目
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, strIDs As String, strIDSQL As String, blnTrans As Boolean
    Dim strNewIDs As String
    Dim blnExecuted As Boolean
    Dim dat导入时间 As Date
    Dim lng天数 As Long
    
    If blnPrompt Then
        If MsgBox("取消生成将删除路径项目对应的医嘱和病历文件。" & vbCrLf & "你确实要取消本次生成的所有路径项目吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    On Error GoTo errH
    
    strSql = "Select A.ID,A.执行时间 From 病人门诊路径执行 A,门诊路径项目 B Where A.路径记录ID = [1] And A.阶段ID = [2] And A.天数 = [3] and A.项目ID=B.ID(+) "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
       
    Do While Not rsTmp.EOF
        If blnExecuted = False Then
            If Not IsNull(rsTmp!执行时间) Then
                blnExecuted = True
            End If
        End If
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If blnExecuted Then
        '不判断权限，不提示，强制取消
        If FuncExecuteAllCancel(False) = False Then
            Exit Function
        End If
    End If
    
    strSql = "Select 导入时间 from 病人门诊路径 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    dat导入时间 = Format(rsTmp!导入时间 & "", "yyyy-MM-dd HH:mm:ss")
    '检查是否已评估
    If mbln启用执行环节 = False Or Not blnExecuted Then
        strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            '强制取消评估，不检查权限
            MsgBox "本次生成的项目已评估，取消生成之前将自动取消评估。", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    
    strIDSQL = "(Select Column_value From Table(f_Str2List([1])))"
    '2.检查医嘱
    '不是当天生成的长嘱，允许取消路径项目，不管是否发送；
    '是当天生成的长嘱，已校对但未作废，不允许取消，未校对的，取消时自动删除对应的医嘱。

    strSql = "Select /*+ Rule*/ distinct A.路径执行id" & vbNewLine & _
             "From 病人门诊路径医嘱 A, 病人门诊路径医嘱 B" & vbNewLine & _
             "Where a.路径执行id In " & strIDSQL & " And a.病人医嘱id = b.病人医嘱id And b.路径执行id <> a.路径执行id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs)
    If rsTmp.RecordCount = 0 Then
        strNewIDs = strIDs
        '没有非当日的长嘱
    Else
      '把以前生成了长嘱的那部分去掉，只检查当天的
        strNewIDs = "," & strIDs & ","
        For i = 1 To rsTmp.RecordCount
            If InStr(strNewIDs, "," & rsTmp!路径执行id & ",") > 0 Then
                strNewIDs = Replace(strNewIDs, "," & rsTmp!路径执行id & ",", ",")
            End If
            rsTmp.MoveNext
        Next
        If strNewIDs = "," Then
            strNewIDs = ""
        Else
            strNewIDs = Mid(strNewIDs, 2, Len(strNewIDs) - 2)
        End If
    End If
    
    If strNewIDs <> "" Then
        '即使已停止的医嘱也不允许删除，加59秒是因为开嘱时间未精确到秒
        strSql = "Select /*+ Rule*/ C.医嘱内容 From 病人门诊路径医嘱 B, 病人医嘱记录 C Where b.路径执行id In " & strIDSQL & _
                 " And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And c.医嘱状态 <> 4 And rownum<2 And to_date(to_char(c.开嘱时间 +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewIDs, dat导入时间)
        If rsTmp.RecordCount > 0 Then
            strIDs = ""
            For i = 1 To rsTmp.RecordCount
                If i > 10 Then strIDs = strIDs & "......": Exit For
                strIDs = strIDs & vbNewLine & rsTmp!医嘱内容
                rsTmp.MoveNext
            Next
            MsgBox "当前生成的项目存在已发送但未作废的医嘱：" & strIDs & vbNewLine & "请先作废医嘱后再执行取消。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '3.检查病历
    strSql = "Select /*+ Rule*/ 1 From 电子病历记录 Where 路径执行id In " & strIDSQL & _
             " And (完成时间 is not null or 打印人 is not null) And rownum<2  And 创建时间 >[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat导入时间)
    If rsTmp.RecordCount > 0 Then
        MsgBox "当前生成的项目对应的病历已签名或已打印，不能整体取消。", vbInformation, gstrSysName
        Exit Function
    End If
        
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(Split(strIDs, ","))
        strSql = "Zl_病人门诊路径生成_Delete(" & Split(strIDs, ",")(i) & ",0)"
        Call zlDatabase.ExecuteProcedure(strSql, "取消门诊路径项目")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    FuncDelAllItem = True

    If blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncDelItem()
'功能：取消生成当前选择的未执行的路径项目
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng执行ID As Long, lng项目ID As Long, blnMust As Boolean, lng天数 As Long
    Dim blnCancel As Boolean, strReason As String, blnTrans As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long

    With vsPath
        If .Cell(flexcpBackColor, .Row, .Col) = &HE0EFED Then
            MsgBox "该项目为必须生成但没有生成的项目，不用取消生成。", vbInformation, gstrSysName
            Exit Sub
        End If
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
    End With

    If mbln启用执行环节 Then
        '已经执行的不允许取消
        strSql = "Select 1 " & vbNewLine & _
                "From 病人门诊路径执行 A, 门诊路径项目 B" & vbNewLine & _
                "Where a.项目id = b.Id(+) And a.Id = [1] And a.执行时间 Is Not Null"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目已执行，不能取消。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '1.检查路径项目
    strSql = "Select b.执行方式,a.天数 From 病人门诊路径执行 a, 门诊路径项目 b Where a.项目ID = b.ID And a.ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then '临时项目，可以取消
        lng天数 = Val("" & rsTmp!天数)
        If rsTmp!执行方式 = 1 Then
            blnMust = True
        ElseIf rsTmp!执行方式 = 2 Or rsTmp!执行方式 = 4 Then  '至少一次或必须一次
            strSql = "Select 开始天数,结束天数 From 门诊路径阶段 Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.当前阶段ID)
            If Not IsNull(rsTmp!开始天数) Then
                If Not IsNull(rsTmp!结束天数) Then
                    blnMust = (lng天数 = Val("" & rsTmp!结束天数))    '是否最后一天
                    If blnMust Then '判断该项目之前有没有执行过(路径外项目除外)
                        strSql = "Select 1 From 病人门诊路径执行 Where 路径记录ID = [1] And 阶段ID = [2] And 项目ID = [3] And 天数<[4] And rownum<2"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, lng项目ID, lng天数)
                        If rsTmp.RecordCount > 0 Then blnMust = False
                    End If
                Else
                    blnMust = True  '单天
                End If
            End If
        End If
    End If

    '2.检查医嘱
    If lng执行ID <> 0 Then
        '即使已停止的医嘱也不允许删除，加59秒是因为开嘱时间未精确到秒
        strSql = "Select /*+ Rule*/ C.医嘱内容 From 病人门诊路径医嘱 B, 病人医嘱记录 C Where b.路径执行id=[1] " & _
                 " And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And c.医嘱状态 <> 4 And rownum<2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前生成的项目存在已发送但未作废的医嘱：" & rsTmp!医嘱内容 & vbNewLine & "请先作废医嘱后再执行取消。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    '3.必须生成的项目填写变异原因
    If blnMust Then
        '取消必须生成的项目时选择变异原因
        strSql = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 门诊变异常见原因 a,门诊变异常见原因 b" & _
                " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
                " Order by 分类,a.编码"
        vPoint = zlControl.GetCoordPos(vsPath.Hwnd, vsPath.CellLeft, vsPath.CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "门诊变异常见原因", True, , , True, True, True, _
                 vPoint.X, vPoint.Y, vsPath.RowHeight(vsPath.Row), blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "系统没有初始门诊变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            strReason = rsTmp!ID
        End If
    End If
    '4.检查病历
    strSql = "Select 1 From 电子病历记录 Where 路径执行id = [1] And (完成时间 is not null or 打印人 is not null) And rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目对应的病历已签名或已打印，不能取消。", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsPath
        If MsgBox("确实要取消路径项目""" & .TextMatrix(.Row, .Col) & """吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    End With
    If Not mbln启用执行环节 Then
        '判断是否已经评估
        strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            '强制取消评估，不检查权限
            MsgBox "本次生成的项目已评估，取消生成之前将自动取消评估。", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If strReason <> "" Then
        strSql = "Zl_病人门诊路径生成_Update(" & lng执行ID & ",'" & vsPath.TextMatrix(vsPath.Row, 0) & "',Null,NULL,NULL,NULL,'" & strReason & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "修改路径项目")
    End If
    strSql = "Zl_病人门诊路径生成_Delete(" & lng执行ID & "," & IIf(strReason <> "", "2", "0") & ")"

    Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAppendItemModify()
'功能：修改路径外项目
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lng执行ID As Long

    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P门诊路径应用), ";阶段评估;") = 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再修改路径外项目。", vbInformation, gstrSysName
            Exit Sub
        Else
            '取消评估
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能修改路径外项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If

    If frmPathAppendOut.ShowMe(mfrmParent, mPati, mPP, "", 2, "", lng执行ID) Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncAppendItem(ByVal bytUseType As Byte, Optional ByVal strItemType As String, Optional ByVal strAdviceIDs As String, _
                                Optional ByVal lng执行ID As Long, Optional ByVal datDate As Date) As Boolean
'功能：添加路径外项目(通过clsDockPath中的接口开放给医嘱部件调用)
'参数：bytUseType=0-直接添加,1-医嘱新开时添加
'       strItemType=医嘱接口调用时传入（当天最后一个项目的分类）
'       strAdviceIDs=医嘱接口调用时传入,医嘱序号
'       datDate =医嘱的开始执行日期（同一批路径外医嘱）
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim DatCur As Date
    Dim blnRefresh As Boolean

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P门诊路径应用), ";阶段评估;") = 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再添加路径外项目。", vbInformation, gstrSysName
            Exit Function
        Else
            '取消评估
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能添加路径外项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Function
            End If
        End If
    Else
        '未评估则检查是否是当天未评估，不是的话则提示是否要添加到最近一个阶段
        If bytUseType = 0 Then
            DatCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            If DatCur <> Format(mPP.当前日期, "yyyy-MM-dd") Then
                If MsgBox("你要添加路径外项目到""" & mPP.当前日期 & """?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call FuncSendItem
                    Exit Function
                End If
            End If
        End If
    End If

    If bytUseType = 0 Then
        With vsPath
            If .Row > 0 And .Row < .Rows - 2 Then strItemType = .TextMatrix(.Row, .FixedCols - 1) '最后一行是"路径评估"
        End With
    End If
    If frmPathAppendOut.ShowMe(mfrmParent, mPati, mPP, strItemType, bytUseType, strAdviceIDs, lng执行ID, datDate) Or blnRefresh Then
        FuncAppendItem = True
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImport(Optional ByVal blnask As Boolean, Optional blnImport As Boolean) As Boolean
'功能：导入路径

    If frmPathImportOut.ShowMe(mfrmParent, mPati, blnImport) Then
        FuncImport = True
        
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
        RaiseEvent RequestRefresh(mPP.病人路径状态)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncUnImport(Optional ByVal blnPrompt As Boolean = True)
'功能：取消导入,未生成路径时可取消导入
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str审核人 As String
    Dim lngPPStatus As Long

    '先检查是否有取消路径的权限
    If InStr(GetInsidePrivs(P门诊路径应用), ";取消导入;") = 0 Then
        str审核人 = zlDatabase.UserIdentify(Me, "没有取消导入权限需要审核。", glngSys, P门诊路径应用, "取消导入")
        If str审核人 = "" Then Exit Sub
    Else
        str审核人 = UserInfo.姓名
    End If
    strSql = "Select 1 From 病人门诊路径执行 Where 路径记录ID = [1] And rownum<2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then

        If MsgBox("当前阶段的路径项目已生成，首先将进行取消生成操作。" & vbCrLf & "你确实要取消该病人已导入的临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If

        '要刷新，以便重取路径信息(当前阶段等)
        If FuncDelAllItem(True, False) Then
            Call FuncUnImport(False)    '重新调用，再次检查
        End If
        Exit Sub
    ElseIf blnPrompt Then
        If MsgBox("你确实要取消该病人已导入的临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    lngPPStatus = mPP.病人路径状态
    
    gcnOracle.BeginTrans: blnTrans = True
    strSql = "Zl_病人门诊路径导入_Delete(" & mPP.病人路径ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消导入")
    '插入取消导入记录
    strSql = "Zl_病人门诊路径取消_Insert(" & mPati.病人ID & "," & mPati.挂号ID & ",'" & UserInfo.姓名 & "','" & str审核人 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "取消导入")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'功能：取消评估,未变异时才能取消（变异后自动结束，只能取消结束）
'参数：blnPrompt=是否弹出询问提示
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim lngPPStatus As Long

    On Error GoTo errH

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人在" & mPP.当前日期 & "还没有进行评估。", vbInformation, gstrSysName
        Exit Function
    End If

    If blnPrompt Then
        If MsgBox("你确定要取消第" & mPP.当前天数 & "天的评估吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If
    lngPPStatus = mPP.病人路径状态

    strSql = "Zl_病人门诊路径评估_Delete(" & mPP.病人路径ID & ", " & mPP.当前阶段ID & ",To_Date('" & mPP.当前日期 & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSql, "取消评估")
    FuncEvaluateCancel = True
    If blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncEvaluate()
'功能：阶段评估,只能对当前阶段的最后一次且执行了的进行
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim strTmp As String
    Dim bln补录 As Boolean
    Dim blnRefresh As Boolean
    Dim strDate  As String
    Dim lngPPStatus As Long

    '1.已评估的不能再评估 '已结束的不能再评估(禁用了菜单项)
    '2.只能对最后一次执行的记录进行评估(凡定义了评估指标的，必须评估后才能生成次日路径，没有定义的则不评估就执行生成)，因为评估为变异则可能结束路径
    '3.必须该阶段的所有项目都执行后才能评估
    On Error GoTo errH

    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        MsgBox "该病人在" & mPP.当前日期 & "已进行了评估。", vbInformation, gstrSysName
        Exit Sub
    End If

    '执行登记检查
    If Not CheckPathIsExecuted(blnRefresh) Then
        '强制刷新
        If blnRefresh Then
            Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
        End If
        Exit Sub
    End If

    lngPPStatus = mPP.病人路径状态

    If frmEvaluateOut.ShowMe(mfrmParent, 1, 1, mPati, mPP, , , , , , bln补录) Or bln补录 Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If

    If mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3 Then RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncReEvaluate()
'功能：修改评估，如果产生了后续阶段的项目，不能修改评估结果为变异来结束评估，在保存的存储过程中判断。
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln补录 As Boolean
    Dim lng阶段ID As Long
    Dim strSysDate As String
    Dim lng天数 As Long

    On Error GoTo errH


    strSql = "Select 1 From 病人门诊路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人在当前阶段还没有进行评估。", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmEvaluateOut.ShowMe(mfrmParent, 1, 2, mPati, mPP, , , , , , bln补录) Or bln补录 Then
        Call zlRefresh(mPati.病人ID, mPati.挂号ID, mPati.挂号NO, mPati.科室ID, mPati.病人状态)
    End If

    If mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3 Then RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewReport(ByVal str报告ID As String, ByVal lng医嘱ID As Long)
'功能：查阅报告

    '先判断是否可以继续操作
    If IsNumeric(str报告ID) Then
        If CheckEPRReport(Val(str报告ID), lng医嘱ID) = 2 Then
            If InStr(GetInsidePrivs(p住院医嘱下达), "查阅未完成报告") > 0 Then
                MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
            Else
                MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        RaiseEvent ViewEPRReport(Val(str报告ID), False)
    Else
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(0, str报告ID, Val(zlDatabase.GetPara("自动标记报告查阅状态", glngSys, p住院医嘱下达, "1")) = 1, mfrmParent)
    End If
End Sub

Public Function CheckEPRReport(ByVal lng报告ID As Long, ByVal lng医嘱ID As Long) As Integer
'功能：检查对应项目的报告填写情况
'参数：lng路径执行ID=病人门诊路径执行记录中的ID
'      lng报告ID=返回报告病历ID
'返回：
'      1-报告已填写完成(已签名,包括修订后签名,或已执行完成)
'      2-报告未填写完成(未签名,或修订后未签名,且未执行完成)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    On Error GoTo errH

    '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
    '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
    strSql = " Select 2 as 排序,医嘱ID,执行过程,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
            " Union ALL" & _
            " Select 排序,医嘱ID,执行过程,执行状态,发送时间" & _
            " From (" & _
                " Select 1 as 排序,B.医嘱ID,B.执行过程,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
                " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                    " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
                " Order by A.序号" & _
            " ) Where Rownum=1" & _
            " Order by 排序,发送时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng医嘱ID)
    If NVL(rsTmp!执行过程, 0) >= 5 Or NVL(rsTmp!执行状态, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mcolReason = Nothing
    SaveWinState Me, App.ProductName
    Set mobjPublicPACS = Nothing
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String, lngId As Long, i As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    lngId = fraMore.Tag
    If lngId = 0 Then
        Call zlCommFun.ShowTipInfo(0, strInfo)
    Else
        strSql = "Select " & IIf(mbln启用执行环节, "A.执行结果,A.执行说明,A.执行人,to_char(A.执行时间,'yyyy-mm-dd hh24:mi') as 执行时间,", "") & _
                " A.登记人,to_char(A.登记时间,'yyyy-mm-dd hh24:mi') as 登记时间 From 病人门诊路径执行 A,门诊路径项目 B Where A.项目ID=B.ID(+) And A.ID = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                For i = 0 To .Fields.count - 1
                    strInfo = strInfo & .Fields(i).Name & "：" & .Fields(i).Value & vbCrLf
                Next
            End With
            Call zlCommFun.ShowTipInfo(fraMore.Hwnd, strInfo, True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsFlow_DblClick()
    Dim lngPhaseID As Long
    If mPP.病人路径状态 = 0 And mPP.病人路径ID <> 0 Then   '导入失败
        Call frmEvaluateOut.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Else
        lngPhaseID = Val(vsFlow.ColData(vsFlow.Col))
        If lngPhaseID <> 0 Then
            Call frmPathSendOut.ShowMe(mfrmParent, 2, mPati, mPP, lngPhaseID, 0)
        ElseIf vsFlow.Col = 0 And mPP.路径ID <> 0 Then
            Call frmPathDefinition.ShowMe(mfrmParent, mPP.路径ID, 1)
        Else
            If vsFlow.Col = vsFlow.Cols - 2 And gstrDBUser = "ZLHIS" Then
                vsFlow.Editable = flexEDKbdMouse
            End If
        End If
    End If
End Sub

Private Sub vsFlow_LostFocus()
    If Not (mPP.病人路径状态 = 0 And mPP.病人路径ID <> 0) Then
        vsFlow.ForeColorSel = vsFlow.CellForeColor
    End If
End Sub

Private Sub vsFlow_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：用于测试时强制删除最后一天的项目(选中最后一个箭头后，输入DELA)
    Dim strPass As String, i As Long

    If vsFlow.Col = vsFlow.Cols - 2 Then
        strPass = UCase(vsFlow.EditText)
        vsFlow.EditText = ""
        If strPass = "DELA" Then
            If MsgBox("您确定要删除最后一天的所有项目吗？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            Call FuncDelPhaseItem
        End If
        vsFlow.Editable = flexEDNone
    End If
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or OldCol <> NewCol Then
        If fraMore.Visible Then fraMore.Visible = False

        If NewRow <> -1 And NewCol <> -1 And mblnUnChange = False Then
            '显示路径项目生成的医嘱清单
            Dim strTmp As String

            strTmp = vsPath.Cell(flexcpData, NewRow, NewCol)
            If InStr(strTmp, "|") > 0 Then
                Call UCAdvice.ShowAdvice(1, "", Val("" & Split(strTmp, "|")(0)), , , , , 1)
            Else
                Call UCAdvice.ShowAdvice(1, "", 0, , , , , 1)
            End If
        Else
            Call UCAdvice.ShowAdvice(1, "", 0, , , , , 1)
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45
End Sub

Private Sub vsPath_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub vsPath_DblClick()
    Dim lng项目ID As Long

    With vsPath
        If Trim(.TextMatrix(.Row, .Col)) <> "" And .Cell(flexcpData, .Row, .Col) <> "" And .Row <> .Rows - 1 Then
            lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
            If lng项目ID <> 0 Then
                Call frmPathItemEditOut.ShowView(mfrmParent, lng项目ID)
            End If
        End If
    End With
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngId As Long, lngRow As Long, lngCol As Long, lngItemID As Long
    
    With vsPath
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            lngRow = .MouseRow: lngCol = .MouseCol
            If .Cell(flexcpData, lngRow, lngCol) <> "" And lngRow <> .Rows - 1 Then
                lngId = Split(.Cell(flexcpData, lngRow, lngCol), "|")(0)
                lngItemID = Split(.Cell(flexcpData, lngRow, lngCol), "|")(1)
                If lngItemID = 0 Then
                    .ToolTipText = ""
                    Call zlCommFun.ShowTipInfo(.Hwnd, mcolReason("C" & lngId), True)      '路径外项目的添加原因
                Else
                    If .ToolTipText = "" Then .ToolTipText = "双击查看路径项目定义"
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                .ToolTipText = ""
            End If

            If lngId = 0 Then
                If imgMore.Visible Then fraMore.Visible = False
                fraMore.Tag = ""
            Else
                If lngRow = .Row And lngCol = .Col Then
                    fraMore.BackColor = .BackColorSel
                Else
                    fraMore.BackColor = .BackColor
                End If

                fraMore.Tag = lngId
                If fraMore.Visible = False Then fraMore.Visible = True
                fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - imgMore.Height - 30
                fraMore.Left = .Left + .ColPos(lngCol) + .ColWidth(lngCol) - imgMore.Width - 30
            End If
        Else
            If fraMore.Visible Then fraMore.Visible = False
        End If
    End With
End Sub

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim lng项目ID As Long

    '显示编辑菜单下面的内容
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置路径表流程图和清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    mlngFontSize = IIf(bytSize = 0, CON_SmallFontSize, CON_BigFontSize)

    vsFlow.Font.Size = mlngFontSize
    vsFlow.Redraw = flexRDDirect

    Call Grid.SetFontSize(vsPath, mlngFontSize)
    If vsPath.FixedRows > 1 Then vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45 '在要Draw之后才生效

    Call UCAdvice.SetVsAdviceFontSize(mlngFontSize)
End Sub

Private Sub MovePathItem(ByVal lngWay As Long)
'功能:当前单元格选中路径外项目时，可在路径外项目中上下移动
'参数:lngWay=1上移一行,-1下移一行(相当于下一行上移一行)
    Dim lngId       As Long
    Dim lngItemNum  As Long
    Dim arrSQL()    As Variant
    Dim i           As Integer
    Dim blnTran     As Boolean
    Dim blnDo As Boolean, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long

    blnDo = True: blnFind = False

    With vsPath
        Do While blnDo
            If .TextMatrix(.Row, .FixedCols - 1) <> .TextMatrix(.Row - lngWay, .FixedCols - 1) Or .Cell(flexcpData, .Row - lngWay, .Col) = "" Then
                MsgBox "项目内容:" & .TextMatrix(.Row, .Col) & vbCrLf & _
                       "已处于【" & .TextMatrix(.Row, .FixedCols - 1) & "】分类的" & IIf(lngWay > 0, "第一项", "最后一项"), vbInformation, gstrSysName
                blnDo = False: blnFind = False: Exit Do
            Else
                lngRow = .Row - lngWay: lngCol = .Col
                blnFind = True: Exit Do
            End If
        Loop
        '交换项目序号
        If blnFind Then
            arrSQL = Array()

            lngId = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人门诊路径序号_update(" & lngId & "," & lngItemNum & ")"

            lngId = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人门诊路径序号_update(" & lngId & "," & lngItemNum & ")"

            On Error GoTo errH
            gcnOracle.BeginTrans: blnTran = True
            For i = LBound(arrSQL) To UBound(arrSQL)
                zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
            Next i
            gcnOracle.CommitTrans: blnTran = False

            Call ClearPathItem(True)
            Call LoadPathItem

            '焦点移动
            .Row = lngRow: .Col = lngCol
            .ShowCell lngRow, lngCol
        End If
    End With
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncConvertPathTable() As VSFlexGrid
'功能:转换临床路径表单,用于应对特殊的打印需求 78233
'返回:转换后的路径表单
    Dim lngFirstSameCol As Long
    Dim lngLastRow As Long
    Dim i As Long, j As Long, k As Long

    Grid.CopyTo vsPath, vsPathPrint(0)
    vsPath.Redraw = flexRDNone
    vsPathPrint(0).MergeCol(0) = True
    vsPathPrint(0).MergeRow(0) = True

    With vsPathPrint(0)
        '一个阶段打印一列，不打印天数和日期
        For i = 2 To .Cols - 1
            If .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, i - 1) Then
            '相邻两列为同一阶段要合并
                For j = 1 To i - 1
                    If .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, j) Then
                        lngFirstSameCol = j '找到与当前列在同一阶段的首列
                        Exit For
                    End If
                Next
                'j-当前行,i-当前列
                For j = R2日期 + 1 To .Rows - 1
                    If .TextMatrix(j, i) <> "" And .TextMatrix(j, 0) <> "评估情况" Then
                        k = 0
                        For k = R2日期 + 1 To .Rows - 1
                            If .TextMatrix(j, i) = .TextMatrix(k, lngFirstSameCol) Then
                                Exit For
                            End If
                        Next
                        If k = .Rows Then
                            '新增
                            k = 0
                            lngLastRow = 0
                            For k = R2日期 + 1 To .Rows - 1
                                If .TextMatrix(j, 0) = .TextMatrix(k, 0) Then
                                    If .TextMatrix(k, lngFirstSameCol) = "" Then
                                        '同分类下空行新增
                                        .TextMatrix(k, lngFirstSameCol) = .TextMatrix(j, i)
                                        Exit For
                                    End If
                                    lngLastRow = k
                                End If
                            Next
                            If k = .Rows Then
                                '同分类最后一行新增一行
                                .AddItem "", lngLastRow + 1
                                .TextMatrix(lngLastRow + 1, lngFirstSameCol) = .TextMatrix(j, i)
                                .TextMatrix(lngLastRow + 1, 0) = .TextMatrix(j, 0)
                                .RowHeight(lngLastRow + 1) = .RowHeight(j)
                            End If
                        Else
                            '前一列存在，所以不处理
                        End If
                    End If
                Next
                '标记删除列
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        '删除列
        For i = .Cols - 1 To 0 Step -1
            If .ColHidden(i) = True And .ColWidth(i) = 0 Then
                '最后一行直接删除
                If i = .Cols - 1 Then
                    .Cols = .Cols - 1
                Else
                    '后面往前移
                    For k = i + 1 To .Cols - 1
                        For j = 0 To .Rows - 1
                            .TextMatrix(j, k - 1) = .TextMatrix(j, k)
                        Next
                    Next
                    .Cols = .Cols - 1
                End If
            End If
        Next
        '隐藏日期和天数
        .RowHidden(R1天数) = True: .RowHidden(R2日期) = True
        .RowHeight(R1天数) = 0: .RowHeight(R2日期) = 0
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效
    End With
    Set FuncConvertPathTable = vsPathPrint(0)
End Function

Private Function CheckPathIsExecuted(Optional ByRef blnRefresh As Boolean) As Boolean
'-------------------------------------------------------------------------------------------
'功能:检查当前阶段是否存在未执行的路径项目
'参数：=1 生成环节调用,=2 评估时候调用
'返回：F-存在未完成执行登记的路径项目,不允许生成或评估
'     T-不存在未完成执行登记的路径项目\不检查执行登记情况，允许生成或评估
'说明：1.护士生成时,当前未执行时不允许生成新的阶段,医生评估时,当前未执行不允许生成新的阶段
'      2.考虑医生要提前生成后续阶段 mbln启用不评估=true时,医生站生成时不检查是否完成执行登记，放在评估环节检查
'      3.护士站没有评估环节,需要每次生成都要检查前一次的执行登记情况
'-------------------------------------------------------------------------------------------
    Dim blnHave As Boolean          '不启用执行环节的检查
    Dim blnReturn As Boolean
    Dim blnExePath As Boolean
    Dim blnUnExe As Boolean         '用于标记没有执行路径权限且存在操作员执行的路径项目时,需要给予用户提示
    Dim strSql As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    blnHave = True                  '默认检查执行登记情况
    blnExePath = InStr(GetInsidePrivs(P门诊路径应用), ";执行路径;") > 0
    blnReturn = True
    blnHave = mbln启用执行环节

    If blnHave Then
        strSql = "Select Nvl(b.项目内容,a.项目内容) 项目内容 From 病人门诊路径执行 a,门诊路径项目 b " & vbNewLine & _
                        "Where a.项目id=b.id(+) And a.路径记录ID = [1] And a.阶段ID = [2] And a.日期 = [3] And a.执行时间 Is null "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            '医生场合评估环节检查,生成时不检查
            If rsTmp.RecordCount > 0 Then
                Call FuncGetRSTipInfo(rsTmp, "项目内容", strTmp)
                If blnExePath Then
                    If MsgBox("该病人还有未执行的项目:" & vbCrLf & strTmp & vbCrLf & "必须先执行。你现在要进行执行操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, 0, , 1) Then
                            blnRefresh = True
                        Else
                            blnReturn = False
                        End If
                    Else
                        blnReturn = False
                    End If
                Else
                    blnUnExe = True: blnReturn = False
                End If
            End If

            If blnUnExe Then
                '没有执行路径权限且存在操作员执行的路径项目时 , 需要给予用户提示
                MsgBox "该病人还有未执行的项目：" & vbCrLf & strTmp & vbCrLf & "必须执行后才能继续。", vbInformation, gstrSysName
            End If
        End If
    End If
    CheckPathIsExecuted = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncGetRSTipInfo(ByVal rsTmp As ADODB.Recordset, ByVal strFieldName As String, ByRef strTipInfo As String)
'功能:循环读取记录集中简要信息
    Dim i As Long

    strTipInfo = ""
    For i = 1 To rsTmp.RecordCount
        strTipInfo = IIf(i = 1, "", strTipInfo & vbCrLf) & rsTmp.Fields(strFieldName)
        If Len(strTipInfo) > 500 Then strTipInfo = strTipInfo & "…": Exit For
        rsTmp.MoveNext
    Next
End Sub

Private Sub GetPathCurrPhase(ByVal bytType As Byte, ByRef lng阶段ID As Long, ByRef lng天数 As Long, Optional ByRef strDate As String)
'--------------------------------------------------
'功能:获取批量执行登记或批量取消执行登记的当前阶段
'参数:bytType =1 批量执行,=2 批量取消执行
'--------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If bytType = 1 Then
        strSql = " Select *" & vbNewLine & _
                 " From (Select Distinct a.阶段id, a.日期, a.天数, Min(a.登记时间) As 登记时间" & vbNewLine & _
                 "       From 病人门诊路径执行 A, 门诊路径项目 B" & vbNewLine & _
                 "       Where a.项目id = b.Id(+) And a.路径记录id = [1] " & _
                 "             And a.执行时间 Is Null" & vbNewLine & _
                 "       Group By a.阶段id, a.日期, a.天数" & vbNewLine & _
                 "       Order By Min(a.登记时间))" & vbNewLine & _
                 " Where Rownum < 2"
    Else
        strSql = " Select *" & vbNewLine & _
                 " From (Select Distinct a.阶段id, a.日期, a.天数, Min(a.登记时间) As 登记时间" & vbNewLine & _
                 "       From 病人门诊路径执行 A, 门诊路径项目 B" & vbNewLine & _
                 "       Where a.项目id = b.Id(+) And a.路径记录id = [1] " & _
                 "             And a.执行时间 Is Not Null" & vbNewLine & _
                 "       Group By a.阶段id, a.日期, a.天数" & vbNewLine & _
                 "       Order By Min(a.登记时间) Desc )  " & vbNewLine & _
                 " Where Rownum < 2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then
        lng阶段ID = Val(NVL(rsTmp!阶段ID))
        strDate = rsTmp!日期 & ""
        lng天数 = Val(NVL(rsTmp!天数))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableChange(ByRef vsBody As VSFlexGrid, ByVal lngPageCOL As Long, Optional vsHead As VSFlexGrid)
'功能:将打印表单转换成固定列,便于打印输出
'主要解决问题: 当阶段行高超过打印有效范围时要求下一页继续补打当前阶段剩余行
'              每一阶段的字体自动缩放行间距,剔除空白行。
'参数: 出参:vsBody打印表体
'      入参:lngPageCOL 打印列数(不含固定列)
    Dim lngRow As Long
    Dim lngCol As Long

    On Error Resume Next
    
    Load vsPathPrint(1)
    
    Err.Clear: On Error GoTo 0

    With vsPathPrint(1)
        '清空
        .Rows = 0
        .Cols = 0

        If lngPageCOL = 0 Then Exit Sub
        If (vsBody.Cols - vsBody.FixedCols) Mod lngPageCOL <> 0 Then Exit Sub
        
        .Rows = ((vsBody.Cols - vsBody.FixedCols) / lngPageCOL) * vsBody.Rows
        .Cols = vsBody.FixedCols + lngPageCOL
        .FixedCols = vsBody.FixedCols
        .FixedRows = vsBody.FixedRows

        '从vsBody将数据复制到vsPathPrint(1)
        '固定列
        For lngCol = 0 To .FixedCols
            lngRow = 0
            Do
                '（原本是固定行转换成非固定行时需要特殊标记便于打印部件识别）
                If lngRow Mod vsBody.Rows < vsBody.FixedRows And lngRow >= vsBody.FixedRows And lngCol = 0 Then
                    .RowData(lngRow) = UCase("FIXEDROW")
                End If
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        '非固定列
        For lngCol = .FixedCols To (.FixedCols + lngPageCOL) - 1
            lngRow = 0
            Do
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, (lngPageCOL * (lngRow \ vsBody.Rows)) + lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next

        '清空多列都是空白的行
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To lngPageCOL
                If .RowData(lngRow) = UCase("FIXEDROW") Then
                    .Cell(flexcpAlignment, lngRow, lngCol, lngRow, .Cols - 1) = flexAlignCenterCenter
                    Exit For
                ElseIf .TextMatrix(lngRow, 0) = "医生签名" Then
                    Exit For
                ElseIf .TextMatrix(lngRow, lngCol) <> "" Then
                    Exit For
                ElseIf lngCol = lngPageCOL Then
                    '记录下要删除的空白行
                   .RemoveItem lngRow
                   lngRow = lngRow - 1  '删除一行,下一行向上填充
                End If
            Next
            If lngRow = .Rows - 1 Then Exit For
        Next
        '显示风格处理
        .MergeCol(0) = True
        '设定字体，宽度
        .FontSize = IIf(mlngFontSize = 0, CON_SmallFontSize, mlngFontSize) '路径跟踪批量打印mlngFontSize=0
        '显示风格
        .Cell(flexcpAlignment, 0, .FixedCols, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
    End With
    Set vsBody = vsPathPrint(1)
End Sub

Private Sub FuncPathCellCopy(ByRef vsSource As VSFlexGrid, ByRef vsCopy As VSFlexGrid, _
        ByVal lngSourRow As Long, ByVal lngSourCol As Long, ByVal lngCopyRow As Long, ByVal lngCopyCol As Long)
'功能:复制单元格
'参数：vsSource-被Copy的表单
'      vsCopy-copy后的表单
'      lngSourRow ,lngSourCol 被Copy的表单对应行和列
'      lngCopyRow，lngCopyCol Copy后表单对应行和列
    With vsCopy
        .Cell(flexcpText, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpText, lngSourRow, lngSourCol)
        .Cell(flexcpAlignment, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpAlignment, lngSourRow, lngSourCol)
        .Cell(flexcpBackColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpBackColor, lngSourRow, lngSourCol)
        .Cell(flexcpForeColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpForeColor, lngSourRow, lngSourCol)
        .Cell(flexcpPicture, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpPicture, lngSourRow, lngSourCol)
    End With
End Sub

Private Sub OutLogModi()
    Dim colSQL As New Collection, i As Long, blnTrans As Boolean

    Call frmPathOutLogOut.ShowMe(mfrmParent, mPati.病人ID, mPati.挂号ID, 2, colSQL, mPP.路径ID, mPP.病人路径ID)

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    '执行出径登记表的SQL
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "修改出径登记表")
    Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
