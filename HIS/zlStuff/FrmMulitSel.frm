VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMulitSel 
   BorderStyle     =   0  'None
   Caption         =   "材料选择器"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ControlBox      =   0   'False
   Icon            =   "FrmMulitSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkContinue 
      Caption         =   "连续选择(&M)"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pic选定区 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4440
      Width           =   4815
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "FrmMulitSel.frx":0E42
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picOK 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3240
         Picture         =   "FrmMulitSel.frx":1184
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "选定"
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf选定 
         Height          =   2085
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   4275
         _cx             =   7541
         _cy             =   3678
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   32
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMulitSel.frx":14C6
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
      End
      Begin VB.Label lbl选定 
         BackColor       =   &H00FFEDDD&
         Caption         =   "选定卫材"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   3210
      Left            =   3285
      TabIndex        =   2
      Top             =   735
      Visible         =   0   'False
      Width           =   2700
      _cx             =   4762
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMulitSel.frx":192F
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   2
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHead 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8040
      _cx             =   14182
      _cy             =   4445
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483637
      BackColorAlternate=   -2147483628
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483644
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMulitSel.frx":197C
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
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   30
         Picture         =   "FrmMulitSel.frx":1A51
         Top             =   30
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBatch 
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Width           =   8025
      _cx             =   14155
      _cy             =   1852
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483637
      BackColorAlternate=   -2147483628
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483644
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMulitSel.frx":1FDB
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
      Begin VB.Image imgBatch 
         Height          =   240
         Left            =   30
         Picture         =   "FrmMulitSel.frx":20B0
         Top             =   30
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   6480
      Top             =   4680
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
            Picture         =   "FrmMulitSel.frx":263A
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMulitSel.frx":298C
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMulitSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--输入参数--
Private mintEditState As Integer                 '编辑状态(1-入库;2-出库)
Private mlng源库房ID As Long                     '源库房ID
Private mlng目库房ID As Long                     '目库房ID
Private mlng使用部门ID As Long                   '使用部门ID
Private mlng供应商ID As Long                     '供应商ID
Private mstrInput As String                      '输入字串
Private mobjOut As Form                          '使用本程序的窗体（必须提供一个公共记录集，用以返回）
Private mblnSelect As Boolean                    '是否允许选择

Private mblnStartUp As Boolean                   '启动成功
Private mblnFirstStart As Boolean                '第一次启动
Private mrsUnit As New ADODB.Recordset           '单位
Private mstrUnit As String                       '单位名称
Private mstrUnitString As String                 'SQL字串
Private mintStockCheck As Integer                '库存检测
Private mstrFindStyle As String                  '匹配方式
Private mbln盘点单 As Boolean                    '盘点单据标志
Private mbln空批次 As Boolean                    '是否增加空批次供输入
Private mblnCheck As Boolean                     '是否遵守出库原则(非批次或时价卫生材料)
Private mblnPrice As Boolean                     '是否允许时价或批次卫材零出库
Private mstrCode As String                       '条码
Private mblnTrackUsing As Boolean                '跟踪在用参数

'本程序使用记录集
Private mrsData As New ADODB.Recordset           '卫材用途分类
Private mrsCard As New ADODB.Recordset           '卫材卡片
Private mrsStock As New ADODB.Recordset          '卫材规格
Private mstrTittle As String                     '选择器名称

'返回记录集
Private mrsReturn As ADODB.Recordset             '返回记录集(卫材信息所有列,卫材目录所有列,卫材库存所有列)
Private mint库房 As Integer                      '1-卫材库;2-在用;3-制剂室
Private mint分批 As Integer                      '0-不分批;1-库房分批;2-在用分批;3-卫材库在用分批
Private mbln时价 As Boolean                      '时价
Private mblnStock As Boolean
Private mstrCardSortBy As String                 '卫材卡片排序列
Private mstrPhysicSortBy As String               '卫材规格排序列
Private mlngCardRow As Long
Private mlngPhysicRow As Long
Private mlngLastSelect材料ID As Long             '上次选择的材料ID（用于是否刷新）
Private mbln散装单位 As Boolean
Private mbln只显示在用物资 As Boolean
Private mbln仅显示库存物资 As Boolean
Private mbln显示批次 As Boolean                 '是否显示批次列表， true-显示批次，false-不能显示批次,主要是出库业务用来判断是否是特殊业务，如申领不明确批次模式则不显示批次列表
Private mlngModule As Long
Private mblnSelectSucess As Boolean
Private mbln盘无存储库房材料 As Boolean
Private mblnCostView As Boolean                 '查看成本价相关信息 true-允许查看 false-不允许查看
Private mblnProvider As Boolean                 '查看上次供应商相关信息 true-允许查看 false-不允许查看
Private mstrPrivs As String                     '操作员权限
Private mbln是否过滤 As Boolean

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'调用get可用库存后，返回的可用数量，实际数量，实际金额及实际差价
Private msin可用数量 As Single
Private msin实际数量 As Single
Private msin实际金额 As Single
Private msin实际差价 As Single
Private Const MFRM_MIN_WIDTH = 8040
Private Const MFRM_MIN_HEIGHT = 3630
Private mstr盘点时间 As String

'--公共--
'Private Const strFormat As String = "'999999999990.9999'"
Private Type WinLocate
    Left As Double
    Top As Double
    lngTxtW As Long
    lngTxtH As Long
End Type
Private mWindowPosition As WinLocate           '窗体位置

Public Property Get In_编辑状态() As Integer
    In_编辑状态 = mintEditState
End Property

Public Property Let In_编辑状态(ByVal vNewValue As Integer)
    mintEditState = vNewValue
End Property

Public Property Get In_源库房() As Long
    In_源库房 = mlng源库房ID
End Property

Public Property Let In_源库房(ByVal vNewValue As Long)
    mlng源库房ID = vNewValue
End Property

Public Property Get In_字串() As String
    In_字串 = mstrInput
End Property

Public Property Let In_字串(ByVal vNewValue As String)
    mstrInput = vNewValue
End Property

Public Property Get In_目库房() As Long
    In_目库房 = mlng目库房ID
End Property

Public Property Let In_目库房(ByVal vNewValue As Long)
    mlng目库房ID = vNewValue
End Property

Public Property Get In_部门() As Long
    In_部门 = mlng使用部门ID
End Property

Public Property Let In_部门(ByVal vNewValue As Long)
    mlng使用部门ID = vNewValue
End Property

Public Property Let In_MainFrm(ByVal vNewValue As Form)
    Set mobjOut = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntMain As Integer = 1, Optional ByVal BlnSetHeader As Boolean = False)
    Dim intCol As Integer
    
    '设置各列表控件的格式
    Select Case IntMain
    Case 1
        With vsHead
            For intCol = 0 To .Cols - 1
                .FixedAlignment(intCol) = flexAlignCenterCenter
                .ColKey(intCol) = UCase(.TextMatrix(0, intCol))
                If InStr(1, .ColKey(intCol), "ID") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    '.coldata(i):1-固定,-1-不能选,0-可选
                    .ColData(intCol) = -1
                ElseIf InStr(1, .ColKey(intCol), "数量") > 0 Or _
                   (InStr(1, .ColKey(intCol), "价") > 0 And .ColKey(intCol) <> "时价") Then
                    .ColAlignment(intCol) = flexAlignRightCenter
                    .ColWidth(intCol) = 1000
                    '.coldata(i):1-固定,-1-不能选,0-可选
                    .ColData(intCol) = 0
                Else
                    .ColData(intCol) = 0
                    Select Case .ColKey(intCol)
                    Case "有效期", "散装", "包装", "系数", "一次性材料", _
                        "无菌性材料", "灭菌效期", "灭菌失效期", "库房分批", "在用分批", "时价"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                    Case "编码", "通用名称"
                        '.coldata(i):1-固定,-1-不能选,0-可选
                        .ColAlignment(intCol) = flexAlignLeftCenter
                        .ColData(intCol) = 1
                    Case Else
                        .ColAlignment(intCol) = flexAlignLeftCenter
                    End Select
                End If
            Next
            '自动调用列宽
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            .Row = 1
            '恢复列宽
             
            zl_vsGrid_Para_Restore mlngModule, vsHead, mstrTittle, "规格信息", False
            
            If mlngModule = 1725 Or .ColWidth(.ColIndex("上次供应商")) = 0 Then .ColWidth(.ColIndex("上次供应商")) = IIf(mblnProvider = True, 1300, 0)
            If mlngModule = 1725 Or .ColHidden(.ColIndex("上次供应商")) = True Then .ColHidden(.ColIndex("上次供应商")) = Not mblnProvider
            If .ColWidth(.ColIndex("指导差价率")) <> 0 Then .ColWidth(.ColIndex("指导差价率")) = 0
            If .ColHidden(.ColIndex("指导差价率")) = False Then .ColHidden(.ColIndex("指导差价率")) = True
            
            If mblnCostView = False Then
                If .ColHidden(.ColIndex("库存差价")) = False Then .ColHidden(.ColIndex("库存差价")) = True
                If .ColHidden(.ColIndex("最新成本价")) = False Then .ColHidden(.ColIndex("最新成本价")) = True
                If .ColHidden(.ColIndex("指导批发价")) = False Then .ColHidden(.ColIndex("指导批发价")) = True
                .ColWidth(.ColIndex("库存差价")) = 0
                .ColData(.ColIndex("库存差价")) = -1
                .ColWidth(.ColIndex("最新成本价")) = 0
                .ColData(.ColIndex("最新成本价")) = -1
                .ColWidth(.ColIndex("指导批发价")) = 0
                .ColData(.ColIndex("指导批发价")) = -1
            Else
                If .ColHidden(.ColIndex("库存差价")) = True Then .ColHidden(.ColIndex("库存差价")) = False
                If .ColHidden(.ColIndex("最新成本价")) = True Then .ColHidden(.ColIndex("最新成本价")) = False
                If .ColHidden(.ColIndex("指导批发价")) = True Then .ColHidden(.ColIndex("指导批发价")) = False
                If .ColWidth(.ColIndex("库存差价")) = 0 Then .ColWidth(.ColIndex("库存差价")) = 1000
                If .ColWidth(.ColIndex("最新成本价")) = 0 Then .ColWidth(.ColIndex("最新成本价")) = 1000
                If .ColWidth(.ColIndex("指导批发价")) = 0 Then .ColWidth(.ColIndex("指导批发价")) = 1000
            End If
        End With
    Case 0
        With vsBatch
            DoEvents
            For intCol = 0 To .Cols - 1
                .FixedAlignment(intCol) = flexAlignCenterCenter
                .ColKey(intCol) = UCase(.TextMatrix(0, intCol))
                '.coldata(i):1-固定,-1-不能选,0-可选
                If InStr(1, .ColKey(intCol), "ID") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    .ColData(intCol) = -1
                ElseIf InStr(1, .ColKey(intCol), "批次") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    .ColData(intCol) = 0
                ElseIf InStr(1, .ColKey(intCol), "数量") > 0 Or _
                    InStr(1, .ColKey(intCol), "价") > 0 Then
                    .ColAlignment(intCol) = flexAlignRightCenter
                    .ColWidth(intCol) = 1000
                    .ColData(intCol) = 0
                Else
                    Select Case .ColKey(intCol)
                    Case "批号"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                        .ColData(intCol) = 1
                    Case "生产日期"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                        .ColData(intCol) = 0
                    Case Else
                        .ColAlignment(intCol) = flexAlignLeftCenter
                        .ColData(intCol) = 0
                    End Select
                End If
            Next
            '自动调用列宽
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            If .Rows >= 2 Then .Row = 1
            '恢复列宽
            zl_vsGrid_Para_Restore mlngModule, vsBatch, mstrTittle, "批次信息", False
            If mblnCostView = False Then
                .ColWidth(.ColIndex("库存差价")) = 0
                .ColWidth(.ColIndex("上次购价")) = 0
                .ColWidth(.ColIndex("成本价")) = 0
                .ColData(.ColIndex("库存差价")) = -1
                .ColData(.ColIndex("上次购价")) = -1
                .ColData(.ColIndex("成本价")) = -1
            End If
        End With
    End Select
End Sub

Private Sub OnCancel()
    Unload Me
    Exit Sub
End Sub

Private Sub OnSelect()
    Dim blnValid As Boolean
    
    If chkContinue.Value = 0 Then
        If In_编辑状态 = 2 Then If CheckData = False Then Exit Sub
        '检查分批属性与库存数据是否一致
        If In_编辑状态 = 2 Then
            blnValid = 检查库存数据(mlng源库房ID, mlngLastSelect材料ID)
        Else
            blnValid = 检查库存数据(mlng目库房ID, mlngLastSelect材料ID)
        End If
        If Not blnValid Then
            MsgBox "发现该卫材在当前库房中的库存记录存在错误（可能是基础数据设置错误，请检查当前库房的部门性质及该卫材的分批属性）！", vbInformation, gstrSysName
            Exit Sub
        End If
        '组装记录集
        If CombinateRec = False Then Exit Sub
        mblnSelectSucess = True
        Unload Me
        Exit Sub
    Else
        If CombinateRec = False Then Exit Sub
        mblnSelectSucess = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub chkContinue_Click()

    If vsf选定.Rows > 2 And chkContinue.Value = 0 Then
        If MsgBox("已经有选定卫材存在，取消“连续选择”将清除已选定的卫材，你确定吗？" _
            , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsf选定.Rows = 1
            vsf选定.Rows = 2
            lbl选定.Caption = "选定药品"
        Else
            chkContinue.Value = 1
            Exit Sub
        End If
        
    End If

    pic选定区.Visible = chkContinue.Value = 1
    picSplit02_S.Visible = chkContinue.Value = 1
    Form_Resize
    
    
    If chkContinue.Value = 0 Then
        pic选定区.Tag = "展开"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
    End If
    
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If

    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    mblnStartUp = False
    mblnFirstStart = False
    Call vsBatch_LostFocus:
    Call vsHead_LostFocus
    '取售价单位
    mstrUnit = ""
    mstrFindStyle = IIf(gstrMatchMethod = "0", "%", "")
    mstrUnitString = ""
    mintStockCheck = 0
    mlngLastSelect材料ID = 0
    vsBatch.Visible = (In_编辑状态 = 2)
    
    pic选定区.Visible = False
    picSplit02_S.Visible = False
    chkContinue.Visible = mbln是否过滤 = False
    pic选定区.Tag = "展开"
    
    
    On Error GoTo ErrHandle
    '初始化记录集
    InitRec
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    If mlngModule = 1725 Then
        mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "查看供应商")
    Else
        mblnProvider = True
    End If
    
    If mstrInput = "" Then Exit Sub
    If mobjOut Is Nothing Then
        MsgBox "请指定主窗体！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '定位
    With mWindowPosition
        Me.Left = .Left
        Me.Top = .Top
    End With
    
    '提取当前库存控制参数
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"
    Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
    
    If Not mrsUnit.EOF Then
        mintStockCheck = mrsUnit!库存检查
    End If
    
    '检查源库房是否为卫材库
    If mlng源库房ID <> 0 Then
        mint库房 = 3
        gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '发料部门' Or 工作性质 like '%制剂室') And 部门id=[1]"
        Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
        If mrsUnit.EOF Then
            gstrSQL = "select 部门ID from 部门性质说明 where 工作性质 In ('卫材库', '虚拟库房') And 部门id=[1]"
            Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
            If Not mrsUnit.EOF Then mint库房 = 1
        Else
            mint库房 = 2
        End If
    End If
    
    '读出该使用的单位级数
    If mlng使用部门ID <> 0 Then
        If mbln散装单位 Then
            mstrUnitString = "/1"
        Else
            mstrUnitString = "/nvl(换算系数,1)"
        End If
    End If
    
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_成本价, True)
        .FM_金额 = GetFmtString(IIf(mbln散装单位, 0, 1), g_金额, True)
        .FM_零售价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_售价, True)
        .FM_数量 = GetFmtString(IIf(mbln散装单位, 0, 1), g_数量, True)
    End With
    
    mblnStartUp = RefreshData
    
    On Error Resume Next
    If mrsCard.RecordCount = 1 Then
        If Not (((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And mblnPrice) Or mintEditState = 1 Then OnSelect
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    mblnFirstStart = True
    
    With vsHead
        .Top = chkContinue.Height + chkContinue.Top * 2
        .Height = IIf(vsBatch.Visible = False, Me.ScaleHeight - .Top, (Me.ScaleHeight - .Top) / 2)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
        .Width = Me.ScaleWidth
        
    End With
    With vsBatch
        .Top = vsHead.Height + vsHead.Top
        .Width = vsHead.Width
        .Height = Me.ScaleHeight - (vsHead.Height + vsHead.Top)
        .Left = ScaleLeft
    End With
    
    If picSplit02_S.Visible Then
        '设置分界线的top
        If vsBatch.Visible Then '批次可见
            vsBatch.Height = vsBatch.Height - (lbl选定.Height + picSplit02_S.Height)
            picSplit02_S.Top = vsBatch.Top + vsBatch.Height
        Else
            vsHead.Height = vsHead.Height - (lbl选定.Height + picSplit02_S.Height)
            picSplit02_S.Top = vsHead.Top + vsHead.Height
        End If
        picSplit02_S.Left = 0
        picSplit02_S.Width = vsHead.Width
    End If
    
    If pic选定区.Visible Then
        pic选定区.Width = vsHead.Width
        pic选定区.Height = lbl选定.Height
        
        pic选定区.Top = picSplit02_S.Top + picSplit02_S.Height
        
        With lbl选定
            .Top = 0
            .Left = 0
            .Width = pic选定区.Width
        End With
        With picUpDown01
            .Left = pic选定区.Width - .Width
            .Top = 0
        End With
        
        With picOK
            .Left = picUpDown01.Left - .Width
            .Top = 0
        End With
        
        With vsf选定
            .Top = lbl选定.Height
            .Left = 0
            .Width = lbl选定.Width
        End With
        
        If pic选定区.Tag = "收缩" Then
            pic选定区.Tag = "展开"
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            picSplit02_S.MousePointer = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrCode = ""
    SaveWinState Me
    zl_vsGrid_Para_Save mlngModule, vsHead, mstrTittle, "规格信息", False
    zl_vsGrid_Para_Save mlngModule, vsBatch, mstrTittle, "批次信息", False
End Sub

 

Private Sub imgBatch_Click()
    Call LoadFulltoColSel(True)
End Sub


Private Sub picOK_Click()
    OnSelect
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit02_S
        If .Top + y < vsHead.Top + 500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 500 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    pic选定区.Move pic选定区.Left, pic选定区.Top + y, pic选定区.Width, pic选定区.Height - y
    vsf选定.Move vsf选定.Left, vsf选定.Top, pic选定区.Width, vsf选定.Height - y
    
End Sub

Private Sub picUpDown01_Click()
    If pic选定区.Tag = "展开" Then
        pic选定区.Tag = "收缩"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
        picSplit02_S.MousePointer = 7

        picSplit02_S.Top = picSplit02_S.Top - vsf选定.Height
        pic选定区.Top = pic选定区.Top - vsf选定.Height
        pic选定区.Height = pic选定区.Height + vsf选定.Height
    Else
        pic选定区.Tag = "展开"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
        
        Form_Resize
    End If
End Sub

Private Sub vsBatch_Click()
'    Dim StrHeader As String
'    Dim intCol As Integer
'    '实现列排序
'    With vsBatch
'        If .MouseRow <> 0 Then Exit Sub
'        If mrsStock.EOF Then Exit Sub
'
'        StrHeader = .TextMatrix(0, .MouseCol)
'        Set .DataSource = Nothing
'        If Mid(mstrPhysicSortBy, 2) = StrHeader Then
'            mstrPhysicSortBy = IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
'            mrsStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
'        Else
'            mstrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
'            mrsStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
'        End If
'        Set .DataSource = mrsStock
'
'        For intCol = 0 To .Cols - 1
'            .ColAlignmentFixed(intCol) = 4
'        Next
'
'        Call SetFormat(0, False)
'    End With
End Sub

Private Sub vsBatch_DblClick()
    With mrsStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    
    If mblnSelect Then
        If chkContinue.Value = 1 Then
            FillVSF选定
            Exit Sub
        End If
        
        OnSelect
    End If
End Sub
Private Sub vsBatch_GotFocus()
    With vsBatch
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
        .BackColorSel = &H8000000D
    End With
End Sub
Private Sub vsBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then vsBatch_DblClick
End Sub

Private Sub vsBatch_LostFocus()
    With vsBatch
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        .BackColorSel = &H8000000A
    End With
End Sub

  


Private Sub vsf选定_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsf选定.Rows > 2 Then
            If vsf选定.Row <> vsf选定.Rows - 1 Then vsf选定.RemoveItem vsf选定.Row
            If vsf选定.Rows = 2 Then
                lbl选定.Caption = "选定卫材"
            Else
                lbl选定.Caption = "选定卫材（" & vsf选定.Rows - 2 & "条）"
            End If
        End If
    End If
End Sub

Private Sub vsHead_DblClick()
    If mrsCard.EOF Then Exit Sub
    If mrsCard.RecordCount = 0 Then Exit Sub
    
    If mblnSelect Then '允许选择才进入
        If chkContinue.Value = 1 Then
            FillVSF选定
            Exit Sub
        End If

        OnSelect
    Else
        MsgBox "该卫材没有库存，不能继续操作！", vbInformation, gstrSysName
    End If
End Sub


Private Sub FillVSF选定()
    Dim blnEof As Boolean         '是否存在批次库存
    Dim i As Integer
    Dim blnValid    As Boolean
    
    '检查药卫材重复
    If chkContinue.Value = 1 Then
        For i = 1 To vsf选定.Rows - 2
            If Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("材料ID"))) = Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID"))) Then
                If vsBatch.Visible Then
                    If vsf选定.TextMatrix(i, vsf选定.ColIndex("批次")) = vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批次")) Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If
        Next
    End If
    
    If In_编辑状态 = 2 Then If CheckData = False Then Exit Sub
        
    '检查分批属性与库存数据是否一致
    If In_编辑状态 = 2 Then
        blnValid = 检查库存数据(mlng源库房ID, mlngLastSelect材料ID)
    Else
        blnValid = 检查库存数据(mlng目库房ID, mlngLastSelect材料ID)
    End If
    
    If Not blnValid Then
        ShowMsgBox "发现该卫材在当前库房中的库存记录存在错误（可能是基础数据设" & vbCrLf & "置错误，请检查当前库房的部门性质及该卫材的分批属性）！"
        Exit Sub
    End If
    
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "材料ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID")))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mbln显示批次 = True Then '只有显示批次的情况下才需要做如下操作
            If ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And In_编辑状态 = 2 Then
                With mrsStock
                    If .RecordCount <> 0 Then .MoveFirst
                    .Find "批次=" & Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批次")))
                    If .EOF Then
                        blnEof = True
                        If mblnPrice Then
                            MsgBox "发生内部错误！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End With
            End If
        End If
    End With
    
    
    '装数据写入记录集，供其它窗体使用
    With vsf选定
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 2, .ColIndex("材料ID")) = mrsCard!材料ID
        .TextMatrix(.Rows - 2, .ColIndex("诊疗id")) = mrsCard!诊疗id
        .TextMatrix(.Rows - 2, .ColIndex("分类id")) = mrsCard!分类id
        
        .TextMatrix(.Rows - 2, .ColIndex("编码")) = mrsCard!编码
        .TextMatrix(.Rows - 2, .ColIndex("名称")) = zlStr.Nvl(mrsCard!通用名称)
        .TextMatrix(.Rows - 2, .ColIndex("商品名")) = zlStr.Nvl(mrsCard!商品名)
        .TextMatrix(.Rows - 2, .ColIndex("规格")) = mrsCard!规格
        .TextMatrix(.Rows - 2, .ColIndex("产地")) = "" & mrsCard!产地
        .TextMatrix(.Rows - 2, .ColIndex("售价")) = zlStr.Nvl(mrsCard!售价, 0)
        .TextMatrix(.Rows - 2, .ColIndex("散装单位")) = mrsCard!散装
        .TextMatrix(.Rows - 2, .ColIndex("换算系数")) = mrsCard!系数
        .TextMatrix(.Rows - 2, .ColIndex("包装单位")) = mrsCard!包装
        .TextMatrix(.Rows - 2, .ColIndex("最大效期")) = "" & mrsCard!有效期
        .TextMatrix(.Rows - 2, .ColIndex("灭菌效期")) = "" & mrsCard!灭菌效期
        .TextMatrix(.Rows - 2, .ColIndex("灭菌失效期")) = "" & mrsCard!灭菌失效期
        .TextMatrix(.Rows - 2, .ColIndex("一次性材料")) = IIf(mrsCard!一次性材料 = "是", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("无菌性材料")) = IIf(mrsCard!无菌性材料 = "是", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("库房分批")) = IIf(mrsCard!库房分批 = "是", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("在用分批")) = IIf(mrsCard!在用分批 = "是", 1, 0)
        
        .TextMatrix(.Rows - 2, .ColIndex("时价")) = IIf(mrsCard!时价 = "是", 1, 0)

        '出库且分批
        If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) Then
            If mbln显示批次 = True Then '只有显示批次情况下才需要做如下操作，否则无意义
                If vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批号")) = "新增批次卫生材料" Then
                    .TextMatrix(.Rows - 2, .ColIndex("批次")) = -1
                Else
                    If Not blnEof Then
                        .TextMatrix(.Rows - 2, .ColIndex("批次")) = Val(mrsStock!批次)
                        .TextMatrix(.Rows - 2, .ColIndex("批号")) = "" & mrsStock!批号
                        .TextMatrix(.Rows - 2, .ColIndex("效期")) = "" & mrsStock!失效期
                        .TextMatrix(.Rows - 2, .ColIndex("灭菌失效期")) = "" & mrsStock!灭菌失效期
                        .TextMatrix(.Rows - 2, .ColIndex("产地")) = "" & mrsStock!产地
                        .TextMatrix(.Rows - 2, .ColIndex("生产日期")) = "" & mrsStock!生产日期
                        .TextMatrix(.Rows - 2, .ColIndex("批准文号")) = "" & mrsStock!批准文号
                        .TextMatrix(.Rows - 2, .ColIndex("供药单位ID")) = "" & mrsStock!上次供应商id
                        .TextMatrix(.Rows - 2, .ColIndex("可用数量")) = IIf(IsNull(mrsStock!可用数量), 0, mrsStock!可用数量)
                        .TextMatrix(.Rows - 2, .ColIndex("实际数量")) = IIf(IsNull(mrsStock!库存数量), 0, mrsStock!库存数量)
                        .TextMatrix(.Rows - 2, .ColIndex("实际金额")) = IIf(IsNull(mrsStock!库存金额), 0, mrsStock!库存金额)
                        .TextMatrix(.Rows - 2, .ColIndex("实际差价")) = IIf(IsNull(mrsStock!库存差价), 0, mrsStock!库存差价)
                        If Not mblnStock Then Call Get可用库存(.TextMatrix(.Rows - 2, .ColIndex("材料ID")), .TextMatrix(.Rows - 2, .ColIndex("批次")))
                    End If
                End If
            Else
                If Not mblnStock Then Call Get可用库存(mrsCard!材料ID, 0)
            End If
        Else
        '入库或不分批
            .TextMatrix(.Rows - 2, .ColIndex("可用数量")) = IIf(IsNull(mrsCard!可用数量), 0, mrsCard!可用数量)
            .TextMatrix(.Rows - 2, .ColIndex("实际数量")) = IIf(IsNull(mrsCard!库存数量), 0, mrsCard!库存数量)
            .TextMatrix(.Rows - 2, .ColIndex("实际金额")) = IIf(IsNull(mrsCard!库存金额), 0, mrsCard!库存金额)
            .TextMatrix(.Rows - 2, .ColIndex("实际差价")) = IIf(IsNull(mrsCard!库存差价), 0, mrsCard!库存差价)
            If In_编辑状态 = 1 Then
                .TextMatrix(.Rows - 2, .ColIndex("批准文号")) = "" & mrsCard!批准文号
            Else
                If mrsStock.RecordCount > 0 Then
                    mrsStock.MoveFirst
                    .TextMatrix(.Rows - 2, .ColIndex("批准文号")) = zlStr.Nvl(mrsStock!批准文号)
                Else
                    .TextMatrix(.Rows - 2, .ColIndex("批准文号")) = ""
                End If
            End If

            If Not mblnStock Then Call Get可用库存(.TextMatrix(.Rows - 2, .ColIndex("材料ID")), 0)
        End If
        
        '如果不显示对方库房的库存，需重新提取并更新
        If Not mblnStock Then
            .TextMatrix(.Rows - 2, .ColIndex("msin可用数量")) = msin可用数量
            .TextMatrix(.Rows - 2, .ColIndex("msin实际数量")) = msin实际数量
            .TextMatrix(.Rows - 2, .ColIndex("msin实际金额")) = msin实际金额
            .TextMatrix(.Rows - 2, .ColIndex("msin实际差价")) = msin实际差价
        End If
        .TextMatrix(.Rows - 2, .ColIndex("指导批发价")) = mrsCard!指导批发价
        .TextMatrix(.Rows - 2, .ColIndex("指导差价率")) = mrsCard!指导差价率
    End With
    
    lbl选定.Caption = "选定卫材（" & vsf选定.Rows - 2 & "条）"
End Sub



Private Sub vsHead_EnterCell()
    Dim lng收费细目ID As Long, intCol As Integer, LngSelectRow As Long
    Dim strTmp As String, recGetPrice As New ADODB.Recordset
    Dim strKc As String
    Dim i As Integer
    
   ' On Error Resume Next
    err = 0: On Error GoTo ErrHand:
    With vsHead
        '如果该规格卫材的价格到执行时间还未执行,则触发
        If Not mrsCard.EOF Then
            lng收费细目ID = Val(.TextMatrix(.Row, .ColIndex("材料ID")))
        End If
        If lng收费细目ID = 0 Then
            vsBatch.Clear 1
            vsBatch.Rows = 2
            mlngLastSelect材料ID = 0
            Exit Sub
        End If
        
        If mlngLastSelect材料ID = lng收费细目ID Then Exit Sub
        mlngLastSelect材料ID = lng收费细目ID
        
        
        '如果已到执行日期而价格未执行，执行计算过程
        gstrSQL = " Select ID From 收费价目 Where 收费细目ID=[1]" & _
                 " And 变动原因=0" & GetPriceClassString("")
        Set recGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng收费细目ID)
        
        With recGetPrice
            If Not .EOF Then
                If Not IsNull(!Id) Then
                    lng收费细目ID = !Id
                    gstrSQL = "zl_材料收发记录_Adjust(" & lng收费细目ID & ")"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption & "-产生价格调整记录"
                End If
            End If
        End With
    End With
    
    If In_编辑状态 = 2 Then
        vsBatch.Visible = False
        '读出该卫材规格下所有的卫材批次库存信息
        mbln时价 = (vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("时价")) = "是")
        mint分批 = 0
        If vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("库房分批")) = "是" Or vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("在用分批")) = "是" Then
            If vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("库房分批")) = "是" And vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("在用分批")) = "是" Then
                mint分批 = 3
            ElseIf vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("库房分批")) = "是" Then
                mint分批 = 1
            Else
                mint分批 = 2
            End If
        End If
        If Not ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) Then         '如果该卫材不分批
            vsBatch.Visible = False
            Form_Resize
        Else
            If vsBatch.Visible = False Then
                If mbln显示批次 = True Then '此参数控制能不能显示批次列表，如申领不明确批次模式
                    vsBatch.Visible = True
                End If
            Else
                If mbln显示批次 = False Then '此参数控制能不能显示批次列表，如申领不明确批次模式
                    vsBatch.Visible = False
                End If
            End If
        End If
        Form_Resize
        
        With mrsStock
            If .State = 1 Then .Close
            gstrSQL = ""
            If mbln空批次 Then
                gstrSQL = "Select " & IIf(mstr盘点时间 <> "", "/*+ Rule*/", "") & " 1 RID,名称 库房,0 批次,'新增批次卫生材料' 批号,sysdate 失效期," & _
                          "to_char(0," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char(0," & gOraFmt_Max.FM_数量 & ") 库存数量,to_char(0," & gOraFmt_Max.FM_金额 & ") 库存金额,to_char(0," & gOraFmt_Max.FM_金额 & ") 库存差价,sysdate as 灭菌失效期,to_char(0," & gOraFmt_Max.FM_零售价 & ") 售价,'' As 成本价,to_char(0," & gOraFmt_Max.FM_成本价 & ") 上次购价,'' 产地 , Sysdate As 生产日期, 0 As 上次供应商id,'' 批准文号,'' 商品条码,'' 内部条码 " & _
                          " From 部门表" & _
                          " Where ID=[1]" & _
                          " Union "
            End If
             
            gstrSQL = gstrSQL & " Select " & IIf(mstr盘点时间 <> "", "/*+ Rule*/", "") & " 2 RID,P.名称 库房,K.批次,K.上次批号 批号,K.效期 失效期,"
            If mblnStock Then
                If mbln散装单位 Then
                    strTmp = " to_char( K.可用数量," & gOraFmt_Max.FM_数量 & ") 可用数量," & _
                             " to_char( K.实际数量," & gOraFmt_Max.FM_数量 & ") as 库存数量,"
                Else
                    strTmp = " to_char( K.可用数量" & mstrUnitString & "," & gOraFmt_Max.FM_数量 & ") 可用数量," & _
                             " to_char( K.实际数量" & mstrUnitString & "," & gOraFmt_Max.FM_数量 & ") 库存数量,"
                End If
            Else
                strTmp = "to_char( ''," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char( ''," & gOraFmt_Max.FM_数量 & ") 库存数量,"
            End If
            
            
            '取库存
            '20060731:刘兴宏加入，主要解决盘点时间的库存
            strKc = "" & _
                "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次,a.上次供应商ID, a.上次采购价,A.零售价,平均成本价," & _
                "           a.实际数量,a.实际金额, a.实际差价, a.可用数量,a.上次批号,a.上次产地,a.效期,a.灭菌效期,a.上次生产日期,a.批准文号,a.商品条码,a.内部条码 " & _
                "   FROM 药品库存 a " & _
                "   Where a.药品id = [3]" & _
                "           AND a.性质=1 " & _
                "           AND a.库房id+0 = "
            If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
                strKc = strKc & IIf(mlng源库房ID = 0, "[1]", "[2]")
            End If
            
            If mstr盘点时间 <> "" Then
                strKc = strKc & _
                    "   UNION ALL " & _
                    "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次, a.供药单位ID 上次供应商ID,max(a.成本价) 上次采购价,max(A.零售价) as 零售价,0 as 平均成本价, " & _
                    "           -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量, " & _
                    "           -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
                    "           -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价,-SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数))  AS 可用数量,a.批号,a.产地 , A.效期,a.灭菌效期,a.生产日期,a.批准文号,a.商品条码,a.内部条码 " & _
                    "   FROM 药品收发记录 a " & _
                    "   Where a.药品id+0=[3]  " & _
                    "           AND a.库房id + 0 ="
                If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
                    strKc = strKc & IIf(mlng源库房ID = 0, "[1]", "[2]")
                End If
                strKc = strKc & " AND a.审核日期 >[5] " & _
                    " GROUP BY A.库房id, a.药品id,a.供药单位id, A.批次, A.批号, A.产地, A.效期, A.灭菌效期,a.生产日期,a.批准文号,a.商品条码,a.内部条码"
            End If
                      
            strKc = "" & _
                "   Select 库房id,药品id,nvl(批次,0) 批次,max(上次批号) 上次批号,min(灭菌效期) as 灭菌失效期,max(上次供应商ID) 上次供应商ID, " & _
                "       Sum(nvl(可用数量,0)) 可用数量," & _
                "       Sum(实际数量) 实际数量," & _
                "       Sum(实际金额) 实际金额," & _
                "       Sum(实际差价) 实际差价," & _
                "       max(上次采购价) 上次采购价,max(零售价) as 零售价,max(平均成本价) as 平均成本价, " & _
                "        Min(灭菌效期) 灭菌效期,Min(效期) 效期,max(上次产地) 上次产地 ,max(上次生产日期) 上次生产日期,max(批准文号) as 批准文号,max(商品条码) as 商品条码,max(内部条码) as 内部条码,1 As 性质" & _
                "   From (" & strKc & ")" & _
                "   Group by 库房id,药品id,nvl(批次,0) "
             
            gstrSQL = gstrSQL & strTmp & _
                     IIf(mblnStock, "to_char(K.实际金额," & gOraFmt_Max.FM_金额 & ")  as 库存金额,", "to_char(0," & gOraFmt_Max.FM_金额 & ")  库存金额,") & _
                     IIf(mblnStock, " to_char(K.实际差价," & gOraFmt_Max.FM_金额 & ")  as 库存差价", "to_char(0," & gOraFmt_Max.FM_金额 & ")  库存差价") & " ,K.灭菌效期 as 灭菌失效期," & _
                     IIf(mblnStock, "to_char(Decode(nvl(M.是否变价,0),0,G.现价,decode(nvl(K.零售价,0),0,nvl(K.实际金额,0)/decode(K.实际数量,null,1,0,1,K.实际数量),nvl(K.零售价,0)))" & IIf(mbln散装单位, "", "*nvl(D.换算系数,1)") & "," & gOraFmt_Max.FM_零售价 & ") 售价,", "to_char(0," & gOraFmt_Max.FM_零售价 & ") 售价,") & _
            " to_char(k.平均成本价," & gOraFmt_Max.FM_成本价 & ") as 成本价, " & _
                     IIf(mblnStock, "to_char(decode(nvl(K.上次采购价,0),0,(nvl(K.实际金额,0)-nvl(K.实际差价,0))/decode(K.实际数量,null,1,0,1,K.实际数量),K.上次采购价)" & IIf(mbln散装单位, "", "*nvl(D.换算系数,1)") & "," & gOraFmt_Max.FM_成本价 & ") 上次购价", "to_char(0," & gOraFmt_Max.FM_成本价 & ") 上次购价") & _
            "        ,K.上次产地 产地,k.上次生产日期 生产日期 ,k.上次供应商ID,k.批准文号,k.商品条码,k.内部条码 " & _
            " From 部门表 P,材料特性 D," & IIf(mstr盘点时间 <> "", "(" & strKc & ")", " 药品库存") & " K,收费项目目录 M,收费价目 G " & _
            " Where K.库房ID = P.ID And D.材料ID = K.药品ID And K.库房ID " & IIf(mstr盘点时间 <> "", " +0=", "=") & IIf(mlng源库房ID = 0, "[1]", "[2]") & _
            " And K.药品ID " & IIf(mstr盘点时间 <> "", " +0=", "=") & " [3]  And K.性质=1 " & _
            " And D.材料id=G.收费细目ID(+) " & _
            " And D.材料ID=M.ID And (M.站点=[7] or M.站点 is null) " & _
            " And m.Id = g.收费细目id And (Sysdate Between g.执行日期 And Nvl(g.终止日期, Sysdate)) " & _
            GetPriceClassString("G")
            
            Dim dtDate As Date
            If mstr盘点时间 <> "" Then
                dtDate = CDate(mstr盘点时间)
            Else
                dtDate = Now
            End If
                     
            If mbln盘点单 Then
                gstrSQL = gstrSQL & " And (K.实际数量<>0 Or K.实际金额<>0 Or K.实际差价<>0)"
            Else
                gstrSQL = gstrSQL & " And K.实际数量<>0 "
            End If
            
            If mstrCode <> "" Then
                gstrSQL = gstrSQL & " And (K.商品条码=[6] Or K.内部条码=[6]) "
            End If
             
            ' If mlng供应商ID <> 0 Then gstrSQL = gstrSQL & " And K.上次供应商ID=[4]"
             
            If gSystem_Para.P156_出库算法 = 0 Then
                gstrSQL = gstrSQL & " Order by RID,批次"
            Else
                gstrSQL = gstrSQL & " Order by RID,失效期,批次"
            End If
            
            Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng目库房ID, IIf(mlng源库房ID = 0, mlng目库房ID, mlng源库房ID), _
                           mlngLastSelect材料ID, mlng供应商ID, dtDate, mstrCode, gstrNodeNo)
        End With
        Dim blnState As Boolean
           
        With vsBatch
            .Redraw = flexRDNone
            Set .DataSource = mrsStock
            If mrsStock.EOF Then
                .Clear 1
                .Rows = 2
            End If
            Call SetFormat(0, mrsStock.EOF)
            
            .Redraw = flexRDBuffered
            If mbln空批次 And mrsStock.RecordCount <> 0 Then
                If .Rows >= 3 Then .Row = 2
                If .Rows = 2 Then .Row = 1
            End If
            blnState = ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And Not mrsStock.EOF        '如果该卫材不分批
            If mbln显示批次 = True And blnState = True Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        Form_Resize
    End If
    
    '设置按钮状态
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "材料ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID")))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Sub
        End If
        'mint库房:1-卫材库;2-在用;3-制剂室
        'mint分批:0-不分批;1-库房分批;2-在用分批;3-卫材库在用分批
        If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And mblnPrice Then
            If mbln显示批次 = False Then
                mblnSelect = True
            Else
                mblnSelect = blnState
            End If
        Else
            mblnSelect = True
        End If
    End With
    'Call ReSetWindowsFormLocal
    
    With vsBatch
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 And vsBatch.Visible = True Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("库存数量"))) > 0 Then
                        .TextMatrix(i, .ColIndex("库存数量")) = "有"
                    Else
                        .TextMatrix(i, .ColIndex("库存数量")) = "无"
                    End If
                    .TextMatrix(i, .ColIndex("库存金额")) = ""
                    .TextMatrix(i, .ColIndex("库存差价")) = ""
                Next
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsHead_GotFocus()
    With vsHead
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsHead_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    If KeyCode = vbKeyReturn Then vsHead_DblClick: Exit Sub
    
    With vsHead
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
    
End Sub

Private Sub vsHead_LostFocus()
    With vsHead
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        .BackColorSel = &H8000000A
    End With
End Sub

Private Function RefreshData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取满足条件的卫生材料,不区分批次,即表头部分
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, StrGroupBy As String
    Dim strLike As String
    Dim strSerach As String
    Dim strKc As String
    Dim strInput As String
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    Dim strCode As String
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim int入库产地取值方式 As Integer
    Dim lng材料ID As Long
    
    On Error GoTo ErrHandle

    RefreshData = False
    
    If mlngModule = 1712 Or mlngModule = 1714 Then
        int入库产地取值方式 = Val(zlDatabase.GetPara(268, glngSys))
    End If
    
    '先按条码查找
    If gblnCode = True Then
        gstrSQL = "Select 药品id From 药品库存 Where 性质 = 1 And 库房id = [1] And (商品条码 = [2] Or 内部条码 = [2])"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlng源库房ID, UCase(mstrInput))
        If Not rsData.EOF Then
            mstrCode = UCase(mstrInput)
            mstrInput = UCase(mstrInput)
            lng材料ID = rsData!药品id
        Else
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlng源库房ID, mstrInput)
            If Not rsData.EOF Then
                mstrCode = mstrInput
                lng材料ID = rsData!药品id
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlng源库房ID, LCase(mstrInput))
                If Not rsData.EOF Then
                    mstrCode = LCase(mstrInput)
                    mstrInput = LCase(mstrInput)
                    lng材料ID = rsData!药品id
                End If
            End If
        End If
    End If
    
    mbln只显示在用物资 = 判断只具备发料部门(mlng目库房ID)
    
    '判断虚拟库房
    gstrSQL = "select count(*) rec from 部门性质说明 where 工作性质='虚拟库房' and 部门id=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断虚拟库房", mlng目库房ID)
    If rsTmp!rec = 1 And (mobjOut.Name = "frmPurchaseCard" Or mobjOut.Name = "frmOtherInputCard") Then blnVirtualStock = True
    
    strLike = "" & GetMatchingSting(UCase(mstrInput), False) & ""
        
    '输入匹配：匹配编码，简码，名称，商品条码（固定左匹配），内部条码（固定左匹配）
    strSerach = " And (A.编码 Like [4] OR B.名称 Like [4] OR ( B.简码 LIKE [4] and B.码类=[6]))"
    If IsNumeric(mstrInput) Then                         '如果是数字,则只取编码
        If Mid(gSystem_Para.Para_输入方式, 1, 1) = "1" Then strSerach = " And (A.编码 Like [4] And B.码类=[6])"
    ElseIf zlStr.IsCharAlpha(mstrInput) Then          '输入全是字母时只匹配简码
        If Mid(gSystem_Para.Para_输入方式, 2, 1) = "1" Then strSerach = " And B.简码 Like [4] And B.码类=[6] "
    ElseIf zlStr.IsCharChinese(mstrInput) Then
        strSerach = " And B.名称 Like [4] And B.码类=[6] "
    End If
    
    strInput = " Select a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.服务对象, a.是否变价 " & _
            " From 收费项目目录 A,收费项目别名 B " & _
            " Where A.ID=B.收费细目ID And (A.站点=[8] or A.站点 is null) AND A.类别 ='4' And (A.撤档时间 is null Or A.撤档时间>[5]) " & strSerach & _
            " Union All " & _
            " Select a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.服务对象, a.是否变价 " & _
            " From 收费项目目录 A, 药品库存 B " & _
            " Where a.Id = b.药品id And 性质 = 1 And 库房id + 0 = [1] And (商品条码 = [7] Or 内部条码 = [7]) "
    
    '读出该卫材用途分类、属于指定规格卫材
    '--认为输入的是名称或简码，按此方式查找指定规格卫材--
    '对列头排顺序
    gstrSQL = "" & _
    " Select  D.诊疗id,D.材料id,D.分类ID,D.编码,D.通用名称,D.商品名,D.规格,D.产地,D.批准文号,d.注册证号,x.名称 As 上次供应商,to_char(D.售价," & gOraFmt_Max.FM_零售价 & ") 售价,to_char(d.成本价," & gOraFmt_Max.FM_成本价 & ") as 最新成本价,D.散装单位 散装,D.换算系数 系数,D.包装单位 包装," & _
                 IIf(mblnStock, "to_char(S.可用数量 " & IIf(mbln散装单位, "", "/D.换算系数") & "," & gOraFmt_Max.FM_数量 & ") 可用数量, " & _
                                "to_char(S.库存数量 " & IIf(mbln散装单位, "", "/D.换算系数") & "," & gOraFmt_Max.FM_数量 & ") 库存数量, " & _
                                "to_char(S.库存金额," & gOraFmt_Max.FM_金额 & ") 库存金额,to_char(S.库存差价," & gOraFmt_Max.FM_金额 & ") 库存差价,", _
                      "to_char(''," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char(''," & gOraFmt_Max.FM_数量 & ") 库存数量,to_char(''," & gOraFmt_Max.FM_金额 & ") 库存金额,to_char(''," & gOraFmt_Max.FM_金额 & ") 库存差价,") & _
    "           D.最大效期 有效期,D.灭菌效期,S.灭菌失效期,D.一次性材料,D.无菌性材料,D.库房分批,D.在用分批,D.时价,to_char(D.指导批发价," & gOraFmt_Max.FM_零售价 & ") 指导批发价,D.指导差价率,E.库房货位 " & _
    " From "
   
    
    '材料信息，材料目录
    If mbln只显示在用物资 Then
        gstrSQL = gstrSQL & " (Select Distinct u.诊疗id,u.材料id,H.分类ID,V.编码,v.名称 As 通用名称,B.名称 As 商品名,V.规格," & IIf(int入库产地取值方式 = 0, "decode(u.上次产地,null,v.产地,u.上次产地)", "decode(v.产地,null,u.上次产地,v.产地)") & " as 产地,u.批准文号,u.注册证号,V.计算单位 as 散装单位,U.包装单位," & _
                    "                       To_Char(U.换算系数," & GFM_XS & " ) 换算系数,nvl(To_Char(U.灭菌效期,'9999990'),0) 灭菌效期,nvl(To_Char(U.最大效期,'9999990'),0) 最大效期," & _
                    "                       Decode(U.库房分批,1,'是','否') 库房分批,Decode(U.在用分批,1,'是','否') 在用分批,Decode(U.一次性材料,1,'是','否')  一次性材料,Decode(U.无菌性材料,1,'是','否') 无菌性材料,Decode(V.是否变价,1,'是','否') 时价," & _
                    "                       U.指导批发价,To_Char(U.指导差价率," & GFM_CJL & " ) 指导差价率,现价 售价,u.成本价,Nvl(u.上次供应商id, 0) As 上次供应商id " & _
                    "               From 材料特性 U, " & _
                    "                    ( " & strInput & ") V," & _
                    "                    诊疗项目目录 H, " & _
                    "                   (SELECT 收费细目id, 执行科室id FROM 收费执行科室 WHERE 执行科室ID" & IIf(mlng源库房ID <> 0, "+0=[1]", IIf(mlng目库房ID <> 0, "+0=[2]", " Is Not NULL")) & ") K," & _
                    "                   (Select 收费细目ID, 执行科室ID From 收费执行科室 Where 执行科室ID" & IIf(mlng目库房ID <> 0, "+0=[2]", IIf(mlng源库房ID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                    "               收费项目别名 B, 收费价目 P " & _
                    "               where U.材料id=v.id and U.诊疗id=H.id And V.ID = B.收费细目id(+) And B.性质(+) = 3 " & _
                    "                       AND U.材料id=K.收费细目ID  " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                    "                       And U.材料id=i.收费细目Id " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                                            IIf(mbln只显示在用物资, " And U.跟踪在用=1 ", IIf(mblnTrackUsing = True, " and  U.跟踪在用 =0 ", "")) & " And v.Id = p.收费细目id And (Sysdate Between p.执行日期 And Nvl(p.终止日期, Sysdate)) " & _
                                            GetPriceClassString("P")
    Else
        gstrSQL = gstrSQL & " (Select Distinct u.诊疗id,u.材料id,H.分类ID,V.编码,v.名称 As 通用名称,B.名称 As 商品名,V.规格," & IIf(int入库产地取值方式 = 0, "decode(u.上次产地,null,v.产地,u.上次产地)", "decode(v.产地,null,u.上次产地,v.产地)") & " as 产地,u.批准文号,u.注册证号,V.计算单位 as 散装单位,U.包装单位," & _
                    "                       To_Char(U.换算系数," & GFM_XS & " ) 换算系数,nvl(To_Char(U.灭菌效期,'9999990'),0) 灭菌效期,nvl(To_Char(U.最大效期,'9999990'),0) 最大效期," & _
                    "                       Decode(U.库房分批,1,'是','否') 库房分批,Decode(U.在用分批,1,'是','否') 在用分批,Decode(U.一次性材料,1,'是','否')  一次性材料,Decode(U.无菌性材料,1,'是','否') 无菌性材料,Decode(V.是否变价,1,'是','否') 时价," & _
                    "                       U.指导批发价 ,To_Char(U.指导差价率," & GFM_CJL & " ) 指导差价率,现价 售价,u.成本价,Nvl(u.上次供应商id, 0) As 上次供应商id " & _
                    "               From 材料特性 U," & _
                    "                    (" & strInput & ") V," & _
                    "                   诊疗项目目录 H,收费项目别名 B, 收费价目 P," & _
                    "                   (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(mlng源库房ID <> 0, "=[1]", IIf(mlng目库房ID <> 0, "=[2]", " Is Not NULL")) & " ) K," & _
                    "                   (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(mlng目库房ID <> 0, "+0=[2]", IIf(mlng源库房ID <> 0, "+0=[1]", " Is Not NULL")) & " ) i" & _
                    "               where U.材料id=v.id and U.诊疗id=H.id And V.ID = B.收费细目id(+) And B.性质(+) = 3 " & _
                    "               AND U.材料id=K.收费细目ID " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                    "               And U.材料id=I.收费细目ID  " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                                    IIf(mbln只显示在用物资, " And U.跟踪在用=1 ", IIf(mblnTrackUsing = True, " and  U.跟踪在用 =0 ", "")) & " And v.Id = p.收费细目id And (Sysdate Between p.执行日期 And Nvl(p.终止日期, Sysdate)) " & _
                                    GetPriceClassString("P")

    End If
    
'    '只查找未停用的规格卫材
'    If mstr盘点时间 <> "" Then      '对盘点时间来说，如果盘点时间小于停用的时间也应该显示出来
'        gstrSQL = gstrSQL & " And (V.撤档时间 Is Null Or V.撤档时间>[5]) "
'    Else
'        gstrSQL = gstrSQL & " And (V.撤档时间 Is Null Or To_char(V.撤档时间,'yyyy-MM-dd')='3000-01-01')"
'    End If
'
    
    gstrSQL = gstrSQL & IIf(blnVirtualStock, " And nvl(u.高值材料,0)=1 and nvl(u.跟踪病人,0)=1 and nvl(u.跟踪在用,0)=1 and nvl(u.在用分批,0)=1", "")

    If mlng目库房ID > 0 Then
        gstrSQL = gstrSQL & " And " & _
            "     ( exists(select 1 from 部门性质说明 where 工作性质 In ('制剂室', '卫材库', '发料部门', '虚拟库房')  and 部门id=[2]" & ")  " & _
            "       or v.服务对象=(select distinct '1' from 部门性质说明 where 工作性质 like '发料部门' and 部门id=[2] and 服务对象 in(1,3))" & _
            "       or v.服务对象=(select distinct '2' from 部门性质说明 where 工作性质 like '发料部门' and 部门id=[2] and 服务对象 in(2,3)))"
    End If
    
    '只查找指定材质分类的规格材料
    gstrSQL = gstrSQL & " ) D,"

    '取库存
    '20060731:刘兴宏加入，主要解决盘点时间的库存
    strKc = "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次,a.上次供应商ID," & _
        "           a.实际数量,a.实际金额, a.实际差价, a.可用数量,a.上次批号,a.上次产地,a.效期,a.灭菌效期 " & _
        "   FROM 药品库存 a " & _
        "   Where a.性质=1 AND a.库房id = "
    If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
        strKc = strKc & IIf(mlng源库房ID = 0, "[2]", "[1]")
    End If
    
    '盘点时根据盘点时间计算盘点时间到当前时间的发生额
    If mstr盘点时间 <> "" Then
        strKc = strKc & _
            "   UNION ALL " & _
            "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次, a.供药单位ID 上次供应商ID, " & _
            "           -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量, " & _
            "           -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
            "           -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价,-SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 可用数量,a.批号,a.产地 , A.效期,a.灭菌效期" & _
            "   FROM 药品收发记录 a " & _
            "   Where a.库房id + 0 ="
        If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
            strKc = strKc & IIf(mlng源库房ID = 0, "[2]", "[1]")
        End If
        strKc = strKc & " AND a.审核日期 >[5] " & _
            " GROUP BY A.库房id, a.药品id,a.供药单位id, A.批次, A.批号, A.产地, A.效期, A.灭菌效期 "
    End If

    If mblnStock Then
        gstrSQL = gstrSQL & " (Select 药品id as 材料id,min(灭菌效期) as 灭菌失效期 , Sum(可用数量) 可用数量," & _
                " Sum(实际数量) 库存数量," & _
                " Sum(实际金额)  库存金额," & _
                " Sum(实际差价) 库存差价"
    Else
        gstrSQL = gstrSQL & " (Select 药品id as 材料id,min(灭菌效期) as 灭菌失效期, 0 可用数量," & _
                " 0 库存数量,0 库存金额,0 库存差价"
    End If
    If mstr盘点时间 <> "" Then
         gstrSQL = gstrSQL & " From (" & strKc & ") where 1=1 "
    Else
         gstrSQL = gstrSQL & " From 药品库存 Where 性质=1 "
    End If
    
    
    'If mlng供应商ID <> 0 Then gstrSQL = gstrSQL & " And 上次供应商ID=[3]"
    
    If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
        gstrSQL = gstrSQL & " And 库房ID" & IIf(mstr盘点时间 <> "", " +0 =", "=") & IIf(mlng源库房ID = 0, "[2]", "[1]") & "  Group By 药品id) S"
    Else
        gstrSQL = gstrSQL & " Group By 药品id) S"
    End If
    gstrSQL = gstrSQL & ",(Select 材料id,库房ID,库房货位 From 材料储备限额 " & _
              " Where 库房ID=" & IIf(mintEditState = 2, "[1]", "[2]") & ") E,供应商 X"
    
    '总条件
    gstrSQL = gstrSQL & " Where D.材料ID=S.材料ID"
    
    If mbln仅显示库存物资 And mblnStock Then
        gstrSQL = gstrSQL & " And S.可用数量<>0"
    Else
        '当系统参数“卫材出库库存检查”为不足禁止时，不提库存为零
        If Not (mintStockCheck = 2 And In_编辑状态 = 2) Or mbln盘点单 Or Not mblnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_编辑状态 = 2 Then gstrSQL = gstrSQL & " And S.可用数量<>0"
    End If
    gstrSQL = gstrSQL & " And D.材料ID=E.材料ID(+)  And d.上次供应商id = x.Id(+) Order By D.编码"
        
    Dim dtDate As Date
    If mstr盘点时间 <> "" Then
        dtDate = CDate(mstr盘点时间)
    Else
        dtDate = CDate("2999-12-31")
    End If

    Set mrsCard = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID, mlng目库房ID, mlng供应商ID, _
                    strLike, dtDate, gSystem_Para.int简码方式 + 1, mstrInput, gstrNodeNo)
    
    If lng材料ID = 0 Then
        gstrSQL = "Select distinct a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.服务对象, a.是否变价" & vbNewLine & _
                    "From 收费项目目录 A, 收费项目别名 B" & vbNewLine & _
                    "Where a.Id = b.收费细目id And (a.站点 =[3] Or a.站点 Is Null) And a.类别 = '4' And" & vbNewLine & _
                    "      (a.撤档时间 Is Null Or a.撤档时间 > To_Date('2999-12-31 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) And (b.简码 Like [1] or a.编码 Like [1] or b.名称 Like [1])"
    Else
        gstrSQL = "Select distinct a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, a.服务对象, a.是否变价" & vbNewLine & _
                    "From 收费项目目录 A, 收费项目别名 B" & vbNewLine & _
                    "Where a.Id = b.收费细目id And (a.站点 =[3] Or a.站点 Is Null) And a.类别 = '4' And" & vbNewLine & _
                    "      (a.撤档时间 Is Null Or a.撤档时间 > To_Date('2999-12-31 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) And a.id = [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在卫材查询", "%" & UCase(mstrInput) & "%", lng材料ID, gstrNodeNo)
    
    If In_编辑状态 = 2 Then
        '出库
        If rsTmp.RecordCount = 0 Then
            MsgBox "无此卫材，请重新输入！", vbInformation, gstrSysName
            Exit Function
        ElseIf rsTmp.RecordCount > 0 And mrsCard.RecordCount = 0 Then
            If blnVirtualStock = False Then
                MsgBox "此卫材无库存！", vbInformation, gstrSysName
            Else
                MsgBox "此卫材无库存或" & _
                vbCrLf & "虚拟库房流通需要卫材具有高值材料、跟踪病人、跟踪在用、在用分批属性，请检查！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    Else
        '入库
        If rsTmp.RecordCount = 0 Then
            MsgBox "无此卫材，请重新输入！", vbInformation, gstrSysName
            Exit Function
        ElseIf mrsCard.RecordCount = 0 Then
            MsgBox "未找到满足条件的卫材，可能是未设置存储库房，请检查", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With vsHead
        Set .DataSource = mrsCard
        If mrsCard.EOF Then
            .Rows = 2
        End If
        
        Call SetFormat(1, mrsCard.EOF)
        mblnSelect = (mrsCard.EOF <> True)
    End With
    
    With vsHead
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("库存数量"))) > 0 Then
                        .TextMatrix(i, .ColIndex("库存数量")) = "有"
                    Else
                        .TextMatrix(i, .ColIndex("库存数量")) = "无"
                    End If
                    .TextMatrix(i, .ColIndex("库存金额")) = ""
                    .TextMatrix(i, .ColIndex("库存差价")) = ""
                Next
            End If
        End If
    End With
    
    Call vsHead_EnterCell
    RefreshData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitRec()
    '----------------------------------------------------------------------------------------
    '功能:构建返回数据集结构
    '----------------------------------------------------------------------------------------
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn
        If .State = 1 Then .Close
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "诊疗ID", adDouble, 18, adFldIsNullable
        .Fields.Append "分类ID", adDouble, 18, adFldIsNullable
        .Fields.Append "供药单位ID", adDouble, 18, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 82, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "售价", adDouble, 18, adFldIsNullable
        .Fields.Append "散装单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 11, adFldIsNullable
        .Fields.Append "包装单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "最大效期", adDouble, 5, adFldIsNullable
        .Fields.Append "灭菌效期", adDouble, 5, adFldIsNullable
        
        .Fields.Append "一次性材料", adDouble, 2, adFldIsNullable
        .Fields.Append "无菌性材料", adDouble, 2, adFldIsNullable
        .Fields.Append "库房分批", adDouble, 2, adFldIsNullable
        .Fields.Append "在用分批", adDouble, 2, adFldIsNullable
        .Fields.Append "批准文号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "生产日期", adDate, , adFldIsNullable
        
        .Fields.Append "时价", adDouble, 2, adFldIsNullable
        .Fields.Append "批次", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "效期", adDate, , adFldIsNullable
        .Fields.Append "灭菌失效期", adDate, , adFldIsNullable
        .Fields.Append "可用数量", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际数量", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际金额", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "实际差价", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "指导批发价", adDouble, 11, adFldIsNullable
        .Fields.Append "指导差价率", adDouble, 11, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Function CombinateRec() As Boolean
    '组装记录集
    '定位记录集
    Dim blnEof As Boolean                   '是否存在库存批次记录
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    CombinateRec = False
    
    If chkContinue.Value = 0 Then '组装一条数据
        With mrsCard
            If .RecordCount <> 0 Then .MoveFirst
            .Find "材料ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID")))
            If .EOF Then
                MsgBox "发生内部错误！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mbln显示批次 = True Then '只有显示批次的情况下才需要做如下操作
                If ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And In_编辑状态 = 2 Then
                    With mrsStock
                        If .RecordCount <> 0 Then .MoveFirst
                        .Find "批次=" & Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批次")))
                        If .EOF Then
                            blnEof = True
                            If mblnPrice Then
                                MsgBox "发生内部错误！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End With
                End If
            End If
        End With
        
        '装数据写入记录集，供其它窗体使用
        With mrsReturn
            If .EOF Then .AddNew
            !材料ID = mrsCard!材料ID
            !诊疗id = mrsCard!诊疗id
            !分类id = mrsCard!分类id
            
            !编码 = mrsCard!编码
            !名称 = zlStr.Nvl(mrsCard!通用名称)
            !商品名 = zlStr.Nvl(mrsCard!商品名)
            !规格 = mrsCard!规格
            !产地 = mrsCard!产地
            !售价 = zlStr.Nvl(mrsCard!售价, 0)
            !散装单位 = mrsCard!散装
            !换算系数 = mrsCard!系数
            !包装单位 = mrsCard!包装
            !最大效期 = mrsCard!有效期
            !灭菌效期 = mrsCard!灭菌效期
            !一次性材料 = IIf(mrsCard!一次性材料 = "是", 1, 0)
            !无菌性材料 = IIf(mrsCard!无菌性材料 = "是", 1, 0)
            !库房分批 = IIf(mrsCard!库房分批 = "是", 1, 0)
            !在用分批 = IIf(mrsCard!在用分批 = "是", 1, 0)
            !灭菌失效期 = mrsCard!灭菌失效期
              
            !时价 = IIf(mrsCard!时价 = "是", 1, 0)
            
            '出库且分批
            If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) Then
                If mbln显示批次 = True Then
                    If vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批号")) = "新增批次卫生材料" Then
                        !批次 = -1
                    Else
                        If Not blnEof Then
                            !产地 = zlStr.Nvl(mrsStock!产地)
                            !生产日期 = mrsStock!生产日期
                            !批准文号 = zlStr.Nvl(mrsStock!批准文号)
                            !供药单位ID = mrsStock!上次供应商id
                            !批次 = Val(zlStr.Nvl(mrsStock!批次))
                            !批号 = zlStr.Nvl(mrsStock!批号)
                            !效期 = mrsStock!失效期
                            !灭菌失效期 = mrsStock!灭菌失效期
                            !可用数量 = IIf(IsNull(mrsStock!可用数量), 0, mrsStock!可用数量)
                            !实际数量 = IIf(IsNull(mrsStock!库存数量), 0, mrsStock!库存数量)
                            !实际金额 = IIf(IsNull(mrsStock!库存金额), 0, mrsStock!库存金额)
                            !实际差价 = IIf(IsNull(mrsStock!库存差价), 0, mrsStock!库存差价)
                            
                            If Not mblnStock Then Call Get可用库存(!材料ID, !批次)
                        End If
                    End If
                Else
                    If Not mblnStock Then Call Get可用库存(mrsCard!材料ID, 0)
                End If
            Else
                '入库或不分批
                !可用数量 = IIf(IsNull(mrsCard!可用数量), 0, mrsCard!可用数量)
                !实际数量 = IIf(IsNull(mrsCard!库存数量), 0, mrsCard!库存数量)
                !实际金额 = IIf(IsNull(mrsCard!库存金额), 0, mrsCard!库存金额)
                !实际差价 = IIf(IsNull(mrsCard!库存差价), 0, mrsCard!库存差价)
                If In_编辑状态 = 1 Then
                    !批准文号 = zlStr.Nvl(mrsCard!批准文号)
                Else
                    If mrsStock.RecordCount > 0 Then
                        mrsStock.MoveFirst
                        !批准文号 = zlStr.Nvl(mrsStock!批准文号)
                    Else
                        !批准文号 = ""
                    End If
                End If
                
                If Not mblnStock Then Call Get可用库存(!材料ID, 0)
            End If
            
            '如果不显示对方库房的库存，需重新提取并更新
            If Not mblnStock Then
                !可用数量 = msin可用数量
                !实际数量 = msin实际数量
                !实际金额 = msin实际金额
                !实际差价 = msin实际差价
            End If
            
            !指导批发价 = mrsCard!指导批发价
            !指导差价率 = mrsCard!指导差价率
            .Update
        End With
    Else '组装多条数据
        With mrsReturn
            For i = 1 To vsf选定.Rows - 2
                .AddNew
                !材料ID = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("材料ID")))
                !诊疗id = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("诊疗id")))
                !分类id = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("分类id")))
                
                !编码 = vsf选定.TextMatrix(i, vsf选定.ColIndex("编码"))
                !名称 = vsf选定.TextMatrix(i, vsf选定.ColIndex("名称"))
                !商品名 = vsf选定.TextMatrix(i, vsf选定.ColIndex("商品名"))
                !规格 = vsf选定.TextMatrix(i, vsf选定.ColIndex("规格"))
                !产地 = vsf选定.TextMatrix(i, vsf选定.ColIndex("产地"))
                !售价 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("售价")))
                !散装单位 = vsf选定.TextMatrix(i, vsf选定.ColIndex("散装单位"))
                !换算系数 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("换算系数")))
                !包装单位 = vsf选定.TextMatrix(i, vsf选定.ColIndex("包装单位"))
                !最大效期 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("最大效期")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("最大效期")))
                !灭菌效期 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("灭菌效期")))
                !灭菌失效期 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("灭菌失效期")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("灭菌失效期")))
                !一次性材料 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("一次性材料")))
                !无菌性材料 = vsf选定.TextMatrix(i, vsf选定.ColIndex("无菌性材料"))
                !库房分批 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("库房分批")))
                !在用分批 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("在用分批")))
                !时价 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("时价")))
                !批次 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("批次")))
                !批号 = vsf选定.TextMatrix(i, vsf选定.ColIndex("批号"))
                !效期 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("效期")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("效期")))
                !生产日期 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("生产日期")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("生产日期")))
                !批准文号 = vsf选定.TextMatrix(i, vsf选定.ColIndex("批准文号"))
                !供药单位ID = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("供药单位ID")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("供药单位ID")))
                !可用数量 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("可用数量")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("可用数量")))
                !实际数量 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("实际数量")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("实际数量")))
                !实际金额 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("实际金额")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("实际金额")))
                !实际差价 = IIf(vsf选定.TextMatrix(i, vsf选定.ColIndex("实际差价")) = "", Null, vsf选定.TextMatrix(i, vsf选定.ColIndex("实际差价")))
                !指导批发价 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("指导批发价")))
                !指导差价率 = Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("指导差价率")))
                
                .Update
            Next
        End With
        
    End If
    
    CombinateRec = True
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '当前库存数
    Dim intCol As Integer
    '检测是否允许选择
    CheckData = False
    
    If mblnSelect = False Then Exit Function
    
    If mbln显示批次 = False Then
        CheckData = True
        Exit Function '如果是不显示批次模式则直接不检查库存
    End If
    
    If vsBatch.Visible Then
        'lng供应商ID不为零，表示退货，无库存时不准继续
        If mlng供应商ID <> 0 Then
            intCol = vsBatch.ColIndex("上次供应商ID")
            If intCol < 0 Then Exit Function
            If Val(vsBatch.TextMatrix(vsBatch.Row, intCol)) <> 0 And mlng供应商ID <> Val(vsBatch.TextMatrix(vsBatch.Row, intCol)) Then
                MsgBox "你选择的退货商不是该卫生材料的供应商，不能继续操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        If mblnStock Then
            DblCurStock = Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("可用数量")))
        Else
            DblCurStock = Get可用库存(Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID"))), Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批次"))))
        End If
    Else
        If Not mrsCard.EOF Then
            If mblnStock Then
                DblCurStock = Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("可用数量")))
            Else
                DblCurStock = Get可用库存(Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("材料ID"))))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '如果源库房与目库房为空，则表明是卫材目录自己在进行常规设置，不判断
    If (mlng源库房ID = 0 And mlng目库房ID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '如果是盘点单调用卫材选择器，则不需判断，直接退出
    If mbln盘点单 Then
        CheckData = True
        Exit Function
    End If
    If vsBatch.Visible Or mbln时价 Then
        If (DblCurStock <> 0) Or Not mblnPrice Or vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("批号")) = "新增批次卫生材料" Then CheckData = True: Exit Function
        MsgBox "该" & IIf(mbln时价, "时价", "批次") & "卫材已经没有库存，不能继续操作！", vbInformation, gstrSysName
        Exit Function
    Else
        If mblnCheck = False Then
           CheckData = True
           Exit Function
        End If
    End If
    
    'mlng供应商ID不为零，表示退货，无库存时不准继续
    If mlng供应商ID <> 0 Then
        MsgBox "该卫材已经没有库存，不能继续操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Select Case mintStockCheck
    Case 1
        If MsgBox("该卫材已经没有库存，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        MsgBox "该卫材已经没有库存，不能继续操作！", vbInformation, gstrSysName
        Exit Function
    End Select
    CheckData = True
End Function

Public Function ShowSelect(ByVal frmMain As Form, ByVal 编辑模式 As Integer, Optional ByVal 源库房 As Long = 0, _
                    Optional ByVal 目库房 As Long = 0, Optional ByVal 使用部门 As Long = 0, Optional ByVal 查询串 As String = "", _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngWidth As Long = 0, Optional ByVal lngTxtHeight As Long = 0, Optional ByVal Bln检测库存 As Boolean = True, _
                    Optional ByVal bln检查批次或时价 As Boolean = True, Optional ByVal mbln盘点单据 As Boolean = False, Optional ByVal bln增加空批次 As Boolean = False, _
                    Optional ByVal bln显示库存 As Boolean = True, Optional ByVal lng供应商 As Long = 0, Optional ByVal bln散装单位 As Boolean = True, _
                    Optional ByVal str盘点时间 As String = "", _
                    Optional ByVal bln仅显示库存物资 As Boolean = False, _
                    Optional ByVal lngModule As Long = 0, _
                    Optional ByVal bln盘无存储库房材料 As Boolean = False, _
                    Optional ByVal strPrivs As String = "", _
                    Optional ByVal bln显示批次 As Boolean = True, Optional ByVal bln是否过滤 As Boolean = True) As ADODB.Recordset
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多选器
    '参数:
    '
    '   bln检查库存:遵守批次卫材及时价卫材零库存不准出库原则，可强制允许not (批次 or 时价) 卫材出库
    '   bln检查批次或时价:允许零库存的批次卫材及时价卫材出库
    '   mlng供应商ID:不为零表示退货
    '   str盘点时间:对盘点有效,主要是计算盘点时该的库存数
    '返回:被选择的卫材的记录集
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    mbln散装单位 = bln散装单位
    
    mlngModule = lngModule
    If mlngModule = 1717 Then   '1717:卫材领用
        mblnTrackUsing = IIf(Val(zlDatabase.GetPara("跟踪在用", glngSys, mlngModule, "0")) = 1, True, False)
    Else
        mblnTrackUsing = False
    End If
    
    '修改:刘兴宏   Bug:12972    日期:2008-05-08 15:28:14
    mbln仅显示库存物资 = bln仅显示库存物资 ': mlngModule = 0  '暂时未分配模块号,以后根据参数来决定
    mblnSelectSucess = False
    If frmMain Is Nothing Then
        mstrTittle = "卫材选择器"
    Else
        mstrTittle = frmMain.Caption
    End If
    With mWindowPosition
        .Left = WinLeft
        .Top = WinTop
        .lngTxtH = lngTxtHeight
        .lngTxtW = lngWidth
    End With
    With Me
        .In_编辑状态 = 编辑模式
        .In_源库房 = 源库房
        .In_目库房 = 目库房
        .In_部门 = 使用部门
        .In_字串 = Trim(查询串)
        .In_MainFrm = frmMain
        mbln盘点单 = mbln盘点单据
        mbln空批次 = bln增加空批次
        mblnCheck = Bln检测库存
        mblnPrice = bln检查批次或时价
        mblnStock = bln显示库存
        mlng供应商ID = lng供应商
        mstr盘点时间 = str盘点时间
        mbln盘无存储库房材料 = bln盘无存储库房材料
        mstrPrivs = strPrivs
        mbln显示批次 = bln显示批次
        mbln是否过滤 = bln是否过滤
        Me.Caption = mstrTittle
        If mblnSelectSucess Then GoTo GoOk:
        .Show 1, frmMain
    End With
GoOk:
    Set ShowSelect = mrsReturn.Clone
End Function

Public Function Get可用库存(ByVal lng材料ID As Long, Optional ByVal lng批次 As Long = 0) As Single
    Dim rsTemp As New ADODB.Recordset
     
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        " Select Sum(A.可用数量" & mstrUnitString & ") 可用数量,Sum(A.实际数量" & mstrUnitString & ") 实际数量,sum(A.实际金额) 实际金额,sum(A.实际差价) 实际差价 " & _
              " From 药品库存 A,材料特性 B " & _
              " Where A.药品ID=B.材料ID and A.性质=1 And A.药品ID=[1]" & IIf(lng批次 = 0, "", " And Nvl(A.批次,0)=[2]")
    
    If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
        gstrSQL = gstrSQL & " And A.库房ID=" & IIf(mlng源库房ID = 0, "[4]", "[3]")
    End If
    
    
    gstrSQL = gstrSQL & " Group By A.药品id"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng材料ID, lng批次, mlng源库房ID, mlng目库房ID)
    
    msin可用数量 = 0
    msin实际差价 = 0
    msin实际金额 = 0
    msin实际数量 = 0
    If Not rsTemp.EOF Then
        msin可用数量 = IIf(IsNull(rsTemp!可用数量), 0, rsTemp!可用数量)
        msin实际差价 = IIf(IsNull(rsTemp!实际差价), 0, rsTemp!实际差价)
        msin实际金额 = IIf(IsNull(rsTemp!实际金额), 0, rsTemp!实际金额)
        msin实际数量 = IIf(IsNull(rsTemp!实际数量), 0, rsTemp!实际数量)
    End If
    Get可用库存 = msin可用数量
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub imgLeft_Click()
    Call LoadFulltoColSel(False)
End Sub

Public Sub ReSetWindowsFormLocal()
    '功能:重新设置窗口的大小和位置
    Dim dblColsWidth As Double, dblMinRowheight As Double, lngScrW As Long
    Dim lngTaskHeight As Long
    Dim dblRowsHeight As Double
    Dim dblRowBatchHeight As Double
    Dim dblTemp As Double
    Dim i As Long
    '定位
    With mWindowPosition
        Me.Left = .Left + 15
        Me.Top = .Top
    End With
    
    dblColsWidth = 0
    For i = 0 To vsHead.Cols - 1
        If Not vsHead.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsHead.ColWidth(i) + 15
        End If
    Next
    dblMinRowheight = vsBatch.RowHeightMin
    lngTaskHeight = GetTaskbarHeight
    dblColsWidth = dblColsWidth + 300
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75

    dblRowsHeight = dblMinRowheight * vsHead.Rows + 30
    dblRowBatchHeight = (dblMinRowheight) * 6 '目前固定六行  'IIf(vsBatch.Visible, 1, 0) *   'IIf(vsBatch.Rows <= 4, 4, vsBatch.Rows)
    
    dblColsWidth = IIf(dblColsWidth < MFRM_MIN_WIDTH, MFRM_MIN_WIDTH, dblColsWidth)
    
    If Me.Top + dblRowsHeight + dblRowBatchHeight <= Screen.Height Then
        '窗体顶部+总行高度+小于等于屏幕高度。
        '看是否比最小高度还小,如果还小,就以最小度高为准
        Me.vsBatch.Height = dblRowBatchHeight
        If dblRowBatchHeight + dblRowsHeight < MFRM_MIN_HEIGHT Then
            Me.Height = MFRM_MIN_HEIGHT
        Else
            Me.Height = dblRowBatchHeight + dblRowsHeight
        End If
        
        Me.vsBatch.Height = dblRowBatchHeight
     '   If Me.ScaleHeight < Me.vsBatch.Top + Me.vsBatch.Height Then Me.vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
        
    Else
        '窗体顶部+总行数高度+批号的总高度大于屏幕高度,需要进一下检查
        '1.看上半屏幕高度是否比下半屏高度要高，如果，以上半屏的高度为准，否则以下半屏为准.
        
        If Screen.Height - Me.Top > Me.Top - mWindowPosition.lngTxtH - 15 Then
            '下半屏要大
            Me.Height = Screen.Height - Me.Top - lngTaskHeight
            '不能完全装下,只能根据情况来分配规格列表和批次列表的高度
            dblTemp = Me.ScaleHeight - dblRowsHeight
            If dblTemp > 6 * dblMinRowheight + 30 Then
               '乘下的高度要大于4行高度,则以批次的高度就为乘下的高度
               vsBatch.Height = dblTemp
            Else
                '乘下的高度不足4行的高度,则以4行高度为准
                vsBatch.Height = 6 * dblMinRowheight + 30
            End If
        Else
            dblTemp = Me.Top - mWindowPosition.lngTxtH - 15
            Me.Top = Me.Top - mWindowPosition.lngTxtH - 15
            '上半屏要大
            If dblTemp - dblRowBatchHeight - dblRowsHeight > 0 Then
                '上半屏能完全能装下
                Me.vsBatch.Height = dblRowBatchHeight
                Me.Height = dblRowBatchHeight + dblRowsHeight
                If Me.Height < MFRM_MIN_HEIGHT Then Me.Height = MFRM_MIN_HEIGHT
            Else
                Me.Height = dblTemp
                '不能完全装下,只能根据情况来分配规格列表和批次列表的高度
                dblTemp = Me.ScaleHeight - dblRowsHeight
                If dblTemp > 4 * dblMinRowheight Then
                   '乘下的高度要大于4行高度,则以批次的高度就为乘下的高度
                   vsBatch.Height = dblTemp
                Else
                    '乘下的高度不足4行的高度,则以4行高度为准
                    vsBatch.Height = 4 * dblMinRowheight
                End If
            End If
            Me.Top = Me.Top - Me.Height
        End If
    End If
    '窗体宽度定位
    '如果列宽总数小于等于当前窗体的宽度,则以列总数为准
    If dblColsWidth + Me.Left < Screen.Width Then
        '总列的宽度完全能显示
        Me.Width = dblColsWidth
    Else
        '检查是否左边屏幕大还是右边屏幕大
        If Screen.Width - Me.Left >= Me.Left Then
            '右边屏幕大
            Me.Width = Screen.Width - Me.Left
        Else
            Me.Left = Me.Left + mWindowPosition.lngTxtW
            '左边屏幕大
            If dblColsWidth < Me.Left Then
                Me.Width = dblColsWidth
            Else
                Me.Width = Me.Left
            End If
            Me.Left = Me.Left - Me.Width
        End If
    End If
 
    vsBatch.Top = Me.ScaleHeight - vsBatch.Height
    With vsHead
        .Height = IIf(vsBatch.Visible = False, Me.ScaleHeight - .Top, vsBatch.Top - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
        .Width = Me.ScaleWidth
        
    End With
    With vsBatch
        .Width = vsHead.Width
        .Left = ScaleLeft
    End With
End Sub

Private Function LoadFulltoColSel(ByVal blnBatch As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载列设置
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    
    If blnBatch Then
        Set vsGrid = vsBatch
        vsColSet.Tag = "Batch"
    Else
        Set vsGrid = vsHead
        vsColSet.Tag = "Head"
    End If
    vsColSet.Clear 1
    vsColSet.Rows = 2
    With vsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            '.coldata(i):1-固定,-1-不能选,0-可选
            If Trim(.ColKey(i)) <> "" And (.ColData(i) = 1 Or .ColData(i) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("列名")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("选择")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = .ColData(i)
                If .ColData(i) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    sngFrmHeight = Me.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = vsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        If blnBatch Then
            .Height = IIf(vsGrid.Top > sngSelSumHeight, sngSelSumHeight, vsGrid.Top)
            .Top = vsBatch.Top - .Height
        Else
            .Top = vsGrid.Top + vsGrid.RowHeight(0) + 15
            sngFrmHeight = sngFrmHeight - .Top
            If sngFrmHeight > sngSelSumHeight Then
                .Height = sngSelSumHeight
            Else
                .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
            End If
        End If
        .SetFocus
    End With
End Function
Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置显示列
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    If blnBatch Then
        Set vsGrid = vsBatch
    Else
        Set vsGrid = vsHead
    End If
    With vsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
    If blnBatch Then
        zl_vsGrid_Para_Save mlngModule, vsBatch, mstrTittle, "批次信息", False
    End If
End Function
Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '修改后
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("选择"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("列名")), blnShow, IIf(.Tag = "Head", False, True))
        Case Else
        End Select
    End With
End Sub

Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            'rowdata(i):1-固定,-1-不能选,0-可选
            If (.TextMatrix(Row, 1) = "库存差价" Or .TextMatrix(Row, 1) = "成本价" Or .TextMatrix(Row, 1) = "上次采购价" Or .TextMatrix(Row, 1) = "上次购价") And mblnCostView = False Then
                Cancel = True
            End If
            If .RowData(Row) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsColSet_LostFocus()
    vsColSet.Visible = False
End Sub

