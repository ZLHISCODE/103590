VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm材料选择器 
   Caption         =   "卫生材料选择器"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   Icon            =   "Frm材料选择器.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9465
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   2625
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "连续选择(&M)"
      Height          =   180
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pic选定区 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   2625
      ScaleHeight     =   2535
      ScaleWidth      =   4815
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
      Width           =   4815
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "Frm材料选择器.frx":0E42
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf选定 
         Height          =   2085
         Left            =   0
         TabIndex        =   10
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
         FormatString    =   $"Frm材料选择器.frx":1184
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
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   4
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   7020
      TabIndex        =   3
      Top             =   5850
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf材料规格 
      Height          =   3675
      Left            =   2640
      TabIndex        =   1
      Top             =   405
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6482
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   2010
      Top             =   1320
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
            Picture         =   "Frm材料选择器.frx":15ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm材料选择器.frx":2C47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   10081
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgTvw"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   8820
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm材料选择器.frx":4951
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf批次 
      Height          =   1620
      Left            =   2625
      TabIndex        =   2
      Top             =   4155
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2858
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   240
      Top             =   6000
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
            Picture         =   "Frm材料选择器.frx":665B
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm材料选择器.frx":69AD
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Image ImgLeftRight_S 
      Height          =   4485
      Left            =   2580
      MousePointer    =   9  'Size W E
      Top             =   1290
      Width           =   45
   End
   Begin VB.Image ImgUpDown_S 
      Height          =   45
      Left            =   2640
      MousePointer    =   7  'Size N S
      Top             =   4080
      Width           =   6765
   End
End
Attribute VB_Name = "Frm材料选择器"
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
Private mobjOut As Form                          '使用本程序的窗体（必须提供一个公共记录集，用以返回）

Private mblnStartUp As Boolean                   '启动成功
Private mblnFirstStart As Boolean                '第一次启动
Private mrsUnit As New ADODB.Recordset           '单位
Private mstrUnit As String                       '单位名称
Private mstrUnitString As String                 'SQL字串
Private mintStockCheck As Integer                '库存检测
Private mbln盘点单 As Boolean                    '盘点单据标志
Private mbln空批次 As Boolean                    '是否增加空批次供输入
Private mblnCheck As Boolean                     '是否检测库存(盘点、领用、申领用)
Private mblnPrice As Boolean                     '是否允许时价或批次卫材零出库
Private mblnTrackUsing As Boolean                '跟踪在用参数

'本程序使用记录集
Private mrsData As New ADODB.Recordset           '卫材用途分类
Private mrsCard As New ADODB.Recordset           '卫材卡片
Private mrsStock As New ADODB.Recordset          '卫材规格

'返回记录集
Private mrsReturn As ADODB.Recordset            '返回记录集(卫材信息所有列,卫材目录所有列,卫材库存所有列)
Private mint库房 As Integer                      '1-卫材库;2-发料部门;3-制剂室
Private mint分批 As Integer                      '0-不分批;1-库房分批;2-在用分批;3-卫材库在用分批
Private mbln只显示跟踪材料 As Boolean
Private mbln时价 As Boolean                      '时价
Private mblnStock  As Boolean
Private mstrCardSortBy As String                 '卫材卡片排序列
Private mstrPhysicSortBy As String               '卫材规格排序列
Private mlngCardRow As Long
Private mlngPhysicRow As Long
Private mlngLastSelect材料ID As Long             '上次选择的材料ID（用于是否刷新）
Private mbln仅显示库存物资 As Boolean
Private mbln盘无存储库房材料 As Boolean
Private mblnCostView As Boolean                 '查看成本价 true-允许查看 false-不允许查看
Private mblnProvider As Boolean                 '查看上次供应商相关信息 true-允许查看 false-不允许查看
Private mbln显示批次 As Boolean                 'true-显示批次列表，false-不显示批次列表
Private mstrPrivs As String                    '操作员权限
Private mlngModule As Long
Private mbln是否过滤 As Boolean                '是否由过滤打开

'调用get可用库存后，返回的可用数量，实际数量，实际金额及实际差价
Private msin可用数量 As Single
Private msin实际数量 As Single
Private msin实际金额 As Single
Private msin实际差价 As Single
Private mbln散装单位 As Boolean
Private mstr盘点时间 As String
Private Enum mCol
    诊疗id = 0
    材料ID
    分类id
    编码
    卫材名称
    商品名
    规格
    产地
    批准文号
    注册证号
    上次供应商
    售价
    最新成本价
    散装单位
    换算系数
    包装单位
    可用数量
    库存数量
    库存金额
    库存差价
    有效期
    灭菌效期
    灭菌失效期
    一次性材料
    无菌性材料
    库房分批
    在用分批
    时价
    指导批发价
    指导差价率
    库房货位
End Enum


'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'--公共--
Private Const mCols = 31            '总列数

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
        With msf材料规格
            
            If BlnSetHeader Then
                .Cols = mCols
                .TextMatrix(0, mCol.诊疗id) = "诊疗ID"
                .TextMatrix(0, mCol.材料ID) = "材料ID"
                .TextMatrix(0, mCol.分类id) = "分类ID"
                .TextMatrix(0, mCol.编码) = "编码"
                .TextMatrix(0, mCol.卫材名称) = "卫材名称"
                .TextMatrix(0, mCol.商品名) = "商品名"
                .TextMatrix(0, mCol.规格) = "规格"
                .TextMatrix(0, mCol.产地) = "产地"
                .TextMatrix(0, mCol.售价) = "售价"
                .TextMatrix(0, mCol.散装单位) = "散装单位"
                .TextMatrix(0, mCol.换算系数) = "换算系数"
                .TextMatrix(0, mCol.包装单位) = "包装单位"
                .TextMatrix(0, mCol.可用数量) = "可用数量"
                .TextMatrix(0, mCol.库存数量) = "库存数量"
                .TextMatrix(0, mCol.库存金额) = "库存金额"
                .TextMatrix(0, mCol.库存差价) = "库存差价"
                .TextMatrix(0, mCol.有效期) = "有效期"
                .TextMatrix(0, mCol.库房分批) = "库房分批"
                .TextMatrix(0, mCol.在用分批) = "在用分批"
                .TextMatrix(0, mCol.一次性材料) = "一次性材料"
                .TextMatrix(0, mCol.无菌性材料) = "无菌性材料"
                .TextMatrix(0, mCol.灭菌效期) = "灭菌效期"
                .TextMatrix(0, mCol.灭菌失效期) = "灭菌失效期"
                .TextMatrix(0, mCol.时价) = "时价"
                .TextMatrix(0, mCol.指导批发价) = "指导批发价"
                .TextMatrix(0, mCol.指导差价率) = "指导差价率"
                .TextMatrix(0, mCol.库房货位) = "库房货位"
                .TextMatrix(0, mCol.最新成本价) = "最新成本价"
                
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol >= mCol.可用数量 And intCol <= mCol.库存差价 Or intCol >= mCol.时价 And intCol <= mCol.指导差价率 Or intCol = mCol.售价 Or intCol = mCol.换算系数 Or intCol = mCol.最新成本价 Then
                    .ColAlignment(intCol) = 7
                ElseIf intCol = mCol.散装单位 Or intCol = mCol.包装单位 Or intCol = mCol.灭菌效期 Or intCol = mCol.一次性材料 Or intCol = mCol.无菌性材料 Then
                    .ColAlignment(intCol) = 4
                Else
                    .ColAlignment(intCol) = 1
                End If
            Next
            
            If mblnStartUp = False Then
                .ColWidth(mCol.诊疗id) = 0
                .ColWidth(mCol.材料ID) = 0
                .ColWidth(mCol.分类id) = 0
                .ColWidth(mCol.编码) = 800
                .ColWidth(mCol.卫材名称) = 2000
                .ColWidth(mCol.商品名) = 2000
                .ColWidth(mCol.规格) = 1600
                .ColWidth(mCol.产地) = 1500
                .ColWidth(mCol.售价) = 1000
                .ColWidth(mCol.散装单位) = 800
                .ColWidth(mCol.换算系数) = 800
                .ColWidth(mCol.包装单位) = 800
                .ColWidth(mCol.可用数量) = 1000
                .ColWidth(mCol.库存数量) = 1000
                .ColWidth(mCol.库存金额) = 1000
                .ColWidth(mCol.有效期) = 1000
                .ColWidth(mCol.灭菌失效期) = 1000
                .ColWidth(mCol.灭菌效期) = 0
                .ColWidth(mCol.灭菌效期) = 0
                .ColWidth(mCol.一次性材料) = 800
                .ColWidth(mCol.无菌性材料) = 800
                .ColWidth(mCol.库房分批) = 800
                .ColWidth(mCol.在用分批) = 800
                .ColWidth(mCol.时价) = 1000
                .ColWidth(mCol.指导差价率) = 0
                .ColWidth(mCol.库房货位) = 1000
                .Row = 1
                RestoreFlexState msf材料规格, Me.Name
                .ColWidth(mCol.库存差价) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(mCol.最新成本价) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(mCol.指导批发价) = IIf(mblnCostView = False, 0, 1000)
                If mlngModule = 1725 Or .ColWidth(mCol.上次供应商) = 0 Then .ColWidth(mCol.上次供应商) = IIf(mblnProvider = False, 0, 1300)
            End If
        End With
    Case 0
        With Msf批次
            
            If BlnSetHeader Then
                .Cols = 17
                .TextMatrix(0, 0) = ""
                .TextMatrix(0, 1) = "库房"
                .TextMatrix(0, 2) = "批次"
                .TextMatrix(0, 3) = "批号"
                .TextMatrix(0, 4) = "失效期"
                .TextMatrix(0, 5) = "可用数量"
                .TextMatrix(0, 6) = "库存数量"
                .TextMatrix(0, 7) = "库存金额"
                .TextMatrix(0, 8) = "库存差价"
                .TextMatrix(0, 9) = "灭菌失效期"
                .TextMatrix(0, 10) = "售价"
                .TextMatrix(0, 11) = "成本价"
                .TextMatrix(0, 12) = "上次购价"
                .TextMatrix(0, 13) = "产地"
                .TextMatrix(0, 14) = "生产日期"
                .TextMatrix(0, 15) = "上次供应商ID"
                .TextMatrix(0, 16) = "批准文号"
                
                
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                .ColAlignment(intCol) = 1
            Next
            .ColWidth(0) = 0
            .ColWidth(15) = 0
            
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            .ColAlignment(7) = 7
            .ColAlignment(8) = 7
            .ColAlignment(10) = 7
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            
            If mblnStartUp = False Then
                .ColWidth(0) = 0
                .ColWidth(1) = 1200
                .ColWidth(2) = 0
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 1200
                .ColWidth(6) = 1200
                .ColWidth(7) = 1200
                .ColWidth(9) = 1200
                .ColWidth(10) = 1200
                .ColWidth(13) = 1200
                .ColWidth(14) = 1200
                .ColWidth(16) = 1200
                                
                .Row = 1
                Call RestoreFlexState(Msf批次, Me.Name)
                .ColWidth(8) = IIf(mblnCostView = False, 0, 1200)
                .ColWidth(11) = IIf(mblnCostView = False, 0, 1200)
                .ColWidth(12) = IIf(mblnCostView = False, 0, 1200)
            End If
        End With
    End Select
End Sub

Private Sub chkContinue_Click()
    Dim blnState As Boolean

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
    
    '判断确认按钮是否可用
    If In_编辑状态 = 1 Then cmd确定.Enabled = True: Exit Sub
    
    blnState = ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And Not mrsStock.EOF

    If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And mblnPrice Then
        If mbln显示批次 = False Then
            cmd确定.Enabled = True
        Else
            cmd确定.Enabled = blnState
        End If
    Else
        cmd确定.Enabled = True
    End If
    
    If chkContinue.Value = 1 Then cmd确定.Enabled = True
    
End Sub

Private Sub Cmd取消_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmd确定_Click()
    Dim blnValid As Boolean
    
    If chkContinue.Value = 0 Then '不可多选时
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
        '组装记录集
        If CombinateRec = False Then Exit Sub
        Unload Me
        Exit Sub
    Else '可多选数据
        If CombinateRec = False Then Exit Sub
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    mblnStartUp = False
    mblnFirstStart = False
    mbln只显示跟踪材料 = False
    '取售价单位
    mstrUnit = ""
    mstrUnitString = ""
    mintStockCheck = 0
    mlngLastSelect材料ID = 0
    
    chkContinue.Visible = mbln是否过滤 = False
    
    Msf批次.Visible = (In_编辑状态 = 2)
    pic选定区.Visible = False
    picSplit02_S.Visible = False
    pic选定区.Tag = "展开"
    
    On Error GoTo ErrHandle

    '初始化记录集
    InitRec
    
    If mobjOut Is Nothing Then
        ShowMsgBox "请指定主窗体！"
        Exit Sub
    End If
    
    '初始化并检测相关数据完整性
    If LoadTvwData() = False Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    If mlngModule = 1725 Then
        mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "查看供应商")
    Else
        mblnProvider = True
    End If
    
    '提取当前库存控制参数
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"
    Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
        
    With mrsUnit
        If Not mrsUnit.EOF Then
            mintStockCheck = mrsUnit!库存检查
        End If
    End With
        
    '检查源库房是否为卫材库
    If mlng源库房ID <> 0 Then
        mint库房 = 3
        
        gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '发料部门' Or 工作性质 like '%制剂室') And 部门id=[1]"
        Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
        
        If mrsUnit.EOF Then
            gstrSQL = "select 部门ID from 部门性质说明 where 工作性质 In('卫材库','虚拟库房') And 部门id=[1]"
            Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID)
            If Not mrsUnit.EOF Then mint库房 = 1
        Else
            mint库房 = 2
        End If
    End If
    
    '读出该使用的单位级数
    If mbln散装单位 Then
        mstrUnitString = "/1"
    Else
        mstrUnitString = "/nvl(换算系数,1)"
    End If
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_成本价)
        .FM_金额 = GetFmtString(IIf(mbln散装单位, 0, 1), g_金额)
        .FM_零售价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_售价)
        .FM_数量 = GetFmtString(IIf(mbln散装单位, 0, 1), g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_成本价, True)
        .FM_金额 = GetFmtString(IIf(mbln散装单位, 0, 1), g_金额, True)
        .FM_零售价 = GetFmtString(IIf(mbln散装单位, 0, 1), g_售价, True)
        .FM_数量 = GetFmtString(IIf(mbln散装单位, 0, 1), g_数量, True)
    End With
    
    
    tvwClass_NodeClick tvwClass.SelectedItem
    mblnStartUp = True

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadTvwData() As Boolean
    Dim NodeThis As Node, ItemThis As ListItem
    
    Dim Int末级 As Integer
    Dim rs材质分类 As New ADODB.Recordset
    
    '材料用途分类是否有数据
    On Error GoTo ErrHandle
    LoadTvwData = False
    
    With mrsData
        
        gstrSQL = "" & _
            "   Select ID,上级ID,名称,1 as 末级 " & _
            "   From 诊疗分类目录 where 类型=7" & _
            "   Start With 上级ID IS NULL Connect By Prior ID=上级ID " & _
            "   Order by level,ID"
        
        zlDatabase.OpenRecordset mrsData, gstrSQL, Me.Caption
        
        If .EOF Then
            ShowMsgBox "请初始化卫材分类（卫材目录管理）！"
            Exit Function
        End If
        
        
        '将卫材用途分类数据装入
        tvwClass.Nodes.Clear
        tvwClass.Nodes.Add , 4, "Root", "所有卫生材料", 1, 1
        
        Do While Not .EOF
            
            If IsNull(!上级ID) Then
                Set NodeThis = tvwClass.Nodes.Add("Root", 4, "K_" & !Id, !名称, 2, 2)
            Else
                Set NodeThis = tvwClass.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称, 2, 2)
            End If
            .MoveNext
        Loop
    End With
    
    With tvwClass
        .Nodes(1).Selected = True
    End With
    
    LoadTvwData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    mblnFirstStart = True
    If Me.Height < 5835 Then Me.Height = 5835
    If Me.Width < 8415 Then Me.Width = 8415
    
    With ImgLeftRight_S
        .Top = 0
        .Height = Me.ScaleHeight - 200 - Cmd取消.Height - .Top
    End With
    
    With tvwClass
        .Top = 0
        .Height = ImgLeftRight_S.Height
        .Width = ImgLeftRight_S.Left
    End With
    
    With ImgUpDown_S
        .Left = ImgLeftRight_S.Left + ImgLeftRight_S.Width
        .Width = Me.ScaleWidth - .Left
    End With
    
    With msf材料规格
        .Left = ImgUpDown_S.Left
        .Top = ImgLeftRight_S.Top + (chkContinue.Height + 2 * chkContinue.Top)
        .Width = ImgUpDown_S.Width
    End With
    
    With Msf批次
        If .Visible Then
            .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
            .Height = ImgLeftRight_S.Top + ImgLeftRight_S.Height - .Top
            .Left = msf材料规格.Left
            .Width = msf材料规格.Width
        End If
    End With
    
    With Cmd取消
        .Top = tvwClass.Top + tvwClass.Height + 150
        .Left = Me.ScaleWidth - .Width - 150
    End With
    With cmd确定
        .Top = Cmd取消.Top
        .Left = Cmd取消.Left - .Width - 100
    End With
    
    With msf材料规格
        .Height = IIf(Msf批次.Visible = False, tvwClass.Top + tvwClass.Height - .Top, Msf批次.Top - 45 - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
    End With
    
    If picSplit02_S.Visible Then
        '设置分界线的top
        If Msf批次.Visible Then '批次可见
            Msf批次.Height = Msf批次.Height - (lbl选定.Height + picSplit02_S.Height)
            picSplit02_S.Top = Msf批次.Top + Msf批次.Height
        Else
            msf材料规格.Height = msf材料规格.Height - (lbl选定.Height + picSplit02_S.Height)
            picSplit02_S.Top = msf材料规格.Top + msf材料规格.Height
        End If
        
        picSplit02_S.Width = msf材料规格.Width
    End If
    
    If pic选定区.Visible Then
        pic选定区.Width = msf材料规格.Width
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

        If pic选定区.Tag = "收缩" Then
            pic选定区.Tag = "展开"
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            picSplit02_S.MousePointer = 0
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
    SaveFlexState msf材料规格, Me.Name
    SaveFlexState Msf批次, Me.Name
End Sub

Private Sub ImgLeftRight_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight_S
        If .Left + x < 2500 Then Exit Sub
        If .Left + x > Me.ScaleWidth - 4500 Then Exit Sub
        
        .Move .Left + x
    End With
    
    Form_Resize
End Sub

Private Sub ImgUpDown_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDown_S
        If .Top + y < 2500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 2500 Then Exit Sub
        
        .Move .Left, .Top + y
    End With
    
    Form_Resize
End Sub
Private Sub Msf批次_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    Dim i As Integer
    '实现列排序
    With Msf批次
        If .MouseRow <> 0 Then Exit Sub
        If mrsStock.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(mstrPhysicSortBy, 2) = StrHeader Then
            mstrPhysicSortBy = IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            mrsStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            mstrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
            mrsStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = mrsStock

        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(0, False)
        
    End With
    
    With Msf批次
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 6)) > 0 Then
                        .TextMatrix(i, 6) = "有"
                    Else
                        .TextMatrix(i, 6) = "无"
                    End If
                    .TextMatrix(i, 7) = ""
                    .TextMatrix(i, 8) = ""
                Next
            End If
        End If
    End With
End Sub

Private Sub Msf批次_DblClick()
    On Error Resume Next
    If cmd确定.Enabled = False Then Exit Sub
    
    With mrsStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    
    If chkContinue.Value = 1 Then
        FillVSF选定
        Exit Sub
    End If
    
    Call cmd确定_Click
End Sub

Private Sub Msf批次_EnterCell()
    Dim intCol As Integer, LngSelectRow As Long
    Dim recGetPrice As New ADODB.Recordset
    Dim lng收费细目ID As Long
    On Error Resume Next
    
    With Msf批次
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngPhysicRow <> 0 Then
            .Row = mlngPhysicRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngPhysicRow = LngSelectRow
        .Row = mlngPhysicRow     '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf批次_GotFocus()
    With Msf批次
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf批次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf批次_DblClick
End Sub

Private Sub Msf批次_LostFocus()
    With Msf批次
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf材料规格_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    Dim i As Integer
    
    '实现列排序
    With msf材料规格
        If .MouseRow <> 0 Then Exit Sub
        If mrsCard.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(mstrCardSortBy, 2) = StrHeader Then
            mstrCardSortBy = IIf(Mid(mstrCardSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            mrsCard.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrCardSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            mstrCardSortBy = "A" & .TextMatrix(0, .MouseCol)
            mrsCard.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = mrsCard

        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(1, False)
    End With
    
    With msf材料规格
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, mCol.库存数量)) > 0 Then
                        .TextMatrix(i, mCol.库存数量) = "有"
                    Else
                        .TextMatrix(i, mCol.库存数量) = "无"
                    End If
                    .TextMatrix(i, mCol.库存金额) = ""
                    .TextMatrix(i, mCol.库存差价) = ""
                Next
            End If
        End If
    End With
End Sub

Private Sub Msf材料规格_DblClick()
    If mrsCard.EOF Then Exit Sub
    If mrsCard.RecordCount = 0 Then Exit Sub
    
    If chkContinue.Value = 1 Then
        FillVSF选定
        Exit Sub
    End If
    
    If cmd确定.Enabled Then
        cmd确定_Click
    Else
        MsgBox "该卫材没有库存，不能继续操作！", vbInformation, gstrSysName
    End If
End Sub

Private Sub FillVSF选定()
    Dim blnEof As Boolean         '是否存在批次库存
    Dim i As Integer
    Dim blnValid    As Boolean
    
    '检查药品重复
    If chkContinue.Value = 1 Then
        For i = 1 To vsf选定.Rows - 2
            If Val(vsf选定.TextMatrix(i, vsf选定.ColIndex("材料ID"))) = Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID)) Then
                If Msf批次.Visible Then
                    If vsf选定.TextMatrix(i, vsf选定.ColIndex("批次")) = Msf批次.TextMatrix(Msf批次.Row, 2) Then
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
        .Find "材料ID=" & Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID))
        
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mbln显示批次 = True Then '只有显示批次的情况下才需要做如下操作
            If ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And In_编辑状态 = 2 Then
                With mrsStock
                    If .RecordCount <> 0 Then .MoveFirst
                    .Find "批次=" & Val(Msf批次.TextMatrix(Msf批次.Row, 2))
                    If .EOF Then
                        blnEof = True
                        If mblnPrice Then
                            MsgBox "无库存数据！", vbInformation, gstrSysName
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
        .TextMatrix(.Rows - 2, .ColIndex("散装单位")) = mrsCard!散装单位
        .TextMatrix(.Rows - 2, .ColIndex("换算系数")) = mrsCard!换算系数
        .TextMatrix(.Rows - 2, .ColIndex("包装单位")) = mrsCard!包装单位
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
                If Msf批次.TextMatrix(Msf批次.Row, 3) = "新增批次卫生材料" Then
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

Private Sub Msf材料规格_EnterCell()
    Dim lng收费细目ID As Long, intCol As Integer, LngSelectRow As Long
    Dim strTmp As String, recGetPrice As New ADODB.Recordset
    Dim strKc As String
    Dim i As Integer
    
'    On Error Resume Next
    On Error GoTo ErrHandle

    With msf材料规格
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngCardRow <> 0 Then
            If mlngCardRow <= .Rows - 1 Then
                .Row = mlngCardRow       '清除上次选中行
            End If
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngCardRow = LngSelectRow
        .Row = mlngCardRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
        
        '如果该规格卫材的价格到执行时间还未执行,则触发
        lng收费细目ID = Val(.TextMatrix(.Row, mCol.材料ID))
        
        If lng收费细目ID = 0 Then
            If Msf批次.Visible Then
                Msf批次.Clear
                Msf批次.Rows = 2
                Call SetFormat(0, True)
                Msf批次_EnterCell
            Else
                Call SetFormat(0, True)
            End If
            mlngLastSelect材料ID = 0
            Exit Sub
        End If
        
        If mlngLastSelect材料ID = lng收费细目ID Then Exit Sub
        mlngLastSelect材料ID = lng收费细目ID
        
        
        '如果已到执行日期而价格未执行，执行计算过程
        
        gstrSQL = "Select ID From 收费价目 " & _
                  "Where 收费细目ID=[1] And 变动原因=0" & GetPriceClassString("")
        Set recGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng收费细目ID)
        
        With recGetPrice
            If Not recGetPrice.EOF Then
                If Not IsNull(recGetPrice!Id) Then
                    lng收费细目ID = recGetPrice!Id
                    gstrSQL = "zl_材料收发记录_Adjust(" & lng收费细目ID & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-产生材料价格调整记录")
                End If
            End If
        End With
    End With
    
    If In_编辑状态 = 2 Then
        Msf批次.Visible = False
        '读出该卫材规格下所有的卫材批次库存信息
        
        mbln时价 = (msf材料规格.TextMatrix(msf材料规格.Row, mCol.时价) = "是")
        mint分批 = 0
        If msf材料规格.TextMatrix(msf材料规格.Row, mCol.库房分批) = "是" Or msf材料规格.TextMatrix(msf材料规格.Row, mCol.在用分批) = "是" Then
            If msf材料规格.TextMatrix(msf材料规格.Row, mCol.库房分批) = "是" And msf材料规格.TextMatrix(msf材料规格.Row, mCol.在用分批) = "是" Then
                mint分批 = 3
            ElseIf msf材料规格.TextMatrix(msf材料规格.Row, mCol.库房分批) = "是" Then
                mint分批 = 1
            Else
                mint分批 = 2
            End If
        End If
        'mint库房 1-卫材库;2-发料部门;3-制剂室
        'mint分批 0-不分批;1-库房分批;2-在用分批;3-卫材库在用分批
        If Not ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) Then '如果该卫材不分批
            Msf批次.Visible = False
            Form_Resize
        Else
            If Msf批次.Visible = False Then
                If mbln显示批次 = True Then '此参数控制能不能显示批次列表，如申领不明确批次模式
                    Msf批次.Visible = True
                End If
            Else
                If mbln显示批次 = False Then '此参数控制能不能显示批次列表，如申领不明确批次模式
                    Msf批次.Visible = False
                End If
            End If
        End If
        Form_Resize
        
        gstrSQL = ""
        
        If mbln空批次 Then
            gstrSQL = "" & _
                "   Select " & IIf(mstr盘点时间 <> "", "/*+ Rule*/", "") & " 1 RID,名称 库房,0 批次,'新增批次卫生材料' 批号,sysdate 失效期,to_char(0," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char(0," & gOraFmt_Max.FM_数量 & ") 库存数量,to_char(0," & gOraFmt_Max.FM_金额 & ") 库存金额" & _
                "           ,to_char(0," & gOraFmt_Max.FM_金额 & ") 库存差价,sysdate as 灭菌失效期,to_char(0," & gOraFmt_Max.FM_零售价 & ") 售价,'' As 成本价,to_char(0," & gOraFmt_Max.FM_成本价 & ") 上次购价,'' 产地 , Sysdate As 生产日期, 0 As 上次供应商id,'' 批准文号" & _
                "   From 部门表" & _
                "   Where ID=[1]" & _
                "   Union "
        End If
        
        gstrSQL = gstrSQL & " Select " & IIf(mstr盘点时间 <> "", "/*+ Rule*/", "") & " 2 RID,P.名称 库房,K.批次,K.上次批号 批号,K.效期 失效期,"
        
        If mblnStock Then
            If mbln散装单位 Then
                strTmp = " to_char(K.可用数量," & gOraFmt_Max.FM_数量 & ") 可用数量," & _
                         " to_char(K.实际数量," & gOraFmt_Max.FM_数量 & ") as 库存数量,"
            Else
                strTmp = " to_char(K.可用数量," & gOraFmt_Max.FM_数量 & ") 可用数量," & _
                         " to_char(K.实际数量," & gOraFmt_Max.FM_数量 & ") as 库存数量,"
            End If
        Else
            strTmp = "to_char(0," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char(0," & gOraFmt_Max.FM_数量 & ") 库存数量,"
        End If
                 
        '取库存
        '20060731:刘兴宏加入，主要解决盘点时间的库存
        strKc = "" & _
            "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次,a.上次供应商ID, a.上次采购价," & _
            "           a.实际数量,a.实际金额, a.实际差价, a.可用数量,A.零售价,平均成本价,a.上次批号,a.上次产地,a.效期,a.灭菌效期,a.上次生产日期,a.批准文号 " & _
            "   FROM 药品库存 a " & _
            "   Where a.药品id=[4]" & _
            "       AND a.性质=1 " & _
            "       AND a.库房id+0 = "
        If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
            strKc = strKc & IIf(mlng源库房ID = 0, "[1]", "[2]")
        End If
        
        If mstr盘点时间 <> "" Then
            strKc = strKc & _
                "   UNION ALL " & _
                "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次, a.供药单位ID 上次供应商ID,max(a.成本价) 上次采购价, " & _
                "           -SUM (DECODE (a.入出系数, 1, a.实际数量*a.付数, -a.实际数量*a.付数)) AS 实际数量, " & _
                "           -SUM (DECODE (a.入出系数, 1, a.零售金额, -a.零售金额)) AS 实际金额," & _
                "           -SUM (DECODE (a.入出系数, 1, a.差价, -a.差价)) AS 实际差价, " & _
                "           -SUM (DECODE (a.入出系数, 1, a.填写数量*a.付数, -a.填写数量*a.付数)) AS 可用数量, " & _
                "           Max(零售价) as 零售价,0 as 平均成本价,a.批号,a.产地 , A.效期,a.灭菌效期,a.生产日期,a.批准文号" & _
                "   FROM 药品收发记录 a " & _
                "   Where  a.药品id+0=[4]  " & _
                "           AND a.库房id + 0 ="
            If mlng源库房ID <> 0 Or mlng目库房ID <> 0 Then
                strKc = strKc & IIf(mlng源库房ID = 0, "[1]", "[2]")
            End If
            strKc = strKc & " AND a.审核日期 >[5] " & _
                " GROUP BY A.库房id, a.药品id,a.供药单位id, A.批次, A.批号, A.产地, A.效期, A.灭菌效期,a.生产日期,a.批准文号 "
        End If
        
        strKc = "" & _
            "   Select 库房id,药品id,nvl(批次,0) 批次,max(上次批号) 上次批号,min(灭菌效期) as 灭菌失效期,max(上次供应商ID) 上次供应商ID, " & _
            "       Sum(nvl(可用数量,0)) 可用数量," & _
            "       Sum(实际数量) 实际数量," & _
            "       Sum(实际金额) 实际金额," & _
            "       Sum(实际差价) 实际差价," & _
            "       max(上次采购价) 上次采购价,Max(零售价) as 零售价,max(平均成本价) as 平均成本价, " & _
            "        Min(灭菌效期) 灭菌效期,Min(效期) 效期,max(上次产地) 上次产地 ,max(上次生产日期) 上次生产日期,max(批准文号) as 批准文号,1 As 性质" & _
            "   From (" & strKc & ")" & _
            "   Group by 库房id,药品id,nvl(批次,0) "
                 
                 
        '1.实价:如果分批:库存中的零售价 ,否则为实际金额/实际数量*比例系数=售价
        '2.定价:收费价目中的现价=售价
        '3.成本价:
        '       a.如果存在上次购价,则以上次购价为准
        '       b.如果不存在上次购价,则以（库存金额-库存差价)/库存数量为准.
        
        gstrSQL = gstrSQL & strTmp & _
                 IIf(mblnStock, " to_char(K.实际金额," & gOraFmt_Max.FM_金额 & ") as 库存金额,", "to_char(''," & gOraFmt_Max.FM_金额 & ") 库存金额,") & _
                 IIf(mblnStock, " to_char(K.实际差价," & gOraFmt_Max.FM_金额 & ") as 库存差价", "to_char(''," & gOraFmt_Max.FM_金额 & ") 库存差价") & ",K.灭菌效期 灭菌失效期," & _
                 IIf(mblnStock, "to_char(Decode(nvl(M.是否变价,0),0,G.现价,decode(nvl(K.零售价,0),0,nvl(K.实际金额,0)/decode(K.实际数量,null,1,0,1,K.实际数量),K.零售价))" & IIf(mbln散装单位, "", "*nvl(D.换算系数,1)") & "," & gOraFmt_Max.FM_零售价 & ") 售价,", "to_char(0," & gOraFmt_Max.FM_零售价 & ") 售价,") & _
        " to_char(k.平均成本价," & gOraFmt_Max.FM_成本价 & ") as 成本价, " & _
                 IIf(mblnStock, "to_char(decode(nvl(K.上次采购价,0),0,(nvl(K.实际金额,0)-nvl(K.实际差价,0))/decode(K.实际数量,null,1,0,1,K.实际数量),K.上次采购价)" & IIf(mbln散装单位, "", "*nvl(D.换算系数,1)") & "," & gOraFmt_Max.FM_成本价 & ") 上次购价", "to_char(0," & gOraFmt_Max.FM_成本价 & ") 上次购价") & _
        "       ,K.上次产地 产地,k.上次生产日期 生产日期 ,k.上次供应商ID,k.批准文号 " & _
        " From 部门表 P, 材料特性 D, " & IIf(mstr盘点时间 <> "", "(" & strKc & ")", " 药品库存") & " K,收费项目目录 M,收费价目 G " & _
        " Where     K.库房ID = P.ID And D.材料ID = K.药品ID " & _
        " And K.库房ID " & IIf(mstr盘点时间 <> "", "+0=", "=") & IIf(mlng源库房ID = 0, "[1]", "[2]") & _
        " And K.药品ID " & IIf(mstr盘点时间 <> "", "+0=", "=") & "[4] And K.性质=1 " & _
        " And D.材料id=G.收费细目ID(+) " & _
        " And D.材料ID=M.ID And (M.站点=[6] or M.站点 is null) " & _
        " And m.Id = g.收费细目id And (Sysdate Between g.执行日期 And Nvl(g.终止日期, Sysdate)) " & _
        GetPriceClassString("G")
                 
        If mbln盘点单 Then
            gstrSQL = gstrSQL & " And (K.实际数量<>0 Or K.实际金额<>0 Or K.实际差价<>0)"
        Else
            gstrSQL = gstrSQL & " And K.实际数量<>0 "
        End If
        
'        If mlng供应商ID <> 0 Then gstrSQL = gstrSQL & " And K.上次供应商ID=[3]"
        
        If gSystem_Para.P156_出库算法 = 0 Then
            gstrSQL = gstrSQL & " Order by RID,批次"
        Else
            gstrSQL = gstrSQL & " Order by RID,失效期,批次"
        End If
        
        Dim dtDate As Date
        If mstr盘点时间 <> "" Then
            dtDate = CDate(mstr盘点时间)
        Else
            dtDate = Now
        End If
        
        Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng目库房ID, mlng源库房ID, mlng供应商ID, mlngLastSelect材料ID, dtDate, gstrNodeNo, gstrPriceClass)
          
        Dim blnState As Boolean
        With Msf批次
            If Not mrsStock.EOF Then
                Set .DataSource = mrsStock
                .ColWidth(0) = 0
            Else
                .Clear
                .Rows = 2
            End If
            
            Call SetFormat(0, mrsStock.EOF)
            If mbln空批次 And mrsStock.RecordCount <> 0 Then
                If .Row > 2 Then
                    .Row = 2
                End If
            End If
            blnState = ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And Not mrsStock.EOF
            If mbln显示批次 = True And blnState = True Then '只有允许显示批次列表后再查询具体批次信息，否则无用
                .Visible = True
            Else
                .Visible = False
            End If

            Msf批次_EnterCell
        End With
        Form_Resize
    End If
    
    With Msf批次
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 And Msf批次.Visible = True Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 6)) > 0 Then
                        .TextMatrix(i, 6) = "有"
                    Else
                        .TextMatrix(i, 6) = "无"
                    End If
                    .TextMatrix(i, 7) = ""
                    .TextMatrix(i, 8) = ""
                Next
            End If
        End If
    End With
    
    '设置按钮状态
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "材料ID=" & Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And mblnPrice Then
            If mbln显示批次 = False Then
                cmd确定.Enabled = True
            Else
                cmd确定.Enabled = blnState
            End If
        Else
            cmd确定.Enabled = True
        End If
        If chkContinue.Value = 1 Then cmd确定.Enabled = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Msf材料规格_GotFocus()
    With msf材料规格
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf材料规格_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf材料规格_DblClick
End Sub

Private Sub Msf材料规格_LostFocus()
    With msf材料规格
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit02_S
        If .Top + y < msf材料规格.Top + 1000 Then Exit Sub
        If .Top + y > tvwClass.Height - 1500 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    pic选定区.Move pic选定区.Left, pic选定区.Top + y, pic选定区.Width, pic选定区.Height - y
    
End Sub

Private Sub picUpDown01_Click()
    If pic选定区.Tag = "展开" Then
        pic选定区.Tag = "收缩"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
        picSplit02_S.MousePointer = 7

        
        picSplit02_S.Top = Me.tvwClass.Height / 2
        pic选定区.Top = picSplit02_S.Top + picSplit02_S.Height
        pic选定区.Height = tvwClass.Height - pic选定区.Top
        
    Else
        pic选定区.Tag = "展开"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
        
        Form_Resize
    End If
End Sub

Private Sub pic选定区_Resize()
    lbl选定.Width = pic选定区.Width
    
    With vsf选定
        .Top = lbl选定.Height
        .Left = 0
        .Width = lbl选定.Width
        .Height = pic选定区.Height - lbl选定.Height
    End With
End Sub


Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strTmp As String, StrGroupBy As String
    Dim strKc As String
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    Dim i As Integer
    Dim int入库产地取值方式 As Integer
    
    '读出该卫材用途分类的规格卫材
    '    如果目标库房不明确（如其他出库和领用）或是制剂室，则不限制卫材材质
    '    如果目标库房不明确（如其他出库和领用）或是卫材库、制剂室，则不限制服务对象
    '    如果目标库房是服务于门诊病人，则门诊用药可以进入；
    '    如果目标库房是服务于住院病人，则住院用药可以进入；
    
    On Error GoTo ErrHandle

    If mlngModule = 1712 Or mlngModule = 1714 Then
        int入库产地取值方式 = Val(zlDatabase.GetPara(268, glngSys))
    End If
    
    If mlng目库房ID <> 0 Then
        mbln只显示跟踪材料 = 判断只具备发料部门(mlng目库房ID)
        If mbln只显示跟踪材料 = False Then
            mbln只显示跟踪材料 = 判断只具备发料部门(mlng源库房ID)
        End If
    End If
    
    '判断虚拟库房
    gstrSQL = "select count(*) rec from 部门性质说明 where 工作性质='虚拟库房' and 部门id=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断虚拟库房", mlng目库房ID)
    If rsTmp!rec = 1 And (mobjOut.Name = "frmPurchaseCard" Or mobjOut.Name = "frmOtherInputCard") Then
        blnVirtualStock = True
    End If
    
    '对列头排顺序
    gstrSQL = "" & _
        " Select " & IIf(mstr盘点时间 <> "", "/*+ Rule*/", "") & " D.诊疗id,D.材料id,D.分类ID,D.编码,D.通用名称,D.商品名,D.规格,D.产地,d.批准文号,d.注册证号,x.名称 As 上次供应商 ,to_char(D.售价," & gOraFmt_Max.FM_零售价 & ") 售价,to_char(d.成本价," & gOraFmt_Max.FM_成本价 & ") as 最新成本价,D.散装单位,D.换算系数,D.包装单位," & _
          IIf(mblnStock, "  to_char(S.可用数量 " & IIf(mbln散装单位, "", "/D.换算系数") & "," & gOraFmt_Max.FM_数量 & ") 可用数量, to_char(S.库存数量 " & _
          IIf(mbln散装单位, "", "/D.换算系数") & "," & gOraFmt_Max.FM_数量 & ") 库存数量 ,to_char(S.库存金额," & gOraFmt_Max.FM_金额 & ") 库存金额,to_char(S.库存差价," & gOraFmt_Max.FM_金额 & ") 库存差价,", "to_char(''," & gOraFmt_Max.FM_数量 & ") 可用数量,to_char(''," & gOraFmt_Max.FM_数量 & ") 库存数量,to_char(''," & gOraFmt_Max.FM_金额 & ") 库存金额,to_char(''," & gOraFmt_Max.FM_金额 & ") 库存差价,") & _
        "     D.最大效期 有效期,D.灭菌效期,S.灭菌失效期,D.一次性材料,D.无菌性材料,D.库房分批,D.在用分批,D.时价,to_char(D.指导批发价," & gOraFmt_Max.FM_零售价 & ") 指导批发价,D.指导差价率,E.库房货位" & _
        " From "

    '材料信息，材料目录
    If mbln只显示跟踪材料 Then
        gstrSQL = gstrSQL & _
                "     (Select Distinct u.诊疗id,u.材料id,H.分类ID,V.编码,V.名称 As 通用名称,B.名称 As 商品名,V.规格," & IIf(int入库产地取值方式 = 0, "decode(u.上次产地,null,v.产地,u.上次产地)", "decode(v.产地,null,u.上次产地,v.产地)") & " as 产地,u.批准文号,u.注册证号,V.计算单位 as 散装单位,U.包装单位," & _
                "          To_Char(U.换算系数," & GFM_XS & " ) 换算系数,nvl(To_Char(U.灭菌效期,'9999990'),0) 灭菌效期,nvl(To_Char(U.最大效期,'9999990'),0) 最大效期," & _
                "          Decode(U.库房分批,1,'是','否') 库房分批,Decode(U.在用分批,1,'是','否') 在用分批,Decode(U.一次性材料,1,'是','否')  一次性材料,Decode(U.无菌性材料,1,'是','否') 无菌性材料,Decode(V.是否变价,1,'是','否') 时价," & _
                "          U.指导批发价 ,To_Char(U.指导差价率," & GFM_CJL & " ) 指导差价率,现价 as 售价,u.成本价,Nvl(u.上次供应商id, 0) As 上次供应商id " & _
                "      From 材料特性 U,收费项目目录 V,诊疗项目目录 H," & _
                "       (SELECT 收费细目id, 执行科室id FROM 收费执行科室 WHERE 执行科室ID" & IIf(mlng源库房ID <> 0, "+0=[1]", IIf(mlng目库房ID <> 0, "+0=[2]", " Is Not NULL")) & ") K," & _
                "       (Select 收费细目ID, 执行科室ID From 收费执行科室 Where 执行科室ID" & IIf(mlng目库房ID <> 0, "+0=[2]", IIf(mlng源库房ID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                "       收费项目别名 B, 收费价目 P " & _
                "      Where U.材料id=v.id And (v.站点=[5] or v.站点 is null) And U.诊疗id=H.id  And V.ID = B.收费细目id(+) And B.性质(+) = 3 " & _
                "          AND U.材料id=K.收费细目ID " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                "          AND U.材料id=i.收费细目ID " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                           IIf(mbln只显示跟踪材料, " and  U.跟踪在用 =1 ", IIf(mblnTrackUsing = True, " and  U.跟踪在用 =0 ", "")) & " And v.Id = p.收费细目id And (Sysdate Between p.执行日期 And Nvl(p.终止日期, Sysdate)) " & _
                           GetPriceClassString("P")
    Else
        gstrSQL = gstrSQL & _
                "     (Select Distinct u.诊疗id,u.材料id,H.分类ID,V.编码,V.名称 As 通用名称,B.名称 As 商品名,V.规格," & IIf(int入库产地取值方式 = 0, "decode(u.上次产地,null,v.产地,u.上次产地)", "decode(v.产地,null,u.上次产地,v.产地)") & " as 产地,u.批准文号,u.注册证号,V.计算单位 as 散装单位,U.包装单位," & _
                "          To_Char(U.换算系数," & GFM_XS & " ) 换算系数,nvl(To_Char(U.灭菌效期,'9999990'),0) 灭菌效期,nvl(To_Char(U.最大效期,'9999990'),0) 最大效期," & _
                "          Decode(U.库房分批,1,'是','否') 库房分批,Decode(U.在用分批,1,'是','否') 在用分批,Decode(U.一次性材料,1,'是','否')  一次性材料,Decode(U.无菌性材料,1,'是','否') 无菌性材料,Decode(V.是否变价,1,'是','否') 时价," & _
                "          U.指导批发价,To_Char(U.指导差价率," & GFM_CJL & " ) 指导差价率,现价 售价,u.成本价,Nvl(u.上次供应商id, 0) As 上次供应商id " & _
                "      From 材料特性 U,收费项目目录 V,诊疗项目目录 H," & _
                "     (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(mlng源库房ID <> 0, "=[1]", IIf(mlng目库房ID <> 0, "=[2]", " Is Not NULL")) & " ) K," & _
                "     (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(mlng目库房ID <> 0, "+0=[2]", IIf(mlng源库房ID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                "      收费项目别名 B, 收费价目 P  " & _
                "      Where U.材料id=v.id And (v.站点=[5] or v.站点 is null) And U.诊疗id=H.id  And V.ID = B.收费细目id(+) And B.性质(+) = 3 " & _
                "          AND U.材料id=K.收费细目ID " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                "          AND U.材料id=i.收费细目ID " & IIf(mbln盘无存储库房材料, "(+)", "") & _
                           IIf(mbln只显示跟踪材料, " and  U.跟踪在用 =1 ", IIf(mblnTrackUsing = True, " and  U.跟踪在用 =0 ", "")) & " And v.Id = p.收费细目id And (Sysdate Between p.执行日期 And Nvl(p.终止日期, Sysdate)) " & _
                           GetPriceClassString("P")
    End If
    
    If mlng目库房ID > 0 Then
        gstrSQL = gstrSQL & " And" & _
            "     ( exists(select 1 from 部门性质说明 where 工作性质 In('制剂室','卫材库','发料部门','虚拟库房')  and 部门id=[2])  " & _
            "       or v.服务对象=(select distinct '1' from 部门性质说明 where 工作性质 like '发料部门' and 部门id=[2] and 服务对象 in(1,3))" & _
            "       or v.服务对象=(select distinct '2' from 部门性质说明 where 工作性质 like '发料部门' and 部门id=[2] and 服务对象 in(2,3)))"
    End If
    
    '查找指定材料用途分类的规格材料
    If Not (Node.Key Like "Root") Then
        gstrSQL = gstrSQL & " And H.分类ID IN (Select ID from 诊疗分类目录 where 类型=7 Start With ID=" & Mid(Node.Key, 3) & " Connect By Prior ID=上级ID)"
    Else
        gstrSQL = gstrSQL & " "
    End If
    
    '只查找未停用的规格卫材
    If mstr盘点时间 <> "" Then      '对盘点时间来说，如果盘点时间小于停用的时间也应该显示出来
        gstrSQL = gstrSQL & " And (V.撤档时间 Is Null Or V.撤档时间>[4])"
    Else
        gstrSQL = gstrSQL & " And (V.撤档时间 Is Null Or To_char(V.撤档时间,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    '只查找指定材质分类的规格卫材
    gstrSQL = gstrSQL & IIf(blnVirtualStock, " and nvl(u.高值材料,0)=1 and nvl(u.跟踪病人,0)=1 and nvl(u.跟踪在用,0)=1 and nvl(u.在用分批,0)=1", "") & " ) D,"

    '取库存
    '20060731:刘兴宏加入，主要解决盘点时间的库存
    strKc = "" & _
        "   SELECT a.库房id, a.药品id, NVL (a.批次, 0) AS 批次,a.上次供应商ID," & _
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
        strKc = strKc & " AND a.审核日期 >[4] " & _
            " GROUP BY A.库房id, a.药品id,a.供药单位id, A.批次, A.批号, A.产地, A.效期, A.灭菌效期 "
    End If
    
    If mblnStock Then
        gstrSQL = gstrSQL & " (Select 药品id as 材料id,min(灭菌效期) as 灭菌失效期 , Sum(nvl(可用数量,0)) 可用数量," & _
                " Sum(nvl(实际数量,0)) 库存数量," & _
                " Sum(nvl(实际金额,0)) 库存金额," & _
                " Sum(nvl(实际差价,0)) 库存差价"
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
        gstrSQL = gstrSQL & " And 库房ID" & IIf(mstr盘点时间 <> "", "+0=", "=") & IIf(mlng源库房ID = 0, "[2]", "[1]") & "  Group By 药品id) S"
    Else
        gstrSQL = gstrSQL & " Group By 药品id) S"
    End If
    
    gstrSQL = gstrSQL & ",(Select 材料id,库房ID,库房货位 From 材料储备限额" & _
              " Where 库房ID=" & IIf(mintEditState = 2, "[1]", "[2]") & ") E,供应商 X"
    
    '总条件
    gstrSQL = gstrSQL & " Where D.材料ID=S.材料ID"
    
    If mbln仅显示库存物资 And mblnStock Then
        gstrSQL = gstrSQL & " And S.可用数量<>0"
    Else
        '当系统参数“材料出库库存检查”为不足禁止时，不提库存为零
        If Not (mintStockCheck = 2 And In_编辑状态 = 2) Or mbln盘点单 Or Not mblnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_编辑状态 = 2 Then gstrSQL = gstrSQL & " And S.可用数量<>0"
    End If
    
    gstrSQL = gstrSQL & " And D.材料ID=E.材料ID(+) And d.上次供应商id = x.Id(+) Order By D.编码"
    Dim dtDate As Date
    If mstr盘点时间 <> "" Then
        dtDate = CDate(mstr盘点时间)
    Else
        dtDate = Now
    End If
    Set mrsCard = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng源库房ID, mlng目库房ID, mlng供应商ID, dtDate, gstrNodeNo)
    
    With msf材料规格
        If Not mrsCard.EOF Then
            Set .DataSource = mrsCard
        Else
            .Clear
            .Rows = 2
        End If
        Call SetFormat(1, mrsCard.EOF)
    End With
    cmd确定.Enabled = (mrsCard.EOF <> True)
    
    Call Msf材料规格_EnterCell
    
    With msf材料规格
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "显示对方库存") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, mCol.库存数量)) > 0 Then
                        .TextMatrix(i, mCol.库存数量) = "有"
                    Else
                        .TextMatrix(i, mCol.库存数量) = "无"
                    End If
                    .TextMatrix(i, mCol.库存金额) = ""
                    .TextMatrix(i, mCol.库存差价) = ""
                Next
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
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
            
            .Fields.Append "时价", adDouble, 2, adFldIsNullable
            .Fields.Append "批次", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "效期", adDate, , adFldIsNullable
            .Fields.Append "灭菌失效期", adDate, , adFldIsNullable
            .Fields.Append "生产日期", adDate, , adFldIsNullable
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
    Dim blnEof As Boolean               '是否存在批次库存
    Dim i As Integer
    
    CombinateRec = False
    
    On Error GoTo ErrHandle:
    
    If chkContinue.Value = 0 Then '组装一条数据
        With mrsCard
            If .RecordCount <> 0 Then .MoveFirst
            .Find "材料ID=" & Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID))
            
            If .EOF Then
                MsgBox "发生内部错误！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mbln显示批次 = True Then '只有显示批次的情况下才需要做如下操作
                If ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) And In_编辑状态 = 2 Then
                    With mrsStock
                        If .RecordCount <> 0 Then .MoveFirst
                        .Find "批次=" & Val(Msf批次.TextMatrix(Msf批次.Row, 2))
                        If .EOF Then
                            blnEof = True
                            If mblnPrice Then
                                MsgBox "无库存数据！", vbInformation, gstrSysName
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
            !散装单位 = mrsCard!散装单位
            !换算系数 = mrsCard!换算系数
            !包装单位 = mrsCard!包装单位
            !最大效期 = mrsCard!有效期
            !灭菌效期 = mrsCard!灭菌效期
            !灭菌失效期 = mrsCard!灭菌失效期
            !一次性材料 = IIf(mrsCard!一次性材料 = "是", 1, 0)
            !无菌性材料 = IIf(mrsCard!无菌性材料 = "是", 1, 0)
            !库房分批 = IIf(mrsCard!库房分批 = "是", 1, 0)
            !在用分批 = IIf(mrsCard!在用分批 = "是", 1, 0)
            
            !时价 = IIf(mrsCard!时价 = "是", 1, 0)
            
            '出库且分批
            If In_编辑状态 = 2 And ((mint分批 = 3 And mint库房 <> 3) Or (mint分批 = 1 And mint库房 = 1) Or (mint分批 = 2 And mint库房 = 2)) Then
                If mbln显示批次 = True Then '只有显示批次情况下才需要做如下操作，否则无意义
                    If Msf批次.TextMatrix(Msf批次.Row, 3) = "新增批次卫生材料" Then
                        !批次 = -1
                    Else
                        If Not blnEof Then
                            !批次 = Val(mrsStock!批次)
                            !批号 = mrsStock!批号
                            !效期 = mrsStock!失效期
                            !灭菌失效期 = mrsStock!灭菌失效期
                            !产地 = mrsStock!产地
                            !生产日期 = mrsStock!生产日期
                            !批准文号 = mrsStock!批准文号
                            !供药单位ID = mrsStock!上次供应商id
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
                    !批准文号 = mrsCard!批准文号
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
    
    If cmd确定.Enabled = False Then Exit Function
    
    If mbln显示批次 = False Then
        CheckData = True
        Exit Function '如果是不显示批次模式则直接不检查库存
    End If
    
    If Msf批次.Visible Then
        'lng供应商ID不为零，表示退货，无库存时不准继续
        If mlng供应商ID <> 0 Then
            intCol = GetCol(Msf批次, "上次供应商ID")
            If intCol < 0 Then Exit Function
            If Val(Msf批次.TextMatrix(Msf批次.Row, intCol)) <> 0 And mlng供应商ID <> Val(Msf批次.TextMatrix(Msf批次.Row, intCol)) Then
                MsgBox "你选择的退货商不是该卫生材料的供应商，不能继续操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mblnStock Then
            DblCurStock = Val(Msf批次.TextMatrix(Msf批次.Row, 5))
        Else
            DblCurStock = Get可用库存(Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID)), Val(Msf批次.TextMatrix(Msf批次.Row, 2)))
        End If
    Else
        If Not mrsCard.EOF Then
            If mblnStock Then
                DblCurStock = Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.可用数量))
            Else
                DblCurStock = Get可用库存(Val(msf材料规格.TextMatrix(msf材料规格.Row, mCol.材料ID)))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '如果源库房与目库房为空，则表明是材料目录自己在进行常规设置，不判断
    If (mlng源库房ID = 0 And mlng目库房ID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '如果是盘点单调用材料选择器，则不需判断，直接退出
    If mbln盘点单 Then
        CheckData = True
        Exit Function
    End If
    If Msf批次.Visible Or mbln时价 Then
        If (DblCurStock <> 0) Or Not mblnPrice Or Msf批次.TextMatrix(Msf批次.Row, 3) = "新增批次卫生材料" Then CheckData = True: Exit Function
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

Public Function ShowMe(ByVal frmMain As Form, ByVal 编辑模式 As Integer, Optional ByVal 源库房 As Long, _
                    Optional ByVal 目库房 As Long = 0, Optional ByVal 使用部门 As Long = 0, Optional ByVal Bln检测库存 As Boolean = True, _
                    Optional ByVal bln检查批次或时价 As Boolean = True, Optional ByVal mbln盘点单据 As Boolean = False, Optional ByVal bln增加空批次 As Boolean = False, _
                    Optional ByVal bln显示库存 As Boolean = True, Optional ByVal lng供应商 As Long = 0, Optional ByVal bln散装单位 As Boolean = True, _
                    Optional bln只显示跟踪材料 As Boolean = False, _
                    Optional str盘点时间 As String = "", _
                    Optional bln仅显示库存物资 As Boolean = False, _
                    Optional lngModule As Long = 0, _
                    Optional ByVal bln盘无存储库房材料 As Boolean = False, _
                    Optional ByVal strPrivs As String = "", _
                    Optional ByVal bln显示批次 As Boolean = True, Optional bln是否过滤 As Boolean = True) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------
    '功能:显示选择器
    '参数:  bln检查库存-遵守批次卫材及时价卫材零库存不准出库原则，可强制允许not (批次 or 时价) 卫材出库
    '       bln检查批次或时价:允许零库存的批次卫材及时价卫材出库
    '       mlng供应商ID:不为零表示退货
    '返回:所选择的记录集
    '------------------------------------------------------------------------------------------------------
    
    mbln散装单位 = bln散装单位
    mlngModule = lngModule
    If lngModule = 1717 Then    '1717:卫材领用
        mblnTrackUsing = IIf(Val(zlDatabase.GetPara("跟踪在用", glngSys, lngModule, "0")) = 1, True, False)
    Else
        mblnTrackUsing = False
    End If
    
    With Me
        .In_编辑状态 = 编辑模式
        .In_源库房 = 源库房
        .In_目库房 = 目库房
        .In_部门 = 使用部门
        .In_MainFrm = frmMain
        mbln盘点单 = mbln盘点单据
        mbln空批次 = bln增加空批次
        mblnCheck = Bln检测库存
        mblnPrice = bln检查批次或时价
        mblnStock = bln显示库存
        mlng供应商ID = lng供应商
        mbln只显示跟踪材料 = bln只显示跟踪材料
        mstr盘点时间 = str盘点时间
        '修改:刘兴宏   Bug:12792    日期:2008-05-08 15:03:47
        mbln仅显示库存物资 = bln仅显示库存物资
        mbln盘无存储库房材料 = bln盘无存储库房材料
        mstrPrivs = strPrivs
        mbln显示批次 = bln显示批次
        mbln是否过滤 = bln是否过滤
        .Show 1, frmMain
    End With
    Set ShowMe = mrsReturn.Clone
End Function

Public Function Get可用库存(ByVal lng材料ID As Long, Optional ByVal lng批次 As Long = 0) As Single
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        " Select Sum(A.可用数量" & mstrUnitString & ") 可用数量,Sum(A.实际数量" & mstrUnitString & ") 实际数量,sum(A.实际金额) 实际金额,sum(A.实际差价) 实际差价 " & _
              " From 药品库存 A,材料特性 B " & _
              " Where A.药品ID=B.材料ID and A.性质=1  And A.药品ID=[1]" & IIf(lng批次 = 0, "", " And Nvl(A.批次,0)=[2]")
    
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
