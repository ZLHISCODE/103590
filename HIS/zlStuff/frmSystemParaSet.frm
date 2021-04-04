VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSystemParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫材参数设置"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmSystemParaSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8460
      TabIndex        =   0
      Top             =   375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8460
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8490
      TabIndex        =   2
      Top             =   5520
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   -315
      Top             =   5865
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
            Picture         =   "frmSystemParaSet.frx":000C
            Key             =   "Limit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystemParaSet.frx":045E
            Key             =   "bm"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stbPage 
      Height          =   7365
      Left            =   75
      TabIndex        =   5
      Top             =   90
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   12991
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本参数(&0)"
      TabPicture(0)   =   "frmSystemParaSet.frx":09F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl出库算法"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl卫材条码前缀"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl条码前缀提示"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl定价单位"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmb出库算法"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt卫材条码前缀"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "vsfParameter"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optUnit(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optUnit(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "卫材流向控制(&1)"
      TabPicture(1)   =   "frmSystemParaSet.frx":0A14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(23)"
      Tab(1).Control(1)=   "Image1(0)"
      Tab(1).Control(2)=   "vsf流向"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "库存检查(&2)"
      TabPicture(2)   =   "frmSystemParaSet.frx":0A30
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1(1)"
      Tab(2).Control(1)=   "lbl提示"
      Tab(2).Control(2)=   "vsf库房检查"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "虚拟库房对照(&3)"
      TabPicture(3)   =   "frmSystemParaSet.frx":0A4C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(1)=   "Image1(3)"
      Tab(3).Control(2)=   "vsf对照"
      Tab(3).ControlCount=   3
      Begin VB.OptionButton optUnit 
         Caption         =   "散装单位"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   17
         Top             =   6705
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "包装单位"
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   16
         Top             =   6705
         Width           =   1425
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf对照 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   8055
         _cx             =   14208
         _cy             =   10398
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VSFlex8Ctl.VSFlexGrid vsf流向 
         Height          =   6135
         Left            =   -74640
         TabIndex        =   13
         Top             =   1080
         Width           =   7815
         _cx             =   13785
         _cy             =   10821
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSystemParaSet.frx":0A68
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
      Begin VSFlex8Ctl.VSFlexGrid vsfParameter 
         Height          =   5055
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   8055
         _cx             =   14208
         _cy             =   8916
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSystemParaSet.frx":0B6C
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
      Begin VB.TextBox txt卫材条码前缀 
         Height          =   300
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   10
         Top             =   6240
         Width           =   2145
      End
      Begin VB.ComboBox cmb出库算法 
         Height          =   300
         ItemData        =   "frmSystemParaSet.frx":1096
         Left            =   1800
         List            =   "frmSystemParaSet.frx":1098
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5797
         Width           =   2145
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf库房检查 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   8055
         _cx             =   14208
         _cy             =   10821
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSystemParaSet.frx":109A
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
      Begin VB.Label lbl定价单位 
         Caption         =   "卫材指导批发价定价单位"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label lbl条码前缀提示 
         AutoSize        =   -1  'True
         Caption         =   "（请录入2-8位数字或字母）"
         Height          =   180
         Left            =   3960
         TabIndex        =   11
         Top             =   6300
         Width           =   2250
      End
      Begin VB.Label lbl卫材条码前缀 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卫材条码前缀"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   6270
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   -74520
         Picture         =   "frmSystemParaSet.frx":1190
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "    在这里可以选择卫材发料部门，卫材仓库，科室虚拟库房三者对应的关系。"
         Height          =   315
         Left            =   -73920
         TabIndex        =   8
         Top             =   600
         Width           =   7080
      End
      Begin VB.Label lbl出库算法 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卫材出库优先算法"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5820
         Width           =   1440
      End
      Begin VB.Label lbl提示 
         Caption         =   "    在这里可以选择各库房是否检查库存及库存检查方式。当库房选中时双击或按“C”键可改变库房的检查方式。"
         Height          =   435
         Left            =   -74040
         TabIndex        =   4
         Top             =   525
         Width           =   7080
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   -74700
         Picture         =   "frmSystemParaSet.frx":1A5A
         Top             =   450
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   -74670
         Picture         =   "frmSystemParaSet.frx":20DB
         Stretch         =   -1  'True
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "控制材料在不同库房间的流通方向"
         Height          =   180
         Index           =   23
         Left            =   -74025
         TabIndex        =   3
         Top             =   720
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmSystemParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnChange As Boolean
Private mblnLoad As Boolean
Private mblnChkClick As Boolean     '是否是程序设置chk控件的值.
Private mintOldChkValue As Integer      '存放下库存参数的旧值.
Private mrs对照 As New ADODB.Recordset
Private mstrPrivs As String
Private Const mlngColor As Long = &H8000000F        '不能修改的列将背景颜色改成灰色
Private Const MCON_LNGCOLOR As Long = &H80000005    '能修改的列背景颜色
Private mstrOld加成销售 As String         '记录旧的 时价卫材入库按扣前加成销售
    
Private Enum mPara
    mint时价卫生材料以加价率入库 = 1
    mint按批次申领卫生材料 = 2
    mint卫材填单下可用库存
    mint负数出库按最后一次入库的成本价计算差价
    mint卫材按分段加成率入库
    mint不严格控制卫材指导批价和指导售价
    mint时价卫材入库按扣前加成销售
    mint允许向发料部门领用卫生材料
    mint时价卫材直接确定售价
    mint外购入库单需要核查
    mint时价卫材入库取上次售价
    mintCount = 12
End Enum

Private Enum m流向
    mint所在库房 = 1
    mint对方库房 = 2
    mint流向
    mint所在库房id
    mint对方库房id
    mintCount = 6
End Enum

Private Enum m库房检查
    mintid = 0
    mint编码
    mint名称
    mint检查方式
    mintCount = 4
End Enum

Private Enum m库房对照
    mint科室id = 0
    mint发料部门 = 1
    mint库房id = 2
    mint卫材仓库 = 3
    mint虚拟库房id
    mint虚拟库房
    mint启用
    mintCount = 7
End Enum

'Private Sub bill对照_AfterAddRow(Row As Long)
'    bill对照.TextMatrix(bill对照.Rows - 1, 6) = "√"
'End Sub

'Private Sub bill对照_cboClick(ListIndex As Long)
'    Dim intRow As Integer
'    Dim lng科室id As Long
'
'    With vsf对照
'        If ListIndex < 0 Then Exit Sub
'        If .Col = 1 Then
'            .TextMatrix(.Row, 0) = .ItemData(ListIndex)
'        ElseIf .Col = 3 Then
'            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
'        ElseIf .Col = 5 Then
'            .TextMatrix(.Row, 4) = .ItemData(ListIndex)
'        End If
'
'        lng科室id = Val(.TextMatrix(.Row, 0))
'
'        For intRow = 1 To .Rows - 1
'            If Val(.TextMatrix(intRow, 0)) > 0 Then
'                If Val(.TextMatrix(intRow, 0)) = lng科室id And intRow <> .Row Then
'                    .TextMatrix(.Row, 0) = ""
'                    .TextMatrix(.Row, 1) = ""
'                    .TextMatrix(.Row, 2) = ""
'                    .TextMatrix(.Row, 3) = ""
'                    .TextMatrix(.Row, 4) = ""
'                    .TextMatrix(.Row, 5) = ""
'                    .TextMatrix(.Row, 6) = ""
'                    Exit For
'                End If
'            End If
'        Next
'    End With
'
'    mblnChange = True
'End Sub

Private Sub bill对照_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim lng科室id As Long
    
    With vsf流向
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .Col = 1 Then
                .TextMatrix(.Row, 0) = .ItemData(.ListIndex)
            ElseIf .Col = 3 Then
                .TextMatrix(.Row, 1) = .ItemData(.ListIndex)
            ElseIf .Col = 5 Then
                .TextMatrix(.Row, 3) = .ItemData(.ListIndex)
            End If
            
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, 0)) > 0 Then
                    If Val(.TextMatrix(intRow, 0)) = lng科室id And intRow <> .Row Then
                        .TextMatrix(.Row, 0) = ""
                        .TextMatrix(.Row, 1) = ""
                        .TextMatrix(.Row, 2) = ""
                        .TextMatrix(.Row, 3) = ""
                        .TextMatrix(.Row, 4) = ""
                        .TextMatrix(.Row, 5) = ""
                        .TextMatrix(.Row, 6) = ""
                        Exit For
                    End If
                End If
            Next
        End If
        mblnChange = True
    End With
End Sub

Private Sub vsf对照_ChangeEdit()
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim str名称 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
    
    If InStr(1, vsf对照.EditText, "-") <= 0 Then Exit Sub
    strID = Mid(vsf对照.EditText, 1, InStr(1, vsf对照.EditText, "-") - 1)
    str名称 = Mid(vsf对照.EditText, InStr(1, vsf对照.EditText, "-") + 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门查询", strID, str名称)
    If rsTemp.RecordCount > 0 Then
        With vsf对照
            If .Col = m库房对照.mint发料部门 Then
                .TextMatrix(.Row, m库房对照.mint科室id) = rsTemp!Id
            ElseIf .Col = m库房对照.mint卫材仓库 Then
                .TextMatrix(.Row, m库房对照.mint库房id) = rsTemp!Id
            ElseIf .Col = m库房对照.mint虚拟库房 Then
                .TextMatrix(.Row, m库房对照.mint虚拟库房id) = rsTemp!Id
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf对照_DblClick()
    With vsf对照
        If .Col = m库房对照.mint启用 Then
            If .TextMatrix(.Row, m库房对照.mint启用) = "" Then
                .TextMatrix(.Row, m库房对照.mint启用) = "√"
            Else
                .TextMatrix(.Row, m库房对照.mint启用) = ""
            End If
        End If
    End With
End Sub



Private Sub bill流向_cboClick(ListIndex As Long)
   
    With vsf流向
        If ListIndex < 0 Then Exit Sub
        If .Col = 0 Then
            .RowData(.Row) = .ItemData(ListIndex)
        ElseIf .Col = 1 Then
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
        End If
'        .TextMatrix(.Row, .Col) = .CboText
        
        If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
    End With
    mblnChange = True
End Sub

Private Sub bill流向_cboKeyDown(KeyCode As Integer, Shift As Integer)
  With vsf流向
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                .RowData(.Row) = .ItemData(.ListIndex)
            End If
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
        End If
        mblnChange = True
    End With
End Sub

'Private Sub bill流向_DblClick(Cancel As Boolean)
'    '处理最后一列的变化
'    With vsf流向
'        If .MouseRow = 0 Then Exit Sub
'        If .MouseCol <> .Cols - 1 Then Exit Sub
'        Select Case Left(.TextMatrix(.Row, .Col), 1)
'            Case "1"
'                .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
'            Case "2"
'                .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
'            Case Else
'                .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
'        End Select
'        mblnChange = True
'End With
'End Sub

Private Sub bill流向_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)

    With vsf流向
            If .Col = 2 Then
                '报警值列只处理回车键
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TextMatrix(.Row, 2) = "" Then
                    '到下一个控件
                    OS.PressKey vbKeyTab
                End If
            ElseIf .Col >= 3 Then
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            End If
    End With
End Sub

Private Sub bill流向_KeyPress(KeyAscii As Integer)
    With vsf流向
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                            Case "2"
                                .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                        End Select
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                        mblnChange = True
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                        mblnChange = True
                End Select
                mblnChange = True
            End If
    End With
End Sub

Private Function Check移库单() As Boolean
    '功能:检查移库单是否存在未审核的单据
    Dim rsTemp As New ADODB.Recordset
    Dim blnTemp As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From 药品收发记录 where 单据=19 and 审核日期 is null and rownum<=3 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then
        blnTemp = True
    Else
        blnTemp = rsTemp.RecordCount = 0
    End If
    If blnTemp = False Then
        ShowMsgBox "还存在申领单或移库单的未审单据,请先处理后再设置!"
    End If
    Check移库单 = blnTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check入库单() As Boolean
    '功能:检查移库单是否存在未审核的单据
    Dim rsTemp As New ADODB.Recordset
    Dim blnTemp As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From 药品收发记录 where 单据=15 and 审核日期 is null and rownum<=3 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    
    If rsTemp.EOF Then
        blnTemp = True
    Else
        blnTemp = rsTemp.RecordCount = 0
    End If
    If blnTemp = False Then
        ShowMsgBox "还存在未审核的外购入库单,请处理后再设置!"
    End If
    Check入库单 = blnTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub chkcheck_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
'End Sub
'
'Private Sub chk申领下库存_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))

End Sub

Private Sub cmdOk_Click()
    If ISValid() = False Then Exit Sub
    If Save数据() = False Then Exit Sub
    Call InitSystemPara
    mblnChange = False
    Unload Me
End Sub
Private Function ISValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngIndex As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim i As Integer
    Dim j As Integer
    
    With vsf流向
        
        For lngRow = 1 To .Rows - 1
            If (.TextMatrix(lngRow, 1) = "" Or .TextMatrix(lngRow, 2) = "" Or .TextMatrix(lngRow, 3) = "") And lngRow <> .Rows - 1 Then
                MsgBox "第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 1
                stbPage.Tab = 1
                Exit Function
            End If
'            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
'                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
'                .Row = lngRow
'                .Col = 0
'                stbPage.Tab = 1
'                Exit Function
'            End If
            If .TextMatrix(lngRow, 1) = .TextMatrix(lngRow, 2) And lngRow <> .Rows - 1 Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 1
                stbPage.Tab = 1
                Exit Function
            End If
            For j = 1 To .Rows - 1
                If .TextMatrix(i, 1) = .TextMatrix(j, 1) And .TextMatrix(i, 2) = .TextMatrix(j, 2) And i <> j Then
                    MsgBox "第" & i & "行与第" & j & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = i
                    .Col = 1
                    stbPage.Tab = 1
                    Exit Function
                End If
            Next
            
'            For lngTemp = lngRow + 1 To .Rows - 1
'                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
'                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
'                    .Row = lngTemp
'                    .Col = 1
'                    stbPage.Tab = 1
'                    Exit Function
'                End If
'            Next
        Next
    End With
    
    With vsfParameter
        If mintOldChkValue <> Val(.TextMatrix(mPara.mint卫材填单下可用库存, 1)) Then
            If Check移库单 = False Then Exit Function
        End If
        If IIf(.TextMatrix(mPara.mint时价卫材入库按扣前加成销售, 1) = "", 0, .TextMatrix(mPara.mint时价卫材入库按扣前加成销售, 1)) <> mstrOld加成销售 Then '记录原卫材入库按扣前加成销售和现在的值是否一样
            '需要验证是否入库了的
            If Check入库单 = False Then Exit Function
        End If
    End With
    ISValid = True
End Function

Private Function Save数据() As Boolean
    Dim str流向  As String
    Dim i As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim str科室id As String
    Dim arr库房 As Variant
    Dim str所在库房id As String
    Dim str对方库房id As String
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim bln次数 As Boolean
    Dim arrSQL  As Variant

    On Error GoTo ErrHandle

    gcnOracle.BeginTrans
    
    With vsfParameter
        Call zlDatabase.SetPara(82, IIf(.TextMatrix(mPara.mint时价卫生材料以加价率入库, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(83, IIf(.TextMatrix(mPara.mint按批次申领卫生材料, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(95, IIf(.TextMatrix(mPara.mint卫材填单下可用库存, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(120, IIf(.TextMatrix(mPara.mint负数出库按最后一次入库的成本价计算差价, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(121, IIf(.TextMatrix(mPara.mint卫材按分段加成率入库, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(123, IIf(.TextMatrix(mPara.mint不严格控制卫材指导批价和指导售价, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(127, IIf(.TextMatrix(mPara.mint时价卫材入库按扣前加成销售, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(132, IIf(.TextMatrix(mPara.mint允许向发料部门领用卫生材料, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(136, IIf(.TextMatrix(mPara.mint时价卫材直接确定售价, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(140, IIf(.TextMatrix(mPara.mint外购入库单需要核查, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(229, IIf(.TextMatrix(mPara.mint时价卫材入库取上次售价, 1) = "1", 1, 0), glngSys, 0)
    End With
    
    Call zlDatabase.SetPara(88, IIf(optUnit(0).Value = True, 0, 1), glngSys, 0)
    Call zlDatabase.SetPara(156, IIf(cmb出库算法.ListIndex = -1, 0, cmb出库算法.ListIndex), glngSys, 0)

    If zlStr.IsHavePrivs(mstrPrivs, "设置条码前缀") = True Then
        Call zlDatabase.SetPara(159, IIf(Trim(txt卫材条码前缀.Text) = "", "", UCase(Trim(txt卫材条码前缀.Text))), glngSys, 0)
    End If

    strTemp = ""
    arrSQL = Array()
    With vsf流向
        For lngRow = 1 To .Rows - 1
            str流向 = Left(.TextMatrix(lngRow, m流向.mint流向), 1)
            If str流向 = "" Then str流向 = "3"
            
            str所在库房id = ""
            str对方库房id = ""
            
            If .TextMatrix(lngRow, m流向.mint所在库房id) = "" And lngRow <> .Rows - 1 Then
                gstrSQL = "select id from 部门表 where 编码=[1]"
                strID = Mid(.TextMatrix(lngRow, m流向.mint所在库房), 1, InStr(1, .TextMatrix(lngRow, m流向.mint所在库房), "-") - 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID)
                If rsTemp.RecordCount > 0 Then
                    str所在库房id = rsTemp!Id
                End If
            Else
                str所在库房id = .TextMatrix(lngRow, m流向.mint所在库房id)
            End If
            
            If .TextMatrix(lngRow, m流向.mint对方库房id) = "" And lngRow <> .Rows - 1 Then
                strID = Mid(.TextMatrix(lngRow, m流向.mint对方库房), 1, InStr(1, .TextMatrix(lngRow, m流向.mint对方库房), "-") - 1)
                gstrSQL = "select id from 部门表 where 编码=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "对方库房查询", strID)
                If rsTemp.RecordCount > 0 Then
                    str对方库房id = rsTemp!Id
                End If
            Else
                str对方库房id = .TextMatrix(lngRow, m流向.mint对方库房id)
            End If
            If str所在库房id <> "" Or str对方库房id <> "" Then
                If LenB(StrConv(strTemp & str所在库房id & "," & str对方库房id & "," & str流向 & ",", vbFromUnicode)) >= 4000 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strTemp
                    strTemp = str所在库房id & "," & str对方库房id & "," & str流向 & ","
                    bln次数 = True
                Else
                    strTemp = strTemp & str所在库房id & "," & str对方库房id & "," & str流向 & ","
                End If
            End If
        Next
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTemp
    End With
    
    For i = 0 To UBound(arrSQL)
        If bln次数 = True Then
            If i = 0 Then
                Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "删除调价记录")
            Else
                Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',1" & ")", "删除调价记录")
            End If
        Else
            Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "删除调价记录")
        End If
    Next

    '保存库房检查
    gstrSQL = ""
    With vsf库房检查
        For i = 1 To .Rows - 1
            gstrSQL = gstrSQL & .TextMatrix(i, m库房检查.mintid) & "," & Switch(.TextMatrix(i, m库房检查.mint检查方式) = "0-不检查", "0", .TextMatrix(i, m库房检查.mint检查方式) = "1-检查，不足提醒", "1", .TextMatrix(i, m库房检查.mint检查方式) = "2-检查，不足禁止", "2") & ","
        Next
    End With

    gstrSQL = "Zl_材料出库检查_insert('" & gstrSQL & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption


    '保存虚拟库房对照
    If zlStr.IsHavePrivs(mstrPrivs, "设置虚拟库房对照") = True Then
        strTemp = ""
        With vsf对照
            For i = 1 To .Rows - 1
                If .TextMatrix(i, m库房对照.mint启用) = "√" And Val(.TextMatrix(i, m库房对照.mint科室id)) > 0 And Val(.TextMatrix(i, m库房对照.mint库房id)) > 0 And Val(.TextMatrix(i, m库房对照.mint虚拟库房id)) > 0 Then
                    If InStr(1, "," & str科室id & ",", "," & Val(.TextMatrix(i, m库房对照.mint科室id)) & ",") = 0 Then
                        str科室id = IIf(str科室id = "", "", str科室id & ",") & .TextMatrix(i, m库房对照.mint科室id)
                        strTemp = IIf(strTemp = "", "", strTemp & "|") & .TextMatrix(i, m库房对照.mint科室id) & "," & .TextMatrix(i, m库房对照.mint库房id) & "," & .TextMatrix(i, m库房对照.mint虚拟库房id)
                    End If
                End If
            Next
        End With

        gstrSQL = "Zl_虚拟库房对照_Update('" & strTemp & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If

    '保存完毕，事务提交
    gcnOracle.CommitTrans
    Save数据 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    mblnLoad = True
    mblnChkClick = False
    mstrPrivs = gstrPrivs

    '进行初始化
    Call InitCtrl
    Call Load库房检查
    Call Load材料流向
    Call Load虚拟库房对照
    Call initVsfPara
    
'    Me.lvwCheckMed.Sorted = True
    
    '加载其他参数
    Call LoadPara
    Call SetColor
    
    If zlStr.IsHavePrivs(mstrPrivs, "设置虚拟库房对照") = False Then
        stbPage.TabVisible(3) = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "设置条码前缀") = False Then
        lbl卫材条码前缀.Visible = False
        txt卫材条码前缀.Visible = False
        lbl条码前缀提示.Visible = False
        lbl定价单位.Top = lbl卫材条码前缀.Top
        optUnit(0).Top = lbl定价单位.Top
        optUnit(1).Top = optUnit(0).Top
    End If
    
    '恢复列宽
    RestoreFlexState vsf流向, App.ProductName & "\" & Me.Name & "\材料流向"
    
    '初始化成功
    mblnChange = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPara()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnChkClick = True
'    chkCheck(5).Tag = ""
    Set rsTemp = ReturnParaData(glngSys, "82,83,88,95,120,121,123,127,132,136,140,156,159,229")
    With rsTemp
        Do While Not .EOF
            Select Case zlStr.Nvl(!参数号, 0)
                Case 88
                    optUnit(0).Value = IIf(zlStr.Nvl(!参数值, 0) = 0, True, False)
                    optUnit(1).Value = IIf(zlStr.Nvl(!参数值, 0) = 0, False, True)
                Case 156
                    '卫材出库算法
                    If zlStr.Nvl(!参数值, 0) = 1 Then
                        cmb出库算法.ListIndex = 1
                    Else
                        cmb出库算法.ListIndex = 0
                    End If
                Case 159
                    '卫材条码前缀
                    txt卫材条码前缀.Text = IIf(IsNull(!参数值), "", !参数值)
            End Select
            With vsfParameter
                For i = 1 To .Rows - 1
                    If zlStr.Nvl(rsTemp!参数号, 0) = .TextMatrix(i, 0) And rsTemp!参数值 = 1 Then
                        .TextMatrix(i, 1) = 1
                    End If
                    If zlStr.Nvl(rsTemp!参数号, 0) = "95" Then
                        mintOldChkValue = Val(rsTemp!参数值)
                    End If
                    If zlStr.Nvl(rsTemp!参数号, 0) = "127" Then
                        mstrOld加成销售 = Val(rsTemp!参数值)
                    End If
                Next
            End With
            
            .MoveNext
        Loop
    End With
    mblnChkClick = False
End Function
Private Sub InitCtrl()
    Dim lngIndex As Long
    
    With vsf流向
        .Cols = m流向.mintCount  '多了一列隐藏列
        .ColWidth(0) = 0
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    With vsf对照
        .Cols = m库房对照.mintCount
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, m库房对照.mint科室id) = "科室id"
        .TextMatrix(0, m库房对照.mint发料部门) = "发料部门"
        .TextMatrix(0, m库房对照.mint库房id) = "卫材仓库ID"
        .TextMatrix(0, m库房对照.mint卫材仓库) = "卫材仓库"
        .TextMatrix(0, m库房对照.mint虚拟库房id) = "虚拟库房ID"
        .TextMatrix(0, m库房对照.mint虚拟库房) = "虚拟库房"
        .TextMatrix(0, m库房对照.mint启用) = "启用"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 0
        .ColWidth(5) = 2000
        .Editable = flexEDKbdMouse
    End With
    
    With cmb出库算法
        .Clear
        .AddItem "0-按批次先进先出"
        .AddItem "1-按效期最近先出"
    End With
End Sub

Private Sub Load材料流向()
    '功能:装入材料流向数据
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTemp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With vsf流向
        '首向装入库房
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('卫材库','制剂室','虚拟库房','发料部门') " & _
                   "   and  b.部门ID=a.ID and " & Where撤档时间("A") & _
                   " order by 编码"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        .Rows = 1
        .Rows = 2
        strTemp = ""
        If Not rsTemp.EOF Then
            rsTemp.MoveFirst
            For i = 1 To rsTemp.RecordCount
                strTemp = strTemp & rsTemp!编码 & "-" & rsTemp!名称 & "|"
                rsTemp.MoveNext
            Next
        End If
        .ColComboList(m流向.mint所在库房) = strTemp
        .ColComboList(m流向.mint对方库房) = strTemp
        
        .ColComboList(m流向.mint流向) = "1-所在库房可流向对方库房|2-对方库房可流向所在库房|3-两库房间可双向流通"
'        Do Until rsTemp.EOF
'            .AddItem rsTemp("编码") & "-" & rsTemp("名称")
'            .ItemData(.NewIndex) = rsTemp("ID")
'            rsTemp.MoveNext
'        Loop
        
        '装入流向控制数据
        gstrSQL = "select A.所在库房ID,A.对方库房ID,A.流向" & _
                ",B.编码 as 所在编码,B.名称 as 所在名称,C.编码 as 对方编码,C.名称 as 对方名称 " & _
                " from 材料流向控制 A,部门表 B,部门表 C " & _
                " where A.所在库房ID= B.ID and A.对方库房ID=C.ID " & _
                "   and (b.撤档时间=to_date('3000-1-1','yyyy-mm-dd') or b.撤档时间 is null) " & _
                " order by b.编码, c.编码 "
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            
'            .RowData(lngRow) = rsTemp("所在库房ID")
'            .TextMatrix(lngRow, 0) = rsTemp("所在编码") & "-" & rsTemp("所在名称")
'            .TextMatrix(lngRow, 1) = rsTemp("对方编码") & "-" & rsTemp("对方名称")
'            .TextMatrix(lngRow, 2) = rsTemp("对方库房ID")
'            .TextMatrix(lngRow, 3) = Switch(rsTemp("流向") = 1, "1-所在库房可流向对方库房", _
'                                            rsTemp("流向") = 2, "2-对方库房可流向所在库房", _
'                                                          True, "3-两库房间可双向流通")
                                                          
            .TextMatrix(lngRow, m流向.mint所在库房) = IIf(IsNull(rsTemp!所在库房id), "", rsTemp!所在编码 & "-" & rsTemp!所在名称)
            .TextMatrix(lngRow, m流向.mint所在库房id) = rsTemp!所在库房id
            .TextMatrix(lngRow, m流向.mint对方库房) = IIf(IsNull(rsTemp!对方库房ID), "", rsTemp!对方编码 & "-" & rsTemp!对方名称)
            .TextMatrix(lngRow, m流向.mint对方库房id) = rsTemp!对方库房ID
            .TextMatrix(lngRow, m流向.mint流向) = Switch(rsTemp("流向") = 1, "1-所在库房可流向对方库房", _
                                            rsTemp("流向") = 2, "2-对方库房可流向所在库房", _
                                                          True, "3-两库房间可双向流通")
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        For i = 0 To vsf流向.Rows - 1
            .RowHeight(i) = 300
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load虚拟库房对照()
    '功能:装入卫材虚拟库房对照关系
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With vsf对照
        '取所有发料部门，卫材库，虚拟库房
        mrs对照.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码, b.工作性质 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('卫材库','发料部门','虚拟库房') " & _
                   " and  b.部门ID=a.ID and " & Where撤档时间("A") & " order by 编码"
        
        zlDatabase.OpenRecordset mrs对照, gstrSQL, Me.Caption
        
        '装入目前的虚拟库房对照关系
        gstrSQL = "Select b.Id As 科室id, b.编码 || '-' || b.名称 As 发料部门, c.Id As 库房id, c.编码 || '-' || c.名称 As 卫材仓库," & _
                  " d.Id As 虚拟库房id,d.编码 || '-' || d.名称 As 虚拟库房 " & _
                  "From 虚拟库房对照 A, 部门表 B, 部门表 C, 部门表 D " & _
                  "Where a.科室id = b.Id And a.库房id = c.Id And a.虚拟库房id = d.Id " & _
                  "  And (b.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd') or b.撤档时间 is null) " & _
                  "Order by b.编码 "
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .TextMatrix(lngRow, m库房对照.mint科室id) = rsTemp!科室id
            .TextMatrix(lngRow, m库房对照.mint发料部门) = rsTemp!发料部门
            .TextMatrix(lngRow, m库房对照.mint库房id) = rsTemp!库房ID
            .TextMatrix(lngRow, m库房对照.mint卫材仓库) = rsTemp!卫材仓库
            .TextMatrix(lngRow, m库房对照.mint虚拟库房id) = rsTemp!虚拟库房id
            .TextMatrix(lngRow, m库房对照.mint虚拟库房) = rsTemp!虚拟库房
            .TextMatrix(lngRow, m库房对照.mint启用) = "√"
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs对照 = Nothing
    
    SaveFlexState vsf流向, App.ProductName & "\" & Me.Name & "\材料流向"
    If mblnChange = False Then Exit Sub
    
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

'Private Sub lvwCheckMed_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If lvwCheckMed.SortKey = ColumnHeader.Index - 1 Then
'        lvwCheckMed.SortOrder = IIf(lvwCheckMed.SortOrder = lvwAscending, lvwDescending, lvwAscending)
'    Else
'        lvwCheckMed.SortKey = ColumnHeader.Index - 1
'        lvwCheckMed.SortOrder = lvwAscending
'    End If
'End Sub

Private Sub stbPage_Click(PreviousTab As Integer)
    Select Case stbPage.Tab
        Case 0
            vsfParameter.SetFocus
        Case 1
            vsf流向.SetFocus
        Case 2
            vsf库房检查.SetFocus
        Case Else
    End Select
End Sub


Private Sub Load库房检查()
    '功能：初始化库房
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim objItem As ListItem
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.编码, B.名称, NVL(C.检查方式, 0) 检查方式" & vbCrLf & _
        " FROM 部门性质说明 A, 部门表 B, 材料出库检查 C" & vbCrLf & _
        " WHERE A.部门ID = B.ID AND A.部门ID = C.库房ID(+) AND" & vbCrLf & _
        "      A.工作性质 IN" & vbCrLf & _
        "      ('卫材库','制剂室','发料部门','虚拟库房') " & vbCrLf & _
        "     And (b.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd') or b.撤档时间 is null) " & vbCrLf & _
        " GROUP BY B.ID,B.编码, B.名称, NVL(C.检查方式, 0)" & vbCrLf & _
        " ORDER BY B.编码 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.vsf库房检查.Rows = 1
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        vsf库房检查.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
'            Set objItem = Me.lvwCheckMed.ListItems.Add(, "C_" & rsTmp!Id, "[" & zlStr.Nvl(rsTmp!编码) & "]", "bm", "bm")
'            objItem.SubItems(1) = zlStr.Nvl(rsTmp!名称)
'            objItem.SubItems(2) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
'            objItem.Tag = rsTmp!Id
            With vsf库房检查
                .TextMatrix(i, m库房检查.mintid) = rsTmp!Id
                .TextMatrix(i, m库房检查.mint名称) = rsTmp!名称
                .TextMatrix(i, m库房检查.mint编码) = rsTmp!编码
                .TextMatrix(i, m库房检查.mint检查方式) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
            End With
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'Private Sub lvwCheckMed_DblClick()
'    If Not Me.lvwCheckMed.SelectedItem Is Nothing Then
'        lvwCheckMed.SelectedItem.SubItems(2) = Switch(lvwCheckMed.SelectedItem.SubItems(2) = "0-不检查", "1-检查，不足提醒", lvwCheckMed.SelectedItem.SubItems(2) = "1-检查，不足提醒", "2-检查，不足禁止", lvwCheckMed.SelectedItem.SubItems(2) = "2-检查，不足禁止", "0-不检查")
'    End If
'End Sub

'Private Sub lvwCheckMed_KeyPress(KeyAscii As Integer)
'    If UCase(Chr(KeyAscii)) = "C" Then
'        Call lvwCheckMed_DblClick
'    End If
'End Sub

Private Sub optUnit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If optUnit(1).Enabled Then optUnit(1).SetFocus
    Else
        stbPage.Tab = 1
    End If
End Sub

Private Sub txt卫材条码前缀_Change()
    txt卫材条码前缀.Text = UCase(txt卫材条码前缀.Text)
End Sub

Private Sub txt卫材条码前缀_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Exit Sub
        End If
    End Select
    KeyAscii = 0
End Sub


Private Sub txt卫材条码前缀_LostFocus()
    If Len(txt卫材条码前缀) > 8 Then
        txt卫材条码前缀.Text = Mid(txt卫材条码前缀.Text, 1, 8)
    End If
End Sub

Private Sub initVsfPara()
    Dim i As Integer
    '初始化常规界面vsflexgrid控件
    With vsfParameter
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        
        .GridLines = flexGridInset
        .GridColor = &H0&
        .AllowUserResizing = flexResizeColumns
        .Rows = mPara.mintCount
        .Cols = 4
        .ExtendLastCol = True '最后一列填充满
        .WordWrap = True
        .AutoSize 3, 3, False, 0 = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ColDataType(1) = flexDTBoolean
        .ScrollBars = flexScrollBarVertical '将横向滚动条取消掉
        .ColHidden(0) = True
    End With
    
    With vsf库房检查
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        .ExtendLastCol = True
        .Cols = 4
        For i = 0 To .Rows - 1
            .RowHeight(i) = 300
        Next
        .ColComboList(m库房检查.mint检查方式) = "0-不检查|1-检查，不足提醒|2-检查，不足禁止"
        .ColHidden(0) = True
    End With
    
    With vsf对照
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        For i = 0 To .Rows - 1
            .RowHeight(i) = 300
        Next
    End With
End Sub

Private Sub vsfParameter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With vsfParameter
            If .CellBackColor = mlngColor Then
                Exit Sub
            End If
            Select Case .Row
                Case mPara.mint卫材填单下可用库存
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mint按批次申领卫生材料, 0, mPara.mint按批次申领卫生材料, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mint按批次申领卫生材料, 1) = "1"
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mint按批次申领卫生材料, 0, mPara.mint按批次申领卫生材料, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mint时价卫材入库取上次售价
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mint时价卫生材料以加价率入库, 1) = ""
                        .TextMatrix(mPara.mint卫材按分段加成率入库, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mint时价卫生材料以加价率入库
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mint时价卫材入库取上次售价, 1) = ""
                        .TextMatrix(mPara.mint卫材按分段加成率入库, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mint卫材按分段加成率入库
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mint时价卫材入库取上次售价, 1) = ""
                        .TextMatrix(mPara.mint时价卫生材料以加价率入库, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case Else
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                    Else
                        .TextMatrix(.Row, 1) = ""
                    End If
            End Select
        End With
    End If
End Sub

Private Sub SetColor()
    '不可修改列设置颜色为灰色
    With vsfParameter
        If .TextMatrix(mPara.mint卫材填单下可用库存, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mint按批次申领卫生材料, 0, mPara.mint按批次申领卫生材料, .Cols - 1) = mlngColor
        End If
        If .TextMatrix(mint时价卫材入库取上次售价, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mint时价卫生材料以加价率入库, 1) = ""
            .TextMatrix(mPara.mint卫材按分段加成率入库, 1) = ""
        End If
        If .TextMatrix(mint时价卫生材料以加价率入库, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mint卫材按分段加成率入库, 0, mPara.mint卫材按分段加成率入库, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mint时价卫材入库取上次售价, 1) = ""
            .TextMatrix(mPara.mint卫材按分段加成率入库, 1) = ""
        End If
        If .TextMatrix(mint卫材按分段加成率入库, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mint时价卫生材料以加价率入库, 0, mPara.mint时价卫生材料以加价率入库, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mint时价卫材入库取上次售价, 0, mPara.mint时价卫材入库取上次售价, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mint时价卫生材料以加价率入库, 1) = ""
            .TextMatrix(mPara.mint时价卫材入库取上次售价, 1) = ""
        End If
    End With
End Sub

Private Sub vsf对照_EnterCell()
    Dim strTemp As String
    
    With vsf对照
        If .Col = m库房对照.mint启用 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = 1 Then
            mrs对照.Filter = "工作性质='发料部门'"
        ElseIf .Col = 3 Then
            mrs对照.Filter = "工作性质='卫材库'"
        ElseIf .Col = 5 Then
            mrs对照.Filter = "工作性质='虚拟库房'"
        End If
        
'        .Clear
        strTemp = ""
        Do While Not mrs对照.EOF
            strTemp = strTemp & mrs对照("编码") & "-" & mrs对照("名称") & "|"
'            .AddItem mrs对照("编码") & "-" & mrs对照("名称")
'            .ItemData(.NewIndex) = mrs对照("ID")
            mrs对照.MoveNext
        Loop
        .ColComboList(.Col) = strTemp
    End With
End Sub

Private Sub vsf对照_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf对照
            If .Col = m库房对照.mint启用 - 1 And .TextMatrix(.Row, m库房对照.mint发料部门) <> "" And .TextMatrix(.Row, m库房对照.mint卫材仓库) <> "" And .TextMatrix(.Row, m库房对照.mint虚拟库房) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m库房对照.mint启用 - 1 And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 2
            End If
        End With
    End If
End Sub

Private Sub vsf库房检查_DblClick()
    With vsf库房检查
        If .Col = m库房检查.mint检查方式 Then
            Select Case .TextMatrix(.Row, m库房检查.mint检查方式)
                Case "0-不检查"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "1-检查，不足提醒"
                Case "1-检查，不足提醒"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "2-检查，不足禁止"
                Case "2-检查，不足禁止"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "0-不检查"
                Case Else
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "0-不检查"
            End Select
        End If
    End With
End Sub

Private Sub vsf流向_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strID As String
    Dim str名称 As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    With vsf流向
        strTemp = .TextMatrix(Row, Col)
        If strTemp <> "" Then
            If Col = m流向.mint所在库房 Then
                gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str名称 = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID, str名称)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, m流向.mint所在库房id) = rsTemp!Id
                End If
            ElseIf Col = m流向.mint对方库房 Then
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str名称 = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID, str名称)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, m流向.mint对方库房id) = rsTemp!Id
                End If
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


Private Sub vsf流向_DblClick()
    With vsf流向
        If .Col = m流向.mint流向 Then
            If .MouseRow = 0 Then Exit Sub
            .Editable = flexEDNone
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
            End Select
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsf流向_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsf流向.Rows > 1 Then
        vsf流向.RemoveItem vsf流向.Row
    End If
End Sub

Private Sub vsf流向_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf流向
            If .Col = m流向.mint流向 And .TextMatrix(.Row, m流向.mint所在库房) <> "" And .TextMatrix(.Row, m流向.mint对方库房) <> "" And .TextMatrix(.Row, m流向.mint流向) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m流向.mint流向 And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 1
            End If
        End With
    End If
End Sub





