VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmInAdviceSend 
   AutoRedraw      =   -1  'True
   Caption         =   "住院临嘱发送"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmInAdviceSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9615
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6615
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   3
      Top             =   6210
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInAdviceSend.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   900
      BandCount       =   2
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      _CBWidth        =   9615
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   2895
      MinHeight1      =   450
      Width1          =   2895
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "tbrSys"
      MinHeight2      =   450
      Width2          =   9195
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   450
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   794
         ButtonWidth     =   2514
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发送到住院"
               Key             =   "发送到住院"
               Description     =   "发送到住院"
               Object.ToolTipText     =   "发送到住院(Ctrl+1)"
               Object.Tag             =   "发送到住院"
               ImageKey        =   "发送"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发送到门诊"
               Key             =   "发送到门诊"
               Description     =   "发送到门诊"
               Object.ToolTipText     =   "发送到门诊(Ctrl+2)"
               Object.Tag             =   "发送到门诊"
               ImageKey        =   "发送"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrSys 
         Height          =   450
         Left            =   3120
         TabIndex        =   6
         Top             =   30
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "选项"
               Key             =   "选项"
               Description     =   "选项"
               Object.ToolTipText     =   "发送条件选项(F12)"
               Object.Tag             =   "选项"
               ImageKey        =   "选项"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "全选"
               Object.ToolTipText     =   "全选(Ctrl+A)"
               Object.Tag             =   "全选"
               ImageKey        =   "全选"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "全清"
               Object.ToolTipText     =   "全清(Ctrl+R)"
               Object.Tag             =   "全清"
               ImageKey        =   "全清"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出(ALT+X)"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   4605
      Width           =   9495
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   8
      Top             =   525
      Width           =   9435
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   825
      Width           =   9540
      _cx             =   16828
      _cy             =   6641
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
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInAdviceSend.frx":1A92
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
         Left            =   3435
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":1B2D
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":20C7
               Key             =   "F"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":2661
               Key             =   "签名"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9525
      _cx             =   16801
      _cy             =   2593
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":29B3
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":2BCD
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":2DE7
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3001
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":321B
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3435
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   960
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":364F
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3869
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3A83
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3C9D
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3EB7
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":40D1
            Key             =   "选项"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInAdviceSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String 'IN
Private mlng病人ID As Long 'IN
Private mlng主页ID As Long 'IN
Private mlng前提ID As Long 'IN
Private mblnSend As Boolean 'OUT:是否成功发送过。
Private mblnRefresh As Boolean 'OUT'是否需要刷新主界面

Private mcolStock As Collection '存放各个药品库房的出库检查方式
Private mrsPati As ADODB.Recordset '包含病人信息
Private mrsPrice As ADODB.Recordset '包含计价关系
Private mrsBill As ADODB.Recordset
Private mstrLike As String
Private mblnFirst As Boolean
Private mint简码 As Integer

'----------------------------------------------
Private Const COL_选择 = 0
Private Const COL_婴儿 = 1
Private Const COL_医嘱内容 = 2
Private Const COL_总量 = 3
Private Const COL_总量单位 = 4
Private Const COL_单量 = 5
Private Const COL_单量单位 = 6
Private Const COL_金额 = 7
Private Const COL_频率 = 8
Private Const COL_用法 = 9
Private Const COL_医生嘱托 = 10
Private Const COL_执行时间 = 11
Private Const COL_执行科室 = 12
Private Const COL_执行性质 = 13
Private Const COL_ID = 14 '隐藏列
Private Const COL_相关ID = 15
Private Const COL_医嘱状态 = 16
Private Const COL_病人科室ID = 17
Private Const COL_开嘱科室ID = 18
Private Const COL_开嘱医生 = 19
Private Const COL_开嘱时间 = 20
Private Const COL_诊疗类别 = 21
Private Const COL_诊疗项目ID = 22
Private Const COL_计价特性 = 23
Private Const COL_执行性质ID = 24
Private Const COL_执行科室ID = 25
Private Const COL_操作类型 = 26
Private Const COL_药品ID = 27
Private Const COL_剂量系数 = 28
Private Const COL_住院包装 = 29
Private Const COL_住院单位 = 30
Private Const COL_可否分零 = 31
Private Const COL_库存 = 32
Private Const COL_次数 = 33
Private Const COL_分解时间 = 34
Private Const COL_首次时间 = 35
Private Const COL_末次时间 = 36
Private Const COL_签名ID = 37

'-------------------------------------------------
Private Const COLP_行号 = 0
Private Const COLP_收费细目ID = 1
Private Const COLP_固定 = 2
Private Const COLP_变价 = 3
Private Const COLP_计价医嘱 = 4 '可见列
Private Const COLP_类别 = 5
Private Const COLP_收费项目 = 6
Private Const COLP_计价数量 = 7
Private Const COLP_数量 = 8
Private Const COLP_单位 = 9
Private Const COLP_单价 = 10
Private Const COLP_应收金额 = 11
Private Const COLP_实收金额 = 12
Private Const COLP_执行科室 = 13
Private Const COLP_费用类型 = 14
Private Const COLP_从项 = 15
Private Const COLP_收费类别 = 16
Private Const COLP_执行科室ID = 17
Private Const COLP_跟踪在用 = 18

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng前提ID As Long, blnRefresh As Boolean) As Boolean
'功能：发送医嘱
'参数：blnRefresh=是否刷新整个主界面
    mstrPrivs = strPrivs
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng前提ID = lng前提ID
    
    On Error Resume Next
    Me.Show 1, frmParent
    Err.Clear: On Error GoTo 0
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Dim str类别s As String
    
    If mblnFirst Then
        mblnFirst = False
        
        '读取发送清单
        Me.Refresh
        str类别s = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "住院临嘱发送类别", "")
        If Not LoadAdviceSend(str类别s) Then Unload Me: Exit Sub
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'功能：读取病人信息
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
        " Union ALL" & _
        " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
    strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
    
    strSQL = "Select A.住院号,A.姓名,A.性别,A.年龄,B.出院病床 as 床号," & _
        " B.当前病区ID,B.出院科室ID,B.费别,B.医疗付款方式,B.险类,C.剩余款," & _
        " B.状态,D.编码 as 付款码,Decode(D.编码,'1',1,Decode(Nvl(B.险类,0),0,0,1)) as 医保," & _
        " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
        " From 病人信息 A,病案主页 B,(" & strSQL & ") C,医疗付款方式 D" & _
        " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
        " And B.医疗付款方式=D.名称(+) And B.病人ID=[1] And B.主页ID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    lblPati.Caption = _
        "住院号:" & Nvl(mrsPati!住院号) & "　姓名:" & mrsPati!姓名 & "　性别:" & Nvl(mrsPati!性别) & "　年龄:" & Nvl(mrsPati!年龄) & _
        "　床号:" & Nvl(mrsPati!床号) & "　费别:" & Nvl(mrsPati!费别) & "　医疗付款方式:" & Nvl(mrsPati!医疗付款方式) & _
        "　剩余款:" & Format(Nvl(mrsPati!剩余款, 0), "0.00")
    
    '保险病人用红色显示
    If Not IsNull(mrsPati!险类) Then lblPati.ForeColor = vbRed
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("帮助"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("退出"))
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("全选"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("全清"))
    ElseIf KeyCode = vbKey1 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("发送为收费单").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("发送为收费单"))
        End If
    ElseIf KeyCode = vbKey2 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("发送为记帐单").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("发送为记帐单"))
        End If
    ElseIf KeyCode = vbKeyF7 Then '切换输入法
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
        
    If InStr(mstrPrivs, "发送门诊费用") = 0 Then
        tbrMain.Buttons("发送到门诊").Visible = False
        tbrMain.Buttons("发送到住院").Caption = "发送"
        tbrMain.Buttons("发送到住院").Tag = "发送"
        tbrMain.Buttons("发送到住院").ToolTipText = "发送(Ctrl+1)"
        cbr.Bands(1).MinWidth = cbr.Bands(1).MinWidth / 3
        cbr.Bands(1).Width = cbr.Bands(1).MinWidth
    End If
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔
    Select Case mint简码
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
   
    mblnSend = False
    mblnRefresh = False
    mblnFirst = True
    
    '各个库房药品出库检查方式
    Set mcolStock = InitStockCheck(2, True)
    
    '显示病人信息
    If Not GetPatiInfo Then Unload Me: Exit Sub
End Sub

Private Function GetStockCheck(ByVal lng库房ID As Long) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng库房ID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - vsPrice.Height - fraUD.Height - cbr.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = 0
    fraUD.Width = Me.ScaleWidth
    
    vsPrice.Left = 0
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = Me.ScaleWidth
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '释放私有及IN变量
    mstrPrivs = ""
    mlng主页ID = 0
    mlng病人ID = 0
    Set mrsPati = Nothing
    Set mrsPrice = Nothing
    Set mrsBill = Nothing
    Set mcolStock = Nothing
    
    gbln加班加价 = False
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsPrice.Top = vsPrice.Top + y
        vsPrice.Height = vsPrice.Height - y
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", _
            IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
        mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng发送号 As Long, strMsg As String, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "当前没有可以发送的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If Button.Key = "发送到住院" Then
        strMsg = "本次医嘱发送的费用将产生为住院记帐单据，确实要发送已选择的医嘱吗？"
    ElseIf Button.Key = "发送到门诊" Then
        strMsg = "本次医嘱发送的费用将产生为门诊收费单据，确实要发送已选择的医嘱吗？"
    End If
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    lng发送号 = SendAdvice(IIF(Button.Key = "发送到门诊", True, False))
    If lng发送号 <> 0 Then
        mblnSend = True
        '打印诊疗单据
        Call frmSendBillPrint.ShowMe(lng发送号, 2, Me, mlng前提ID)
        
        '如果全部发送完毕,则退出
        If vsAdvice.Rows = 2 Then
            If Val(vsAdvice.TextMatrix(1, COL_ID)) = 0 Then
                Unload Me: Exit Sub
            End If
        End If
        Call GetPatiInfo
    End If
End Sub

Private Sub tbrSys_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    Select Case Button.Key
        Case "全选"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_选择) = img16.ListImages("T").Picture
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "全清"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "选项"
            With frmInAdviceSendCond
                .Show 1, Me
                If .mblnOK Then
                    Call LoadAdviceSend(.mstr类别s)
                End If
            End With
        Case "帮助"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, _
    Optional rsUpload As ADODB.Recordset, Optional str医嘱IDs As String)
'功能：根据可见行的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_选择 Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            '取消选择时
            If Not (.Cell(flexcpData, lngRow, COL_选择) = 0 And Not .Cell(flexcpPicture, lngRow, COL_选择) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)))
                '1.清除对应的费用及发送记录填写
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "医嘱ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '因为要使用BookMark，因此恢复
                End If
                '2.清除对应的发送计价数量累计
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "医嘱ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.清除对应的医保上传单据号
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "医嘱ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
                '4.清除不发送的签名医嘱组ID
                If str医嘱IDs <> "" Then
                    str医嘱IDs = str医嘱IDs & ","
                    str医嘱IDs = Replace(str医嘱IDs, "," & i & ",", ",")
                    If str医嘱IDs <> "" Then
                        str医嘱IDs = Left(str医嘱IDs, Len(str医嘱IDs) - 1)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'功能：根据指定医嘱行，返回该医嘱中可见的行
    Dim lng组ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        '一并给药的定位到第一药品行
        If blnFirst Then
            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 _
                And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_相关ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowAdvicePrice(NewRow)
                
                '缺省选择计价医嘱(如果可能)
                Call ShowDefaultRow
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_选择 + 1 - .FixedCols Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = COL_医嘱内容 Then
            .AutoSize COL_医嘱内容
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_频率: lngRight = COL_用法
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
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_选择 Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_选择) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_选择) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_选择) = img16.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_选择) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        End If
    End With
End Sub

Private Sub ShowDefaultRow()
'功能：对于可以计价的医嘱,缺省增加一行并设置缺省计价医嘱
'说明：ComboList="#医嘱ID1;计价医嘱1|#医嘱ID2;计价医嘱2|..."
'      仅在第一次显示计价表和回车新增行时调用
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lng医嘱ID As Long, lng行号 As Long, str计价医嘱 As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_计价医嘱) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_计价医嘱), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_行号)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_收费细目ID)) <> 0 Then
                '第一次显示时缺省增加一行
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '不是第一次显示时缺省计价医嘱与上一行相同
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_固定)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_行号)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lng医嘱ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str计价医嘱 = Replace(arrCombo(i), "#" & lng医嘱ID & ";", "")
                lng行号 = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                If blnHave Then
                    If lng行号 = Val(.TextMatrix(lngRow - 1, COLP_行号)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            '模拟选中这个计价医嘱
            .TextMatrix(lngRow, COLP_行号) = lng行号
            .TextMatrix(lngRow, COLP_计价医嘱) = str计价医嘱
            .Cell(flexcpData, lngRow, COLP_计价医嘱) = .TextMatrix(lngRow, COLP_计价医嘱)
            
            '只有一个计价医嘱时不必停留
            If UBound(arrCombo) = 0 Then
                .Col = COLP_收费项目
            Else
                .Col = COLP_计价医嘱
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lng原嘱ID As Long, lng医嘱ID As Long
    Dim lng收费细目ID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_计价医嘱 Then
            '如果绑定了ComboData,TextMatrix取值就为ComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lng医嘱ID = .ComboData
                lng原嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
                lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                                
                '检查该计价医嘱是否已有相同收费细目
                If lng收费细目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng收费细目ID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """已经设置了收费项目""" & .TextMatrix(Row, COLP_收费项目) & """。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '原来的医嘱如果有从项至少要保留一个(主项是固定不可动的)
                If lng原嘱ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '标明输入了的计价医嘱部份
                i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                .TextMatrix(Row, COLP_行号) = i
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng收费细目ID <> 0 Then
                    '新选择的医嘱是否有从项决定修改后的项目是否从项
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 从项=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_从项) = IIF(blnHaveSub, "√", "")
                
                    '更新或增加记录集内容
                    If lng原嘱ID = 0 Then
                        mrsPrice.AddNew '加入
                    Else '更新
                        mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 收费细目ID=" & lng收费细目ID
                    End If
                    mrsPrice!医嘱ID = lng医嘱ID
                    If Val(vsAdvice.TextMatrix(i, COL_相关ID)) <> 0 Then
                        mrsPrice!相关ID = vsAdvice.TextMatrix(i, COL_相关ID)
                    Else
                        mrsPrice!相关ID = Null
                    End If
                    If lng原嘱ID = 0 Then
                        mrsPrice!收费细目ID = lng收费细目ID
                        mrsPrice!数量 = Val(.TextMatrix(Row, COLP_计价数量))
                        mrsPrice!单价 = Val(.TextMatrix(Row, COLP_单价))
                        mrsPrice!固定 = 0
                    End If
                    mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
            End If
        ElseIf Col = COLP_收费项目 Or Col = COLP_执行科室 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        ElseIf Col = COLP_计价数量 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf Col = COLP_单价 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), "0.00000")
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '根据可否编辑设置
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_计价医嘱 Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_收费项目 Or NewCol = COLP_执行科室 Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    '显示药品库存
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_行号))
            If lngRow <> 0 And .TextMatrix(NewRow, COLP_收费类别) <> "" Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_收费类别)) > 0 _
                    Or .TextMatrix(NewRow, COLP_收费类别) = "4" And Val(.TextMatrix(NewRow, COLP_跟踪在用)) = 1 Then
                    '显示药品及跟踪卫材的库存
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                        stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_医嘱内容) & "," & vsAdvice.TextMatrix(lngRow, COL_执行科室) & "可用库存:" & FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_库存)), 5) & vsAdvice.TextMatrix(lngRow, COL_住院单位)
                    Else
                        '同一个函数取:药品按住院单位,卫材按售价单位
                        stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "可用库存:" & _
                            FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID))), 5) & .TextMatrix(NewRow, COLP_单位)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明：返回的行号范围不包括给药途径的行号
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'功能：初始化清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;" & _
        "婴儿,550,1;医嘱内容,3000,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;金额,850,7;" & _
        "频率,1000,1;用法,1000,1;医生嘱托,1500,1;执行时间,1000,1;执行科室,850,1;执行性质,850,1;" & _
        "ID;相关ID;医嘱状态;病人科室ID;开嘱科室ID;开嘱医生;开嘱时间;诊疗类别;诊疗项目ID;计价特性;执行性质ID;" & _
        "执行科室ID;操作类型;药品ID;剂量系数;住院包装;住院单位;可否分零;库存;次数;分解时间;首次时间;末次时间;签名否"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_选择 + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "行号;收费细目ID;固定;变价;计价医嘱,2000,1;类别,650,1;收费项目,2000,1;计价数量,900,7;" & _
        "数量,800,7;单位,500,1;单价,1000,7;应收金额,1050,7;实收金额,1050,7;执行科室,1000,1;费用类型,850,1;" & _
        "从项,450,4;收费类别;执行科室ID;跟踪在用"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub DeleteCurRow(ByVal lngRow As Long, Optional ByVal blnDelCur As Boolean = True)
'功能：在处理待发送清单的过程中删除最近加入的行(含药疗或非药)
'参数：blnDelCur=是否删除当前行
    Dim lng医嘱ID As Long, lng相关ID As Long, i As Long
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
        lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
                
        '删除当前行
        If blnDelCur Then .RemoveItem lngRow
        
        '删除相关行
        If lng相关ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID _
                    Or Val(.TextMatrix(i, COL_ID)) = lng相关ID Then
                    .RemoveItem i
                End If
            Next
        Else
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub InitPriceRecordset()
'功能：初始化医嘱计价记录集
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "收费类别", adVarChar, 1
    mrsPrice.Fields.Append "收费细目ID", adBigInt
    mrsPrice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble
    mrsPrice.Fields.Append "单价", adDouble, , adFldIsNullable '变价价格
    mrsPrice.Fields.Append "在用", adInteger '卫材是否跟踪在用
    mrsPrice.Fields.Append "从项", adInteger
    mrsPrice.Fields.Append "固定", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset)
'初始化记录集
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-计价,2-签名,3-校对,4-发送,5-费用,6-发料
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '计价数量累计记录集
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsTotal.Fields.Append "项目ID", adBigInt
    rsTotal.Fields.Append "库房ID", adBigInt
    rsTotal.Fields.Append "数量", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    '医保上传记帐单
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsUpload.Fields.Append "NO", adVarChar, 10
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
End Sub

Private Function LoadAdvicePrice(ByVal lngRow As Long, rsSend As ADODB.Recordset, cur金额 As Currency) As Boolean
'功能：读取指定医嘱(仅当前行)的计价关系到临时记录集,并计算缺省发送金额(按费别打折)
'返回：cur金额=计算出的医嘱发送金额(非药变价未算,需要输入价格后才行)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean, i As Long
    Dim dbl数量 As Double, dbl单价 As Double
    Dim cur应收 As Currency, cur实收 As Currency
    Dim bln附加手术 As Boolean, lng项目ID As Long
    Dim lng主收入ID As Long, blnHaveSub As Boolean
    Dim lng执行科室ID As Long
    
    On Error GoTo errH
    
    cur金额 = 0
    With vsAdvice
        If InStr(",5,6,7,", rsSend!诊疗类别) > 0 Then
            '不为院外执行(自备药),药品不可能为叮嘱,且固定正常计价
            If Nvl(rsSend!执行性质, 0) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!医嘱ID = rsSend!ID
                mrsPrice!相关ID = rsSend!相关ID
                mrsPrice!收费类别 = rsSend!诊疗类别
                mrsPrice!收费细目ID = rsSend!药品ID
                mrsPrice!执行科室ID = rsSend!执行科室ID
                mrsPrice!数量 = 1 '药品固定为1
                mrsPrice!在用 = 0 '药品固定
                mrsPrice!固定 = 1 '药品固定
                mrsPrice!从项 = 0
                                
                '发送的零售数量
                If rsSend!诊疗类别 = "7" Then
                    '中药药房单位按不可分零处理:每付
                    If Nvl(rsSend!可否分零, 0) = 0 Then
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量)) / Nvl(rsSend!剂量系数, 1)
                    Else
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_单量)) / Nvl(rsSend!剂量系数, 1) / Nvl(rsSend!住院包装, 1)) * Nvl(rsSend!住院包装, 1)
                    End If
                Else
                    dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Nvl(rsSend!住院包装, 1)
                End If
                dbl数量 = Format(dbl数量, "0.00000")
                                
                '记录售价单价
                If Nvl(rsSend!是否变价, 0) = 0 Then
                    mrsPrice!单价 = Format(CalcPrice(rsSend!药品ID, , , True), "0.00000")
                Else '以售价计算药品时价,自备药时无对应药房
                    mrsPrice!单价 = Format(CalcDrugPrice(rsSend!药品ID, Nvl(rsSend!执行科室ID, 0), dbl数量, , True), "0.00000")
                End If
                mrsPrice.Update
                                
                '计算医嘱发送金额(按费别打折的实收金额)
                If Not IsNull(mrsPati!费别) Then
                    If Nvl(rsSend!是否变价, 0) = 0 Then
                        cur金额 = Format(CalcPrice(rsSend!药品ID, mrsPati!费别, dbl数量, , Nvl(rsSend!执行科室ID, 0)), gstrDec)
                    Else
                        cur金额 = Format(CalcDrugPrice(rsSend!药品ID, Nvl(rsSend!执行科室ID, 0), dbl数量, mrsPati!费别), "0.00000")
                    End If
                Else
                    If gbln加班加价 Then
                        '处理加班加价
                        If Nvl(rsSend!是否变价, 0) = 0 Then
                            dbl单价 = Format(CalcPrice(rsSend!药品ID), "0.00000")
                        Else '以售价计算药品时价,自备药时无对应药房
                            dbl单价 = Format(CalcDrugPrice(rsSend!药品ID, Nvl(rsSend!执行科室ID, 0), dbl数量), "0.00000")
                        End If
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * dbl单价, gstrDec)
                    Else
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * mrsPrice!单价, gstrDec)
                    End If
                End If
            End If
        Else
            '取诊疗收费关系中的对照(发送时才定计价):正常计价,不为叮嘱、院外执行
            If Nvl(rsSend!计价特性, 0) = 0 And InStr(",0,5,", Nvl(rsSend!执行性质, 0)) = 0 Then
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                bln附加手术 = (rsSend!诊疗类别 = "F" And Not IsNull(rsSend!相关ID))
                
                '先读取已有的计价
                strSQL = IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "")
                strSQL = _
                    " Select C.类别,A.收费细目ID as 收费项目ID,A.数量 as 收费数量,Nvl(E.固有对照,0) as 固有对照," & _
                    " B.收入项目ID,C.加班加价,B.加班加价率,Decode(C.是否变价,1,A.单价,B.现价)" & strSQL & " as 单价," & _
                    " C.是否变价,Nvl(A.从项,0) as 从项,D.跟踪在用,Nvl(A.执行科室ID,[3]) as 执行科室ID,C.屏蔽费别" & _
                    " From 病人医嘱计价 A,收费价目 B,收费项目目录 C,材料特性 D,诊疗收费关系 E" & _
                    " Where A.医嘱ID=[1] And E.诊疗项目ID(+)=[2] And A.收费细目ID=E.收费项目ID(+)" & _
                    " And A.收费细目ID=B.收费细目ID And A.收费细目ID=C.ID And A.收费细目ID=D.材料ID(+)" & _
                    " And C.服务对象 IN(2,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Order by 从项,A.收费细目ID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!ID), Val(rsSend!诊疗项目ID), Val(Nvl(rsSend!执行科室ID, 0)))
                
                '没有则读取默认的计价
                If rsTmp.EOF Then
                    strSQL = IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "")
                    strSQL = _
                        " Select C.类别,A.收费项目ID,A.收费数量,A.固有对照,B.收入项目ID," & _
                        " C.加班加价,B.加班加价率,Decode(C.是否变价,1,NULL,B.现价)" & strSQL & " as 单价," & _
                        " C.是否变价,Nvl(A.从属项目,0) as 从项,D.跟踪在用,[2] as 执行科室ID,C.屏蔽费别" & _
                        " From 诊疗收费关系 A,收费价目 B,收费项目目录 C,材料特性 D" & _
                        " Where A.诊疗项目ID=[1]" & _
                        " And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And A.收费项目ID=D.材料ID(+)" & _
                        " And C.服务对象 IN(2,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " Order by 从项,A.收费项目ID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!诊疗项目ID), Val(Nvl(rsSend!执行科室ID, 0)))
                End If
                
                '确定计价之中是否包含从项以及主项收入ID
                If Not rsTmp.EOF And gbln从项汇总折扣 Then
                    Do While Not rsTmp.EOF
                        If Nvl(rsTmp!从项, 0) = 0 Then
                            'SQL中主项排在前面,只取主项目的第一个收入
                            If lng主收入ID = 0 Then lng主收入ID = rsTmp!收入项目ID
                        ElseIf Nvl(rsTmp!从项, 0) = 1 Then
                            blnHaveSub = True: Exit Do
                        End If
                        rsTmp.MoveNext
                    Loop
                    rsTmp.MoveFirst
                End If
                
                Do While True
                    blnDo = False
                    If rsTmp.EOF Then
                        If lng项目ID <> 0 Then blnDo = True
                    Else
                        If rsTmp!收费项目ID <> lng项目ID And lng项目ID <> 0 Then blnDo = True
                    End If
                    If blnDo Then
                        If Not IsNull(mrsPrice!单价) Then
                            mrsPrice!单价 = Format(mrsPrice!单价, "0.00000")
                        End If
                        mrsPrice.Update
                        
                        '医嘱发送金额
                        cur金额 = cur金额 + Format(cur实收, gstrDec)
                    End If
                    If rsTmp.EOF Then Exit Do
                    
                    '------------------------------------
                    If rsTmp!收费项目ID <> lng项目ID Then
                        cur实收 = 0
                        mrsPrice.AddNew
                        mrsPrice!医嘱ID = rsSend!ID
                        mrsPrice!相关ID = rsSend!相关ID
                        mrsPrice!收费类别 = rsTmp!类别
                        mrsPrice!收费细目ID = rsTmp!收费项目ID
                        mrsPrice!数量 = Nvl(rsTmp!收费数量, 0)
                        mrsPrice!在用 = Nvl(rsTmp!跟踪在用, 0)
                        mrsPrice!固定 = Nvl(rsTmp!固有对照, 0)
                        mrsPrice!从项 = Nvl(rsTmp!从项, 0)
                        
                        '执行科室:非药嘱药品及跟踪卫材的专门取
                        lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                        If rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0 Then
                            lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!收费项目ID, 4, Nvl(rsSend!病人科室ID, 0), 0, 2, lng执行科室ID)
                        End If
                        If lng执行科室ID <> 0 Then
                            mrsPrice!执行科室ID = lng执行科室ID
                        Else
                            mrsPrice!执行科室ID = Null
                        End If
                    End If
                    lng项目ID = rsTmp!收费项目ID
                    
                    '计算单价和实收
                    If Nvl(rsTmp!是否变价, 0) = 0 Then '固定价格
                        mrsPrice!单价 = Nvl(mrsPrice!单价, 0) + rsTmp!单价
                        
                        cur应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(rsTmp!单价, "0.00000")
                        
                        '处理加班加价
                        If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                            cur应收 = cur应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
                        End If
                        
                        cur应收 = Format(cur应收, gstrDec)
                        
                        If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And Nvl(rsTmp!屏蔽费别, 0) = 0 Then
                            cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                        Else
                            cur实收 = cur实收 + cur应收
                        End If
                    ElseIf InStr(",5,6,7,", rsTmp!类别) > 0 Then
                        '非药嘱药品计价按时价计算(仅一个收入),其它变价需要由医生输入
                        mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, Nvl(mrsPrice!执行科室ID, 0), dbl数量 * Nvl(rsTmp!收费数量, 0), , True)
                        
                        cur应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, "0.00000")
                        
                        '处理加班加价
                        If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                            cur应收 = cur应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
                        End If

                        cur应收 = Format(cur应收, gstrDec)
                        
                        If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And Nvl(rsTmp!屏蔽费别, 0) = 0 Then
                            cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                        Else
                            cur实收 = cur实收 + cur应收
                        End If
                    ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Then
                        '跟踪在用的时价卫材和药品一样计算
                        mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, Nvl(mrsPrice!执行科室ID, 0), dbl数量 * Nvl(rsTmp!收费数量, 0), , True)
                        
                        cur应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, "0.00000")
                        
                        '处理加班加价
                        If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                            cur应收 = cur应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
                        End If

                        cur应收 = Format(cur应收, gstrDec)
                        
                        If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And Nvl(rsTmp!屏蔽费别, 0) = 0 Then
                            cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                        Else
                            cur实收 = cur实收 + cur应收
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                
                '从属项目汇总计算折扣
                If gbln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 Then
                    cur金额 = Format(ActualMoney(Nvl(mrsPati!费别), lng主收入ID, cur金额), gstrDec)
                End If
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'功能：根据当前医嘱行获取可选择的计价医嘱内容
'参数：lngRow=可见行(药疗或非药)
'说明：注意这里是根据具体医嘱在取,与住院不同
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If .Cell(flexcpData, lngRow, COL_ID) = 3 Then
            '中药用法：中药用法,中药煎法
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", CLng(.Cell(flexcpData, i, COL_ID))) > 0 Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        For j = 1 To mrsPrice.RecordCount
                            If Nvl(mrsPrice!固定, 0) = 0 Then
                                If .Cell(flexcpData, i, COL_ID) = 2 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药煎法-" & .Cell(flexcpData, i, COL_医嘱内容)
                                ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药用法-" & .Cell(flexcpData, i, COL_医嘱内容)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                End If
            Next
        ElseIf .Cell(flexcpData, lngRow, COL_ID) = 4 Then
            '采集方法行
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    For j = 1 To mrsPrice.RecordCount
                        If Nvl(mrsPrice!固定, 0) = 0 Then
                            If .TextMatrix(i, COL_诊疗类别) = "C" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";检验项目-" & .Cell(flexcpData, i, COL_医嘱内容)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";采集方法-" & .Cell(flexcpData, i, COL_医嘱内容)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                        mrsPrice.MoveNext
                    Next
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '首行成药：给药途径
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    For j = 1 To mrsPrice.RecordCount
                        If Nvl(mrsPrice!固定, 0) = 0 Then
                            strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";给药途径-" & .Cell(flexcpData, lngTmp, COL_医嘱内容)
                            Exit For
                        End If
                        mrsPrice.MoveNext
                    Next
                End If
            End If
        Else
            '一组手术或检查或独立医嘱
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        For j = 1 To mrsPrice.RecordCount
                            If Nvl(mrsPrice!固定, 0) = 0 Then
                                If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";附加手术-" & .Cell(flexcpData, i, COL_医嘱内容)
                                ElseIf .TextMatrix(i, COL_诊疗类别) = "G" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";手术麻醉-" & .Cell(flexcpData, i, COL_医嘱内容)
                                ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";检查部位-" & .Cell(flexcpData, i, COL_医嘱内容)
                                Else
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, COL_医嘱内容)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'功能：根据医嘱计价关系，计算并显示指定医嘱的费用(整个医嘱，可能多行)
'参数：lngRow=可见行(药疗或非药)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str计价医嘱 As String
    Dim str单位 As String, dbl数量 As Double
    Dim bln附加手术 As Boolean, strCombo As String, str行号 As String
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim dbl当前单价 As Double, cur当前应收 As Currency, cur当前实收 As Currency
    Dim lng行号 As Long, cur合计 As Currency
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
        
    On Error GoTo errH
    
    '用于汇总计算折扣的临时记录集
    rsMain.Fields.Append "医嘱行号", adBigInt
    rsMain.Fields.Append "主项行号", adBigInt
    rsMain.Fields.Append "主收入ID", adBigInt
    rsMain.Fields.Append "医嘱合计", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnFirst = False '一并给药中是否第一药品行
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or 医嘱ID=" & Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or 相关ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '计价医嘱
            bln附加手术 = False
            lng行号 = .FindRow(CStr(mrsPrice!医嘱ID), , COL_ID)
            If InStr(",5,6,7", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                str计价医嘱 = "药品医嘱-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 1 Then
                str计价医嘱 = "给药途径-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 2 Then
                str计价医嘱 = "中药煎法-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 3 Then
                str计价医嘱 = "中药用法-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 4 Then
                str计价医嘱 = "采集方法-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "C" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检验项目-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "F" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                bln附加手术 = True
                str计价医嘱 = "附加手术-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "G" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "手术麻醉-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "D" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检查部位-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            Else
                str计价医嘱 = .Cell(flexcpData, lng行号, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, lng行号, COL_医嘱内容)
            End If
            str计价医嘱 = Replace(str计价医嘱, "'", "''")
            
            '数量:药品按住院单位的数量,其它按零售数量
            If InStr(",5,6,", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                If Val(.TextMatrix(lng行号, COL_可否分零)) = 0 Then
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量)) * Val(.TextMatrix(lng行号, COL_单量)) _
                        / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装))
                Else
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量)) _
                        * IntEx(Val(.TextMatrix(lng行号, COL_单量)) / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装)))
                End If
            Else
                dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
            End If
            dbl数量 = Format(dbl数量 * Nvl(mrsPrice!数量, 0), "0.00000")
                        
            '组合SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as 序号," & mrsPrice!医嘱ID & " as 医嘱ID,ID as 收费细目ID," & _
                Nvl(mrsPrice!固定, 0) & " as 固定,'" & str计价医嘱 & "' as 计价医嘱,类别,名称,产地,规格," & _
                "计算单位 as 单位," & Nvl(mrsPrice!数量, 0) & " as 计价数量," & dbl数量 & " as 数量," & _
                Format(Nvl(mrsPrice!单价, 0), "0.00000") & " as 单价,费用类型," & lng行号 & " as 行号," & _
                " 是否变价,加班加价," & IIF(bln附加手术, 1, 0) & " as 附加手术," & mrsPrice!从项 & " as 从项," & _
                Nvl(mrsPrice!执行科室ID, 0) & " as 执行科室ID,屏蔽费别 From 收费项目目录 Where ID=" & mrsPrice!收费细目ID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '需要计价的医嘱选择
        '根据待发送医嘱取可计价医嘱(不能从mrsPrice取,因为可能无收费关系或已删除,而且也许现在计价已全部删除)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_计价医嘱) = strCombo
            .Editable = flexEDKbdMouse '可以选择则可以编辑
        Else
            .ColData(COLP_计价医嘱) = ""
        End If
        
        '显示已有的计价项目
        If strSQL <> "" Then
            strSQL = "Select A.行号,A.收费细目ID,A.固定,A.从项,A.计价医嘱,A.类别,C.名称 as 类别名称,A.执行科室ID,G.名称 as 执行科室," & _
                " Nvl(E.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 名称," & _
                " A.单位,A.计价数量,A.数量,D.住院包装,D.住院单位,Decode(A.是否变价,1,A.单价,B.现价) as 单价,F.跟踪在用," & _
                " A.费用类型,A.屏蔽费别,A.是否变价,A.加班加价,B.加班加价率,B.原价,B.现价,A.附加手术,B.附术收费率,B.收入项目ID" & _
                " From (" & strSQL & ") A,收费价目 B,收费项目类别 C,药品规格 D,收费项目别名 E,材料特性 F,部门表 G" & _
                " Where A.收费细目ID=B.收费细目ID And A.类别=C.编码 And A.收费细目ID=D.药品ID(+)" & _
                " And A.收费细目ID=F.材料ID(+) And A.执行科室ID=G.ID(+)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIF(gbln商品名, 3, 1) & _
                " Order by A.序号"
                '因为输入后是调用本函数刷新,要保持动态记录集中记录顺序
                '要保证主项排在前面,LoadAdvicePrice时，主项是排在前面，而且编辑后只可能加了从项
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) '没法
            
            If Not rsTmp.EOF And gbln从项汇总折扣 Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str行号 <> rsTmp!行号 & "_" & rsTmp!收费细目ID Then
                    If str行号 <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                            .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, "0.00000")
                            .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                            .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                        End If
                        cur合计 = cur合计 + Format(cur实收, gstrDec)
                    End If
                    str行号 = rsTmp!行号 & "_" & rsTmp!收费细目ID
                    dbl单价 = 0: cur应收 = 0: cur实收 = 0
                    .Rows = .Rows + 1
                    
                    '标识固定对照为灰色
                    If rsTmp!固定 <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_行号) = rsTmp!行号
                    .TextMatrix(.Rows - 1, COLP_收费细目ID) = rsTmp!收费细目ID
                    .TextMatrix(.Rows - 1, COLP_固定) = rsTmp!固定
                    .TextMatrix(.Rows - 1, COLP_计价医嘱) = rsTmp!计价医嘱
                    .TextMatrix(.Rows - 1, COLP_类别) = rsTmp!类别名称
                    .TextMatrix(.Rows - 1, COLP_收费类别) = rsTmp!类别
                    .TextMatrix(.Rows - 1, COLP_收费项目) = rsTmp!名称
                    .TextMatrix(.Rows - 1, COLP_计价数量) = Nvl(rsTmp!计价数量, 0) '相对数量
                    
                    dbl数量 = Nvl(rsTmp!数量, 0) '售价数量用于后面按成本打折计算
                    If InStr(",5,6,7,", rsTmp!类别) > 0 Then '住院包装
                        .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!住院单位)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0), 5)
                            dbl数量 = dbl数量 * Nvl(rsTmp!住院包装, 1)
                        Else
                            '中药药房单位按不可分零处理:每付
                            '非药嘱药品计价:因为这里预定了售价数量,因此转换为药房单位显示时不作不分零处理
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0) / Nvl(rsTmp!住院包装, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!单位)
                        .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_执行科室) = Nvl(rsTmp!执行科室)
                    .TextMatrix(.Rows - 1, COLP_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                    .TextMatrix(.Rows - 1, COLP_费用类型) = Nvl(rsTmp!费用类型)
                    .TextMatrix(.Rows - 1, COLP_从项) = IIF(Nvl(rsTmp!从项, 0) = 0, "", "√")
                    .TextMatrix(.Rows - 1, COLP_跟踪在用) = Nvl(rsTmp!跟踪在用, 0)
                    
                    '记录用于输入恢复
                    .Cell(flexcpData, .Rows - 1, COLP_计价医嘱) = .TextMatrix(.Rows - 1, COLP_计价医嘱)
                    .Cell(flexcpData, .Rows - 1, COLP_收费项目) = .TextMatrix(.Rows - 1, COLP_收费项目)
                    .Cell(flexcpData, .Rows - 1, COLP_计价数量) = .TextMatrix(.Rows - 1, COLP_计价数量)
                    .Cell(flexcpData, .Rows - 1, COLP_执行科室) = .TextMatrix(.Rows - 1, COLP_执行科室)
                    
                    '记录从属主项的信息，以便计算
                    If gbln从项汇总折扣 And rsTmp!从项 = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!行号 & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!行号 & ",") = 0 Then
                            rsClone.Filter = "行号=" & rsTmp!行号 & " And 从项=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!医嘱行号 = rsTmp!行号
                                rsMain!主项行号 = .Rows - 1
                                rsMain!主收入ID = rsTmp!收入项目ID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!行号
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!行号
                            End If
                        End If
                    End If
                    
                    '非药嘱药品及跟踪卫材:即使固定也可以修改执行科室
                    If InStr(",5,6,7,", rsTmp!类别) > 0 _
                        Or rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '单价计算处理
                If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                    If Nvl(rsTmp!是否变价, 0) = 0 Then
                        dbl当前单价 = Nvl(rsTmp!单价, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, Nvl(rsTmp!执行科室ID, 0), Format(Nvl(rsTmp!数量, 0) * Nvl(rsTmp!住院包装, 1), "0.00000"), , True)
                        Else
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, Nvl(rsTmp!执行科室ID, 0), Format(Nvl(rsTmp!数量, 0), "0.00000"), , True)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                        dbl当前单价 = Format(dbl当前单价 * Nvl(rsTmp!住院包装, 1), "0.00000")
                        cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    Else
                        cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                        dbl当前单价 = Format(dbl当前单价 * Nvl(rsTmp!住院包装, 1), "0.00000")
                    End If
                ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 And Nvl(rsTmp!是否变价, 0) = 1 Then
                    '跟踪在用的时价卫材和药品一样计算
                    dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, Nvl(rsTmp!执行科室ID, 0), Format(Nvl(rsTmp!数量, 0), "0.00000"), , True)
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                Else
                    dbl当前单价 = Format(Nvl(rsTmp!单价, 0), "0.00000") '其它如果为变价则是用户输入的
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    If Nvl(rsTmp!是否变价, 0) = 1 Then '记录非药变价范围
                        .TextMatrix(.Rows - 1, COLP_变价) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_应收金额) = CCur(Nvl(rsTmp!原价, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_实收金额) = CCur(Nvl(rsTmp!现价, 0))
                        .Editable = flexEDKbdMouse '非药品变价,即使固定也可以定价
                    End If
                End If
                '应收
                If rsTmp!附加手术 = 1 Then
                    cur当前应收 = cur当前应收 * Nvl(rsTmp!附术收费率, 100) / 100
                End If
                '处理加班加价
                If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                    cur当前应收 = cur当前应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
                End If
                cur当前应收 = Format(cur当前应收, gstrDec)
                
                '实收
                If gbln从项汇总折扣 And (rsTmp!从项 = 1 Or InStr(strHaveSub & ",", "," & rsTmp!行号 & ",") > 0) Then
                    cur当前实收 = Format(cur当前应收, gstrDec)
                    '累计医嘱合计来计算折扣
                    rsMain.Filter = "医嘱行号=" & rsTmp!行号
                    rsMain!医嘱合计 = Nvl(rsMain!医嘱合计, 0) + cur当前实收
                    rsMain.Update
                ElseIf Nvl(rsTmp!屏蔽费别, 0) = 0 And Not IsNull(mrsPati!费别) Then
                    cur当前实收 = Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur当前应收, rsTmp!收费细目ID, Nvl(rsTmp!执行科室ID, 0), _
                        dbl数量, IIF(gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                Else
                    cur当前实收 = Format(cur当前应收, gstrDec)
                End If
                
                dbl单价 = dbl单价 + dbl当前单价
                cur应收 = cur应收 + cur当前应收
                cur实收 = cur实收 + cur当前实收
                
                rsTmp.MoveNext
            Next
            If str行号 <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                    .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, "0.00000")
                    .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                    .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                End If
                cur合计 = cur合计 + Format(cur实收, gstrDec)
            End If
        End If
        
        '汇总计算折扣
        If gbln从项汇总折扣 And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur当前实收 = Format(ActualMoney(Nvl(mrsPati!费别), rsMain!主收入ID, rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 - Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                .TextMatrix(rsMain!主项行号, COLP_实收金额) = Format(Val(.TextMatrix(rsMain!主项行号, COLP_实收金额)) + (cur当前实收 - rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 + Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                rsMain.MoveNext
            Loop
        End If
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '定位缺省单元
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_计价医嘱 And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_计价医嘱
        End If
        '定位表格输入位置
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_计价医嘱 And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '重新汇总显示可见行的发送医嘱金额
    vsAdvice.TextMatrix(lngRow, COL_金额) = Format(cur合计, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln非本科 As Boolean) As Boolean
'功能：判断价表中单元格是否可以编辑
    Dim lng行号 As Long
    
    With vsPrice
        bln非本科 = False
        CellEditable = .Editable
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lngCol = COLP_执行科室 Then
            '跟踪在用的卫材,非药嘱药品计价的执行科室可以修改
            If Not (.TextMatrix(lngRow, COLP_收费类别) = "4" And Val(.TextMatrix(lngRow, COLP_跟踪在用)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_收费类别)) > 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(lng行号, COL_诊疗类别)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_收费项目) = "" Or .TextMatrix(lngRow, COLP_行号) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_固定)) <> 0 Then
            '固定对照行仅可以修改变价
            If Not (Val(.TextMatrix(lngRow, COLP_变价)) = 1 And lngCol = COLP_单价) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_单价 Then
                If Val(.TextMatrix(lngRow, COLP_变价)) <> 1 Then
                    CellEditable = False
                Else
                    '非本科执行的变价项目不允许定价格
                    If lng行号 <> 0 Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            bln非本科 = True: CellEditable = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_计价医嘱 And lngCol <> COLP_计价数量 And lngCol <> COLP_收费项目 Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Function LoadAdviceSend(Optional ByVal str类别s As String) As Boolean
'功能：根据条件读取并显示要发送的药品医嘱清单
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_ID：1-给药途径，2-中药煎法，3-中药用法，4-采集方法
'   COL_婴儿：存放婴儿编号
'   COL_诊疗类别：存放诊疗类别名称，用于显示计价医嘱
'   COL_医嘱内容：存放诊疗项目名称或标本部位，用于显示计价医嘱
'   COL_分解时间：存放费用的发生时间(无分解时间时)
'   COL_频率：1-"一次性"临嘱
'   COL_金额：原始的金额而不是累计显示的
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel组ID As Long
    Dim bln分批时价 As Boolean, lng次数 As Long, lng最小次数 As Long
    Dim str分解时间 As String, dbl总量 As Double, cur金额 As Currency
    
    Dim vMsg As VbMsgBoxResult, bln时价提示 As Boolean
    Dim bln库存提示 As Boolean, bln默认发送 As Boolean
    Dim str用法 As String, i As Long, j As Long
        
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": Call Form_Resize
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '有删除行功能
    
    vsAdvice.ColHidden(COL_婴儿) = True
    Me.Refresh
    
    Call InitPriceRecordset '计价关系表
    
    '获取发送清单:新开和已校的每条医嘱记录(药品和非药品),开嘱医生为自已
    '----------------------------------------------------------------------------------------------------------
    '叮嘱、护理等级、术后医嘱不发送,但这里先读取叮嘱(给药途径,用法,煎法,采集方法)
    strSQL = _
        " Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号,A.医嘱状态," & _
        " A.诊疗类别,F.名称 as 类别名称,A.诊疗项目ID,B.名称 as 诊疗项目,A.收费细目ID as 药品ID,A.婴儿," & _
        " A.医嘱内容,A.标本部位,A.天数,A.总给予量,D.住院单位,A.单次用量,B.计算单位,D.剂量系数,D.住院包装," & _
        " A.开始执行时间,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行时间方案," & _
        " A.病人科室ID,A.开嘱科室ID,A.开嘱医生,A.开嘱时间,A.计价特性,A.执行性质,A.执行科室ID,E.名称 as 执行科室," & _
        " B.操作类型,D.可否分零,D.药房分批,C.是否变价,C.撤档时间,C.服务对象,S.签名ID" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,收费项目目录 C,药品规格 D,部门表 E,诊疗项目类别 F,病人医嘱状态 S,病人医嘱记录 X" & _
        " Where A.病人ID=[1] And A.主页ID=[2] And Nvl(A.前提ID,0)=[3] And A.ID=S.医嘱ID And S.操作类型=1" & _
        " And A.医嘱状态 IN(1,3,5) And A.医嘱期效=1 And A.相关ID=X.ID(+) And B.类别=F.编码" & _
        " And A.诊疗项目ID=B.ID And A.收费细目ID=C.ID(+) And A.收费细目ID=D.药品ID(+)" & _
        " And A.执行科室ID=E.ID(+) And A.开始执行时间 is Not NULL And A.病人来源<>3" & _
        " And Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1)=[4]" & _
        " And Exists(Select M.姓名 From 人员表 M,执业类别 N Where M.姓名=Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1) And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))" & _
        " And Not(A.诊疗类别='H' And B.操作类型='1') And Not(A.诊疗类别='Z' And B.操作类型='4')" & _
        " Order by A.婴儿,组号,组ID,A.序号"
    
    On Error GoTo errH
    Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mlng前提ID, UserInfo.姓名)
    
    '计算并显示发送清单
    '----------------------------------------------------------------------------------------------------------
    If Not rsSend.EOF Then
        With vsAdvice
            bln时价提示 = True: bln库存提示 = True: bln默认发送 = True
            .Redraw = flexRDNone
            For i = 1 To rsSend.RecordCount
                '一组医嘱中的一个不能发送,则整组不能发送
                If lngDel组ID <> 0 Then
                    If Nvl(rsSend!相关ID, rsSend!ID) = lngDel组ID Then
                        GoTo NextLoop
                    Else
                        lngDel组ID = 0
                    End If
                End If
                
                '检查不允许发送的诊疗类别
                '一组医嘱检查第一个医嘱行,其它相关行不检查
                If str类别s <> "" And lngDel组ID = 0 Then
                    If rsSend!诊疗类别 = "7" Then
                        '中药配方
                        If InStr(str类别s, "'8'") = 0 Then lngDel组ID = rsSend!相关ID
                    ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                        '中西成药(包括西成药，中成药混合一并给药的情况)
                        If InStr(str类别s, "'" & rsSend!诊疗类别 & "'") = 0 Then
                            lngDel组ID = rsSend!相关ID
                            '删除已可能加入的其它一并给药行,当前行尚未加入不删除
                            Call DeleteCurRow(lngRow, False)
                            lng最小次数 = 0
                        End If
                    ElseIf rsSend!诊疗类别 = "D" Then
                        '检查组合(或独立的检查)
                        If InStr(str类别s, "'D'") = 0 Then lngDel组ID = rsSend!ID
                    ElseIf rsSend!诊疗类别 = "F" Then
                        '手术组合(或独立的手术)
                        If InStr(str类别s, "'F'") = 0 Then lngDel组ID = rsSend!ID
                    ElseIf rsSend!诊疗类别 = "C" Then
                        '检验组合(或独立的检验)
                        If InStr(str类别s, "'C'") = 0 Then lngDel组ID = Nvl(rsSend!相关ID, rsSend!ID)
                    ElseIf IsNull(rsSend!相关ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_相关ID)) Then
                        '其它独立项目
                        If InStr(str类别s, "'" & rsSend!诊疗类别 & "'") = 0 Then lngDel组ID = rsSend!ID
                    End If
                    If lngDel组ID <> 0 Then GoTo NextLoop
                End If
                                                
                '加入当前行
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
                Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("T").Picture
                
                '隐藏相关行
                If rsSend!诊疗类别 = "7" Then
                    .RowHidden(lngRow) = True '中草药
                ElseIf rsSend!诊疗类别 = "E" Then
                    If Not IsNull(rsSend!相关ID) Then
                        .RowHidden(lngRow) = True
                        .Cell(flexcpData, lngRow, COL_ID) = 2 '中药煎法
                    ElseIf Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                        If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_ID) = 1 '给药途径
                        ElseIf .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 4 '采集方法
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 3 '中药用法
                        End If
                    End If
                ElseIf InStr(",5,6,", rsSend!诊疗类别) = 0 And Not IsNull(rsSend!相关ID) Then
                    '附加手术,手术麻醉,检查部位,一并采集的检验项目
                    .RowHidden(lngRow) = True
                End If
                
                '排开一般的叮嘱(不含给药途径,中药煎法,用法,采集方法)
                If Nvl(rsSend!执行性质, 0) = 0 Then
                    If InStr(",1,2,3,4,", CLng(.Cell(flexcpData, lngRow, COL_ID))) = 0 _
                        And InStr(",5,6,7,", rsSend!诊疗类别) = 0 Then
                        Call .RemoveItem(lngRow): GoTo NextLoop
                    End If
                End If
                
                '一般列赋值
                '---------------------------------------------------------------
                .Cell(flexcpData, lngRow, COL_婴儿) = CLng(Nvl(rsSend!婴儿, 0))
                If Nvl(rsSend!婴儿, 0) = 0 Then
                    .TextMatrix(lngRow, COL_婴儿) = "病人"
                Else
                    .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsSend!婴儿
                    .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
                End If
                
                .TextMatrix(lngRow, COL_ID) = rsSend!ID
                .TextMatrix(lngRow, COL_相关ID) = Nvl(rsSend!相关ID)
                .TextMatrix(lngRow, COL_医嘱状态) = rsSend!医嘱状态
                .TextMatrix(lngRow, COL_诊疗类别) = rsSend!诊疗类别
                .TextMatrix(lngRow, COL_诊疗项目ID) = rsSend!诊疗项目ID
                .TextMatrix(lngRow, COL_医嘱内容) = Nvl(rsSend!医嘱内容)
                
                '电子签名标识
                .TextMatrix(lngRow, COL_签名ID) = Nvl(rsSend!签名ID)
                If Val(.TextMatrix(lngRow, COL_签名ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_医嘱内容) = img16.ListImages("签名").Picture
                End If
                
                '用于显示计价医嘱
                .Cell(flexcpData, lngRow, COL_诊疗类别) = CStr(Nvl(rsSend!类别名称))
                If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "D" Then
                    .Cell(flexcpData, lngRow, COL_医嘱内容) = CStr(Nvl(rsSend!标本部位))
                Else
                    .Cell(flexcpData, lngRow, COL_医嘱内容) = CStr(Nvl(rsSend!诊疗项目))
                End If
                
                .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsSend!医生嘱托)
                .TextMatrix(lngRow, COL_执行时间) = Nvl(rsSend!执行时间方案)
                .TextMatrix(lngRow, COL_频率) = Nvl(rsSend!执行频次)
                
                .TextMatrix(lngRow, COL_病人科室ID) = Nvl(rsSend!病人科室ID)
                .TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsSend!开嘱科室ID)
                .TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsSend!开嘱医生)
                .TextMatrix(lngRow, COL_开嘱时间) = Format(Nvl(rsSend!开嘱时间), "yyyy-MM-dd HH:mm:ss")
                                
                '采集方法行显示检验项目的执行科室
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngRow - 1, COL_执行科室)
                Else
                    .TextMatrix(lngRow, COL_执行科室) = Nvl(rsSend!执行科室)
                End If
                .TextMatrix(lngRow, COL_执行科室ID) = Nvl(rsSend!执行科室ID)
                
                .TextMatrix(lngRow, COL_计价特性) = Nvl(rsSend!计价特性, 0)
                .TextMatrix(lngRow, COL_执行性质ID) = Nvl(rsSend!执行性质, 0)
                .TextMatrix(lngRow, COL_操作类型) = Nvl(rsSend!操作类型)
                                
                '药品相关信息
                If InStr(",5,6,7", rsSend!诊疗类别) > 0 Then
                    '药品对应的规格已撤档则不允许发送(诊疗项目本身也可以相同处理,目前暂未处理)
                    If Format(Nvl(rsSend!撤档时间, "3000-01-01"), "yyyy-MM-dd") <> "3000-01-01" Or InStr(",2,3,", Nvl(rsSend!服务对象, 0)) = 0 Then
                        If rsSend!诊疗类别 = "7" Then
                            strTmp = "该中草药对应的中药配方无法发送：" & vbCrLf & vbCrLf & "　　" & Nvl(rsSend!医嘱内容)
                        Else
                            strTmp = "该药品(及一并给药的其他药品)无法发送：" & vbCrLf & vbCrLf & "　　" & Nvl(rsSend!医嘱内容)
                        End If
                        strTmp = strTmp & vbCrLf & vbCrLf & "没有发现有效的药品规格信息，该药品可能已经被停用或不能用于住院病人。"
                        strTmp = strTmp & vbCrLf & "请先到药品目录管理中处理，按[确定]继续处理其他医嘱。"
                        
                        .Redraw = flexRDDirect
                        Call .ShowCell(lngRow, COL_选择)
                        Screen.MousePointer = 0
                        MsgBox strTmp, vbInformation, gstrSysName
                        
                        '删除当前行(及相关行),及处理下一医嘱
                        Screen.MousePointer = 11
                        lngDel组ID = Nvl(rsSend!相关ID, rsSend!ID)
                        Call DeleteCurRow(lngRow)
                        .Refresh: .Redraw = flexRDNone
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                
                    .TextMatrix(lngRow, COL_药品ID) = rsSend!药品ID
                    .TextMatrix(lngRow, COL_剂量系数) = Nvl(rsSend!剂量系数, 1)
                    .TextMatrix(lngRow, COL_住院包装) = Nvl(rsSend!住院包装, 1)
                    .TextMatrix(lngRow, COL_住院单位) = Nvl(rsSend!住院单位)
                    .TextMatrix(lngRow, COL_可否分零) = Nvl(rsSend!可否分零, 0)
                    .TextMatrix(lngRow, COL_库存) = GetStock(rsSend!药品ID, Nvl(rsSend!执行科室ID, 0), 2) '按住院包装
                End If
                                                                        
                '计算发送次数，执行的分解时间等
                '---------------------------------------------------------------
                If rsSend!诊疗类别 = "7" Then
                    .TextMatrix(lngRow, COL_次数) = rsSend!总给予量
                    If Not IsNull(rsSend!执行时间方案) Then
                        .TextMatrix(lngRow, COL_分解时间) = Calc次数分解时间(rsSend!总给予量, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(0), "MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(rsSend!总给予量 - 1), "MM-dd HH:mm")
                    Else
                        '无分解时间(临嘱可能未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = Nvl(rsSend!单次用量) '单量
                    .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = rsSend!总给予量 '付数
                    .TextMatrix(lngRow, COL_总量单位) = "付"
                ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                    '计算临嘱用药次数
                    If Nvl(rsSend!天数, 0) <> 0 And Not IsNull(rsSend!执行频次) Then
                        '一个频率周期的次数
                        If rsSend!间隔单位 = "周" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / 7))
                        ElseIf rsSend!间隔单位 = "天" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔))
                        ElseIf rsSend!间隔单位 = "小时" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * 24)
                        End If
                    Else
                        '可分零药品时,按总量对单量的倍数计算给药途径的次数,否则按一个频率周期的次数计算
                        If Nvl(rsSend!可否分零, 0) = 0 And Nvl(rsSend!单次用量, 0) <> 0 Then
                            lng次数 = IntEx(rsSend!总给予量 * rsSend!剂量系数 / rsSend!单次用量)
                        Else
                            lng次数 = Nvl(rsSend!频率次数, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!执行时间方案) Then
                        str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        If str分解时间 <> "" Then
                            .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                        End If
                    Else
                        '无分解时间(临嘱可能未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                    .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = FormatEx(rsSend!总给予量 / rsSend!住院包装, 5) '以住院单位显示
                    .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!住院单位)
                    
                    If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                ElseIf rsSend!诊疗类别 = "E" And CLng(.Cell(flexcpData, lngRow, COL_ID)) <> 0 Then
                    '给药途径,中药煎法,中药用法,采集方法
                    '一并给药的按最小次数发送(影响给药途径计费)
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '给药途径
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                    .TextMatrix(j, COL_次数) = lng最小次数
                                    If .TextMatrix(j, COL_分解时间) <> "" Then
                                        .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                        .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "MM-dd HH:mm")
                                        .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng最小次数 = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数) '付数或次数
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    If .Cell(flexcpData, lngRow, COL_ID) = 3 Then '中药用法
                        .TextMatrix(lngRow, COL_总量单位) = "付"
                    Else
                        .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                    End If
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                Else
                    '其它非药临嘱:采集方法在上面的分支中已作处理
                    If IsNull(rsSend!相关ID) Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱,包括检验组合
                        dbl总量 = Nvl(rsSend!总给予量, 1)
                        lng次数 = IntEx(dbl总量 / Nvl(rsSend!单次用量, 1))
                        
                        If IsNull(rsSend!执行时间方案) And (Nvl(rsSend!频率次数, 0) = 0 Or Nvl(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                            '执行频率为"一次性"的项目
                            str分解时间 = "" '不需要
                            .Cell(flexcpData, lngRow, COL_频率) = 1
                        Else
                            '执行频率为"可选频率"的项目:下医嘱时应输入了总量
                            If Not IsNull(rsSend!执行时间方案) Then
                                str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                            Else
                                str分解时间 = "" '临嘱也许未输入执行时间,无法分解
                            End If
                        End If
                        .TextMatrix(lngRow, COL_次数) = lng次数
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                        If str分解时间 <> "" Then
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                        Else
                            '记录费用发生时间(当无分解时间时),以医嘱的开始执行时间
                            .Cell(flexcpData, lngRow, COL_分解时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
                        End If
                        
                        .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                        If Not IsNull(rsSend!单次用量) Then
                            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                        End If
                        .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                        .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                    Else
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                        .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                        .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                        .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                        .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                    End If
                End If
                
                '计算项目发送金额
                cur金额 = 0
                If Not LoadAdvicePrice(lngRow, rsSend, cur金额) Then
                    lngDel组ID = Nvl(rsSend!相关ID, rsSend!ID)
                    Call DeleteCurRow(lngRow)
                    lng最小次数 = 0: GoTo NextLoop
                End If
                .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
                .Cell(flexcpData, lngRow, COL_金额) = CCur(.TextMatrix(lngRow, COL_金额))
                
                '相关行时的一些处理：累计显示金额,给药途径,用法,执行科室,执行性质
                '---------------------------------------------------------------
                If rsSend!诊疗类别 = "E" And InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_ID))) > 0 Then '给药途径或中药用法
                    cur金额 = 0
                    lngTmp = .FindRow(CStr(rsSend!ID), , COL_相关ID)
                    
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '给药途径
                        '一并给药时,给药途径的金额累加显示在第一个成药中
                        .TextMatrix(lngTmp, COL_金额) = Format(Val(.TextMatrix(lngTmp, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                        '显示给药途径,执行性质
                        For j = lngTmp To lngRow - 1
                            strTmp = ""
                            If Val(.TextMatrix(j, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                                strTmp = "自备药"
                            ElseIf Val(.TextMatrix(j, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                                strTmp = "离院带药"
                            End If
                            .TextMatrix(j, COL_执行性质) = strTmp
                            .TextMatrix(j, COL_用法) = rsSend!诊疗项目
                        Next
                    Else
                        '药品的执行性质
                        strTmp = ""
                        If Val(.TextMatrix(lngTmp, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                            strTmp = "自备药"
                        ElseIf Val(.TextMatrix(lngTmp, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                            strTmp = "离院带药"
                        End If
                    
                        '中药用法,煎法
                        str用法 = rsSend!诊疗项目
                        If Val(.Cell(flexcpData, lngRow - 1, COL_ID)) = 2 Then
                            str用法 = str用法 & "|" & GetItemField("诊疗项目目录", Val(.TextMatrix(lngRow - 1, COL_诊疗项目ID)), "名称")
                        End If
                        For j = lngTmp To lngRow
                            .TextMatrix(j, COL_用法) = str用法 '用于填写收发记录
                            cur金额 = cur金额 + Val(.TextMatrix(j, COL_金额))
                        Next
                        .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
                        '显示执行性质
                        .TextMatrix(lngRow, COL_执行性质) = strTmp
                        '显示配方执行科室
                        .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngTmp, COL_执行科室)
                    End If
                    
                    '使相关医嘱选择状态相同(固为库存的原因；非药医嘱不用)
                    For j = lngTmp To lngRow
                        If .Cell(flexcpData, j, COL_选择) <> 0 Then
                            Call RowSelectSame(j, COL_选择)
                            Exit For '一个禁止,全部禁止
                        End If
                    Next
                    If j > lngRow Then
                        For j = lngRow To lngTmp Step -1
                            If InStr(",5,6,7,", .TextMatrix(j, COL_诊疗类别)) > 0 Then
                                If .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                                    Call RowSelectSame(j, COL_选择)
                                    Exit For '最后不选,全部不选
                                End If
                            End If
                        Next
                    End If
                ElseIf InStr(",5,6,7,", rsSend!诊疗类别) = 0 Then
                    If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 <> "C" Then
                        '其它非药医嘱
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!相关ID Then
                                .TextMatrix(j, COL_金额) = Format(Val(.TextMatrix(j, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                                Exit For
                            End If
                        Next
                    ElseIf Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                        '检验标本采集方法为显示行
                        .TextMatrix(lngRow, COL_用法) = rsSend!诊疗项目
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_金额) = Format(Val(.TextMatrix(lngRow, COL_金额)) + Val(.TextMatrix(j, COL_金额)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If

                '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止),自备药不检查
                '---------------------------------------------------------------
                If InStr(",5,6,7,", rsSend!诊疗类别) > 0 And Nvl(rsSend!执行性质, 0) <> 5 Then
                    Call CheckStock(lngRow, rsSend, bln库存提示, bln时价提示, bln默认发送)
                End If
                
NextLoop:       '---------------------------------------------------------------
                Progress = i / rsSend.RecordCount * 100
                txtPer.Text = CInt(psb.Value) & "%"
                txtPer.Refresh
                rsSend.MoveNext
            Next
        End With
    End If
    With vsAdvice
        .AutoSize COL_医嘱内容
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        '电子签名图标对齐
        .Cell(flexcpPictureAlignment, .FixedRows, COL_医嘱内容, .Rows - 1, COL_医嘱内容) = 0
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone: Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub CheckStock(ByVal lngRow As Long, rsSend As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean)
'功能：根据库存检查参数检查发送药品的库存
'参数：lngRow=医嘱行号,rsSend=当前发送医嘱信息
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = GetStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID)))
        bln分批 = Nvl(rsSend!药房分批, 0) = 1
        bln时价 = Nvl(rsSend!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = .TextMatrix(lngRow, COL_住院单位) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品总量:住院包装
            If .TextMatrix(lngRow, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                If Val(.TextMatrix(lngRow, COL_可否分零)) = 0 Then
                    dbl总量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量))
                    dbl总量 = dbl总量 / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装))
                Else
                    dbl总量 = IntEx(Val(.TextMatrix(lngRow, COL_单量)) / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装)))
                    dbl总量 = dbl总量 * Val(.TextMatrix(lngRow, COL_总量))
                End If
            Else
                dbl总量 = Val(.TextMatrix(lngRow, COL_总量))
            End If
            
            '当前可用库存:住院包装,减去前面相同药品要发送的库存
            For i = lngRow - 1 To .FixedRows Step -1
                blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                If blnDo Then
                    blnDo = Val(.TextMatrix(i, COL_药品ID)) = Val(.TextMatrix(lngRow, COL_药品ID)) _
                        And Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(lngRow, COL_执行科室ID))
                End If
                If blnDo Then
                    blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                End If
                If blnDo Then
                    If .TextMatrix(i, COL_诊疗类别) = "7" Then
                        '中药药房单位按不可分零处理:每付
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            dbl已发库存 = dbl已发库存 + _
                                Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                        End If
                    Else
                        dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                    End If
                End If
            Next
            dbl可用库存 = Val(.TextMatrix(lngRow, COL_库存))
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        strTmp = "药房分批或时价药品""" & .TextMatrix(lngRow, COL_医嘱内容) & """库存不足：" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                            "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    Else
                        strTmp = """" & .TextMatrix(lngRow, COL_医嘱内容) & """库存不足：" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                            "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该药品吗？"
                    End If
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                        End If
                    End If
                    
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    rsTotal As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean) As Boolean
'功能：发送过程中时，对非药嘱药品及跟踪在用的卫材计价进行库存检查(累计检查)
'参数：lngRow=医嘱行号
'      dbl数量=已计算好的计价数量(售价单位)
'      rsTotal=当前病人前面已累计发送的计价药品或卫材数量(售价单位)
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = GetStockCheck(lng库房ID)
        bln分批 = Nvl(rsPrice!分批, 0) = 1
        bln时价 = Nvl(rsPrice!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = Nvl(rsPrice!住院单位, Nvl(rsPrice!计算单位)) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品或卫材总量:住院包装
            dbl总量 = Format(dbl数量 / Nvl(rsPrice!住院包装, 1), "0.00000")
            
            '当前可用库存:住院包装,减去前面相同药品医嘱要发送的库存
            If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_药品ID)) = rsPrice!ID And Val(.TextMatrix(i, COL_执行科室ID)) = lng库房ID
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_诊疗类别) = "7" Then
                            '中药药房单位按不可分零处理:每付
                            If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                dbl已发库存 = dbl已发库存 + _
                                    Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                    / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                            Else
                                dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                    * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                            End If
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                        End If
                    End If
                Next
            End If
            '计价部份要发送的累计数量
            rsTotal.Filter = "项目ID=" & rsPrice!ID & " And 库房ID=" & lng库房ID
            Do While Not rsTotal.EOF
                dbl已发库存 = dbl已发库存 + Format(rsTotal!数量 / Nvl(rsPrice!住院包装, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl可用库存 = Format(GetStock(rsPrice!ID, lng库房ID, 2), "0.00000")
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        strTmp = "医嘱""" & .TextMatrix(lngRow, COL_医嘱内容) & """的分批或时价计价项目""" & rsPrice!名称 & """库存不足：" & _
                            vbCrLf & vbCrLf & Get部门名称(lng库房ID) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    Else
                        strTmp = "医嘱""" & .TextMatrix(lngRow, COL_医嘱内容) & """的计价项目""" & rsPrice!名称 & """库存不足：" & _
                            vbCrLf & vbCrLf & Get部门名称(lng库房ID) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该医嘱吗？"
                    End If
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '如果未提示或要发送,加入累计发送数量
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!项目ID = rsPrice!ID
            rsTotal!库房ID = lng库房ID
            rsTotal!数量 = dbl数量
            rsTotal.Update
        End If
    End With
End Function

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String, blnCancel As Boolean
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim vPoint As POINTAPI
    
    With vsPrice
        lng行号 = Val(.TextMatrix(Row, COLP_行号))
        If Col = COLP_收费项目 Then
            '不能选择已有的项目
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_行号)) = lng行号 And lng行号 <> 0 And i <> Row Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                End If
            Next
            str项目IDs = Mid(str项目IDs, 2)
            
            strSQL = _
                " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
                " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7)"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,-ID as ID,Nvl(-上级ID,To_Number('999999999'||类型)) as 上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 末级,ID,上级ID,编码,名称,单位,规格,产地,类别,费用类型,说明," & _
                " Decode(Nvl(是否变价,0),1,Decode(Instr('567',类别ID),0,Sum(原价)||'-'||Sum(现价),'时价'),Sum(现价)) as 价格," & _
                " Sum(原价) as 原价ID,Sum(现价) as 现价ID,是否变价 as 是否变价ID,类别ID,跟踪在用ID" & _
                " From (" & _
                " Select Distinct 1 as 末级,A.ID,Decode(Instr('567',A.类别),0,A.分类ID,-E.分类ID) as 上级ID,A.编码,A.名称," & _
                " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,A.说明,B.原价,B.现价,A.是否变价," & _
                " A.类别 as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费项目目录 A,收费价目 B,收费项目类别 C,药品规格 D,诊疗项目目录 E" & _
                " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.类别 Not IN('4','J','1') And A.类别=C.编码 And A.ID=D.药品ID(+) And D.药名ID=E.ID(+)"
            If DeptExist("发料部门", 2) Then
                strSQL = strSQL & " Union ALL " & _
                    " Select Distinct 1 as 末级,A.ID,-E.分类ID as 上级ID,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,A.说明," & _
                    " B.原价,B.现价,A.是否变价,A.类别 as 类别ID,D.跟踪在用 as 跟踪在用ID" & _
                    " From 收费项目目录 A,收费价目 B,收费项目类别 C,材料特性 D,诊疗项目目录 E" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.类别='4' And A.类别=C.编码 And A.ID=D.材料ID And D.诊疗ID=E.ID"
            End If
            strSQL = strSQL & " ) Group by 末级,ID,上级ID,类别,编码,名称,单位,规格,产地,费用类型,说明,是否变价,类别ID,跟踪在用ID"
            
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "收费项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str项目IDs & ",")
            If Not rsTmp Is Nothing Then
                '非本科执行的医嘱不允许输入变价项目
                If lng行号 <> 0 Then
                    If Nvl(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And Nvl(rsTmp!跟踪在用ID, 0) = 1) Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                '医保对码检查
                If CheckItemInsure(rsTmp) Then
                    .SetFocus: Exit Sub
                End If
                
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                Call SetItemInput(Row, rsTmp, lng医嘱ID, lng原项目ID)
                If lng行号 <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有可用的收费项目，请先到收费项目管理中设置！", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_执行科室 Then
            vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_收费类别) = "4" Then
                '跟踪在用的卫材
                strSQL = _
                    " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                    " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                    " And A.收费细目ID=[1]" & _
                    " Order by B.服务对象,C.编码"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                '药品
                '药品从系统指定的储备药房中找
                If Not Check上班安排(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                    Decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!名称
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!执行科室ID = rsTmp!ID
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset) As Boolean
'功能：检查输入(选择)计价项目是否医保对码
'返回：如果未对码，并且提示选择不继续，则返回真。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int险类 As Integer
    
    If gint医保对码 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select 险类 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then int险类 = Nvl(rsTmp!险类, 0)
    If int险类 <> 0 Then
        If Not ItemExistInsure(rsInput!ID, int险类) Then
            If gint医保对码 = 1 Then
                If MsgBox("项目""" & rsInput!名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gint医保对码 = 2 Then
                MsgBox "项目""" & rsInput!名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_计价医嘱 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_固定)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_行号)) <> 0 And Val(.TextMatrix(.Row, COLP_收费细目ID)) <> 0 Then
                    '医嘱如果有从项至少要保留一个(主项是固定不可动的)
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_计价医嘱) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("确实要删除当前计价行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & " And 收费细目ID=" & Val(.TextMatrix(.Row, COLP_收费细目ID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_计价医嘱
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_收费项目 Or .Col = COLP_执行科室) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：定位到价表中下一个可以输入的单元格
    Dim i As Long, j As Long
    
    With vsPrice
        '当前单元格如果未输入完整,则退出
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_单价 And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '从下一单元开始循环搜索
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_计价医嘱) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '当前表格内没有找到下一个可编辑单元,如果有需计价医嘱,则增加一新行
            If CStr(.ColData(COLP_计价医嘱)) <> "" Then
                '当前行未输入完整,则定位到不完整单元
                If .TextMatrix(lngRow, COLP_计价医嘱) = "" Then
                    .Col = COLP_计价医嘱
                ElseIf .TextMatrix(lngRow, COLP_计价数量) = "" Then
                    .Col = COLP_计价数量
                ElseIf .TextMatrix(lngRow, COLP_收费项目) = "" Then
                    .Col = COLP_收费项目
                ElseIf Val(.TextMatrix(lngRow, COLP_变价)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_单价)) = 0 _
                    And CellEditable(lngRow, COLP_单价) Then
                    .Col = COLP_单价
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_计价医嘱
                    
                    '缺省选择计价医嘱(如果可能)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '不可编辑时随意定一个
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim StrInput As String, strMatch As String
    Dim vPoint As POINTAPI
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng行号 = Val(.TextMatrix(Row, COLP_行号))
            If Col = COLP_计价医嘱 Then
                '下拉时回车
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '不然EnterNextCell函数要退出
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_计价数量 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "计价数量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_单价 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费单价输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '检查变价输入范围
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, "0.00000")
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_收费项目 And .EditText <> "" Then
                '不能选择已有的项目
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_行号)), COL_ID)) = Val(vsAdvice.TextMatrix(lng行号, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng行号, COL_ID)) <> 0 And i <> Row Then
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                    End If
                Next
                str项目IDs = Mid(str项目IDs, 2)
                
                '不同的输入匹配方式
                StrInput = UCase(.EditText)
                strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
                If IsNumeric(StrInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
                ElseIf zlCommFun.IsCharAlpha(StrInput) Then         '01,11.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
                ElseIf zlCommFun.IsCharChinese(StrInput) Then
                    strMatch = " And C.名称 Like [2] And C.码类=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("发料部门", 2) Then strSQL = " And A.类别<>'4'"
                strSQL = _
                    " Select A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地,A.费用类型,A.说明," & _
                    " Decode(Nvl(A.是否变价,0),1,Decode(Instr('567',A.类别ID),0,Sum(A.原价)||'-'||Sum(A.现价),'时价'),Sum(A.现价)) as 价格," & _
                    " Sum(A.原价) as 原价ID,Sum(A.现价) as 现价ID,A.是否变价 as 是否变价ID,A.类别ID,B.跟踪在用 as 跟踪在用ID" & _
                    " From (" & _
                    " Select Distinct 1 as 末级,A.ID,A.类别 as 类别ID,D.名称 as 类别,A.编码,A.名称,A.计算单位 as 单位," & _
                    " A.规格,A.产地,A.费用类型,A.说明,B.原价,B.现价,A.是否变价" & _
                    " From 收费项目目录 A,收费价目 B,收费项目别名 C,收费项目类别 D" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.收费细目ID And A.类别=D.编码 And A.类别 Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,材料特性 B" & _
                    " Where A.ID=B.材料ID(+)" & _
                    " Group by A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价,A.类别ID,B.跟踪在用" & _
                    " Order by A.类别,A.编码"
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "收费项目", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    StrInput & "%", mstrLike & StrInput & "%", mint简码 + 1, "," & str项目IDs & ",")
                If Not rsTmp Is Nothing Then
                    '非本科执行的医嘱不允许输入变价项目
                    If lng行号 <> 0 Then
                        If Nvl(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And Nvl(rsTmp!跟踪在用ID, 0) = 1) Then
                            If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                                MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    '医保对码检查
                    If CheckItemInsure(rsTmp) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                        .SetFocus: Exit Sub
                    End If
                    
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    Call SetItemInput(Row, rsTmp, lng医嘱ID, lng原项目ID)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    If lng行号 <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的收费项目！", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            ElseIf Col = COLP_执行科室 And .EditText <> "" Then '执行科室
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_收费类别) = "4" Then
                    '跟踪在用的卫材
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                        " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And A.收费细目ID=[1] And (C.编码 Like [3] Or C.名称 Like [4] Or C.简码 Like [4])" & _
                        " Order by B.服务对象,C.编码"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                    '药品从系统指定的储备药房中找
                    If Not Check上班安排(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(2,3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                            " And (A.病人来源 is NULL Or A.病人来源=2)" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                        Decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!名称
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    
                    '更新记录集
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                        mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                        mrsPrice!执行科室ID = rsTmp!ID
                        mrsPrice.Update
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_计价数量 Or Col = COLP_单价 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lng医嘱ID As Long, ByVal lng原项目ID As Long)
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng行号 As Long, dbl单价 As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '记录集内容
        '表格内容:仅临时显示标记输入了项目,也可以处理为未定计价医嘱不允许输入项目
        .TextMatrix(lngRow, COLP_类别) = rsInput!类别
        .TextMatrix(lngRow, COLP_收费类别) = rsInput!类别ID
        .TextMatrix(lngRow, COLP_收费细目ID) = rsInput!ID
        .TextMatrix(lngRow, COLP_收费项目) = rsInput!名称
        If Not IsNull(rsInput!产地) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & "(" & rsInput!产地 & ")"
        End If
        If Not IsNull(rsInput!规格) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & " " & rsInput!规格
        End If
        .TextMatrix(lngRow, COLP_单位) = Nvl(rsInput!单位) '都按零售单位(包括非药嘱药品计价)
        .TextMatrix(lngRow, COLP_计价数量) = 1 '缺省相对计价1,药品为计1个零售单位
        
        '执行科室
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lng行号 <> 0 Then
            lng执行科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))
            '非药嘱药品和跟踪在用的卫材专门求执行科室
            If rsInput!类别ID = "4" And Nvl(rsInput!跟踪在用ID, 0) = 1 Or InStr(",5,6,7,", rsInput!类别ID) > 0 Then
                lng病人科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID))
                lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsInput!类别ID, rsInput!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
            End If
        End If
        .TextMatrix(lngRow, COLP_执行科室) = Get部门名称(lng执行科室ID)
        .TextMatrix(lngRow, COLP_执行科室ID) = lng执行科室ID
        
        '单价计算处理:药嘱的药品计价不可能在这里处理
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            If Nvl(rsInput!是否变价ID, 0) = 0 Then
                dbl单价 = Nvl(rsInput!现价ID, 0)
            ElseIf lng行号 <> 0 Then
                '按每次缺省一个零售单位,当前发送数次计算
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True)
            End If
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, "0.00000")
                        
            '时价药品不输入价格
            .TextMatrix(lngRow, COLP_变价) = 0
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        ElseIf rsInput!类别ID = "4" And Nvl(rsInput!跟踪在用ID, 0) = 1 And Nvl(rsInput!是否变价ID, 0) = 1 Then
            '跟踪在用的时价卫材和药品一样计算
            dbl单价 = 0
            If lng行号 <> 0 Then
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True)
            End If
            .TextMatrix(lngRow, COLP_变价) = 0
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, "0.00000")
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        Else
            If Nvl(rsInput!是否变价ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_变价) = 0
                .TextMatrix(lngRow, COLP_单价) = Format(Nvl(rsInput!现价ID, 0), "0.00000")
                .Cell(flexcpData, lngRow, COLP_应收金额) = 0
                .Cell(flexcpData, lngRow, COLP_实收金额) = 0
            Else
                .TextMatrix(lngRow, COLP_变价) = 1
                .TextMatrix(lngRow, COLP_单价) = ""
                .Cell(flexcpData, lngRow, COLP_应收金额) = Nvl(rsInput!原价ID, 0)
                .Cell(flexcpData, lngRow, COLP_实收金额) = Nvl(rsInput!现价ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_费用类型) = Nvl(rsInput!费用类型)
        .TextMatrix(lngRow, COLP_固定) = 0
        
        '用于输入恢复
        .Cell(flexcpData, lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目)
        .Cell(flexcpData, lngRow, COLP_计价数量) = .TextMatrix(lngRow, COLP_计价数量)
        .Cell(flexcpData, lngRow, COLP_单价) = .TextMatrix(lngRow, COLP_单价)
        .Cell(flexcpData, lngRow, COLP_执行科室) = .TextMatrix(lngRow, COLP_执行科室)
        
        '记录集内容
        If lng医嘱ID <> 0 Then
            If lng原项目ID = 0 Then
                '当前医嘱是否有从项决定新增的项目是否从项
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 从项=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_从项) = IIF(blnHaveSub, "√", "")
            
                mrsPrice.AddNew '加入
            Else '更新
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
            End If
            If lng原项目ID = 0 Then
                mrsPrice!医嘱ID = lng医嘱ID
                lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
                If Val(vsAdvice.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                    mrsPrice!相关ID = Val(vsAdvice.TextMatrix(lng行号, COL_相关ID))
                Else
                    mrsPrice!相关ID = Null
                End If
                mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!收费类别 = rsInput!类别ID
            mrsPrice!收费细目ID = rsInput!ID
            If lng执行科室ID <> 0 Then
                mrsPrice!执行科室ID = lng执行科室ID
            Else
                mrsPrice!执行科室ID = Null
            End If
            mrsPrice!在用 = Nvl(rsInput!跟踪在用ID, 0)
            mrsPrice!数量 = 1
            mrsPrice!单价 = Val(.TextMatrix(lngRow, COLP_单价))
            mrsPrice!固定 = 0
            mrsPrice.Update
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln非本科 As Boolean
    
    If Not CellEditable(Row, Col, bln非本科) Then
        '非本科执行的变价项目不允许定价格
        If bln非本科 Then
            MsgBox "该医嘱非本科执行，不允许对变价项目定价。该计价项目需要手工计价。", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_计价数量 Or Col = COLP_单价 Or Col = COLP_执行科室 Then
            '必须先确定收费项目
            If vsPrice.TextMatrix(Row, COLP_收费项目) = "" Then Cancel = True
        End If
        If Col = COLP_单价 Then
            '输入变价前必须先确定计价医嘱,以决定是否可以输入(本科执行)
            If vsPrice.TextMatrix(Row, COLP_计价医嘱) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_计价数量 Or Col = COLP_单价 Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub InitBillSet()
'功能：初始化医嘱记帐单据生成记录集
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.Fields.Append "费用序号", adBigInt
    mrsBill.Fields.Append "发送序号", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng费用序号 As Long, lng发送序号 As Long, bln到门诊 As Boolean)
'功能：获取当前费用单据的NO及序号
'参数：lng费用序号=费用记录中的序号,为-1时表示不取费用序号
'      lng发送序号=发送记录中的序号,为-1时表示不取发送序号
'说明：strKey=根据记帐单据生成规则定的唯一关键字
'1.中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
'2.一个配方中的所有草药分配一个独立单据号
'3.材料医嘱与成药分号规则相同。
'4.其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
'5.检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
'6.一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNO(IIF(bln到门诊, 13, 14))
        mrsBill!费用序号 = IIF(lng费用序号 = -1, 0, 1)
        mrsBill!发送序号 = IIF(lng发送序号 = -1, 0, 1)
        mrsBill.Update
    Else
        If lng费用序号 <> -1 Then
            mrsBill!费用序号 = mrsBill!费用序号 + 1
        End If
        If lng发送序号 <> -1 Then
            mrsBill!发送序号 = mrsBill!发送序号 + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng费用序号 <> -1 Then lng费用序号 = mrsBill!费用序号
    If lng发送序号 <> -1 Then lng发送序号 = mrsBill!发送序号
End Sub

Private Sub DeleteSendRow()
'功能：将待发送医嘱清单中已发送成功的的行删除
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_选择
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Get实收金额(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Get实收金额 = CCur(strSQL)
End Function

Private Function Set实收金额(ByVal strSQL As String, ByVal cur金额 As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Set实收金额 = strLeft & cur金额 & strRight
End Function

Private Function CheckSignSend() As Boolean
'功能：检查一起签名的医嘱是否一起发送的
'说明：这里只检查新开的医嘱，已校对的医嘱发送不再受影响(不同于门诊没有校对)
    Dim col签名ID As New Collection, str签名ID As String
    Dim lng签名ID As Long, strTmp As String
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                '收集已签名医嘱的发送状态
                lng签名ID = Val(.TextMatrix(i, COL_签名ID))
                If lng签名ID <> 0 Then
                    If InStr(str签名ID & ",", "," & lng签名ID & ",") > 0 Then
                        strTmp = Split(col签名ID("_" & lng签名ID), "=")(1)
                        If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                            If InStr(strTmp, "1") = 0 Then
                                col签名ID.Remove "_" & lng签名ID
                                col签名ID.Add lng签名ID & "=" & strTmp & "1", "_" & lng签名ID
                            End If
                        Else
                            If InStr(strTmp, "0") = 0 Then
                                col签名ID.Remove "_" & lng签名ID
                                col签名ID.Add lng签名ID & "=" & strTmp & "0", "_" & lng签名ID
                            End If
                        End If
                    Else
                        str签名ID = str签名ID & "," & lng签名ID
                        If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                            col签名ID.Add lng签名ID & "=1", "_" & lng签名ID
                        Else
                            col签名ID.Add lng签名ID & "=0", "_" & lng签名ID
                        End If
                    End If
                End If
            End If
        Next
            
        '检查签名情况(一次签名的医嘱必须一起发送)
        strTmp = ""
        For i = 1 To col签名ID.Count
            lng签名ID = Split(col签名ID(i), "=")(0)
            str签名ID = Split(col签名ID(i), "=")(1)
            If Not (str签名ID = "1" Or str签名ID = "0") Then
                '这次签名的内容不是"都要发送或都不发送"的情况
                j = .FindRow(CStr(lng签名ID), , COL_签名ID)
                Do While j <> -1
                    If Not .RowHidden(j) Then
                        If .Cell(flexcpData, j, COL_选择) = 1 Or .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                            strTmp = strTmp & vbCrLf & "●" & .TextMatrix(j, COL_医嘱内容)
                        End If
                    End If
                    j = .FindRow(CStr(lng签名ID), j + 1, COL_签名ID)
                Loop
                Exit For '暂只提示第一组
            End If
        Next
    End With
    
    If strTmp <> "" Then
        MsgBox "以下医嘱与其他本次要发送的医嘱一起签名，但当前处理为不发送：" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "一起签名的医嘱必须一起发送，请调整相关医嘱的发送状态。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSignSend = True
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng项目ID As Long, ByVal lngCol As Long)
'功能：定位到并显示指定医嘱的指定计价行
'参数：lngRow=医嘱行号
'      lng项目ID=计价项目ID
'      lngCol=计价表格显示列
    Dim k As Long
    
    With vsAdvice
        .Col = COL_医嘱内容 '进入行自动ShowPrice,mrsPrice发生变化
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                '附加手术,手术麻醉,检查部位,检验组合项目
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '给药途径
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '中药煎法
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_行号)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_收费细目ID)) = lng项目ID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function SendAdvice(ByVal bln到门诊 As Boolean) As Long
'功能：处理医嘱发送(这个过程中记帐报警)
'说明：逐个病人发送提交
'返回：如果成功则返回发送号
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset

    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date
    Dim blnTran As Boolean, blnBool As Boolean, strTmp As String
    Dim strWarn As String, intWarn As Integer, str类别 As String, str类别名称 As String
    
    Dim lng发送号 As Long, int计费状态 As Integer, int划价 As Integer, strNO As String
    Dim lng细目ID As Long, lng费用序号 As Long, lng费用父号 As Long, lng发送序号 As Long
    Dim int付数 As Integer, dbl数量 As Double, cur合计 As Currency
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim int配方数 As Integer, strNOKey As String, str自动发料 As String
    Dim str发生时间 As String, str登记时间 As String
    Dim dbl发送数次 As Double, blnFirst As Boolean '配方数及分号关键字
    Dim lng药品类别ID As Long, lng卫材类别ID As Long
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim bln离院带药 As Boolean, blnVarZero As Boolean
    Dim bln附加手术 As Boolean, blnHaveSub As Boolean
    Dim int父序号 As Integer, var父索引 As Variant
    Dim lng父收入ID As Long, str实收 As String
    Dim cur医嘱合计 As Currency
    
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    
    Dim rsAudit As ADODB.Recordset
    Dim strAudit As String
    
    '电子签名
    Dim lng组ID As Long, str医嘱IDs As String, strSource As String
    Dim intRule As Integer, strSign As String
    Dim lng证书ID As Long, lng签名ID As Long
    
    On Error GoTo errH
    
    '检查一起签名的医嘱是否一起发送
    If Not CheckSignSend Then Exit Function
    
    With vsAdvice
        '先检查并提示特殊医嘱:3-转科,5-出院,6-转院,11-死亡
        strTmp = ""
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                    strTmp = strTmp & vbCrLf & mrsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, COL_医嘱内容)
                End If
            End If
        Next
        If strTmp <> "" Then
            If MsgBox("要发送的医嘱中包含下列特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送当前选择的医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End With
    
    '提取当前病人的审批项目清单
    strAudit = ""
    If Not IsNull(mrsPati!险类) Then
        Set rsAudit = GetAuditRecord(mlng病人ID, mlng主页ID)
    Else
        Set rsAudit = Nothing '以Nothing为标志该病人不需要判断
    End If
    
    '读取药品/卫材入出类别
    lng药品类别ID = ExistIOClass(IIF(bln到门诊, 8, 9))
    lng卫材类别ID = ExistIOClass(IIF(bln到门诊, 40, 41))
    
    Screen.MousePointer = 11
    
    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
    
    Call InitBillSet
    Call InitRecordSet(rsSQL, rsTotal, rsUpload)
    lng发送号 = zlDatabase.GetNextNO(10)
    
    '这个时间发送过程中未用于停止时间,为避免与校对时间重复(取的Sysdate)
    curDate = zlDatabase.Currentdate
    intWarn = -1 '记帐报警时缺省要提示,与病人无关
    int配方数 = 1 '表示发送的第几付配方,用于分单据号
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '特殊医嘱：3-转科;5-出院;6-转院,11-死亡
                If .TextMatrix(i, COL_诊疗类别) = "Z" Then
                    '转科,出院,转院,死亡医嘱发送时，病人要处于正常状态
                    If .Cell(flexcpData, i, COL_婴儿) = 0 Then
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 And Nvl(mrsPati!状态, 0) <> 0 Then
                            MsgBox "病人""" & mrsPati!姓名 & """当前处于""" & Decode(Nvl(mrsPati!状态, 0), 1, "等待入科", 2, "正在转科", 3, "已预出院") & """状态，" & _
                                "不能发送""" & .TextMatrix(i, COL_医嘱内容) & """医嘱。", vbInformation, gstrSysName
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                            GoTo NextAdvice
                        End If
                    End If
                    
                    '如果是转科、出院、转院医嘱,检查病人是否有未执行的医技项目及未发药品
                    If InStr(",3,5,6,", .TextMatrix(i, COL_操作类型)) > 0 Then
                        strTmp = ExistWaitExe(mlng病人ID, mlng主页ID, .Cell(flexcpData, i, COL_婴儿))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_医嘱内容): .Refresh
                            If MsgBox("发现病人""" & mrsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """存在尚未执行完成的内容：" & _
                                vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(i, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                GoTo NextAdvice
                            End If
                        End If
                        strTmp = ExistWaitDrug(mlng病人ID, mlng主页ID, .Cell(flexcpData, i, COL_婴儿))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_医嘱内容): .Refresh
                            If MsgBox("发现病人""" & mrsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """" & _
                                strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(i, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                GoTo NextAdvice
                            End If
                        End If
                    End If
                End If
            
                '产生单据号分配关键字
                '-----------------------------------------------------------------------------------------
                If InStr(",5,6,M,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '中西成药和材料按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                    strNOKey = "成药材料_" & mlng病人ID & "_" & mlng主页ID & "_" & _
                        Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                        .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                    '再按要打印的诊疗单据分号
                    strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                    '一个配方中的所有草药分配一个独立单据号
                    strNOKey = "中药配方_" & mlng病人ID & "_" & mlng主页ID & "_" & int配方数
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "C" Then
                    '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                    strNOKey = "一并采集_" & Val(.TextMatrix(i, COL_相关ID))
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_相关ID))
                Else
                    '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                End If
                
                '是否离院带药
                bln离院带药 = False
                If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    If .TextMatrix(i, COL_执行性质) = "离院带药" Then bln离院带药 = True
                ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                    If j <> -1 Then
                        If .TextMatrix(j, COL_执行性质) = "离院带药" Then bln离院带药 = True
                    End If
                End If
                
                '产生医嘱记帐费用:以最新价格计算
                '-----------------------------------------------------------------------------------------
                strSQL = "": lng细目ID = 0
                If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
                    If Val(.TextMatrix(i, COL_执行性质ID)) <> 5 Then
                        strSQL = _
                            " Select A.ID,A.类别,D.名称 as 类别名称,RTrim(A.名称||' '||A.规格) as 名称," & _
                            " A.计算单位,A.是否变价,A.屏蔽费别,A.加班加价,B.加班加价率,100 as 附术收费率," & _
                            " Y.住院单位,Y.住院包装,Y.剂量系数,Y.药房分批 as 分批,0 as 跟踪在用,B.收入项目ID," & _
                            " C.收据费目,1 as 数量,B.现价 as 单价,[2] as 执行科室ID,0 as 从项,I.要求审批" & _
                            " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y,保险支付项目 I" & _
                            " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                            " And A.ID=Y.药品ID(+) And A.ID=[1] And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                            " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                            " Order by A.编码"
                    End If
                Else
                    '先删除原非药医嘱的计价(应该没有)
                    rsSQL.AddNew
                    rsSQL!类型 = 1: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    rsSQL!SQL = "ZL_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                    rsSQL.Update
                    
                    '不计价,手工计价；叮嘱,院外执行的医嘱不读取
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                blnVarZero = False '是否变价且价格为零
                                If Nvl(mrsPrice!单价, 0) = 0 Then
                                    blnVarZero = ItemIsVarPrice(mrsPrice!收费细目ID)
                                End If
                                If Not blnVarZero Then
                                    If Nvl(mrsPrice!数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                        '填写病人医嘱计价:只保存非药嘱药品及跟踪卫材计价
                                        If InStr(",5,6,7,", mrsPrice!收费类别) > 0 _
                                            Or mrsPrice!收费类别 = "4" And Nvl(mrsPrice!在用, 0) = 1 Then
                                            lng执行科室ID = Nvl(mrsPrice!执行科室ID, 0)
                                            
                                            '卫材必须设置执行科室
                                            If lng执行科室ID = 0 And mrsPrice!收费类别 = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!收费细目ID, COLP_执行科室)
                                                Screen.MousePointer = 0
                                                MsgBox "卫材""" & vsPrice.TextMatrix(vsPrice.Row, COLP_收费项目) & """没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                                    "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                                vsPrice.SetFocus: GoTo FuncEnd
                                            End If
                                        Else
                                            lng执行科室ID = 0
                                        End If
                                        rsSQL.AddNew
                                        rsSQL!类型 = 1: rsSQL!项目ID = mrsPrice!收费细目ID: rsSQL!序号 = i
                                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                        rsSQL!SQL = "ZL_病人医嘱计价_INSERT(" & _
                                            mrsPrice!医嘱ID & "," & mrsPrice!收费细目ID & "," & _
                                            Nvl(mrsPrice!数量, 0) & "," & Nvl(mrsPrice!单价, 0) & "," & _
                                            Nvl(mrsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & ")"
                                        rsSQL.Update
                                        
                                        '临时病人医嘱计价表
                                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                            "Select " & mrsPrice!收费细目ID & " as 收费细目ID," & _
                                            Nvl(mrsPrice!执行科室ID, 0) & " as 执行科室ID," & _
                                            Nvl(mrsPrice!数量, 0) & " as 数量," & _
                                            Format(Nvl(mrsPrice!单价, 0), "0.00000") & " as 单价," & _
                                            Nvl(mrsPrice!从项, 0) & " as 从项 From Dual"
                                    End If
                                Else 'If Check本科执行(Val(.TextMatrix(i, COL_执行科室ID))) Then
                                    '单价为零,表明是变价未定价(正常价格不可能为0)
                                    '本科执行的需要定价,否则是手工计价
                                    Call SeekPriceRow(i, mrsPrice!收费细目ID, COLP_单价)
                                    Screen.MousePointer = 0
                                    MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, gstrSysName
                                    vsPrice.SetFocus: GoTo FuncEnd
                                End If
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    
                    If strSQL <> "" Then
                        strSQL = _
                            " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,A.是否变价," & _
                            " A.屏蔽费别,A.加班加价,B.加班加价率,B.附术收费率,Y.住院单位,Y.住院包装,Y.剂量系数," & _
                            " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,B.收入项目ID," & _
                            " C.收据费目,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,X.执行科室ID,X.从项,I.要求审批" & _
                            " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,(" & strSQL & ") X,药品规格 Y,保险支付项目 I" & _
                            " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.ID=E.材料ID(+)" & _
                            " And A.类别=D.编码 And X.收费细目ID=A.ID And A.ID=Y.药品ID(+)" & _
                            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                            " And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                            " Order by X.从项,A.ID"
                            '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                    End If
                End If
                                
                '汇总折扣变量初始
                blnHaveSub = False
                var父索引 = Empty: int父序号 = 0
                cur医嘱合计 = 0: lng父收入ID = 0
                
                int计费状态 = IIF(Val(.TextMatrix(i, COL_计价特性)) = 1, -1, 0) '无需计费或未计费
                If strSQL <> "" Then
                    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_药品ID)), Val(.TextMatrix(i, COL_执行科室ID)), Val(Nvl(mrsPati!险类, 0)))
                    If Not rsMoney.EOF Then
                        int计费状态 = 1 '已计费
                                                
                        '确定是否主从关系:即使不汇总折扣,也要记录
                        rsMoney.Filter = "从项=1"
                        If Not rsMoney.EOF Then blnHaveSub = True
                        rsMoney.Filter = 0
                    End If
                    
                    '处理收入项目级的费用明细
                    bln附加手术 = .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0
                    For j = 1 To rsMoney.RecordCount
                        '检查是否需要和已经审批
                        If Nvl(rsMoney!要求审批, 0) = 1 And Not rsAudit Is Nothing Then
                            rsAudit.Filter = "项目ID=" & rsMoney!ID
                            If rsAudit.EOF Then
                                If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                    If InStr(strAudit, "●" & rsMoney!名称) = 0 Then
                                        strAudit = strAudit & vbCrLf & "●" & rsMoney!名称
                                    End If
                                ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                    strAudit = strAudit & vbCrLf & "… …"
                                End If
                            End If
                        End If
                    
                        '执行科室ID
                        lng执行科室ID = Nvl(rsMoney!执行科室ID, 0)
                        '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                        If rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 _
                            Or InStr(",5,6,7", rsMoney!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                            lng病人科室ID = Val(.TextMatrix(i, COL_病人科室ID))
                            lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsMoney!类别, rsMoney!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
                        End If
                        If InStr(",5,6,7", rsMoney!类别) > 0 Then
                            If lng药品类别ID = 0 Then
                                MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                        
                            If InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                    int付数 = Val(.TextMatrix(i, COL_总量))
                                    '中药药房单位按不可分零处理:每付
                                    If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                        dbl数量 = Val(.TextMatrix(i, COL_单量)) / Nvl(rsMoney!剂量系数, 1)
                                    Else
                                        dbl数量 = IntEx(Val(.TextMatrix(i, COL_单量)) / Nvl(rsMoney!剂量系数, 1) / Nvl(rsMoney!住院包装, 1)) * Nvl(rsMoney!住院包装, 1)
                                    End If
                                Else
                                    int付数 = 1
                                    dbl数量 = Val(.TextMatrix(i, COL_总量)) * Nvl(rsMoney!住院包装, 1)
                                End If
                            Else
                                int付数 = 1
                                '中药药房单位按不可分零处理:每付
                                '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                                dbl数量 = Val(.TextMatrix(i, COL_总量)) * Nvl(rsMoney!数量, 0)
                            End If
                            dbl数量 = Format(dbl数量, "0.00000")
                            
                            If Nvl(rsMoney!是否变价, 0) = 1 Then
                                dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, int付数 * dbl数量, , True), "0.00000")
                            Else
                                dbl单价 = Format(Nvl(rsMoney!单价, 0), "0.00000")
                            End If
                        ElseIf rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Then
                            '检查卫生材料入出类别
                            If lng卫材类别ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                            
                            int付数 = 1
                            dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * Nvl(rsMoney!数量, 0), "0.00000")
                            
                            '确定时价卫材价格
                            If Nvl(rsMoney!是否变价, 0) = 1 Then
                                dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, dbl数量, , True), "0.00000")
                            Else
                                dbl单价 = Format(Nvl(rsMoney!单价, 0), "0.00000")
                            End If
                        Else
                            int付数 = 1
                            dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * Nvl(rsMoney!数量, 0), "0.00000")
                            dbl单价 = Format(Nvl(rsMoney!单价, 0), "0.00000")
                        End If
                        
                        '非药嘱药品及跟踪卫材的库存检查
                        If rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 _
                            Or InStr(",5,6,7", rsMoney!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                            If GetStockCheck(lng执行科室ID) <> 0 Or Nvl(rsMoney!是否变价, 0) = 1 Or Nvl(rsMoney!分批, 0) = 1 Then
                                If rsMoney!类别 = "4" Then
                                    blnBool = CheckPriceStock(i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                                Else
                                    blnBool = CheckPriceStock(i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_选择, rsSQL, rsTotal, rsUpload, str医嘱IDs)
                                    GoTo NextAdvice
                                End If
                            End If
                        End If
                            
                        '发送金额
                        cur应收 = int付数 * dbl数量 * dbl单价
                        If bln附加手术 Then
                            cur应收 = cur应收 * Nvl(rsMoney!附术收费率, 100) / 100
                        End If
                        
                        '处理加班加价
                        If gbln加班加价 And Nvl(rsMoney!加班加价, 0) = 1 Then
                            cur应收 = cur应收 * (1 + Nvl(rsMoney!加班加价率, 0) / 100)
                        End If
                        
                        cur应收 = Format(cur应收, gstrDec)
                        
                        '计算汇总折扣合计
                        If gbln从项汇总折扣 And blnHaveSub Then
                            cur实收 = cur应收
                            cur医嘱合计 = cur医嘱合计 + cur实收
                        ElseIf Nvl(rsMoney!屏蔽费别, 0) = 0 Then
                            cur实收 = Format(ActualMoney(Nvl(mrsPati!费别), rsMoney!收入项目ID, cur应收, rsMoney!ID, lng执行科室ID, int付数 * dbl数量, _
                                IIF(gbln加班加价 And Nvl(rsMoney!加班加价, 0) = 1, Nvl(rsMoney!加班加价率, 0) / 100, 0)), gstrDec)
                        Else
                            cur实收 = cur应收
                        End If
                        
                        '医保相关字段
                        bln保险项目否 = False: lng保险大类ID = 0: cur统筹金额 = 0: str保险编码 = "": str费用类型 = ""
                        If Not IsNull(mrsPati!险类) Then
                            strTmp = gclsInsure.GetItemInsure(mlng病人ID, rsMoney!ID, cur实收, False, mrsPati!险类)
                            If strTmp <> "" Then
                                bln保险项目否 = Val(Split(strTmp, ";")(0)) <> 0
                                lng保险大类ID = Val(Split(strTmp, ";")(1))
                                cur统筹金额 = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                str保险编码 = CStr(Split(strTmp, ";")(3))
                                If UBound(Split(strTmp, ";")) >= 5 Then
                                    If Split(strTmp, ";")(5) <> "" Then
                                        str费用类型 = Split(strTmp, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                        
                        '收集记帐报警类别
                        cur合计 = cur合计 + cur实收
                        If InStr(str类别, rsMoney!类别) = 0 Then
                            str类别 = str类别 & rsMoney!类别
                            str类别名称 = str类别名称 & "," & rsMoney!类别名称
                        End If
                        
                        'NO,序号
                        Call GetCurBillSet(strNOKey, strNO, lng费用序号, -1, bln到门诊)
                        rsSQL.AddNew: blnBool = False
                        If rsMoney!ID <> lng细目ID Then
                            lng费用父号 = lng费用序号
                            '主从关系时，记录主项信息
                            If rsMoney!从项 = 0 And blnHaveSub Then
                                int父序号 = lng费用序号
                                lng父收入ID = rsMoney!收入项目ID
                                var父索引 = rsSQL.Bookmark
                                blnBool = True
                            End If
                        End If
                        lng细目ID = rsMoney!ID
                        
                        '汇总折扣时，对主项的实收金额作特殊处理
                        If gbln从项汇总折扣 And blnHaveSub And blnBool Then
                            str实收 = Chr(0) & Chr(1) & "Begin" & cur实收 & "End" & Chr(0) & Chr(1)
                        Else
                            str实收 = cur实收
                        End If
                        
                        '发生时间
                        If .TextMatrix(i, COL_分解时间) <> "" Then
                            str发生时间 = "To_Date('" & Split(.TextMatrix(i, COL_分解时间), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str发生时间 = "To_Date('" & .Cell(flexcpData, i, COL_分解时间) & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        
                        '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                        rsSQL!类型 = 5: rsSQL!项目ID = rsMoney!ID: rsSQL!序号 = i
                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                        If bln到门诊 Then
                            '暂未取发药窗口
                            rsSQL!SQL = "ZL_门诊划价记录_INSERT(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & mlng病人ID & "," & ZVal(mlng主页ID) & "," & _
                                ZVal(Nvl(mrsPati!住院号, 0)) & ",'" & Nvl(mrsPati!床号) & "','" & Nvl(mrsPati!姓名) & "'," & _
                                "'" & Nvl(mrsPati!性别) & "','" & Nvl(mrsPati!年龄) & "'," & _
                                "'" & Nvl(mrsPati!费别) & "',NULL," & ZVal(Nvl(mrsPati!当前病区ID, 0)) & "," & _
                                ZVal(.TextMatrix(i, COL_病人科室ID)) & "," & ZVal(.TextMatrix(i, COL_开嘱科室ID)) & "," & _
                                "'" & .TextMatrix(i, COL_开嘱医生) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!类别 & "','" & Nvl(rsMoney!计算单位) & "',NULL," & _
                                int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                                "'" & Nvl(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                str发生时间 & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "'医嘱发送','" & UserInfo.姓名 & "'," & IIF(rsMoney!类别 = "4", lng卫材类别ID, lng药品类别ID) & "," & _
                                "'" & .TextMatrix(i, COL_医嘱内容) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_频率) & "'," & _
                                ZVal(.TextMatrix(i, COL_单量)) & ",'" & .TextMatrix(i, COL_用法) & "',1," & _
                                IIF(bln离院带药, 3, Val(.TextMatrix(i, COL_计价特性))) & ",2)"
                        Else
                            '是否划价费用
                            If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                int划价 = IIF(InStr(gstr发送划价单, "5") > 0, 1, 0)
                            Else
                                int划价 = IIF(InStr(gstr发送划价单, .TextMatrix(i, COL_诊疗类别)) > 0, 1, 0)
                            End If
                            
                            '登记时间
                            If int划价 = 1 Then '与非划价的时间上区分开
                                str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str登记时间 = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            '收集医保上传单据号:mrsBill中的不一定产生了费用
                            If int划价 = 0 Then
                                rsUpload.Filter = "NO='" & strNO & "'"
                                If rsUpload.EOF Then
                                    rsUpload.AddNew
                                    rsUpload!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                    rsUpload!NO = strNO
                                    rsUpload.Update
                                End If
                            End If
                            
                            rsSQL!SQL = "ZL_住院记帐记录_Insert(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & mlng病人ID & "," & ZVal(mlng主页ID) & "," & _
                                ZVal(Nvl(mrsPati!住院号)) & ",'" & Nvl(mrsPati!姓名) & "'," & _
                                "'" & Nvl(mrsPati!性别) & "','" & Nvl(mrsPati!年龄) & "'," & _
                                "'" & Nvl(mrsPati!床号) & "','" & Nvl(mrsPati!费别) & "'," & _
                                ZVal(Nvl(mrsPati!当前病区ID, 0)) & "," & ZVal(.TextMatrix(i, COL_病人科室ID)) & ",0," & _
                                Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                ZVal(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & rsMoney!ID & "," & _
                                "'" & rsMoney!类别 & "','" & Nvl(rsMoney!计算单位) & "'," & _
                                IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                                int付数 & "," & dbl数量 & ",NULL," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                                "'" & Nvl(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                cur统筹金额 & "," & str发生时间 & "," & str登记时间 & "," & _
                                "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                                IIF(rsMoney!类别 = "4", lng卫材类别ID, lng药品类别ID) & "," & _
                                "NULL,'" & .TextMatrix(i, COL_医嘱内容) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_频率) & "'," & ZVal(.TextMatrix(i, COL_单量)) & "," & _
                                "'" & .TextMatrix(i, COL_用法) & "',1," & _
                                IIF(bln离院带药, 3, Val(.TextMatrix(i, COL_计价特性))) & ",Null,'" & str费用类型 & "')"
                        End If
                        rsSQL.Update
                        
                        '记录自动发料的SQL
                        If gbln住院自动发料 And Not bln到门诊 And int划价 = 0 And lng执行科室ID <> 0 And rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Then
                            If InStr(str自动发料 & ";", ";" & strNO & "," & lng执行科室ID & ";") = 0 Then
                                rsSQL.AddNew
                                rsSQL!类型 = 6
                                rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                rsSQL!项目ID = 0
                                rsSQL!序号 = i
                                rsSQL!SQL = "zl_材料收发记录_处方发料(" & lng执行科室ID & ",25,'" & strNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                                rsSQL.Update
                                str自动发料 = str自动发料 & ";" & strNO & "," & lng执行科室ID
                            End If
                        End If
                        
                        rsMoney.MoveNext
                    Next
                End If
                
                '对医嘱金额进行汇总折扣处理
                If gbln从项汇总折扣 And blnHaveSub And var父索引 <> Empty And lng父收入ID <> 0 Then
                    rsSQL.Bookmark = var父索引
                    cur实收 = Format(ActualMoney(Nvl(mrsPati!费别), lng父收入ID, cur医嘱合计), gstrDec)
                    cur实收 = cur实收 - cur医嘱合计 '打折差额
                    cur实收 = Get实收金额(rsSQL!SQL) + cur实收
                    rsSQL!SQL = Set实收金额(rsSQL!SQL, cur实收)
                    rsSQL.Update
                End If
                
                '选择要发送的医嘱自动进行校对(实际可能因为叮嘱不发送)
                If Val(.TextMatrix(i, COL_医嘱状态)) = 1 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    rsSQL.AddNew
                    rsSQL!类型 = 3: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = Val(.TextMatrix(i, COL_ID))
                    rsSQL!SQL = "ZL_病人医嘱记录_校对(" & Val(.TextMatrix(i, COL_ID)) & ",3," & _
                        "To_Date('" & Format(.TextMatrix(i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),0)"
                End If
                
                
                '产生医嘱发送记录
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then '叮嘱不发送(给药途径，配方煎法、用法,采集方法可能为)
                    '发送了出院,转院,死亡医嘱
                    If .TextMatrix(i, COL_诊疗类别) = "Z" _
                        And InStr(",5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                        mblnRefresh = True
                    End If
                    
                    '一样要产生费用NO
                    Call GetCurBillSet(strNOKey, strNO, -1, lng发送序号, bln到门诊)
                                                            
                    '是否一组医嘱的第一医嘱行:药疗的第一药品行为第一医嘱行
                    blnFirst = False
                    If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = True
                        End If
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "C" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = True '检验组合中的第一检验行
                        End If
                    ElseIf InStr(",1,2,3,4,", CLng(.Cell(flexcpData, i, COL_ID))) = 0 Then '排开采集方法
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            blnFirst = True
                        End If
                    End If
                                        
                    '发送数次:药品为剂量单位的总量,其它为次数
                    If .TextMatrix(i, COL_诊疗类别) = "7" Then
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量))
                    ElseIf InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_住院包装)) * Val(.TextMatrix(i, COL_剂量系数))
                    Else
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量))
                    End If
                    dbl发送数次 = Format(dbl发送数次, "0.00000")
                                                            
                    '首末时间
                    str分解时间 = .TextMatrix(i, COL_分解时间)
                    If str分解时间 <> "" Then
                        str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        str末次时间 = "To_Date('" & Split(str分解时间, ",")(Val(.TextMatrix(i, COL_次数)) - 1) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '无法分解或为"一次性"临嘱
                        str首次时间 = "NULL"
                        str末次时间 = "NULL"
                    End If

                    rsSQL.AddNew
                    rsSQL!类型 = 4: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    
                    rsSQL!SQL = "ZL_病人医嘱发送_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng发送号 & "," & IIF(bln到门诊, 1, 2) & ",'" & strNO & "'," & _
                        lng发送序号 & "," & dbl发送数次 & "," & str首次时间 & "," & str末次时间 & "," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "0," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & int计费状态 & "," & IIF(blnFirst, 1, 0) & ")"
                    rsSQL.Update
                    
                    '要发送的尚未签名的新开医嘱ID(组ID,一组中的叮嘱也会被签名)
                    If Val(.TextMatrix(i, COL_签名ID)) = 0 And Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            lng组ID = Val(.TextMatrix(i, COL_相关ID))
                        Else
                            lng组ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(str医嘱IDs & ",", "," & lng组ID & ",") = 0 Then
                            str医嘱IDs = str医嘱IDs & "," & lng组ID
                        End If
                    End If
                End If
                
                '计算中药配方数
                If .Cell(flexcpData, i, COL_ID) = 3 Then '中药用法
                    int配方数 = int配方数 + 1
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            txtPer.Text = CInt(psb.Value) & "%"
            txtPer.Refresh
        Next
                
        '提示未审核项目
        If strAudit <> "" Then
            MsgBox "病人""" & mrsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
            GoTo errH
        End If
        
        '自动进行电子签名(未签名部份)
        '-----------------------------------------------------------------------------------------
        If Not gobjESign Is Nothing And Mid(gstrESign, 2, 1) = "1" And str医嘱IDs <> "" Then
            str医嘱IDs = Mid(str医嘱IDs, 2) '这里是组ID,返回为明细的ID
            intRule = ReadAdviceSignSource(1, mlng病人ID, mlng主页ID, str医嘱IDs, 0, False, strSource, mlng前提ID)
            If intRule = 0 Then GoTo FuncEnd
            If strSource = "" Then
                Screen.MousePointer = 0
                MsgBox "不能读取要签名的医嘱源文。", vbInformation, gstrSysName
                GoTo FuncEnd
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
            If strSign = "" Then GoTo FuncEnd
            lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
            rsSQL.AddNew
            rsSQL!类型 = 2: rsSQL!医嘱ID = 0: rsSQL!项目ID = 0: rsSQL!序号 = 0
            rsSQL!SQL = "zl_医嘱签名记录_Insert(" & lng签名ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱IDs & "')"
            rsSQL.Update
        End If
        
        '提交病人数据
        '-----------------------------------------------------------------------------------------
        If Not CompletePatiSend(bln到门诊, rsSQL, rsUpload, cur合计, str类别, str类别名称, strWarn, intWarn, blnTran) Then GoTo errH
    End With
    SendAdvice = lng发送号
FuncEnd:
    '删除所有已成功发送的行
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If Err.Number <> 0 Then
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function CompletePatiSend(ByVal bln到门诊 As Boolean, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur合计 As Currency, ByVal str类别 As String, ByVal str类别名称 As String, _
    strWarn As String, intWarn As Integer, blnTran As Boolean) As Boolean
'功能：提交一个病人的医嘱发送数据,在这之前处理记帐报警
'参数：
'      rsSQL=包含所有要执行的SQL
'      rsUpload=用于医保上传的记帐单据号
'      cur合计=病人本次要发送医嘱的记帐金额合计,用于记帐报警
'      str类别=病人本次发送记帐费用的收费类别,用于记帐报警
'      str类别=病人本次发送记帐费用的收费类别名称,用于记帐报警
'      strWarn(I/O)=用于记录当前病人已报警类别
'      intWarn(I/O)=用于记录本次发送报警提示时的选择项
'说明：如果出错,则在调用函数中处理,blnTran返回是否启用了事务
    Dim rsWarn As New ADODB.Recordset
    Dim strSQL As String, intR As Integer
    Dim cur当日 As Currency, i As Long
    Dim arrNOs() As String, strMsg As String
    
    '病人费用报警
    If Not bln到门诊 And cur合计 > 0 Then
        strSQL = "Select Nvl(适用病人,1) as 适用病人,Nvl(报警方法,1) as 报警方法," & _
            " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
            " Where 病区ID=[1] And Nvl(适用病人,1)=[2]"
        Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsPati!当前病区ID), IIF(Nvl(mrsPati!医保, 0) = 1, 2, 1))
        If Not rsWarn.EOF Then
            If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(mlng病人ID)
            str类别名称 = Mid(str类别名称, 2)
            For i = 1 To Len(str类别)
                intR = BillingWarn(Me, mstrPrivs, rsWarn, mrsPati!姓名, Nvl(mrsPati!剩余款, 0), cur当日, cur合计, Nvl(mrsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名称, ",")(i - 1), strWarn, intWarn, Nvl(mrsPati!医保, 0) = 1)
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        '执行顺序:1-计价,2-签名,3-校对,4-发送,5-费用,6-发料
        '1.对费用记录按收费细目ID排序插入
        rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        rsSQL.Sort = "类型,项目ID,序号"
        rsUpload.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!SQL, Me.Caption)
            rsSQL.MoveNext
        Loop
            
        '医保数据上传
        If Not IsNull(mrsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, , mrsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , mrsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '因为现在一张NO内肯定为一个病人的,所以最后病人参数可以不传
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!险类) Then
                        '未提交前上传失败则回滚并中止发送
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName '每张提示
                        Else
                            MsgBox mrsPati!姓名 & "的费用上传失败，发送操作将被中止。", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName '每张提示
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
        gcnOracle.CommitTrans: blnTran = False
        
        '医保数据上传
        If Not IsNull(mrsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, , mrsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, , mrsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!险类) Then
                        '提交后上传失败,仅提示
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox mrsPati!姓名 & "的记帐单""" & rsUpload!NO & """上传失败，HIS端数据已提交，按确定继续发送。", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
            
        '提交成功,将病人医嘱行标记为可删除
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    .RowData(i) = -1
                End If
            Next
        End With
    End If
    CompletePatiSend = True
End Function

Private Sub ShowSendTotal()
'功能：根据当前选择要发送的医嘱，计算并显示发送的医嘱合计
    Dim cur金额 As Currency, cur药品金额 As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '可见行的金额:是一组的汇总金额
                If Not .RowHidden(i) Then
                    cur金额 = cur金额 + Val(.TextMatrix(i, COL_金额))
                End If
                '药品的金额,取原始金额
                If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    cur药品金额 = cur药品金额 + Val(.Cell(flexcpData, i, COL_金额))
                End If
            End If
        Next
    End With
    stbThis.Panels(3).Text = "金额:" & FormatEx(cur金额, gbytDec) & "(药" & FormatEx(cur药品金额, gbytDec) & ")"
    Call Form_Resize
End Sub
