VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceOperate 
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmAdviceOperate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   10875
   StartUpPosition =   3  '窗口缺省
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   10875
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   450
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   450
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   10755
         _ExtentX        =   18971
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
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "全选"
               Object.ToolTipText     =   "全选(Ctrl+A)"
               Object.Tag             =   "全选"
               ImageKey        =   "全选"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "全清"
               Object.ToolTipText     =   "全清(Ctrl+R)"
               Object.Tag             =   "全清"
               ImageKey        =   "全清"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "执行"
               Key             =   "执行"
               Description     =   "执行"
               Object.ToolTipText     =   "执行(Ctrl+E)"
               Object.Tag             =   "执行"
               ImageKey        =   "执行"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Description     =   "重置"
               Object.ToolTipText     =   "重新设置条件(F12)"
               Object.Tag             =   "重置"
               ImageKey        =   "重置"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "刷新"
               Description     =   "刷新"
               Object.ToolTipText     =   "重新读取数据(F5)"
               Object.Tag             =   "刷新"
               ImageKey        =   "刷新"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10875
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   10875
      Begin VB.CommandButton cmdAlley 
         Caption         =   "过敏史/病生状态"
         Height          =   350
         Left            =   9285
         TabIndex        =   4
         Top             =   20
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.ComboBox cbo医生 
         Height          =   300
         Left            =   9285
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   45
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lbl医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "停止医生(&D)"
         Height          =   180
         Left            =   8250
         TabIndex        =   2
         Top             =   105
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblPati 
         BackStyle       =   0  'Transparent
         Caption         =   "姓名: 住院号: 床号: 科室:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   105
         Width           =   6825
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6735
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6930
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   1560
      TabIndex        =   12
      Top             =   6885
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Align           =   1  'Align Top
      Height          =   1590
      Left            =   0
      TabIndex        =   1
      Top             =   5220
      Visible         =   0   'False
      Width           =   10875
      _cx             =   19182
      _cy             =   2805
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
      Rows            =   6
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
      Editable        =   2
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
   Begin VB.PictureBox picUD 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   10875
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   10875
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Align           =   1  'Align Top
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   10875
      _cx             =   19182
      _cy             =   7541
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
      SheetBorder     =   -2147483643
      FocusRect       =   2
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
      FormatString    =   $"frmAdviceOperate.frx":058A
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList img16 
         Left            =   2235
         Top             =   855
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
               Picture         =   "frmAdviceOperate.frx":0625
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":0BBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1159
               Key             =   "签名"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   2835
         Top             =   840
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
               Picture         =   "frmAdviceOperate.frx":14AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":17A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1A9F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1D99
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":2093
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":238D
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":25A7
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":27C1
            Key             =   "执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":29DB
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":2BF5
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":2E0F
            Key             =   "刷新"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3509
            Key             =   "重置"
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3723
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":393D
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3B57
            Key             =   "执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3D71
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3F8B
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":41A5
            Key             =   "刷新"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":489F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6825
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceOperate.frx":4AB9
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12859
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":534D
            Text            =   "通过"
            TextSave        =   "通过"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":5937
            Text            =   "疑问"
            TextSave        =   "疑问"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":5F21
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":655B
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnuPass 
      Caption         =   "Pass"
      Visible         =   0   'False
      Begin VB.Menu mnuPassItem 
         Caption         =   "药物临床信息参考(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "药品说明书(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "中国药典(&N)"
         Index           =   2
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "病人用药教育(&S)"
         Index           =   3
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "检验值(&T)"
         Index           =   4
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "专项信息(&P)"
         Index           =   6
         Begin VB.Menu mnuPassSpec 
            Caption         =   "药物-药物相互作用(&D)"
            Index           =   0
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "药物-食物相互作用(&F)"
            Index           =   1
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "国内注射剂配伍(&M)"
            Index           =   3
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "国外注射剂配伍(&T)"
            Index           =   4
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "禁忌症(&C)"
            Index           =   6
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "副作用(&S)"
            Index           =   7
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "老年人用药(&G)"
            Index           =   9
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "儿童用药(&P)"
            Index           =   10
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "妊娠期用药(&E)"
            Index           =   11
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "哺乳期用药(&L)"
            Index           =   12
         End
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "医药信息中心(&I)"
         Index           =   8
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "药品配对信息(&M)"
         Index           =   10
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "给药途径配对信息(&R)"
         Index           =   11
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "医院药品信息(&F)"
         Index           =   12
      End
   End
End
Attribute VB_Name = "frmAdviceOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'功能：
'0-医嘱作废:
'    只需选择要作废的医嘱
'1-停止医嘱:
'    需要指定终止时间(缺省为当前,次日生效缺省为次日零点,预定的不变)
'    护士停时需要指定停止医生
'2-确认停止:
'    只需选择需要确认的医嘱
'3-校对医嘱:
'    补录的医嘱可以修改校对时间(非补录的缺省为当前不可改,补录的缺省为开嘱时间+1m)
'4-调整计价项目:
'    增删改每个医嘱的计价项目
'5-暂停医嘱
'    选择需要暂停的医嘱
'6-启用医嘱
'    选择需要启用的医嘱
Public mstrPrivs As String
Public mlng医嘱ID As Long '用于缺省定位
Public mint类型 As Integer '0-医嘱作废,1-停止医嘱,2-确认停止,3-医嘱校对,4-调整计价项目,5-暂停医嘱,6-启用医嘱
Public mlng病区ID As Long
Public mlng病人ID As Long
Public mlng主页ID As Long
Public mbln护士站 As Boolean
Public mblnOK As Boolean

Private mrsPrice As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mstrLike As String
Private mblnReturn As Boolean
Private mint简码 As Integer
Private mstrRollNotify As String '操作后要进行超期收回提醒的病人(病人ID,主页ID;...)
Private mlngPassPati As Long 'Pass:上次已传入PASS的病人ID

'重置条件
Private mblnFirst As Boolean
Private mstr病人IDs As String
Private mint期效 As Integer
Private mint类别 As Integer
Private mblnPauseLast As Boolean

'隐藏列
Private Const COL_ID = 0
Private Const COL_相关ID = 1
Private Const COL_组ID = 2
Private Const COL_组号 = 3
Private Const COL_诊疗类别 = 4
Private Const COL_毒理分类 = 5
Private Const COL_类型 = 6 '1-中药配方,2-检验组合
'Pass警示列
Private Const COL_警示 = 7
'输入列
Private Const COL_选择 = 8 '
Private Const COL_输入 = 9 '
'可见列
Private Const COL_姓名 = 10
Private Const COL_住院号 = 11
Private Const COL_床号 = 12
Private Const COL_婴儿 = 13
Private Const COL_期效 = 14
Private Const COL_开嘱时间 = 15
Private Const COL_开始时间 = 16
Private Const COL_医嘱内容 = 17
Private Const COL_皮试 = 18
Private Const COL_总量 = 19
Private Const COL_单量 = 20
Private Const COL_频率 = 21
Private Const COL_用法 = 22
Private Const COL_医生嘱托 = 23
Private Const COL_执行时间 = 24
Private Const COL_终止时间 = 25 '
Private Const COL_执行科室 = 26
Private Const COL_执行性质 = 27
Private Const COL_上次执行 = 28 '
Private Const COL_标志 = 29
Private Const COL_开嘱医生 = 30
Private Const COL_校对护士 = 31 '
Private Const COL_校对时间 = 32 '
Private Const COL_停嘱医生 = 33 '
Private Const COL_停嘱时间 = 34 '
'隐藏
Private Const COL_病人ID = 35
Private Const COL_主页ID = 36
Private Const COL_操作类型 = 37
Private Const COL_执行科室ID = 38
Private Const COL_病人科室ID = 39
Private Const COL_收费细目ID = 40
Private Const COL_单量单位 = 41
Private Const COL_前提ID = 42
Private Const COL_签名ID = 43
Private Const COL_操作人员 = 44

'计价清单的列值
Private Const COLP_医嘱ID = 0 '附加存放变价信息
Private Const COLP_相关ID = 1 '附加存放变价信息
Private Const COLP_诊疗类别 = 2 '附加存放变价信息
Private Const COLP_诊疗项目ID = 3
Private Const COLP_收费细目ID = 4
Private Const COLP_固定 = 5
Private Const COLP_计价医嘱 = 6
Private Const COLP_类别 = 7 '收费类别名称
Private Const COLP_收费项目 = 8
Private Const COLP_单位 = 9
Private Const COLP_数量 = 10
Private Const COLP_单价 = 11
Private Const COLP_执行科室 = 12
Private Const COLP_费用类型 = 13
Private Const COLP_从项 = 14
Private Const COLP_收费类别 = 15
Private Const COLP_执行科室ID = 16
Private Const COLP_跟踪在用 = 17

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

Private Sub cmdAlley_Click()
'功能：对病人过敏史/病生状态进行查看
    'Pass
    Call AdviceCheckWarn(21, vsAdvice.Row)
End Sub

Private Function ResetCond() As Boolean
'功能：重置发送条件
    Dim blnSeek As Boolean
    Me.Refresh
    With frmAdviceOperateCond
        .mstrPrivs = mstrPrivs
        .mint类型 = mint类型
        .mlng病区ID = mlng病区ID
        .mlng病人ID = mlng病人ID
        .Show 1, Me
        If .mblnOK Then
            mlng病区ID = .mlng病区ID
            mstr病人IDs = .mstr病人IDs
            mint期效 = .mint期效
            mint类别 = .mint类别
            mblnPauseLast = .mblnPauseLast
                        
            '只选择了当前病人才定位当前医嘱
            If UBound(Split(mstr病人IDs, ";")) = 0 Then
                If Val(Split(mstr病人IDs, ",")(0)) = mlng病人ID Then blnSeek = True
            End If
            Call RefreshData(IIF(blnSeek, mlng医嘱ID, 0), True)
        End If
        ResetCond = .mblnOK
    End With
End Function

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        If tbr.Buttons("重置").Visible Then
            If Not ResetCond Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全选"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全清"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("执行"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("退出"))
    ElseIf KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("帮助"))
    ElseIf KeyCode = vbKeyF5 Then
        Call tbr_ButtonClick(tbr.Buttons("刷新"))
    ElseIf KeyCode = vbKeyF12 Then
        If tbr.Buttons("重置").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("重置"))
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
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Call InitAdviceTable
    Call SetAdviceCol '先设置一次列属性,以便正确恢复个性化
    If mint类型 = 3 Or mint类型 = 4 Then
        Call InitPriceTable
    End If
    Call RestoreWinState(Me, App.ProductName, mint类型)
    
    mblnOK = False
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
    If Not (mint类型 = 3 Or mint类型 = 4) Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    
    '设置重置可作否,缺省重置条件
    mblnFirst = True
    mblnPauseLast = False
    mint期效 = 0: mint类别 = 0
    mstr病人IDs = mlng病人ID & "," & mlng主页ID
    If mbln护士站 And InStr(",3,5,6,", mint类型) > 0 Then
        If mint类型 = 3 Then
            tbr.Buttons("重置").Enabled = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "批量医嘱校对", 0)) <> 0
        ElseIf mint类型 = 5 Or mint类型 = 6 Then
            tbr.Buttons("重置").Enabled = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "批量医嘱启停", 0)) <> 0
        End If
    Else
        tbr.Buttons("重置").Enabled = False
    End If
    tbr.Buttons("重置").Visible = tbr.Buttons("重置").Enabled 'Enabled用于判断
    
    If mint类型 = 0 Then
        Caption = "病人医嘱作废"
        tbr.Buttons("执行").Caption = "作废"
        tbr.Buttons("执行").ToolTipText = "作废选择的医嘱(Ctrl+E)"
    ElseIf mint类型 = 1 Then
        Caption = "病人医嘱停止"
        tbr.Buttons("执行").Caption = "停止"
        tbr.Buttons("执行").ToolTipText = "停止选择的医嘱(Ctrl+E)"
        If mbln护士站 Then
            lbl医生.Visible = True
            cbo医生.Visible = True
        End If
    ElseIf mint类型 = 2 Then
        Caption = "确认医嘱停止"
        tbr.Buttons("执行").Caption = "确认"
        tbr.Buttons("执行").ToolTipText = "确认选择的医嘱(Ctrl+E)"
    ElseIf mint类型 = 3 Then
        Caption = "病人医嘱校对"
        tbr.Buttons("执行").Caption = "校对"
        tbr.Buttons("执行").ToolTipText = "确认选择的医嘱(Ctrl+E)"
        
        stbThis.Panels(3).Visible = True
        stbThis.Panels(4).Visible = True
        
        picUD.Visible = True
        vsPrice.Visible = True
        
        '病人过敏史/病生状态可用检测
        mlngPassPati = 0
        If gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 Then 'Pass
            cmdAlley.Visible = True
            vsAdvice.ColHidden(COL_警示) = False
            cmdAlley.Enabled = PassGetState("AlleyEnable") = 1
        End If
    ElseIf mint类型 = 4 Then
        Caption = "调整计价项目"
        tbr.Buttons("执行").Caption = "确认"
        tbr.Buttons("执行").ToolTipText = "确认选择项目的价目(Ctrl+E)"
        
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint类型 = 5 Then
        Caption = "病人医嘱暂停"
        tbr.Buttons("执行").Caption = "暂停"
        tbr.Buttons("执行").ToolTipText = "暂停选择的医嘱(Ctrl+E)"
    ElseIf mint类型 = 6 Then
        Caption = "病人医嘱启用"
        tbr.Buttons("执行").Caption = "启用"
        tbr.Buttons("执行").ToolTipText = "启用选择的医嘱(Ctrl+E)"
    End If
        
    '读取部门信息
    If mint类型 = 3 Or mint类型 = 4 Then
        strSQL = "Select ID,名称 From 部门表"
        Set mrsDept = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDept, strSQL, Me.Caption)
    End If
    
    '显示病人信息：一个病人操作的情况
    strSQL = _
        " Select A.住院号,A.姓名,A.性别,A.年龄,B.出院病床," & _
        " B.住院医师,B.出院科室ID,C.名称 as 科室" & _
        " From 病人信息 A,病案主页 B,部门表 C" & _
        " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    lblPati.Caption = "姓名:" & rsTmp!姓名 & "　住院号:" & Nvl(rsTmp!住院号) & _
        "　床号:" & Nvl(rsTmp!出院病床) & "　科室:" & Nvl(rsTmp!科室)
    
    '可选的停嘱医生:缺省为病人的住院医师或病人科室的第一个医生
    '目前不支持批量停止医嘱,因此肯定是以传入的当前病人为准读取
    If mint类型 = 1 And mbln护士站 Then
        Call Get开嘱医生(rsTmp!出院科室ID, True, Nvl(rsTmp!住院医师), 0, cbo医生)
        If cbo医生.ListIndex = -1 And cbo医生.ListCount > 0 Then cbo医生.ListIndex = 0
    End If
    
    '显示医嘱内容
    If Not tbr.Buttons("重置").Enabled Then Call RefreshData(mlng医嘱ID, True)
End Sub

Private Sub RefreshData(Optional ByVal lng医嘱ID As Long, Optional ByVal blnNotify As Boolean)
'功能：刷新数据
'参数：lng医嘱ID=用于医嘱定位
'      blnNotify=是否提醒特殊医嘱
    Dim blnChange As Boolean, i As Long
    Dim strPatis As String, arrPatis As Variant
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strMsg As String, strTmp As String
    
    '显示医嘱内容
    Call LoadAdvice(strPatis)
    
    '读取计价数据
    If mint类型 = 3 Or mint类型 = 4 Then
        Call InitPriceRecordset
        Screen.MousePointer = 11
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            Progress = i / (vsAdvice.Rows - 1) * 100
            blnChange = False
            Call LoadPrice(i, blnChange)
            If blnChange And mint类型 = 4 Then Call SelectRow(i)
        Next
        Progress = 0: Screen.MousePointer = 0
    End If
    
    If lng医嘱ID <> 0 Then
        i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If i <> -1 Then vsAdvice.Row = i
    End If
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    '特殊医嘱提醒
    If blnNotify And InStr(",3,4,6,", mint类型) > 0 And strPatis <> "" Then
        arrPatis = Split(strPatis, ";")
        For i = 0 To UBound(arrPatis)
            lng病人ID = Split(arrPatis(i), ",")(0)
            lng主页ID = Split(arrPatis(i), ",")(1)
            strTmp = ExistsSpecAdvice(lng病人ID, lng主页ID)
            If strTmp <> "" Then
                strTmp = Replace(Replace(strTmp, "提醒您，", ""), vbCrLf & vbCrLf, vbCrLf)
                strMsg = strMsg & vbCrLf & strTmp
            End If
        Next
        If strMsg <> "" Then MsgBox Mid(strMsg, 3), vbInformation, gstrSysName & " - 提醒您"
    End If
End Sub

Private Sub SelectRow(ByVal lngRow As Long)
'功能：使指定行选中(包括一并给药)
    With vsAdvice
        If mint类型 = 3 Then
            Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages(1).Picture
            .Cell(flexcpData, lngRow, COL_选择) = 1
        Else
            .TextMatrix(lngRow, COL_选择) = -1 '直接对TextMatrix时,不要用True
        End If
    End With
    Call vsAdvice_AfterEdit(lngRow, COL_选择)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsAdvice.Height = Me.ScaleHeight - cbr.Height - stbThis.Height - picPati.Height - IIF(picUD.Visible, picUD.Height + vsPrice.Height, 0)
    If cbo医生.Visible Then
        lblPati.Width = Me.ScaleWidth - lbl医生.Width - cbo医生.Width - lblPati.Left - 350
        cbo医生.Left = Me.ScaleWidth - cbo医生.Width - 200
        lbl医生.Left = cbo医生.Left - lbl医生.Width - 45
    ElseIf cmdAlley.Visible Then
        lblPati.Width = Me.ScaleWidth - cmdAlley.Width - lblPati.Left - 350
        cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 200
    Else
        lblPati.Width = Me.ScaleWidth - lblPati.Left
    End If
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mint类型)
    
    Set mrsPrice = Nothing
    Set mrsDept = Nothing
    mstrPrivs = ""
    mlng医嘱ID = 0
    mint类型 = 0
    mlng病区ID = 0
    mlng病人ID = 0
    mlng主页ID = 0
    mbln护士站 = False
End Sub

Private Sub picUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        vsAdvice.Height = vsAdvice.Height + y
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

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    Select Case Button.Key
        Case "全选"
            If vsAdvice.ColHidden(COL_选择) Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows + 1 And Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, COL_ID)) = 0 Then Exit Sub
            
            If mint类型 = 3 Then
                For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                    If vsAdvice.Cell(flexcpData, i, COL_选择) = Empty Then '保持疑问的不变
                        Set vsAdvice.Cell(flexcpPicture, i, COL_选择) = img16.ListImages(1).Picture
                        vsAdvice.Cell(flexcpData, i, COL_选择) = 1
                    End If
                Next
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = True
            End If
        Case "全清"
            If mint类型 = 3 Then
                Set vsAdvice.Cell(flexcpPicture, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = Nothing
                vsAdvice.Cell(flexcpData, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = Empty
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = False
            End If
        Case "刷新"
            Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        Case "重置"
            Call ResetCond
        Case "执行"
            If Not CheckValid Then Exit Sub
            If Not CheckSignValid Then Exit Sub
            If ExecuteOperate Then
                '医嘱校对时检查并提醒超期收回(自动)停止的医嘱
                If mint类型 = 3 And mstrRollNotify <> "" Then
                    Call ShowRollNotify
                End If
                
                mblnOK = True: Unload Me
            End If
        Case "帮助"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub ShowRollNotify()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim lng病人ID As Long, lng主页ID As Long, i As Long
    
    On Error GoTo errH
    
    For i = 0 To UBound(Split(mstrRollNotify, ";"))
        '条件与超期收回中一致，但只包含当前状态为(自动)停止的。
        strSQL = "(A.执行时间方案 is NULL And (Nvl(A.频率次数,0)=0 Or Nvl(A.频率间隔,0)=0 Or A.频率间隔 is NULL))"
        strSQL = _
            " Select A.姓名,A.医嘱内容 From 病人医嘱记录 A,诊疗项目目录 E" & _
            " Where A.诊疗项目ID=E.ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And Not(A.诊疗类别='H' And E.操作类型='1') And Not(A.诊疗类别='Z' And E.操作类型='4')" & _
            " And Nvl(A.执行性质,0)<>0 And A.总给予量 is NULL And Nvl(A.医嘱期效,0)=0" & _
            " And ((Not " & strSQL & " And A.执行终止时间<A.上次执行时间)" & _
            " Or (" & strSQL & " And Trunc(A.执行终止时间)<Trunc(A.上次执行时间)+1))" & _
            " And A.医嘱状态=8 And (A.相关ID is Null Or A.诊疗类别 IN('5','6'))" & _
            " And A.开始执行时间 is Not NULL And A.病人来源<>3 And Not Exists(" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 诊疗类别 IN('5','6') And X.相关ID=A.ID" & _
                " And 病人ID=[1] And 主页ID=[2])" & _
            " Order by A.序号"
        lng病人ID = Split(Split(mstrRollNotify, ";")(i), ",")(0)
        lng主页ID = Split(Split(mstrRollNotify, ";")(i), ",")(1)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            strMsg = strMsg & vbCrLf & vbCrLf & "病人""" & Nvl(rsTmp!姓名) & """的医嘱："
            Do While Not rsTmp.EOF
                strMsg = strMsg & vbCrLf & "●　" & rsTmp!医嘱内容
                rsTmp.MoveNext
            Loop
        End If
    Next
    If strMsg <> "" Then
        MsgBox "下列已停止的病人医嘱被超期发送：" & strMsg & vbCrLf & vbCrLf & "该类医嘱可以在护士工作站中使用""超期发送收回""进行处理。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    '显示计价项目
    If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
        If (mint类型 = 3 Or mint类型 = 4) And Not mrsPrice Is Nothing Then
            Call ShowPrice(NewRow)
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_医嘱内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：一并给药的一起输入
    Dim lngBegin As Long, lngEnd As Long
    Dim strTmp As String, vPause As Date, i As Long
        
    With vsAdvice
        If Col = COL_输入 And Not mblnReturn Then
            '非回车焦点转移确认
            strTmp = .TextMatrix(Row, Col)
            If strTmp <> "" Then strTmp = GetFullDate(strTmp)
            If Not IsDate(strTmp) Then
                .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
            End If
            
            If mint类型 = 1 Then '检查终止时间
                If IsDate(.Cell(flexcpData, Row, COL_上次执行)) Then
                    If .TextMatrix(Row, COL_执行时间) = "" And Format(.TextMatrix(Row, COL_上次执行), "HH:mm") = "00:00" Then
                        '"持续性"长嘱,停止当天不发送
                        If Format(strTmp, "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, Row, COL_上次执行)) + 1, "yyyy-MM-dd") Then
                            .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                        End If
                    Else
                        If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                            .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                        End If
                    End If
                End If
                If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
                If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
            ElseIf mint类型 = 2 Then  '检查确认停止时间
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
            ElseIf mint类型 = 3 Then
                '不能小于开嘱时间,开始时间较小者(开始时间可以改成比开嘱时间小)
                If Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                Else
                    If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            ElseIf mint类型 = 5 Then '检查暂停时间
                '应>=开始执行时间,因为该时间点尚未执行
                If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
                '应>上次执行时间,因为该时间点已执行
                If .TextMatrix(Row, COL_上次执行) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                '应<执行终止时间,因为该时间点执行有效
                If .TextMatrix(Row, COL_终止时间) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
                If vPause <> CDate(0) Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            ElseIf mint类型 = 6 Then '检查启用时间
                '应>暂停时间
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
                If vPause <> CDate(0) Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                '应<=执行终止时间
                If .TextMatrix(Row, COL_终止时间) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            End If
            .TextMatrix(Row, Col) = strTmp
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '一并给药一起更改
            mblnReturn = True
            Call vsAdvice_AfterEdit(Row, Col)
        Else
            '一并给药的一起选择或输入
            If (Col = COL_选择 Or Col = COL_输入) And InStr(",5,6,", .TextMatrix(Row, COL_诊疗类别)) > 0 Then
                If RowIn一并给药(Row, lngBegin, lngEnd) Then
                    For i = lngBegin To lngEnd
                        If i <> Row Then
                            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                            .Cell(flexcpData, i, Col) = .Cell(flexcpData, Row, Col)
                            Set .Cell(flexcpPicture, i, Col) = .Cell(flexcpPicture, Row, Col)
                        End If
                    Next
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Or Col = COL_输入 Or Col = COL_警示 Then Cancel = True 'Pass
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If mint类型 = 3 And .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
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
        lngLeft = COL_婴儿: lngRight = COL_开始时间
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_频率: lngRight = COL_用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
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
'功能：定位到下一输入单元或输入校对标志
    Dim blnGroup As Boolean, i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        With vsAdvice
            If .ColHidden(COL_选择) And .ColHidden(COL_输入) Then
                If .Row + 1 <= .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    .Row = .FixedRows
                End If
            Else
                If .Col = COL_选择 Then
                    If Not .ColHidden(COL_输入) Then
                        .Col = COL_输入
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                    End If
                ElseIf .Col = COL_输入 Then
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    .Col = COL_选择
                Else
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    If Not .ColHidden(COL_选择) Then .Col = COL_选择
                End If
            End If
            Call .ShowCell(.Row, .Col)
        End With
    ElseIf KeyAscii = 32 Then
        With vsAdvice
            If mint类型 = 3 And .Col = COL_选择 Then
                KeyAscii = 0
                
                If .Cell(flexcpData, .Row, .Col) = Empty Then
                    Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages(1).Picture
                    .Cell(flexcpData, .Row, .Col) = 1
                ElseIf .Cell(flexcpData, .Row, .Col) = 1 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages(2).Picture
                    .Cell(flexcpData, .Row, .Col) = 2
                ElseIf .Cell(flexcpData, .Row, .Col) = 2 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                    .Cell(flexcpData, .Row, .Col) = Empty
                End If
            
                If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
                    If .Row - 1 >= .FixedRows Then
                        blnGroup = Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID))
                    End If
                    If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                        blnGroup = Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID))
                    End If
                    If blnGroup Then
                        For i = .Row - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                        For i = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strTmp As String, vPause As Date
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        mblnReturn = True
        With vsAdvice
            '检查输入的有效性
            If .EditText <> "" Then .EditText = GetFullDate(.EditText)
            If Not IsDate(.EditText) Then
                MsgBox "请输入一个有效的" & .TextMatrix(0, Col) & " 。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            If mint类型 = 1 Then '检查终止时间
                '必须大于开始执行时间
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "输入的执行终止时间必须大于医嘱的开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '不能小于开嘱时间
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "输入的执行终止时间不应小于开嘱时间 " & Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '不应小于上次执行时间
                If IsDate(.Cell(flexcpData, Row, COL_上次执行)) Then
                    If .TextMatrix(Row, COL_执行时间) = "" And Format(.TextMatrix(Row, COL_上次执行), "HH:mm") = "00:00" Then

                        '"持续性"长嘱,停止当天不发送
                        If Format(.EditText, "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, Row, COL_上次执行)) + 1, "yyyy-MM-dd") Then
                            strTmp = .EditText 'MsgBox一现,EditText就空了,所以要记录
                            If MsgBox("对持续性长嘱，执行终止日期应晚于上次执行日期 " & Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                            End If
                        End If
                    Else
                        If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                            strTmp = .EditText 'MsgBox一现,EditText就空了,所以要记录
                            If MsgBox("输入的执行终止时间小于医嘱的上次执行时间 " & Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                            End If
                        End If
                    End If
                End If
            ElseIf mint类型 = 2 Then  '检查确认停止时间
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "确认停止医嘱的时间不能小于医嘱的执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
            ElseIf mint类型 = 3 Then  '检查校对时间(补录的才能改)
                '不能小于开嘱时间,开始时间较小者(开始时间可以改成比开嘱时间小)
                If Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then
                        MsgBox "输入的校对时间不能小于开嘱时间 " & Format(.Cell(flexcpData, Row, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                Else
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                        MsgBox "输入的校对时间不能小于医嘱的开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            ElseIf mint类型 = 5 Then '检查暂停时间
                '应>=开始执行时间,因为该时间点尚未执行
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") Then
                    MsgBox "医嘱的暂停时间应大于等于开始执行时间 " & Format(.Cell(flexcpData, Row, COL_开始时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '应>上次执行时间,因为该时间点已执行
                If .TextMatrix(Row, COL_上次执行) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                        MsgBox "医嘱的暂停时间应大于上次执行时间 " & Format(.Cell(flexcpData, Row, COL_上次执行), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                '应<执行终止时间,因为该时间点执行有效
                If .TextMatrix(Row, COL_终止时间) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                        MsgBox "医嘱的暂停时间应小于执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
                If vPause <> CDate(0) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        MsgBox "医嘱的暂停时间应大于上次暂停后的启用时间 " & Format(vPause, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            ElseIf mint类型 = 6 Then '检查启用时间
                '应>暂停时间
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
                If vPause <> CDate(0) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        MsgBox "医嘱的启用时间应大于上次暂停时间 " & Format(vPause, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                
                '应<=执行终止时间
                If .TextMatrix(Row, COL_终止时间) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") Then
                        MsgBox "医嘱的启用时间应小于等于执行终止时间 " & Format(.Cell(flexcpData, Row, COL_终止时间), "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            End If
            .TextMatrix(Row, Col) = IIF(.EditText = "" And strTmp <> "", strTmp, .EditText)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            Call vsAdvice_AfterEdit(Row, Col) '一并给药的一并更改:提示后不会自动执行该事件
            
            '设置为相同时间(校对,暂停,启用)
            If Row = .FixedRows And .Rows > .FixedRows + 1 Then
                If mint类型 = 3 Then
                    If MsgBox("要设置所有的医嘱都在这个时间校对吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call SetSameTime(Row)
                    End If
                ElseIf (mint类型 = 5 Or mint类型 = 6) Then
                    If MsgBox("要设置所有医嘱都在这个时间" & IIF(mint类型 = 5, "暂停", "启用") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call SetSameTime(Row)
                    End If
                End If
            End If
            Call vsAdvice_KeyPress(13) '定位到一下输入单元
        End With
    Else
        If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnReturn = False
    If Col <> COL_选择 And Col <> COL_输入 Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(Row, COL_ID)) = 0 Then
        Cancel = True
    ElseIf mint类型 = 1 And Col = COL_输入 And vsAdvice.TextMatrix(Row, COL_类型) = "1" Then
        Cancel = True '停止医嘱时,中药配方(长嘱)的终止时间不可修改
    ElseIf mint类型 = 3 Then
        If Col = COL_输入 And Not (vsAdvice.TextMatrix(Row, COL_标志) = "补录" Or InStr(mstrPrivs, "修改校对时间") > 0) Then
            Cancel = True '校对医嘱时,非补录的校对时间不可更改
        ElseIf Col = COL_选择 Then
            Cancel = True '不能直接编辑
        End If
    End If
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;相关ID;组ID;组号;诊疗类别;毒理分类;中药;,240,4;,300,4;,1530,1;" & _
        "姓名,750,1;住院号,750,1;床号,500,1;婴儿,500,1;期效,500,4;开嘱时间,1080,1;开始时间,1080,1;" & _
        "医嘱内容,3000,1;,375,4;总量,850,1;单量,850,1;频率,1000,1;用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;" & _
        "终止时间,1080,1;执行科室,850,1;执行性质,850,1;上次执行,1080,1;标志,500,4;" & _
        "开嘱医生,850,1;校对护士,850,1;校对时间,1080,1;停嘱医生,850,1;停嘱时间,1080,1;" & _
        "病人ID;主页ID;操作类型;执行科室ID;病人科室ID;收费细目ID;单量单位;前提ID;签名ID;操作人员"
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
        .ColHidden(COL_警示) = True 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "医嘱ID;相关ID;诊疗类别;诊疗项目ID;收费细目ID;固定;" & _
        "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;单位,500,4;数量,650,1;单价,850,7;" & _
        "执行科室,1000,1;费用类型,850,1;从项,450,4;收费类别;执行科室ID;跟踪在用"
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

Private Sub SetAdviceCol()
'功能：设置一些可见列及编辑属性,应在表格数据装入后调用
    With vsAdvice
        .TextMatrix(0, COL_选择) = ""
        .Editable = flexEDKbdMouse
        
        '根据情况设置列的可见性
        If mint类型 = 0 Then
            '医嘱作废
            .ColHidden(COL_输入) = True
            .ColHidden(COL_上次执行) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 1 Then
            '停止医嘱
            .TextMatrix(0, COL_输入) = "终止时间"
            .ColHidden(COL_终止时间) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 2 Then
            '确认停止
            .TextMatrix(0, COL_输入) = "确认时间"
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 3 Then
            '医嘱校对
            .TextMatrix(0, COL_输入) = "校对时间"
            .ColHidden(COL_上次执行) = True
            .ColHidden(COL_校对护士) = True
            .ColHidden(COL_校对时间) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .Cell(flexcpPictureAlignment, .FixedRows, COL_选择, .Rows - 1, COL_选择) = 4
        ElseIf mint类型 = 4 Then
            '调整计价项目
            .ColHidden(COL_输入) = True
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 5 Then
            '暂停医嘱
            .TextMatrix(0, COL_输入) = "暂停时间"
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        ElseIf mint类型 = 6 Then
            '启用医嘱
            .TextMatrix(0, COL_输入) = "启用时间"
            .ColHidden(COL_停嘱医生) = True
            .ColHidden(COL_停嘱时间) = True
            .ColDataType(COL_选择) = flexDTBoolean
        End If
        
        '设置冻结列
        If Not .ColHidden(COL_输入) Then
            .FrozenCols = COL_输入 + 1 - .FixedCols
            .SheetBorder = vbBlack
        ElseIf Not .ColHidden(COL_选择) Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
            .SheetBorder = vbBlack
        End If
        
        '可输入列标识
        .Cell(flexcpBackColor, .FixedRows, COL_选择, .Rows - 1, COL_输入) = &HC0FFC0
    End With
End Sub

Private Function GetWhere() As String
'功能：根据窗体功能产生医嘱条件串
'说明：假设"病人医嘱记录"别名为"A"
    Dim strSQL As String
    
    If mint类型 = 0 Then
        '医嘱作废:已校对,但未发送过的临嘱或长嘱。已暂停的长嘱也可以直接作废。
        '临时自由医嘱校对后自动停止，这种也允许作废
        strSQL = " And (A.医嘱状态 Not IN(1,2,4,8,9) And A.上次执行时间 is NULL Or A.医嘱期效=1 And A.诊疗项目ID is Null And A.医嘱状态=8)"
    ElseIf mint类型 = 1 Then
        '停止医嘱:长嘱,已暂停的也可以直接停止,不含中药配方(有付数,自动停)
        strSQL = " And A.医嘱状态 Not IN(1,2,4,8,9) And Nvl(A.医嘱期效,0)=0 And A.总给予量 is NULL"
    ElseIf mint类型 = 2 Then
        '确认停止:停止状态的长嘱(含被自动停止的中药配方长嘱)
        strSQL = " And A.医嘱状态=8 And Nvl(A.医嘱期效,0)=0"
    ElseIf mint类型 = 3 Then
        '医嘱校对:对新开的，开嘱医生具有资格的或已审核的医嘱进行校对。
        strSQL = " And A.医嘱状态=1 And Exists(" & _
            "Select M.姓名 From 人员表 M,执业类别 N" & _
            " Where M.姓名=Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1)" & _
            " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
            " )"
    ElseIf mint类型 = 4 Then
        '调整计价项目
        strSQL = " And A.医嘱状态 Not IN(1,2,4,8,9)"
    ElseIf mint类型 = 5 Then
        '暂停医嘱:长嘱,不含中药配方(有付数,不准暂停)
        strSQL = " And A.医嘱状态 IN(3,5,7) And Nvl(A.医嘱期效,0)=0 And A.总给予量 is NULL"
    ElseIf mint类型 = 6 Then
        '启用医嘱
        strSQL = " And A.医嘱状态=6"
    End If
    GetWhere = strSQL
End Function

Private Function LoadAdvice(strPatis As String) As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
'参数：str病人IDs=用于返回实际有数据的病人串:"病人ID,主页ID,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsPause As New ADODB.Recordset
    Dim str成药 As String, str中药 As String
    Dim strSQL As String, strWhere As String
    Dim vCurDate As Date, bln给药途径 As Boolean
    Dim lng病人ID As Long, lng主页ID As Long
    Dim i As Long, j As Long, k As Long
    Dim str婴儿 As String, str科室s As String, strTmp As String
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
        
    '----------------------------------------------------------------------
    strPatis = ""
    With vsAdvice
        .Rows = .FixedRows
        .ColHidden(COL_姓名) = True
        .ColHidden(COL_住院号) = True
        .ColHidden(COL_床号) = True
        .ColHidden(COL_婴儿) = True
    End With
    
    '----------------------------------------------------------------------
    strWhere = GetWhere
    '医生操作时不显示医技的,护士编辑时显示所有
    strWhere = strWhere & IIF(Not mbln护士站, " And A.前提ID is NULL", "")
    
    '校对的医嘱范围限制
    If mint类型 = 3 And InStr(mstrPrivs, "全院医嘱校对") = 0 Then
        strWhere = strWhere & " And A.开嘱医生 In(" & _
            " Select Distinct B.姓名" & _
            " From 部门人员 A,人员表 B,人员性质说明 C" & _
            " Where A.人员ID=B.ID And B.ID=C.人员ID And C.人员性质='医生'" & _
            "   And A.部门ID In(" & _
            "     Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
            "     Where A.人员ID=(Select 人员ID From 上机人员表 Where 用户名=User)" & _
            "       And A.部门ID=B.病区ID)" & _
            ")"
    End If
    
    '批量操作时设置的条件
    If mint期效 <> 0 Then
        strWhere = strWhere & " And Nvl(A.医嘱期效,0)=" & mint期效 - 1
    End If
    If mint类别 <> 0 Then
        If mint类别 = 1 Then
            '药品类
            strWhere = strWhere & _
                " And (A.诊疗类别 IN('5','6','7')" & _
                " Or (A.诊疗类别='E' And A.相关ID is Not NULL)" & _
                " Or Exists(Select ID From 病人医嘱记录 S Where 诊疗类别 IN('5','6','7') And S.相关ID=A.ID And 病人ID=[1])" & _
                " )"
        ElseIf mint类别 = 2 Then
            '其他类
            strWhere = strWhere & _
                " And Not A.诊疗类别 IN('5','6','7')" & _
                " And Not(A.诊疗类别='E' And A.相关ID is Not NULL)" & _
                " And Not Exists(Select ID From 病人医嘱记录 S Where 诊疗类别 IN('5','6','7') And S.相关ID=A.ID And 病人ID=[1])"
        End If
    End If
    
    vCurDate = zlDatabase.Currentdate
    
    '----------------------------------------------------------------------
    For k = 0 To UBound(Split(mstr病人IDs, ";"))
        lng病人ID = Split(Split(mstr病人IDs, ";")(k), ",")(0)
        lng主页ID = Split(Split(mstr病人IDs, ";")(k), ",")(1)
        
        '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法
        strSQL = _
            "Select /*+ RULE */ A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
                " Nvl(A.诊疗类别,'*') as 诊疗类别,C.毒理分类,NULL as 中药,A.审查结果,NULL as 选择,NULL as 输入," & _
                " P.姓名,P.住院号,P.当前床号 as 床号,Decode(Nvl(A.婴儿,0),0,'病人','婴儿'||A.婴儿) as 婴儿,Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
                " To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间,To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,A.医嘱内容,A.皮试结果 as 皮试," & _
                " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
                " Decode(A.单次用量,NULL,NULL,A.单次用量||B.计算单位) as 单量," & _
                " A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('246',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
                " A.医生嘱托,A.执行时间方案 as 执行时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间," & _
                " Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
                " Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质," & _
                " To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & _
                " Decode(A.紧急标志,1,'紧急',2,'补录','普通') as 标志," & _
                " A.开嘱医生,A.校对护士,To_Char(A.校对时间,'YYYY-MM-DD HH24:MI') as 校对时间," & _
                " A.停嘱医生,To_Char(A.停嘱时间,'YYYY-MM-DD HH24:MI') as 停嘱时间,A.病人ID,A.主页ID," & _
                " B.操作类型,A.执行科室ID,A.病人科室ID,A.收费细目ID,B.计算单位 as 单量单位,A.前提ID,S.签名ID,S.操作人员" & _
            " From 病人医嘱记录 A,病人信息 P,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,病人医嘱状态 S,病人医嘱记录 X" & _
            " Where A.病人ID=P.病人ID And A.诊疗项目ID=B.ID" & IIF(InStr(",0,1,2,3,", mint类型) > 0, "(+)", "") & _
                " And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
                " And A.收费细目ID=D.药品ID(+) And A.相关ID=X.ID(+)" & _
                " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
                " And A.ID=S.医嘱ID And S.操作类型=1 And A.病人ID=[1] And A.主页ID=[2]" & _
                " And A.开始执行时间 is Not NULL And A.病人来源<>3" & strWhere & _
            " Order by Nvl(A.婴儿,0),组号,组ID,A.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        
        If Not rsTmp.EOF Then
            strPatis = strPatis & ";" & lng病人ID & "," & lng主页ID
            If InStr(str科室s & ",", "," & rsTmp!病人科室ID & ",") = 0 Then
                str科室s = str科室s & "," & rsTmp!病人科室ID
            End If
            
            '暂停医嘱时读取医嘱的上次启用时间(不一定有)
            '启用医嘱时读取医嘱的暂停时间
            If mint类型 = 5 Or mint类型 = 6 Then
                strSQL = "Select B.医嘱ID,Max(B.操作时间) as 上次时间" & _
                    " From 病人医嘱记录 A,病人医嘱状态 B" & _
                    " Where A.ID=B.医嘱ID And B.操作类型=" & IIF(mint类型 = 5, 7, 6) & _
                    " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
                    " And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And A.病人来源<>3" & strWhere & _
                    " Group by B.医嘱ID"
                Set rsPause = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
            End If
            
            With vsAdvice
                .Redraw = flexRDNone
                Do While Not rsTmp.EOF
                    '添加新行
                    strTmp = ""
                    For i = 0 To rsTmp.Fields.Count - 1
                        strTmp = strTmp & vbTab & Nvl(rsTmp.Fields(i).Value)
                    Next
                    .AddItem Mid(strTmp, 2): i = .Rows - 1
                    
                    '是否显示婴儿列
                    If InStr(str婴儿 & ",", "," & .TextMatrix(i, COL_婴儿) & ",") = 0 Then
                        If str婴儿 <> "" Then .ColHidden(COL_婴儿) = False
                        str婴儿 = str婴儿 & "," & .TextMatrix(i, COL_婴儿)
                    End If
                    
                    '病人之间的间隔线
                    If .TextMatrix(i, COL_住院号) <> .TextMatrix(i - 1, COL_住院号) And i - 1 >= .FixedRows Then
                        .CellBorderRange i - 1, .FixedCols, i - 1, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                    End If
                    
                    '成药及中药的一些处理
                    bln给药途径 = False
                    If .TextMatrix(i, COL_诊疗类别) = "E" Then
                        If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                            If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                                bln给药途径 = True
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        '显示成药的给药途径
                                        .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法)
                                        '显示成药的执行性质
                                        If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                            .TextMatrix(j, COL_执行性质) = "自备药"
                                        ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                            .TextMatrix(j, COL_执行性质) = "离院带药"
                                        Else
                                            .TextMatrix(j, COL_执行性质) = ""
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                                If .TextMatrix(i - 1, COL_诊疗类别) = "7" Then
                                    .TextMatrix(i, COL_类型) = "1" '中药配方
                                ElseIf .TextMatrix(i - 1, COL_诊疗类别) = "C" Then
                                    .TextMatrix(i, COL_类型) = "2" '检验组合
                                End If
                                
                                '显示中药配方或检验组合的执行科室
                                .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)
                                
                                If .TextMatrix(i - 1, COL_诊疗类别) = "7" Then
                                    '显示中药配方执行性质
                                    If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        .TextMatrix(i, COL_执行性质) = "自备药"
                                    ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(i, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(i, COL_执行性质) = ""
                                    End If
                                Else
                                    .TextMatrix(i, COL_执行性质) = ""
                                End If
                                
                                '删除单味中药行,以及检验组合中的检验项目;同时判断检验申请
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        .RemoveItem j: i = .Rows - 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            .TextMatrix(i, COL_执行性质) = ""
                        End If
                    End If
                                                                    
                    '处理可见行的的一些标识
                    If Not bln给药途径 And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                        '处理小数点问题,暂未想到办法
                        If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                            .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                        End If
                        If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                            .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                        End If
                    
                        '时间以MM-DD HH:MI格式显示,以CellData进行判断
                        .Cell(flexcpData, i, COL_开始时间) = .TextMatrix(i, COL_开始时间)
                        .Cell(flexcpData, i, COL_开嘱时间) = .TextMatrix(i, COL_开嘱时间)
                        .Cell(flexcpData, i, COL_上次执行) = .TextMatrix(i, COL_上次执行)
                        .Cell(flexcpData, i, COL_终止时间) = .TextMatrix(i, COL_终止时间)
                        .Cell(flexcpData, i, COL_校对时间) = .TextMatrix(i, COL_校对时间)
                        .Cell(flexcpData, i, COL_停嘱时间) = .TextMatrix(i, COL_停嘱时间)
                        .TextMatrix(i, COL_开始时间) = Format(.TextMatrix(i, COL_开始时间), "MM-dd HH:mm")
                        .TextMatrix(i, COL_开嘱时间) = Format(.TextMatrix(i, COL_开嘱时间), "MM-dd HH:mm")
                        .TextMatrix(i, COL_上次执行) = Format(.TextMatrix(i, COL_上次执行), "MM-dd HH:mm")
                        .TextMatrix(i, COL_终止时间) = Format(.TextMatrix(i, COL_终止时间), "MM-dd HH:mm")
                        .TextMatrix(i, COL_校对时间) = Format(.TextMatrix(i, COL_校对时间), "MM-dd HH:mm")
                        .TextMatrix(i, COL_停嘱时间) = Format(.TextMatrix(i, COL_停嘱时间), "MM-dd HH:mm")
                        
                        If mint类型 = 1 Then
                            '停嘱时缺省的医嘱终止时间
                            If .Cell(flexcpData, i, COL_终止时间) = "" Then
                                '缺省执行终止时间
                                If gbln长期医嘱次日生效 Then
                                    .TextMatrix(i, COL_输入) = Format(vCurDate + 1, "yyyy-MM-dd 00:00")
                                Else
                                    .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                                '如果发送过,缺省在上次执行时间
                                If .TextMatrix(i, COL_上次执行) <> "" Then
                                    If .TextMatrix(i, COL_执行时间) = "" And Format(.TextMatrix(i, COL_上次执行), "HH:mm") = "00:00" Then
                                        '"持续性"的长嘱,停止当日不发送
                                        If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_上次执行)) + 1, "yyyy-MM-dd") Then
                                            .TextMatrix(i, COL_输入) = Format(CDate(.Cell(flexcpData, i, COL_上次执行)) + 1, "yyyy-MM-dd HH:mm")
                                        End If
                                    Else
                                        If .TextMatrix(i, COL_输入) < CStr(.Cell(flexcpData, i, COL_上次执行)) Then
                                            .TextMatrix(i, COL_输入) = CStr(.Cell(flexcpData, i, COL_上次执行))
                                        End If
                                    End If
                                End If
                            Else
                                .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_终止时间)
                            End If
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 2 Then
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            '应>=终止时间
                            If .TextMatrix(i, COL_输入) < .Cell(flexcpData, i, COL_终止时间) Then
                                .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_终止时间))), "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 3 Then
                            '校对时的缺省校对时间
                            If .TextMatrix(i, COL_标志) = "补录" Then
                                .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_开嘱时间))), "yyyy-MM-dd HH:mm")
                            Else
                                .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 5 Then
                            If mblnPauseLast Then
                                If .TextMatrix(i, COL_上次执行) <> "" Then
                                    '缺省在上次执行时间之后暂停
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_上次执行))), "yyyy-MM-dd HH:mm")
                                Else
                                    '如无上次执行时间则以开始时间为准
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_开始时间)
                                End If
                            Else
                                '暂停医嘱时间:暂停段中,医嘱暂停点无效,启用点有效。
                                .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            
                            '应>=开始执行时间,因为该时间点尚未执行
                            If .TextMatrix(i, COL_输入) < .Cell(flexcpData, i, COL_开始时间) Then
                                .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_开始时间)
                            End If
                            '应>上次执行时间,因为该时间点已执行
                            If .TextMatrix(i, COL_上次执行) <> "" Then
                                If .TextMatrix(i, COL_输入) <= .Cell(flexcpData, i, COL_上次执行) Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_上次执行))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            '应<执行终止时间,因为该时间点执行有效
                            If .TextMatrix(i, COL_终止时间) <> "" Then
                                If .TextMatrix(i, COL_输入) >= .Cell(flexcpData, i, COL_终止时间) Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", -1, CDate(.Cell(flexcpData, i, COL_终止时间))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                            rsPause.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_输入) <= Format(rsPause!上次时间, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, rsPause!上次时间), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        ElseIf mint类型 = 6 Then
                            '启用医嘱时间
                            .TextMatrix(i, COL_输入) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            
                            '应>暂停时间
                            rsPause.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_输入) <= Format(rsPause!上次时间, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_输入) = Format(DateAdd("n", 1, rsPause!上次时间), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            '应<=执行终止时间
                            If .TextMatrix(i, COL_终止时间) <> "" Then
                                If .TextMatrix(i, COL_输入) > .Cell(flexcpData, i, COL_终止时间) Then
                                    .TextMatrix(i, COL_输入) = .Cell(flexcpData, i, COL_终止时间)
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_输入) = .TextMatrix(i, COL_输入) '用于输入恢复
                        End If
                        
                        '行高
                        If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                        
                        '毒麻精药品标识
                        If .TextMatrix(i, COL_毒理分类) <> "" Then
                            If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                                .Cell(flexcpFontBold, i, COL_医嘱内容) = True
                            End If
                        End If
                        
                        '皮试结果标识
                        If .TextMatrix(i, COL_皮试) = "(+)" Then
                            .Cell(flexcpForeColor, i, COL_皮试) = vbRed
                        ElseIf .TextMatrix(i, COL_皮试) = "(-)" Then
                            .Cell(flexcpForeColor, i, COL_皮试) = vbBlue
                        End If
                        
                        'Pass:根据审查结果显示警示灯
                        If .TextMatrix(i, COL_警示) <> "" Then
                            Set .Cell(flexcpPicture, i, COL_警示) = imgPass.ListImages(Val(.TextMatrix(i, COL_警示)) + 1).Picture
                            .Cell(flexcpData, i, COL_警示) = .TextMatrix(i, COL_警示) '用于单药警告
                            .TextMatrix(i, COL_警示) = ""
                        End If
                        
                        '电子签名标识
                        If Val(.TextMatrix(i, COL_签名ID)) <> 0 Then
                            Set .Cell(flexcpPicture, i, COL_医嘱内容) = img16.ListImages("签名").Picture
                        End If
                    End If
                    
                    If bln给药途径 Then .RemoveItem i
                    
                    Progress = rsTmp.AbsolutePosition / rsTmp.RecordCount * 100
                    
                    rsTmp.MoveNext
                Loop
            End With
        End If
    Next
        
    '----------------------------------------------------------------------
    '病人信息显示
    If strPatis <> "" Then strPatis = Mid(strPatis, 2)
    If UBound(Split(strPatis, ";")) = 0 Then
        '只有一个病人的数据的情况
        lng病人ID = Split(strPatis, ",")(0)
        lng主页ID = Split(strPatis, ",")(1)
        If lng病人ID <> mlng病人ID Then '不是当前病人重新取来显示
            strSQL = _
                " Select A.住院号,A.姓名,A.性别,A.年龄,B.出院病床," & _
                " B.住院医师,B.出院科室ID,C.名称 as 科室" & _
                " From 病人信息 A,病案主页 B,部门表 C" & _
                " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                " And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
            lblPati.Caption = "姓名:" & rsTmp!姓名 & "　住院号:" & Nvl(rsTmp!住院号) & _
                "　床号:" & Nvl(rsTmp!出院病床) & "　科室:" & Nvl(rsTmp!科室)
        End If
    ElseIf UBound(Split(strPatis, ";")) > 0 Then
        '有多个病人数据的情况
        vsAdvice.ColHidden(COL_姓名) = False
        vsAdvice.ColHidden(COL_住院号) = False
        vsAdvice.ColHidden(COL_床号) = False
                
        strSQL = "Select 名称 From 部门表 Where ID IN(" & Mid(str科室s, 2) & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        str科室s = ""
        Do While Not rsTmp.EOF
            str科室s = str科室s & "," & rsTmp!名称
            rsTmp.MoveNext
        Loop
        lblPati.Caption = "共有(" & Mid(str科室s, 2) & ") " & UBound(Split(strPatis, ";")) + 1 & " 个病人的医嘱"
    ElseIf UBound(Split(strPatis, ";")) = -1 Then
        '没有任何病人数据的情况
        lblPati.Caption = ""
    End If
    
    '----------------------------------------------------------------------
    If vsAdvice.Rows = vsAdvice.FixedRows Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        vsPrice.Rows = vsPrice.FixedRows
        vsPrice.Rows = vsPrice.FixedRows + 1
    Else
        '电子签名图标对齐
        vsAdvice.Cell(flexcpPictureAlignment, vsAdvice.FixedRows, COL_医嘱内容, vsAdvice.Rows - 1, COL_医嘱内容) = 0
        '自动调整行高
        vsAdvice.AutoSize COL_医嘱内容
    End If
    Call SetAdviceCol
    vsAdvice.Row = vsAdvice.FixedRows
    If Not vsAdvice.ColHidden(COL_选择) Then
        vsAdvice.Col = COL_选择
    Else
        vsAdvice.Col = COL_医嘱内容
    End If
    vsAdvice.Redraw = flexRDDirect
    
    Progress = 0: Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    strPatis = ""
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Function CheckValid() As Boolean
'功能：确认前检查合法性
    Dim str超期 As String, str超长 As String
    Dim str特殊 As String, strTmp As String
    Dim curDate As Date, i As Long, k As Long
    Dim strPatis As String
    
    mstrRollNotify = ""
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '是否有可以操作的记录
        If .Rows = .FixedRows + 1 And Val(.TextMatrix(.FixedRows, COL_ID)) = 0 Then
            If mint类型 = 0 Then
                '医嘱作废
                strTmp = "当前没有可以作废的医嘱。"
            ElseIf mint类型 = 1 Then
                '停止医嘱
                strTmp = "当前没有可以停止的医嘱。"
            ElseIf mint类型 = 2 Then
                '确认停止
                strTmp = "当前没有被停止的医嘱。"
            ElseIf mint类型 = 3 Then
                '医嘱校对
                strTmp = "当前没有新开的医嘱。"
            ElseIf mint类型 = 4 Then
                '调整计价项目
                strTmp = "当前没有通过校对的有效医嘱。"
            ElseIf mint类型 = 5 Then
                '暂停医嘱
                strTmp = "当前没有可以暂停的医嘱。"
            ElseIf mint类型 = 6 Then
                '启用医嘱
                strTmp = "当前没有暂停后需要启用的医嘱。"
            End If
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            Exit Function
        End If
        
        '是否有选择
        str超期 = "": str超长 = "": str特殊 = ""
        If Not .ColHidden(COL_选择) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And (Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty) Then
                    k = k + 1
                    If InStr(strPatis & ",", "," & .TextMatrix(i, COL_病人ID)) = 0 Then
                        strPatis = strPatis & "," & .TextMatrix(i, COL_病人ID)
                    End If
                    
                    If mint类型 = 1 Then
                        '收集超期发送的医嘱
                        If IsDate(.Cell(flexcpData, i, COL_上次执行)) Then
                            If .TextMatrix(i, COL_执行时间) = "" And Format(.TextMatrix(i, COL_上次执行), "HH:mm") = "00:00" Then
                                '"持续性"长嘱,停止当天不发送
                                If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_上次执行)) + 1, "yyyy-MM-dd") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, COL_医嘱内容)
                                End If
                            Else
                                If Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, COL_医嘱内容)
                                End If
                            End If
                        End If
                        
                        '收集超长停止的医嘱
                        If CDate(.TextMatrix(i, COL_输入)) - curDate > 7 Then
                            str超长 = str超长 & vbCrLf & "●　" & .TextMatrix(i, COL_医嘱内容) & "，停止时间：" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm")
                        End If
                    ElseIf mint类型 = 2 Then
                        '收集超期发送的医嘱
                        If IsDate(.Cell(flexcpData, i, COL_上次执行)) Then
                            If .TextMatrix(i, COL_执行时间) = "" And Format(.TextMatrix(i, COL_上次执行), "HH:mm") = "00:00" Then
                                '"持续性"长嘱,停止当天不发送
                                If Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_上次执行)) + 1, "yyyy-MM-dd") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, COL_医嘱内容)
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then
                                    str超期 = str超期 & vbCrLf & "●　" & .TextMatrix(i, COL_医嘱内容)
                                End If
                            End If
                        End If
                    ElseIf mint类型 = 3 Then
                        '收集术后医嘱,通过校对的才判断
                        If .Cell(flexcpData, i, COL_选择) = 1 And _
                            .TextMatrix(i, COL_诊疗类别) = "Z" And .TextMatrix(i, COL_操作类型) = "4" Then
                            If InStr(str特殊 & ";", ";" & .TextMatrix(i, COL_病人ID) & "," & .TextMatrix(i, COL_主页ID) & ";") = 0 Then
                                str特殊 = str特殊 & ";" & .TextMatrix(i, COL_病人ID) & "," & .TextMatrix(i, COL_主页ID)
                            End If
                        End If
                    End If
                End If
            Next
            If k = 0 Then
                MsgBox "没有选择任何医嘱，请选择需要" & tbr.Buttons("执行").Caption & "的医嘱。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
                
        '医生
        If mint类型 = 1 And mbln护士站 Then
            If cbo医生.ListIndex = -1 Then
                MsgBox "请选择停止医嘱的医生。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    strTmp = ""
    strPatis = IIF(UBound(Split(Mid(strPatis, 2), ",")) > 0, "你选择了多个病人的医嘱，请仔细进行检查以避免出现差错。" & vbCrLf & vbCrLf, "")
    If mint类型 = 0 Then
        '医嘱作废
        strTmp = "确实要作废已经选择的医嘱吗？"
    ElseIf mint类型 = 1 Then
        '停止医嘱
        If str超期 <> "" Then '检查是否有需要退回超前摆药的情况
            If MsgBox("下列要停止的医嘱被超期发送：" & vbCrLf & str超期 & _
                vbCrLf & vbCrLf & "该类医嘱可以在护士工作站中使用""超期发送收回""进行处理。" & _
                vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str超长 <> "" Then
            If MsgBox("下列医嘱的停止时间超过当前时间太久：" & vbCrLf & str超长 & _
                vbCrLf & vbCrLf & "如果停止时间不正确，将会对医嘱的发送和计费产生影响。" & _
                vbCrLf & "确实要在指定的时间停止这些医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str超期 = "" And str超长 = "" Then
            strTmp = "确实要停止已经选择的医嘱吗？"
        End If
    ElseIf mint类型 = 2 Then
        '确认停止
        If str超期 <> "" Then
            If MsgBox("下列停止的医嘱被超期发送：" & vbCrLf & str超期 & _
                vbCrLf & vbCrLf & "该类医嘱可以在护士工作站中使用""超期发送收回""进行处理。" & _
                vbCrLf & "确认已经选择的医嘱停止吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            strTmp = "确认已经选择的医嘱停止吗？"
        End If
    ElseIf mint类型 = 3 Then
        '医嘱校对
        If str特殊 <> "" Then
            If MsgBox(strPatis & "要校对的医嘱中包括术后医嘱，校对后会停止其它长期医嘱。" & _
                vbCrLf & "确实要进行校对吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            mstrRollNotify = Mid(str特殊, 2)
        Else
            strTmp = strPatis & "确实要对已经选择的医嘱进行校对处理吗？"
        End If
    ElseIf mint类型 = 5 Then
        '暂停医嘱
        strTmp = strPatis & "确实要暂停已经选择的医嘱吗？"
    ElseIf mint类型 = 6 Then
        '启用医嘱
        strTmp = strPatis & "确实要启用已经选择的医嘱吗？"
    End If
    If strTmp <> "" Then
        If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    CheckValid = True
End Function

Private Function CheckSignValid() As Boolean
'功能：1.检查未签名的医嘱不能进行校对
'      2.一次签名的医嘱必须一起通过校对
    Dim col医嘱ID As New Collection, str医嘱ID As String
    Dim col签名ID As New Collection, str签名ID As String
    Dim str住院 As String, str医技 As String
    Dim lng签名ID As Long, strTmp As String
    Dim int状态 As Integer, i As Long, j As Long
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    If mint类型 <> 3 Then CheckSignValid = True: Exit Function
    
    With vsAdvice
        '获取护士人员列表：只是护士，不是医生
        If Mid(gstrESign, 2, 1) = "1" Or Mid(gstrESign, 3, 1) = "1" Then
            strNurse = ""
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                    If .Cell(flexcpData, i, COL_选择) = 1 And Val(.TextMatrix(i, COL_签名ID)) = 0 Then
                        If InStr(strNurse & ",", "," & .TextMatrix(i, COL_操作人员) & ",") = 0 Then
                            strNurse = strNurse & "," & .TextMatrix(i, COL_操作人员)
                        End If
                    End If
                End If
            Next
            If strNurse <> "" Then
                strSQL = "Select /*+ Rule*/ A.姓名" & _
                    " From 人员表 A,(Select * From Table(Cast(f_Str2List([1]) As zlTools.t_StrList))) B" & _
                    " Where A.姓名=B.Column_Value" & _
                    " And Exists(Select 1 From 人员性质说明 X Where X.人员ID=A.ID And X.人员性质='护士')" & _
                    " And Not Exists(Select 1 From 人员性质说明 Y Where Y.人员ID=A.ID And Y.人员性质='医生')"
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strNurse, 2))
                On Error GoTo 0
                
                strNurse = ""
                Do While Not rsTmp.EOF
                    strNurse = strNurse & "," & rsTmp!姓名
                    rsTmp.MoveNext
                Loop
                strNurse = strNurse & ","
            End If
        End If
        
        For i = .FixedRows To .Rows - 1
            'flexcpData:0-不处理,1-校对,2-疑问
            If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                '1.收集未签名的医嘱内容
                If .Cell(flexcpData, i, COL_选择) = 1 And Val(.TextMatrix(i, COL_签名ID)) = 0 Then
                    '设置为使用签名的场合
                    If InStr(strNurse, "," & .TextMatrix(i, COL_操作人员) & ",") = 0 Then '护士录入的医嘱不进行签名检查
                        If Val(.TextMatrix(i, COL_前提ID)) = 0 And Mid(gstrESign, 2, 1) = "1" Then
                            If UBound(Split(str住院, vbCrLf)) < 10 Then
                                str住院 = str住院 & vbCrLf & "●" & .TextMatrix(i, COL_医嘱内容)
                            ElseIf InStr(str住院, "… …") = 0 Then
                                str住院 = str住院 & vbCrLf & "… …"
                            End If
                        ElseIf Val(.TextMatrix(i, COL_前提ID)) <> 0 And Mid(gstrESign, 3, 1) = "1" Then
                            If UBound(Split(str医技, vbCrLf)) < 10 Then
                                str医技 = str医技 & vbCrLf & "●" & .TextMatrix(i, COL_医嘱内容)
                            ElseIf InStr(str医技, "… …") = 0 Then
                                str医技 = str医技 & vbCrLf & "… …"
                            End If
                        End If
                    End If
                End If
                
                '2.收集已签名医嘱的校对状态
                lng签名ID = Val(.TextMatrix(i, COL_签名ID))
                If lng签名ID <> 0 Then
                    j = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID))) '组ID
                    int状态 = .Cell(flexcpData, i, COL_选择)
                    If int状态 = 2 Then int状态 = 0 '这里疑问等同于不校对
                    If InStr(str签名ID & ",", "," & lng签名ID & ",") > 0 Then
                        '收集各个签名在界面上的校对状态
                        strTmp = Split(col签名ID("_" & lng签名ID), "=")(1)
                        If InStr(strTmp, int状态) = 0 Then
                            col签名ID.Remove "_" & lng签名ID
                            col签名ID.Add lng签名ID & "=" & strTmp & int状态, "_" & lng签名ID
                        End If
                        
                        '收集各个签名已读到界面的医嘱(组ID)
                        strTmp = col医嘱ID("_" & lng签名ID)
                        If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                            col医嘱ID.Remove "_" & lng签名ID
                            col医嘱ID.Add strTmp & "," & j, "_" & lng签名ID
                        End If
                    Else
                        str签名ID = str签名ID & "," & lng签名ID
                        col签名ID.Add lng签名ID & "=" & int状态, "_" & lng签名ID
                        col医嘱ID.Add j, "_" & lng签名ID
                    End If
                End If
            End If
        Next
        
        '检查已签名医嘱校对情况
        strTmp = "": str医嘱ID = Mid(str医嘱ID, 2)
        For i = 1 To col签名ID.Count
            lng签名ID = Split(col签名ID(i), "=")(0)
            str签名ID = Split(col签名ID(i), "=")(1)
            
            '本次一起签名的未读入界面的未校对医嘱
            str医嘱ID = col医嘱ID("_" & lng签名ID)
            str医嘱ID = ExistOtherSignAdvice(lng签名ID, str医嘱ID)
            If str医嘱ID <> "" Then
                If InStr(str签名ID, "0") = 0 Then
                    str签名ID = str签名ID & "0"
                    strTmp = strTmp & str医嘱ID
                End If
            End If
            
            If Not (str签名ID = "1" Or str签名ID = "0") Then
                '这次签名的内容不是"都要通过校对或都不通过校对(包括疑问)"的情况
                j = .FindRow(CStr(lng签名ID), , COL_签名ID)
                Do While j <> -1
                    If Val(.TextMatrix(j, COL_ID)) <> 0 And Not .RowHidden(j) Then
                        If InStr(",0,2,", .Cell(flexcpData, j, COL_选择)) > 0 Then
                            strTmp = strTmp & vbCrLf & .TextMatrix(j, COL_姓名) & "：" & IIF(Len(.TextMatrix(j, COL_医嘱内容)) > 40, Left(.TextMatrix(j, COL_医嘱内容), 40) & "...", .TextMatrix(j, COL_医嘱内容))
                        End If
                    End If
                    j = .FindRow(CStr(lng签名ID), j + 1, COL_签名ID)
                Loop
                Exit For '暂只提示第一组
            End If
        Next
    End With
    
    '1.没有签名的医嘱不允许校对：对住院医嘱和医技医嘱分别进行检查
    If str住院 <> "" Then
        MsgBox "以下医嘱医生还没有签名，不能进行校对：" & vbCrLf & str住院, vbInformation, gstrSysName
        Exit Function
    End If
    If str医技 <> "" Then
        MsgBox "以下医嘱医生还没有签名，不能进行校对：" & vbCrLf & str医技, vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.一起签名的医嘱必须一起通过校对
    If strTmp <> "" Then
        MsgBox "以下医嘱与其他本次要通过校对的医嘱一起签名，但当前处理为不校对或校对疑问：" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "一起签名的医嘱必须一起通过校对，请调整相关医嘱的校对状态。", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckSignValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExistOtherSignAdvice(ByVal lng签名ID As Long, ByVal str医嘱ID As String) As String
'功能：检查是否存在某次新开医嘱签名中本次没有读取到界面上的医嘱(因为要一起通过校对,如果有,这些医嘱也是没校对的)
'返回：未读取到界面的未校对医嘱内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.姓名,B.医嘱内容 From 病人医嘱状态 A,病人医嘱记录 B" & _
        " Where A.医嘱ID=B.ID And A.操作类型=1 And B.医嘱状态 IN(1,2)" & _
        " And (B.相关ID is Null Or B.诊疗类别 IN('5','6'))" & _
        " And Not Exists(Select 1 From 病人医嘱记录 S Where 诊疗类别 IN('5','6') And S.相关ID=B.ID)" & _
        " And Instr([2],','||Nvl(B.相关ID,B.ID)||',')=0 And A.签名ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng签名ID, "," & str医嘱ID & ",")
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & vbCrLf & Nvl(rsTmp!姓名) & "：" & IIF(Len(Nvl(rsTmp!医嘱内容)) > 40, Left(Nvl(rsTmp!医嘱内容), 40) & "...", Nvl(rsTmp!医嘱内容))
        rsTmp.MoveNext
    Loop
    ExistOtherSignAdvice = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng医嘱ID As Long, ByVal lng项目ID As Long, ByVal lngCol As Long)
'功能：定位到并显示指定医嘱的指定计价行
'参数：lngRow=医嘱行号,lng医嘱ID=计价医嘱ID
'      lng项目ID=计价项目ID,lngCol=计价表格显示列
    Dim k As Long
    
    With vsAdvice
        .Row = lngRow: .Col = COL_医嘱内容 '进入行自动ShowPrice,mrsPrice发生变化
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_医嘱ID)) = lng医嘱ID _
                And Val(vsPrice.TextMatrix(k, COLP_收费细目ID)) = lng项目ID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function ExecuteOperate() As Boolean
    Dim arrSQL As Variant, lng相关ID As Long
    Dim blnExe As Boolean, i As Long, j As Long
    Dim lng医嘱ID As Long, lng执行科室ID As Long
    Dim strOper As String, blnVarZero As Boolean
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名ID As Long, lng证书ID As Long
    Dim strSource As String, strSign As String
    Dim colStopTime As New Collection
    
    Screen.MousePointer = 11
    
    '产生SQL
    arrSQL = Array()
    With vsAdvice
        If mint类型 <> 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) <> Empty Then
                    '一组医嘱只校对一次,除一并给药外,其它医嘱只有一个显示行
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> lng相关ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    If blnExe Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        '(组ID)使用相关ID为NULL的医嘱的ID(给药途径,中药用法,检查项目,主要手术,及独立医嘱)
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            lng医嘱ID = Val(.TextMatrix(i, COL_相关ID))
                        Else
                            lng医嘱ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If mint类型 = 0 Then      '医嘱作废
                            '医生作废医嘱电子签名
                            If Val(.TextMatrix(i, COL_签名ID)) <> 0 Then
                                str医嘱ID = str医嘱ID & "," & lng医嘱ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ")"
                        ElseIf mint类型 = 1 Then  '停止医嘱
                            '医生停止医嘱电子签名
                            If Val(.TextMatrix(i, COL_签名ID)) <> 0 Then
                                str医嘱ID = str医嘱ID & "," & lng医嘱ID
                                '记录停止医嘱的执行终止时间：由于是在执行过程之前取签名源文,这时还未写入数据库
                                colStopTime.Add Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_停止(" & lng医嘱ID & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                "'" & IIF(mbln护士站, NeedName(cbo医生.Text), UserInfo.姓名) & "')"
                        ElseIf mint类型 = 2 Then  '确认停止
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_确认停止(" & lng医嘱ID & "," & _
                            "To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint类型 = 3 Then  '医嘱校对
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_校对(" & lng医嘱ID & "," & _
                                IIF(.Cell(flexcpData, i, COL_选择) = 1, 3, 2) & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint类型 = 5 Then  '暂停医嘱
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_暂停(" & lng医嘱ID & ",To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint类型 = 6 Then  '启用医嘱
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_启用(" & lng医嘱ID & ",To_Date('" & Format(.TextMatrix(i, COL_输入), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        End If
                    End If
                End If
                lng相关ID = Val(.TextMatrix(i, COL_相关ID))
            Next
        End If
        
        '医嘱计价部分
        lng相关ID = 0
        If mint类型 = 3 Or mint类型 = 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_选择)) <> 0 Or .Cell(flexcpData, i, COL_选择) = 1 Then
                    '一并给药的只需处理一次
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> lng相关ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    
                    If blnExe Then
                        '删除对应的计价
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_相关ID)) & ")"
                        Else
                            arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        End If
                        
                        '生成新的计价
                        '本来用一次性循环快些,但为了判断是否要保存及输入合法性,必须用Filter
                        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) <> 0 Then
                            mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                " Or 医嘱ID=" & Val(vsAdvice.TextMatrix(i, COL_相关ID))
                        Else
                            mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                " Or 相关ID=" & vsAdvice.TextMatrix(i, COL_ID)
                        End If
                        For j = 1 To mrsPrice.RecordCount
                            '之中存在收费细目ID为空的无用记录(最初用于确定可选的计价医嘱)
                            If Not IsNull(mrsPrice!收费细目ID) And InStr(",5,6,7,", mrsPrice!诊疗类别) = 0 Then
                                If Nvl(mrsPrice!数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                    blnVarZero = False
                                    If Nvl(mrsPrice!单价, 0) = 0 Then
                                        blnVarZero = ItemIsVarPrice(mrsPrice!收费细目ID)
                                    End If
                                    If blnVarZero Then
                                        Call SeekPriceRow(i, mrsPrice!医嘱ID, mrsPrice!收费细目ID, COLP_单价)
                                        Screen.MousePointer = 0
                                        MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, gstrSysName
                                        vsPrice.SetFocus: Exit Function
                                    End If
                                    
                                    '计价执行科室:只保存非药嘱药品及卫材计价
                                    If InStr(",5,6,7,", mrsPrice!收费类别) > 0 _
                                        Or mrsPrice!收费类别 = "4" And Nvl(mrsPrice!在用, 0) = 1 Then
                                        lng执行科室ID = Nvl(mrsPrice!执行科室ID, 0)
                                        
                                        '卫材必须设置执行科室
                                        If lng执行科室ID = 0 And mrsPrice!收费类别 = "4" Then
                                            Call SeekPriceRow(i, mrsPrice!医嘱ID, mrsPrice!收费细目ID, COLP_执行科室)
                                            Screen.MousePointer = 0
                                            MsgBox "卫材""" & vsPrice.TextMatrix(vsPrice.Row, COLP_收费项目) & """没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                                "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                            vsPrice.SetFocus: Exit Function
                                        End If
                                    Else
                                        lng执行科室ID = 0
                                    End If
                                    
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "zl_病人医嘱计价_Insert(" & mrsPrice!医嘱ID & "," & _
                                        mrsPrice!收费细目ID & "," & mrsPrice!数量 & "," & Nvl(mrsPrice!单价, 0) & "," & _
                                        Nvl(mrsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & ")"
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                End If
                lng相关ID = Val(.TextMatrix(i, COL_相关ID))
            Next
        End If
    End With
    
    '作废或停止时的电子签名
    If (mint类型 = 0 Or mint类型 = 1) And str医嘱ID <> "" Then
        strOper = Decode(mint类型, 0, "作废", 1, "停止")
        
        '护士不能作废、停止医生已签名的医嘱
        If mbln护士站 Then
            MsgBox "你要" & strOper & "的医嘱中包含医生已签名的医嘱，只能由医生来" & strOper & "并签名。", vbInformation, gstrSysName
            Screen.MousePointer = 0: Exit Function
        End If
        
        '医生停止,作废时必须要签名
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox strOper & "已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能" & strOper & "。", vbInformation, gstrSysName
            Else
                MsgBox strOper & "已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能" & strOper & "。", vbInformation, gstrSysName
            End If
            Screen.MousePointer = 0: Exit Function
        End If
        
        '获取签名医嘱源文
        str医嘱ID = Mid(str医嘱ID, 2) '组ID,返回为明细ID
        intRule = ReadAdviceSignSource(Decode(mint类型, 0, 4, 1, 8), mlng病人ID, mlng主页ID, str医嘱ID, 0, False, strSource, , colStopTime)
        If intRule = 0 Then Screen.MousePointer = 0: Exit Function
        If strSource = "" Then
            Screen.MousePointer = 0
            MsgBox "不能读取需要" & strOper & "的已签名医嘱源文内容。", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
        If strSign <> "" Then
            lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_医嘱签名记录_Insert(" & lng签名ID & "," & Decode(mint类型, 0, 4, 1, 8) & "," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱ID & "')"
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    
    '执行SQL
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    Screen.MousePointer = 0
    ExecuteOperate = True
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitPriceRecordset()
'说明：编辑时,当计价医嘱及收费项目都输入后,才加入记录集
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "诊疗类别", adVarChar, 1
    mrsPrice.Fields.Append "诊疗项目ID", adBigInt
    mrsPrice.Fields.Append "收费类别", adVarChar, 1, adFldIsNullable
    mrsPrice.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "单价", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "在用", adInteger '卫材是否跟踪在用
    mrsPrice.Fields.Append "从项", adInteger
    mrsPrice.Fields.Append "固定", adInteger '现有的收费关系中是否固定对照
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub ShowDefaultRow()
'功能：对于可以计价的医嘱,缺省增加一行并设置缺省计价医嘱
'说明：ComboList="#医嘱ID1;计价医嘱1|#医嘱ID2;计价医嘱2|..."
'      仅在第一次显示计价表和回车新增行时调用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrCombo As Variant, lngRow As Long
    Dim lng医嘱ID As Long, str计价医嘱 As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If .ColData(COLP_计价医嘱) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_计价医嘱), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_医嘱ID)) <> 0 _
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
                    And Val(.TextMatrix(lngRow - 1, COLP_医嘱ID)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lng医嘱ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str计价医嘱 = Replace(arrCombo(i), "#" & lng医嘱ID & ";", "")
                If blnHave Then
                    If lng医嘱ID = Val(.TextMatrix(lngRow - 1, COLP_医嘱ID)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            '模拟选中这个计价医嘱
            strSQL = "Select 相关ID,诊疗类别,诊疗项目ID From 病人医嘱记录 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COLP_医嘱ID) = lng医嘱ID
                .TextMatrix(lngRow, COLP_计价医嘱) = str计价医嘱
                .TextMatrix(lngRow, COLP_相关ID) = Nvl(rsTmp!相关ID)
                .TextMatrix(lngRow, COLP_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(lngRow, COLP_诊疗类别) = rsTmp!诊疗类别
                .Cell(flexcpData, lngRow, COLP_计价医嘱) = .TextMatrix(lngRow, COLP_计价医嘱)
                
                '只有一个计价医嘱时不必停留
                If UBound(arrCombo) = 0 Then
                    .Col = COLP_收费项目
                Else
                    .Col = COLP_计价医嘱
                End If
            End If
        End If
        Call .ShowCell(.Row, .Col)
        If blnFirst Then .TopRow = .Row '第一次显示时,ShowCell居然不起作用
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset, strSQL As String, i As Long
    Dim lng原嘱ID As Long, lng医嘱ID As Long, lng收费细目ID As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_计价医嘱 Then
            '如果绑定了ComboData,TextMatrix取值就为ComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lng医嘱ID = .ComboData
                lng原嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
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
                                                                
                '表格内容：mrsPrice中可能已删除,所以要从数据库读
                strSQL = "Select 相关ID,诊疗类别,诊疗项目ID From 病人医嘱记录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If rsTmp.EOF Then
                    MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """可能已经被其它人删除,请退出重新进入。", vbInformation, gstrSysName
                    Exit Sub
                End If
                .TextMatrix(Row, COLP_医嘱ID) = lng医嘱ID
                .TextMatrix(Row, COLP_相关ID) = Nvl(rsTmp!相关ID)
                .TextMatrix(Row, COLP_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(Row, COLP_诊疗类别) = rsTmp!诊疗类别
                
                '记录集内容
                If lng收费细目ID <> 0 Then
                    '新选择的医嘱是否有从项决定修改后的项目是否从项
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 从项=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_从项) = IIF(blnHaveSub, "√", "")
                    
                    If lng原嘱ID = 0 Then
                        mrsPrice.AddNew '加入
                    Else '更新
                        mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 收费细目ID=" & lng收费细目ID
                    End If
                    mrsPrice!医嘱ID = lng医嘱ID
                    mrsPrice!相关ID = rsTmp!相关ID
                    mrsPrice!诊疗项目ID = rsTmp!诊疗项目ID
                    mrsPrice!诊疗类别 = rsTmp!诊疗类别
                    If lng原嘱ID = 0 Then
                        mrsPrice!收费细目ID = lng收费细目ID
                        mrsPrice!数量 = Val(.TextMatrix(Row, COLP_数量))
                        mrsPrice!单价 = Val(.TextMatrix(Row, COLP_单价))
                        mrsPrice!固定 = 0
                    End If
                    mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            End If
        ElseIf Col = COLP_收费项目 Or Col = COLP_执行科室 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        ElseIf Col = COLP_数量 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        ElseIf Col = COLP_单价 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), "0.00000")
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str项目IDs As String, blnCancel As Boolean
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim vPoint As POINTAPI
    
    With vsPrice
        If Col = COLP_收费项目 Then
            '不能选择已有的项目
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_医嘱ID)) = Val(.TextMatrix(Row, COLP_医嘱ID)) _
                    And Val(.TextMatrix(Row, COLP_医嘱ID)) <> 0 And i <> Row Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                End If
            Next
            str项目IDs = Mid(str项目IDs, 2)
            
            strSQL = _
                " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
                " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7)"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,-ID as ID,Nvl(-上级ID,To_Number('999999999'||类型)) as 上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,ID,上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 是否变价ID,Null as 类别ID,-NULL as 跟踪在用ID" & _
                " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 末级,ID,上级ID,编码,名称,单位,规格,产地,类别,费用类型,说明," & _
                " Decode(Nvl(是否变价,0),1,Decode(Instr('567',类别ID),0,Sum(原价)||'-'||Sum(现价),'时价'),Sum(现价)) as 价格," & _
                " Sum(原价) as 原价ID,Sum(现价) as 现价ID,是否变价 as 是否变价ID,类别ID,跟踪在用ID" & _
                " From (" & _
                " Select Distinct 1 as 末级,A.ID,Decode(Instr('567',A.类别),0,A.分类ID,-E.分类ID) as 上级ID,A.编码,A.名称," & _
                " A.计算单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,C.名称 as 类别,A.费用类型,A.说明,B.原价,B.现价,A.是否变价," & _
                " -NULL as 跟踪在用ID" & _
                " From 收费项目目录 A,收费价目 B,收费项目类别 C,药品规格 D,诊疗项目目录 E" & _
                " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.类别 Not IN('4','J','1') And A.类别=C.编码 And A.ID=D.药品ID(+) And D.药名ID=E.ID(+)"
            If DeptExist("发料部门", 2) Then
                strSQL = strSQL & " Union ALL" & _
                    " Select Distinct 1 as 末级,A.ID,-E.分类ID as 上级ID,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,C.名称 as 类别,A.费用类型,A.说明," & _
                    " B.原价,B.现价,A.是否变价,D.跟踪在用 as 跟踪在用ID" & _
                    " From 收费项目目录 A,收费价目 B,收费项目类别 C,材料特性 D,诊疗项目目录 E" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(2,3)" & IIF(str项目IDs <> "", " And A.ID Not IN(" & str项目IDs & ")", "") & _
                    " And A.类别='4' And A.类别=C.编码 And A.ID=D.材料ID And D.诊疗ID=E.ID"
            End If
            strSQL = strSQL & " ) Group by 末级,ID,上级ID,编码,名称,单位,规格,产地,类别,费用类型,说明,是否变价,类别ID,跟踪在用ID"
            
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "收费项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str项目IDs & ",")
            If Not rsTmp Is Nothing Then
                '医保对码检查
                If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID))) Then
                    .SetFocus: Exit Sub
                End If
                
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                Call SetItemInput(Row, rsTmp, lng医嘱ID, lng原项目ID)
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
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!执行科室ID = rsTmp!ID
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
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

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查输入(选择)计价项目是否医保对码
'返回：如果未对码，并且提示选择不继续，则返回真。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int险类 As Integer
    
    If gint医保对码 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select 险类 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", lng病人ID, lng主页ID)
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

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lng医嘱ID As Long, ByVal lng原项目ID As Long)
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim lng行号 As Long, dbl单价 As Double
    Dim blnHaveSub As Boolean, dbl总量 As Double
    Dim rsTmp As ADODB.Recordset
    
    With vsPrice
        '表格内容
        .TextMatrix(lngRow, COLP_收费类别) = rsInput!类别ID
        .TextMatrix(lngRow, COLP_收费细目ID) = rsInput!ID
        .TextMatrix(lngRow, COLP_类别) = rsInput!类别
        .TextMatrix(lngRow, COLP_收费项目) = rsInput!名称
        If Not IsNull(rsInput!产地) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & "(" & rsInput!产地 & ")"
        End If
        If Not IsNull(rsInput!规格) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & " " & rsInput!规格
        End If
        
        '如果是加入药品计价(非药嘱),按零售单位处理
        .TextMatrix(lngRow, COLP_数量) = 1 '缺省计价数量为1
        .TextMatrix(lngRow, COLP_单位) = Nvl(rsInput!单位)
                
        '单价计算处理:药嘱计价不可能在这里处理,非药嘱药品计价按售价处理
        .Cell(flexcpData, lngRow, 0) = 0
        .Cell(flexcpData, lngRow, 1) = 0
        .Cell(flexcpData, lngRow, 2) = 0
        
        '执行科室
        lng行号 = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If lng行号 = -1 Then
            Set rsTmp = GetItemField("病人医嘱记录", lng医嘱ID)
            lng病人ID = rsTmp!病人ID
            lng主页ID = Nvl(rsTmp!主页ID, 0)
            lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
            lng病人科室ID = Nvl(rsTmp!病人科室ID, 0)
            dbl总量 = Nvl(rsTmp!总给予量, 0)
            If dbl总量 = 0 Then dbl总量 = 1
        Else
            lng病人ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人ID))
            lng主页ID = Val(vsAdvice.TextMatrix(lng行号, COL_主页ID))
            lng执行科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))
            lng病人科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID))
            dbl总量 = Val(vsAdvice.TextMatrix(lng行号, COL_总量))
            If dbl总量 = 0 Then dbl总量 = 1
        End If
            
        '非药嘱和跟踪在用的卫材专门求执行科室
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Or rsInput!类别ID = "4" And Nvl(rsInput!跟踪在用ID, 0) = 1 Then
            lng执行科室ID = Get收费执行科室ID(lng病人ID, lng主页ID, rsInput!类别ID, rsInput!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
            '记录卫材是否跟踪在用
            If rsInput!类别ID = "4" Then
                .TextMatrix(lngRow, COLP_跟踪在用) = Nvl(rsInput!跟踪在用ID, 0)
            End If
        End If
        If lng执行科室ID <> 0 Then
            mrsDept.Filter = "ID=" & lng执行科室ID
            If Not mrsDept.EOF Then
                .TextMatrix(lngRow, COLP_执行科室) = mrsDept!名称
            End If
        End If
        .TextMatrix(lngRow, COLP_执行科室ID) = lng执行科室ID
                
        '单价
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            If Nvl(rsInput!是否变价ID, 0) = 0 Then
                dbl单价 = Nvl(rsInput!现价ID, 0)
            Else '未确定计价医嘱时,药品无法计算价格
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, dbl总量, , True) '按缺省计价数量为1个零售单位计算
            End If
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, "0.00000")
        ElseIf rsInput!类别ID = "4" And Nvl(rsInput!跟踪在用ID, 0) = 1 And Nvl(rsInput!是否变价ID, 0) = 1 Then
            '跟踪在用的时价卫材和药品一样计算
            dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, dbl总量, , True)
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, "0.00000")
        Else
            If Nvl(rsInput!是否变价ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_单价) = Format(Nvl(rsInput!现价ID, 0), "0.00000")
            Else
                .Cell(flexcpData, lngRow, 0) = 1
                .Cell(flexcpData, lngRow, 1) = Nvl(rsInput!原价ID, 0)
                .Cell(flexcpData, lngRow, 2) = Nvl(rsInput!现价ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_费用类型) = Nvl(rsInput!费用类型)
        .TextMatrix(lngRow, COLP_固定) = "0"
        
        '用于输入恢复
        .Cell(flexcpData, lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目)
        .Cell(flexcpData, lngRow, COLP_数量) = .TextMatrix(lngRow, COLP_数量)
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
                mrsPrice!相关ID = Val(.TextMatrix(lngRow, COLP_相关ID))
                mrsPrice!诊疗类别 = .TextMatrix(lngRow, COLP_诊疗类别)
                mrsPrice!诊疗项目ID = Val(.TextMatrix(lngRow, COLP_诊疗项目ID))
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
            Call SelectRow(vsAdvice.Row)
        End If
    End With
End Sub

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
                If Val(.TextMatrix(.Row, COLP_医嘱ID)) <> 0 And Val(.TextMatrix(.Row, COLP_收费细目ID)) <> 0 Then
                    '医嘱如果有从项至少要保留一个(主项是固定不可动的)
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(.Row, COLP_医嘱ID)) & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_计价医嘱) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("确定要删除当前计价行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(.Row, COLP_医嘱ID)) & " And 收费细目ID=" & Val(.TextMatrix(.Row, COLP_收费细目ID))
                    mrsPrice.Delete
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_计价医嘱
                End If
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

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str项目IDs As String
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim StrInput As String, strMatch As String
    Dim vPoint As POINTAPI
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COLP_计价医嘱 Then
                '下拉时回车
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '不然EnterNextCell函数要退出
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_数量 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费数量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_单价 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费单价输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '检查变价输入范围
                strTmp = CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, "0.00000")
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_收费项目 And .EditText <> "" Then
                '不能选择已有的项目
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COLP_医嘱ID)) = Val(.TextMatrix(Row, COLP_医嘱ID)) _
                        And Val(.TextMatrix(Row, COLP_医嘱ID)) <> 0 And i <> Row Then
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
                    " Select Distinct 1 as 末级,A.ID,A.类别 as 类别ID,D.名称 as 类别,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,A.费用类型,A.说明,B.原价,B.现价,A.是否变价" & _
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
                    '医保对码检查
                    If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_主页ID))) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                        .SetFocus: Exit Sub
                    End If
                    
                    lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    Call SetItemInput(Row, rsTmp, lng医嘱ID, lng原项目ID)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
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
                    
                    '更新记录集
                    lng医嘱ID = Val(.TextMatrix(Row, COLP_医嘱ID))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                        mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 收费细目ID=" & lng原项目ID
                        mrsPrice!执行科室ID = rsTmp!ID
                        mrsPrice.Update
                        Call SelectRow(vsAdvice.Row)
                    End If
                    
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
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
            If Col = COLP_数量 Or Col = COLP_单价 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
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
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = COLP_数量 Or Col = COLP_单价 Or Col = COLP_执行科室 Then
        If vsPrice.TextMatrix(Row, COLP_收费项目) = "" Then
            Cancel = True '必须先确定收费项目
        End If
    End If
    
    If Col = COLP_数量 Or Col = COLP_单价 Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断价表中单元格是否可以编辑
    CellEditable = vsPrice.Editable
    With vsPrice
        If lngCol = COLP_执行科室 Then
            '跟踪在用的卫材,非药嘱药品计价的执行科室可以修改
            If Not (.TextMatrix(lngRow, COLP_收费类别) = "4" And Val(.TextMatrix(lngRow, COLP_跟踪在用)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_收费类别)) > 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COLP_诊疗类别)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_收费项目) = "" Or .TextMatrix(lngRow, COLP_诊疗类别) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_固定)) <> 0 Then
            '固定对照行仅可以修改变价
            If Not (.Cell(flexcpData, lngRow, 0) = 1 And lngCol = COLP_单价) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_单价 Then
                If .Cell(flexcpData, lngRow, 0) <> 1 Then CellEditable = False
            ElseIf lngCol <> COLP_计价医嘱 And lngCol <> COLP_数量 And lngCol <> COLP_收费项目 Then
                CellEditable = False
            End If
        End If
    End With
End Function

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
                ElseIf .TextMatrix(lngRow, COLP_数量) = "" Then
                    .Col = COLP_数量
                ElseIf .TextMatrix(lngRow, COLP_收费项目) = "" Then
                    .Col = COLP_收费项目
                ElseIf .Cell(flexcpData, lngRow, 0) = 1 And Val(.TextMatrix(lngRow, COLP_单价)) = 0 Then
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

Private Function LoadPrice(ByVal lngRow As Long, Optional blnChange As Boolean) As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费关系进行更新
'返回：blnChange=是否根据当前的诊疗收费关系对现有的计价内容进行了调整
    Dim rsMan As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim rsAdd As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim blnLoad As Boolean, lng诊疗项目ID As Long
    Dim dblPrice As Double, blnSubItem As Boolean
    Dim lng执行科室ID As Long
    
    On Error GoTo errH
    
    With vsAdvice
        '已经读取过了,不再重复读取
        If .TextMatrix(lngRow, COL_ID) = "" Then LoadPrice = True: Exit Function
        If .RowData(lngRow) = 1 Then LoadPrice = True: Exit Function
                            
        '药品的计价(这里仅用于显示；数量为相对数量,药品固定为1；实价药品单价显示时计算)
        '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算1个住院包装的单价
            strSQL = _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,C.类别 as 收费类别,C.ID as 收费细目ID," & _
                " 1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价,0 as 从项," & _
                " A.执行科室ID,0 as 跟踪在用,C.撤档时间" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(2,3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        ElseIf .TextMatrix(lngRow, COL_类型) = "1" Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,C.类别 as 收费类别,C.ID as 收费细目ID," & _
                " 1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价,0 as 从项," & _
                " A.执行科室ID,0 as 跟踪在用,C.撤档时间" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN(2,3)" & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价：除药品外的计价,包含相关医嘱计价
        blnLoad = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '给药途径:一并给药的只读取一次来共用
            If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                If .TextMatrix(lngRow - 1, COL_相关ID) = .TextMatrix(lngRow, COL_相关ID) Then
                    blnLoad = False
                End If
            End If
        End If
        If blnLoad Then
            '成药的给药途径；中药配方的煎法，用法；检查及部位；手术及附加手术,麻醉项目
            '不计价,手工计价；叮嘱,院外执行；的医嘱不读取
            '用Union方式可以利用索引
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,B.类别 as 收费类别,A.收费细目ID," & _
                "   A.数量,A.单价,Nvl(A.从项,0) as 从项,A.执行科室ID,C.跟踪在用,B.撤档时间" & _
                " From (" & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID" & _
                " From 病人医嘱记录 A,病人医嘱计价 B" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.ID=[1]" & _
                " Union ALL" & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID" & _
                " From 病人医嘱记录 A,病人医嘱计价 B" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.ID=[2]" & _
                " Union ALL" & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,B.收费细目ID,B.数量,B.单价,B.从项,Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID" & _
                " From 病人医嘱记录 A,病人医嘱计价 B" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.ID=B.医嘱ID(+) And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And A.相关ID=[1]" & _
                " ) A,收费项目目录 B,材料特性 C" & _
                " Where A.收费细目ID=B.ID(+) And A.收费细目ID=C.材料ID(+)" & _
                " Order by 序号,从项"
        End If
        Set rsMan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)))
        
        '诊疗收费关系中收费数量及固有对照是否变化
        strSQL = "Select C.诊疗项目ID,C.收费项目ID,C.收费数量,C.固有对照,C.从属项目" & _
            " From 病人医嘱记录 A,病人医嘱计价 B,诊疗收费关系 C" & _
            " Where A.ID=B.医嘱ID And A.诊疗项目ID=C.诊疗项目ID And B.收费细目ID=C.收费项目ID" & _
            " And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])"
        Set rsCur = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)))
        
        '加入药品及现有的计价
        For i = 1 To rsMan.RecordCount
            mrsPrice.AddNew '暂未输入计价关系的也要加入用于确定可计价医嘱(该记录无用)
            mrsPrice!医嘱ID = rsMan!ID
            mrsPrice!相关ID = rsMan!相关ID
            mrsPrice!诊疗类别 = rsMan!诊疗类别
            mrsPrice!诊疗项目ID = rsMan!诊疗项目ID
            mrsPrice!固定 = IIF(InStr(",5,6,7,", rsMan!诊疗类别) > 0, 1, 0)
            
            '在设置医嘱计价时,对于原已设置,但现已撤档的项目,当作未设置(以便重新增加)
            If Not IsNull(rsMan!收费细目ID) _
                And Format(Nvl(rsMan!撤档时间, "3000-01-01"), "yyyy-MM-dd") = "3000-01-01" Then
                mrsPrice!收费类别 = rsMan!收费类别
                mrsPrice!收费细目ID = rsMan!收费细目ID
                mrsPrice!执行科室ID = rsMan!执行科室ID
                mrsPrice!在用 = Nvl(rsMan!跟踪在用, 0)
                mrsPrice!数量 = rsMan!数量
                
                '药品(仅用于显示)：如果为时价，显示时计算；否则就是取的最新价格
                '非药品：如果为变价,则取以前定的(如果有)；否则下面取最新价格
                mrsPrice!单价 = rsMan!单价
                mrsPrice!从项 = Nvl(rsMan!从项, 0)
                        
                '诊疗收费关系中收费数量及固有对照是否变化
                If InStr(",5,6,7,", rsMan!诊疗类别) = 0 Then '包含非药嘱的药品计价
                    rsCur.Filter = "诊疗项目ID=" & rsMan!诊疗项目ID & " And 收费项目ID=" & rsMan!收费细目ID
                    If Not rsCur.EOF Then
                        If Nvl(rsCur!固有对照, 0) <> 0 And Nvl(rsMan!数量, 0) <> Nvl(rsCur!收费数量, 0) Then
                            mrsPrice!数量 = rsCur!收费数量 '变成了固有对照才取新设置的数量
                            blnChange = True
                        End If
                        mrsPrice!从项 = Nvl(rsCur!从属项目, 0)
                        mrsPrice!固定 = Nvl(rsCur!固有对照, 0)
                    End If
                    '价格取最新的(非变价)
                    dblPrice = CalcPrice(rsMan!收费细目ID)
                    If dblPrice <> 0 Then mrsPrice!单价 = Format(dblPrice, "0.00000")
                End If
            End If
            mrsPrice.Update
            If mrsPrice!从项 = 1 Then blnSubItem = True '存在从属项目
            
            '诊疗收费关系中新增了的对照(在未校对之前,病人医嘱计价没有内容,这时也是相对增加的)
            If InStr(",5,6,7,", rsMan!诊疗类别) = 0 Then '包含非药嘱的药品计价
                lng诊疗项目ID = rsMan!诊疗项目ID
                blnLoad = False: rsMan.MoveNext
                If rsMan.EOF Then
                    blnLoad = True
                ElseIf rsMan!诊疗项目ID <> lng诊疗项目ID Then
                    blnLoad = True
                End If
                rsMan.MovePrevious
                If blnLoad Then
                    strSQL = _
                        " Select A.诊疗项目ID,C.类别 as 收费类别,A.收费项目ID,A.收费数量,A.固有对照,Nvl(A.从属项目,0) as 从属项目," & _
                        " C.类别,B.病人科室ID,B.执行科室ID,E.跟踪在用,Sum(Decode(Nvl(C.是否变价,0),1,NULL,D.现价)) as 单价" & _
                        " From 诊疗收费关系 A,病人医嘱记录 B,收费项目目录 C,收费价目 D,材料特性 E" & _
                        " Where A.诊疗项目ID=B.诊疗项目ID And B.ID=[1] And C.ID=E.材料ID(+)" & _
                        " And A.收费项目ID Not IN(Select 收费细目ID From 病人医嘱计价 Where 医嘱ID=[1])" & _
                        " And A.收费项目ID=C.ID And A.收费项目ID=D.收费细目ID And C.服务对象 IN(2,3)" & _
                        " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                        " Group by A.诊疗项目ID,C.类别,A.收费项目ID,A.收费数量,A.固有对照,Nvl(A.从属项目,0),C.类别,B.病人科室ID,B.执行科室ID,E.跟踪在用" & _
                        " Order by 从属项目"
                    Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!ID))
                    If Not rsAdd.EOF Then
                        For j = 1 To rsAdd.RecordCount
                            '非药嘱和跟踪在用的卫材专门求执行科室
                            lng执行科室ID = Nvl(rsAdd!执行科室ID, 0)
                            If InStr(",5,6,7,", rsAdd!类别) > 0 Or rsAdd!类别 = "4" And Nvl(rsAdd!跟踪在用, 0) = 1 Then
                                lng执行科室ID = Get收费执行科室ID(Val(.TextMatrix(lngRow, COL_病人ID)), Val(.TextMatrix(lngRow, COL_主页ID)), rsAdd!类别, rsAdd!收费项目ID, 4, Nvl(rsAdd!病人科室ID, 0), 0, 2, lng执行科室ID)
                            End If
                            
                            mrsPrice.AddNew
                            mrsPrice!医嘱ID = rsMan!ID
                            mrsPrice!相关ID = rsMan!相关ID
                            mrsPrice!诊疗类别 = rsMan!诊疗类别
                            mrsPrice!诊疗项目ID = rsMan!诊疗项目ID
                            mrsPrice!收费类别 = rsAdd!收费类别
                            mrsPrice!收费细目ID = rsAdd!收费项目ID
                            If lng执行科室ID <> 0 Then
                                mrsPrice!执行科室ID = lng执行科室ID
                            Else
                                mrsPrice!执行科室ID = Null
                            End If
                            mrsPrice!在用 = Nvl(rsAdd!跟踪在用, 0)
                            mrsPrice!数量 = rsAdd!收费数量
                            mrsPrice!单价 = rsAdd!单价
                            mrsPrice!从项 = Nvl(rsAdd!从属项目, 0)
                            mrsPrice!固定 = Nvl(rsAdd!固有对照, 0)
                            mrsPrice.Update
                            
                            If mrsPrice!从项 = 1 Then blnSubItem = True '存在从属项目
                            If Nvl(mrsPrice!数量, 0) <> 0 Then blnChange = True '有变化
                            
                            rsAdd.MoveNext
                        Next
                        
                        '确定了对应收费项目,删除无收费项目的无用记录
                        mrsPrice.Filter = "医嘱ID=" & rsMan!ID
                        Do While Not mrsPrice.EOF
                            If IsNull(mrsPrice!收费细目ID) Then
                                mrsPrice.Delete
                                mrsPrice.Update
                            End If
                            mrsPrice.MoveNext
                        Loop
                    End If
                    
                    '对存在从项的计价进行处理，保证只有一个主项
                    If blnSubItem Then
                        j = 0
                        strSQL = _
                            " Select Sum(Decode(从属项目,1,1,0)) as 从项数," & _
                            " Max(Decode(从属项目,1,NULL,收费项目ID)) as 主项ID" & _
                            " From 诊疗收费关系 Where 诊疗项目ID=[1]"
                        Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!诊疗项目ID))
                        If Not rsMan.EOF Then j = Nvl(rsAdd!从项数, 0)
                        If j = 0 Then
                            '如果现有计价没有从属项目，则取消所有从项属性
                            mrsPrice.Filter = "医嘱ID=" & rsMan!ID
                            Do While Not mrsPrice.EOF
                                If mrsPrice!从项 = 1 Then
                                    mrsPrice!从项 = 0
                                    mrsPrice.Update
                                    blnChange = True
                                End If
                                mrsPrice.MoveNext
                            Loop
                        Else
                            '如果存在从属项目，则除主项以外全部设置为从项
                            mrsPrice.Filter = "医嘱ID=" & rsMan!ID
                            Do While Not mrsPrice.EOF
                                If mrsPrice!收费细目ID = Val(Nvl(rsAdd!主项ID, 0)) Then '为什么一定要加Val?
                                    If mrsPrice!从项 = 1 Then
                                        mrsPrice!从项 = 0 '主项肯定有且只有一个
                                        mrsPrice.Update
                                        blnChange = True
                                    End If
                                Else
                                    If mrsPrice!从项 = 0 Then
                                        mrsPrice!从项 = 1
                                        mrsPrice.Update
                                        blnChange = True
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Loop
                        End If
                    End If
                    blnSubItem = False '新的一条医嘱开始判断
                End If
            End If
            
            rsMan.MoveNext
        Next
        .RowData(lngRow) = 1
    End With
    
    LoadPrice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowPrice(ByVal lngRow As Long)
'功能：显示当前医嘱行的计价内容(包含相关医嘱的计价项目),同时设置一些编辑属性
    Dim rs诊疗项目 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    Dim str医嘱IDs As String, str收费细目IDs As String
    Dim strSQL As String, strTmp As String
    Dim str计价医嘱 As String, i As Long, j As Long
    Dim blnNoFirst As Boolean, lngBegin As Long
    Dim blnAllFixed As Boolean, blnHavePrice As Boolean
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    
    On Error GoTo errH
    
    With vsPrice
        .Redraw = False
        '清除价目表格
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Editable = flexEDNone
        
        '是否一并给药中的非第一药品行
        If RowIn一并给药(lngRow, lngBegin, 0) Then
            If lngRow > lngBegin Then blnNoFirst = True
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            If blnNoFirst Then
                '一并给药时仅第一行显示给药途径的计价
                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
            Else
                mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                    " Or 医嘱ID=" & Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
            End If
        Else
            mrsPrice.Filter = "医嘱ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                " Or 相关ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
        End If
        
        If Not mrsPrice.EOF Then
'            If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
'                mrsPrice.Sort = "诊疗类别" '一并给药时显示顺序要求药品在前
'            Else
'                mrsPrice.Sort = ""
'            End If
                        
            '获取诊疗项目,收费细目,价格信息
            For i = 1 To mrsPrice.RecordCount
                str医嘱IDs = str医嘱IDs & "," & mrsPrice!医嘱ID
                If Not IsNull(mrsPrice!收费细目ID) Then
                    str收费细目IDs = str收费细目IDs & "," & mrsPrice!收费细目ID
                End If
                mrsPrice.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
            str收费细目IDs = Mid(str收费细目IDs, 2)
                        
            strSQL = "Select B.ID,C.名称 as 类别名称,B.名称,B.标本部位" & _
                " From 病人医嘱记录 A,诊疗项目目录 B,诊疗项目类别 C" & _
                " Where A.ID IN(" & str医嘱IDs & ") And A.诊疗项目ID=B.ID And B.类别=C.编码"
            Call zlDatabase.OpenRecordset(rs诊疗项目, strSQL, Me.Caption) 'In
            
            '读取是否变价及变价范围等项目信息
            If str收费细目IDs <> "" Then
                strSQL = _
                    " Select A.ID,C.名称 as 类别名称,A.编码,A.名称,A.规格," & _
                    " A.产地,A.计算单位,D.住院单位,A.费用类型,A.是否变价,D.住院包装" & _
                    " From 收费项目目录 A,收费项目类别 C,药品规格 D" & _
                    " Where A.类别=C.编码 And A.ID=D.药品ID" & _
                    " And A.类别 IN('5','6','7') And A.ID IN(" & str收费细目IDs & ")"
                strSQL = strSQL & " Union ALL " & _
                    " Select A.ID,C.名称 as 类别名称,A.编码,A.名称,A.规格,A.产地," & _
                    " A.计算单位,NULL as 住院单位,A.费用类型,A.是否变价,-NULL as 住院包装" & _
                    " From 收费项目目录 A,收费项目类别 C" & _
                    " Where A.类别=C.编码 And A.类别 Not IN('5','6','7')" & _
                    " And A.ID IN(" & str收费细目IDs & ")"
                
                strSQL = _
                    " Select A.ID,A.类别名称,A.编码,A.名称,A.规格,A.产地,A.计算单位," & _
                    " A.住院单位,A.费用类型,A.是否变价,A.住院包装,Sum(B.原价) as 原价,Sum(B.现价) as 现价" & _
                    " From (" & strSQL & ") A,收费价目 B Where A.ID=B.收费细目ID" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Group by A.ID,A.类别名称,A.编码,A.名称,A.规格,A.产地,A.计算单位,A.住院包装,A.费用类型,A.是否变价,A.住院单位"

                strSQL = _
                    " Select A.ID,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称,A.规格,A.产地," & _
                    " A.计算单位,A.住院单位,A.费用类型,A.是否变价,A.原价,A.现价,A.住院包装" & _
                    " From (" & strSQL & ") A,收费项目别名 B" & _
                    " Where A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIF(gbln商品名, 3, 1)
                Call zlDatabase.OpenRecordset(rs收费细目, strSQL, Me.Caption) 'In
            End If
                        
            '确定显示行数
            If str收费细目IDs <> "" Then
                .Rows = .FixedRows + UBound(Split(str收费细目IDs, ",")) + 1
            End If
                                    
            '显示每行内容
            j = .FixedRows
            blnAllFixed = True: blnHavePrice = False
            mrsPrice.MoveFirst
            For i = 1 To mrsPrice.RecordCount
                '确定计价医嘱内容
                rs诊疗项目.Filter = "ID=" & mrsPrice!诊疗项目ID
                If InStr(",5,6,7,", mrsPrice!诊疗类别) > 0 Then
                    str计价医嘱 = "药品医嘱-" & rs诊疗项目!名称
                ElseIf mrsPrice!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    str计价医嘱 = "给药途径-" & rs诊疗项目!名称
                ElseIf mrsPrice!诊疗类别 = "E" And InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_类型))) > 0 Then
                    If vsAdvice.TextMatrix(lngRow, COL_类型) = "2" Then
                        str计价医嘱 = "采集方法-" & rs诊疗项目!名称
                    ElseIf Not IsNull(mrsPrice!相关ID) Then
                        str计价医嘱 = "中药煎法-" & rs诊疗项目!名称
                    Else
                        str计价医嘱 = "中药用法-" & rs诊疗项目!名称
                    End If
                ElseIf Not IsNull(mrsPrice!相关ID) Then
                    If mrsPrice!诊疗类别 = "C" Then
                        str计价医嘱 = "检验项目-" & rs诊疗项目!名称
                    ElseIf mrsPrice!诊疗类别 = "D" Then
                        str计价医嘱 = "检查部位-" & rs诊疗项目!标本部位
                    ElseIf mrsPrice!诊疗类别 = "F" Then
                        str计价医嘱 = "附加手术-" & rs诊疗项目!名称
                    ElseIf mrsPrice!诊疗类别 = "G" Then
                        str计价医嘱 = "麻醉项目-" & rs诊疗项目!名称
                    End If
                Else
                    str计价医嘱 = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称
                End If
                
                '可以选择的计价医嘱(包含暂未设置收费关系的)
                If mrsPrice!固定 = 0 Then
                    If InStr(strTmp, "|#" & mrsPrice!医嘱ID & ";" & str计价医嘱) = 0 Then
                        strTmp = strTmp & "|#" & mrsPrice!医嘱ID & ";" & str计价医嘱
                    End If
                End If
                '如果已对应的收费关系全为固定,则不允许选择
                If InStr(",5,6,7,", mrsPrice!诊疗类别) = 0 Then
                    If Not IsNull(mrsPrice!收费细目ID) Then
                        '除开药品,是否有计价关系
                        blnHavePrice = True
                        '除开药品的计价关系是否全部为固定
                        blnAllFixed = blnAllFixed And (mrsPrice!固定 <> 0)
                    End If
                End If
                
                '暂未设置收费关系的不显示,但可以选择
                If Not IsNull(mrsPrice!收费细目ID) Then
                    rs收费细目.Filter = "ID=" & mrsPrice!收费细目ID
                    
                    '显示计价的医嘱内容
                    .TextMatrix(j, COLP_计价医嘱) = str计价医嘱
                    .TextMatrix(j, COLP_医嘱ID) = mrsPrice!医嘱ID
                    .TextMatrix(j, COLP_相关ID) = Nvl(mrsPrice!相关ID)
                    .TextMatrix(j, COLP_诊疗类别) = mrsPrice!诊疗类别
                    .TextMatrix(j, COLP_诊疗项目ID) = mrsPrice!诊疗项目ID
                        
                    '显示具体计价的项目
                    .TextMatrix(j, COLP_收费类别) = mrsPrice!收费类别
                    .TextMatrix(j, COLP_收费细目ID) = mrsPrice!收费细目ID
                    .TextMatrix(j, COLP_类别) = rs收费细目!类别名称
                    .TextMatrix(j, COLP_收费项目) = rs收费细目!名称
                    If Not IsNull(rs收费细目!产地) Then
                        .TextMatrix(j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目) & "(" & rs收费细目!产地 & ")"
                    End If
                    If Not IsNull(rs收费细目!规格) Then
                        .TextMatrix(j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目) & " " & rs收费细目!规格
                    End If
                    
                    '非药嘱药品以售价单位设置
                    If InStr(",5,6,7,", mrsPrice!诊疗类别) Then
                        .TextMatrix(j, COLP_单位) = Nvl(rs收费细目!住院单位)
                    Else
                        .TextMatrix(j, COLP_单位) = Nvl(rs收费细目!计算单位)
                    End If
                    '药嘱缺省为1,非药嘱药品可设置(售价单位)
                    .TextMatrix(j, COLP_数量) = FormatEx(mrsPrice!数量, 5)
                    
                    '药嘱药品为按1个住院单位计算的价格
                    .TextMatrix(j, COLP_单价) = Format(Nvl(mrsPrice!单价), "0.00000")
                    
                    '执行科室
                    lng执行科室ID = Nvl(mrsPrice!执行科室ID, 0)
                    '非药嘱药品或跟踪在用的卫材计价可以设置执行科室
                    If mrsPrice!收费类别 = "4" And Nvl(mrsPrice!在用, 0) = 1 _
                        Or InStr(",5,6,7,", mrsPrice!收费类别) > 0 And InStr(",5,6,7,", mrsPrice!诊疗类别) = 0 Then
                        '以当前值作为缺省重新取有效的执行科室
                        lng病人科室ID = Val(vsAdvice.TextMatrix(lngRow, COL_病人科室ID))
                        lng执行科室ID = Get收费执行科室ID(Val(vsAdvice.TextMatrix(lngRow, COL_病人ID)), Val(vsAdvice.TextMatrix(lngRow, COL_主页ID)), _
                            mrsPrice!收费类别, rs收费细目!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
                        '记录是否跟踪在用
                        If mrsPrice!收费类别 = "4" Then
                            .TextMatrix(j, COLP_跟踪在用) = Val(Nvl(mrsPrice!在用, 0))
                        End If
                        .Editable = flexEDKbdMouse
                    End If
                    If lng执行科室ID <> 0 Then
                        mrsDept.Filter = "ID=" & lng执行科室ID
                        If Not mrsDept.EOF Then
                            .TextMatrix(j, COLP_执行科室) = mrsDept!名称
                        End If
                    End If
                    .TextMatrix(j, COLP_执行科室ID) = lng执行科室ID
                                        
                    '变价的处理
                    If Nvl(rs收费细目!是否变价, 0) = 1 Then
                        If InStr(",5,6,7,", mrsPrice!收费类别) > 0 Then
                            If InStr(",5,6,7,", mrsPrice!诊疗类别) > 0 Then
                                '药嘱药品计算1个住院单位的时价
                                .TextMatrix(j, COLP_单价) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rs收费细目!住院包装, 1))
                                .TextMatrix(j, COLP_单价) = Format(Val(.TextMatrix(j, COLP_单价)) * Nvl(rs收费细目!住院包装, 1), "0.00000")
                            Else
                                '非药嘱药品按零售单位计算
                                .TextMatrix(j, COLP_单价) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, mrsPrice!数量), "0.00000")
                            End If
                        ElseIf mrsPrice!收费类别 = "4" And Nvl(mrsPrice!在用, 0) = 1 Then
                            '时价卫材价格的药品一样计算
                            .TextMatrix(j, COLP_单价) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, mrsPrice!数量), "0.00000")
                        Else
                            '记录可以输入的价格范围
                            .Cell(flexcpData, j, 0) = 1 '标识为变价(药品不管)
                            .Cell(flexcpData, j, 1) = Nvl(rs收费细目!原价, 0)
                            .Cell(flexcpData, j, 2) = Nvl(rs收费细目!现价, 0)
                            '也许以前定了变价,现在变价范围变了
                            If .TextMatrix(j, COLP_单价) <> "" Then
                                If CheckScope(Nvl(rs收费细目!原价, 0), Nvl(rs收费细目!现价, 0), Nvl(mrsPrice!单价, 0)) <> "" Then
                                    .TextMatrix(j, COLP_单价) = ""
                                End If
                            End If
                            '非药品变价,即使固定也可以编辑
                            .Editable = flexEDKbdMouse
                        End If
                    End If

                    .TextMatrix(j, COLP_费用类型) = Nvl(rs收费细目!费用类型)
                    .TextMatrix(j, COLP_固定) = mrsPrice!固定
                    .TextMatrix(j, COLP_从项) = IIF(Nvl(mrsPrice!从项, 0) = 0, "", "√")
                    
                    '记录用于恢复输入
                    .Cell(flexcpData, j, COLP_计价医嘱) = .TextMatrix(j, COLP_计价医嘱)
                    .Cell(flexcpData, j, COLP_收费项目) = .TextMatrix(j, COLP_收费项目)
                    .Cell(flexcpData, j, COLP_数量) = .TextMatrix(j, COLP_数量)
                    .Cell(flexcpData, j, COLP_单价) = .TextMatrix(j, COLP_单价)
                    .Cell(flexcpData, j, COLP_执行科室) = .TextMatrix(j, COLP_执行科室)
                    
                    '标识固定对照为灰色
                    If mrsPrice!固定 <> 0 Then
                        .Cell(flexcpBackColor, j, .FixedCols, j, .Cols - 1) = &HE0E0E0
                    End If
                    
                    j = j + 1
                End If
                
                mrsPrice.MoveNext
            Next
            
            '设置编辑数据
            '------------------------------------------------------------------
            '需要计价的医嘱选择
            If strTmp <> "" And Not (blnHavePrice And blnAllFixed) Then
                .ColData(COLP_计价医嘱) = Mid(strTmp, 2)
                .Editable = flexEDKbdMouse '可以选择则可以编辑
            Else
                .ColData(COLP_计价医嘱) = ""
            End If
        End If
        .Row = .FixedRows: .Col = COLP_计价医嘱
        
        '缺省选择计价医嘱(如果可能)
        Call ShowDefaultRow
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
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

Private Sub SetSameTime(ByVal lngRow As Long)
'功能：设置其它医嘱行为相同的校对,暂停,启用时间
    Dim strTime As String, vPause As Date
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        strTime = Format(.TextMatrix(lngRow, COL_输入), "yyyy-MM-dd HH:mm")
        For i = .FixedRows To .Rows - 1
            If i <> lngRow Then
                blnDo = True
                If mint类型 = 3 Then
                    '应>=开嘱时间
                    If strTime < Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                ElseIf mint类型 = 5 Then
                    '应>=开始执行时间,因为该时间点尚未执行
                    If strTime < Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                    '应>上次执行时间,因为该时间点已执行
                    If .TextMatrix(i, COL_上次执行) <> "" Then
                        If strTime <= Format(.Cell(flexcpData, i, COL_上次执行), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    '应<执行终止时间,因为该时间点执行有效
                    If .TextMatrix(i, COL_终止时间) <> "" Then
                        If strTime >= Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    '应>上次暂停后的启用时间(如果有,操作时间不能重复,应>)
                    vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 7)
                    If vPause <> CDate(0) Then
                        If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                ElseIf mint类型 = 6 Then
                    '应>暂停时间
                    vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 6)
                    If vPause <> CDate(0) Then
                        If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    '应<=执行终止时间
                    If .TextMatrix(i, COL_终止时间) <> "" Then
                        If strTime > Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                End If
                If blnDo Then
                    .TextMatrix(i, COL_输入) = strTime
                    .Cell(flexcpData, i, COL_输入) = strTime
                End If
            End If
        Next
    End With
End Sub

Private Function GetPauseTime(ByVal lng医嘱ID As Long, ByVal int状态 As Integer) As Date
'功能：读取指定医嘱的暂停时间(该医嘱当前应已暂停)或上次启用时间(如果有)
'参数：int状态=6-暂停,7-启用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(操作时间) as 上次时间 From 病人医嘱状态 Where 操作类型=[2] And 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, int状态)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!上次时间) Then
            GetPauseTime = rsTmp!上次时间
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceCheckWarn(ByVal lngCmd As Long, ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号:lngCmd=0时需要,多病人批量操作时需要当前病人行
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    If Not (lngRow >= vsAdvice.FixedRows) Then Exit Function '必须要确定病人所在行
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If
    
    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    lng病人ID = Val(vsAdvice.TextMatrix(lngRow, COL_病人ID))
    lng主页ID = Val(vsAdvice.TextMatrix(lngRow, COL_主页ID))
    If lng病人ID <> mlngPassPati Then
        strSQL = _
            " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
            " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
            " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
            " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
            " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    
        Call PassSetPatientInfo(lng病人ID, lng主页ID, rsTmp!姓名, Nvl(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
            rsTmp!科室码 & "/" & rsTmp!科室名, IIF(Not IsNull(rsTmp!医生名), Nvl(rsTmp!医生码) & "/" & Nvl(rsTmp!医生名), ""), _
            IIF(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        mlngPassPati = lng病人ID
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With vsAdvice
            If Val(.TextMatrix(lngRow, COL_ID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                '取药品名称
                str药品 = .TextMatrix(lngRow, COL_医嘱内容)
                If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
                If InStr(str药品, "(") > 0 Then str药品 = Left(str药品, InStr(str药品, "(") - 1)
                '取药品给药途径
                str用法 = .TextMatrix(lngRow, COL_用法)
                
                '传入查询药品信息
                Call PassSetQueryDrug(.TextMatrix(lngRow, COL_收费细目ID), str药品, .TextMatrix(lngRow, COL_单量单位), str用法)
                    
                '设置菜单可用状态
                Call SetPassMenuState
                
                AdviceCheckWarn = 1 '表示可以弹出菜单
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    
    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Pass
    If Button = 2 And gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 And mint类型 = 3 Then
        If AdviceCheckWarn(0, vsAdvice.Row) >= 0 Then PopupMenu mnuPass, 2
    End If
End Sub

Private Sub SetPassMenuState()
'功能：设置Pass菜单可用状态
    'Pass
    '一级菜单
    '药物临床信息参考
    mnuPassItem(0).Enabled = PassGetState("CPRRes") = 1
    '药品说明书
    mnuPassItem(1).Enabled = PassGetState("Directions") = 1
    '中国药典
    mnuPassItem(2).Enabled = PassGetState("Chp") = 1
    '病人用药教育
    mnuPassItem(3).Enabled = PassGetState("CPERes") = 1
    '检验值
    mnuPassItem(4).Enabled = PassGetState("CheckRes") = 1
    '专项信息
    'mnuPassItem(6).Enabled = PassGetState("") = 1
    '医药信息中心
    mnuPassItem(8).Enabled = PassGetState("MEDInfo") = 1
    '药品配对信息
    mnuPassItem(10).Enabled = PassGetState("MATCH-DRUG") = 1
    '给药途径配对信息
    mnuPassItem(11).Enabled = PassGetState("MATCH-ROUTE") = 1
    '医院药品信息
    mnuPassItem(12).Enabled = PassGetState("HisDrugInfo") = 1
    
    '二菜菜单
    '药物-药物相互作用
    mnuPassSpec(0).Enabled = PassGetState("DDIM") = 1
    '药物-食物相互使用
    mnuPassSpec(1).Enabled = PassGetState("DFIM") = 1
    '国内注射剂体外配伍
    mnuPassSpec(3).Enabled = PassGetState("MatchRes") = 1
    '国外注射剂体外配伍
    mnuPassSpec(4).Enabled = PassGetState("TriessRes") = 1
    '禁忌症
    mnuPassSpec(6).Enabled = PassGetState("DDCM") = 1
    '副作用
    mnuPassSpec(7).Enabled = PassGetState("SIDE") = 1
    '老年人用药
    mnuPassSpec(9).Enabled = PassGetState("GERI") = 1
    '儿童用药
    mnuPassSpec(10).Enabled = PassGetState("PEDI") = 1
    '妊娠期用药
    mnuPassSpec(11).Enabled = PassGetState("PREG") = 1
    '哺乳期用药
    mnuPassSpec(12).Enabled = PassGetState("LACT") = 1
End Sub

Private Sub mnuPassItem_Click(Index As Integer)
'功能：执行PASS命令
    'Pass
    Select Case Index
    Case 0 '药物临床信息参考
        Call PassDoCommand(101)
    Case 1 '药品说明书
        Call PassDoCommand(102)
    Case 2 '中国药典
        Call PassDoCommand(107)
    Case 3 '病人用药教育
        Call PassDoCommand(103)
    Case 4 '检验值
        Call PassDoCommand(104)
    Case 8 '医药信息中心
        Call PassDoCommand(106)
    Case 10 '药品配对信息
        Call PassDoCommand(13)
    Case 11 '给药途径配对信息
        Call PassDoCommand(14)
    Case 12 '医院药品信息
        Call PassDoCommand(105)
    End Select
End Sub

Private Sub mnuPassSpec_Click(Index As Integer)
'功能：执行专项PASS命令
    'Pass
    Select Case Index
    Case 0 '药物-药物相互作用
        Call PassDoCommand(201)
    Case 1 '药物-食物相互使用
        Call PassDoCommand(202)
    Case 3 '国内注射剂配伍
        Call PassDoCommand(203)
    Case 4 '国外注射剂配伍
        Call PassDoCommand(204)
    Case 6 '禁忌症
        Call PassDoCommand(205)
    Case 7 '副作用
        Call PassDoCommand(206)
    Case 9 '老年人用药
        Call PassDoCommand(207)
    Case 10 '儿童用药
        Call PassDoCommand(208)
    Case 11 '妊娠期用药
        Call PassDoCommand(209)
    Case 12 '哺乳期用药
        Call PassDoCommand(210)
    End Select
End Sub
