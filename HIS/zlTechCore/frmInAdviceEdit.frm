VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmInAdviceEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "住院医嘱编辑"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmInAdviceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList img16 
      Left            =   5730
      Top             =   1710
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
            Picture         =   "frmInAdviceEdit.frx":058A
            Key             =   "Pati"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   780
      TabIndex        =   26
      Top             =   1575
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病人"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "住院医师"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "费别"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "护理等级"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   $"frmInAdviceEdit.frx":0B24
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "出院日期"
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "付款方式"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   52
      Top             =   7740
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInAdviceEdit.frx":0B31
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11404
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceEdit.frx":13C5
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceEdit.frx":19FF
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "计价"
            TextSave        =   "计价"
            Key             =   "Price"
            Object.ToolTipText     =   "显示诊疗计价面板(F8)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin VB.Frame fraPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      TabIndex        =   35
      Top             =   510
      Width           =   10875
      Begin VB.CommandButton cmdAlley 
         Caption         =   "过敏史/病生状态"
         Height          =   350
         Left            =   9135
         TabIndex        =   32
         Top             =   50
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.CommandButton cmdPati 
         Height          =   240
         Left            =   1740
         Picture         =   "frmInAdviceEdit.frx":2039
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F4)"
         Top             =   105
         Width           =   255
      End
      Begin VB.ComboBox cbo婴儿 
         Height          =   300
         ItemData        =   "frmInAdviceEdit.frx":212F
         Left            =   9435
         List            =   "frmInAdviceEdit.frx":2145
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   75
         Width           =   1395
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   780
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   75
         Width           =   1245
      End
      Begin VB.Label lbl婴儿 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婴儿(&B)"
         Height          =   180
         Left            =   8745
         TabIndex        =   30
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:　床号:　性别:　年龄:　费别:  付款方式:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2085
         TabIndex        =   36
         Top             =   135
         Width           =   4140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人(&A)"
         Height          =   180
         Left            =   135
         TabIndex        =   27
         Top             =   135
         Width           =   630
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   33
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
         TabIndex        =   34
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
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Description     =   "增加"
               Object.ToolTipText     =   "增加一条医嘱(Ctrl+A)"
               Object.Tag             =   "增加"
               ImageKey        =   "增加"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "插入"
               Key             =   "插入"
               Description     =   "插入"
               Object.ToolTipText     =   "插入一条医嘱(Ctrl+I)"
               Object.Tag             =   "插入"
               ImageKey        =   "插入"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前医嘱(Del)"
               Object.Tag             =   "删除"
               ImageKey        =   "删除"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停止"
               Key             =   "停止"
               Description     =   "停止"
               Object.ToolTipText     =   "停止当前医嘱(Ctrl+W)"
               Object.Tag             =   "停止"
               ImageKey        =   "停止"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "一并"
               Key             =   "一并"
               Description     =   "一并"
               Object.ToolTipText     =   "一并给药(Ctrl+K)"
               Object.Tag             =   "一并"
               ImageKey        =   "一并"
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "参考"
               Key             =   "参考"
               Description     =   "参考"
               Object.ToolTipText     =   "查看诊疗项目参考(F6)"
               Object.Tag             =   "参考"
               ImageKey        =   "参考"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "参考_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "复制"
               Key             =   "复制"
               Description     =   "复制"
               Object.ToolTipText     =   "复制产生新的医嘱(Ctrl+Y)"
               Object.Tag             =   "复制"
               ImageKey        =   "复制"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "成套"
               Key             =   "成套"
               Description     =   "成套"
               Object.ToolTipText     =   "保存为成套医嘱(Ctrl+T)"
               Object.Tag             =   "成套"
               ImageKey        =   "成套"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Description     =   "保存"
               Object.ToolTipText     =   "保存医嘱(F2)"
               Object.Tag             =   "保存"
               ImageKey        =   "保存"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "签名"
               Key             =   "签名"
               Description     =   "签名"
               Object.ToolTipText     =   "电子签名"
               Object.Tag             =   "签名"
               ImageKey        =   "签名"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   6555
      TabIndex        =   2
      Top             =   2655
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   71958529
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4800
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   10770
      _cx             =   18997
      _cy             =   8467
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
      ForeColorSel    =   0
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
      Rows            =   18
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInAdviceEdit.frx":2194
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
      Begin MSComctlLib.ImageList imgPass 
         Left            =   1035
         Top             =   360
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
               Picture         =   "frmInAdviceEdit.frx":227C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":2576
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":2870
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":2B6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":2E64
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   405
         Top             =   345
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":315E
               Key             =   "紧急"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":3378
               Key             =   "补录"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSign 
         Left            =   1695
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":3892
               Key             =   "签名"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraAdvice 
      Height          =   2040
      Left            =   45
      TabIndex        =   37
      Top             =   5700
      Width           =   10800
      Begin MSComctlLib.ImageList img24 
         Left            =   4185
         Top             =   615
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":3BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceEdit.frx":42DE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrFree 
         Height          =   450
         Left            =   300
         TabIndex        =   57
         Top             =   810
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "img24"
         DisabledImageList=   "img24"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "自由录入医嘱(F3)"
               ImageIndex      =   1
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo附加执行 
         Height          =   300
         Left            =   6255
         TabIndex        =   20
         Text            =   "cbo附加执行"
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txt天数 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2385
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1635
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmd频率 
         Height          =   240
         Left            =   4860
         Picture         =   "frmInAdviceEdit.frx":49D8
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt频率 
         Height          =   300
         Left            =   3495
         TabIndex        =   10
         Top             =   1275
         Width           =   1665
      End
      Begin VB.TextBox txt单量 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1635
         Width           =   1380
      End
      Begin VB.TextBox txt总量 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1635
         Width           =   1515
      End
      Begin VB.CommandButton cmd用法 
         Height          =   240
         Left            =   2445
         Picture         =   "frmInAdviceEdit.frx":4ACE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt用法 
         Height          =   300
         Left            =   930
         TabIndex        =   8
         Top             =   1275
         Width           =   1815
      End
      Begin VB.CommandButton cmd开嘱时间 
         Height          =   240
         Left            =   10320
         Picture         =   "frmInAdviceEdit.frx":4BC4
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "选择日期(F4)"
         Top             =   1665
         Width           =   255
      End
      Begin VB.ComboBox cbo期效 
         Height          =   300
         ItemData        =   "frmInAdviceEdit.frx":4CBA
         Left            =   930
         List            =   "frmInAdviceEdit.frx":4CC4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   1005
      End
      Begin VB.CommandButton cmd终止时间 
         Height          =   240
         Left            =   10320
         Picture         =   "frmInAdviceEdit.frx":4CD8
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "选择日期(F4)"
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton cmd开始时间 
         Height          =   240
         Left            =   4590
         Picture         =   "frmInAdviceEdit.frx":4DCE
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "选择日期(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.CheckBox chk紧急 
         Caption         =   "紧急(&E)"
         Height          =   225
         Left            =   9705
         TabIndex        =   16
         Top             =   233
         Width           =   930
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   4890
         Picture         =   "frmInAdviceEdit.frx":4EC4
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   600
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   285
         Left            =   4890
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   900
         Width           =   285
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Left            =   6255
         TabIndex        =   18
         Text            =   "cbo执行科室"
         Top             =   912
         Width           =   1725
      End
      Begin VB.TextBox txt终止时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   8805
         TabIndex        =   21
         Top             =   1275
         Width           =   1800
      End
      Begin VB.TextBox txt医嘱内容 
         Height          =   660
         Left            =   930
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "按 ~ 键切换快捷浮动面板"
         Top             =   552
         Width           =   3945
      End
      Begin VB.ComboBox cbo医生 
         Height          =   300
         ItemData        =   "frmInAdviceEdit.frx":4FBA
         Left            =   6255
         List            =   "frmInAdviceEdit.frx":4FBC
         TabIndex        =   23
         Top             =   1635
         Width           =   1725
      End
      Begin VB.TextBox txt开始时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3075
         TabIndex        =   3
         Top             =   195
         Width           =   1800
      End
      Begin VB.TextBox txt开嘱时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   8805
         TabIndex        =   24
         Top             =   1635
         Width           =   1800
      End
      Begin VB.ComboBox cbo执行时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   17
         Top             =   555
         Width           =   4350
      End
      Begin VB.ComboBox cbo执行性质 
         Height          =   300
         ItemData        =   "frmInAdviceEdit.frx":4FBE
         Left            =   8805
         List            =   "frmInAdviceEdit.frx":4FCB
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   915
         Width           =   1800
      End
      Begin VB.ComboBox cbo医生嘱托 
         Height          =   300
         Left            =   6255
         TabIndex        =   15
         Top             =   195
         Width           =   3345
      End
      Begin VB.Label lbl附加执行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "附加执行"
         Height          =   180
         Left            =   5490
         TabIndex        =   56
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl天数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用    天"
         Height          =   180
         Left            =   2205
         TabIndex        =   55
         Top             =   1695
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl频率 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "频率"
         Height          =   180
         Left            =   3105
         TabIndex        =   46
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl单量单位 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "单位"
         Height          =   180
         Left            =   4905
         TabIndex        =   42
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl单量 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单量"
         Height          =   180
         Left            =   3105
         TabIndex        =   41
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl总量单位 
         BackStyle       =   0  'Transparent
         Caption         =   "单位"
         Height          =   180
         Left            =   2490
         TabIndex        =   44
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl总量 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总量"
         Height          =   180
         Left            =   540
         TabIndex        =   43
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl开嘱时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开嘱时间"
         Height          =   180
         Left            =   8055
         TabIndex        =   54
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label lbl期效 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱期效"
         Height          =   180
         Left            =   180
         TabIndex        =   53
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl医生嘱托 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生嘱托"
         Height          =   180
         Left            =   5490
         TabIndex        =   51
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl执行科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Left            =   5490
         TabIndex        =   49
         Top             =   972
         Width           =   720
      End
      Begin VB.Label lbl终止时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终止时间"
         Height          =   180
         Left            =   8055
         TabIndex        =   48
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl用法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用法"
         Height          =   180
         Left            =   540
         TabIndex        =   45
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl医嘱内容 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱内容"
         Height          =   180
         Left            =   180
         TabIndex        =   40
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl开嘱医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开嘱医生"
         Height          =   180
         Left            =   5490
         TabIndex        =   39
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label lbl开始时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   2325
         TabIndex        =   38
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl执行时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行时间"
         Height          =   180
         Left            =   5490
         TabIndex        =   47
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lbl执行性质 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行性质"
         Height          =   180
         Left            =   8055
         TabIndex        =   50
         Top             =   975
         Width           =   720
      End
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":4FED
            Key             =   "增加"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5207
            Key             =   "插入"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5421
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":563B
            Key             =   "一并"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5855
            Key             =   "参考"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5A6F
            Key             =   "复制"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5C89
            Key             =   "成套"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":5EA3
            Key             =   "保存"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":659D
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":67B7
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":69D1
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":70CB
            Key             =   "签名"
         EndProperty
      EndProperty
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":77C5
            Key             =   "增加"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":79DF
            Key             =   "插入"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":7BF9
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":7E13
            Key             =   "一并"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":802D
            Key             =   "参考"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":8247
            Key             =   "复制"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":8461
            Key             =   "成套"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":867B
            Key             =   "保存"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":8D75
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":8F8F
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":91A9
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceEdit.frx":98A3
            Key             =   "签名"
         EndProperty
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
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "系统设置(&U)"
         Index           =   14
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "用药研究(&M)"
         Index           =   16
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "警告(&W)"
         Index           =   18
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "审查(&V)"
         Index           =   19
      End
   End
End
Attribute VB_Name = "frmInAdviceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
'入口参数
Private mblnModal As Boolean
Private mfrmParent As Object
Private mstrPrivs As String
Private mbln护士站 As Boolean
Private mbln补录 As Boolean '补录时不能删改非补录医嘱，非补录时可以删改补录医嘱(并保持补录属性)
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng前提ID As Long '医技工作站下医嘱时用
Private mint婴儿 As Integer '修改时用
Private mlng医嘱ID As Long '修改时用
Private mbln会诊 As Boolean '是否给会诊病人下医嘱
Private mbln审核 As Boolean '是否审核医嘱模式(仅医生使用)

'程序变量
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset

Private WithEvents mfrmShortCut As frmClinicShortCut
Attribute mfrmShortCut.VB_VarHelpID = -1
Private WithEvents mfrmPrice As frmAdvicePrice
Attribute mfrmPrice.VB_VarHelpID = -1
Private mcolStock As Collection '存放各个药品库房的出库检查方式
Private mstrDelIDs As String '记录需要被删除的医嘱ID
Private mstr性别 As String '用于项目输入限制判断
Private mint年龄 As Integer '病人的整数年龄
Private mint险类 As Integer '当前病人险类
Private mstr付款码 As String '当前病人医疗付款方式编码
Private mlng病人科室id As Long '病人当前科室ID
Private mlngPassPati As Long 'Pass:上次已传入PASS的病人ID

'本地参数
Private mint简码 As Integer
Private mstrLike As String
Private mbln校对 As Boolean
Private mbln一次性 As Boolean '临嘱缺省为一次性
Private mbln天数 As Boolean
Private msng天数 As Single
Private mbln提醒对码 As Boolean
Private mbln检查诊断 As Boolean

'事件状态控制变量
Private mblnRunFirst As Boolean
Private mblnClickItem As Boolean
Private mblnRowChange As Boolean
Private mblnDoCheck As Boolean

Private Const TIME_LIMIT = 30 '非补录医嘱允许早于的时间
'执行时间示例
Private Const COL_按周执行 = _
    "每周三次 1/8-3/8-5/8 或 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "表示在每周星期一的8:00,星期三的8:00,星期五的8:00这几个时间执行"
Private Const COL_按天执行 = _
    "每天三次 8-12-16 或 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "表示在每天8:00,12:00,16:00这几个时间执行" & vbCrLf & _
    "两天一次 1/8 或 1/8:00" & vbCrLf & _
        vbTab & "表示在每两天中的第1天8:00这个时间执行"
Private Const COL_按时执行 = _
    "每小时两次 1:20-1:40" & vbCrLf & _
        vbTab & "表示在每小时内的20和40分钟这两个时间执行" & vbCrLf & _
    "两小时一次 2:30 或 1:30 或 1:00" & vbCrLf & _
        vbTab & "表示在每两小时内的第2的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时这个时间执行"

'固定列
Private Const COL_F标志 = 0
'可见列索引
Private Const COL_警示 = 1 'Pass:以字符串类型处理,空表示没有审查结果
Private Const COL_期效 = 2
Private Const COL_开始时间 = 3
Private Const COL_医嘱内容 = 4
Private Const COL_总量 = 5
Private Const COL_总量单位 = 6
Private Const COL_单量 = 7
Private Const COL_单量单位 = 8
Private Const COL_频率 = 9
Private Const COL_用法 = 10
Private Const COL_医生嘱托 = 11
Private Const COL_执行时间 = 12
Private Const COL_终止时间 = 13
Private Const COL_开嘱医生 = 14
Private Const COL_开嘱时间 = 15
Private Const COL_校对护士 = 16

'隐藏列索引
Private Const COL_EDIT = 17 '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号,它的Data值=新下的成套方案ID
Private Const COL_相关ID = 18
Private Const COL_婴儿 = 19
Private Const COL_序号 = 20 'Pass:Data值用于记录是否更改了审核结果
Private Const COL_状态 = 21
Private Const COL_类别 = 22
Private Const COL_诊疗项目ID = 23
Private Const COL_名称 = 24
Private Const COL_标本部位 = 25
Private Const COL_收费细目ID = 26
Private Const COL_天数 = 27
Private Const COL_频率次数 = 28
Private Const COL_频率间隔 = 29
Private Const COL_间隔单位 = 30
Private Const COL_计价性质 = 31
Private Const COL_执行科室ID = 32
Private Const COL_执行性质 = 33 '病人医嘱记录.执行性质=诊疗项目目录.执行科室
Private Const COL_开嘱科室ID = 34
Private Const COL_标志 = 35

Private Const COL_计算方式 = 36 '诊疗项目目录.计算方式
Private Const COL_频率性质 = 37 '诊疗项目目录.执行频率
Private Const COL_操作类型 = 38 '诊疗项目目录.操作类型
Private Const COL_库存 = 39 '按住院包装存放的可用库存
Private Const COL_可否分零 = 40
Private Const COL_剂量系数 = 41
Private Const COL_住院单位 = 42
Private Const COL_住院包装 = 43
Private Const COL_处方限量 = 44 '非药诊疗项目为录入限量
Private Const COL_处方职务 = 45
Private Const COL_毒理分类 = 46
Private Const COL_药品剂型 = 47
Private Const COL_签名否 = 48

Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    Optional ByVal lng前提ID As Long, Optional ByVal bln护士站 As Boolean, Optional ByVal bln补录 As Boolean, _
    Optional ByVal int婴儿 As Integer, Optional ByVal lng医嘱ID As Long, Optional ByVal blnModal As Boolean, _
    Optional ByVal bln会诊 As Boolean, Optional bln审核 As Boolean) As Boolean
    
    Set mfrmParent = frmParent
    mblnModal = blnModal
    mstrPrivs = strPrivs
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng前提ID = lng前提ID
    mbln护士站 = bln护士站
    mbln补录 = bln补录
    mint婴儿 = int婴儿
    mlng医嘱ID = lng医嘱ID
    mbln会诊 = bln会诊
    mbln审核 = bln审核
        
    On Error Resume Next
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Private Property Let mblnNoSave(ByVal vData As Boolean)
    tbr.Buttons("保存").Enabled = vData
End Property

Private Property Get mblnNoSave() As Boolean
    mblnNoSave = tbr.Buttons("保存").Enabled
End Property

Private Sub InitAdviceTable()
'功能：初始化表格内容，用在窗体个性化设置恢复之前
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    strHead = _
        ",240,4;期效,500,4;开始时间,1080,1;医嘱内容,3500,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;" & _
        "频率,1200,1;用法,1200,1;医生嘱托,1000,1;执行时间;终止时间,1080,1;" & _
        "开嘱医生,850,1;开嘱时间,1080,1;校对护士,850,1;" & _
        "EDIT;相关ID;婴儿;序号;医嘱状态;诊疗类别;诊疗项目ID;名称;标本部位;收费细目ID;" & _
        "天数;频率次数;频率间隔;间隔单位;计价性质;执行科室ID;执行性质;开嘱科室ID;标志;" & _
        "计算方式;频率性质;操作类型;库存;可否分零;剂量系数;住院单位;住院包装;处方限量;处方职务;毒理分类;药品剂型;签名否"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        
        .ColHidden(COL_警示) = True 'Pass
        '.FrozenCols = COL_医嘱内容 + 1 - .FixedCols
        .ColWidth(0) = 14 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub cbo附加执行_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo附加执行.ListIndex = -1 Then Exit Sub
    
    If cbo附加执行.ItemData(cbo附加执行.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = GetControlRect(cbo附加执行.Hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl附加执行.Caption, , , , , , True, vRect.Left, vRect.Top, txt用法.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = SeekCboIndex(cbo附加执行, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call zlControl.CboSetIndex(cbo附加执行.Hwnd, intIdx)
        End If
    Else
        cbo附加执行.Tag = "1"
        lngRow = vsAdvice.Row
        
        '更新更改了的执行科室医嘱内容
       Call AdviceChange
    End If
End Sub

Private Sub cbo附加执行_GotFocus()
    Call zlControl.TxtSelAll(cbo附加执行)
End Sub

Private Sub cbo附加执行_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo附加执行.ListIndex = -1 Then
            Call cbo附加执行_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo附加执行_Validate(False)
        End If
    End If
End Sub

Private Sub cbo附加执行_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    If cbo附加执行.ListIndex <> -1 Then Exit Sub '已选中
    If cbo附加执行.Text = "" Then Cancel = True: Exit Sub '无输入
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo附加执行.ListCount > 0 Then
        If cbo附加执行.ItemData(cbo附加执行.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    StrInput = UCase(NeedName(cbo附加执行.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, StrInput & "%", mstrLike & StrInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then cbo附加执行.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo附加执行.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo附加执行.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl附加执行.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Refresh开始时间(ByVal lngRow As Long)
'功能：根据参数设置，重新刷新正确的开始时间
    If Not IsDate(txt开始时间.Text) Then
        txt开始时间.Text = GetDefaultTime(lngRow)
    Else
        If cbo期效.ListIndex = 0 Then
            '长嘱
            If gbln长期医嘱次日生效 And Format(txt开始时间.Text, "HH:mm") <> "00:00" Then
                txt开始时间.Text = GetDefaultTime(lngRow)
            End If
        ElseIf cbo期效.ListIndex = 1 Then
            '临嘱
            If gbln长期医嘱次日生效 And Format(txt开始时间.Text, "HH:mm") = "00:00" Then
                If Abs(DateDiff("n", txt开始时间.Text, zlDatabase.Currentdate)) > 30 Then
                    txt开始时间.Text = GetDefaultTime(lngRow)
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo期效_Click()
'功能：更改项目期效时,清空当前行的数据
    Dim lngRow As Long, i As Long
    Dim strTime As String
    
    With vsAdvice
        lngRow = .Row
        If .RowData(lngRow) = 0 Then
            Call Refresh开始时间(lngRow): Exit Sub
        End If
        
        If NeedName(cbo期效.Text) = .TextMatrix(lngRow, COL_期效) Then Exit Sub
        
        '自由录入医嘱直接更改期效
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            .TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2
                .TextMatrix(lngRow, COL_状态) = 1
                Call ReSetColor(lngRow)
            End If
            mblnNoSave = True: Exit Sub
        End If
        
        If CanAlterType(lngRow) Then
            Call AdviceAlterType(lngRow)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, COL_医嘱内容)
        Else
            '一并给药中某一个不准改(因为规格原因),则当前行内容不能清除
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If RowIn一并给药(lngRow) Then
                    MsgBox "一并给药的药品中存在未按规格下达的药品，不能更改为临嘱。", vbInformation, gstrSysName
                    Call zlControl.CboSetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
                    Exit Sub
                End If
            End If
        
            If MsgBox("更改医嘱期效后需要重新输入医嘱内容,要更改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call zlControl.CboSetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
                Exit Sub
            End If
            
            '清除医嘱数据行
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '西成药、中成药:只可能是单独给药的,删除给药途径行,并清除当前行
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow, True)
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '检查组合项目及手术项目
                '删除部位行或手术附加行(附加手术,麻醉项目)
                Call Delete检查手术(lngRow)
                '清除当前行
                Call DeleteRow(lngRow, True)
            ElseIf RowIn配方行(lngRow) Then
                '中药配方：顺序(序号)要求必须严格控制
                '删除组成味药及煎法行:删除之后重新定位的当前行
                lngRow = Delete中药配方(lngRow)
                '清除当前行(中药用法行)
                Call DeleteRow(lngRow, True)
            Else
                '其它项目直接清除当前行内容
                Call DeleteRow(lngRow, True)
            End If
            
            '重新进入行
            strTime = txt开始时间.Text '保留开始时间
            i = cbo期效.ListIndex '保留当前选择的期效
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, COL_医嘱内容)
            txt开始时间.Text = strTime
            cbo期效.ListIndex = i '就是需要再激活以设置开始时间值
            
            Call Refresh开始时间(.Row)
        End If
    End With
End Sub

Private Sub cbo期效_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo期效.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo期效.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo期效.ListCount > 0 Then lngIdx = 0
        cbo期效.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo医生_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo医生.ListIndex = -1 Then Exit Sub
    
    If cbo医生.ItemData(cbo医生.ListIndex) = -1 Then
        If vsAdvice.RowData(vsAdvice.Row) = 0 Then
            cbo医生.Text = "": Exit Sub
        End If
        
        '全院住院科室的医生
        strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(2,3)"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
        vRect = GetControlRect(cbo医生.Hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl开嘱医生.Caption, , , , , , True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo医生, rsTmp!ID)
            If intIdx <> -1 Then
                cbo医生.ListIndex = intIdx
            Else
                cbo医生.AddItem Nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名, cbo医生.ListCount - 1
                cbo医生.ItemData(cbo医生.NewIndex) = rsTmp!ID
                cbo医生.ListIndex = cbo医生.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有住院医生数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的医生(不引发Click)
            intIdx = SeekCboIndex(cbo医生, vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱医生))
            Call zlControl.CboSetIndex(cbo医生.Hwnd, intIdx)
        End If
    Else
        cbo医生.Tag = "1"
        Call AdviceChange
    End If
End Sub

Private Sub Set用法Input(rsInput As ADODB.Recordset, ByVal int类型 As Integer)
'功能：输入给药途径或中药用法后调用
'参数：rsInput=输入或选择的返回记录
'      int类型=2-给药途径,4-中药用法
'说明：如果可选频率,则配合给药途径处理可用执行时间方案的变化
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnValid As Boolean, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim vMsg As VbMsgBoxResult, strMsg As String
    
    cmd用法.Tag = rsInput!ID
    txt用法.Text = rsInput!名称
    txt用法.Tag = "1"
    
    With vsAdvice
        '重新获取可用的缺省时间方案
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            Call Get时间方案(cbo执行时间, Get频率范围(.Row), .TextMatrix(.Row, COL_频率), rsInput!ID)
            If cbo执行时间.ListCount > 0 Then
                Call zlControl.CboSetIndex(cbo执行时间.Hwnd, 0)
                cbo执行时间.Tag = "1"
            Else
                '判断当前执行时间是否合法
                If cbo执行时间.Text <> "" Then
                    blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位))
                    If Not blnValid Then '如果不合法,则另取,否则保持
                        cbo执行时间.Text = ""
                        cbo执行时间.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '根据诊疗用法用量作缺省设置
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            strSQL = "Select 频次,小儿剂量,成人剂量,医生嘱托,疗程 From 诊疗用法用量 Where 性质>0 And 项目ID=[1] And 用法ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_诊疗项目ID)), Val(rsInput!ID))
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!频次) And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 Then '已为一次性时不管
                    Call Get频率信息_编码(rsTmp!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    txt频率.Text = str频率
                    cmd频率.Tag = str频率
                    txt频率.Tag = "1"
                End If
                
                '根据新的频率重新设置执行时间
                If cbo执行时间.Enabled Then
                    Call Get时间方案(cbo执行时间, Get频率范围(.Row), str频率, rsInput!ID)
                    If cbo执行时间.ListCount > 0 Then
                        Call zlControl.CboSetIndex(cbo执行时间.Hwnd, 0)
                        cbo执行时间.Tag = "1"
                    Else
                        '判断当前执行时间是否合法
                        If cbo执行时间.Text <> "" Then
                            blnValid = ExeTimeValid(cbo执行时间.Text, int频率次数, int频率间隔, str间隔单位)
                            If Not blnValid Then '如果不合法,则另取,否则保持
                                cbo执行时间.Text = ""
                                cbo执行时间.Tag = "1"
                            End If
                        End If
                    End If
                End If

                '药品单量
                If mint年龄 > 12 Then
                    If Nvl(rsTmp!成人剂量, 0) <> 0 Then
                        txt单量.Text = FormatEx(rsTmp!成人剂量, 5)
                        txt单量.Tag = "1"
                    End If
                Else
                    If Nvl(rsTmp!小儿剂量, 0) <> 0 Then
                        txt单量.Text = FormatEx(rsTmp!小儿剂量, 5)
                        txt单量.Tag = "1"
                    ElseIf Nvl(rsTmp!成人剂量, 0) <> 0 Then
                        txt单量.Text = FormatEx(rsTmp!成人剂量 * (mint年龄 + 2) * 5 / 100, 5)
                        txt单量.Tag = "1"
                    End If
                End If
                
                '药品临嘱总量:住院包装
                If .TextMatrix(.Row, COL_期效) = "临嘱" And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 Then
                    '取缺省的天数
                    sng天数 = msng天数
                    If mbln天数 Then
                        If str间隔单位 = "周" Then
                            sng天数 = IIF(7 > sng天数, 7, sng天数)
                        ElseIf str间隔单位 = "天" Then
                            sng天数 = IIF(int频率间隔 > sng天数, int频率间隔, sng天数)
                        ElseIf str间隔单位 = "小时" Then
                            sng天数 = IIF(int频率间隔 \ 24 > sng天数, int频率间隔 \ 24, sng天数)
                        End If
                        If sng天数 = 0 Then sng天数 = 1
                    End If
                    If Nvl(rsTmp!疗程, 1) > sng天数 Then
                        sng天数 = Nvl(rsTmp!疗程, 1)
                    End If
                    If Val(.TextMatrix(.Row, COL_天数)) > sng天数 Then
                        sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    End If
                    If Val(.TextMatrix(.Row, COL_天数)) <> sng天数 Then
                        txt天数.Text = sng天数
                        txt天数.Tag = "1"
                    End If
                    
                    If str频率 <> "" And Val(txt单量.Text) <> 0 _
                        And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                        And Val(.TextMatrix(.Row, COL_住院包装)) <> 0 Then
                                                
                        txt总量.Text = FormatEx(Calc缺省药品总量( _
                            Val(txt单量.Text), sng天数, _
                            int频率次数, int频率间隔, str间隔单位, _
                            .TextMatrix(.Row, COL_执行时间), _
                            Val(.TextMatrix(.Row, COL_剂量系数)), _
                            Val(.TextMatrix(.Row, COL_住院包装)), _
                            Val(.TextMatrix(.Row, COL_可否分零))), 5)
                        txt总量.Tag = "1"
                    End If
                End If
                
                '医生嘱托
                If Not IsNull(rsTmp!医生嘱托) Then
                    cbo医生嘱托.Text = rsTmp!医生嘱托
                    cbo医生嘱托.Tag = "1"
                End If
            End If
        End If
    End With
    
    '处理当前医嘱给药途径/煎法的变化
    Call AdviceChange
    
    '对保险对码进行检查
    strMsg = CheckAdviceInsure(vsAdvice.Row)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
End Sub

Private Sub Set频率Input(rsInput As ADODB.Recordset, ByVal int范围 As Integer, ByVal int项目频率 As Integer)
'功能：输入执行频率后调用
'参数：rsInput=输入或选择的返回记录
'      int范围=1-西医;2-中医;-1-一次性;-2-持续性
'      int项目频率=项目本身的执行频率属性
'说明：配合用法处理可用执行时间方案的变化
    Dim lng用法ID As Long, blnValid As Boolean
    Dim sng天数 As Single, str原总量 As String
    Dim str原执行时间 As String, str原天数 As String
    Dim i As Long
    
    str原执行时间 = cbo执行时间.Text
    str原天数 = txt天数.Text
    str原总量 = txt总量.Text
    With vsAdvice
        .TextMatrix(.Row, COL_频率性质) = Decode(int范围, 1, 0, 2, 0, -1, 1, -2, 2)
        If RowIn检验行(.Row) Then '同步赋值,因为后续以检验项目的执行性质作判断
            For i = .Row - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(.Row) Then
                    .TextMatrix(i, COL_频率性质) = .TextMatrix(.Row, COL_频率性质)
                Else
                    Exit For
                End If
            Next
        End If
        cmd频率.Tag = rsInput!名称
        txt频率.Text = rsInput!名称
        txt频率.Tag = "1"
        
        '先设置临嘱药品天数的可用性
        If mbln天数 And cbo期效.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            If Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                If txt天数.Enabled Then SetDayState -1, -1
            Else
                If Not txt天数.Enabled Then SetDayState 1, 1
            End If
        End If
        
        '先设置总量的可用性:临嘱"计次"可选频率的设置为一次性后不输入总量(除药品外)
        If cbo期效.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) = 0 And Not RowIn配方行(.Row) Then
            If Val(.TextMatrix(.Row, COL_计算方式)) = 3 And int项目频率 = 0 Then
                If txt总量.Enabled And Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                    SetItemEditable , -1
                    txt总量.Text = "1"
                ElseIf Not txt总量.Enabled And Val(.TextMatrix(.Row, COL_频率性质)) = 0 Then
                    SetItemEditable , 1
                End If
                lbl总量单位.Caption = .TextMatrix(.Row, COL_总量单位)
            End If
        End If
        
        '先设置执行时间的可用性(临嘱可选频率项目可能在一次性之间切换)
        If cbo期效.ListIndex = 1 Then
            If int项目频率 = 0 And Decode(int范围, 1, 0, 2, 0, -1, 1, -2, 2) <> 1 Then
                If Not cbo执行时间.Enabled Then SetItemEditable , , , , 1
            Else
                If cbo执行时间.Enabled Then SetItemEditable , , , , -1
            End If
        End If
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            '处理可用执行时间方案的变化
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                '查找给药途径对应的行
                lng用法ID = .FindRow(CLng(.TextMatrix(.Row, COL_相关ID)), .Row + 1)
                If lng用法ID <> -1 Then '未找到给药途径的情况,应该不可能
                    lng用法ID = .TextMatrix(lng用法ID, COL_诊疗项目ID)
                Else
                    lng用法ID = 0
                End If
            ElseIf RowIn配方行(.Row) Then
                '得到对应的中药用法ID
                lng用法ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
            End If
            
            Call Get时间方案(cbo执行时间, int范围, txt频率.Text, lng用法ID)
            '取新的频率的默认执行时间
            If cbo执行时间.ListCount > 0 Then
                Call zlControl.CboSetIndex(cbo执行时间.Hwnd, 0)
                cbo执行时间.Tag = "1"
            Else
                '判断当前执行时间是否合法
                If cbo执行时间.Text <> "" Then
                    blnValid = ExeTimeValid(cbo执行时间.Text, rsInput!频率次数, rsInput!频率间隔, rsInput!间隔单位)
                    If Not blnValid Then '如果不合法,则另取,否则保持
                        cbo执行时间.Text = ""
                        cbo执行时间.Tag = "1"
                    End If
                End If
            End If
            
            '重新计算总量
            If mbln天数 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 _
                And .TextMatrix(.Row, COL_期效) = "临嘱" And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 Then
                sng天数 = Val(txt天数.Text)
                If sng天数 = 0 Then sng天数 = 1
                
                If txt频率.Text <> "" And Val(txt单量.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_住院包装)) <> 0 Then
                    
                    txt总量.Text = FormatEx(Calc缺省药品总量( _
                        Val(txt单量.Text), sng天数, rsInput!频率次数, _
                        rsInput!频率间隔, rsInput!间隔单位, cbo执行时间.Text, _
                        Val(.TextMatrix(.Row, COL_剂量系数)), _
                        Val(.TextMatrix(.Row, COL_住院包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    txt总量.Tag = "1"
                End If
            End If
        End If
    End With
    
    '检查总量,天数,执行时间是否变化
    If txt总量.Text <> str原总量 Then txt总量.Tag = "1"
    If txt天数.Text <> str原天数 Then txt天数.Tag = "1"
    If cbo执行时间.Text <> str原执行时间 Then cbo执行时间.Tag = "1"
    
    '处理当前医嘱执行频率的变化
    Call AdviceChange
End Sub

Private Sub cbo医生_GotFocus()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    If cbo医生.ListIndex <> -1 Then Exit Sub '已选中
    If vsAdvice.RowData(vsAdvice.Row) = 0 Then '无医嘱
        cbo医生.Text = "": Exit Sub
    End If
    If cbo医生.Text = "" Then '无输入
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱医生) = "" Then Cancel = True
        Exit Sub
    End If
    
    StrInput = UCase(NeedName(cbo医生.Text))
    '全院住院科室的医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.简码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl开嘱医生.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo医生, rsTmp!ID)
        If intIdx <> -1 Then
            cbo医生.ListIndex = intIdx
        Else
            cbo医生.AddItem Nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名, cbo医生.ListCount - 1
            cbo医生.ItemData(cbo医生.NewIndex) = rsTmp!ID
            cbo医生.ListIndex = cbo医生.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAlley_Click()
'功能：对病人过敏史/病生状态进行管理
    'Pass
    Call AdviceCheckWarn(22)
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
        
    With vsAdvice
        If cbo期效.ListIndex = 1 Then
            int频率 = Get项目频率(.Row)
            If Not RowIn配方行(.Row) And int频率 = 0 Then
                str范围 = "1,-1" '临嘱可以为一次性
            Else
                str范围 = Get频率范围(.Row)
            End If
        Else
            str范围 = Get频率范围(.Row)
            int频率 = Decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2)
        End If
        
        strSQL = _
            " Select Rownum as ID,A.编码,A.名称,A.简码," & _
            " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
            " From 诊疗频率项目 A" & _
            " Where Instr([1],','||A.适用范围||',')>0" & _
            " Order by A.适用范围,A.编码"
        vRect = GetControlRect(txt频率.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗频率", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, "," & str范围 & ",")
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
            End If
            txt频率.Text = .TextMatrix(.Row, COL_频率)
            Call zlControl.TxtSelAll(txt频率)
            txt频率.SetFocus: Exit Sub
        End If
        Call Set频率Input(rsTmp, rsTmp!范围ID, int频率)
        txt频率.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    
    If mblnNoSave Then
        If mbln审核 Then
            strMsg = "确实要放弃审核当前这些医嘱而退出吗？"
        Else
            strMsg = "当前医嘱内容编辑后尚未保存，确实要退出吗？"
        End If
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    If Not mfrmShortCut Is Nothing Then
        mfrmShortCut.SaveShowState '系统自动卸载该子窗体
    End If
End Sub

Private Sub lvwPati_LostFocus()
    lvwPati.Visible = False
End Sub

Private Sub mfrmPrice_PanelHide()
    Call stbThis_PanelClick(stbThis.Panels("Price"))
End Sub

Private Sub mfrmShortCut_ItemClick(ByVal 类型 As Integer, ByVal 分类ID As Long)
    If cmdSel.Enabled And cmdSel.Visible Then
        Call ClinicSelecter(分类ID)
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Price" Then
        If Panel.Bevel <> sbrNoBevel Then
            Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Panel.Tag = IIF(Panel.Bevel = sbrInset, "Show", "")
            Call ShowPrice(vsAdvice.Row)
        End If
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", _
            IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
        mint简码 = IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
    End If
End Sub

Private Sub tbrFree_ButtonClick(ByVal Button As MSComctlLib.Button)
    '强起时清除已有内容
    If Button.Value = 0 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If MsgBox("取消自由录入状态将清除已录入的医嘱内容，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Button.Value = 1
                Call zlControl.TxtSelAll(txt医嘱内容)
                txt医嘱内容.SetFocus: Exit Sub
            End If
            Call DeleteRow(vsAdvice.Row, True)
            mblnNoSave = True
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        End If
    End If
    
    txt医嘱内容.Text = ""
    txt医嘱内容.SetFocus
End Sub

Private Sub txt频率_GotFocus()
    Call zlControl.TxtSelAll(txt频率)
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmd频率.Tag <> "" And txt频率.Text = .TextMatrix(.Row, COL_频率) And txt频率.Text <> "" Then
                Call SeekNextControl
            ElseIf txt频率.Text = "" Then
                If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
            Else
                If cbo期效.ListIndex = 1 Then
                    int频率 = Get项目频率(.Row)
                    If Not RowIn配方行(.Row) And int频率 = 0 Then
                        str范围 = "1,-1" '临嘱可以为一次性
                    Else
                        str范围 = Get频率范围(.Row)
                    End If
                Else
                    str范围 = Get频率范围(.Row)
                    int频率 = Decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2)
                End If
                
                strSQL = _
                    " Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
                    " From 诊疗频率项目 A" & _
                    " Where Instr([3],','||A.适用范围||',')>0" & _
                    " And (A.编码 Like [1] Or Upper(A.名称) Like [2]" & _
                    " Or Upper(A.简码) Like [2] Or Upper(A.英文名称) Like [2])" & _
                    " Order by A.适用范围,A.编码"
                vRect = GetControlRect(txt频率.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗频率", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, UCase(txt频率.Text) & "%", _
                    mstrLike & UCase(txt频率.Text) & "%", "," & str范围 & ",")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                    End If
                    txt频率.Text = .TextMatrix(.Row, COL_频率)
                    Call zlControl.TxtSelAll(txt频率)
                    txt频率.SetFocus: Exit Sub
                End If
                Call Set频率Input(rsTmp, rsTmp!范围ID, int频率)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Function GetBaseRow(ByVal lngRow As Long) As Long
'功能：由当前可见行获取主项目的行
    If RowIn配方行(lngRow) Then
        '获取中药配方第一味中药行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    ElseIf RowIn检验行(lngRow) Then
        '获取一并采样的第一个项目行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    Else
        GetBaseRow = lngRow
    End If
End Function

Private Function Get项目频率(ByVal lngRow As Long) As Integer
'功能：获取指定项目的原始执行频率属性
'参数：lngRow=当前可见行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    lngRow = GetBaseRow(lngRow)
    strSQL = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)))
    If Not rsTmp.EOF Then Get项目频率 = Nvl(rsTmp!执行频率, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt天数_Change()
    txt天数.Tag = "1"
End Sub

Private Sub txt天数_GotFocus()
    Call zlControl.TxtSelAll(txt天数)
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        '为临嘱
        If (IsNumeric(txt单量.Text) Or txt单量.Text = "") _
            And (IsNumeric(txt天数.Text) Or txt天数.Text = "") Then
            If SeekNextControl Then Call txt天数_Validate(False)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt天数_Validate(Cancel As Boolean)
    Dim sng天数 As Single, i As Long
    Dim strSame As String, strMsg As String
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    With vsAdvice
        If Val(txt天数.Text) = 0 Then
            txt天数.Text = 1: txt天数.Tag = "1"
        End If
        
        '天数至少需要一个频率同期的天数
        If Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 Then
            If .TextMatrix(.Row, COL_间隔单位) = "周" Then
                sng天数 = 7
            ElseIf .TextMatrix(.Row, COL_间隔单位) = "天" Then
                sng天数 = Val(.TextMatrix(.Row, COL_频率间隔))
            ElseIf .TextMatrix(.Row, COL_间隔单位) = "小时" Then
                sng天数 = Val(.TextMatrix(.Row, COL_频率间隔)) \ 24
            End If
            If Val(txt天数.Text) < sng天数 Then
                If MsgBox("按""" & .TextMatrix(.Row, COL_频率) & """执行时，至少需要 " & sng天数 & " 天的用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt天数_GotFocus: Exit Sub
                End If
            End If
        End If

        '重新计算总量
        If .TextMatrix(.Row, COL_频率) <> "" _
            And Val(.TextMatrix(.Row, COL_单量)) <> 0 _
            And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
            And Val(.TextMatrix(.Row, COL_住院包装)) <> 0 Then
            
            txt总量.Text = FormatEx(Calc缺省药品总量( _
                Val(.TextMatrix(.Row, COL_单量)), Val(txt天数.Text), _
                Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_住院包装)), _
                Val(.TextMatrix(.Row, COL_可否分零))), 5)
            txt总量.Tag = "1"
        End If
    End With
    
    '每次输入天数后，作为下次的缺省
    If txt天数.Tag = "1" Then
        msng天数 = Val(txt天数.Text)
    End If
    
    Call AdviceChange
    
    '成套方案批量处理
    With vsAdvice
        If Val(.Cell(flexcpData, .Row, COL_EDIT)) <> 0 Then
            strSame = CStr(.Cell(flexcpData, .Row, COL_EDIT))
            If InStr(strSame, ",") > 0 Then
                strMsg = "该次复制的所有的药品都按这个天数执行吗？"
            Else
                strMsg = "该成套方案的所有药品都按这个天数执行吗？"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                For i = .FixedRows To .Rows - 1
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And .TextMatrix(i, COL_期效) = "临嘱" Then
                        If Not (Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) _
                            Or .RowData(i) = Val(.TextMatrix(.Row, COL_相关ID)) Or i = .Row) _
                                And CStr(.Cell(flexcpData, i, COL_EDIT)) = strSame Then
                            If .TextMatrix(i, COL_频率) <> "" _
                                And Val(.TextMatrix(i, COL_单量)) <> 0 _
                                And Val(.TextMatrix(i, COL_剂量系数)) <> 0 _
                                And Val(.TextMatrix(i, COL_住院包装)) <> 0 Then
                                .TextMatrix(i, COL_天数) = txt天数.Text
                                .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(i, COL_单量)), Val(txt天数.Text), _
                                    Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), _
                                    .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), _
                                    Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_住院包装)), _
                                    Val(.TextMatrix(i, COL_可否分零))), 5)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txt用法_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd用法.Tag) <> 0 And txt用法.Text = .TextMatrix(.Row, COL_用法) And txt用法.Text <> "" Then
                Call SeekNextControl
            ElseIf txt用法.Text = "" Then
                If cmd用法.Enabled And cmd用法.Visible Then cmd用法_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    int类型 = 2 '给药途径
                ElseIf RowIn检验行(vsAdvice.Row) Then
                    int类型 = 6 '采集方法
                Else
                    int类型 = 4 '中药用法
                End If
                If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                    strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[4] And 性质>0)" & _
                        " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                            " Where A.用法ID=B.ID And B.服务对象 IN(2,3) And A.项目ID=[4] And A.性质>0)<=1)"
                End If
                
                '优化
                strLike = mstrLike
                If Len(txt用法.Text) < 2 Then strLike = ""
                
                strSQL = "Select Distinct A.ID,A.编码,A.名称" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.ID=B.诊疗项目ID" & _
                    " And A.类别='E' And A.操作类型=[3] And A.服务对象 IN(2,3)" & strSQL & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])" & _
                    Decode(mint简码, 0, " And B.码类 IN([5],3)", 1, " And B.码类 IN([5],3)", "") & _
                    " Order by A.编码"
                vRect = GetControlRect(txt用法.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl用法.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, UCase(txt用法.Text) & "%", _
                    strLike & UCase(txt用法.Text) & "%", CStr(int类型), Val(.TextMatrix(.Row, COL_诊疗项目ID)), mint简码 + 1)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的" & lbl用法.Caption & "。", vbInformation, gstrSysName
                    End If
                    txt用法.Text = .TextMatrix(.Row, COL_用法)
                    Call zlControl.TxtSelAll(txt用法)
                    txt用法.SetFocus: Exit Sub
                End If
                
                '对一并给药的其它药品的可用给药途径进行检查
                If int类型 = 2 Then
                    Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), 2) Then
                                .Refresh
                                MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, COL_医嘱内容) & """。", vbInformation, gstrSysName
                                .Refresh
                                txt用法.Text = .TextMatrix(.Row, COL_用法)
                                Call zlControl.TxtSelAll(txt用法)
                                txt用法.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                
                Call Set用法Input(rsTmp, int类型)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub cmd用法_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            int类型 = 2 '给药途径
        ElseIf RowIn检验行(vsAdvice.Row) Then
            int类型 = 6 '采集方法
        Else
            int类型 = 4 '中药用法
        End If
        If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
            strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[2] And 性质>0)" & _
                " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                    " Where A.用法ID=B.ID And B.服务对象 IN(2,3) And A.项目ID=[2] And A.性质>0)<=1)"
        End If
        strSQL = "Select Distinct A.ID,A.编码,A.名称,C.名称 as 分类" & _
            " From 诊疗项目目录 A,诊疗项目别名 B,诊疗分类目录 C" & _
            " Where A.ID=B.诊疗项目ID And A.分类ID=C.ID(+)" & _
            " And A.类别='E' And A.操作类型=[1] And A.服务对象 IN(2,3)" & strSQL & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " Order by A.编码"
        vRect = GetControlRect(txt用法.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl用法.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, CStr(int类型), Val(.TextMatrix(.Row, COL_诊疗项目ID)))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的" & lbl用法.Caption & "，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            txt用法.Text = .TextMatrix(.Row, COL_用法)
            Call zlControl.TxtSelAll(txt用法)
            txt用法.SetFocus: Exit Sub
        End If
        
        '对一并给药的其它药品的可用给药途径进行检查
        If int类型 = 2 Then
            Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), 2) Then
                        .Refresh
                        MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, COL_医嘱内容) & """。", vbInformation, gstrSysName
                        .Refresh
                        txt用法.Text = .TextMatrix(.Row, COL_用法)
                        Call zlControl.TxtSelAll(txt用法)
                        txt用法.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set用法Input(rsTmp, int类型)
        txt用法.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt用法_GotFocus()
    Call zlControl.TxtSelAll(txt用法)
End Sub

Private Sub txt用法_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If Val(cmd用法.Tag) <> 0 And txt用法.Text <> .TextMatrix(.Row, COL_用法) Then
            txt用法.Text = .TextMatrix(.Row, COL_用法)
        End If
    End With
End Sub

Private Sub txt频率_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If cmd频率.Tag <> "" And txt频率.Text <> .TextMatrix(.Row, COL_频率) Then
            txt频率.Text = .TextMatrix(.Row, COL_频率)
        End If
    End With
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo医生.ListIndex = -1 Then
            Call cbo医生_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo医生_Validate(False)
        End If
    End If
End Sub

Private Sub cbo婴儿_Click()
    If Not Visible Then Exit Sub
    If cbo婴儿.ListIndex = Val(cbo婴儿.Tag) Then Exit Sub
    cbo婴儿.Tag = cbo婴儿.ListIndex
    
    Call ShowAdvice
    
    vsAdvice.SetFocus
End Sub

Private Sub cbo执行科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo执行科室.ListIndex = -1 Then Exit Sub
    
    If cbo执行科室.ItemData(cbo执行科室.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = GetControlRect(cbo执行科室.Hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl执行科室.Caption, , , , , , True, vRect.Left, vRect.Top, cbo执行科室.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = SeekCboIndex(cbo执行科室, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call zlControl.CboSetIndex(cbo执行科室.Hwnd, intIdx)
        End If
    Else
        cbo执行科室.Tag = "1"
        lngRow = vsAdvice.Row
        
        '更新更改了的执行科室医嘱内容
        Call AdviceChange
        
        '重新获取库存并显示：以住院单位，中药配方不显示
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                Call GetDrugStock(lngRow)
                stbThis.Panels(3).Text = "库存: " & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & .TextMatrix(lngRow, COL_住院单位)
            ElseIf RowIn配方行(lngRow) Then
                Call GetDrugStock(lngRow)
            End If
        End With
    End If
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行科室.ListIndex = -1 Then
            Call cbo执行科室_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo执行科室_Validate(False)
        End If
    End If
End Sub

Private Sub cbo执行科室_GotFocus()
    Call zlControl.TxtSelAll(cbo执行科室)
End Sub

Private Sub cbo执行科室_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    If cbo执行科室.ListIndex <> -1 Then Exit Sub '已选中
    If cbo执行科室.Text = "" Then '无输入
        If tbrFree.Buttons(1).Value = 0 Then Cancel = True
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo执行科室.ListCount > 0 Then
        If cbo执行科室.ItemData(cbo执行科室.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    StrInput = UCase(NeedName(cbo执行科室.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, StrInput & "%", mstrLike & StrInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then cbo执行科室.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo执行科室.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo执行科室.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl执行科室.Caption, False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo执行时间_Change()
    cbo执行时间.Tag = "1"
End Sub

Private Sub cbo执行时间_Click()
    'cbo执行时间_Change
    '更新数据
    cbo执行时间.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo执行时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行时间.ListIndex <> -1 Or cbo执行时间.Text <> "" Then
            If SeekNextControl Then Call cbo执行时间_Validate(False)
        End If
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cbo执行时间_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cbo执行时间.Text <> "" Then
            '检查长度
            If Len(cbo执行时间.Text) > 50 Then
                MsgBox "输入内容不能超过 50 个字符。", vbInformation, gstrSysName
                Call cbo执行时间_GotFocus
                Cancel = True: Exit Sub
            End If
            '检查合法性
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(lngRow, COL_频率次数)), Val(.TextMatrix(lngRow, COL_频率间隔)), .TextMatrix(lngRow, COL_间隔单位))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                        strTmp = COL_按周执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                        strTmp = COL_按天执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                        strTmp = COL_按时执行
                    End If
                    MsgBox "输入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cbo执行时间_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        Else
            '可选频率的长嘱必须输入执行时间
            If .TextMatrix(lngRow, COL_期效) = "长嘱" And .TextMatrix(lngRow, COL_频率) <> "" Then
                MsgBox "可选频率的长期医嘱必须输入执行时间方案。", vbInformation, gstrSysName
                Call cbo执行时间_GotFocus
                Cancel = True: Exit Sub
            End If
        End If
    End With
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_Click()
    cbo执行性质.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行性质.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo执行性质.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo执行性质.ListCount > 0 Then lngIdx = 0
        cbo执行性质.ListIndex = lngIdx
    End If
End Sub

Private Sub chk紧急_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk紧急.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub chk紧急_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdExt_Click()
'功能：修改现有医嘱的扩充内容
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strTmp As String
    Dim lngRow As Long, lngDrugRow As Long
    Dim lng诊疗项目ID As Long, lng用法ID As Long
    Dim strMsg As String, vMsg As VbMsgBoxResult
    
    lngRow = vsAdvice.Row
        
    If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
        strExtData = Get检查部位IDs(lngRow)
        frmAdviceEditEx.mintType = 0
    ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
        strExtData = Get手术附加IDs(lngRow)
        frmAdviceEditEx.mintType = 1
    ElseIf RowIn配方行(lngRow) Then
        strExtData = Get中药配方IDs(lngRow)
        frmAdviceEditEx.mintType = 2
    ElseIf RowIn检验行(lngRow) Then
        strExtData = Get检验组合IDs(lngRow)
        frmAdviceEditEx.mintType = 4
    Else
        Exit Sub '兼容以前的检验项目
    End If
        
    frmAdviceEditEx.mstrPrivs = mstrPrivs
    frmAdviceEditEx.mlngHwnd = txt医嘱内容.Hwnd
    frmAdviceEditEx.mint期效 = cbo期效.ListIndex '0-长期,1-临时
    frmAdviceEditEx.mstr性别 = mstr性别
    If frmAdviceEditEx.mintType = 4 Then
        frmAdviceEditEx.mlng项目ID = 0
    Else
        frmAdviceEditEx.mlng项目ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
    End If
    frmAdviceEditEx.mstrExtData = strExtData
    
    frmAdviceEditEx.mbln护士站 = mbln护士站
    frmAdviceEditEx.mbln医保 = InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> ""
    
    On Error Resume Next
    frmAdviceEditEx.Show 1, Me
    On Error GoTo 0
    
    '重新设置相关内容
    If frmAdviceEditEx.mblnOK Then
        strExtData = frmAdviceEditEx.mstrExtData
        
        '非补录医嘱更新开嘱时间
        If Val(vsAdvice.TextMatrix(lngRow, COL_标志)) <> 2 Then
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            txt开嘱时间.Text = vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间)
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
            '检查组合
            Call AdviceSet检查手术(1, lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, COL_医嘱内容)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
            '一组手术
            Call AdviceSet检查手术(2, lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, COL_医嘱内容)
            
            '刷新处理手术麻醉的执行科室
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        ElseIf RowIn检验行(lngRow) Then
            '检验组合
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "开始时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱医生", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "开嘱时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "标志", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '采集方法的执行科室可能与检验项目不同
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID)) <> 0 Then
                rsCurr!执行科室ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_总量)) <> 0 Then
                rsCurr!总量 = Val(vsAdvice.TextMatrix(lngRow, COL_总量))
            End If
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngRow, COL_执行时间)
            rsCurr!频率 = vsAdvice.TextMatrix(lngRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngRow, COL_间隔单位)
            rsCurr!开始时间 = vsAdvice.Cell(flexcpData, lngRow, COL_开始时间)
            rsCurr!开嘱医生 = vsAdvice.TextMatrix(lngRow, COL_开嘱医生)
            rsCurr!开嘱科室ID = Val(vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID))
            rsCurr!开嘱时间 = vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!标志 = vsAdvice.TextMatrix(lngRow, COL_标志)
            '修改了检验组合内容,采集方法行应标记为修改
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '完全重新设置该检验组合
            '------------------------
            '删除检验项目行:删除之后重新定位的当前行
            lngRow = Delete检验组合(lngRow)
            '清除当前行(采集方法行)
            Call DeleteRow(lngRow, True, False)
            '重新产生:产生之后重新定位的当前行
            lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData, rsCurr)
            
            '强行显示当前医嘱卡片
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        ElseIf RowIn配方行(lngRow) Then
            '中药配方
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng诊疗项目ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_诊疗项目ID))
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行性质", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "开始时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "终止时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱医生", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "开嘱时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "标志", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!执行性质 = NeedName(cbo执行性质.Text) '正常,自备药,离院带药
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID)) <> 0 Then
                rsCurr!执行科室ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID))
            End If
            rsCurr!频率 = vsAdvice.TextMatrix(lngDrugRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngDrugRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngDrugRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngDrugRow, COL_间隔单位)
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_总量)) <> 0 Then
                rsCurr!总量 = Val(vsAdvice.TextMatrix(lngDrugRow, COL_总量))
            End If
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngDrugRow, COL_执行时间)
            rsCurr!开始时间 = vsAdvice.Cell(flexcpData, lngDrugRow, COL_开始时间)
            rsCurr!终止时间 = vsAdvice.Cell(flexcpData, lngDrugRow, COL_终止时间)
            rsCurr!开嘱医生 = vsAdvice.TextMatrix(lngDrugRow, COL_开嘱医生)
            rsCurr!开嘱科室ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_开嘱科室ID))
            rsCurr!开嘱时间 = vsAdvice.Cell(flexcpData, lngDrugRow, COL_开嘱时间)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!标志 = vsAdvice.TextMatrix(lngRow, COL_标志)
            '修改了配方内容,用法行应标记为修改
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '完全重新设置该中药配方行
            '------------------------
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '清除当前行(中药用法行)
            Call DeleteRow(lngRow, True, False)
            '产生配方:产生之后重新定位的当前行
            lngRow = AdviceSet中药配方(lng诊疗项目ID, lngRow, lng用法ID, strExtData, rsCurr)
            
            '强行显示当前医嘱卡片
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        End If
        
        If InStr(",0,3,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
            vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
            Call ReSetColor(lngRow)
        End If
        
        mblnNoSave = True '标记为未保存
    End If
    
    Call vsAdvice.AutoSize(COL_医嘱内容)
    
    '对保险对码进行检查
    strMsg = CheckAdviceInsure(vsAdvice.Row)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    
    txt医嘱内容.SetFocus
End Sub

Private Sub cmdPati_Click()
    lvwPati.ListItems("_" & mlng病人ID & "_" & mlng主页ID).Selected = True
    lvwPati.SelectedItem.EnsureVisible
    lvwPati.Left = txtPati.Left + fraPati.Left
    lvwPati.Top = txtPati.Top + txtPati.Height + fraPati.Top
    lvwPati.Height = vsAdvice.Height - 300
    lvwPati.Visible = True
    lvwPati.SetFocus
End Sub

Private Sub ClinicSelecter(Optional ByVal lng分类ID As Long)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = frmClinicSelect.ShowSelect(Me, mstrPrivs, cbo期效.ListIndex, mstr性别, , , , lng分类ID)
    If rsTmp Is Nothing Then '取消或无数据
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus: Exit Sub
    End If
        
    '根据选择项目设置缺省医嘱信息
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '显示已缺省设置的值
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        txt医嘱内容.SetFocus '必须先定位
        Call SeekNextControl
    Else
        '恢复原值(AdviceInput函数中可能处理了一下)
        txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
        txt医嘱内容.SetFocus
    End If
End Sub

Private Sub cmdSel_Click()
    ClinicSelecter
End Sub

Private Sub cmd开始时间_Click()
    If IsDate(txt开始时间.Text) Then
        dtpDate.Value = CDate(txt开始时间.Text)
    Else
        dtpDate.Value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "开始时间"
    dtpDate.Left = txt开始时间.Left + fraAdvice.Left
    dtpDate.Top = txt开始时间.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub cmd开嘱时间_Click()
    If IsDate(txt开嘱时间.Text) Then
        dtpDate.Value = CDate(txt开嘱时间.Text)
    Else
        dtpDate.Value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "开嘱时间"
    dtpDate.Left = txt开嘱时间.Left + txt开嘱时间.Width - dtpDate.Width + fraAdvice.Left
    dtpDate.Top = txt开嘱时间.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub cmd终止时间_Click()
    If IsDate(txt终止时间.Text) Then
        dtpDate.Value = CDate(txt终止时间.Text)
    Else
        dtpDate.Value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "终止时间"
    dtpDate.Left = txt终止时间.Left + txt终止时间.Width - dtpDate.Width + fraAdvice.Left
    dtpDate.Top = txt终止时间.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String, bln补录 As Boolean
    
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_标志) <> "" Then
        bln补录 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 2
    Else
        bln补录 = mbln补录
    End If
    
    If dtpDate.Tag = "开始时间" Then
        '取值
        If IsDate(txt开始时间.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt开始时间.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check开始时间(strDate, txt终止时间.Text, bln补录) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt开始时间.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt开始时间_Validate(False) '更新数据
        txt开始时间.SetFocus
    ElseIf dtpDate.Tag = "终止时间" Then
        '取值
        If IsDate(txt终止时间.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt开始时间.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
                
        '判断时间合法性
        If Not Check终止时间(txt开嘱时间.Text, txt开始时间.Text, strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt终止时间.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt终止时间_Validate(False) '更新数据
        txt终止时间.SetFocus
    ElseIf dtpDate.Tag = "开嘱时间" Then
        '取值
        If IsDate(txt开嘱时间.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt开嘱时间.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
                
        '判断时间合法性
        If Not Check开嘱时间(strDate, vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间), bln补录) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt开嘱时间.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt开嘱时间_Validate(False) '更新数据
        txt开嘱时间.SetFocus
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpDate_DateClick(dtpDate.Value)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    dtpDate.Visible = False
    If dtpDate.Tag = "终止时间" And ActiveControl Is txt开始时间 Then
        If txt终止时间.Enabled And txt终止时间.Visible Then txt终止时间.SetFocus
    End If
    dtpDate.Tag = ""
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If mlng医嘱ID = 0 Then
            If cbo期效.Enabled Then cbo期效.SetFocus  '新增
        Else
            If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus '修改
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        If KeyCode = vbKeyX Then
            If tbr.Buttons("退出").Enabled And tbr.Buttons("退出").Visible Then
                Call tbr_ButtonClick(tbr.Buttons("退出"))
            End If
        ElseIf Between(Chr(KeyCode), "1", "6") Then
            Call mfrmShortCut.ShowShortCut(Val(Chr(KeyCode)))
        End If
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA
                If tbr.Buttons("增加").Enabled And tbr.Buttons("增加").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("增加"))
                End If
            Case vbKeyI
                If tbr.Buttons("插入").Enabled And tbr.Buttons("插入").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("插入"))
                End If
            Case vbKeyW
                If tbr.Buttons("停止").Enabled And tbr.Buttons("停止").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("停止"))
                End If
            Case vbKeyK
                If tbr.Buttons("一并").Enabled And tbr.Buttons("一并").Visible Then
                    tbr.Buttons("一并").Value = IIF(tbr.Buttons("一并").Value = tbrPressed, tbrUnpressed, tbrPressed)
                    Call tbr_ButtonClick(tbr.Buttons("一并"))
                End If
            Case vbKeyR
                If tbr.Buttons("申请").Enabled And tbr.Buttons("申请").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("申请"))
                End If
            Case vbKeyY
                If tbr.Buttons("复制").Enabled And tbr.Buttons("复制").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("复制"))
                End If
            Case vbKeyT
                If tbr.Buttons("成套").Visible And tbr.Buttons("成套").Enabled Then
                    Call tbr_ButtonClick(tbr.Buttons("成套"))
                End If
            Case vbKeyS
                If tbr.Buttons("保存").Enabled And tbr.Buttons("保存").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("保存"))
                End If
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                If dtpDate.Visible Then
                    dtpDate.Visible = False
                    If dtpDate.Tag = "终止时间" Then
                        If txt终止时间.Visible And txt终止时间.Enabled Then txt终止时间.SetFocus
                    End If
                    dtpDate.Tag = ""
                ElseIf lvwPati.Visible Then
                    lvwPati.Visible = False
                End If
            Case vbKeyF3
                If tbrFree.Buttons(1).Enabled And tbrFree.Buttons(1).Visible Then
                    tbrFree.Buttons(1).Value = IIF(tbrFree.Buttons(1).Value = 1, 0, 1)
                    Call tbrFree_ButtonClick(tbrFree.Buttons(1))
                End If
            Case vbKeyF4
                If Me.ActiveControl Is txt开始时间 Then
                    If cmd开始时间.Visible And cmd开始时间.Enabled Then cmd开始时间_Click
                ElseIf Me.ActiveControl Is txt终止时间 Then
                    If cmd终止时间.Enabled And cmd终止时间.Visible Then cmd终止时间_Click
                ElseIf Me.ActiveControl Is txt开嘱时间 Then
                    If cmd开嘱时间.Enabled And cmd开嘱时间.Visible Then cmd开嘱时间_Click
                ElseIf Me.ActiveControl Is txt医嘱内容 Then
                    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
                ElseIf Me.ActiveControl Is txt用法 Then
                    If cmd用法.Visible And cmd用法.Enabled Then cmd用法_Click
                ElseIf Me.ActiveControl Is txt频率 Then
                    If cmd频率.Visible And cmd频率.Enabled Then cmd频率_Click
                ElseIf Me.ActiveControl Is txtPati Then
                    If cmdPati.Visible And cmdPati.Enabled Then cmdPati_Click
                End If
            Case vbKeyF1
                Call tbr_ButtonClick(tbr.Buttons("帮助"))
            Case vbKeyF2
                If tbr.Buttons("保存").Enabled And tbr.Buttons("保存").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("保存"))
                End If
            Case vbKeyF6
                If tbr.Buttons("参考").Visible And tbr.Buttons("参考").Enabled Then
                    Call tbr_ButtonClick(tbr.Buttons("参考"))
                End If
            Case vbKeyF7 '切换输入法
                If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
                    If stbThis.Panels("WB").Bevel = sbrRaised Then
                        Call stbThis_PanelClick(stbThis.Panels("WB"))
                    Else
                        Call stbThis_PanelClick(stbThis.Panels("PY"))
                    End If
                End If
            Case vbKeyF8 '切换显示计价项目
                If stbThis.Panels("Price").Visible Then
                    Call stbThis_PanelClick(stbThis.Panels("Price"))
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        Call mfrmShortCut.ShowMe(Me)
    End If
End Sub

Private Sub Form_Load()
    Dim lngRow As Long
    
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    Call zlControl.CboSetHeight(cbo执行科室, Me.Height)
    Call zlControl.CboSetWidth(cbo执行科室.Hwnd, cbo执行科室.Width * 1.3)
    
    mblnOK = False
    mblnNoSave = False
    mblnRunFirst = True
    mblnRowChange = True
    mblnDoCheck = True
    mstrDelIDs = ""
    
    '病人过敏史/病生状态可用检测
    mlngPassPati = 0
    If gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 And Not mbln护士站 Then  'Pass
        cmdAlley.Visible = True
        vsAdvice.ColHidden(COL_警示) = False
        cmdAlley.Enabled = PassGetState("AlleyEnable") = 1
    End If
    
    '医生输入允许状态
    If Not mbln护士站 Then cbo医生.Locked = True
    
    '权限设置
    If InStr(mstrPrivs, "诊疗参考") = 0 And mlng前提ID = 0 Then
        tbr.Buttons("参考").Visible = False
        tbr.Buttons("参考_").Visible = False
    End If
'    If InStr(mstrPrivs, "保存成套方案") = 0 Then
'        tbr.Buttons("成套").Visible = False
'    End If
    If mbln护士站 And InStr(mstrPrivs, "医嘱停止") = 0 Then
        tbr.Buttons("停止").Visible = False
    End If
    
    '电子签名功能
    If gobjESign Is Nothing Or mbln护士站 Then
        tbr.Buttons("签名").Visible = False
    End If
    
    '输入匹配
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    '简码匹配方式：0-拼音,1-五笔
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0))
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
    
    '计价面板状态
    If mblnModal Then
        stbThis.Panels("Price").Visible = False
    Else
        Set mfrmPrice = New frmAdvicePrice
        stbThis.Panels("Price").Tag = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "PricePaneVisible", "")
    End If
    
    '临嘱缺省一次性
    mbln一次性 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "临嘱缺省一次性", 0)) <> 0
    
    '执行天数
    mbln天数 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "医嘱执行天数", 0)) <> 0
    
    '检查出院诊断
    mbln检查诊断 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "要求输入出院诊断", 0)) <> 0 And mlng前提ID = 0
    
    '自动进行校对
    mbln校对 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "自动完成校对计价", 0)) <> 0 And mbln护士站
    
    '药品出库检查方式:这里暂时没用
    Set mcolStock = InitStockCheck(2)
    
    '常用嘱托
    Call ReadEnjoin
    '医嘱内容定义
    Call InitAdviceDefine
    '--------------------------------------------
    '读取病人清单
    Call LoadPatients
    
    '修改时强行定位婴儿
    If mlng医嘱ID = 0 Then '新增
        cbo婴儿.ListIndex = 0 '缺省新增病人的医嘱
    Else '修改
        cbo婴儿.ListIndex = mint婴儿
    End If
    cbo婴儿.Tag = cbo婴儿.ListIndex
    
    '读取并显示病人医嘱
    Call ReLoadAdvice(mlng医嘱ID)
    
    '医嘱审核设置
    If mbln审核 Then
        '没有需要审核的医嘱
        If vsAdvice.Rows = vsAdvice.FixedRows + 1 And vsAdvice.RowData(vsAdvice.FixedRows) = 0 Then
            MsgBox "该病人没有需要审核的医嘱。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        Me.Caption = "住院医嘱审核"
        tbr.Buttons("保存").Caption = "审核"
        tbr.Buttons("保存").ToolTipText = "完成审核并退出(F2)"
        tbr.Buttons("停止").Visible = False
        vsAdvice.BackColor = &HC0FFFF
        vsAdvice.BackColorBkg = &HC0FFFF
    ElseIf Not mbln护士站 Then
        If Not HaveAuditPriv Then '不具有资格的医生下医嘱
            tbr.Buttons("停止").Visible = False
            tbr.Buttons("签名").Visible = False
        End If
    End If
    
    '处理快捷输入窗体
    Set mfrmShortCut = New frmClinicShortCut
    mfrmShortCut.ShowMe Me, True '根据上次上否显示
End Sub

Private Function GetStockCheck(ByVal lng库房ID As Long) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng库房ID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Sub InitAdviceDefine()
'功能：初始化医嘱内容定义相关内容
'说明：当mrsDefine不为Nothing时，可以正常使用
    Dim strSQL As String
    
    On Error Resume Next
    Set mobjVBA = CreateObject("ScriptControl")
    Err.Clear: On Error GoTo 0
    
    If Not mobjVBA Is Nothing Then
        mobjVBA.Language = "VBScript"
        Set mobjScript = New clsScript
        mobjVBA.AddObject "clsScript", mobjScript, True
        
        On Error GoTo errH
        strSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
        Set mrsDefine = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDefine, strSQL, Me.Caption)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsDefine = Nothing
End Sub

Private Sub ReLoadAdvice(Optional ByVal lng医嘱ID As Long)
'功能：重新读取并显示病人的当前医嘱清单
'参数：lng医嘱ID=用于定位
    Dim lngRow As Long
    
    If LoadAdvice Then
        '显示医嘱
        Call ShowAdvice
        
        If lng医嘱ID = 0 Then
            If vsAdvice.RowData(vsAdvice.Row) <> 0 And Not mbln审核 Then
                Call tbr_ButtonClick(tbr.Buttons("增加"))
            End If
        Else
            '修改的医嘱ID应该是显示行
            lngRow = vsAdvice.FindRow(lng医嘱ID)
            If lngRow <> -1 Then
                If Not vsAdvice.RowHidden(lngRow) Then
                    mblnRowChange = False
                    vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngRow
                    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
                    mblnRowChange = True
                End If
            End If
        End If
        '进入时屏蔽了ShowAdvice中的调用,强行进入
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Function ReadEnjoin() As Boolean
'功能：读取并加入常用嘱托
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPre As String
        
    On Error GoTo errH
    
    strPre = cbo医生嘱托.Text '加入后保持原有值
    cbo医生嘱托.Clear
    
    strSQL = "Select Upper(编码) as 编码,名称,Upper(名称) as 大写名,Upper(简码) as 简码 From 常用嘱托 Where 名称 is Not Null Order by 名称"
    rsTmp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        AddComboItem cbo医生嘱托.Hwnd, CB_ADDSTRING, 0, rsTmp!名称
        rsTmp.MoveNext
    Loop
    cbo医生嘱托.Text = strPre
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If dtpDate.Visible Then
        dtpDate.Visible = False
        dtpDate.Tag = ""
    ElseIf lvwPati.Visible Then
        lvwPati.Visible = False
    End If
    
    On Error Resume Next
    
    fraPati.Left = 0
    fraPati.Top = cbr.Height
    fraPati.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = cbr.Height + fraPati.Height
    vsAdvice.Height = Me.ScaleHeight - fraPati.Height - cbr.Height - stbThis.Height - (fraAdvice.Height - 80)
    vsAdvice.Width = Me.ScaleWidth
    
    fraAdvice.Left = 0
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 80
    fraAdvice.Width = Me.ScaleWidth
    
    'Pass
    cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 30
    cbo婴儿.Left = Me.ScaleWidth - IIF(cmdAlley.Visible, cmdAlley.Width + 30, 0) - cbo婴儿.Width - 30
    lbl婴儿.Left = cbo婴儿.Left - lbl婴儿.Width - 30
    
    If cmdAlley.Visible Or lbl婴儿.Visible Then
        lblPati.Width = IIF(lbl婴儿.Visible, lbl婴儿.Left, cmdAlley.Left) - lblPati.Left - 90
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msng天数 = 0
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    
    '计价面板状态
    If Not mfrmPrice Is Nothing Then
        Unload mfrmPrice
        Set mfrmPrice = Nothing
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "PricePaneVisible", stbThis.Panels("Price").Tag
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If mblnClickItem Then Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnClickItem = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            If Val(lvwPati.SelectedItem.Tag) = mlng病人ID _
                And Val(lvwPati.SelectedItem.ListSubItems(1).Tag) = mlng主页ID Then
                lvwPati.Visible = False
                txtPati.Text = lvwPati.SelectedItem.Text
                txtPati.ForeColor = lvwPati.SelectedItem.ForeColor
                vsAdvice.SetFocus: Exit Sub
            End If
            
            '已经转出的病人是不允许下医嘱的
            If Val(lvwPati.SelectedItem.ListSubItems(9).Tag) = 1 Then
                MsgBox "该病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mblnNoSave Then
                If MsgBox("当前病人的医嘱编辑后尚未保存，确实要更换病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            
            With lvwPati.SelectedItem
                mlng病人ID = Val(.Tag)
                mlng主页ID = Val(.ListSubItems(1).Tag)
                mlng病人科室id = Val(.ListSubItems(2).Tag)
                mint险类 = Val(.ListSubItems(3).Tag)
                mstr付款码 = Get医疗付款码(.SubItems(10))
                mstr性别 = .SubItems(4)
                txtPati.Text = .Text
                txtPati.ForeColor = .ForeColor
                lblPati.Caption = "住院号:" & .SubItems(1) & "　床号:" & Trim(.SubItems(2)) & _
                    "　性别:" & .SubItems(4) & "　年龄:" & .SubItems(5) & _
                    "　费别:" & .SubItems(6) & "  付款方式:" & .SubItems(10)
                mbln提醒对码 = True
            End With
            mint年龄 = GetPatiYear(mlng病人ID)
            Call Show费用信息
            
            lvwPati.Visible = False
            mblnNoSave = False
            mstrDelIDs = ""
            
            '清除当前医嘱内容
            mblnRowChange = False
            With vsAdvice
                .Redraw = flexRDNone
                Call .Select(.FixedRows, 0, .Rows - 1, .Cols - 1)
                Call .Clear(flexClearSelection)
                .Rows = .FixedRows + 1 '相当于删除行
                
                'Clear不会清除RowData,RowHidden,需要额外处理
                .RowData(.FixedRows) = Empty
                .RowHidden(.FixedRows) = False
                                
                .Row = .FixedRows: .Col = .FixedCols
                .Redraw = flexRDDirect
            End With
            mblnRowChange = True
            
            '产科才有婴儿医嘱
            Call SetBabyVisible(mlng病人科室id)
            
            '读取并显示病人医嘱
            If LoadAdvice Then Call ShowAdvice
            
            vsAdvice.SetFocus
            
            '特殊医嘱提醒
            strMsg = ExistsSpecAdvice(mlng病人ID, mlng主页ID)
            If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnClickItem = False
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'功能：判断两行是否可以一并给药
'参数：lngRow1=前面一条已经输入的药品行
'      lngRow2=当前行(已输入或未输入)
'返回：如果不可以，则strMsg返回提示信息
    Dim lngFind As Long
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '必须全部为成药且类别相同
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
            
            '不能包含已校对的医嘱
            If InStr(",1,2,", .TextMatrix(lngRow1, COL_状态)) = 0 Then
                strMsg = "要设置为一并给药的药品包含已经校对的医嘱。"
                Exit Function
            End If
                
            '补录医嘱时不能设置非补录的其它医嘱为一并给药
            If (mbln补录 And Val(.TextMatrix(lngRow1, COL_标志)) <> 2) _
                Or (Not mbln补录 And Val(.TextMatrix(lngRow1, COL_标志)) = 2) Then
                strMsg = "不能将补录的医嘱与其它医嘱一起设置为一并给药。"
                Exit Function
            End If
            '不能包含已签名的医嘱
            If Val(.TextMatrix(lngRow1, COL_签名否)) = 1 Then
                strMsg = "要设置为一并给药的药品包含已经签名的医嘱。"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
'            '必须全部为成药且类别相同
'            If Not (.TextMatrix(lngRow1, COL_类别) = .TextMatrix(lngRow2, COL_类别) _
'                And InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) > 0) Then
'                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
'                Exit Function
'            End If
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
            
            '期效必须相同
            If .TextMatrix(lngRow1, COL_期效) <> .TextMatrix(lngRow2, COL_期效) Then
                strMsg = "一并给药的药品医嘱期效必须相同。"
                Exit Function
            End If
            
            '不能包含已校对的医嘱
            If InStr(",1,2,", .TextMatrix(lngRow1, COL_状态)) = 0 _
                Or InStr(",1,2,", .TextMatrix(lngRow2, COL_状态)) = 0 Then
                strMsg = "要设置为一并给药的药品包含已经校对的医嘱。"
                Exit Function
            End If
                
            '不能将补录的医嘱与其它医嘱一起设置为一并给药
            If Val(.TextMatrix(lngRow1, COL_标志)) = 2 And Val(.TextMatrix(lngRow2, COL_标志)) <> 2 _
                Or Val(.TextMatrix(lngRow1, COL_标志)) <> 2 And Val(.TextMatrix(lngRow2, COL_标志)) = 2 Then
                strMsg = "不能将补录的医嘱与其它医嘱一起设置为一并给药。"
                Exit Function
            End If
            
            '补录医嘱时不能设置非补录的其它医嘱为一并给药
            If mbln补录 And (Val(.TextMatrix(lngRow1, COL_标志)) <> 2 Or Val(.TextMatrix(lngRow2, COL_标志)) <> 2) Then
                strMsg = "补录医嘱时不能设置非补录的其它医嘱为一并给药。"
                Exit Function
            End If
            
            '不能包含已签名的医嘱
            If Val(.TextMatrix(lngRow1, COL_签名否)) = 1 Or Val(.TextMatrix(lngRow2, COL_签名否)) = 1 Then
                strMsg = "要设置为一并给药的药品包含已经签名的医嘱。"
                Exit Function
            End If
            
            '一并给药(前面药品)的给药途径是否适用于当前药品
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_相关ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check适用用法(Val(.TextMatrix(lngFind, COL_诊疗项目ID)), Val(.TextMatrix(lngRow2, COL_诊疗项目ID)), 2) Then
                    strMsg = """" & .TextMatrix(lngRow2, COL_医嘱内容) & """不能使用""" & .TextMatrix(lngFind, COL_医嘱内容) & """给药途径，" & _
                    vbCrLf & "不能与""" & .TextMatrix(lngRow1, COL_医嘱内容) & """设置为一并给药。"
                    Exit Function
                End If
            End If
        End If
    End With
    RowCanMerge = True
End Function

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng医嘱ID As Long, lng相关ID As Long
    Dim str类别 As String, str补录 As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim str期效 As String, lng诊疗项目ID As Long
    Dim lngTmp As Long, i As Long, j As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim blnMoved As Boolean, strAlter As String
    
    Call AdviceChange '强制更新医嘱内容
    
    With vsAdvice
        Select Case Button.Key
            Case "增加"
                If .RowData(.Row) = 0 Then
'                    If .Row <> .Rows - 1 Then
'                        MsgBox "当前行无内容，请先在当前行录入有效医嘱或删除当前行。", vbInformation, gstrSysName
'                    Else
'                        MsgBox "当前行无内容，请先在当前行录入有效医嘱。", vbInformation, gstrSysName
'                    End If
'                    Exit Sub
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '先删除中间间隔的空行
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                
                Call .ShowCell(.Row, .Col)
                If Visible And cbo期效.Enabled Then cbo期效.SetFocus
            Case "插入"
                If .RowData(.Row) = 0 Then
                    MsgBox "当前行无内容，请先在当前行录入有效医嘱。", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '插入后成自动成为一并给药:插入在一并给药的中间才行
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                        
                        '不能在已校对的一并给药中插入
                        If InStr(",1,2,", .TextMatrix(.Row, COL_状态)) = 0 Then
                            MsgBox "该组一并给药的医嘱已经过校对，不能再插入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        '补录医嘱时,不允许在非补录的一并给药中插入补录医嘱；反之亦然
                        If (mbln补录 And Val(.TextMatrix(.Row, COL_标志)) <> 2) _
                            Or (Not mbln补录 And Val(.TextMatrix(.Row, COL_标志)) = 2) Then
                            MsgBox "不能将补录的医嘱与其它医嘱一起设置为一并给药。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        '不能在已签名的一并给药中插入
                        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                            MsgBox "该组一并给药的医嘱已经签名，不能再插入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        lng相关ID = Val(.TextMatrix(lngPreRow, COL_相关ID))
                    End If
                End If
                
                '先删除中间间隔的空行
                mblnRowChange = False
                lng医嘱ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lng医嘱ID)
                mblnRowChange = True
                            
                '当前行之前插入新行
                '--------------------------------------------------------------
                If RowIn配方行(.Row) Or RowIn检验行(.Row) Then
                    '中药配方及检验组合行是前面的行隐藏
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                cbo期效.SetFocus '先定位避免光标晃
            Case "一并" '一并给药
                If Button.Value = tbrPressed Then
                    lngBegin = GetPreRow(.Row)
                    '前面没有行
                    If lngBegin = -1 Then
                        MsgBox "前面没有可以一并给药的医嘱行。", vbInformation, gstrSysName
                        Button.Value = tbrUnpressed: Exit Sub
                    End If
                    '两行不符合条件
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Button.Value = tbrUnpressed: Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '当前行尚未输入内容的情况
                        cbo期效.ListIndex = IIF(.TextMatrix(lngBegin, COL_期效) = "临嘱", 1, 0)
                        If mbln补录 Then
                            txt开始时间.Text = .Cell(flexcpData, lngBegin, COL_开始时间)
                        ElseIf DateDiff("n", CDate(.Cell(flexcpData, lngBegin, COL_开始时间)), zlDatabase.Currentdate) <= TIME_LIMIT Then
                            txt开始时间.Text = .Cell(flexcpData, lngBegin, COL_开始时间)
                        End If
                        txt医嘱内容.SetFocus: Exit Sub
                    Else
                        '要把当前行与前面行一起一并给药
                        Call MergeRow(lngBegin, .Row, False)
                        Call ReSetColor(.Row) '一并之后再一并设置
                    End If
                Else
                    If .RowData(.Row) = 0 Then
                        '当前行尚未输入内容的情况
                        If RowIn一并给药(.Row) Then Button.Value = tbrPressed
                        Exit Sub
                    Else
                        '当前行是一并给药中的行
                        Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                                                
                        '先判断可否取消一并给药
                        '不能包含已校对的医嘱
                        If InStr(",1,2,", .TextMatrix(.Row, COL_状态)) = 0 Then
                            MsgBox "当前医嘱已经过校对。", vbInformation, gstrSysName
                            Button.Value = tbrPressed: Exit Sub
                        End If
                        '补录医嘱时不能取消非补录的一并给药药品
                        If mbln补录 And Val(.TextMatrix(.Row, COL_标志)) <> 2 Then
                            MsgBox "补录医嘱时不能设置非补录的药品医嘱。", vbInformation, gstrSysName
                            Button.Value = tbrPressed: Exit Sub
                        End If
                                                
                        '不能包含已签名的医嘱
                        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                            MsgBox "当前医嘱已经签名。", vbInformation, gstrSysName
                            Button.Value = tbrPressed: Exit Sub
                        End If

                        '先提示
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '整个一并给药取消为单独给药
                            If MsgBox("要将该组一并给药的药品全部取消为单独给药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Button.Value = tbrPressed: Exit Sub
                            End If
                        End If
                        
                        '删除中间的空行
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '从一并给药中分离该行
                            Call ReSetColor(.Row) '在取消之前一并设置
                            Call SplitRow(.Row)
                        Else
                            '取消一并给药
                            Call ReSetColor(.Row) '在取消之前一并设置
                            lngTmp = .RowData(.Row) '记录用于恢复行定位
                            Call AdviceSet单独给药(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case "删除"
                If .RowSel <> .Row Then
                    MsgBox "一次只能删除一条医嘱，请选择要删除的医嘱行。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    '已校对的医嘱不能删除
                    If InStr(",1,2,", .TextMatrix(.Row, COL_状态)) = 0 Then
                        MsgBox "该条医嘱已经过校对，不能删除。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    '补录医嘱时不能删除非补录的医嘱
                    If mbln补录 And Val(.TextMatrix(.Row, COL_标志)) <> 2 Then
                        MsgBox "补录医嘱时不能删除其它医嘱。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    '已签名的医嘱不能删除
                    If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                        MsgBox "该条医嘱已经签名，不能删除。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                '删除当前行
                Call AdviceDelete(.Row)
                .SetFocus
            Case "停止"
                Call AdviceStop
            Case "参考"
                If Val(.TextMatrix(.Row, COL_诊疗项目ID)) <> 0 Then
                    If RowIn配方行(.Row) Or RowIn检验行(.Row) Then
                        i = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                        If i <> -1 Then
                            lng诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                        End If
                    Else
                        lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                    End If
                End If
                Call ShowClinicHelp(IIF(mblnModal, 1, 0), Me, lng诊疗项目ID)
            Case "复制"
                lng病人ID = mlng病人ID: lng主页ID = mlng主页ID: blnMoved = False
                strMsg = frmAdviceCopy.ShowMe(Me, mstrPrivs, lng病人ID, lng主页ID, blnMoved, mbln护士站, mlng前提ID, strAlter)
                If strMsg <> "" Then
                    Call tbr_ButtonClick(tbr.Buttons("增加"))
                    Call AdviceSet复制医嘱(lng病人ID, lng主页ID, strMsg, blnMoved, strAlter)
                End If
            Case "成套"
                '病人选择时已限制了住院数据未转出
                Call frmAdviceScheme.ShowMe(mstrPrivs, 2, mlng病人ID, mlng主页ID, "", cbo婴儿.ListIndex, Me)
            Case "保存"
                If Not CheckAdvice Then Exit Sub '检查中处理了光标定位
                If Not SaveAdvice Then .SetFocus: Exit Sub
                If mbln审核 Then Unload Me: Exit Sub
            Case "签名"
                Call AdviceSign
            Case "帮助"
                ShowHelp App.ProductName, Me.Hwnd, Me.Name
            Case "退出"
                Unload Me
        End Select
    End With
End Sub

Private Sub AdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名ID As Long, lng证书ID As Long
    Dim intRule As Integer
    
    If gobjESign Is Nothing Then Exit Sub
    
    '自动保存
    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mlng主页ID, strIDs, 0, False, strSource, mlng前提ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
    If strSign <> "" Then
        lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
        strSQL = "zl_医嘱签名记录_Insert(" & lng签名ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        '重新读取显示医嘱
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        mblnOK = True
        If txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        Else
            vsAdvice.SetFocus
        End If

        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceStop() As Boolean
'功能：当前医嘱停止
    Dim strSQL As String, lng医嘱ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngCur As Long, i As Long
    Dim strStopTime As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = .RowData(.Row)
        End If
        If RowIn一并给药(.Row) Then
            If MsgBox("该组一并给药的医嘱将会一起停止，确实要停止吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("确实要停止医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        '停嘱时缺省的医嘱终止时间
        If .TextMatrix(.Row, COL_终止时间) = "" Then
            If gbln长期医嘱次日生效 Then
                strStopTime = "To_Date('" & Format(zlDatabase.Currentdate + 1, "yyyy-MM-dd 00:00") & "','YYYY-MM-DD HH24:MI')"
            Else
                strStopTime = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            End If
        Else
            strStopTime = "To_Date('" & .Cell(flexcpData, .Row, COL_终止时间) & "','YYYY-MM-DD HH24:MI')"
        End If
        
        '医生在当前时间停止
        strSQL = "ZL_病人医嘱记录_停止(" & lng医嘱ID & "," & strStopTime & ",'" & UserInfo.姓名 & "')"
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    '删除已停止的医嘱显示
    With vsAdvice
        '要删除的行范围
        lngBegin = .Row: lngEnd = .Row
        For i = .Row To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = lng医嘱ID Or .RowData(i) = lng医嘱ID Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = .Row + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = lng医嘱ID Or .RowData(i) = lng医嘱ID Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        '记录下一条有效显示的医嘱
        lngCur = GetNextRow(lngEnd)
        If lngCur <> -1 Then lngCur = .RowData(lngCur)
        
        mblnRowChange = False
        '反序删除行
        For i = lngEnd To lngBegin Step -1
            .RemoveItem i
        Next
        
        '重新定位新的行
        If lngCur <> -1 Then
            lngCur = .FindRow(lngCur)
            .Row = lngCur
        Else
            Call tbr_ButtonClick(tbr.Buttons("增加"))
        End If
        .Col = COL_医嘱内容
        Call .ShowCell(.Row, .Col)
        
        mblnRowChange = True
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    If txt医嘱内容.Enabled Then
        txt医嘱内容.SetFocus
    Else
        vsAdvice.SetFocus
    End If
    
    mblnOK = True
    AdviceStop = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Get一并给药范围(ByVal lng相关ID As Long, lngBegin As Long, lngEnd As Long)
'功能：根据相关的给药途径医嘱ID,确定一并给药的一组药品的起止行号
'说明：中间可能包含有空行
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng相关ID), , COL_相关ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txtPati_GotFocus()
    zlControl.TxtSelAll txtPati
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '根据输入定位病人
        If Not txtPati.Locked And txtPati.Text <> "" Then
            For i = 1 To lvwPati.ListItems.Count
                With lvwPati.ListItems(i)
                    If Left(txtPati.Text, 1) = "-" And IsNumeric(Mid(txtPati.Text, 2)) Then '-病人ID
                        If .Tag <> "" And Val(.Tag) = Val(Mid(txtPati.Text, 2)) Then Exit For
                    ElseIf Left(txtPati.Text, 1) = "+" And IsNumeric(Mid(txtPati.Text, 2)) Then '+住院号
                        If .SubItems(1) <> "" And Val(.SubItems(1)) = Val(Mid(txtPati.Text, 2)) Then Exit For
                    ElseIf Left(txtPati.Text, 1) = "/" Then  '/床号
                        If Trim(.SubItems(2)) <> "" And UCase(Trim(.SubItems(2))) = UCase(Mid(txtPati.Text, 2)) Then Exit For
                    Else '当作姓名
                        If .Text <> "" And UCase(.Text) Like "*" & UCase(txtPati.Text) & "*" Then Exit For
                    End If
                End With
            Next
            If i <= lvwPati.ListItems.Count Then
                lvwPati.ListItems(i).EnsureVisible
                lvwPati.ListItems(i).Selected = True
                Call lvwPati_KeyPress(13)
            Else
                MsgBox "没有找到指定的病人，请重新输入。", vbInformation, gstrSysName
                txtPati.Text = lvwPati.SelectedItem.Text
                txtPati.ForeColor = lvwPati.SelectedItem.ForeColor
                txtPati.SetFocus: Exit Sub
            End If
        ElseIf txtPati.Locked Then
            Call SeekNextControl
        End If
    ElseIf Not txtPati.Locked Then
        txtPati.ForeColor = Me.ForeColor
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If Not lvwPati.SelectedItem Is Nothing Then
        If Val(lvwPati.SelectedItem.Tag) = mlng病人ID _
            And Val(lvwPati.SelectedItem.ListSubItems(1).Tag) = mlng主页ID Then
            txtPati.Text = lvwPati.SelectedItem.Text
            txtPati.ForeColor = lvwPati.SelectedItem.ForeColor
        End If
    End If
End Sub

Private Sub txt单量_Change()
    txt单量.Tag = "1"
End Sub

Private Sub txt单量_GotFocus()
    zlControl.TxtSelAll txt单量
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt单量.Text) Or (txt单量.Text = "" And cbo期效.ListIndex = 1) Then
            If SeekNextControl Then Call txt单量_Validate(False)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt单量_Validate(Cancel As Boolean)
    Dim strMsg As String, blnTag As Boolean
    Dim dbl次数 As Double, sng天数 As Single
    Dim dbl总量 As Double, dbl总剂量 As Double
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    With vsAdvice
        If Val(txt单量.Text) = 0 Then txt单量.Text = ""
        If Not IsNumeric(txt单量.Text) Then
            If txt单量.Text <> "" Then
                Cancel = True: txt单量_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 And .TextMatrix(.Row, COL_期效) = "长嘱" Then
                '恢复人为的清除
                If IsNumeric(.TextMatrix(.Row, COL_单量)) Then
                    txt单量.Text = .TextMatrix(.Row, COL_单量)
                End If
            End If
        ElseIf CDbl(txt单量.Text) <= 0 Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        ElseIf CDbl(txt单量.Text) > LONG_MAX Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        Else
            '单量合法性检查
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_收费细目ID)) <> 0 Then
                dbl次数 = IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, 1, Val(.TextMatrix(.Row, COL_总量))) * _
                    Val(.TextMatrix(.Row, COL_住院包装)) * Val(.TextMatrix(.Row, COL_剂量系数)) / Val(txt单量.Text)
                If dbl次数 > 200 Then
                    If MsgBox("该药品按每次 " & FormatEx(txt单量.Text, 5) & .TextMatrix(.Row, COL_单量单位) & " 使用，" & _
                        IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, "每", Val(.TextMatrix(.Row, COL_总量))) & _
                        .TextMatrix(.Row, COL_住院单位) & "可以使用 " & FormatEx(dbl次数, 5) & " 次。" & _
                        vbCrLf & vbCrLf & "你确认单量输入正确吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt单量_GotFocus: Exit Sub
                    End If
                End If
            End If
            
            '检查处方限量,录入限量
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_处方限量)) <> 0 And Val(txt单量.Text) <> 0 Then
                If Val(.TextMatrix(.Row, COL_收费细目ID)) = 0 Then
                    '长嘱按品种下达时，只单纯比较单量
                    If Val(txt单量.Text) > Val(.TextMatrix(.Row, COL_处方限量)) Then
                        If MsgBox(.TextMatrix(.Row, COL_名称) & " 的单量:" & FormatEx(txt单量.Text, 5) & lbl单量单位.Caption & " 超过处方限量:" & _
                            FormatEx(Val(.TextMatrix(.Row, COL_处方限量)), 5) & lbl单量单位.Caption & "，你确认单量输入正确吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: txt单量_GotFocus: Exit Sub
                        End If
                    End If
                ElseIf .TextMatrix(.Row, COL_期效) = "长嘱" And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 Then
                    '成药长嘱按规格下达时，按一个频率周期的缺省总量比较
                    dbl总量 = Calc缺省药品总量(Val(txt单量.Text), 1, _
                        Val(.TextMatrix(.Row, COL_频率次数)), _
                        Val(.TextMatrix(.Row, COL_频率间隔)), _
                        .TextMatrix(.Row, COL_间隔单位), _
                        .TextMatrix(.Row, COL_执行时间), _
                        Val(.TextMatrix(.Row, COL_剂量系数)), _
                        Val(.TextMatrix(.Row, COL_住院包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零)))
                    dbl总剂量 = dbl总量 * Val(.TextMatrix(.Row, COL_住院包装)) * Val(.TextMatrix(.Row, COL_剂量系数))
                    If dbl总剂量 > Val(.TextMatrix(.Row, COL_处方限量)) Then
                        If MsgBox(.TextMatrix(.Row, COL_名称) & " 按每次 " & FormatEx(txt单量.Text, 5) & lbl单量单位.Caption & "，" & _
                            .TextMatrix(.Row, COL_频率) & "的剂量:" & FormatEx(dbl总量, 5) & .TextMatrix(.Row, COL_住院单位) & _
                            "(" & FormatEx(dbl总剂量, 5) & lbl单量单位.Caption & ")超过处方限量:" & _
                            FormatEx(Val(.TextMatrix(.Row, COL_处方限量)), 5) & lbl单量单位.Caption & "，要继续吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: txt单量_GotFocus: Exit Sub
                        End If
                    End If
                End If
            ElseIf .TextMatrix(.Row, COL_期效) = "长嘱" And Val(txt单量.Text) <> 0 _
                And InStr(",5,6,7,", .TextMatrix(.Row, COL_类别)) = 0 And Val(.TextMatrix(.Row, COL_处方限量)) > 0 Then
                '按一个频率周期的缺省总量进行检查
                If Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 Then
                    If .TextMatrix(.Row, COL_间隔单位) = "周" Then
                        dbl总量 = Val(txt单量.Text) * Val(.TextMatrix(.Row, COL_频率次数)) / 7
                    ElseIf .TextMatrix(.Row, COL_间隔单位) = "天" Then
                        dbl总量 = Val(txt单量.Text) * Val(.TextMatrix(.Row, COL_频率次数)) / Val(.TextMatrix(.Row, COL_频率间隔))
                    ElseIf .TextMatrix(.Row, COL_间隔单位) = "小时" Then
                        dbl总量 = Val(txt单量.Text) * (Val(.TextMatrix(.Row, COL_频率次数)) / Val(.TextMatrix(.Row, COL_频率间隔))) * 24
                    End If
                Else
                    dbl总量 = Val(txt单量.Text) '持续性长嘱
                End If
                If dbl总量 > Val(.TextMatrix(.Row, COL_处方限量)) Then
                    If MsgBox(.TextMatrix(.Row, COL_名称) & " 按每回 " & FormatEx(txt单量.Text, 5) & lbl单量单位.Caption & "，" & _
                        .TextMatrix(.Row, COL_频率) & "的总量为:" & FormatEx(dbl总量, 5) & lbl单量单位.Caption & _
                        "，超过允许录入的最大限量:" & .TextMatrix(.Row, COL_处方限量) & lbl单量单位.Caption & "，要继续吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt单量_GotFocus: Exit Sub
                    End If
                End If
            End If
            
            txt单量.Text = FormatEx(txt单量.Text, 5)
            
            '重新计算药品总量(先输入单量时)
            If mbln天数 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "临嘱" Then
                If .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 _
                    And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_住院包装)) <> 0 Then
                    
                    sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    If sng天数 = 0 Then sng天数 = 1
                    
                    txt总量.Text = FormatEx(Calc缺省药品总量( _
                        Val(txt单量.Text), sng天数, _
                        Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                        .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_住院包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    txt总量.Tag = "1"
                End If
            End If
        End If
        
        '更新数据
        blnTag = txt单量.Tag <> ""
        Call AdviceChange
        
        '药品库存检查:仅对成药长嘱(其它的用总量),只提醒,按一个频率周期算总量
        If blnTag Then
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "长嘱" Then
                strMsg = CheckStock(.Row)
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub txt开始时间_Change()
    txt开始时间.Tag = "1"
End Sub

Private Sub txt开始时间_GotFocus()
    If txt开始时间.Text = "" Then txt开始时间.Text = GetDefaultTime(vsAdvice.Row)
    zlControl.TxtSelAll txt开始时间
End Sub

Private Sub txt开始时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt开始时间.Text <> "" Then
            txt开始时间.Text = GetFullDate(txt开始时间.Text)
            If SeekNextControl Then Call txt开始时间_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt开始时间_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt开始时间.Locked Then
        glngTXTProc = GetWindowLong(txt开始时间.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt开始时间.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt开始时间_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt开始时间.Locked Then
        Call SetWindowLong(txt开始时间.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt开始时间_Validate(Cancel As Boolean)
    Dim bln补录 As Boolean
            
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    If txt开始时间.Locked Then Exit Sub
        
    If Not IsDate(txt开始时间.Text) Then
        If txt开始时间.Text <> "" Then
            Cancel = True
            txt开始时间_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)) Then
                '恢复人为的清除
                txt开始时间.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)
            End If
        End If
    Else
        '检查时间合法性
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_标志) <> "" Then
            bln补录 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 2
        Else
            bln补录 = mbln补录
        End If
        If Not Check开始时间(txt开始时间.Text, txt终止时间.Text, bln补录) Then
            Cancel = True
            txt开始时间_GotFocus
            Exit Sub
        End If
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo医生嘱托_Change()
    cbo医生嘱托.Tag = "1"
End Sub

Private Sub cbo医生嘱托_Click()
    cbo医生嘱托.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo医生嘱托_GotFocus()
    zlControl.TxtSelAll cbo医生嘱托
End Sub

Private Sub cbo医生嘱托_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo医生嘱托_Validate(False)
    Else
        Call zlControl.CboAppendText(cbo医生嘱托, KeyAscii)
    End If
End Sub

Private Sub cbo医生嘱托_Validate(Cancel As Boolean)
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    If zlCommFun.ActualLen(cbo医生嘱托.Text) > 100 Then
        MsgBox "输入内容不过超过 50 个汉字或 100 个字符。", vbInformation, gstrSysName
        cbo医生嘱托_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt医嘱内容_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt医嘱内容_GotFocus()
    If txt开始时间.Text = "" Then txt开始时间_GotFocus
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt医嘱内容)
    End If
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt医嘱内容.Text = "" Then Exit Sub
        If txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        If tbrFree.Buttons(1).Value = 0 Then
            Set rsTmp = frmClinicSelect.ShowSelect(Me, mstrPrivs, cbo期效.ListIndex, mstr性别, txt医嘱内容.Text, txt医嘱内容)
            If rsTmp Is Nothing Then '取消或无数据
                '恢复原值
                txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
                zlControl.TxtSelAll txt医嘱内容
                txt医嘱内容.SetFocus: Exit Sub
            End If
            '新项目的录入
            '成套项目中如果包含成药,则不能按规格下医嘱
            
            '根据选择项目设置缺省医嘱信息
            Me.Refresh
            If AdviceInput(rsTmp, vsAdvice.Row) Then
                '显示已缺省设置的值
                Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
                Call SeekNextControl
            Else
                '恢复原值
                txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
                zlControl.TxtSelAll txt医嘱内容
                txt医嘱内容.SetFocus: Exit Sub
            End If
        ElseIf tbrFree.Buttons(1).Value = 1 Then
            If txt医嘱内容.Text <> "" Then
                If zlCommFun.ActualLen(txt医嘱内容.Text) > txt医嘱内容.MaxLength Then
                    MsgBox "输入内容不过超过 " & txt医嘱内容.MaxLength \ 2 & " 个汉字或 " & txt医嘱内容.MaxLength & " 个字符。", vbInformation, gstrSysName
                    Call txt医嘱内容_GotFocus: Exit Sub
                End If
                Call AdviceInputFree(vsAdvice.Row)
                Call SeekNextControl
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    End If
End Sub

Private Sub cbo执行时间_GotFocus()
    zlControl.TxtSelAll cbo执行时间
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    If tbrFree.Buttons(1).Value = 0 Then
        '恢复人为的改变
        If txt医嘱内容.Text <> vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容) Then
            txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
        End If
    ElseIf tbrFree.Buttons(1).Value = 1 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 And txt医嘱内容.Text = "" Then
            '因为必须录入,所以自动恢复
            txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
            Exit Sub
        End If
        
        If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
        
        If txt医嘱内容.Text <> "" Then
            If zlCommFun.ActualLen(txt医嘱内容.Text) > txt医嘱内容.MaxLength Then
                MsgBox "输入内容不过超过 " & txt医嘱内容.MaxLength \ 2 & " 个汉字或 " & txt医嘱内容.MaxLength & " 个字符。", vbInformation, gstrSysName
                Call txt医嘱内容_GotFocus: Cancel = True: Exit Sub
            End If
            Call AdviceInputFree(vsAdvice.Row)
        End If
    End If
End Sub

Private Sub txt终止时间_Change()
    txt终止时间.Tag = "1"
End Sub

Private Sub txt终止时间_GotFocus()
    zlControl.TxtSelAll txt终止时间
End Sub

Private Sub txt终止时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt终止时间.Text <> "" Then
            txt终止时间.Text = GetFullDate(txt终止时间.Text)
        End If
        If SeekNextControl Then Call txt终止时间_Validate(False)
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt终止时间_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt终止时间.Locked Then
        glngTXTProc = GetWindowLong(txt终止时间.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt终止时间.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt终止时间_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt终止时间.Locked Then
        Call SetWindowLong(txt终止时间.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt终止时间_Validate(Cancel As Boolean)
    If txt终止时间.Locked Then Exit Sub
    
    If Not IsDate(txt终止时间.Text) And txt终止时间.Text <> "" Then
        Cancel = True
        txt终止时间_GotFocus
        Exit Sub
    ElseIf txt终止时间.Text <> "" Then
        '检查时间合法性
        If Not Check终止时间(txt开嘱时间.Text, txt开始时间.Text, txt终止时间.Text) Then
            Cancel = True
            txt终止时间_GotFocus
            Exit Sub
        End If
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt开嘱时间_Change()
    txt开嘱时间.Tag = "1"
End Sub

Private Sub txt开嘱时间_GotFocus()
    zlControl.TxtSelAll txt开嘱时间
End Sub

Private Sub txt开嘱时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt开嘱时间.Text <> "" Then
            txt开嘱时间.Text = GetFullDate(txt开嘱时间.Text)
        End If
        If SeekNextControl Then Call txt开嘱时间_Validate(False)
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt开嘱时间_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt开嘱时间.Locked Then
        glngTXTProc = GetWindowLong(txt开嘱时间.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt开嘱时间.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt开嘱时间_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt开嘱时间.Locked Then
        Call SetWindowLong(txt开嘱时间.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt开嘱时间_Validate(Cancel As Boolean)
    Dim bln补录 As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    If txt开嘱时间.Locked Then Exit Sub
        
    If Not IsDate(txt开嘱时间.Text) And txt开嘱时间.Text <> "" Then
        Cancel = True
        txt开嘱时间_GotFocus
        Exit Sub
    ElseIf txt开嘱时间.Text <> "" Or vsAdvice.RowData(vsAdvice.Row) <> 0 Then
        '检查时间合法性
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_标志) <> "" Then
            bln补录 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 2
        Else
            bln补录 = mbln补录
        End If
    
        '检查时间合法性
        If Not Check开嘱时间(txt开嘱时间.Text, txt开始时间.Text, bln补录) Then
            Cancel = True
            txt开嘱时间_GotFocus
            Exit Sub
        End If
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt总量_Change()
    txt总量.Tag = "1"
End Sub

Private Sub txt总量_GotFocus()
    zlControl.TxtSelAll txt总量
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt总量.Text) Then
            If SeekNextControl Then Call txt总量_Validate(False)
        End If
    Else
        If RowIn配方行(vsAdvice.Row) Then
            strMask = "0123456789" '中药配方只能输入整数
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 Then
            If InStr(mstrPrivs, "药品小数输入") > 0 Then
                strMask = "0123456789."
            Else
                strMask = "0123456789"
            End If
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt总量_Validate(Cancel As Boolean)
    Dim blnTag As Boolean, strMsg As String
    Dim dbl总量 As Double, sng天数 As Single, bln配方行 As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("退出")) Then Exit Sub
    
    With vsAdvice
        If Val(txt总量.Text) = 0 Then txt总量.Text = ""
        If Not IsNumeric(txt总量.Text) Then
            If txt总量.Text <> "" Then
                Cancel = True: txt总量_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
                '恢复人为的清除
                If IsNumeric(.TextMatrix(.Row, COL_总量)) Then
                    txt总量.Text = .TextMatrix(.Row, COL_总量)
                End If
            End If
        ElseIf CDbl(txt总量.Text) <= 0 Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        ElseIf CDbl(txt总量.Text) > LONG_MAX Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        Else
            txt总量.Text = FormatEx(txt总量.Text, 5)
        End If
        
        bln配方行 = RowIn配方行(.Row)
        
        If IsNumeric(txt总量.Text) Then
            If bln配方行 Then
                txt总量.Text = CInt(txt总量.Text)
            ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                If InStr(mstrPrivs, "药品小数输入") = 0 Then
                    txt总量.Text = Int(txt总量.Text)
                End If
            ElseIf Val(.TextMatrix(.Row, COL_计算方式)) = 3 Then
                '计次项目总量限制为整数。计次项目不输入单量,因此单量不管
                'txt总量.Text = Int(txt总量.Text)
            End If
        End If
        
        '检查总量够否
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "临嘱" Then
            If .TextMatrix(.Row, COL_频率) <> "" _
                And Val(.TextMatrix(.Row, COL_单量)) <> 0 _
                And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                And Val(.TextMatrix(.Row, COL_住院包装)) <> 0 Then
                
                If Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                    dbl总量 = FormatEx(Calc缺省药品总量( _
                        Val(.TextMatrix(.Row, COL_单量)), 1, 1, 1, "天", "", _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_住院包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                Else
                    sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    If sng天数 = 0 Then sng天数 = 1
                    
                    dbl总量 = FormatEx(Calc缺省药品总量( _
                        Val(.TextMatrix(.Row, COL_单量)), sng天数, _
                        Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                        .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_住院包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                End If
                If Val(txt总量.Text) < dbl总量 Then
                    If MsgBox(.TextMatrix(.Row, COL_名称) & "按每次 " & _
                        .TextMatrix(.Row, COL_单量) & .TextMatrix(.Row, COL_单量单位) & "," & _
                        .TextMatrix(.Row, COL_频率) & IIF(mbln天数, ",用药 " & sng天数 & " 天", "") & _
                        "执行时,至少需要 " & FormatEx(dbl总量, 5) & .TextMatrix(.Row, COL_总量单位) & ",要继续吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt总量_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '检查处方限量,录入限量
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "临嘱" Then
            If Val(.TextMatrix(.Row, COL_处方限量)) <> 0 Then
                dbl总量 = Val(txt总量.Text) * Val(.TextMatrix(.Row, COL_住院包装)) * Val(.TextMatrix(.Row, COL_剂量系数))
                If dbl总量 > Val(.TextMatrix(.Row, COL_处方限量)) Then
                    If MsgBox(.TextMatrix(.Row, COL_名称) & " 的总用量:" & txt总量.Text & lbl总量单位.Caption & "(" & dbl总量 & lbl单量单位.Caption & ")超过处方限量:" & _
                        FormatEx(Val(.TextMatrix(.Row, COL_处方限量)), 5) & lbl单量单位.Caption & "，要继续吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt总量_GotFocus: Exit Sub
                    End If
                End If
            End If
        ElseIf RowIn配方行(.Row) Then
            If Not CheckCHLimited(.Row, Val(txt总量.Text)) Then
                Cancel = True: txt总量_GotFocus: Exit Sub
            End If
        ElseIf InStr(",5,6,7,", .TextMatrix(.Row, COL_类别)) = 0 And Val(.TextMatrix(.Row, COL_处方限量)) > 0 Then
            If Val(txt总量.Text) > Val(.TextMatrix(.Row, COL_处方限量)) Then
                If MsgBox(.TextMatrix(.Row, COL_名称) & " 的总量:" & txt总量.Text & lbl总量单位.Caption & " 超过允许录入的最大限量:" & _
                    .TextMatrix(.Row, COL_处方限量) & lbl总量单位.Caption & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt总量_GotFocus: Exit Sub
                End If
            End If
        End If
        
        '更新数据
        blnTag = txt总量.Tag <> ""
        Call AdviceChange
        
        '药品库存检查:只提醒,修改了才提醒
        If blnTag Then
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Or bln配方行 Then
                strMsg = CheckStock(.Row)
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Function CheckCHLimited(ByVal lngRow As Long, ByVal int付数 As Integer) As Boolean
'功能：检查中药配方每味药的处方限量
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    CheckCHLimited = True
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_类别) = "7" Then
                    strSQL = strSQL & " Union ALL " & _
                        "Select ID,名称,计算单位," & FormatEx(Val(.TextMatrix(i, COL_单量)), 5) & " as 单量 From 诊疗项目目录 Where ID=" & Val(.TextMatrix(i, COL_诊疗项目ID))
                End If
            Else
                Exit For
            End If
        Next
    End With
    If strSQL = "" Then Exit Function
    strSQL = "Select A.ID,A.名称,A.计算单位,A.单量,B.处方限量 From (" & Mid(strSQL, 11) & ") A,药品特性 B Where A.ID=B.药名ID And Nvl(B.处方限量,0)<>0"
    
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) '没法

    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        If int付数 * rsTmp!单量 > rsTmp!处方限量 Then
            strSQL = strSQL & vbCrLf & rsTmp!名称 & "：剂量:" & FormatEx(rsTmp!单量, 5) & Nvl(rsTmp!计算单位) & "," & int付数 & "付;处方限量:" & FormatEx(rsTmp!处方限量, 5) & Nvl(rsTmp!计算单位) & vbTab
        End If
        rsTmp.MoveNext
    Next
    If strSQL <> "" Then
        If MsgBox("该配方中以下药品超出处方限量：" & vbCrLf & strSQL & vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckCHLimited = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearAdviceCard()
'功能：清除医嘱显示卡片相关的内容
'参数：bln开始时间=是否清除开始时间
    Call SetCardEditable(True)
    
    txt开始时间.Text = ""
    txt医嘱内容.Text = ""
    cbo医生嘱托.Text = ""
    cbo执行科室.Clear
    cbo附加执行.Clear
    cbo医生.Text = "" '不清除以保留
    chk紧急.Visible = Not mbln补录
    
    mblnDoCheck = False
    chk紧急.Value = 0
    mblnDoCheck = True
    
    txt开嘱时间.Text = ""
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1)
    Call SetStartTime(True)
    Call SetWriteTime(mbln补录)
    
    stbThis.Panels(3).Text = ""
    stbThis.Panels(4).Text = ""
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'功能：用颜色标识当前医嘱是否可以编辑
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
    cmdSel.Enabled = fraAdvice.Enabled
End Sub

Private Function Get频率范围(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn配方行(lngRow) Then
            Get频率范围 = 2 '中医
        Else
            If RowIn检验行(lngRow) Then '以检验项目行为准
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                Get频率范围 = 1 '可选频率的项目使用西医频率项目
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 1 Then
                Get频率范围 = -1 '一次性
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 2 Then
                Get频率范围 = -2 '持续性
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'功能：当前行为隐藏行时，定位到它所属的可见行
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "E" And Val(.TextMatrix(.Row, COL_相关ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub SetCbo执行性质(ByVal bln离院带药 As Boolean)
    cbo执行性质.Clear
    If bln离院带药 Then
        cbo执行性质.AddItem "0-正常"
        cbo执行性质.AddItem "1-自备药"
        cbo执行性质.AddItem "2-离院带药"
    Else
        cbo执行性质.AddItem "0-正常"
        cbo执行性质.AddItem "1-自备药"
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能：当行改变时，更新卡片内容
    Dim rsItem As New ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    Dim lng用法ID As Long, blnEditable As Boolean
    Dim lngBaseRow As Long, blnGroup As Boolean '中药配方的第一味组成药行
    Dim dblPrice As Double, strTmp As String, i As Long
    Dim lng药品ID As Long
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_开始时间)
    End If
    
    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub
    
    Me.Refresh
    LockWindowUpdate Me.Hwnd

    lngRow = NewRow
    blnGroup = RowIn一并给药(lngRow) '空行也可能在一并给药的范围中
    tbr.Buttons("一并").Value = IIF(blnGroup, tbrPressed, tbrUnpressed)
        
    On Error GoTo errH
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            '无效行清除卡片内容
            Call ClearAdviceCard
            
            '缺省为非自由录入
            tbrFree.Buttons(1).Value = 0
            tbrFree.Buttons(1).Enabled = Not RowIn一并给药(lngRow)
            tbrFree.Buttons(1).Image = IIF(tbrFree.Buttons(1).Enabled, 1, 2)
            
            '缺省期效根据上一行的显示
            i = GetPreRow(lngRow)
            If i = -1 Or Not Visible Then
                Call zlControl.CboSetIndex(cbo期效.Hwnd, 0) '缺省为长嘱
            Else
                Call zlControl.CboSetIndex(cbo期效.Hwnd, IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1))
            End If
            
            '缺省开始时间
            Call txt开始时间_GotFocus
        ElseIf Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由录入医嘱
            blnEditable = True
            If InStr(",1,2,", .TextMatrix(lngRow, COL_状态)) = 0 Then blnEditable = False
            If mbln补录 And Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then blnEditable = False
            
            '已签名的医嘱不可修改
            If Val(.TextMatrix(lngRow, COL_签名否)) = 1 Then blnEditable = False
            
            '修改资格判断
            If blnEditable Then
                If Not mbln护士站 Then
                    '无资格的医生不能修改有资格的医生下达或审核了的医嘱
                    If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(lngRow, COL_开嘱医生))) Then
                        blnEditable = False
                    End If
                Else
                    '护士不能修改医生经过审核的医嘱
                    If .TextMatrix(lngRow, COL_开嘱医生) Like "*/*" Then blnEditable = False
                End If
            End If
            
            Call SetCardEditable(blnEditable)
            
            tbrFree.Buttons(1).Value = 1
            tbrFree.Buttons(1).Enabled = blnEditable
            tbrFree.Buttons(1).Image = IIF(blnEditable, 1, 2)
            cmdExt.Enabled = False
            cmdSel.Enabled = False
            
            '其它输入项禁用
            Call SetDayState(-1, -1)
            SetItemEditable -1, -1, -1, -1, -1, , -1, -1, -1
            stbThis.Panels(3).Text = "": stbThis.Panels(4).Text = ""
            
            '显示当前医嘱卡片内容
            '--------------------------------------------------------------------------------------------
            Call zlControl.CboSetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
            
            '开始时间：只有新增医嘱时或是补录的医嘱可以修改开始时间
            txt开始时间.Text = .Cell(flexcpData, lngRow, COL_开始时间)
            Call SetStartTime(Val(.TextMatrix(lngRow, COL_标志)) = 2 Or .TextMatrix(lngRow, COL_EDIT) = "1")
            
            '医嘱内容
            txt医嘱内容.Text = .TextMatrix(lngRow, COL_医嘱内容)
            
            '医生嘱托
            cbo医生嘱托.Text = .TextMatrix(lngRow, COL_医生嘱托)
            
            '紧急标志
            If Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then
                chk紧急.Visible = True
                mblnDoCheck = False
                chk紧急.Value = Val(.TextMatrix(lngRow, COL_标志))
                mblnDoCheck = True
            Else
                chk紧急.Visible = False
            End If
            
            '可选执行科室
            SetItemEditable , , , , , 1
            Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cbo执行科室, "*", 0, 0, 4, mlng病人科室id, 0, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo期效.ListIndex)
            
            '开嘱医生
            If .TextMatrix(lngRow, COL_开嘱医生) Like "*/*" Then
                cbo医生.Clear: cbo医生.Locked = True
                cbo医生.AddItem .TextMatrix(lngRow, COL_开嘱医生)
                Call zlControl.CboSetIndex(cbo医生.Hwnd, cbo医生.NewIndex)
            Else
                cbo医生.Locked = Not mbln护士站
                If mbln护士站 Then
                    '加入当前医生:不清除历史选择
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), 0, cbo医生, 2, True)
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), 0, cbo医生)
                End If
            End If
            
            '开嘱时间
            txt开嘱时间.Text = .Cell(flexcpData, lngRow, COL_开嘱时间)
            Call SetWriteTime(Val(.TextMatrix(lngRow, COL_标志)) = 2)
        Else
            '卡片编辑：已校对的医嘱不能修改,补录医嘱时不能更改非补录的内容
            blnEditable = True
            If InStr(",1,2,", .TextMatrix(lngRow, COL_状态)) = 0 Then blnEditable = False
            If mbln补录 And Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then blnEditable = False
            
            '已签名的医嘱不可修改
            If Val(.TextMatrix(lngRow, COL_签名否)) = 1 Then blnEditable = False
            
            '修改资格判断
            If blnEditable Then
                If Not mbln护士站 Then
                    '无资格的医生不能修改有资格的医生下达或审核了的医嘱
                    If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(lngRow, COL_开嘱医生))) Then
                        blnEditable = False
                    End If
                Else
                    '护士不能修改医生经过审核的医嘱
                    If .TextMatrix(lngRow, COL_开嘱医生) Like "*/*" Then blnEditable = False
                End If
            End If
            
            Call SetCardEditable(blnEditable)
            
            '已有诊疗项目，不可变为自由录入
            tbrFree.Buttons(1).Value = 0
            tbrFree.Buttons(1).Enabled = False
            tbrFree.Buttons(1).Image = 2
            
            '获取诊疗项目基本信息
            '---------------------
            If InStr(",5,6,7,", Val(.TextMatrix(lngRow, COL_类别))) > 0 Then
                lng药品ID = Val(.TextMatrix(lngRow, COL_收费细目ID))
            End If
            
            If RowIn配方行(lngRow) Then
                txt总量.MaxLength = 3
                '获取中药配方第一味中药行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                lng药品ID = Val(.TextMatrix(lngBaseRow, COL_收费细目ID))
            ElseIf RowIn检验行(lngRow) Then
                '获取一并采样的第一个项目行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                txt总量.MaxLength = txt单量.MaxLength
            Else
                lngBaseRow = lngRow
                txt总量.MaxLength = txt单量.MaxLength
            End If
            strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
            Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngBaseRow, COL_诊疗项目ID)))
            
            '扩展按钮可用状态(检查组合,检验组合,手术,中药配方)
            cmdExt.Enabled = InStr(",7,C,F,", rsItem!类别) > 0 Or (rsItem!类别 = "D" And Nvl(rsItem!组合项目, 0) = 1)
            
            '显示当前医嘱卡片内容
            '--------------------------------------------------------------------------------------------
            Call zlControl.CboSetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
            
            '开始时间：只有新增医嘱时或是补录的医嘱可以修改开始时间
            txt开始时间.Text = .Cell(flexcpData, lngRow, COL_开始时间)
            Call SetStartTime(Val(.TextMatrix(lngRow, COL_标志)) = 2 Or .TextMatrix(lngRow, COL_EDIT) = "1")
            
            '医嘱内容
            txt医嘱内容.Text = .TextMatrix(lngRow, COL_医嘱内容)
            
            '单量
            '----------------------
            If rsItem!类别 = "7" Then '中药配方(中草药)虽然有单量,但不在这里填写
                SetItemEditable -1
            ElseIf cbo期效.ListIndex = 0 Then
                '长嘱：成药或计时,计量项目可以录入
                If InStr(",1,2,", Nvl(rsItem!计算方式, 0)) > 0 Or InStr(",5,6,", rsItem!类别) > 0 Then
                    SetItemEditable 1
                    txt单量.Text = .TextMatrix(lngRow, COL_单量)
                    lbl单量单位.Caption = .TextMatrix(lngRow, COL_单量单位)
                Else
                    SetItemEditable -1
                End If
            ElseIf cbo期效.ListIndex = 1 Then
                '临嘱:成药或可选择频率的计时,计量项目可以录入(注意这是原始频率,当前可能已是一次性)
                If (Nvl(rsItem!执行频率, 0) = 0 And InStr(",1,2,", Nvl(rsItem!计算方式, 0)) > 0) _
                    Or InStr(",5,6,", rsItem!类别) > 0 Then
                    SetItemEditable 1
                    txt单量.Text = .TextMatrix(lngRow, COL_单量)
                    lbl单量单位.Caption = .TextMatrix(lngRow, COL_单量单位)
                Else
                    SetItemEditable -1
                End If
            End If
            
            '天数：西药，中成药临嘱才使用，用于计算总量
            '一般：临嘱的药品(非中药)或可选择频率的计时,计量项目可以使用天数来自动计算总量
            blnEditable = False
            If cbo期效.ListIndex = 1 And InStr(",5,6,", rsItem!类别) > 0 Then
                If mbln天数 And Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then blnEditable = True
            End If
            If blnEditable Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt天数.Text = Val(.TextMatrix(lngRow, COL_天数))
            If Val(txt天数.Text) = 0 Then txt天数.Text = ""
            
            '总量
            '--------------------
            If rsItem!类别 = "7" Then
                '中药配方(中草药)填写为付数
                SetItemEditable , 1
                lbl总量单位.Caption = "付"
                txt总量.Text = .TextMatrix(lngRow, COL_总量) '付数
            ElseIf cbo期效.ListIndex = 1 Then
                '临嘱都需要填写总量:临嘱发送以总量为准
                If rsItem!类别 = "Z" And Nvl(rsItem!操作类型) <> "0" Then
                    SetItemEditable , -1 '特殊医嘱不允许修改总量(固定为1次)
                ElseIf InStr(",5,6,", rsItem!类别) = 0 And Nvl(rsItem!计算方式, 0) = 3 _
                    And (Nvl(rsItem!执行频率, 0) = 1 Or Val(.TextMatrix(lngRow, COL_频率性质)) = 1) Then
                    SetItemEditable , -1 '非药品一次性计次项目不输入总量(原始频率为一次性或当前设置为一次性)
                Else
                    SetItemEditable , 1
                End If
                lbl总量单位.Caption = .TextMatrix(lngRow, COL_总量单位)
                txt总量.Text = .TextMatrix(lngRow, COL_总量)
            Else
                '其它长嘱不允许填写总量
                SetItemEditable , -1
            End If
            
            '给药途径和中药用法
            '--------------
            If InStr(",5,6,", rsItem!类别) > 0 Then
                SetItemEditable , , 1
                lbl用法.Caption = "给药途径"
                '查找给药途径对应的行:查找的Rowdata(Variant)数据要转为Long型,才能精确匹配
                lng用法ID = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            ElseIf rsItem!类别 = "7" Then
                SetItemEditable , , 1
                lbl用法.Caption = "中药用法"
                
                '中药配方显示行就是中药用法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            ElseIf RowIn检验行(lngRow) Then '不用类别判断,兼容以前的检验
                '检验组合
                SetItemEditable , , 1
                lbl用法.Caption = "采集方法"
                
                '检验组合显示行就是采集方法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            Else
                SetItemEditable , , -1
            End If
            
            '频率：都可以选择(临嘱输入用于指导使用)
            If True Then
                SetItemEditable , , , 1
                cmd频率.Tag = .TextMatrix(lngRow, COL_频率)
                txt频率.Text = .TextMatrix(lngRow, COL_频率)
            Else
                SetItemEditable , , , -1
            End If
                    
            '执行时间："可选频率"或药品(当前未被设置为一次性)。
            If Nvl(rsItem!执行频率, 0) = 0 And Val(.TextMatrix(lngBaseRow, COL_频率性质)) <> 1 Then
                SetItemEditable , , , , 1
                Call Get时间方案(cbo执行时间, Get频率范围(lngRow), .TextMatrix(lngRow, COL_频率), lng用法ID)
                cbo执行时间.Text = .TextMatrix(lngRow, COL_执行时间)
            Else
                SetItemEditable , , , , -1
            End If
                    
            '医生嘱托
            cbo医生嘱托.Text = .TextMatrix(lngRow, COL_医生嘱托)
                    
            '终止时间：长嘱可以修改
            If cbo期效.ListIndex = 0 Then
                If rsItem!类别 = "7" Then
                    SetItemEditable , , , , , , 2 '中药配方长嘱不可修改终止时间(锁定)
                ElseIf rsItem!类别 = "H" And Nvl(rsItem!操作类型, 0) = 1 Then
                    SetItemEditable , , , , , , -1 '护理等级为自动停止,不可指定
                Else
                    SetItemEditable , , , , , , 1
                End If
            Else
                SetItemEditable , , , , , , -1
            End If
            txt终止时间.Text = .Cell(flexcpData, lngRow, COL_终止时间)
                    
            '执行性质:长嘱目前可以使用"自备药"
            If InStr(",5,6,7,", rsItem!类别) > 0 Then
                If rsItem!类别 = "7" Then
                    '对于中药配方,根据诊疗项目管理中限制及本程序处理,不可能用法和煎法一个为院外执行,一个不为
                    If Val(.TextMatrix(lngBaseRow, COL_执行性质)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 Then
                        strTmp = "自备药"
                    ElseIf Val(.TextMatrix(lngBaseRow, COL_执行性质)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                        strTmp = "离院带药"
                    Else
                        strTmp = "正常"
                    End If
                Else
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                        strTmp = "自备药"
                    ElseIf Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                        strTmp = "离院带药"
                    Else
                        strTmp = "正常"
                    End If
                End If
                Call SetCbo执行性质(cbo期效.ListIndex = 1)
                SetItemEditable , , , , , , , 1
                Call SeekIndex(cbo执行性质, strTmp)
            Else
                SetItemEditable , , , , , , , -1
            End If
                    
            '执行科室
            If rsItem!类别 = "Z" And Nvl(rsItem!操作类型, 0) = 3 Then
                '转科医嘱用临床科室
                SetItemEditable , , , , , 1
                lbl执行科室.Caption = "转入科室"
                Call Get临床科室(2, mlng病人科室id, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室, Not gbln病区科室独立)
            ElseIf rsItem!类别 = "Z" And Nvl(rsItem!操作类型, 0) = 7 Then
                '会诊医嘱用临床科室
                SetItemEditable , , , , , 1
                lbl执行科室.Caption = "会诊科室"
                Call Get临床科室(2, mlng病人科室id, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室)
            Else
                '是药品则以药品行为准显示,检验组合以检验项目为准显示
                i = lngRow
                If rsItem!类别 = "7" Then
                    i = lngBaseRow
                ElseIf RowIn检验行(lngRow) Then '不用类别判断,兼容以前的检验
                    i = lngBaseRow
                End If
                
                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                    '非叮嘱和院外执行时才显示和可以选择(包括药品)
                    SetItemEditable , , , , , 1
                    Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cbo执行科室, rsItem!类别, rsItem!ID, lng药品ID, Nvl(rsItem!执行科室, 0), _
                        mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), Val(.TextMatrix(i, COL_执行科室ID)), cbo期效.ListIndex)
                ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                    SetItemEditable , , , , , -1
                    If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                        cbo执行科室.AddItem "<无执行叮嘱>"
                    Else
                        cbo执行科室.AddItem "<院外执行>"
                    End If
                    Call zlControl.CboSetIndex(cbo执行科室.Hwnd, 0)
                End If
            End If
            
            '附加执行:指给药途径,中药用法,手术麻醉,采集方式的执行科室
            If Should附加执行(lngRow, i, strTmp) Then
                SetItemEditable , , , , , , , , 1
                Call Get诊疗执行科室(mlng病人ID, mlng主页ID, cbo附加执行, .TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), lng药品ID, _
                    Val(.TextMatrix(i, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), Val(.TextMatrix(i, COL_执行科室ID)), cbo期效.ListIndex)
            Else
                SetItemEditable , , , , , , , , -1
                If i <> -1 Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                        If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                            cbo附加执行.AddItem "<无执行叮嘱>"
                        ElseIf Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            cbo附加执行.AddItem "<院外执行>"
                        End If
                        Call zlControl.CboSetIndex(cbo附加执行.Hwnd, 0)
                    End If
                End If
            End If
            lbl附加执行.Caption = strTmp
            
            '紧急标志
            If Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then
                chk紧急.Visible = True
                mblnDoCheck = False
                chk紧急.Value = Val(.TextMatrix(lngRow, COL_标志))
                mblnDoCheck = True
            Else
                chk紧急.Visible = False
            End If
            
            '开嘱医生
            If .TextMatrix(lngRow, COL_开嘱医生) Like "*/*" Then
                cbo医生.Clear: cbo医生.Locked = True
                cbo医生.AddItem .TextMatrix(lngRow, COL_开嘱医生)
                Call zlControl.CboSetIndex(cbo医生.Hwnd, cbo医生.NewIndex)
            Else
                cbo医生.Locked = Not mbln护士站
                If mbln护士站 Then
                    '加入当前医生:不清除历史选择
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), 0, cbo医生, 2, True)
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), 0, cbo医生)
                End If
            End If
                        
            '开嘱时间
            txt开嘱时间.Text = .Cell(flexcpData, lngRow, COL_开嘱时间)
            Call SetWriteTime(Val(.TextMatrix(lngRow, COL_标志)) = 2)
                    
            '显示药品库存：以住院单位，中药配方不显示
            '----------------------------------------
            If InStr(",5,6,", rsItem!类别) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                If .TextMatrix(lngRow, COL_库存) = "" Then Call GetDrugStock(lngRow)
                If .TextMatrix(lngRow, COL_库存) <> "" Then
                    stbThis.Panels(3).Text = "库存:" & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & .TextMatrix(lngRow, COL_住院单位)
                Else
                    stbThis.Panels(3).Text = ""
                End If
            Else
                If rsItem!类别 = "7" And InStr(",1,2,", .TextMatrix(lngRow, COL_状态)) > 0 Then
                    Call GetDrugStock(lngRow)
                End If
                stbThis.Panels(3).Text = ""
            End If
            
            '显示医嘱单价
            dblPrice = GetItemPrice(lngRow)
            If dblPrice <> 0 Then
                If InStr(",5,6,", rsItem!类别) > 0 Then
                    stbThis.Panels(4).Text = "每" & .TextMatrix(lngRow, COL_住院单位) & ":" & FormatEx(dblPrice, 5) & "元"
                ElseIf rsItem!类别 = "7" Then
                    stbThis.Panels(4).Text = "每付:" & FormatEx(dblPrice, 5) & "元"
                Else
                    stbThis.Panels(4).Text = IIF(IsNull(rsItem!计算单位), "价格:", "每" & Nvl(rsItem!计算单位) & ":") & FormatEx(dblPrice, 5) & "元"
                End If
            Else
                stbThis.Panels(4).Text = ""
            End If
            
            '显示费用类型
            strTmp = Get费用类型(lngRow)
            If strTmp <> "" Then
                stbThis.Panels(4).Text = stbThis.Panels(4).Text & IIF(stbThis.Panels(4).Text = "", "类型:", ",类型:") & strTmp
            End If
        End If
    End With
    
    '清除编辑标志
    Call ClearItemTag
    
    '设置医嘱功能可用性
    Call SetFuncEnabled
    
    '显示计价窗体
    Call ShowPrice(lngRow)
    
    LockWindowUpdate 0
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'根据当前行的情况显示计价窗体
    If mblnModal Then Exit Sub
    
    If vsAdvice.RowData(lngRow) = 0 Or Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_状态))) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf RowIn配方行(lngRow) Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf stbThis.Panels("Price").Bevel = sbrNoBevel Then
        stbThis.Panels("Price").Visible = True
        If stbThis.Panels("Price").Tag <> "" Then
            stbThis.Panels("Price").Bevel = sbrInset
        Else
            stbThis.Panels("Price").Bevel = sbrRaised
        End If
    End If
    
    If stbThis.Panels("Price").Bevel <> sbrInset Then
        '关闭计价窗体
        mfrmPrice.HideMe
    Else
        Call mfrmPrice.ShowMe(Me, vsAdvice, mlng病人ID, mlng主页ID, mlng病人科室id, _
            COL_序号 & "," & COL_相关ID & "," & COL_状态 & "," & COL_类别 & "," & COL_诊疗项目ID & "," & _
            COL_收费细目ID & "," & COL_标本部位 & "," & COL_计价性质 & "," & COL_执行性质 & "," & COL_执行科室ID)
    End If
End Sub

Private Sub SetFuncEnabled()
'功能：设置医嘱功能可用性
    Dim blnEnabled As Boolean
    With vsAdvice
        '停止功能
        blnEnabled = True
        If .RowData(.Row) = 0 Then blnEnabled = False
        If InStr(",1,2,4,8,9,", Val(.TextMatrix(.Row, COL_状态))) > 0 Then blnEnabled = False
        If .TextMatrix(.Row, COL_期效) <> "长嘱" Then blnEnabled = False
        If RowIn配方行(.Row) Then blnEnabled = False
        tbr.Buttons("停止").Enabled = blnEnabled
        
        '删除功能
        blnEnabled = True
        If .RowData(.Row) <> 0 Then
            If Not fraAdvice.Enabled Then blnEnabled = False
            If InStr(",1,2,", .TextMatrix(.Row, COL_状态)) = 0 Then blnEnabled = False
            '已签名医嘱不可删除
            If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then blnEnabled = False
        End If
        tbr.Buttons("删除").Enabled = blnEnabled
        
        '一并功能
        blnEnabled = True
        If Not fraAdvice.Enabled Then blnEnabled = False
        tbr.Buttons("一并").Enabled = blnEnabled
        
        '插入功能
        blnEnabled = True
        If Not fraAdvice.Enabled Then
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 _
                And Val(.TextMatrix(.Row, COL_相关ID)) = Val(.TextMatrix(.Row - 1, COL_相关ID)) Then
                blnEnabled = False
            End If
        End If
        tbr.Buttons("插入").Enabled = blnEnabled
    End With
End Sub

Private Function Get费用类型(ByVal lngRow As Long) As String
'功能：获取指定行的费用类型
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str类型 As String
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
            '取医保的费用类型
            If mint险类 <> 0 Then
                str类型 = gclsInsure.GetItemInsure(mlng病人ID, Val(.TextMatrix(lngRow, COL_收费细目ID)), 0, False, mint险类)
                If str类型 <> "" Then
                    If UBound(Split(str类型, ";")) >= 5 Then
                        str类型 = Split(str类型, ";")(5)
                    Else
                        str类型 = ""
                    End If
                End If
            End If
            '没有则取HIS的费用类型
            If str类型 = "" Then
                strSQL = "Select 费用类型 From 收费项目目录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
                If Not rsTmp.EOF Then str类型 = Nvl(rsTmp!费用类型)
            End If
        End If
    End With
    Get费用类型 = str类型
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Should附加执行(ByVal lngRow As Long, lngRow2 As Long, str执行科室 As String) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
'参数：lngRow2=返回附加行的医嘱行号
'      str执行科室=附加执行科室类型
    Dim i As Long
    
    lngRow2 = -1
    str执行科室 = "附加执行"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function

        If RowIn配方行(lngRow) Then
            '中药用法
            lngRow2 = lngRow
            str执行科室 = "用法执行"
            Should附加执行 = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            '给药途径
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
            str执行科室 = "给药执行"
            Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
            '手术麻醉
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            str执行科室 = "麻醉执行"
            If lngRow2 <> -1 Then Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
            And .TextMatrix(lngRow - 1, COL_类别) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
            '采集方式
            lngRow2 = lngRow
            str执行科室 = "采集执行"
            Should附加执行 = True
        End If
        
        '叮嘱或院外执行
        If Should附加执行 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_执行性质))) > 0 Then
                Should附加执行 = False
            End If
        End If
    End With
End Function

Private Function GetItemPrice(ByVal lngRow As Long) As Double
'功能：获取当前医嘱行的价格(药品为一个药房包装的单价,其它根据收费对照)
'说明：药品不包含给药途径及中药用法煎法
    Dim rsTmp As New ADODB.Recordset
    Dim str医嘱IDs As String, str项目IDs As String, str单量s As String
    Dim strAdviceIDs As String, lng执行科室ID As Long
    Dim dblPrice As Double, dbl数量 As Double
    Dim bln药品 As Boolean, strSQL As String, i As Long
    
    With vsAdvice
        bln药品 = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
            '西药及中成药按规格下才能计算价格
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                str项目IDs = str项目IDs & "," & Val(.TextMatrix(lngRow, COL_收费细目ID))
            End If
            lng执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        ElseIf RowIn配方行(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" And Val(.TextMatrix(i, COL_收费细目ID)) <> 0 Then
                        If lng执行科室ID = 0 Then
                            lng执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                        End If
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COL_收费细目ID))
                        str单量s = str单量s & ";" & Val(.TextMatrix(i, COL_单量))
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            bln药品 = False
            '其它医嘱,未校对(计价)的按收费对照计算,否则直接取医嘱计价
            '不包含不计价和手工计价的项目
            If Val(.TextMatrix(lngRow, COL_计价性质)) = 0 Then
                If InStr(",1,2,", .TextMatrix(lngRow, COL_状态)) > 0 Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                Else
                    str医嘱IDs = str医嘱IDs & "," & .RowData(lngRow)
                End If
            End If
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_计价性质)) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_状态)) > 0 Then
                            str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                        Else
                            str医嘱IDs = str医嘱IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1 '检验组合
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_计价性质)) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_状态)) > 0 Then
                            str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                        Else
                            str医嘱IDs = str医嘱IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    str医嘱IDs = Mid(str医嘱IDs, 2)
    str项目IDs = Mid(str项目IDs, 2)
    str单量s = Mid(str单量s, 2)
    
    On Error GoTo errH
    
    If bln药品 Then
        If str项目IDs = "" Then Exit Function
    
        '不排序时,ID顺序为从右向左
        strSQL = "Select A.ID,A.是否变价,B.住院包装,B.剂量系数,B.可否分零 From 收费项目目录 A,药品规格 B Where A.ID=B.药品ID And A.ID IN(" & str项目IDs & ")"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'In
        For i = 1 To rsTmp.RecordCount
            '数量:住院包装
            If str单量s <> "" Then '中药配方才管每味剂量
                dbl数量 = Val(Split(str单量s, ";")(rsTmp.RecordCount - i))
                '中药药房单位按不可分零处理:每付
                If Nvl(rsTmp!可否分零, 0) = 0 Then
                    dbl数量 = Format(dbl数量 / Nvl(rsTmp!剂量系数, 1) / Nvl(rsTmp!住院包装, 1), "0.00000")
                Else
                    dbl数量 = IntEx(dbl数量 / Nvl(rsTmp!剂量系数, 1) / Nvl(rsTmp!住院包装, 1))
                End If
            Else
                dbl数量 = 1
            End If
            If Nvl(rsTmp!是否变价, 0) = 0 Then
                dblPrice = dblPrice + CalcPrice(rsTmp!ID) * Nvl(rsTmp!住院包装, 1) * dbl数量
            Else
                dblPrice = dblPrice + CalcDrugPrice(rsTmp!ID, lng执行科室ID, dbl数量 * Nvl(rsTmp!住院包装, 1)) * Nvl(rsTmp!住院包装, 1) * dbl数量
            End If
            rsTmp.MoveNext
        Next
    Else
        If str项目IDs = "" And str医嘱IDs = "" Then Exit Function
    
        If str医嘱IDs <> "" Then
            strSQL = _
                " Select B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价" & _
                " From 病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And B.医嘱ID IN(" & str医嘱IDs & ")" & _
                " Group by B.数量,C.是否变价,B.单价"
        End If
        If str项目IDs <> "" Then
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select B.收费数量 as 数量,Decode(C.是否变价,1,0,Sum(D.现价)) as 单价" & _
                " From 诊疗收费关系 B,收费项目目录 C,收费价目 D" & _
                " Where B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And B.诊疗项目ID IN(" & str项目IDs & ")" & _
                " Group by B.收费数量,C.是否变价"
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Name) 'In
        For i = 1 To rsTmp.RecordCount
            dblPrice = dblPrice + Format(Nvl(rsTmp!数量, 0) * Nvl(rsTmp!单价, 0), "0.00000")
            rsTmp.MoveNext
        Next
    End If
    
    GetItemPrice = Format(dblPrice, "0.00000")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDrugStock(ByVal lngRow As Long)
'功能：重新获取指定药品行的药品库存
'参数：lngRow=成药行或中药用法行
'说明：如果是中药配方行,一次性获取整个配方中的所有中药的库存
    Dim i As Long
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Or Val(.TextMatrix(lngRow, COL_收费细目ID)) = 0 Then
                .TextMatrix(lngRow, COL_库存) = ""
            Else
                .TextMatrix(lngRow, COL_库存) = GetStock(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)))
            End If
        ElseIf RowIn配方行(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Or Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then
                            .TextMatrix(i, COL_库存) = ""
                        Else
                            .TextMatrix(i, COL_库存) = GetStock(Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_执行科室ID)))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function Get住院医师() As String
'功能：获取当前病人的住院医师
    Dim i As Integer
    For i = 1 To lvwPati.ListItems.Count
        If Val(Mid(lvwPati.ListItems(i).Key, 2)) = mlng病人ID Then
            Get住院医师 = lvwPati.ListItems(i).SubItems(3)
            Exit For
        End If
    Next
End Function

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = COL_医嘱内容 Then Call vsAdvice.AutoSize(COL_医嘱内容)
    End If
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpDate.Visible Or lvwPati.Visible Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            If .MouseCol >= .FixedCols And .MouseCol <= .Cols - 1 Then
                If .MouseCol = COL_医嘱内容 Then
                    '复制医嘱
                    If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
                        Call AdviceCopyCurr(vsAdvice.Row)
                        If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
                    End If
                Else
                    Call vsAdvice_KeyPress(13) '定位到对应的编辑控件
                End If
            ElseIf .MouseCol = 0 Then
                '填写申请
                '##
            End If
        End If
    End With
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否最后一可见行
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

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
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
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
            lngLeft = COL_期效: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_频率: lngRight = COL_用法
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowIn一并给药(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Get一并给药范围(Val(.TextMatrix(Row - 1, COL_相关ID)), lngBegin, lngEnd)
            Else
                Call Get一并给药范围(Val(.TextMatrix(Row, COL_相关ID)), lngBegin, lngEnd)
            End If
            
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
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If tbr.Buttons("删除").Enabled And tbr.Buttons("删除").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("删除"))
        End If
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    
    If KeyAscii = 13 Then
        '定位到对应的编辑控件
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_期效
                Set objEdit = cbo期效
            Case COL_开始时间
                Set objEdit = txt开始时间
            Case COL_医嘱内容
                Set objEdit = txt医嘱内容
            Case COL_单量
                Set objEdit = txt单量
            Case COL_总量
                Set objEdit = txt总量
            Case COL_用法
                Set objEdit = txt用法
            Case COL_频率
                Set objEdit = txt频率
            Case COL_执行时间
                Set objEdit = cbo执行时间
            Case COL_执行科室ID
                Set objEdit = cbo执行科室
            Case COL_终止时间
                Set objEdit = txt终止时间
            Case COL_开嘱医生
                Set objEdit = cbo医生
            Case COL_开嘱时间
                Set objEdit = txt开嘱时间
            Case COL_医生嘱托
                Set objEdit = cbo医生嘱托
            Case COL_标志
                Set objEdit = chk紧急
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    End If
End Sub

Private Sub ClearItemTag()
'功能：清除控件编辑标志
    txt开始时间.Tag = ""
    txt单量.Tag = ""
    txt天数.Tag = ""
    txt总量.Tag = ""
    txt用法.Tag = ""
    txt频率.Tag = ""
    cbo执行时间.Tag = ""
    cbo医生嘱托.Tag = ""
    cbo执行科室.Tag = ""
    cbo执行性质.Tag = ""
    cbo附加执行.Tag = ""
    txt终止时间.Tag = ""
    txt开嘱时间.Tag = ""
    cbo医生.Tag = ""
    chk紧急.Tag = ""
End Sub

Private Sub SetStartTime(ByVal Editable As Boolean)
'功能：设置开始时间是否允许编辑
    txt开始时间.TabStop = Editable
    txt开始时间.Locked = Not Editable
    cmd开始时间.Enabled = Editable
    If Editable Then
        txt开始时间.BackColor = lvwPati.BackColor
    Else
        txt开始时间.BackColor = &HE0E0E0
    End If
End Sub

Private Sub SetWriteTime(ByVal Editable As Boolean)
'功能：设置开嘱时间是否允许编辑
    txt开嘱时间.TabStop = Editable
    txt开嘱时间.Locked = Not Editable
    cmd开嘱时间.Enabled = Editable
    If Editable Then
        txt开嘱时间.BackColor = lvwPati.BackColor
    Else
        txt开嘱时间.BackColor = &HE0E0E0
    End If
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'功能：设置执行天数可用和或见状态
'参数：0-保持不变,-1-禁止,1-允许
    If intEnabled = -1 Then
        txt天数.Enabled = False
        txt天数.BackColor = Me.BackColor
        txt天数.Text = ""
    ElseIf intEnabled = 1 Then
        txt天数.TabStop = True
        txt天数.Enabled = True
        txt天数.BackColor = lvwPati.BackColor
    End If
    
    If intVisible = -1 Then
        lbl天数.Visible = False
        txt天数.Visible = False
        txt天数.Text = ""
        
        lbl总量.Left = lbl用法.Left + lbl用法.Width - lbl总量.Width
        txt总量.Left = txt用法.Left
        txt总量.Width = txt用法.Width - cmd用法.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        lbl单量.Left = lbl频率.Left + lbl频率.Width - lbl单量.Width
        txt单量.Left = txt频率.Left
        txt单量.Width = txt频率.Width - cmd频率.Width - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        txt总量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt总量.TabIndex + 1
        txt单量.TabIndex = txt天数.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl天数.Visible = True
        txt天数.Visible = True
        
        lbl单量.Left = lbl用法.Left + lbl用法.Width - lbl单量.Width
        txt单量.Left = txt用法.Left
        txt单量.Width = txt用法.Width - txt天数.Width - Me.TextWidth("三个字!") - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        lbl总量.Left = lbl频率.Left + lbl频率.Width - lbl总量.Width
        txt总量.Left = txt频率.Left
        txt总量.Width = txt频率.Width - cmd频率.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        txt单量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt单量.TabIndex + 1
        txt总量.TabIndex = txt天数.TabIndex + 1
    End If
End Sub

Private Sub SetItemEditable(Optional int单量 As Integer, Optional int总量 As Integer, _
    Optional int用法 As Integer, Optional int频率 As Integer, _
    Optional int执行时间 As Integer, Optional int执行科室 As Integer, _
    Optional int终止时间 As Integer, Optional int执行性质 As Integer, _
    Optional int附加执行 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-禁止,1-允许,2-锁定
'说明：禁止时,同时清除该项目数据(不是全部)

    '依次设置为禁止时,会引发焦点改变,从而可能引发Validate事件,所以先禁止焦点顺序
    If int单量 = -1 Then txt单量.TabStop = False
    If int总量 = -1 Then txt总量.TabStop = False
    If int用法 = -1 Then txt用法.TabStop = False
    If int频率 = -1 Then txt频率.TabStop = False
    If int执行时间 = -1 Then cbo执行时间.TabStop = False
    If int执行科室 = -1 Then cbo执行科室.TabStop = False
    If int终止时间 = -1 Then txt终止时间.TabStop = False
    If int执行性质 = -1 Then cbo执行性质.TabStop = False
    If int附加执行 = -1 Then cbo附加执行.TabStop = False
    
    If int单量 = -1 Then
        txt单量.Enabled = False
        txt单量.BackColor = Me.BackColor
        txt单量.Text = ""
        lbl单量单位.Caption = "" '"单位"
    ElseIf int单量 = 1 Then
        txt单量.TabStop = True
        txt单量.Enabled = True
        txt单量.BackColor = lvwPati.BackColor
    End If

    If int总量 = -1 Then
        txt总量.Enabled = False
        txt总量.BackColor = Me.BackColor
        txt总量.Text = ""
        lbl总量单位.Caption = "" '"单位"
    ElseIf int总量 = 1 Then
        txt总量.TabStop = True
        txt总量.Enabled = True
        txt总量.BackColor = lvwPati.BackColor
    End If
    
    If int用法 = -1 Then
        txt用法.Enabled = False
        txt用法.BackColor = Me.BackColor
        txt用法.Text = ""
        cmd用法.Enabled = False
        lbl用法.Caption = "用法"
    ElseIf int用法 = 1 Then
        txt用法.TabStop = True
        txt用法.Enabled = True
        cmd用法.Enabled = True
        txt用法.BackColor = lvwPati.BackColor
    End If

    If int频率 = -1 Then
        txt频率.Enabled = False
        cmd频率.Enabled = False
        txt频率.BackColor = Me.BackColor
        txt频率.Text = ""
    ElseIf int频率 = 1 Then
        txt频率.TabStop = True
        txt频率.Enabled = True
        cmd频率.Enabled = True
        txt频率.BackColor = lvwPati.BackColor
    End If

    If int执行时间 = -1 Then
        cbo执行时间.Enabled = False
        cbo执行时间.BackColor = Me.BackColor
        cbo执行时间.Clear
    ElseIf int执行时间 = 1 Then
        cbo执行时间.TabStop = True
        cbo执行时间.Enabled = True
        cbo执行时间.BackColor = lvwPati.BackColor
    End If

    If int执行科室 = -1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.Enabled = False
        cbo执行科室.BackColor = Me.BackColor
        cbo执行科室.Clear
    ElseIf int执行科室 = 1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.TabStop = True
        cbo执行科室.Enabled = True
        cbo执行科室.BackColor = lvwPati.BackColor
    End If

    If int执行性质 = -1 Then
        cbo执行性质.Enabled = False
        cbo执行性质.BackColor = Me.BackColor
        Call zlControl.CboSetIndex(cbo执行性质.Hwnd, -1) '不清除
    ElseIf int执行性质 = 1 Then
        cbo执行性质.TabStop = True
        cbo执行性质.Enabled = True
        cbo执行性质.BackColor = lvwPati.BackColor
    End If
    
    If int附加执行 = -1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.Enabled = False
        cbo附加执行.BackColor = Me.BackColor
        cbo附加执行.Clear
    ElseIf int附加执行 = 1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.TabStop = True
        cbo附加执行.Enabled = True
        cbo附加执行.BackColor = lvwPati.BackColor
    End If
    
    If int终止时间 = -1 Then
        cmd终止时间.Enabled = False
        txt终止时间.Enabled = False
        txt终止时间.Locked = False
        txt终止时间.BackColor = Me.BackColor
        txt终止时间.Text = ""
    ElseIf int终止时间 = 1 Then
        txt终止时间.TabStop = True
        cmd终止时间.Enabled = True
        txt终止时间.Enabled = True
        txt终止时间.Locked = False
        txt终止时间.BackColor = lvwPati.BackColor
    ElseIf int终止时间 = 2 Then '锁定(不可编辑,不清除内容)
        txt终止时间.TabStop = False
        cmd终止时间.Enabled = False
        txt终止时间.Enabled = True
        txt终止时间.Locked = True
        txt终止时间.BackColor = &HE0E0E0
    End If
End Sub

Private Function LoadPatients() As Boolean
'功能：读取与调用界面相同范围的病人列表
    Dim rsPati As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim strMsg As String, i As Integer, j As Integer
    Dim str科室IDs As String, lng部门ID As Long
    Dim intBedLen As Integer, blnDo As Boolean
    
    On Error GoTo errH
    
    strSQL = _
        "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄," & _
        " B.入院日期,B.出院日期,B.住院医师,B.出院病床 as 床号,B.费别," & _
        " B.险类,B.出院科室ID as 科室ID,B.当前病区ID as 病区ID,C.名称 as 护理等级," & _
        " B.状态,B.数据转出,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
        " From 病人信息 A,病案主页 B,收费项目目录 C" & _
        " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+)" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        
    '会诊调用，医技调用，或非在院病人：不能选择
    blnDo = False
    If mlng前提ID = 0 And Not mbln会诊 Then
        If IsNull(rsPati!出院日期) And Nvl(rsPati!状态, 0) <> 3 Then
            blnDo = True
        End If
    End If
    If blnDo Then
        '提供当前科室/病区的在院病人清单供选择
        lng部门ID = IIF(mbln护士站, Nvl(rsPati!病区ID, 0), Nvl(rsPati!科室ID, 0))
        intBedLen = GetMaxBedLen(lng部门ID, Not mbln护士站)
        strSQL = _
            "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄,B.入院日期,B.出院日期," & _
            " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
            " B.出院科室ID as 科室ID,C.名称 as 护理等级,B.状态,B.数据转出," & _
            " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
            " From 病人信息 A,病案主页 B,收费项目目录 C" & _
            " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+) And A.病人ID=[1] And B.主页ID=[2]"
        strSQL = strSQL & " Union " & _
            "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄,B.入院日期,B.出院日期," & _
            " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
            " B.出院科室ID as 科室ID,C.名称 as 护理等级,B.状态,B.数据转出," & _
            " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
            " From 病人信息 A,病案主页 B,收费项目目录 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0" & _
            " And B.出院日期 is NULL And B.状态<>3 And B.护理等级ID=C.ID(+)" & _
            IIF(mbln护士站, " And B.当前病区ID=[3]", " And B.出院科室ID=[3]") & _
            IIF(Not mbln护士站 And InStr(mstrPrivs, "本科病人") = 0, " And B.住院医师=[4]", "") & _
            " Order by 床号"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng部门ID, UserInfo.姓名)
    Else
        cmdPati.Visible = False
        txtPati.Locked = True
        txtPati.BackColor = Me.BackColor
    End If
    
    '下面判断权限备用
    If Not (mbln护士站 Or mlng前提ID <> 0 Or mbln会诊) Then
        str科室IDs = GetUser科室IDs(True)
    End If
    
    lvwPati.ListItems.Clear
    If Not rsPati.EOF Then
        For i = 1 To rsPati.RecordCount
            If mbln护士站 Or mlng前提ID <> 0 Or mbln会诊 Then
                blnDo = True
            Else
                '判断医生站下医嘱的权限
                blnDo = False
                If Nvl(rsPati!住院医师) = UserInfo.姓名 Then
                    blnDo = True '当前医生经治病人,可以下医嘱
                Else
                    If InStr("," & str科室IDs & ",", "," & rsPati!科室ID & ",") > 0 Then
                        '是否有本科病人医嘱下达权限
                        If InStr(";" & mstrPrivs & ";", ";本科医嘱下达;") > 0 _
                            Or InStr(";" & mstrPrivs & ";", ";全院医嘱下达;") > 0 Then
                            blnDo = True
                        End If
                    Else
                        '是否有全院病人医嘱下达权限
                        If InStr(";" & mstrPrivs & ";", ";全院医嘱下达;") > 0 Then
                            blnDo = True
                        End If
                    End If
                End If
                If rsPati!病人ID = mlng病人ID And rsPati!主页ID = mlng主页ID Then
                    blnDo = True
                End If
            End If
            If blnDo Then
                Set objItem = lvwPati.ListItems.Add(, "_" & rsPati!病人ID & "_" & rsPati!主页ID, rsPati!姓名, , "Pati")
                objItem.SubItems(1) = Nvl(rsPati!住院号)
                objItem.SubItems(2) = Nvl(rsPati!床号)
                objItem.SubItems(3) = Nvl(rsPati!住院医师)
                objItem.SubItems(4) = Nvl(rsPati!性别)
                objItem.SubItems(5) = Nvl(rsPati!年龄)
                objItem.SubItems(6) = Nvl(rsPati!费别)
                objItem.SubItems(7) = Nvl(rsPati!护理等级)
                objItem.SubItems(8) = Format(rsPati!入院日期, "MM-dd HH:mm")
                objItem.SubItems(9) = Format(Nvl(rsPati!出院日期), "MM-dd HH:mm")
                objItem.SubItems(10) = Nvl(rsPati!医疗付款方式)
                
                objItem.Tag = rsPati!病人ID
                objItem.ListSubItems(1).Tag = rsPati!主页ID
                objItem.ListSubItems(2).Tag = rsPati!科室ID
                objItem.ListSubItems(3).Tag = Nvl(rsPati!险类, 0)
                objItem.ListSubItems(4).Tag = Format(rsPati!入院日期, "yyyy-MM-dd HH:mm:ss")
                objItem.ListSubItems(5).Tag = Format(Nvl(rsPati!出院日期), "yyyy-MM-dd HH:mm:ss")
                objItem.ListSubItems(6).Tag = Nvl(rsPati!状态, 0)
                
                '用于历史数据查询
                objItem.ListSubItems(9).Tag = Nvl(rsPati!数据转出, 0)
                
                '保险病人用红色显示
                If Not IsNull(rsPati!险类) Then
                    objItem.ForeColor = vbRed
                    For j = 1 To objItem.ListSubItems.Count
                        objItem.ListSubItems(j).ForeColor = vbRed
                    Next
                End If
                
                '显示初始病人的信息
                If rsPati!病人ID = mlng病人ID And rsPati!主页ID = mlng主页ID Then
                    With objItem
                        txtPati.ForeColor = .ForeColor
                        txtPati.Text = .Text
                        mstr性别 = .SubItems(4)
                        mlng病人科室id = Val(.ListSubItems(2).Tag)
                        mint险类 = Val(.ListSubItems(3).Tag)
                        mstr付款码 = Get医疗付款码(.SubItems(10))
                        lblPati.Caption = "住院号:" & .SubItems(1) & "　床号:" & Trim(.SubItems(2)) & _
                            "　性别:" & .SubItems(4) & "　年龄:" & .SubItems(5) & _
                            "　费别:" & .SubItems(6) & "　付款方式:" & .SubItems(10)
                        .Selected = True '一定要选中当前病人
                        mbln提醒对码 = True
                    End With
                    mint年龄 = GetPatiYear(mlng病人ID)
                    Call Show费用信息
                    
                    '产科才有婴儿医嘱
                    Call SetBabyVisible(mlng病人科室id)
                    
                    '特殊医嘱提醒
                    strMsg = ExistsSpecAdvice(mlng病人ID, mlng主页ID)
                    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                End If
            End If
            rsPati.MoveNext
        Next
    End If
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetBabyVisible(ByVal lng科室ID As Long)
'功能：根据科室性质设置婴儿医嘱是否可以选择
'说明：产科才有婴儿医嘱
    If DeptIsWoman(lng科室ID) Then
        lbl婴儿.Visible = True
        cbo婴儿.Visible = True
    Else
        Call zlControl.CboSetIndex(cbo婴儿.Hwnd, 0)
        cbo婴儿.Tag = 0
        lbl婴儿.Visible = False
        cbo婴儿.Visible = False
    End If
End Sub

Private Function Show费用信息() As Boolean
'功能：获取当前病人的费用信息
'参数：bytType:0-费用余额,1-预交余额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select 费用余额,预交余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
        " Union ALL" & _
        " Select 0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2]"
    strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额,Sum(预结费用) as 预结费用 From (" & strSQL & ")"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        stbThis.Panels(2).Text = _
            "预交余额:" & FormatEx(Nvl(rsTmp!预交余额, 0), 2) & ",未结费用:" & FormatEx(Nvl(rsTmp!费用余额, 0), 2) & _
            IIF(Nvl(rsTmp!预结费用, 0) <> 0, ",预结费用:" & FormatEx(Nvl(rsTmp!预结费用, 0), 2), "") & _
            ",剩余款:" & FormatEx(Nvl(rsTmp!预交余额, 0) - Nvl(rsTmp!费用余额, 0) + Nvl(rsTmp!预结费用, 0), 2)
    Else
        stbThis.Panels(2).Text = ""
    End If
    Show费用信息 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPreRow(ByVal lngRow As Long) As Long
'功能：取上一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'功能：取下一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'功能：获取组ID相同的一组医嘱行号范围(注意考虑一并给药中的空行)
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_相关ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function CheckAdviceInsure(ByVal lngRow As Long) As String
'功能：医保病人下达医嘱时，医嘱录入后，对医嘱涉及的计价项目的保险对码情况进行检查
'参数：lngRow=可见行，一组医嘱的第一行
'返回：提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim strSQL As String, i As Long
    
    If gint医保对码 = 0 Or mint险类 = 0 Or Not mbln提醒对码 Then Exit Function
    If gclsInsure.GetCapability(support允许不设置医保项目, , mint险类) Then Exit Function
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    With vsAdvice
        '为利用索引,用Union方式
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                If InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 Then
                    '药品无对应关系,只处理按规格下达时
                    If Val(.TextMatrix(i, COL_收费细目ID)) <> 0 Then
                        strSQL = strSQL & " Union ALL " & _
                            "Select " & Val(.TextMatrix(i, COL_收费细目ID)) & " as 收费项目ID From Dual"
                    End If
                Else
                    '包含了收费数量为0的
                    strSQL = strSQL & " Union ALL " & _
                        "Select 收费项目ID From 诊疗收费关系 Where 诊疗项目ID=" & Val(.TextMatrix(i, COL_诊疗项目ID))
                End If
            End If
        Next
    End With
    strSQL = Mid(strSQL, 12)
    If strSQL = "" Then Exit Function
    
    strSQL = "Select Distinct C.名称,B.收费细目ID" & _
        " From (" & strSQL & ") A,保险支付项目 B,收费项目目录 C" & _
        " Where A.收费项目ID=B.收费细目ID(+) And A.收费项目ID=C.ID" & _
        " And B.险类(+)=" & mint险类
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'IN
    strSQL = "": i = 0
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!收费细目ID) Then
            If i = 8 Then
                strSQL = strSQL & vbCrLf & "… …"
                Exit Do
            End If
            strSQL = strSQL & vbCrLf & "●" & rsTmp!名称
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSQL <> "" Then
        CheckAdviceInsure = "当前病人是医保病人，但医嘱的以下计价项目没有设置对应的保险项目！" & vbCrLf & vbCrLf & _
            "医嘱内容：" & vbCrLf & Left(vsAdvice.TextMatrix(lngRow, COL_医嘱内容), 50) & vbCrLf & vbCrLf & "计价项目：" & strSQL
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TimeIs长嘱次日(ByVal lngRow As Long) As Boolean
'功能：判断指定医嘱行的开始执行时间是否按次日生效设置的
    With vsAdvice
        TimeIs长嘱次日 = gbln长期医嘱次日生效 And .TextMatrix(lngRow, COL_期效) = "长嘱" _
            And Format(.Cell(flexcpData, lngRow, COL_开始时间), "HH:mm") = "00:00"
    End With
End Function

Private Function GetDefaultTime(lngRow As Long) As String
'功能：获取新开医嘱的缺省开始时间
'说明：
'      长嘱次日生效时,缺省时间为次日零点或当日零点(补录)。
'      最近一条有效时间为当天，且间隔现在在半小时以内，则与该条相同
'      如果没有,则取最近新开(补录时为新开或修改)的一条的时间
'      如果没有,则取当前时间
    Dim curDate As Date, strDate As String
    Dim int期效 As Integer, i As Long
    
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '当前行的期效:可能才输入,也可能已输入
        If .RowData(lngRow) <> 0 Then
            int期效 = IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1)
        Else
            int期效 = cbo期效.ListIndex
        End If
        
        If int期效 = 0 And gbln长期医嘱次日生效 Then
            If Not mbln补录 Then
                strDate = Format(curDate + 1, "yyyy-MM-dd 00:00")
            Else
                strDate = Format(curDate, "yyyy-MM-dd 00:00")
            End If
        Else
            '先从当前行向回找:跳过缺省为次日生效的时间
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) Then
                    If Not TimeIs长嘱次日(i) Then
                        If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                            If DateAdd("n", 30, CDate(.Cell(flexcpData, i, COL_开始时间))) >= curDate Then
                                strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
                
            '再从最后行向回找:跳过缺省为次日生效的时间
            If strDate = "" Then
                For i = .Rows - 1 To lngRow + 1 Step -1
                    If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) Then
                        If Not TimeIs长嘱次日(i) Then
                            If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                                If DateAdd("n", 30, CDate(.Cell(flexcpData, i, COL_开始时间))) >= curDate Then
                                    strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        If strDate = "" Then
            '先从当前行向回找
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) _
                    And (Val(.TextMatrix(i, COL_EDIT)) = 1 And Not mbln补录 _
                         Or InStr(",1,2,", Val(.TextMatrix(i, COL_EDIT))) > 0 And mbln补录) Then
                    If Not TimeIs长嘱次日(i) Then
                        strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                End If
            Next
            '再从最后行向回找
            If strDate = "" Then
                For i = .Rows - 1 To lngRow + 1 Step -1
                    If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) _
                        And (Val(.TextMatrix(i, COL_EDIT)) = 1 And Not mbln补录 _
                             Or InStr(",1,2,", Val(.TextMatrix(i, COL_EDIT))) > 0 And mbln补录) Then
                        If Not TimeIs长嘱次日(i) Then
                            strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End With
    If strDate = "" Then strDate = Format(curDate, "yyyy-MM-dd HH:mm")
    GetDefaultTime = strDate
End Function

Private Function GetCurRow序号(lngRow As Long) As Long
'功能：获取指定行可用的的序号
'参数：lngRow=要取序号的行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng序号 As Long, i As Long
    Dim lng序号1 As Long, lng序号2 As Long
            
    '取之后最近一个有效序号,直接使用
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                lng序号 = Val(vsAdvice.TextMatrix(i, COL_序号))
                Exit For
            End If
        End If
    Next
    If lng序号 = 0 Then
        '后面没有,则取数据库之中的最大序号与之前的最大序号比较
        On Error GoTo errH
        strSQL = "Select Max(序号) as 序号 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(婴儿,0)=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cbo婴儿.ListIndex)
        If Not rsTmp.EOF Then lng序号1 = Nvl(rsTmp!序号, 0)
        On Error GoTo 0
        
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                    And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                    lng序号2 = Val(vsAdvice.TextMatrix(i, COL_序号))
                    Exit For
                End If
            End If
        Next
        
        If lng序号1 > lng序号2 Then
            lng序号 = lng序号1
        Else
            lng序号 = lng序号2
        End If

        If lng序号 <> 0 Then lng序号 = lng序号 + 1 '最大序号+1
    End If
    If lng序号 = 0 Then lng序号 = 1
    GetCurRow序号 = lng序号
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet医嘱序号(lngRow As Long, intStep As Integer)
'功能：将当前病人医嘱记录中序号前移或后移
'参数：lngRow=起始调整行,intStep=调整步长,如1或-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                vsAdvice.TextMatrix(i, COL_序号) = Val(vsAdvice.TextMatrix(i, COL_序号)) + intStep
                If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 3 '标志修改了序号
                End If
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'功能：指定的医嘱删除处理
    Dim lngBegin As Long, lngEnd As Long
    Dim lng相关ID As Long, blnGroup As Boolean
    Dim lng医嘱ID As Long, i As Integer
    
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            lng医嘱ID = vsAdvice.RowData(lngRow)
            lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
            blnGroup = RowIn一并给药(lngRow)
            If blnGroup Then
                '先删除一并给药中的空行(一定要删)
                Call Get一并给药范围(lng相关ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '必须反向
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                '删除之后当前行号可能变了
                lngRow = vsAdvice.FindRow(lng医嘱ID, lngBegin)
                
                '一并给药只删除当前行
                Call DeleteRow(lngRow)
            Else
                '单独的成药：删除给药途径行及当前行
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            Call Delete检查手术(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn配方行(lngRow) Then
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '删除当前行(中药用法行)
            Call DeleteRow(lngRow)
        ElseIf RowIn检验行(lngRow) Then
            lngRow = Delete检验组合(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '标记为未保存
    Else
        '空行直接删除
        Call DeleteRow(lngRow)
    End If
    
    '重新定位行
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'功能：删除表格中的一行,但不改变当前行
'参数：blnClear=是否仅清除该行内容,不删除
'      blnDelID=是否记录要删除的医嘱ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '调整序号
            Call AdviceSet医嘱序号(lngRow + 1, -1)
            
            '记录要删除的ID(除了才新增的)
            If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 And blnDelID Then
                mstrDelIDs = mstrDelIDs & "," & .RowData(lngRow)
            End If
        End If
            
        '如果为行1且仅剩行1或仅清除,则保留
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '清除该行数据
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '文字
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '数据
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '粗体
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '文字色
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '固定列文字色
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '固定列背景色
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '单元图片
            Set .Cell(flexcpPicture, lngRow, COL_警示) = Nothing 'Pass警示灯
            
            '单元格边框
            .Select lngRow, .FixedCols, lngRow, COL_标志
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '因为有删除行,所以调用程序肯定有行定位,所以不必恢复行
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete检查手术(ByVal lngRow As Long)
'功能：1.删除检查组合项目的部位行
'      2.删除手术项目的附加手术行及麻醉项目行
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID) '不一定有,所以用查找
    If i <> -1 Then
        lngBegin = i
        For i = lngBegin To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngEnd To lngBegin Step -1
            Call DeleteRow(i)
        Next
    End If
End Sub

Private Function Delete中药配方(ByVal lngRow As Long) As Long
'功能：删除中药配方的组成味药及煎法行
'参数：lngRow=中药配方用法行(可见)
'返回：删除之后重新定位的当前行(中药用法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到中药用法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete中药配方 = vsAdvice.Row
End Function

Private Function Delete检验组合(ByVal lngRow As Long) As Long
'功能：删除一并采集的多个检验项目行
'参数：lngRow=采集方法行(可见)
'返回：删除之后重新定位的当前行(采集方法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到采集方法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete检验组合 = vsAdvice.Row
End Function

Private Function Get检查部位IDs(ByVal lngRow As Long) As String
'功能：获取指定行的检查部位ID串
'返回："部位ID1,部位ID2,..."
    Dim strTmp As String, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
            Else
                Exit For
            End If
        Next
    End If
    Get检查部位IDs = Mid(strTmp, 2)
End Function

Private Function Get手术附加IDs(ByVal lngRow As Long) As String
'功能：获取指定手术行的附加手术及麻醉项目ID串
'返回："手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
    Dim strTmp As String, lng麻醉ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_类别) = "G" Then
                    lng麻醉ID = Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get手术附加IDs = Mid(strTmp, 2) & ";" & IIF(lng麻醉ID = 0, "", lng麻醉ID)
End Function

Private Function Get中药配方IDs(ByVal lngRow As Long) As String
'功能：获取中药配方的组成味药及煎法ID串
'返回："中药ID1,单量1,脚注1;中药ID2,单量2,脚注2;...|煎法ID"
    Dim lng煎法ID As Long, str中药IDs As String, i As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_类别) = "E" Then
                    lng煎法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                ElseIf .TextMatrix(i, COL_类别) = "7" Then
                    str中药IDs = Val(.TextMatrix(i, COL_诊疗项目ID)) & "," & _
                        .TextMatrix(i, COL_单量) & "," & .TextMatrix(i, COL_医生嘱托) & _
                        ";" & str中药IDs
                End If
            Else
                Exit For
            End If
        Next
    End With
    Get中药配方IDs = Mid(str中药IDs, 1, Len(str中药IDs) - 1) & "|" & lng煎法ID
End Function

Private Function Get检验组合IDs(ByVal lngRow As Long) As String
'功能：获取一并采集的检验组合项目ID及标本
'返回："项目ID1,项目ID2,...;检验标本"
    Dim str项目IDs As String, str标本 As String, i As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                str项目IDs = Val(.TextMatrix(i, COL_诊疗项目ID)) & "," & str项目IDs
                str标本 = .TextMatrix(i, COL_标本部位)
            Else
                Exit For
            End If
        Next
    End With
    Get检验组合IDs = Left(str项目IDs, Len(str项目IDs) - 1) & ";" & str标本
End Function

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, COL_类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "C" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, COL_类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "7" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn一并给药(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中
'参数：lngRow=可见的行,可能是空行
'说明：一并给药的范围中可能存在空行
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng相关ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngNextRow, COL_相关ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 _
            And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            
            lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
        End If
    End With
    RowIn一并给药 = blnGroup
End Function

Private Function Calc终止时间(ByVal lngRow As Long) As String
'功能：根据中药配方长嘱的开始时间，付数，频次，执行时间，确定终止时间
'参数：lngRow=中药配方的一个中药组成味药行
'返回：最后一次执行的时间点作为终止时间(yyyy-MM-dd HH:mm)
    Dim vStartTime As Date, vCurTime As Date, vTmpTime As Date
    Dim arrTime() As String, strTime As String, strLastTime As String
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim int付数 As Long, strTmp As String, i As Integer
    
    '这些条件必须具备才能计算
    If Not (IsDate(vsAdvice.Cell(flexcpData, lngRow, COL_开始时间)) _
        And Val(vsAdvice.TextMatrix(lngRow, COL_总量)) <> 0 _
        And vsAdvice.TextMatrix(lngRow, COL_频率) <> "" _
        And vsAdvice.TextMatrix(lngRow, COL_执行时间) <> "") Then Exit Function
    
    strTime = vsAdvice.TextMatrix(lngRow, COL_执行时间)
    int频率次数 = vsAdvice.TextMatrix(lngRow, COL_频率次数)
    int频率间隔 = vsAdvice.TextMatrix(lngRow, COL_频率间隔)
    str间隔单位 = vsAdvice.TextMatrix(lngRow, COL_间隔单位)
    If Not ExeTimeValid(strTime, int频率次数, int频率间隔, str间隔单位) Then Exit Function
    
    arrTime = Split(strTime, "-")
    vStartTime = Format(vsAdvice.Cell(flexcpData, lngRow, COL_开始时间), "yyyy-MM-dd HH:mm")
    vCurTime = vStartTime
    int付数 = Val(vsAdvice.TextMatrix(lngRow, COL_总量))
    
    If str间隔单位 = "周" Then
        vCurTime = GetWeekBase(vStartTime) '该周星期一作为基准时间
        Do While int付数 > 0
            '1/8:00-1/15:00-2/9:00
            For i = 1 To int频率次数
                vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                    strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                Else
                    strTmp = Split(arrTime(i - 1), "/")(1)
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm")
                If vTmpTime >= vStartTime Then
                    strLastTime = Format(vTmpTime, "yyyy-MM-dd HH:mm")
                    int付数 = int付数 - 1
                    If int付数 = 0 Then Exit Do
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str间隔单位 = "天" Then
        Do While int付数 > 0
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If InStr(arrTime(i - 1), ":") = 0 Then
                        strTmp = arrTime(i - 1) & ":00"
                    Else
                        strTmp = arrTime(i - 1)
                    End If
                    vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm")
                    If vTmpTime >= vStartTime Then
                        strLastTime = Format(vTmpTime, "yyyy-MM-dd HH:mm")
                        int付数 = int付数 - 1
                        If int付数 = 0 Then Exit Do
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm")
                    If vTmpTime >= vStartTime Then
                        strLastTime = Format(vTmpTime, "yyyy-MM-dd HH:mm")
                        int付数 = int付数 - 1
                        If int付数 = 0 Then Exit Do
                    End If
                Next
            End If
            vCurTime = vCurTime + int频率间隔
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        Do While int付数 > 0
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                strLastTime = Format(vTmpTime, "yyyy-MM-dd HH:mm")
                int付数 = int付数 - 1
                If int付数 = 0 Then Exit Do
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    Calc终止时间 = Format(strLastTime, "yyyy-MM-dd HH:mm")
End Function

Private Function Check过敏试验(ByVal lng药名ID As Long, ByVal str名称 As String) As String
'功能：检查西成药，中成药的过敏试验
'参数：lng药名ID=药品诊疗项目ID
'      str名称=药品名称,用于提示
'返回：为空表示通过
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    '取有效时间内的最后一次过敏结果登记
    strSQL = "Select 药物名,结果,记录时间 From 病人过敏记录" & _
        " Where 病人ID=[1] And 药物ID=[2] And Trunc(记录时间)>=Trunc(Sysdate-[3])" & _
        " Order by 记录时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, lng药名ID, gint过敏登记有效天数)
    If Not rsTmp.EOF Then
        '有过敏结果登记记录,根据是否阳性决定是否提示
        If Nvl(rsTmp!结果, 0) = 1 Then
            strMsg = "该病人在" & Format(rsTmp!记录时间, "M月d日") & "的过敏实验中对""" & Nvl(rsTmp!药物名, str名称) & """过敏(+)。" & _
                vbCrLf & vbCrLf & "是否仍然使用该药品？"
        Else
            strMsg = "" '为阴性,通过
        End If
    Else
        '无过敏结果登记记录,则先看该药品是否需要皮试
        strSQL = "Select A.用法ID,B.名称" & _
            " From 诊疗用法用量 A,诊疗项目目录 B" & _
            " Where A.用法ID=B.ID And A.性质=0 And A.项目ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药名ID)
        If Not rsTmp.EOF Then
            '要求皮试,则提示皮试
            strMsg = "在对病人使用""" & str名称 & """前，要求先进行""" & rsTmp!名称 & """，" & vbCrLf & _
                "但没有发现有效的过敏试验结果，是否仍然使用该药品？"
        Else
            strMsg = "" '没有皮试要求,通过
        End If
    End If
    Check过敏试验 = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集,lngRow=当前输入行
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, intType As Integer
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim bln中医 As Boolean, i As Long
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) '取上一有效行,某些内容缺省与上一行相同
    lngNextRow = GetNextRow(lngRow) '取下一有效行
    
    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    txt医嘱内容.Text = rsInput!名称 '暂时显示
    
    '药品处方职务检查(护士站在保存时检查)
    If Not mbln护士站 Then
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            strMsg = CheckOneDuty(rsInput!名称, Nvl(rsInput!处方职务ID), UserInfo.姓名, InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> "")
            If strMsg <> "" Then
                vsAdvice.Refresh
                MsgBox strMsg, vbInformation, gstrSysName
                vsAdvice.Refresh: Exit Function
            End If
        End If
    End If
    
    '出院医嘱诊断检查
    If mbln检查诊断 And rsInput!类别ID = "Z" And (Nvl(rsInput!项目特性) = "出院" Or Nvl(rsInput!项目特性) = "转院") Then
        bln中医 = Have部门性质(mlng病人科室id, "中医科")
        If Not ExistsDiagNoses(mlng病人ID, mlng主页ID, IIF(bln中医, "3,13", "3")) Then
            vsAdvice.Refresh
            MsgBox "病人的出院诊断还没有输入，请先输入病人的出院诊断再下达出院医嘱。", vbInformation, gstrSysName
            vsAdvice.Refresh: Exit Function
        End If
    End If
    
    '医保病人输入内容时的提示
    If InStr(",5,6,7,", rsInput!类别ID) > 0 And Not IsNull(rsInput!收费细目ID) Then 'mint险类 <> 0
        Call gclsInsure.GetItemInfo(mint险类, mlng病人ID, rsInput!收费细目ID) '非医保病人也要调
    End If
    
    With vsAdvice
        '检验项目：采集方法判断
        If rsInput!类别ID = "C" Then
            '所有数据中取一个缺省的采集方法,同时判断是否有采集方法数据
            lng用法ID = Get缺省用法ID(6, 2)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的标本采集方法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '缺省与上一行相同
            If lngPreRow <> -1 Then
                If RowIn检验行(lngPreRow) Then
                    lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
                End If
            End If
        End If
        
        '中药配方：给成与中药用法判断
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            If rsInput!类别ID = "8" Then
                If GetGroupCount(rsInput!诊疗项目ID, 2, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!名称 & """是一个中药配方，但没有设置有效的组成中药。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
                
                '部份药无效的提示
                strMsg = GetGroupNone(rsInput!诊疗项目ID, 2)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "配方""" & rsInput!名称 & """中以下药品已撤档或服务对象不匹配：" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "这些药品将不会出现在配方中。", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '所有数据中取一个缺省的中药用法,同时判断是否有中药用法数据
            lng用法ID = Get缺省用法ID(4, 2)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的中药用(服)法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '中药用法缺省与上一行相同
            If RowIn配方行(lngPreRow) Then
                lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
            End If
        End If
        
        '中西成药：给药途径判断
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
'            '所有数据中取一个缺省的给药途径,同时判断是否有给药途径数据
'            lng用法ID = Get缺省用法ID(2, 2)
'            If lng用法ID = 0 Then
'                .Refresh
'                MsgBox "没有可用的给药途径,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
'                .Refresh: Exit Function
'            End If
            '给药途径缺省与上一个行相同剂型的相同
            If lngPreRow <> -1 And Not IsNull(rsInput!药品剂型) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 And .TextMatrix(lngPreRow, COL_药品剂型) = Nvl(rsInput!药品剂型) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_相关ID)), lngPreRow + 1)
                    lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                End If
            End If
        End If
        
        '中西成药：过敏试验检查
        If InStr(",5,6,", rsInput!类别ID) > 0 And gint过敏登记有效天数 <> 0 Then
            str过敏 = Check过敏试验(rsInput!诊疗项目ID, rsInput!名称)
            If str过敏 <> "" Then
                .Refresh
                If MsgBox(str过敏, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .Refresh: Exit Function
                End If
            End If
        End If
        
        '中西成药：一并给药的判断
        blnGroup = RowIn一并给药(lngRow) Or tbr.Buttons("一并").Value = tbrPressed
        If blnGroup Then
            If rsInput!类别ID = "9" Then
                .Refresh
                MsgBox "不能在一并给药的药品中直接输入成套方案。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            If .RowData(lngRow) = 0 Then
                '一并给药中的待输入空行：只有插入在一并给药的中间,才能自动成为一并给药
                lngGroupRow = lngPreRow
            Else
                '一并给药中的药品行：可能是第一行或最后一行
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngGroupRow = lngPreRow
                Else
                    lngGroupRow = lngNextRow
                End If
            End If
            
            '一并给药的,类别，期效必须相同
            If Decode(rsInput!类别ID, "5", "Y", "6", "Y", "N") <> Decode(.TextMatrix(lngGroupRow, COL_类别), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "该组一并给药的药品必须都为西成药或中成药。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            If NeedName(cbo期效.Text) <> .TextMatrix(lngGroupRow, COL_期效) Then
                .Refresh
                MsgBox "该组一并给药的药品必须都为""" & .TextMatrix(lngGroupRow, COL_期效) & """。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
            lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID)) '一并给药的给药途径相同
            
            '检查一并给药的的给药途径是否适合于当前输入药品(非一并给药的缺省用法在输入函数中作了判断处理)
            If Not Check适用用法(lng用法ID, rsInput!诊疗项目ID, 2) Then
                .Refresh
                MsgBox "一并的给药途径为""" & .TextMatrix(i, COL_医嘱内容) & """，不适用于当前输入药品。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        End If
            
        '成套项目
        If rsInput!类别ID = "9" Then
            If GetGroupCount(rsInput!诊疗项目ID, 2) = 0 Then
                .Refresh
                MsgBox """" & rsInput!名称 & """是一个成套方案，但没有设置有效的组成项目。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!诊疗项目ID, 2)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '需要输入更多数据的一些项目
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        If rsInput!类别ID = "D" And Nvl(GetItemField("诊疗项目目录", rsInput!诊疗项目ID, "组合项目"), 0) = 1 Then
            '检查组合项目
            intType = 0
        ElseIf rsInput!类别ID = "F" Then
            '手术：需要输入麻醉项目，及可选择附加手术
            intType = 1
        ElseIf InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理)
            intType = 2
        ElseIf rsInput!类别ID = "C" Then
            '输入一并采集的多个检验项目及检验标本
            intType = 4
            strExtData = rsInput!诊疗项目ID & ";" & Nvl(rsInput!规格)
        End If
        If intType <> -1 Then
            frmAdviceEditEx.mstrPrivs = mstrPrivs
            frmAdviceEditEx.mlngHwnd = txt医嘱内容.Hwnd
            frmAdviceEditEx.mintType = intType
            frmAdviceEditEx.mint期效 = cbo期效.ListIndex
            frmAdviceEditEx.mstr性别 = mstr性别
            frmAdviceEditEx.mlng项目ID = IIF(rsInput!类别ID = "C", 0, rsInput!诊疗项目ID)
            frmAdviceEditEx.mstrExtData = IIF(rsInput!类别ID = "C", strExtData, "") '新输入项目
            
            frmAdviceEditEx.mbln护士站 = mbln护士站
            frmAdviceEditEx.mbln医保 = InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> ""
            
            On Error Resume Next
            frmAdviceEditEx.Show 1, Me
            On Error GoTo errH
            
            If Not frmAdviceEditEx.mblnOK Then Exit Function
            strExtData = frmAdviceEditEx.mstrExtData
        End If
    
        '修改已有项目时,先删除当前医嘱的内容
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '西成药、中成药
                If Not blnGroup Then
                    '单个成药删除给药途径行,并清除当前行
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    '一组成药时,只清除当前行
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '检查组合项目及手术项目
                '删除部位行或手术附加行(附加手术,麻醉项目)
                Call Delete检查手术(lngRow)
                '清除当前行
                Call DeleteRow(lngRow, True)
            ElseIf RowIn配方行(lngRow) Then
                '中药配方：顺序(序号)要求必须严格控制
                '删除组成味药及煎法行:删除之后重新定位的当前行
                lngRow = Delete中药配方(lngRow)
                '清除当前行(中药用法行)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn检验行(lngRow) Then
                '删除检验项目行:删除之后重新定位的当前行
                lngRow = Delete检验组合(lngRow)
                '清除当前行(采集方法行)
                Call DeleteRow(lngRow, True)
            Else
                '其它项目直接清除当前行内容
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '当前行新增医嘱
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理):处理之后重新定位的当前行
            lngRow = AdviceSet中药配方(rsInput!诊疗项目ID, lngRow, lng用法ID, strExtData)
        ElseIf rsInput!类别ID = "9" Then
            '成套医嘱需要分解为多个项目加入
            Call AdviceSet成套项目(rsInput!诊疗项目ID, lngRow, strExtData)
        ElseIf rsInput!类别ID = "C" Then
            '检验组合
            lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData)
        Else
            '中、西成药，检查(组合)，手术(组合)，及其它诊疗项目
            Call AdviceSet诊疗项目(rsInput, lngRow, lng用法ID, lngGroupRow, strExtData)
            
            '自动设置一并给药
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                If Not RowIn一并给药(lngRow) Then
                    If tbr.Buttons("一并").Value = tbrPressed Then
                        '手工使一并给药
                        Call MergeRow(lngPreRow, lngRow) '本来就是显示当前行的内容,不用再强行RowChange
                    ElseIf lngPreRow <> -1 Then
                        '自动使一并给药
                        If .TextMatrix(lngPreRow, COL_类别) = rsInput!类别ID Then
                            If RowIn一并给药(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                tbr.Buttons("一并").Value = tbrPressed
                                Call MergeRow(lngPreRow, lngRow, False)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        '重新自动调整行高
        Call .AutoSize(COL_医嘱内容)
    End With
    mblnNoSave = True '标记为未保存
    
    '对保险对码进行检查
    strMsg = CheckAdviceInsure(vsAdvice.Row)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'功能：将两行设置为一并给药
'参数：lngRow1=前面行,可能本来已经属于一并给药
'      lngRow2=当前行
'说明：设置完成后,表格仍定位在原lngRow2的当前行
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            lngTmp = .RowData(lngRow2) '记录以再定位到当前行
            '先取消之前的一并给药
            If RowIn一并给药(lngRow1) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow1, COL_相关ID)), lngBegin, lngEnd)
                Call AdviceSet单独给药(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSet一并给药(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'功能：将指定行从一并给药中独立出来(该组一并给药必须至少包含三行)
'参数：lngRow=当前行,且为一并给药中的最后一药品行
'说明：设置完成后,表格仍定位在原lngRow的当前行
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '记录用于恢复定位当前行
        Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
        
        '先取消整个的一并给药
        Call AdviceSet单独给药(lngBegin, lngEnd)
        
        '再设置除最后行外的行为一并给药
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSet一并给药(lngBegin, lngEnd)
        
        '恢复当前行
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Private Sub AdviceSet成套项目(ByVal lng成套ID As Long, ByVal lngRow As Long, Optional ByVal str序号 As String)
'功能：输入成套项目(包括一并给药,检查组合,手术附加,中药配方)
'参数：lngRow=空的输入行(可能是插入的新行,但不位于一并给药中间)
    Dim rsItems As New ADODB.Recordset
    Dim rs规格 As New ADODB.Recordset
    Dim rs疗程 As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng序号 As Long
    Dim lngPreRow As Long, vCurDate As Date, lngTmp As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim bln给药途径 As Boolean, bln采集方法 As Boolean, int频率性质 As Integer
    Dim bln中药用法 As Boolean, bln中药煎法 As Boolean, bln配方 As Boolean
    Dim lng西药房ID As Long, lng成药房ID As Long, lng中药房ID As Long
    Dim lng相关ID As Long, str适用范围 As String, str频率 As String
    Dim str医生 As String, lng医生ID As Long, blnFirst As Boolean
    Dim lng倍数 As Long, vBookMark As Variant, str药房IDs As String
    Dim sng天数 As Single, strSQL序号 As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '产生序号过滤串
    If str序号 <> "" Then
        If Left(str序号, 1) = "+" Then
            strSQL序号 = " And Instr([2],','||A.序号||',')>0"
        ElseIf Left(str序号, 1) = "-" Then
            strSQL序号 = " And Instr([2],','||A.序号||',')=0"
        End If
    End If
    
    '药品规格信息:虽然存了收费细目ID,但长嘱可能没存,以前的数据也没存
    strSQL = "Select A.序号,B.药名ID,B.药品ID,B.剂量系数,B.住院包装,B.住院单位," & _
        " B.可否分零,C.编码,Nvl(D.名称,C.名称) as 名称,C.规格,C.产地" & _
        " From 诊疗项目组合 A,药品规格 B,收费项目目录 C,收费项目别名 D" & _
        " Where A.诊疗项目ID=B.药名ID And B.药品ID=C.ID" & _
        " And C.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[3]" & _
        " And A.诊疗组合ID=[1]" & strSQL序号 & _
        " Order by A.序号,C.编码"
    Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",", IIF(gbln商品名, 3, 1))
    
    '成药疗程信息(因成套中无直接对应配方,中药取不到疗程)
    strSQL = "Select Distinct A.诊疗项目ID,C.疗程" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,诊疗用法用量 C" & _
        " Where A.诊疗项目ID=B.ID And B.类别 IN('5','6')" & _
        " And A.诊疗项目ID=C.项目ID And A.诊疗组合ID=[1]" & strSQL序号
    Set rs疗程 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    
    '按序号排列后应该与医嘱编辑时的次序一致
    strSQL = "Select A.期效,A.序号,A.相关序号,A.诊疗项目ID,A.收费细目ID,A.总给予量,A.单次用量," & _
        " A.医生嘱托,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行科室ID,B.类别,B.名称," & _
        " B.计算单位,Nvl(A.标本部位,B.标本部位) as 标本部位,A.时间方案,Nvl(A.执行性质,B.执行科室) as 执行性质," & _
        " B.计价性质,B.操作类型,B.计算方式,B.执行频率,B.录入限量,C.处方限量,C.处方职务,C.毒理分类,C.药品剂型,C.品种医嘱" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,药品特性 C" & _
        " Where A.诊疗项目ID=B.ID And A.诊疗项目ID=C.药名ID(+)" & _
        " And A.诊疗组合ID=[1]" & strSQL序号 & _
        " Order by A.序号"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        
        lngPreRow = GetPreRow(lngRow) '前一参照行
        intCount = 0 '已经设置的行数
        lng序号 = GetCurRow序号(lngRow) '起始序号
        vCurDate = zlDatabase.Currentdate
        
        For i = 1 To rsItems.RecordCount
            lngCurRow = lngRow + intCount
            If lngCurRow > lngRow Then .AddItem "", lngCurRow
             
            '记录相对ID
            .RowData(lngCurRow) = -1 * rsItems!序号
            If Not IsNull(rsItems!相关序号) Then
                .TextMatrix(lngCurRow, COL_相关ID) = -1 * rsItems!相关序号
            End If
            
            .TextMatrix(lngCurRow, COL_EDIT) = 1 '新增的
            .Cell(flexcpData, lngCurRow, COL_EDIT) = lng成套ID '记录相关的成套项目
            
            .TextMatrix(lngCurRow, COL_婴儿) = cbo婴儿.ListIndex
            .TextMatrix(lngCurRow, COL_序号) = lng序号 + intCount
            .TextMatrix(lngCurRow, COL_状态) = 1 '新开
            .TextMatrix(lngCurRow, COL_期效) = IIF(Nvl(rsItems!期效, 0) = 0, "长嘱", "临嘱")
            .TextMatrix(lngCurRow, COL_类别) = rsItems!类别
            
            If IsDate(txt开始时间.Text) Then
                .TextMatrix(lngCurRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngCurRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
            End If
            
            .TextMatrix(lngCurRow, COL_诊疗项目ID) = rsItems!诊疗项目ID
            .TextMatrix(lngCurRow, COL_名称) = rsItems!名称
            .TextMatrix(lngCurRow, COL_标本部位) = Nvl(rsItems!标本部位)

            '其它
            .TextMatrix(lngCurRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
            .TextMatrix(lngCurRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
            .TextMatrix(lngCurRow, COL_操作类型) = Nvl(rsItems!操作类型)
            .TextMatrix(lngCurRow, COL_毒理分类) = Nvl(rsItems!毒理分类)
            .TextMatrix(lngCurRow, COL_药品剂型) = Nvl(rsItems!药品剂型)
            If InStr(",5,6,7,", rsItems!类别) > 0 Then
                .TextMatrix(lngCurRow, COL_处方限量) = Nvl(rsItems!处方限量)
            Else
                .TextMatrix(lngCurRow, COL_处方限量) = Nvl(rsItems!录入限量)
            End If
            .TextMatrix(lngCurRow, COL_处方职务) = Nvl(rsItems!处方职务)
            
            '药品规格信息:中草药肯定有,成药按单量与剂量单位自动匹配
            lng倍数 = 0: vBookMark = 0
            '不管成套本身是否记录规格，根据当前设置来
            If rsItems!类别 = "7" Or (InStr(",5,6,", rsItems!类别) > 0 _
                And (Nvl(rsItems!期效, 0) = 1 Or gbln药品按规格下医嘱 And Nvl(rsItems!品种医嘱, 0) = 0)) Then
'                '成药长嘱：记录了规格的保持原样；没记录规格的，如果固定按品种下达则不管
'                If Not (InStr(",5,6,", rsItems!类别) > 0 And Nvl(rsItems!期效, 0) = 0 _
'                    And IsNull(rsItems!收费细目ID) And Nvl(rsItems!品种医嘱, 0) = 1) Then
                If Not IsNull(rsItems!收费细目ID) Then
                    rs规格.Filter = "药品ID=" & rsItems!收费细目ID
                Else
                    rs规格.Filter = "药名ID=" & rsItems!诊疗项目ID
                End If
                If Not rs规格.EOF Then
                    If IsNull(rsItems!收费细目ID) Then
                        '取剂量系数为单量的最小整倍数的那一个规格
                        If CInt(Nvl(rsItems!单次用量, 0)) <> 0 Then
                            Do While Not rs规格.EOF
                                If rs规格!剂量系数 / rsItems!单次用量 = Int(rs规格!剂量系数 / rsItems!单次用量) Then
                                    If rs规格!剂量系数 / rsItems!单次用量 < lng倍数 Or lng倍数 = 0 Then
                                        vBookMark = rs规格.Bookmark
                                        lng倍数 = rs规格!剂量系数 / rsItems!单次用量
                                    End If
                                End If
                                rs规格.MoveNext
                            Loop
                            If vBookMark <> 0 Then rs规格.Bookmark = vBookMark
                        End If
                        If rs规格.EOF Then rs规格.MoveFirst
                    End If
                    .TextMatrix(lngCurRow, COL_名称) = Nvl(rs规格!名称)
                    .TextMatrix(lngCurRow, COL_收费细目ID) = rs规格!药品ID
                    .TextMatrix(lngCurRow, COL_剂量系数) = Nvl(rs规格!剂量系数)
                    .TextMatrix(lngCurRow, COL_住院包装) = Nvl(rs规格!住院包装)
                    .TextMatrix(lngCurRow, COL_住院单位) = Nvl(rs规格!住院单位)
                    .TextMatrix(lngCurRow, COL_可否分零) = Nvl(rs规格!可否分零, 0)
                End If
            End If
                                
            '判断是否特定行
            bln给药途径 = False: bln采集方法 = False
            bln中药用法 = False: bln中药煎法 = False: bln配方 = False
            If rsItems!类别 = "E" Then
                If IsNull(rsItems!相关序号) Then
                    If Val(.TextMatrix(lngCurRow - 1, COL_相关ID)) = .RowData(lngCurRow) Then
                        If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_类别)) > 0 Then
                            bln给药途径 = True
                        ElseIf .TextMatrix(lngCurRow - 1, COL_类别) = "C" Then
                            bln采集方法 = True
                        Else
                            bln中药用法 = True
                        End If
                    End If
                Else
                    bln中药煎法 = True
                End If
            End If
            If rsItems!类别 = "7" Or bln中药煎法 Or bln中药用法 Then bln配方 = True
                    
            '获取当前项目的适用范围
            If bln采集方法 Then
                '采集方法以检验项目的为准
                lngTmp = .FindRow(CStr(.RowData(lngCurRow)), , COL_相关ID)
                int频率性质 = .TextMatrix(lngTmp, COL_频率性质)
            Else
                int频率性质 = Nvl(rsItems!执行频率, 0)
            End If
            If bln配方 Then
                str适用范围 = 2 '中药配方(包括煎法,用法)用中医
'            ElseIf bln采集方法 Then
'                str适用范围 = -1 '设置与检验项目相同:一次性
            ElseIf int频率性质 = 1 Then
                str适用范围 = -1 '一次性
            ElseIf int频率性质 = 2 Then
                str适用范围 = -2 '持续性
            ElseIf int频率性质 = 0 Then '可选频率
                If Nvl(rsItems!期效, 0) = 1 Then
                    str适用范围 = "1,-1" '临嘱可能为一次性(光名称不能唯一区分)
                Else
                    str适用范围 = 1
                End If
            End If
                    
            '频率,频率次数,频率间隔,间隔单位
            .TextMatrix(lngCurRow, COL_频率性质) = int频率性质
            If Not IsNull(rsItems!执行频次) Then
                .TextMatrix(lngCurRow, COL_频率) = rsItems!执行频次
                .TextMatrix(lngCurRow, COL_频率次数) = Nvl(rsItems!频率次数, 0)
                .TextMatrix(lngCurRow, COL_频率间隔) = Nvl(rsItems!频率间隔, 0)
                .TextMatrix(lngCurRow, COL_间隔单位) = Nvl(rsItems!间隔单位)
                
'                Call Get频率信息_名称(rsItems!执行频次, int频率次数, int频率间隔, str间隔单位, str适用范围)
'                .TextMatrix(lngCurRow, COL_频率) = rsItems!执行频次
'                .TextMatrix(lngCurRow, COL_频率次数) = int频率次数
'                .TextMatrix(lngCurRow, COL_频率间隔) = int频率间隔
'                .TextMatrix(lngCurRow, COL_间隔单位) = str间隔单位
                
                '临嘱可选频率可能设置为了一次性
                If Nvl(rsItems!期效, 0) = 1 And int频率性质 = 0 And Nvl(rsItems!频率次数, 0) = 0 And Nvl(rsItems!频率间隔, 0) = 0 Then
                    .TextMatrix(lngCurRow, COL_频率性质) = 1
                End If
            Else '取缺省的
                If Nvl(rsItems!期效, 0) = 1 And int频率性质 = 0 Then
                    If mbln一次性 Then '临嘱缺省为一次性
                        str适用范围 = -1
                        .TextMatrix(lngCurRow, COL_频率性质) = 1
                    Else
                        str适用范围 = 1
                    End If
                End If
                Call Get缺省频率(str适用范围, str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngCurRow, COL_频率) = str频率
                .TextMatrix(lngCurRow, COL_频率次数) = int频率次数
                .TextMatrix(lngCurRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngCurRow, COL_间隔单位) = str间隔单位
            End If
            
            '单量
            .TextMatrix(lngCurRow, COL_单量) = FormatEx(Nvl(rsItems!单次用量), 5)
            If Nvl(rsItems!期效, 0) = 0 Then
                If InStr(",5,6,7,", rsItems!类别) > 0 Or InStr(",1,2,", Nvl(rsItems!计算方式, 0)) > 0 Then
                    .TextMatrix(lngCurRow, COL_单量单位) = Nvl(rsItems!计算单位)
                End If
            Else
                If InStr(",5,6,7,", rsItems!类别) > 0 Or (int频率性质 = 0 And InStr(",1,2,", Nvl(rsItems!计算方式, 0)) > 0) Then
                    .TextMatrix(lngCurRow, COL_单量单位) = Nvl(rsItems!计算单位)
                End If
            End If
            
            '总量
            If (InStr(",5,6,", rsItems!类别) > 0 And Nvl(rsItems!期效, 0) = 1) Or rsItems!类别 = "7" Then
                '成药临嘱(有对应规格)或中药配方
                If InStr(",5,6,", rsItems!类别) > 0 Then
                    .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow, COL_住院单位)
                    
                    sng天数 = msng天数
                    If mbln天数 Then
                        If .TextMatrix(lngCurRow, COL_间隔单位) = "周" Then
                            If 7 > sng天数 Then sng天数 = 7
                        ElseIf .TextMatrix(lngCurRow, COL_间隔单位) = "天" Then
                            If Val(.TextMatrix(lngCurRow, COL_频率间隔)) > sng天数 Then
                                sng天数 = Val(.TextMatrix(lngCurRow, COL_频率间隔))
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_间隔单位) = "小时" Then
                            If Val(.TextMatrix(lngCurRow, COL_频率间隔)) \ 24 > sng天数 Then
                                sng天数 = Val(.TextMatrix(lngCurRow, COL_频率间隔)) \ 24
                            End If
                        End If
                        If sng天数 = 0 Then sng天数 = 1
                    End If
                Else
                    .TextMatrix(lngCurRow, COL_总量单位) = "付"
                    sng天数 = 1
                End If
                
                If Not IsNull(rsItems!总给予量) Then
                    If InStr(",5,6,", rsItems!类别) > 0 Then
                        '转换为住院单位
                        .TextMatrix(lngCurRow, COL_总量) = FormatEx(rsItems!总给予量 / Val(.TextMatrix(lngCurRow, COL_住院包装)), 5)
                    Else
                        .TextMatrix(lngCurRow, COL_总量) = rsItems!总给予量
                    End If
                Else
                    '计算缺省总量
                    If .TextMatrix(lngCurRow, COL_频率) <> "" Then
                        If InStr(",5,6,", rsItems!类别) > 0 Then
                            rs疗程.Filter = "诊疗项目ID=" & rsItems!诊疗项目ID
                            If Not rs疗程.EOF Then
                                If Nvl(rs疗程!疗程, 1) > sng天数 Then
                                    sng天数 = Nvl(rs疗程!疗程, 1)
                                End If
                            End If
                        End If
                        
                        If InStr(",5,6,", rsItems!类别) > 0 Then
                            If (Val(.TextMatrix(lngCurRow, COL_单量)) <> 0 _
                                And Val(.TextMatrix(lngCurRow, COL_住院包装)) <> 0 _
                                And Val(.TextMatrix(lngCurRow, COL_剂量系数)) <> 0) Then
                                If Val(.TextMatrix(lngCurRow, COL_频率性质)) = 1 Then
                                    .TextMatrix(lngCurRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                            Val(.TextMatrix(lngCurRow, COL_单量)), 1, 1, 1, "天", "", _
                                            Val(.TextMatrix(lngCurRow, COL_剂量系数)), _
                                            Val(.TextMatrix(lngCurRow, COL_住院包装)), _
                                            Val(.TextMatrix(lngCurRow, COL_可否分零))), 5)
                                Else
                                    .TextMatrix(lngCurRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                            Val(.TextMatrix(lngCurRow, COL_单量)), sng天数, _
                                            Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                            Val(.TextMatrix(lngCurRow, COL_频率间隔)), _
                                            .TextMatrix(lngCurRow, COL_间隔单位), _
                                            .TextMatrix(lngCurRow, COL_执行时间), _
                                            Val(.TextMatrix(lngCurRow, COL_剂量系数)), _
                                            Val(.TextMatrix(lngCurRow, COL_住院包装)), _
                                            Val(.TextMatrix(lngCurRow, COL_可否分零))), 5)
                                End If
                            End If
                        Else
                            .TextMatrix(lngCurRow, COL_总量) = Calc缺省药品总量(1, sng天数, _
                                    Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                    Val(.TextMatrix(lngCurRow, COL_频率间隔)), _
                                    .TextMatrix(lngCurRow, COL_间隔单位))
                        End If
                    End If
                End If
                
                If mbln天数 And InStr(",5,6,", rsItems!类别) > 0 And Val(.TextMatrix(lngCurRow, COL_频率性质)) <> 1 Then
                    .TextMatrix(lngCurRow, COL_天数) = sng天数
                End If
            ElseIf bln配方 Then
                '中药煎法,用法的总量与组成药相同(为了显示)
                .TextMatrix(lngCurRow, COL_总量) = .TextMatrix(lngCurRow - 1, COL_总量)
                .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow - 1, COL_总量单位)
            ElseIf Nvl(rsItems!期效, 0) = 1 Then
                '其它临嘱都需要总量
                '如果为一次性或计次临嘱缺省总量为1
                If Not IsNull(rsItems!总给予量) Then
                    vsAdvice.TextMatrix(lngCurRow, COL_总量) = rsItems!总给予量
                ElseIf int频率性质 = 1 Or Nvl(rsItems!计算方式, 0) = 3 Then
                    vsAdvice.TextMatrix(lngCurRow, COL_总量) = 1
                End If
                .TextMatrix(lngCurRow, COL_总量单位) = Nvl(rsItems!计算单位)
            End If
                    
            '执行时间,终止时间(总量,频率,执行时间之后)
            If .TextMatrix(lngCurRow, COL_频率) <> "" Then
                If Val(.TextMatrix(lngCurRow, COL_执行性质)) <> 1 Then
                    '可能求出缺省执行时间方案
                    If bln给药途径 Or bln中药用法 Then
                        If Not IsNull(rsItems!时间方案) Then
                            If ExeTimeValid(rsItems!时间方案, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位)) Then
                                .TextMatrix(lngCurRow, COL_执行时间) = rsItems!时间方案
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_执行时间) = "" Then
                            .TextMatrix(lngCurRow, COL_执行时间) = Get缺省时间(Val(str适用范围), .TextMatrix(lngCurRow, COL_频率), rsItems!诊疗项目ID)
                        End If
                    ElseIf int频率性质 = 0 Then
                        If Not IsNull(rsItems!时间方案) Then
                            If ExeTimeValid(rsItems!时间方案, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位)) Then
                                .TextMatrix(lngCurRow, COL_执行时间) = rsItems!时间方案
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_执行时间) = "" Then
                            .TextMatrix(lngCurRow, COL_执行时间) = Get缺省时间(Val(str适用范围), .TextMatrix(lngCurRow, COL_频率))
                        End If
                    End If
                End If
                
                If bln采集方法 Then
                    .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
                ElseIf bln给药途径 Or bln中药用法 Then
                    '成药和中药配方的用法,执行时间,配方终止时间
                    If bln中药用法 Then
                        .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
                        '终止时间:中药配方(长嘱)可以计算,之所以要在用法行算,是因为这里才可能得到缺省执行时间
                        If Nvl(rsItems!期效, 0) = 0 Then
                            .Cell(flexcpData, lngCurRow, COL_终止时间) = Calc终止时间(lngCurRow)
                            .TextMatrix(lngCurRow, COL_终止时间) = Format(.Cell(flexcpData, lngCurRow, COL_终止时间), "MM-dd HH:mm")
                        End If
                    End If
                    For j = lngCurRow - 1 To lngRow Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngCurRow) Then
                            If bln给药途径 Then
                                .TextMatrix(j, COL_用法) = rsItems!名称
                            Else
                                .TextMatrix(j, COL_终止时间) = .TextMatrix(lngCurRow, COL_终止时间)
                                .Cell(flexcpData, j, COL_终止时间) = .Cell(flexcpData, lngCurRow, COL_终止时间)
                            End If
                            .TextMatrix(j, COL_执行时间) = .TextMatrix(lngCurRow, COL_执行时间)
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
                                
            '开嘱医生和开嘱科室
            If mbln护士站 Then
                '与上一行相同
                If lngPreRow <> -1 Then
                    .TextMatrix(lngCurRow, COL_开嘱医生) = .TextMatrix(lngPreRow, COL_开嘱医生)
                End If
                '缺省为病人的住院医师或病人科室的第一个医生
                If .TextMatrix(lngCurRow, COL_开嘱医生) = "" Then
                    str医生 = Get住院医师
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str医生, lng医生ID)
                    .TextMatrix(lngCurRow, COL_开嘱医生) = str医生
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngCurRow, COL_开嘱医生), lng医生ID, , , True)
                End If
                .TextMatrix(lngCurRow, COL_开嘱科室ID) = Get开嘱科室ID(lng医生ID, mlng病人科室id)
            Else
                .TextMatrix(lngCurRow, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(lngCurRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            End If
                                
            '执行性质
            If InStr(",5,6,7,", rsItems!类别) > 0 Then
                If Nvl(rsItems!执行性质, 0) = 5 Then
                    .TextMatrix(lngCurRow, COL_执行性质) = 5
                Else
                    .TextMatrix(lngCurRow, COL_执行性质) = 4
                End If
            ElseIf bln给药途径 Or bln中药煎法 Or bln中药用法 Or bln采集方法 Then
                .TextMatrix(lngCurRow, COL_执行性质) = Nvl(rsItems!执行性质, 0)
            Else
                .TextMatrix(lngCurRow, COL_执行性质) = Nvl(rsItems!执行性质, 0)
            End If
            
            '执行科室ID:为0-叮嘱,5-院外执行时取出为0
            If rsItems!类别 = "Z" And Nvl(rsItems!操作类型, 0) = 3 Then
                If Nvl(rsItems!执行科室ID, 0) <> 0 Then
                    .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                Else
                    '转科医嘱取缺省转入临床科室(不管执行性质)
                    Call Get临床科室(2, mlng病人科室id, lngTmp, , Not gbln病区科室独立)
                    .TextMatrix(lngCurRow, COL_执行科室ID) = lngTmp
                End If
            ElseIf rsItems!类别 = "Z" And Nvl(rsItems!操作类型, 0) = 7 Then
                If Nvl(rsItems!执行科室ID, 0) <> 0 Then
                    .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                Else
                    '会诊医嘱取缺省会诊临床科室(不管执行性质)
                    Call Get临床科室(2, mlng病人科室id, lngTmp)
                    .TextMatrix(lngCurRow, COL_执行科室ID) = lngTmp
                End If
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_执行性质))) = 0 Then
                If Nvl(rsItems!执行科室ID, 0) <> 0 Then
                    If InStr(",5,6,7,", rsItems!类别) > 0 Then
                        str药房IDs = Get可用药房IDs(rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), mlng病人科室id)
                        If InStr("," & str药房IDs & ",", "," & rsItems!执行科室ID & ",") > 0 Then
                            .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                        End If
                    ElseIf Val(.TextMatrix(lngCurRow, COL_执行性质)) = 4 Then
                        '4-指定科室时才取,其它的固定生成
                        .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                    End If
                End If
                If Val(.TextMatrix(lngCurRow, COL_执行科室ID)) = 0 Then
                    '药品类的整个成套相同
                    If rsItems!类别 = "5" Then
                        If lng西药房ID = 0 Then
                            lng西药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, Nvl(rsItems!期效, 0), , True)
                        End If
                        .TextMatrix(lngCurRow, COL_执行科室ID) = lng西药房ID
                    ElseIf rsItems!类别 = "6" Then
                        If lng成药房ID = 0 Then
                            lng成药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, Nvl(rsItems!期效, 0), , True)
                        End If
                        .TextMatrix(lngCurRow, COL_执行科室ID) = lng成药房ID
                    ElseIf rsItems!类别 = "7" Then
                        If lng中药房ID = 0 Then
                            lng中药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, Nvl(rsItems!期效, 0), , True)
                        End If
                        .TextMatrix(lngCurRow, COL_执行科室ID) = lng中药房ID
                    Else
                        '之前先求出开嘱科室
                        .TextMatrix(lngCurRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, _
                            rsItems!诊疗项目ID, 0, Val(.TextMatrix(lngCurRow, COL_执行性质)), mlng病人科室id, _
                            Val(.TextMatrix(lngCurRow, COL_开嘱科室ID)), Nvl(rsItems!期效, 0))
                    End If
                End If
            End If
                        
            '医生嘱托
            .TextMatrix(lngCurRow, COL_医生嘱托) = Nvl(rsItems!医生嘱托)
            
            '开嘱时间
            If mbln补录 And IsDate(.Cell(flexcpData, lngCurRow, COL_开始时间)) Then
                .TextMatrix(lngCurRow, COL_开嘱时间) = .TextMatrix(lngCurRow, COL_开始时间)
                .Cell(flexcpData, lngCurRow, COL_开嘱时间) = .Cell(flexcpData, lngCurRow, COL_开始时间)
            Else
                .TextMatrix(lngCurRow, COL_开嘱时间) = Format(vCurDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngCurRow, COL_开嘱时间) = Format(vCurDate, "yyyy-MM-dd HH:mm")
            End If
            
            '紧急标志
            .TextMatrix(lngCurRow, COL_标志) = IIF(mbln补录, 2, chk紧急.Value) '可以在界面先统一设置为紧急
            blnFirst = True
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_类别)) > 0 Then
                If Val(.TextMatrix(lngCurRow, COL_相关ID)) = Val(.TextMatrix(lngCurRow - 1, COL_相关ID)) Then
                    blnFirst = False
                End If
            End If
            If blnFirst Then
                If Val(.TextMatrix(lngCurRow, COL_标志)) = 2 Then
                    Set .Cell(flexcpPicture, lngCurRow, COL_F标志) = imgFlag.ListImages("补录").Picture
                    .Cell(flexcpPictureAlignment, lngCurRow, COL_F标志) = 4
                ElseIf Val(.TextMatrix(lngCurRow, COL_标志)) = 1 Then
                    Set .Cell(flexcpPicture, lngCurRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
                    .Cell(flexcpPictureAlignment, lngCurRow, COL_F标志) = 4
                End If
            End If
            
            '读取药品库存
            If InStr(",5,6,7,", .TextMatrix(lngCurRow, COL_类别)) > 0 Then
                If Val(.TextMatrix(lngCurRow, COL_收费细目ID)) <> 0 And Val(.TextMatrix(lngCurRow, COL_执行科室ID)) <> 0 Then
                    .TextMatrix(lngCurRow, COL_库存) = GetStock(Val(.TextMatrix(lngCurRow, COL_收费细目ID)), Val(.TextMatrix(lngCurRow, COL_执行科室ID)))
                End If
            End If
            
            '----------------------
            '毒麻精药品标识:中药配方及组成味中药不处理
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_类别)) > 0 And .TextMatrix(lngCurRow, COL_毒理分类) <> "" Then
                If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(lngCurRow, COL_毒理分类)) > 0 Then
                    .Cell(flexcpFontBold, lngCurRow, COL_医嘱内容) = True
                End If
            End If
            
            '隐蔽一些附加行
            If (InStr(",F,G,D,7,E,C,", rsItems!类别) > 0 And Not IsNull(rsItems!相关序号)) Or bln给药途径 Then
                .RowHidden(lngCurRow) = True
            End If
            
            '医嘱内容
            If Not .RowHidden(lngCurRow) Then
                If InStr(",F,D,", rsItems!类别) > 0 And IsNull(rsItems!相关序号) Then
                    .TextMatrix(lngCurRow, COL_医嘱内容) = rsItems!名称 '临时
                Else
                    .TextMatrix(lngCurRow, COL_医嘱内容) = AdviceTextMake(lngCurRow)
                End If
            Else
                .TextMatrix(lngCurRow, COL_医嘱内容) = rsItems!名称
            End If
            
            If lngPreRow = -1 And Not .RowHidden(lngCurRow) Then lngPreRow = lngCurRow
            
            '----------------------
            intCount = intCount + 1
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        '再取检查和手术的医嘱内容
        For i = lngRow To lngCurRow
            If InStr(",F,D,", .TextMatrix(i, COL_类别)) > 0 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                .TextMatrix(i, COL_医嘱内容) = AdviceTextMake(i)
            End If
        Next
        
        '调整受影响行的序号
        Call AdviceSet医嘱序号(lngCurRow + 1, intCount)
        
        '产生真实的医嘱ID
        For i = lngRow To lngCurRow
            lng相关ID = .RowData(i)
            .RowData(i) = zlDatabase.GetNextId("病人医嘱记录")
            For j = i - 1 To lngRow Step -1
                If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            For j = i + 1 To lngCurRow
                If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
        Next
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then '寻找可见行(如配方和检验之后)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        
        .Row = lngRow: .Col = COL_医嘱内容
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function AdviceSet中药配方(lng诊疗项目ID As Long, ByVal lngRow As Long, ByVal lng用法ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'功能：(重新)处理中药配方的缺省医嘱数据
'参数：lng诊疗项目ID=输入的中药配方ID或单味中药ID
'      lngRow=当前输入行
'      lng用法ID=缺省中药用法ID
'      strExtData=包含配方组成味药及煎法数据
'      rsCurr=如果是修改了配方内容后调用,则包含要保持的一些当前值
'返回：处理后的中药配方的当前显示行号
    Dim rsItems As New ADODB.Recordset '中药详细信息
    Dim rsUse As New ADODB.Recordset '中药用法信息
    Dim rs煎法 As New ADODB.Recordset '中药煎法项目信息
    Dim rs用法 As New ADODB.Recordset '中药用法项目信息
    Dim arr中药s As Variant, str中药IDs As String, lng相关ID As Long
    Dim lngCopyRow As Long '缺省参照行
    Dim lngDrugRow As Long '如果缺省参照行是中药配方,则为该配方的第一个中药行
    Dim lngFirstRow As Long '当前配方的第一个中药行
    Dim strSQL As String, i As Long
    
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim lng煎法ID As Long, int疗程 As Integer
    Dim str医生 As String, lng医生ID As Long
        
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn配方行(lngCopyRow) Then
            '如果上一有效行是中药配方的,则取它的第一中药行
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_相关ID)
        End If
    End If
    
    '获取相关数据库信息
    '------------------
    arr中药s = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr中药s)
        str中药IDs = str中药IDs & "," & CStr(Split(arr中药s(i), ",")(0))
    Next
    str中药IDs = Mid(str中药IDs, 2)
    lng煎法ID = Val(Split(strExtData, "|")(1))
    
    '配方用法信息:直接输入配方时才有可能有,输入单味中药无
    strSQL = "Select A.用法ID,A.频次,A.疗程,A.医生嘱托" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And B.服务对象 IN(2,3)" & _
        " And Nvl(A.性质,0)=0 And A.项目ID=[1]"
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊疗项目ID)
    If Not rsUse.EOF Then lng用法ID = rsUse!用法ID '缺省设置的中药配方用法优先
    
    '配方组成味中药信息:中药无规格概念,对应的的规格记录一定有且只有一条
    strSQL = "Select A.*,B.药品ID,B.剂量系数,B.住院包装,B.住院单位,B.可否分零,C.处方职务" & _
        " From 诊疗项目目录 A,药品规格 B,药品特性 C" & _
        " Where A.ID=B.药名ID And A.ID=C.药名ID And A.ID IN(" & str中药IDs & ")"
    zlDatabase.OpenRecordset rsItems, strSQL, Me.Caption 'In
    
    '配方煎法项目信息
    strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
    Set rs煎法 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng煎法ID)
    
    '配方用法项目信息
    strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
    Set rs用法 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng用法ID)
    
    '加入配方组成味中药行:按照用户输入顺序
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '中药用法的医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了配方中的内容,用法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的中药配方
        lng相关ID = zlDatabase.GetNextId("病人医嘱记录")
    End If
    
    For i = 0 To UBound(arr中药s)
        rsItems.Filter = "ID=" & CStr(Split(arr中药s(i), ",")(0)) '应该肯定有
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = zlDatabase.GetNextId("病人医嘱记录")
        vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到后面的中药用法行
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '新增
        vsAdvice.TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
        vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
        vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        vsAdvice.TextMatrix(lngRow, COL_类别) = rsItems!类别
        vsAdvice.TextMatrix(lngRow, COL_医嘱内容) = rsItems!名称
        vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
        vsAdvice.TextMatrix(lngRow, COL_频率性质) = Nvl(rsItems!执行频率, 0)
        vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rsItems!操作类型)
        
        vsAdvice.TextMatrix(lngRow, COL_单量) = FormatEx(Val(Split(arr中药s(i), ",")(1)), 5) '单味药的单次用量
        vsAdvice.TextMatrix(lngRow, COL_单量单位) = Nvl(rsItems!计算单位)
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = CStr(Split(arr中药s(i), ",")(2)) '单味药的脚注
        
        '规格信息:中药不存在规格概念,一定有
        vsAdvice.TextMatrix(lngRow, COL_收费细目ID) = rsItems!药品ID
        vsAdvice.TextMatrix(lngRow, COL_剂量系数) = rsItems!剂量系数
        vsAdvice.TextMatrix(lngRow, COL_住院单位) = rsItems!住院单位
        vsAdvice.TextMatrix(lngRow, COL_住院包装) = rsItems!住院包装
        vsAdvice.TextMatrix(lngRow, COL_可否分零) = Nvl(rsItems!可否分零, 0) '对中药实际上无用
        vsAdvice.TextMatrix(lngRow, COL_处方职务) = Nvl(rsItems!处方职务)
        
        '计价性质:各自独立
        vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
        
        If lngFirstRow <> 0 Then
            '与上一行已设置的组成中药相同
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = vsAdvice.TextMatrix(lngFirstRow, COL_执行性质)
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID)
            vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
            vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
            
            vsAdvice.TextMatrix(lngRow, COL_终止时间) = vsAdvice.TextMatrix(lngFirstRow, COL_终止时间)
            vsAdvice.Cell(flexcpData, lngRow, COL_终止时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_终止时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
            vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
            
            vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
        ElseIf Not rsCurr Is Nothing Then
            '修改了配方内容后重新设置,保持与当前的值
            
            '执行性质:修改时根据当前界面设置决定
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "自备药", 5, 4)
            '执行科室
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Nvl(rsCurr!执行科室ID)
            
            vsAdvice.TextMatrix(lngRow, COL_频率) = Nvl(rsCurr!频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = Nvl(rsCurr!频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = Nvl(rsCurr!频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = Nvl(rsCurr!间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = Nvl(rsCurr!总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = Nvl(rsCurr!执行时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开始时间) = Format(Nvl(rsCurr!开始时间), "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = CStr(Nvl(rsCurr!开始时间))
            
            vsAdvice.TextMatrix(lngRow, COL_终止时间) = Format(Nvl(rsCurr!终止时间), "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_终止时间) = CStr(Nvl(rsCurr!终止时间))
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsCurr!开嘱医生)
            vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsCurr!开嘱科室ID)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(Nvl(rsCurr!开嘱时间), "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = CStr(Nvl(rsCurr!开嘱时间))
            
            vsAdvice.TextMatrix(lngRow, COL_标志) = Nvl(rsCurr!标志)
        Else
            '执行性质:中药配方组成中药相同,缺省=4-指定科室
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = 4
        
            '执行科室
            If lngDrugRow <> -1 Then '缺省与上一配方行相同
                vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID)
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!ID, rsItems!药品ID, Nvl(rsItems!执行科室, 0), mlng病人科室id, 0, cbo期效.ListIndex, , True)
            End If
            
            '执行频率
            '根据用法里面设置的优先
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!频次) Then
                    Call Get频率信息_编码(rsUse!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                End If
            End If
            '或缺省与上一行相同
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" And lngDrugRow <> -1 Then
                If Val(vsAdvice.TextMatrix(lngDrugRow, COL_EDIT)) = 1 And vsAdvice.TextMatrix(lngDrugRow, COL_频率) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率)
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngDrugRow, COL_频率次数)
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngDrugRow, COL_频率间隔)
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngDrugRow, COL_间隔单位)
                End If
            End If
            '或取缺省值
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" Then
                Call Get缺省频率(2, str频率, int频率次数, int频率间隔, str间隔单位)
                vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            End If
            
            '总量(付数):长嘱临嘱都需要,长嘱还要由此直接确定终止时间
            If vsAdvice.TextMatrix(lngRow, COL_频率) <> "" Then
                int疗程 = 1
                If Not rsUse.EOF Then int疗程 = Nvl(rsUse!疗程, 1)
                '配方付数
                vsAdvice.TextMatrix(lngRow, COL_总量) = Calc缺省药品总量(1, int疗程, _
                        Val(vsAdvice.TextMatrix(lngRow, COL_频率次数)), _
                        Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔)), _
                        vsAdvice.TextMatrix(lngRow, COL_间隔单位))
            End If
            
            '执行时间
            If lngDrugRow <> -1 Then '缺省与上一行相同
                If vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率) Then
                    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngDrugRow, COL_执行时间)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_执行时间) = "" Then '缺省时间方案
                vsAdvice.TextMatrix(lngRow, COL_执行时间) = Get缺省时间(2, vsAdvice.TextMatrix(lngRow, COL_频率), lng用法ID)
            End If
            
            '开始时间
            If IsDate(txt开始时间.Text) Then
                vsAdvice.TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
            End If
            
            '终止时间:长嘱时根据开始时间、付数、频次,执行时间自动确定
            If cbo期效.ListIndex = 0 Then
                vsAdvice.Cell(flexcpData, lngRow, COL_终止时间) = Calc终止时间(lngRow)
                vsAdvice.TextMatrix(lngRow, COL_终止时间) = Format(vsAdvice.Cell(flexcpData, lngRow, COL_终止时间), "MM-dd HH:mm")
            End If
            
            '开嘱医生和开嘱科室
            If mbln护士站 Then '护士使用时
                '与上一行相同
                If lngCopyRow <> -1 Then
                    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngCopyRow, COL_开嘱医生)
                End If
                '缺省为病人的住院医师或病人科室的第一个医生
                If vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = "" Then
                    str医生 = Get住院医师
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str医生, lng医生ID)
                    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = str医生
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, vsAdvice.TextMatrix(lngRow, COL_开嘱医生), lng医生ID, , , True)
                End If
                vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(lng医生ID, mlng病人科室id)
            Else
                vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            End If
            
            If mbln补录 And IsDate(vsAdvice.Cell(flexcpData, lngRow, COL_开始时间)) Then
                '补录配方时,缺省开嘱时间与开始时间相同
                vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngRow, COL_开始时间)
                vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngRow, COL_开始时间)
            Else
                vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
                vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            End If
            vsAdvice.TextMatrix(lngRow, COL_标志) = IIF(mbln补录, 2, chk紧急.Value)
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '该中药配方的第一个组成中药行
        lngRow = lngRow + 1 '保持当前输入行位置
    Next
    
    '加入中药配方煎法行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = zlDatabase.GetNextId("病人医嘱记录")
    vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID
    vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '新增
    vsAdvice.TextMatrix(lngRow, COL_期效) = vsAdvice.TextMatrix(lngFirstRow, COL_期效)
    vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs煎法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng煎法ID
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rs煎法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rs煎法!操作类型)
    
    '!中药煎法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    
    vsAdvice.TextMatrix(lngRow, COL_医嘱内容) = rs煎法!名称
    
    vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
    
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = vsAdvice.TextMatrix(lngFirstRow, COL_频率性质) '以药品的为准
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    vsAdvice.TextMatrix(lngRow, COL_终止时间) = vsAdvice.TextMatrix(lngFirstRow, COL_终止时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_终止时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_终止时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Nvl(rs煎法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "离院带药", 5, Nvl(rs煎法!执行科室, 0))
    End If
    
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rs煎法!类别, lng煎法ID, 0, _
            Nvl(rs煎法!执行科室, 0), mlng病人科室id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)), cbo期效.ListIndex)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rs煎法!计价性质, 0)
    vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
    
    vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
    
    vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
    
    '保持当前输入行位置
    lngRow = lngRow + 1
    
    '设置中药配方用法行:中药配方的显示行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng相关ID
    
    If Not rsCurr Is Nothing Then
        '修改了配方内容,标记为修改
        If InStr(",0,3,", rsCurr!Edit) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
        Else
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '本来就是新增或修改
        End If
    Else
        '新输入的中药配方,为新增
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_期效) = vsAdvice.TextMatrix(lngFirstRow, COL_期效)
    vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs用法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng用法ID
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rs用法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rs用法!操作类型)
    
    '!中药用法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    vsAdvice.TextMatrix(lngRow, COL_总量单位) = "付"
    
    vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
    
    vsAdvice.TextMatrix(lngRow, COL_名称) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_用法) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = vsAdvice.TextMatrix(lngFirstRow, COL_频率性质)
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    vsAdvice.TextMatrix(lngRow, COL_终止时间) = vsAdvice.TextMatrix(lngFirstRow, COL_终止时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_终止时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_终止时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Nvl(rs用法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "离院带药", 5, Nvl(rs用法!执行科室, 0))
    End If
    
    '中药用法如果未设置执行科室,则缺省为病人所在病区(门诊要改为病人所在科室!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rs用法!类别, lng用法ID, 0, _
            Nvl(rs用法!执行科室, 0), mlng病人科室id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)), cbo期效.ListIndex)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rs用法!计价性质, 0)
    vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
    
    vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
    
    vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
    If Val(vsAdvice.TextMatrix(lngRow, COL_标志)) = 2 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("补录").Picture
        vsAdvice.Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_标志)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
        vsAdvice.Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
    End If
    
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsCurr!医生嘱托)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsUse!医生嘱托)
    End If
    
    '中药配方的中药库存
    Call GetDrugStock(lngRow)
    
    '中药配方医嘱内容
    vsAdvice.TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
        
    AdviceSet中药配方 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet检验组合(ByVal lngRow As Long, ByVal lng采集方法ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'功能：处理新增的检验(组合)
'参数：rsItems=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng采集方法ID=缺省的采集方法
'      strExtData=检查:"项目ID1,项目ID2,...;检验标本"
'      rsCurr=修改检验项目时用
'返回：处理之后的当前显示行号
    Dim rsMore As New ADODB.Recordset '采集方法信息
    Dim rsItems As New ADODB.Recordset '检验项目信息
    Dim arrItems As Variant, strItems As String
    Dim strSQL As String, curDate As Date
    Dim str医生 As String, lng医生ID As Long
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng相关ID As Long, str医嘱内容 As String
    Dim lngCopyRow As Long, lngFirstRow As Long, i As Long
    
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    '当前时间
    curDate = zlDatabase.Currentdate
    
    '检验项目信息
    '----------------------------------------------------------------------------
    '各个检验项目信息:按输入顺序
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        strItems = strItems & "," & Val(arrItems(i))
    Next
    strSQL = "Select * From 诊疗项目目录 Where ID IN(" & Mid(strItems, 2) & ")"
    Call zlDatabase.OpenRecordset(rsItems, strSQL, Me.Caption) 'In
    
    '取某个检验项目的采集方法
    strSQL = "Select A.项目ID,Nvl(A.性质,0) as 序号,A.用法ID" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And B.服务对象 IN(2,3)" & _
        " And A.项目ID IN(" & Mid(strItems, 2) & ")" & _
        " Order by A.项目ID,Nvl(A.性质,0)"
    Call zlDatabase.OpenRecordset(rsMore, strSQL, Me.Caption) 'In
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng采集方法ID = 0 Then
            lng采集方法ID = rsMore!用法ID '修改时不变
        End If
    End If
    
    strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
    Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng采集方法ID)
    
    mblnRowChange = False
    
    '设置各行检验项目
    '----------------------------------------------------------------------------
    '采集方法医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了检验组合中的内容,采集方法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的中药配方
        lng相关ID = zlDatabase.GetNextId("病人医嘱记录")
    End If
    
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = zlDatabase.GetNextId("病人医嘱记录")
            .TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到采集方法行
            .TextMatrix(lngRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
            .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
            .TextMatrix(lngRow, COL_状态) = 1 '新开
            
            .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
            Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
            
            .TextMatrix(lngRow, COL_类别) = rsItems!类别
            .TextMatrix(lngRow, COL_医嘱内容) = rsItems!名称
            .TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
            .TextMatrix(lngRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
            If .TextMatrix(lngRow, COL_期效) = "临嘱" And Nvl(rsItems!执行频率, 0) = 0 And mbln一次性 Then
                .TextMatrix(lngRow, COL_频率性质) = 1 '住院可选择频率的临嘱缺省为一次性
            Else
                .TextMatrix(lngRow, COL_频率性质) = Nvl(rsItems!执行频率, 0)
            End If
            .TextMatrix(lngRow, COL_操作类型) = Nvl(rsItems!操作类型)
            .TextMatrix(lngRow, COL_处方限量) = Nvl(rsItems!录入限量)
            .TextMatrix(lngRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
            .TextMatrix(lngRow, COL_执行性质) = Nvl(rsItems!执行科室, 0)
            '检验标本
            .TextMatrix(lngRow, COL_标本部位) = Split(strExtData, ";")(1)
            
            '部份内容一并采集的检验项目相同
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
                
                '一并采集的检验项目应该相同
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngFirstRow, COL_执行科室ID)
                End If
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
            
                .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngFirstRow, COL_开始时间)
                .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngFirstRow, COL_开始时间)
                
                .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngFirstRow, COL_开嘱医生)
                .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngFirstRow, COL_开嘱科室ID)
                
                .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngFirstRow, COL_开嘱时间)
                .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngFirstRow, COL_开嘱时间)
                
                .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngFirstRow, COL_标志)
            ElseIf Not rsCurr Is Nothing Then
                If cbo期效.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_总量) = Nvl(rsCurr!总量, 1)
                End If
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    If Nvl(rsCurr!执行科室ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = rsCurr!执行科室ID
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!ID, 0, _
                            Nvl(rsItems!执行科室, 0), mlng病人科室id, Nvl(rsCurr!开嘱科室ID, 0), cbo期效.ListIndex)
                    End If
                End If
                
                '执行频率
                .TextMatrix(lngRow, COL_频率) = Nvl(rsCurr!频率)
                .TextMatrix(lngRow, COL_频率次数) = Nvl(rsCurr!频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = Nvl(rsCurr!频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = Nvl(rsCurr!间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = Nvl(rsCurr!执行时间)
                
                '时间/科室/医生
                .TextMatrix(lngRow, COL_开始时间) = Format(Nvl(rsCurr!开始时间), "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = CStr(Nvl(rsCurr!开始时间))
                
                .TextMatrix(lngRow, COL_开嘱时间) = Format(Nvl(rsCurr!开嘱时间), "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = CStr(Nvl(rsCurr!开嘱时间))
                
                .TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsCurr!开嘱医生)
                .TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsCurr!开嘱科室ID)
                
                .TextMatrix(lngRow, COL_标志) = Nvl(rsCurr!标志)
            Else
                '长嘱时无总量
                If cbo期效.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_总量) = 1
                End If
                
                '开嘱医生和科室
                If mbln护士站 Then '护士使用时
                    If lngCopyRow <> -1 Then '与上一行相同
                        .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngCopyRow, COL_开嘱医生)
                    End If
                    '缺省为病人的住院医师或病人科室的第一个医生
                    If .TextMatrix(lngRow, COL_开嘱医生) = "" Then
                        str医生 = Get住院医师
                        Call Get开嘱医生(mlng病人科室id, mbln护士站, str医生, lng医生ID)
                        .TextMatrix(lngRow, COL_开嘱医生) = str医生
                    Else
                        Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), lng医生ID, , , True)
                    End If
                    .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(lng医生ID, mlng病人科室id)
                Else
                    .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                    .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
                End If
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    '之前要求出开嘱科室ID
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsItems!类别, rsItems!ID, 0, _
                        Nvl(rsItems!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), cbo期效.ListIndex)
                End If
                
                '执行频率
                Call Get缺省频率(Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngRow, COL_频率) = str频率
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                
                '执行时间:"可选频率"(药品是可选频率,但可能设置为一次性)
                If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                    If lngCopyRow <> -1 Then '与上一行相同
                        If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                            .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_执行时间) = "" Then  '缺省时间方案
                        .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率))
                    End If
                End If
            
                '开始时间
                If IsDate(txt开始时间.Text) Then
                    .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
                End If
                
                '开嘱时间
                If mbln补录 And IsDate(.Cell(flexcpData, lngRow, COL_开始时间)) Then
                    '补录医嘱时,项目的开嘱时间缺省为开始时间
                    .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开始时间)
                    .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                Else
                    .TextMatrix(lngRow, COL_开嘱时间) = Format(curDate, "MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                End If
                
                '紧急标志
                .TextMatrix(lngRow, COL_标志) = IIF(mbln补录, 2, chk紧急.Value)
            End If
            
            str医嘱内容 = str医嘱内容 & "," & rsItems!名称 '医嘱内容
            If lngFirstRow = 0 Then lngFirstRow = lngRow '第一项目行
            lngRow = lngRow + 1 '保持当前输入行位置
            
            rsItems.MoveNext
        Next
        
        '设置标本的采集方法
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng相关ID
        
        If Not rsCurr Is Nothing Then
            '修改了检验组合内容,标记为修改
            If InStr(",0,3,", rsCurr!Edit) > 0 Then
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
            Else
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '本来就是新增或修改
            End If
        Else
            '新输入的检验组合,为新增
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
        End If
        
        .TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
        .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngRow, COL_状态) = 1 '新开
        
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        .TextMatrix(lngRow, COL_类别) = rsMore!类别
        .TextMatrix(lngRow, COL_名称) = rsMore!名称
        .TextMatrix(lngRow, COL_用法) = rsMore!名称
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsMore!ID
        .TextMatrix(lngRow, COL_计算方式) = Nvl(rsMore!计算方式, 0)
        .TextMatrix(lngRow, COL_操作类型) = Nvl(rsMore!操作类型)
        .TextMatrix(lngRow, COL_计价性质) = Nvl(rsMore!计价性质, 0)
        .TextMatrix(lngRow, COL_标本部位) = .TextMatrix(lngFirstRow, COL_标本部位)
        
        '总量为检验项目的,与检验项目相同
        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
        If cbo期效.ListIndex = 1 Then
            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsMore!计算单位)
        End If
        
        '执行频率
        .TextMatrix(lngRow, COL_频率性质) = .TextMatrix(lngFirstRow, COL_频率性质) '以检验的为准
        .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
        .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
        .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
        .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
        .TextMatrix(lngRow, COL_执行性质) = Nvl(rsMore!执行科室, 0)
        
        '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
            .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsMore!类别, rsMore!ID, 0, _
                Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngFirstRow, COL_开嘱科室ID)), cbo期效.ListIndex)
        End If
        
        '时间/科室/医生
        .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngFirstRow, COL_开始时间)
        .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngFirstRow, COL_开始时间)
        .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngFirstRow, COL_开嘱时间)
        .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngFirstRow, COL_开嘱时间)
        .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngFirstRow, COL_开嘱科室ID)
        .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngFirstRow, COL_开嘱医生)
        
        '显示紧急标志
        .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngFirstRow, COL_标志)
        If Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
            Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("补录").Picture
            .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
        ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
            Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
            .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
        End If
                
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsCurr!医生嘱托)
        End If
        
        '医嘱内容:检验1,检验2(标本 采集方法)
        .TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet检验组合 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet诊疗项目(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng给药途径ID As Long, ByVal lngGroupRow As Long, ByVal strExtData As String)
'功能：处理新增(插入)的中、西成药，检查(组合)，手术(组合)，及其它诊疗项目的缺省医嘱数据
'参数：rsInput=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng给药途径ID=缺省给药途径ID,或一并给药时的给药途径ID
'      lngGroupRow=在一并给药的一组成药中插入新的成药行时,对应一并给药的一行行号
'      strExtData=检查:包含检查部位信息,手术:包含附加手术及麻醉的信息,可能无附加手术
    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '诊疗项目详细信息
    Dim strSQL As String, lngCopyRow As Long
    Dim blnFirst As Boolean, lngTmp As Long, i As Long
    Dim str医生 As String, lng医生ID As Long
    Dim str药房IDs As String, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng药品ID As Long, bln品种 As Boolean
        
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
    With vsAdvice
        '开始设置医嘱缺省内容
        .RowData(lngRow) = zlDatabase.GetNextId("病人医嘱记录")
        .TextMatrix(lngRow, COL_EDIT) = 1 '新增
        .TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
        .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngRow, COL_状态) = 1 '新开
        
        '序号:保持连续,当前行占用新序号后,后面的序号向后移
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_类别) = rsInput!类别ID
        .TextMatrix(lngRow, COL_名称) = rsInput!名称 '该名称可能是别名
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsInput!诊疗项目ID
        
        '药品特性
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            strSQL = "Select 毒理分类,药品剂型,处方限量,处方职务,品种医嘱 From 药品特性 Where 药名ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!诊疗项目ID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(lngRow, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                .TextMatrix(lngRow, COL_处方限量) = Nvl(rsTmp!处方限量)
                .TextMatrix(lngRow, COL_处方职务) = Nvl(rsTmp!处方职务)
                '是否长嘱药品固定按品种下达
                bln品种 = Nvl(rsTmp!品种医嘱, 0) <> 0 And cbo期效.ListIndex = 0
            End If
        End If
        
        '是否长嘱药品固定按品种下达
        lng药品ID = Nvl(rsInput!收费细目ID, 0)
        If bln品种 Then lng药品ID = 0
        
        '药品的规格信息
        .TextMatrix(lngRow, COL_收费细目ID) = lng药品ID
        If lng药品ID <> 0 Then
            strSQL = "Select Nvl(C.名称,A.名称) as 名称," & _
                " B.剂量系数,B.住院单位,B.住院包装,B.可否分零" & _
                " From 收费项目目录 A,药品规格 B,收费项目别名 C" & _
                " Where A.ID=B.药品ID And A.ID=[1]" & _
                " And A.ID=C.收费细目ID(+) And C.码类(+)=1 And C.性质(+)=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药品ID, IIF(gbln商品名, 3, 1))
            .TextMatrix(lngRow, COL_名称) = rsTmp!名称 '将别名换成正式规格名称
            .TextMatrix(lngRow, COL_剂量系数) = rsTmp!剂量系数
            .TextMatrix(lngRow, COL_住院单位) = rsTmp!住院单位
            .TextMatrix(lngRow, COL_住院包装) = rsTmp!住院包装
            .TextMatrix(lngRow, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
        End If
        
        '获取更多诊疗项目信息
        '----------------------------------------------------------------------------
        strSQL = "Select A.*" & _
            " From 诊疗用法用量 A,诊疗项目目录 B" & _
            " Where A.用法ID=B.ID And (Nvl(A.性质,0)=0 Or B.服务对象 IN(2,3))" & _
            " And A.项目ID=[1]"
        strSQL = "Select A.*,Nvl(B.性质,0) as 性质,B.用法ID," & _
            " B.频次,B.成人剂量,B.小儿剂量,B.医生嘱托,B.疗程" & _
            " From 诊疗项目目录 A,(" & strSQL & ") B" & _
            " Where A.ID=B.项目ID(+) And A.ID=[1]" & _
            " Order by 性质"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!诊疗项目ID))
        
        If lng药品ID = 0 Then '将别名换成正式诊疗名称
            .TextMatrix(lngRow, COL_名称) = rsMore!名称
        End If
        
        If cbo期效.ListIndex = 0 Then
            If InStr(",5,6,", rsInput!类别ID) > 0 Or InStr(",1,2,", Nvl(rsMore!计算方式, 0)) > 0 Then
                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsMore!计算单位) '药品为剂量单位
            End If
        Else
            If InStr(",5,6,", rsInput!类别ID) > 0 Or (Nvl(rsMore!执行频率, 0) = 0 And InStr(",1,2,", Nvl(rsMore!计算方式, 0)) > 0) Then
                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsMore!计算单位) '药品为剂量单位
            End If
        End If
        
        If cbo期效.ListIndex = 1 Then
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                '中、西成药临嘱的总量单位就是住院单位
                .TextMatrix(lngRow, COL_总量单位) = .TextMatrix(lngRow, COL_住院单位)
            Else
                '其它临嘱要输入总量
                '如果为一次性或计次临嘱缺省总量为1
                If Nvl(rsMore!执行频率, 0) = 1 Or Nvl(rsMore!计算方式, 0) = 3 Then
                    .TextMatrix(lngRow, COL_总量) = 1
                End If
                .TextMatrix(lngRow, COL_总量单位) = Nvl(rsMore!计算单位)
            End If
        End If
        
        .TextMatrix(lngRow, COL_计算方式) = Nvl(rsMore!计算方式, 0)
        If .TextMatrix(lngRow, COL_期效) = "临嘱" And Nvl(rsMore!执行频率, 0) = 0 And mbln一次性 Then
            .TextMatrix(lngRow, COL_频率性质) = 1 '住院可选择频率的临嘱缺省为一次性
        Else
            .TextMatrix(lngRow, COL_频率性质) = Nvl(rsMore!执行频率, 0)
        End If
        .TextMatrix(lngRow, COL_操作类型) = Nvl(rsMore!操作类型)
        If InStr(",5,6,7,", rsInput!类别ID) = 0 Then
            .TextMatrix(lngRow, COL_处方限量) = Nvl(rsMore!录入限量)
        End If
        
        '标本部位
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            .TextMatrix(lngRow, COL_标本部位) = rsInput!名称 '记录药品输入时选择名称
        Else
            .TextMatrix(lngRow, COL_标本部位) = Nvl(rsMore!标本部位)
        End If
        
        '计价性质
        .TextMatrix(lngRow, COL_计价性质) = Nvl(rsMore!计价性质, 0)
    
        '执行性质:新增项目时根据项目设置,药品=4-指定科室,一并给药的相同
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_执行性质) = .TextMatrix(lngGroupRow, COL_执行性质)
            Else
                .TextMatrix(lngRow, COL_执行性质) = 4
            End If
        Else
            .TextMatrix(lngRow, COL_执行性质) = Nvl(rsMore!执行科室, 0)
        End If
        
        '开嘱医生和科室
        If lngGroupRow = 0 Then
            If mbln护士站 Then '护士使用时
                '与上一行相同
                If lngCopyRow <> -1 Then
                    .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngCopyRow, COL_开嘱医生)
                End If
                '缺省为病人的住院医师或病人科室的第一个医生
                If .TextMatrix(lngRow, COL_开嘱医生) = "" Then
                    str医生 = Get住院医师
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str医生, lng医生ID)
                    .TextMatrix(lngRow, COL_开嘱医生) = str医生
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), lng医生ID, , , True)
                End If
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(lng医生ID, mlng病人科室id)
            Else
                .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            End If
        Else
            .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngGroupRow, COL_开嘱医生)
            .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngGroupRow, COL_开嘱科室ID)
        End If
        
        '执行科室:药品缺省与上一行相同,一并给药的相同
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngGroupRow, COL_执行科室ID)
            ElseIf lngCopyRow <> -1 Then
                If rsInput!类别ID = .TextMatrix(lngCopyRow, COL_类别) Then
                    str药房IDs = Get可用药房IDs(rsInput!类别ID, rsInput!诊疗项目ID, lng药品ID, mlng病人科室id)
                    If InStr("," & str药房IDs & ",", "," & .TextMatrix(lngCopyRow, COL_执行科室ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngCopyRow, COL_执行科室ID)
                    End If
                End If
            End If
        End If
        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
            If rsInput!类别ID = "Z" And Nvl(rsMore!操作类型, 0) = 3 Then
                '转科医嘱取缺省转入临床科室(不管执行性质)
                Call Get临床科室(2, mlng病人科室id, lngTmp, , Not gbln病区科室独立)
                .TextMatrix(lngRow, COL_执行科室ID) = lngTmp
            ElseIf rsInput!类别ID = "Z" And Nvl(rsMore!操作类型, 0) = 7 Then
                '会诊医嘱取缺省会诊临床科室(不管执行性质)
                Call Get临床科室(2, mlng病人科室id, lngTmp)
                .TextMatrix(lngRow, COL_执行科室ID) = lngTmp
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                '执行性质为(0-叮嘱,5-院外执行)无执行科室
                '先要求出开嘱科室ID
                .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsInput!类别ID, rsInput!诊疗项目ID, _
                    lng药品ID, Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), cbo期效.ListIndex, , InStr(",5,6,", rsInput!类别ID) > 0)
            End If
        End If
        
        '药品库存
        If InStr(",5,6,", rsInput!类别ID) > 0 And lng药品ID <> 0 Then
            Call GetDrugStock(lngRow)
        End If
        
        '执行频率:可选频率,一次性或持续性
        If True Then 'If Nvl(rsMore!执行频率, 0) = 0 Then
            '缺省与上一新增行相同
            If lngCopyRow <> -1 Then
                If .TextMatrix(lngRow, COL_期效) = .TextMatrix(lngCopyRow, COL_期效) And Get频率范围(lngRow) = Get频率范围(lngCopyRow) Then
                    If Val(.TextMatrix(lngCopyRow, COL_EDIT)) = 1 And .TextMatrix(lngCopyRow, COL_频率) <> "" _
                        And Not (.TextMatrix(lngRow, COL_类别) = "7" And Not RowIn配方行(lngCopyRow)) _
                        And Not (.TextMatrix(lngRow, COL_类别) <> "7" And RowIn配方行(lngCopyRow)) Then
                        .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率)
                        .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngCopyRow, COL_频率次数)
                        .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngCopyRow, COL_频率间隔)
                        .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngCopyRow, COL_间隔单位)
                    End If
                End If
            End If
            '或取缺省频率
            If .TextMatrix(lngRow, COL_频率) = "" Then
                Call Get缺省频率(Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngRow, COL_频率) = str频率
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            End If
        End If
        
        '中，西成药的一些缺省信息
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '执行频率
            If lngGroupRow <> 0 Then
                '一并给药的相同
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngGroupRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngGroupRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngGroupRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngGroupRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngGroupRow, COL_执行时间)
                '频率性质也要相同,如已强制设置为一次性
                .TextMatrix(lngRow, COL_频率性质) = .TextMatrix(lngGroupRow, COL_频率性质)
            End If
            
            '确定临嘱用药天数：
            '1.最少为一个频率周期天数
            '2-有疗程则为疗程天数(应大于一个频率周期天数)
            If cbo期效.ListIndex = 1 Then
                sng天数 = msng天数
                If mbln天数 Then
                    If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                        If 7 > sng天数 Then sng天数 = 7
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                        If Val(.TextMatrix(lngRow, COL_频率间隔)) > sng天数 Then
                            sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔))
                        End If
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                        If Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24 > sng天数 Then
                            sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24
                        End If
                    End If
                    If sng天数 = 0 Then sng天数 = 1
                End If
            End If
            
            rsMore.Filter = "性质>0" '取第一种给药途径用为缺省设置
            If Not rsMore.EOF Then
                '不是一并给药时,设置的缺省用法频率优先
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!用法ID) Then lng给药途径ID = rsMore!用法ID
                    If Not IsNull(rsMore!频次) And Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then '缺省为一次性优先
                        Call Get频率信息_编码(rsMore!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                        .TextMatrix(lngRow, COL_频率) = str频率
                        .TextMatrix(lngRow, COL_频率次数) = int频率次数
                        .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                        .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                    End If
                End If
                
                '医生嘱托
                .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsMore!医生嘱托) '一般为给药途径的说明
                
                '药品单量
                If mint年龄 > 12 Then
                    If Nvl(rsMore!成人剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!成人剂量, 5)
                    End If
                Else
                    If Nvl(rsMore!小儿剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!小儿剂量, 5)
                    ElseIf Nvl(rsMore!成人剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!成人剂量 * (mint年龄 + 2) * 5 / 100, 5)
                    End If
                End If
                If Val(.TextMatrix(lngRow, COL_单量)) = 0 Then .TextMatrix(lngRow, COL_单量) = ""
                
                '药品临嘱总量:住院包装
                If cbo期效.ListIndex = 1 Then
                    If Nvl(rsMore!疗程, 1) > sng天数 Then sng天数 = Nvl(rsMore!疗程, 1)
                    If .TextMatrix(lngRow, COL_频率) <> "" And Val(.TextMatrix(lngRow, COL_单量)) <> 0 _
                        And Val(.TextMatrix(lngRow, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngRow, COL_住院包装)) <> 0 Then
                        If Val(.TextMatrix(lngRow, COL_频率性质)) = 1 Then '临嘱药品可能缺省为一次性
                            '仅按疗程算改为按最少用药天数算
                            .TextMatrix(lngRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(lngRow, COL_单量)), 1, 1, 1, "天", "", _
                                    Val(.TextMatrix(lngRow, COL_剂量系数)), _
                                    Val(.TextMatrix(lngRow, COL_住院包装)), _
                                    Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                        Else
                            '仅按疗程算改为按最少用药天数算
                            .TextMatrix(lngRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(lngRow, COL_单量)), sng天数, _
                                    Val(.TextMatrix(lngRow, COL_频率次数)), _
                                    Val(.TextMatrix(lngRow, COL_频率间隔)), _
                                    .TextMatrix(lngRow, COL_间隔单位), _
                                    .TextMatrix(lngRow, COL_执行时间), _
                                    Val(.TextMatrix(lngRow, COL_剂量系数)), _
                                    Val(.TextMatrix(lngRow, COL_住院包装)), _
                                    Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                        End If
                    End If
                End If
            End If
            
            '记录缺省天数
            If mbln天数 And cbo期效.ListIndex = 1 And Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then
                .TextMatrix(lngRow, COL_天数) = sng天数
            End If
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        '执行时间:"可选频率"(药品是可选频率,但可能设置为一次性)
        If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
            If .TextMatrix(lngRow, COL_执行时间) = "" Then
                If lngCopyRow <> -1 Then '与上一行相同
                    If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                    End If
                End If
                If .TextMatrix(lngRow, COL_执行时间) = "" Then  '缺省时间方案
                    .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率), lng给药途径ID)
                End If
            End If
        End If
        
        '其它(与项目无关)
        '---------------------------------------------------------------------
        If lngGroupRow = 0 Then
            If IsDate(txt开始时间.Text) Then
                .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
            End If
            
            If mbln补录 And IsDate(.Cell(flexcpData, lngRow, COL_开始时间)) Then
                '补录医嘱时,项目的开嘱时间缺省为开始时间
                .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开始时间)
                .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开始时间)
            Else
                .TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            End If
            .TextMatrix(lngRow, COL_标志) = IIF(mbln补录, 2, chk紧急.Value)
        Else
            .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngGroupRow, COL_开始时间)
            .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngGroupRow, COL_开始时间)
            
            .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngGroupRow, COL_开嘱时间)
            .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngGroupRow, COL_开嘱时间)
            
            .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngGroupRow, COL_标志)
        End If
        
        '紧急标志
        blnFirst = True
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    blnFirst = False
                End If
            End If
        End If
        If blnFirst Then
            If Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("补录").Picture
                .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
            ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
                Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
                .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
            End If
        End If
        
        
        '在主行处理完成之后处理附加行,并组合医嘱内容
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '新增一个给药途径项目,并设置相关
            If lng给药途径ID <> 0 Then
                .TextMatrix(lngRow, COL_用法) = Get项目名称(lng给药途径ID)
            End If
            If lngGroupRow <> 0 Then
                '一并给药的关联相同的给药途径行
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_相关ID) = .TextMatrix(lngGroupRow, COL_相关ID)
                Else
                    '这种情况是仅为了使用一并给药的相同设置
                    .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
                End If
            Else '独立新增的成药关联独立的给药途径行
                .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
            End If
            
            '毒麻精的颜色标识
            If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(lngRow, COL_毒理分类)) > 0 _
                And .TextMatrix(lngRow, COL_毒理分类) <> "" Then
                .Cell(flexcpFontBold, lngRow, COL_医嘱内容) = True
            End If
        ElseIf rsInput!类别ID = "D" And strExtData <> "" Then
            '检查的组合部位行
            Call AdviceSet检查手术(1, lngRow, strExtData)
        ElseIf rsInput!类别ID = "F" And strExtData <> "" Then
            '手术的附加手术及麻醉项目行
            Call AdviceSet检查手术(2, lngRow, strExtData)
        End If
        
        '医嘱内容
        .TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdviceInputFree(ByVal lngRow As Long)
'功能：处理新增自由输入医嘱
    Dim str医生 As String, lng医生ID As Long
    Dim lngCopyRow As Long
    
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) <> 0 Then
            If txt医嘱内容.Text <> .TextMatrix(lngRow, COL_医嘱内容) Then
                .TextMatrix(lngRow, COL_医嘱内容) = txt医嘱内容.Text
                
                '医生站时严格检查开嘱医生
                If Not mbln护士站 And GetAuditName(.TextMatrix(lngRow, COL_开嘱医生)) <> UserInfo.姓名 Then
                    .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                    Call GetCboIndex(cbo医生, UserInfo.姓名) '界面更新
                    .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
                End If
                
                If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                    .TextMatrix(lngRow, COL_EDIT) = 2
                    .TextMatrix(lngRow, COL_状态) = 1
                    Call ReSetColor(lngRow)
                End If
                mblnNoSave = True
            End If
        Else
            .RowData(lngRow) = zlDatabase.GetNextId("病人医嘱记录")
            .TextMatrix(lngRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngRow, COL_期效) = NeedName(cbo期效.Text)
            .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
            .TextMatrix(lngRow, COL_状态) = 1 '新开
            
            '序号:保持连续,当前行占用新序号后,后面的序号向后移
            .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
            Call AdviceSet医嘱序号(lngRow + 1, 1)
                            
            .TextMatrix(lngRow, COL_医嘱内容) = txt医嘱内容.Text
            .TextMatrix(lngRow, COL_类别) = "*" '特殊标记,为程序处理需要
            .TextMatrix(lngRow, COL_诊疗项目ID) = 0
            
            .TextMatrix(lngRow, COL_执行性质) = 4 '按可选执行科室处理，缺省为无
            .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "*", 0, 0, 4, mlng病人科室id, 0, cbo期效.ListIndex)
            If IsDate(txt开始时间.Text) Then
                .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
            End If
                
            '开嘱医生
            If mbln护士站 Then '护士使用时
                '与上一行相同
                If lngCopyRow <> -1 Then
                    .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngCopyRow, COL_开嘱医生)
                End If
                '缺省为病人的住院医师或病人科室的第一个医生
                If .TextMatrix(lngRow, COL_开嘱医生) = "" Then
                    str医生 = Get住院医师
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str医生, lng医生ID)
                    .TextMatrix(lngRow, COL_开嘱医生) = str医生
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, .TextMatrix(lngRow, COL_开嘱医生), lng医生ID, , , True)
                End If
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(lng医生ID, mlng病人科室id)
            Else
                .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            End If
                    
            '开嘱时间
            If mbln补录 And IsDate(.Cell(flexcpData, lngRow, COL_开始时间)) Then
                '补录医嘱时,项目的开嘱时间缺省为开始时间
                .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开始时间)
                .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开始时间)
            Else
                .TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            End If
            
            '紧急标志
            .TextMatrix(lngRow, COL_标志) = IIF(mbln补录, 2, chk紧急.Value)
            If Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("补录").Picture
                .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
            ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
                Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
                .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
            End If
            
            mblnNoSave = True
            
            Call vsAdvice_AfterRowColChange(-1, -1, lngRow, .Col)
        End If
    End With
End Sub

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal lngRow As Long, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      lngRow=当前输入行
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '删除现有的检查部位行或现有的附加手术及麻醉项目行(修改了时)
    Call Delete检查手术(lngRow)
    
    '重新加入部位行或附加手术行及麻醉项目行
    If int类型 = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    End If
    
    If strDataIDs <> "" Then
        strSQL = "Select * From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '按用户输入项目顺序
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '不可能EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = zlDatabase.GetNextId("病人医嘱记录")
                    .TextMatrix(lngRow + i + 1, COL_相关ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_EDIT) = 1 '新增
                    
                    .TextMatrix(lngRow + i + 1, COL_婴儿) = cbo婴儿.ListIndex
                    .TextMatrix(lngRow + i + 1, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_状态) = 1 '新开
                    .TextMatrix(lngRow + i + 1, COL_期效) = .TextMatrix(lngRow, COL_期效)
                    
                    .TextMatrix(lngRow + i + 1, COL_类别) = rsTmp!类别
                    .TextMatrix(lngRow + i + 1, COL_诊疗项目ID) = rsTmp!ID
                    .TextMatrix(lngRow + i + 1, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                    .TextMatrix(lngRow + i + 1, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                    .TextMatrix(lngRow + i + 1, COL_操作类型) = Nvl(rsTmp!操作类型)
                    .TextMatrix(lngRow + i + 1, COL_处方限量) = Nvl(rsTmp!录入限量)
                    
                    .TextMatrix(lngRow + i + 1, COL_标本部位) = Nvl(rsTmp!标本部位)
                    .TextMatrix(lngRow + i + 1, COL_医嘱内容) = rsTmp!名称
                    
                    .TextMatrix(lngRow + i + 1, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
                    
                    .TextMatrix(lngRow + i + 1, COL_单量) = .TextMatrix(lngRow, COL_单量)
                    .TextMatrix(lngRow + i + 1, COL_总量) = .TextMatrix(lngRow, COL_总量)
                    
                    .TextMatrix(lngRow + i + 1, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                    .TextMatrix(lngRow + i + 1, COL_频率) = .TextMatrix(lngRow, COL_频率)
                    .TextMatrix(lngRow + i + 1, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                    .TextMatrix(lngRow + i + 1, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                    .TextMatrix(lngRow + i + 1, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                    .TextMatrix(lngRow + i + 1, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
                    .Cell(flexcpData, lngRow + i + 1, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
                    
                    '执行性质:根据项目自身设置
                    .TextMatrix(lngRow + i + 1, COL_执行性质) = Nvl(rsTmp!执行科室, 0)
                    
                    '叮嘱和院外执行无执行科室,手术麻醉单独执行科室
                    '否则不管其执行科室设置,一个检查或手术组合应该相同
                    If InStr(",0,5,", Nvl(rsTmp!执行科室, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_执行科室ID) = 0
                    Else
                        If rsTmp!类别 = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsTmp!类别, rsTmp!ID, 0, _
                                Nvl(rsTmp!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
                        Else
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                        End If
                    End If
                    
                    .TextMatrix(lngRow + i + 1, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                    .Cell(flexcpData, lngRow + i + 1, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                    
                    .TextMatrix(lngRow + i + 1, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                    .TextMatrix(lngRow + i + 1, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                    
                    .TextMatrix(lngRow + i + 1, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                    .Cell(flexcpData, lngRow + i + 1, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                    
                    .TextMatrix(lngRow + i + 1, COL_标志) = .TextMatrix(lngRow, COL_标志)
                End With
            Next
                
            '调整序号
            Call AdviceSet医嘱序号(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet给药途径(ByVal lngRow As Long, ByVal lng给药途径ID As Long, Optional str执行性质 As String) As Long
'功能：为录入的中，西成药设置对应的给药途径行(新增或修改)
'参数：lngRow=要处理给药途径的药品行
'      lng给药途径ID=给药途径ID
'      str执行性质=修改给药途径时,当前界面设置的执行性质
'返回：被设置的给药途径行的医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng给药途径ID)
    If rsTmp.EOF Then lng给药途径ID = 0 '没有数据，先设置以保持关系
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then '未设置"相关ID"时
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '修改医嘱的内容时重新设置给药途径内容(不是更换诊疗项目)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
        End If
        
        '无效内容：名称,收费细目ID,剂量系数,住院单位,住院包装,标本部位,医生嘱托,单量,总量,用法
        If blnNew Then
            .RowData(lngNewRow) = zlDatabase.GetNextId("病人医嘱记录")
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngNewRow, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + 1
        Else
            '医嘱ID(RowData),序号:保持不变
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '标志为内容修改
                .TextMatrix(lngNewRow, COL_状态) = 1 '修改后变为新开
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngNewRow, COL_状态) = 1 '新开
        .TextMatrix(lngNewRow, COL_期效) = .TextMatrix(lngRow, COL_期效)
        
        .TextMatrix(lngNewRow, COL_类别) = "E" '给药途径属于治疗
        .TextMatrix(lngNewRow, COL_诊疗项目ID) = lng给药途径ID
        
        '如果没有确定给药途径，暂时不设置的内容
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
            .TextMatrix(lngNewRow, COL_操作类型) = Nvl(rsTmp!操作类型)
            .TextMatrix(lngNewRow, COL_医嘱内容) = rsTmp!名称
            
            .TextMatrix(lngNewRow, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
            
            '执行性质:缺省根据项目设置,修改时根据当前界面设置
            If str执行性质 = "" Then
                .TextMatrix(lngNewRow, COL_执行性质) = Nvl(rsTmp!执行科室, 0)
            Else
                .TextMatrix(lngNewRow, COL_执行性质) = Decode(str执行性质, "离院带药", 5, Nvl(rsTmp!执行科室, 0))
            End If
            
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_执行性质))) = 0 Then
                .TextMatrix(lngNewRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", lng给药途径ID, 0, _
                    Nvl(rsTmp!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
            Else
                .TextMatrix(lngNewRow, COL_执行科室ID) = 0
            End If
        End If
        
        '给药途径天数与药品相同
        .TextMatrix(lngNewRow, COL_天数) = .TextMatrix(lngRow, COL_天数)
        
        .TextMatrix(lngNewRow, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质) '以药品的为准
        .TextMatrix(lngNewRow, COL_频率) = .TextMatrix(lngRow, COL_频率)
        .TextMatrix(lngNewRow, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
        .TextMatrix(lngNewRow, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
        .TextMatrix(lngNewRow, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
        .TextMatrix(lngNewRow, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
        
        .TextMatrix(lngNewRow, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
        .Cell(flexcpData, lngNewRow, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
        
        .TextMatrix(lngNewRow, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
        .TextMatrix(lngNewRow, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
        
        .TextMatrix(lngNewRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
        .Cell(flexcpData, lngNewRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
        
        .TextMatrix(lngNewRow, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
        .Cell(flexcpData, lngNewRow, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
        
        .TextMatrix(lngNewRow, COL_标志) = .TextMatrix(lngRow, COL_标志)
            
        '往后调整序号
        If blnNew Then Call AdviceSet医嘱序号(lngNewRow + 1, 1)
        
        AdviceSet给药途径 = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'功能：根据当前医嘱卡片中的内容，更新当前医嘱内容
'说明：对于ListIndex=-1而对应医嘱项又有内容的，保持原内容不更新
    Dim lngRow As Long, lngBeginRow As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim blnCurDo As Boolean, blnOtherDo As Boolean, blnTmp As Boolean
    Dim bln终止时间 As Boolean, lngTmp As Long, strTmp As String
    Dim strCurDate As String, lng开嘱科室ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long
    
    With vsAdvice
        lngRow = .Row
        
        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub '清除编辑标志
        
        If RowIn配方行(lngRow) Then
            '中药配方
            strTmp = "未计算"
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            For i = lngBeginRow To lngRow
                '修改处理配方的所有行内容(包括煎法和用法)
                If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then
                    .TextMatrix(i, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开始时间) = txt开始时间.Text
                    blnCurDo = True: bln终止时间 = True
                End If
                If chk紧急.Visible And chk紧急.Tag <> "" Then
                    .TextMatrix(i, COL_标志) = chk紧急.Value
                    If i = lngRow Then '用法行显示紧急标志
                        If Val(.TextMatrix(i, COL_标志)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
                        ElseIf Val(.TextMatrix(i, COL_标志)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("紧急").Picture
                        Else
                            Set .Cell(flexcpPicture, i, COL_F标志) = Nothing
                        End If
                        .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
                    End If
                    blnCurDo = True
                End If
                If txt总量.Enabled And IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then
                    .TextMatrix(i, COL_总量) = FormatEx(Val(txt总量.Text), 5)
                    blnCurDo = True: bln终止时间 = True
                End If
                If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                    .TextMatrix(i, COL_频率) = txt频率.Text
                    Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, 2) '中医范围
                    .TextMatrix(i, COL_频率次数) = int频率次数
                    .TextMatrix(i, COL_频率间隔) = int频率间隔
                    .TextMatrix(i, COL_间隔单位) = str间隔单位
                    blnCurDo = True: bln终止时间 = True
                End If
                If cbo执行时间.Tag <> "" Then
                    .TextMatrix(i, COL_执行时间) = cbo执行时间.Text
                    blnCurDo = True: bln终止时间 = True
                End If
                
                If .TextMatrix(i, COL_类别) = "7" Then
                    '更改的是组成中药的执行科室(用法煎法的改不到)
                    If cbo执行科室.ListIndex <> -1 And cbo执行科室.Tag <> "" Then
                        .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                        blnCurDo = True
                    End If
                    
                    '执行性质:配方中所有组成的中药相同
                    If cbo执行性质.Tag <> "" Then
                        .TextMatrix(i, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "自备药", 5, 4)
                        If Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            .TextMatrix(i, COL_执行科室ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            '恢复缺省执行科室,缺省与前面相同
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_类别) = "7" And Val(.TextMatrix(j, COL_执行科室ID)) <> 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(j, COL_执行科室ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, .TextMatrix(i, COL_类别), _
                                        Val(.TextMatrix(i, COL_诊疗项目ID)), Val(.TextMatrix(i, COL_收费细目ID)), 4, mlng病人科室id, 0, cbo期效.ListIndex, , True)
                                End If
                            Else
                                .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngBeginRow, COL_执行科室ID)
                            End If
                        End If
                        blnReInRow = True '界面执行科室编辑性变化
                        blnCurDo = True
                    End If
                End If
                
                If cbo医生.ListIndex <> -1 And cbo医生.Tag <> "" Then
                    .TextMatrix(i, COL_开嘱医生) = NeedName(cbo医生.Text)
                    .TextMatrix(i, COL_开嘱科室ID) = Get开嘱科室ID(cbo医生.ItemData(cbo医生.ListIndex), mlng病人科室id)
                    blnCurDo = True
                End If
                
                If txt开嘱时间.Tag <> "" And IsDate(txt开嘱时间.Text) Then
                    .TextMatrix(i, COL_开嘱时间) = Format(txt开嘱时间.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开嘱时间) = Format(txt开嘱时间.Text, "yyyy-MM-dd HH:mm")
                    blnCurDo = True
                End If
                                                    
                '修改时自动更新部份内容
                blnTmp = False
                If cbo医生嘱托.Tag <> "" Or cbo执行性质.Tag <> "" _
                    Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                    blnTmp = True
                End If
                If blnCurDo Or blnTmp Then
                    '非补录医嘱托修改了内容则更新开嘱时间
                    If Val(.TextMatrix(i, COL_标志)) <> 2 Then
                        If strCurDate = "" Then
                            strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                            txt开嘱时间.Text = strCurDate '界面更新
                        End If
                        .TextMatrix(i, COL_开嘱时间) = Format(strCurDate, "MM-dd HH:mm")
                        .Cell(flexcpData, i, COL_开嘱时间) = strCurDate
                    End If
                    
                    '医生站时严格检查开嘱医生:如果是已审核的，修改医嘱时以审核医生为准。
                    If Not mbln护士站 And GetAuditName(.TextMatrix(i, COL_开嘱医生)) <> UserInfo.姓名 Then
                        .TextMatrix(i, COL_开嘱医生) = UserInfo.姓名
                        Call GetCboIndex(cbo医生, UserInfo.姓名) '界面更新
                        If lng开嘱科室ID = 0 Then
                            lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
                        End If
                        .TextMatrix(i, COL_开嘱科室ID) = lng开嘱科室ID
                    End If
                End If
                                                    
                '终止时间:中药配方长嘱重新计算
                If cbo期效.ListIndex = 0 And bln终止时间 Then
                    If strTmp = "未计算" Then
                        .Cell(flexcpData, i, COL_终止时间) = Calc终止时间(i)
                        .TextMatrix(i, COL_终止时间) = Format(.Cell(flexcpData, i, COL_终止时间), "MM-dd HH:mm")
                        txt终止时间.Text = .Cell(flexcpData, i, COL_终止时间)
                        strTmp = .Cell(flexcpData, i, COL_终止时间)
                    Else
                        .TextMatrix(i, COL_终止时间) = Format(strTmp, "MM-dd HH:mm")
                        .Cell(flexcpData, i, COL_终止时间) = strTmp
                    End If
                End If
                                                    
                If .TextMatrix(i, COL_类别) = "E" And i <> lngRow Then lngTmp = i '煎法行号
                                                    
                '---------------
                If blnCurDo Then '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                        If Not .RowHidden(i) Then Call ReSetColor(i) '用法行才设置
                    End If
                    mblnNoSave = True '标记为未保存
                End If
            Next
            
            '涉及中药用法行的内容:直接更改当前行的内容(煎法行在配方编辑中才能改)
            '-----------------------------------------------------------
            blnCurDo = False
                    
            '医生嘱托:是放在中药用法行(显示行)中的
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If
        
            '中药用法
            If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                .TextMatrix(lngRow, COL_用法) = txt用法.Text
                
                '同时更改计价性质和执行性质
                .TextMatrix(lngRow, COL_计价性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "计价性质"), 0)
                i = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", Val(cmd用法.Tag), 0, _
                        Val(.TextMatrix(lngRow, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), cbo期效.ListIndex)
                End If
                
                blnReInRow = True '需要刷新中药用法执行科室
                blnCurDo = True
            End If
            
            '用法和煎法的执行性质
            If cbo执行性质.Tag <> "" Then
                '用法
                i = Nvl(GetItemField("诊疗项目目录", Val(.TextMatrix(lngRow, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, .TextMatrix(lngRow, COL_类别), _
                        Val(.TextMatrix(lngRow, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngRow, COL_执行性质)), _
                        mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), cbo期效.ListIndex)
                End If
                
                '煎法
                i = Nvl(GetItemField("诊疗项目目录", Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngTmp, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngTmp, COL_执行性质)) = 5 Then
                    .TextMatrix(lngTmp, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, .TextMatrix(lngTmp, COL_类别), _
                        Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngTmp, COL_执行性质)), _
                        mlng病人科室id, Val(.TextMatrix(lngTmp, COL_开嘱科室ID)), cbo期效.ListIndex)
                End If
                
                If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                    .TextMatrix(lngTmp, COL_EDIT) = 2
                    .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                End If
                mblnNoSave = True '标记为未保存
                
                blnCurDo = True
            End If
            
            '中药用法执行科室:即配方当前显示行的执行科室
            If cbo附加执行.ListIndex <> -1 And cbo附加执行.Tag <> "" Then
                .TextMatrix(lngRow, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                blnCurDo = True
            End If
            
            '---------------
            If blnCurDo Then '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                    .TextMatrix(lngRow, COL_EDIT) = 2
                    .TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
                    Call ReSetColor(lngRow)
                End If
                mblnNoSave = True '标记为未保存
            End If
        Else '其它诊疗项目
            If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
                blnCurDo = True
            End If
            If chk紧急.Visible And chk紧急.Tag <> "" Then
                .TextMatrix(lngRow, COL_标志) = chk紧急.Value
                
                '显示紧急标志,一并给药显示在第一行
                lngBeginRow = lngRow
                If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
                End If
                If Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                    Set .Cell(flexcpPicture, lngBeginRow, COL_F标志) = imgFlag.ListImages("补录").Picture
                ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
                    Set .Cell(flexcpPicture, lngBeginRow, COL_F标志) = imgFlag.ListImages("紧急").Picture
                Else
                    Set .Cell(flexcpPicture, lngBeginRow, COL_F标志) = Nothing
                End If
                .Cell(flexcpPictureAlignment, lngBeginRow, COL_F标志) = 4
                
                blnCurDo = True
            End If
            If txt单量.Enabled And (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" Then
                .TextMatrix(lngRow, COL_单量) = FormatEx(txt单量.Text, 5)
                blnCurDo = True
            End If
            
            If txt天数.Tag <> "" Then
                .TextMatrix(lngRow, COL_天数) = txt天数.Text
                blnCurDo = True
            End If
            
            If txt总量.Enabled And IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then
                .TextMatrix(lngRow, COL_总量) = FormatEx(Val(txt总量.Text), 5)
                blnCurDo = True
            End If
            
            If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                '频率性质已经在设置时确定(临嘱可能在一次性之间切换)
                .TextMatrix(lngRow, COL_频率) = txt频率.Text
                Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, Get频率范围(lngRow))
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                blnCurDo = True
            End If
            
            If cbo执行时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_执行时间) = cbo执行时间.Text
                blnCurDo = True
            End If
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If
            
            If cbo执行科室.ListIndex <> -1 And cbo执行科室.Tag <> "" Then
                If Not RowIn检验行(lngRow) Then '采集方法的执行科室不同
                    .TextMatrix(lngRow, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                End If
                blnCurDo = True
            End If
                        
            '附加执行科室：给药途径,手术麻醉,采集方法
            If cbo附加执行.ListIndex <> -1 And cbo附加执行.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_类别) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
                    And .TextMatrix(lngRow - 1, COL_类别) = "C" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                End If
                
                '只更新对应行,不影响其它行
                If lngTmp <> -1 Then
                    .TextMatrix(lngTmp, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                    End If
                    mblnNoSave = True '标记为未保存
                End If
            End If
            
            If (IsDate(txt终止时间.Text) Or txt终止时间.Text = "") And txt终止时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_终止时间) = Format(txt终止时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_终止时间) = txt终止时间.Text
                blnCurDo = True
            End If
            
            If cbo医生.ListIndex <> -1 And cbo医生.Tag <> "" Then
                .TextMatrix(lngRow, COL_开嘱医生) = NeedName(cbo医生.Text)
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(cbo医生.ItemData(cbo医生.ListIndex), mlng病人科室id)
                blnCurDo = True
            End If
            
            If txt开嘱时间.Tag <> "" And IsDate(txt开嘱时间.Text) Then
                .TextMatrix(lngRow, COL_开嘱时间) = Format(txt开嘱时间.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(txt开嘱时间.Text, "yyyy-MM-dd HH:mm")
                blnCurDo = True
            End If
                                    
            '执行性质,给药途径:为更新开嘱时间(包括给药途径的同步更改),先判断是否改变
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If cbo执行性质.Tag <> "" Then blnCurDo = True
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then blnCurDo = True
            End If
                                    
            '修改时自动更新部份内容
            blnTmp = False
            If cbo执行性质.Tag <> "" Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                blnReInRow = True '需要刷新给药途径,采集方式的执行科室
                blnTmp = True
            End If
            If blnCurDo Or blnTmp Then
                '非补录医嘱托修改了内容则更新开嘱时间
                If Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then
                    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                    .TextMatrix(lngRow, COL_开嘱时间) = Format(strCurDate, "MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_开嘱时间) = strCurDate
                    txt开嘱时间.Text = strCurDate '界面更新
                End If
                
                '医生站时严格检查开嘱医生:如果是已审核的，修改医嘱时以审核医生为准。
                If Not mbln护士站 And GetAuditName(.TextMatrix(lngRow, COL_开嘱医生)) <> UserInfo.姓名 Then
                    .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                    Call GetCboIndex(cbo医生, UserInfo.姓名) '界面更新
                    If lng开嘱科室ID = 0 Then
                        lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
                    End If
                    .TextMatrix(lngRow, COL_开嘱科室ID) = lng开嘱科室ID
                End If
            End If
                                    
            '其它需要同步处理的关联行
            '----------------------------------------------------------------
            If RowIn检验行(lngRow) Then
                '采集方法
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    .TextMatrix(lngRow, COL_名称) = txt用法.Text
                    
                    '同时更改计价性质和执行性质
                    .TextMatrix(lngRow, COL_计价性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "计价性质"), 0)
                    .TextMatrix(lngRow, COL_执行性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", Val(cmd用法.Tag), 0, _
                            Val(.TextMatrix(lngRow, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), cbo期效.ListIndex)
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    End If
                    
                    blnCurDo = True
                End If
                
                '设置一并采集的各个检验项目
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                                blnOtherDo = True
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                blnOtherDo = True
                            End If
                            If cbo执行科室.Tag <> "" And cbo执行科室.ListIndex <> -1 Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                Else
                                    .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                                End If
                                blnOtherDo = True
                            End If
                            If txt开始时间.Tag <> "" Then
                                .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                blnOtherDo = True
                            End If
                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                blnOtherDo = True
                            End If
                            If txt终止时间.Tag <> "" Then
                                .TextMatrix(i, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
                                .Cell(flexcpData, i, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
                                blnOtherDo = True
                            End If
                            If cbo医生.Tag <> "" Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            If txt开嘱时间.Tag <> "" Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            If chk紧急.Tag <> "" Then
                                .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                blnOtherDo = True
                            End If
                            
                            '开嘱时间
                            If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            
                            '开嘱医生
                            If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                blnOtherDo = True
                            End If
                            
                            '开嘱科室ID
                            If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            
                            '标记为修改
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '中、西成药处理给药途径及一并给药的情况
                
                '执行性质
                If cbo执行性质.Tag <> "" Then
                    .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "自备药", 5, 4)
                    If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                        '恢复缺省药房,缺省与前面的成药相同
                        strTmp = Get可用药房IDs(.TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), mlng病人科室id)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '西成药和中成药的药房可能不同,所以类别要相同
                            If .TextMatrix(i, COL_类别) = .TextMatrix(lngRow, COL_类别) And Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_执行科室ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_执行科室ID) = Val(.TextMatrix(i, COL_执行科室ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                            .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, .TextMatrix(lngRow, COL_类别), _
                                Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), 4, mlng病人科室id, 0, cbo期效.ListIndex, , True)
                        End If
                    End If
                    
                    cbo执行科室.Tag = "1" '标明执行科室一并给药的要同步变
                    blnReInRow = True '界面执行科室编辑性变化
                End If
                
                '给药途径本身及其它相关数据同步更改
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    Call AdviceSet给药途径(lngRow, Val(cmd用法.Tag), NeedName(cbo执行性质.Text))
                ElseIf blnCurDo Then 'cbo执行性质.Tag <> "" Then
                    '如果执行性质更改了,需要强行修改对应的给药途径的执行性质和执行科室
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call AdviceSet给药途径(lngRow, Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), NeedName(cbo执行性质.Text))
                End If
                
                '一并给药:不处理给药途径,前面已单独设置
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 Then '可能现在中间有空行
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                                If txt开始时间.Tag <> "" Then
                                    .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                    .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                    blnOtherDo = True
                                End If
                                If cbo医生.Tag <> "" Then
                                    .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                    .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                    blnOtherDo = True
                                End If
                                If txt开嘱时间.Tag <> "" Then
                                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                    blnOtherDo = True
                                End If
                                If txt用法.Tag <> "" Then
                                    .TextMatrix(i, COL_用法) = .TextMatrix(lngRow, COL_用法)
                                    blnOtherDo = True
                                End If
                                If txt频率.Tag <> "" Then
                                    .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质) '需要同步设置,因为临嘱可能在一次性之间切换
                                    .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                    .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                    .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                    .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                    blnOtherDo = True
                                End If
                                    
                                '一并给药的,天数相同变化,总量重新计算
                                If txt天数.Tag <> "" Then
                                    .TextMatrix(i, COL_天数) = .TextMatrix(lngRow, COL_天数)
                                    If .TextMatrix(i, COL_频率) <> "" And Val(.TextMatrix(i, COL_频率性质)) <> 1 _
                                        And Val(.TextMatrix(i, COL_单量)) <> 0 _
                                        And Val(.TextMatrix(i, COL_剂量系数)) <> 0 _
                                        And Val(.TextMatrix(i, COL_住院包装)) <> 0 Then
                                        
                                        .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                            Val(.TextMatrix(i, COL_单量)), Val(.TextMatrix(i, COL_天数)), _
                                            Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), _
                                            .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), _
                                            Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_住院包装)), _
                                            Val(.TextMatrix(i, COL_可否分零))), 5)
                                    End If
                                    blnOtherDo = True
                                End If
                                    
                                If cbo执行时间.Tag <> "" Then
                                    .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                    blnOtherDo = True
                                End If
                                If txt终止时间.Tag <> "" Then
                                    .TextMatrix(i, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
                                    .Cell(flexcpData, i, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
                                    blnOtherDo = True
                                End If
                                
                                '执行性质:离院带药在一并给药中需一致，其它可单独设置
                                If cbo执行性质.Tag <> "" And NeedName(cbo执行性质.Text) = "离院带药" Then
                                    .TextMatrix(i, COL_执行性质) = .TextMatrix(lngRow, COL_执行性质)
                                    '由自备药转过来时需要重新设置执行科室
                                    If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                    End If
                                    blnOtherDo = True
                                End If
                                
                                '执行科室:执行科室(药房)可以不同
'                                If cbo执行科室.Tag <> "" Then
'                                    .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
'                                    blnOtherDo = True
'                                End If
                                
                                '补录时不允许在一并给药之中插入药品
                                If chk紧急.Tag <> "" Then
                                    .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱时间
                                If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱医生
                                If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                    .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱科室ID
                                If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                    .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                    blnOtherDo = True
                                End If
                                
                                '标记为修改
                                If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                    .TextMatrix(i, COL_EDIT) = 2
                                    .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_类别)) > 0 And blnCurDo Then
                '检查组合项目行或手术附加行
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt单量.Tag <> "" Then
                                .TextMatrix(i, COL_单量) = .TextMatrix(lngRow, COL_单量)
                                blnOtherDo = True
                            End If
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                                blnOtherDo = True
                            End If
                            
                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                blnOtherDo = True
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                blnOtherDo = True
                            End If
                            If cbo执行科室.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                ElseIf .TextMatrix(i, COL_类别) <> "G" Then '手术麻醉的执行科室为单独
                                    .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                End If
                                blnOtherDo = True
                            End If
                            If txt开始时间.Tag <> "" Then
                                .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                blnOtherDo = True
                            End If
                            If txt终止时间.Tag <> "" Then
                                .TextMatrix(i, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
                                .Cell(flexcpData, i, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
                                blnOtherDo = True
                            End If
                            If cbo医生.Tag <> "" Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            If txt开嘱时间.Tag <> "" Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            If chk紧急.Tag <> "" Then
                                .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                blnOtherDo = True
                            End If
                            
                            '开嘱时间
                            If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            
                            '开嘱医生
                            If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                blnOtherDo = True
                            End If
                            
                            '开嘱科室ID
                            If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            
                            '标记为修改
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
                     
            If blnCurDo Then '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                    .TextMatrix(lngRow, COL_EDIT) = 2
                    .TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
                    Call ReSetColor(lngRow)
                End If
                mblnNoSave = True '标记为未保存
            End If
        End If
                
        '更新医嘱内容
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, COL_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = .TextMatrix(lngRow, COL_医嘱内容)
        End If
    End With
        
    '清除编辑标志
    Call ClearItemTag
    
    '某些情况下需要重新设置卡片的项目编辑性(如修改了执行性质时)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub ReSetColor(ByVal lngRow As Long)
'功能：重新设置指定行的颜色
'说明：因为疑问的医嘱编辑后变成新开
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        '一并给药范围
        lngBegin = lngRow: lngEnd = lngRow
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If RowIn一并给药(lngRow) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
            End If
        End If
        '恢复成正常色
        For i = lngBegin To lngEnd
            .Cell(flexcpForeColor, i, .FixedCols, i, COL_校对护士) = .ForeColor
            '毒麻精的颜色标识
            If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(i, COL_毒理分类)) > 0 _
                And .TextMatrix(i, COL_毒理分类) <> "" Then
                .Cell(flexcpFontBold, i, COL_医嘱内容) = True
            End If
        Next
        .ForeColorSel = .Cell(flexcpForeColor, lngRow, COL_开始时间)
    End With
End Sub

Private Sub AdviceSet一并给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：将选择范围内的药品设置为一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
'说明：以第一行药品的给药途径为准,但位置放在最后一行药品之后
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng相关ID As Long, i As Long
    Dim strStart As String, curDate As Date
    
    lngRow1 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngBegin, COL_相关ID)), lngBegin + 1) '第一给药途径行
    lngRow2 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1) '最后给药途径行
    
    '删除给药途径行之前记录执行性质,以便后面作判断
    For i = lngRow2 To lngRow1 Step -1
        If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex And vsAdvice.RowHidden(i) Then
            vsAdvice.Cell(flexcpData, i - 1, COL_执行性质) = Val(vsAdvice.TextMatrix(i, COL_执行性质))
        End If
    Next
    
    '复制第一行的给药途径到最后一行的给药途径
    For i = vsAdvice.FixedCols To vsAdvice.Cols - 1
        If i <> COL_EDIT And i <> COL_相关ID And i <> COL_序号 And i <> COL_状态 Then
            vsAdvice.TextMatrix(lngRow2, i) = vsAdvice.TextMatrix(lngRow1, i)
        End If
    Next
    '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号
    If InStr(",0,3,", vsAdvice.TextMatrix(lngRow2, COL_EDIT)) > 0 Then
        vsAdvice.TextMatrix(lngRow2, COL_EDIT) = 2 '标记为已修改
        vsAdvice.TextMatrix(lngRow2, COL_状态) = 1 '修改后变为新开
    End If
    lng相关ID = vsAdvice.RowData(lngRow2)
    
    varTmp1 = mblnRowChange: varTmp2 = vsAdvice.Redraw
    mblnRowChange = False: vsAdvice.Redraw = flexRDNone
    
    '删除除最后一行给药途径外的其它给药途径
    For i = lngEnd To lngBegin Step -1
        If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
            If vsAdvice.RowHidden(i) Then
                Call DeleteRow(i)
            Else
                vsAdvice.TextMatrix(i, COL_相关ID) = lng相关ID
                If InStr(",0,3,", vsAdvice.TextMatrix(i, COL_EDIT)) > 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                    vsAdvice.TextMatrix(i, COL_状态) = 1 '修改后变为新开
                End If
            End If
        End If
    Next
    
    '行号已变更
    lngRow1 = lngBegin '开始一并给药行
    curDate = zlDatabase.Currentdate
    
    '检查医生是否变更
    If Not mbln护士站 And GetAuditName(vsAdvice.TextMatrix(lngRow1, COL_开嘱医生)) <> UserInfo.姓名 Then
        '更新相关信息:前面已标记为修改,且手工操作完成时已有进入界面刷新
        vsAdvice.TextMatrix(lngRow1, COL_开嘱医生) = UserInfo.姓名
        vsAdvice.TextMatrix(lngRow1, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
        If Val(vsAdvice.TextMatrix(lngRow1, COL_标志)) <> 2 Then '不是补录医嘱时
            vsAdvice.TextMatrix(lngRow1, COL_开嘱时间) = Format(curDate, "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow1, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    For i = lngRow1 + 1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
            lngRow2 = i '记录新的结束行号
            
            '一并给药的部分信息相同
            vsAdvice.TextMatrix(i, COL_开始时间) = vsAdvice.TextMatrix(lngRow1, COL_开始时间)
            vsAdvice.Cell(flexcpData, i, COL_开始时间) = vsAdvice.Cell(flexcpData, lngRow1, COL_开始时间)
            
            vsAdvice.TextMatrix(i, COL_开嘱医生) = vsAdvice.TextMatrix(lngRow1, COL_开嘱医生)
            vsAdvice.TextMatrix(i, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngRow1, COL_开嘱科室ID)
            
            vsAdvice.TextMatrix(i, COL_开嘱时间) = vsAdvice.TextMatrix(lngRow1, COL_开嘱时间) '一并给药的开嘱时间相同
            vsAdvice.Cell(flexcpData, i, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngRow1, COL_开嘱时间)
            
            vsAdvice.TextMatrix(i, COL_用法) = vsAdvice.TextMatrix(lngRow1, COL_用法)
            
            vsAdvice.TextMatrix(i, COL_频率性质) = vsAdvice.TextMatrix(lngRow1, COL_频率性质)
            vsAdvice.TextMatrix(i, COL_频率) = vsAdvice.TextMatrix(lngRow1, COL_频率)
            vsAdvice.TextMatrix(i, COL_频率次数) = vsAdvice.TextMatrix(lngRow1, COL_频率次数)
            vsAdvice.TextMatrix(i, COL_频率间隔) = vsAdvice.TextMatrix(lngRow1, COL_频率间隔)
            vsAdvice.TextMatrix(i, COL_间隔单位) = vsAdvice.TextMatrix(lngRow1, COL_间隔单位)
            vsAdvice.TextMatrix(i, COL_执行时间) = vsAdvice.TextMatrix(lngRow1, COL_执行时间)
            
            vsAdvice.TextMatrix(i, COL_终止时间) = vsAdvice.TextMatrix(lngRow1, COL_终止时间)
            vsAdvice.Cell(flexcpData, i, COL_终止时间) = vsAdvice.Cell(flexcpData, lngRow1, COL_终止时间)
            
            vsAdvice.TextMatrix(i, COL_标志) = vsAdvice.TextMatrix(lngRow1, COL_标志)
            Set vsAdvice.Cell(flexcpPicture, i, COL_F标志) = Nothing '在开始行显示
            
            If Val(vsAdvice.TextMatrix(lngRow1, COL_执行性质)) <> 5 And Val(vsAdvice.Cell(flexcpData, lngRow1, COL_执行性质)) = 5 Then
                '第一行是离院带药,全部设置为离院带药
                vsAdvice.TextMatrix(i, COL_执行性质) = vsAdvice.TextMatrix(lngRow1, COL_执行性质)
                vsAdvice.TextMatrix(i, COL_执行科室ID) = vsAdvice.TextMatrix(lngRow1, COL_执行科室ID)
            ElseIf Val(vsAdvice.TextMatrix(i, COL_执行性质)) <> 5 And Val(vsAdvice.Cell(flexcpData, i, COL_执行性质)) = 5 Then
                '当前行是离院带药,则设置为与第一行相同
                vsAdvice.TextMatrix(i, COL_执行性质) = vsAdvice.TextMatrix(lngRow1, COL_执行性质)
                vsAdvice.TextMatrix(i, COL_执行科室ID) = vsAdvice.TextMatrix(lngRow1, COL_执行科室ID)
            Else
                '否则保持不变
            End If
            
'            '执行性质:一并给的相同,并缺省按第一行设置
'            vsAdvice.TextMatrix(i, COL_执行性质) = vsAdvice.TextMatrix(lngRow1, COL_执行性质)
'            '执行科室:执行科室(药房)可以不同
'            vsAdvice.TextMatrix(i, COL_执行科室ID) = vsAdvice.TextMatrix(lngRow1, COL_执行科室ID)
            
            '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
            If InStr(",0,3,", vsAdvice.TextMatrix(i, COL_EDIT)) > 0 Then
                vsAdvice.TextMatrix(i, COL_EDIT) = 2
                vsAdvice.TextMatrix(i, COL_状态) = 1 '修改后变为新开
            End If
        Else
            Exit For
        End If
    Next
    
    '开始执行时间处理(非补录新开的不能太早)
    If Not mbln补录 Then
        strStart = ""
        For i = lngRow1 To lngRow2
            If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 1 Then
                If DateDiff("n", CDate(vsAdvice.Cell(flexcpData, i, COL_开始时间)), curDate) > 30 Then
                    strStart = GetDefaultTime(i): Exit For
                End If
            End If
        Next
        If strStart <> "" Then
            For i = lngRow1 To lngRow2 + 1
                vsAdvice.Cell(flexcpData, i, COL_开始时间) = strStart
                vsAdvice.TextMatrix(i, COL_开始时间) = Format(strStart, "MM-dd HH:mm")
            Next
        End If
    End If
    mblnRowChange = varTmp1: vsAdvice.Redraw = varTmp2
    mblnNoSave = True '标记为未保存
End Sub

Private Sub AdviceSet单独给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：取消一组药品的一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng给药途径ID As Long, i As Long
    Dim int执行性质 As Integer, str执行性质 As String
    Dim lngRow As Long, curDate As Date, blnUpdate As Boolean
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        '一并给药途径
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1)
        lng给药途径ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
        int执行性质 = Val(.TextMatrix(lngRow, COL_执行性质))
                        
        '检查医生变更:以给药途径行为准变化
        If Not mbln护士站 And GetAuditName(.TextMatrix(lngRow, COL_开嘱医生)) <> UserInfo.姓名 Then
            '更新相关信息:手工操作完成时有进入界面刷新
            .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
            .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            If Val(.TextMatrix(lngRow, COL_标志)) <> 2 Then '不是补录医嘱时
                curDate = zlDatabase.Currentdate
                .TextMatrix(lngRow, COL_开嘱时间) = Format(curDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            End If
            
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2 '标记为已修改
                .TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
            End If
            blnUpdate = True
        End If
                        
        '显示紧急标志:每一行
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_标志)) = 2 Then
                Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
            ElseIf Val(.TextMatrix(i, COL_标志)) = 1 Then
                Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("紧急").Picture
            Else
                Set .Cell(flexcpPicture, i, COL_F标志) = Nothing
            End If
            .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
            
            '药品行相应变化
            If blnUpdate Then
                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                If Val(.TextMatrix(i, COL_标志)) <> 2 Then '不是补录医嘱时
                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                End If
                
                If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    .TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                    .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                End If
            End If
        Next
        
        For i = lngEnd - 1 To lngBegin Step -1 '必须反向
            '设置给药途径行
            If Val(.TextMatrix(i, COL_执行性质)) = 5 And int执行性质 <> 5 Then
                str执行性质 = "自备药"
            ElseIf Val(.TextMatrix(i, COL_执行性质)) <> 5 And int执行性质 = 5 Then
                str执行性质 = "离院带药"
            Else
                str执行性质 = ""
            End If
            .TextMatrix(i, COL_相关ID) = "" '必须清除作为标志
            .TextMatrix(i, COL_相关ID) = AdviceSet给药途径(i, lng给药途径ID, str执行性质)
            
            If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                .TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
            End If
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '标记为未保存
    End With
End Sub

Private Sub ShowAdvice()
'功能：显示当前界面条件下的医嘱记录
'说明：1.根据程序编辑方式,相关的数据行是按序号严格排列在一在的。
'      2.这里不处理一并给药的边框及配方行高，状态颜色等格式内容,它们已在读取或编辑时设置
    Dim lngRow As Long, blnHide As Boolean, i As Long
    
    Screen.MousePointer = 11
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
        
    '先删除无效行
    For i = vsAdvice.Rows - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) = 0 Then vsAdvice.RemoveItem i
    Next
    
    '根据当前期效,婴儿显示
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                blnHide = False
                '隐藏以下数据行：
                '1.成药的给药途径行
                '2.手术的附加手术及麻醉项目行
                '3.检查组合的部位行
                '4.中药配方的组成味中药及中药煎法行
                '5.(一并采集的)检验项目
                If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
            Else
                .RowHidden(i) = True
            End If
        Next
    End With
    
    '没有数据行,添加一行空
    If lngRow = 0 Then
        vsAdvice.AddItem ""
        lngRow = vsAdvice.Rows - 1
    End If
    
    vsAdvice.Row = lngRow
    If vsAdvice.RowData(lngRow) = 0 Then
        vsAdvice.Col = vsAdvice.FixedCols
    Else
        vsAdvice.Col = COL_医嘱内容
    End If
    vsAdvice.Redraw = flexRDDirect
    mblnRowChange = True
    
    '显示当前行:进入时在FormLoad中处理,以加快速度
    If Me.Visible Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Screen.MousePointer = 0
End Sub

Private Function SaveAdvice() As Boolean
'功能：保存当前病人的医嘱记录
    Dim arrSQL As Variant
    Dim arrVerify As Variant, arrDelID() As String
    Dim dbl总量 As Double, i As Long
    
    'Pass自动用药审查
    If gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 And Not mbln护士站 Then
        If AdviceCheckWarn(1) = 3 Then
            If MsgBox("合理用药监测系统审查出存在黑灯用药，要继续保存操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Screen.MousePointer = 11
    
    '生成SQL
    arrSQL = Array()
    arrVerify = Array()
        
    '删除了的记录
    arrDelID = Split(mstrDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & Val(arrDelID(i)) & ")"
        End If
    Next
                
    '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then '所有医嘱记录
                '总量转换
                dbl总量 = 0
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    If Val(.TextMatrix(i, COL_总量)) <> 0 Then
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            '成药转换成零售单位
                            dbl总量 = Format(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_住院包装)), "0.00000")
                        Else
                            '中药配方付数或非药临嘱总量,不转换
                            dbl总量 = Val(.TextMatrix(i, COL_总量))
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_EDIT)) = 3 Then '修改了序号的记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新序号(" & .RowData(i) & "," & Val(.TextMatrix(i, COL_序号)) & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 2 Then '修改了内容的记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Update(" & _
                        .RowData(i) & "," & ZVal(.TextMatrix(i, COL_相关ID)) & "," & _
                        Val(.TextMatrix(i, COL_序号)) & "," & Val(.TextMatrix(i, COL_状态)) & "," & _
                        IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1) & "," & _
                        ZVal(.TextMatrix(i, COL_诊疗项目ID)) & "," & ZVal(.TextMatrix(i, COL_天数)) & "," & _
                        ZVal(.TextMatrix(i, COL_单量)) & "," & ZVal(dbl总量) & "," & _
                        "'" & Replace(.TextMatrix(i, COL_医嘱内容), "'", "''") & "','" & Replace(.TextMatrix(i, COL_医生嘱托), "'", "''") & "'," & _
                        "'" & .TextMatrix(i, COL_标本部位) & "','" & .TextMatrix(i, COL_频率) & "'," & _
                        ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & "," & _
                        "'" & .TextMatrix(i, COL_间隔单位) & "','" & .TextMatrix(i, COL_执行时间) & "'," & _
                        Val(.TextMatrix(i, COL_计价性质)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & _
                        Val(.TextMatrix(i, COL_执行性质)) & "," & Val(.TextMatrix(i, COL_标志)) & "," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        mlng病人科室id & "," & Val(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then '新增的记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                        .RowData(i) & "," & ZVal(.TextMatrix(i, COL_相关ID)) & "," & _
                        Val(.TextMatrix(i, COL_序号)) & ",2," & mlng病人ID & "," & mlng主页ID & "," & _
                        Val(.TextMatrix(i, COL_婴儿)) & "," & Val(.TextMatrix(i, COL_状态)) & "," & _
                        IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1) & "," & _
                        "'" & IIF(.TextMatrix(i, COL_类别) = "*", "", .TextMatrix(i, COL_类别)) & "'," & _
                        ZVal(.TextMatrix(i, COL_诊疗项目ID)) & "," & ZVal(.TextMatrix(i, COL_收费细目ID)) & "," & _
                        ZVal(.TextMatrix(i, COL_天数)) & "," & ZVal(.TextMatrix(i, COL_单量)) & "," & ZVal(dbl总量) & "," & _
                        "'" & Replace(.TextMatrix(i, COL_医嘱内容), "'", "''") & "','" & Replace(.TextMatrix(i, COL_医生嘱托), "'", "''") & "'," & _
                        "'" & .TextMatrix(i, COL_标本部位) & "','" & .TextMatrix(i, COL_频率) & "'," & _
                        ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & "," & _
                        "'" & .TextMatrix(i, COL_间隔单位) & "','" & .TextMatrix(i, COL_执行时间) & "'," & _
                        Val(.TextMatrix(i, COL_计价性质)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & _
                        Val(.TextMatrix(i, COL_执行性质)) & "," & Val(.TextMatrix(i, COL_标志)) & "," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        mlng病人科室id & "," & Val(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "NULL," & ZVal(mlng前提ID) & ")"

                    '自动进行校对计价
                    If mbln校对 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                        '只能对有资格医生下达或已审核的医嘱进行自动校对。
                        If HaveAuditPriv(GetAuditName(.TextMatrix(i, COL_开嘱医生))) Then
                            ReDim Preserve arrVerify(UBound(arrVerify) + 1)
                            arrVerify(UBound(arrVerify)) = "ZL_病人医嘱记录_校对(" & .RowData(i) & ",3," & _
                                "To_Date('" & Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),1)"
                        End If
                    End If
                End If
                
                'Pass:更新审查结果
                If Val(.Cell(flexcpData, i, COL_序号)) = 1 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & .RowData(i) & "," & _
                        IIF(CStr(.Cell(flexcpData, i, COL_警示)) = "", "NULL", Val(.Cell(flexcpData, i, COL_警示))) & ")"
                End If
            End If
        Next
    End With
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    For i = 0 To UBound(arrVerify)
        zlDatabase.ExecuteProcedure CStr(arrVerify(i)), Me.Caption
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    '保存成功后,所有记录变成原始记录
    With vsAdvice
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, COL_EDIT) = 0
                .Cell(flexcpData, i, COL_序号) = Empty 'Pass:保存后清除标志
            End If
        Next
    End With
    
    Screen.MousePointer = 0
    mblnNoSave = False
    mstrDelIDs = ""
    SaveAdvice = True
    mblnOK = True
    
    '调整医嘱序号并刷新
    Call AdviceClearUp(Not mbln校对)
    If Not mbln校对 Then
        '保存后重新进入行(比如开始时间不准改了)
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Else
        '重新读取显示医嘱
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        If txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        Else
            vsAdvice.SetFocus
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceClearUp(ByVal blnRefresh As Boolean)
'功能：重新整理医嘱序号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    strSQL = "Select Count(*) as Num From (Select 序号,Count(ID) From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] Having Count(ID)>1 Group by 序号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Sub
    If Nvl(rsTmp!Num, 0) = 0 Then Screen.MousePointer = 0: Exit Sub
    
    strSQL = "ZL_病人医嘱记录_更新序号(NULL,NULL," & mlng病人ID & "," & mlng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '刷新数据
    If blnRefresh Then
        '重新读取显示医嘱
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        If txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        Else
            vsAdvice.SetFocus
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LoadAdvice() As Boolean
'功能：读取当前病人的医嘱记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln配方 As Boolean
    Dim blnFirst As Boolean, i As Long, j As Long
    Dim strCurr As String, strTime As String
    Dim blnDo As Boolean, strPassWhere As String 'Pass
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
                
    '下医嘱缺省的天数
    If msng天数 = 0 Then msng天数 = 1
                
    '医技编辑时只显示医技的,医生编辑时不显示医技的,护士编辑时显示所有
    strSQL = IIF(mlng前提ID <> 0, " And A.前提ID+0=[3]", IIF(Not mbln护士站, " And A.前提ID is NULL", ""))
    
    '医嘱审核的条件设置
    If mbln审核 Then
        '医生审核时:只包含未审核医嘱
        strSQL = strSQL & " And A.医嘱状态=1 And Instr(A.开嘱医生,'/')=0" & _
            " And Not Exists(Select M.姓名 From 人员表 M,执业类别 N Where M.姓名=A.开嘱医生 And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))"
    ElseIf Not mbln护士站 And HaveAuditPriv Then
        '有资格医生下达时:不包含未审核医嘱(在审核功能中包含,以示区分)
        strSQL = strSQL & " And Not (A.医嘱状态=1 And Instr(A.开嘱医生,'/')=0" & _
            " And Not Exists(Select M.姓名 From 人员表 M,执业类别 N Where M.姓名=A.开嘱医生 And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))" & _
            " )"
    End If
    
    'PASS启用时的条件
    '不读取：4-已作废，8-已停止，9-已确认废止；包含当天的术后医嘱
    If gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 And Not mbln护士站 Then
        '包含当天的临嘱(包括已执行),和未停止的长嘱
        strPassWhere = " And A.医嘱状态<>4 And (A.医嘱状态 Not IN(8,9)" & _
            " Or A.医嘱期效=1 And Trunc(A.开嘱时间)=Trunc(Sysdate)" & _
            " Or A.诊疗类别='Z' And B.操作类型='4' And Trunc(A.开嘱时间)=Trunc(Sysdate))"
    Else
        strPassWhere = " And A.医嘱状态<>4 And (A.医嘱状态 Not IN(8,9)" & _
            " Or A.诊疗类别='Z' And B.操作类型='4' And Trunc(A.开嘱时间)=Trunc(Sysdate))"
    End If
    
    strSQL = _
        " Select A.ID,A.相关ID,Nvl(A.婴儿,0) as 婴儿,A.序号,A.医嘱期效," & _
        " A.医嘱状态,A.诊疗类别,A.诊疗项目ID,B.名称,A.标本部位,A.收费细目ID," & _
        " A.开始执行时间,A.医嘱内容,A.医生嘱托,A.单次用量,A.天数,A.总给予量,B.计算单位," & _
        " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,B.计算方式,B.执行频率,B.操作类型," & _
        " A.计价特性,A.执行时间方案,A.执行性质,A.执行科室ID," & _
        " A.执行终止时间,A.开嘱科室ID,A.开嘱医生,A.开嘱时间,A.校对护士," & _
        " A.紧急标志,C.处方限量,C.处方职务,C.毒理分类,C.药品剂型," & _
        " D.剂量系数,D.住院包装,D.住院单位,D.可否分零,A.申请ID,A.审查结果," & _
        " Decode(S.签名ID,NULL,0,1) as 签名否" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,药品特性 C,药品规格 D,病人医嘱状态 S" & _
        " Where A.诊疗项目ID=B.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
        " And A.收费细目ID=D.药品ID(+) And A.ID=S.医嘱ID And S.操作类型=1" & strPassWhere & strSQL & _
        " And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And A.病人来源<>3" & _
        " Order by 婴儿,序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, mlng前提ID)
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        mblnRowChange = False
        strCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        With vsAdvice
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                bln配方 = False
                
                .RowData(i) = CLng(rsTmp!ID)
                .TextMatrix(i, COL_EDIT) = IIF(mbln审核, 2, 0) '本为原始记录,审核时表示更改了。
                .TextMatrix(i, COL_相关ID) = Nvl(rsTmp!相关ID)
                .TextMatrix(i, COL_婴儿) = Nvl(rsTmp!婴儿, 0)
                .TextMatrix(i, COL_序号) = rsTmp!序号
                .TextMatrix(i, COL_期效) = IIF(Nvl(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .TextMatrix(i, COL_状态) = Nvl(rsTmp!医嘱状态, 0)
                
                .TextMatrix(i, COL_类别) = Nvl(rsTmp!诊疗类别, "*")
                .TextMatrix(i, COL_诊疗项目ID) = Nvl(rsTmp!诊疗项目ID)
                .TextMatrix(i, COL_名称) = Nvl(rsTmp!名称)
                .TextMatrix(i, COL_标本部位) = Nvl(rsTmp!标本部位)
                .TextMatrix(i, COL_收费细目ID) = Nvl(rsTmp!收费细目ID)
                .TextMatrix(i, COL_医嘱内容) = Nvl(rsTmp!医嘱内容)
                .TextMatrix(i, COL_医生嘱托) = Nvl(rsTmp!医生嘱托)
                
                .TextMatrix(i, COL_计价性质) = Nvl(rsTmp!计价特性, 0)
                .TextMatrix(i, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                
                '临嘱可选频率的可能被设置为了一次性
                If Nvl(rsTmp!医嘱期效, 0) = 1 And Nvl(rsTmp!执行频率, 0) = 0 _
                    And Nvl(rsTmp!频率次数, 0) = 0 And Nvl(rsTmp!频率间隔, 0) = 0 Then
                    .TextMatrix(i, COL_频率性质) = 1
                Else
                    .TextMatrix(i, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                End If
                
                .TextMatrix(i, COL_操作类型) = Nvl(rsTmp!操作类型)
                .TextMatrix(i, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(i, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!处方限量)
                .TextMatrix(i, COL_处方职务) = Nvl(rsTmp!处方职务)
                .TextMatrix(i, COL_剂量系数) = Nvl(rsTmp!剂量系数)
                .TextMatrix(i, COL_住院包装) = Nvl(rsTmp!住院包装)
                .TextMatrix(i, COL_住院单位) = Nvl(rsTmp!住院单位)
                If Not IsNull(rsTmp!剂量系数) Then
                    .TextMatrix(i, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
                End If
                
                .TextMatrix(i, COL_开始时间) = Format(rsTmp!开始执行时间, "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_开始时间) = Format(rsTmp!开始执行时间, "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COL_频率) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, COL_频率次数) = Nvl(rsTmp!频率次数)
                .TextMatrix(i, COL_频率间隔) = Nvl(rsTmp!频率间隔)
                .TextMatrix(i, COL_间隔单位) = Nvl(rsTmp!间隔单位)
                .TextMatrix(i, COL_执行时间) = Nvl(rsTmp!执行时间方案)
                
                .TextMatrix(i, COL_终止时间) = Format(Nvl(rsTmp!执行终止时间), "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_终止时间) = Format(Nvl(rsTmp!执行终止时间), "yyyy-MM-dd HH:mm")
                                            
                .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID)
                .TextMatrix(i, COL_执行性质) = Nvl(rsTmp!执行性质, 0)
                
                If Nvl(rsTmp!诊疗类别) = "E" Then
                    If Nvl(rsTmp!相关ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_相关ID)) = rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是成药的给药途径,可能是一并给药的
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = rsTmp!ID Then
                                    '显示给药途径
                                    .TextMatrix(j, COL_用法) = rsTmp!名称
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是中药配方的用法,即配方显示行
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                            bln配方 = True
                        ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                        End If
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        '当前记录是中药配方煎法行
                        bln配方 = True
                    End If
                ElseIf Nvl(rsTmp!诊疗类别) = "7" Then
                    bln配方 = True
                End If
                
                '单量
                .TextMatrix(i, COL_单量) = FormatEx(Nvl(rsTmp!单次用量), 5)
                If InStr(",5,6,7,", Nvl(rsTmp!诊疗类别, "*")) > 0 Or Nvl(rsTmp!计算方式, 0) <> 3 Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!计算单位)
                End If
                
                '天数
                .TextMatrix(i, COL_天数) = Nvl(rsTmp!天数, 0)
                '取最近新开医嘱的开数作为缺省天数
                If InStr(",1,2,", Nvl(rsTmp!医嘱状态, 0)) > 0 _
                    And InStr(",5,6,", Nvl(rsTmp!诊疗类别, "*")) > 0 _
                    And Nvl(rsTmp!医嘱期效, 0) = 1 And Nvl(rsTmp!天数, 0) <> 0 Then
                    msng天数 = Nvl(rsTmp!天数, 1)
                End If
                
                '总量
                If InStr(",5,6,", Nvl(rsTmp!诊疗类别, "*")) > 0 Then
                    '成药临嘱有总量,以零售单位存放,住院单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!住院包装) Then
                        .TextMatrix(i, COL_总量) = FormatEx(rsTmp!总给予量 / rsTmp!住院包装, 5)
                    End If
                    If Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!住院单位)
                    End If
                Else
                    '其它情况有中药和其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then
                        .TextMatrix(i, COL_总量) = rsTmp!总给予量
                    End If
                    If bln配方 Then
                        .TextMatrix(i, COL_总量单位) = "付" '中药配方总量单位为"付"
                    ElseIf Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If

                .TextMatrix(i, COL_开嘱科室ID) = rsTmp!开嘱科室ID
                If mbln审核 Then
                    .TextMatrix(i, COL_开嘱医生) = rsTmp!开嘱医生 & "/" & UserInfo.姓名
                Else
                    .TextMatrix(i, COL_开嘱医生) = rsTmp!开嘱医生
                End If
                
                .TextMatrix(i, COL_开嘱时间) = Format(rsTmp!开嘱时间, "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_开嘱时间) = Format(rsTmp!开嘱时间, "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COL_校对护士) = Nvl(rsTmp!校对护士)
                                                
                '显示紧急标志:一并给药只显示在第一行
                .TextMatrix(i, COL_标志) = Nvl(rsTmp!紧急标志, 0)
                blnFirst = True
                If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Nvl(rsTmp!紧急标志, 0) = 2 Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
                    ElseIf Nvl(rsTmp!紧急标志, 0) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("紧急").Picture
                    End If
                End If
                
                '根据医嘱状态,欺效和药品毒理设置颜色
                '-------------------------------------------------------------------
                '医嘱颜色
                blnDo = False
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
                If rsTmp!医嘱状态 = 2 Then
                    '校对疑问
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80& '深红
                    blnDo = True
                ElseIf InStr(",8,9,", rsTmp!医嘱状态) > 0 Then
                    '已停止,已确认停止:长嘱都以终止时间进行判断
                    If strCurr >= Nvl(rsTmp!执行终止时间) Or Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                        blnDo = True
                    End If
                ElseIf rsTmp!医嘱状态 = 6 Then
                    '已暂停
                    strTime = Format(GetAdviceTime(rsTmp!ID, 6), "yyyy-MM-dd HH:mm")
                    If strCurr >= strTime Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '深绿
                        blnDo = True
                    End If
                ElseIf rsTmp!医嘱状态 = 7 Then
                    '已启用
                    strTime = Format(GetAdviceTime(rsTmp!ID, 7), "yyyy-MM-dd HH:mm")
                    If strCurr < strTime Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '深绿
                        blnDo = True
                    End If
                End If
                If Not blnDo And rsTmp!医嘱状态 <> 1 Then
                    '已通过校对(也包含后续的多个状态)
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '深蓝
                End If
                
                '校对后术后医嘱红色显示
                If .TextMatrix(i, COL_类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 4 _
                    And InStr(",1,2,4,", rsTmp!医嘱状态) = 0 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '红色
                End If
                
                '毒麻精药品标识:中药配方及组成味中药不处理
                If InStr(",5,6,", Nvl(rsTmp!诊疗类别, "*")) > 0 And Not IsNull(rsTmp!毒理分类) Then
                    If InStr(",麻醉药,毒性药,精神药,", rsTmp!毒理分类) > 0 Then
                        .Cell(flexcpFontBold, i, COL_医嘱内容) = True
                    End If
                End If
                
                'Pass根据审查结果显示警示灯
                If Not IsNull(rsTmp!审查结果) Then
                    .Cell(flexcpData, i, COL_警示) = CStr(Nvl(rsTmp!审查结果))
                    Set .Cell(flexcpPicture, i, COL_警示) = imgPass.ListImages(rsTmp!审查结果 + 1).Picture
                End If
                
                '电子签名标识
                .TextMatrix(i, COL_签名否) = Nvl(rsTmp!签名否)
                If Val(.TextMatrix(i, COL_签名否)) = 1 Then
                    Set .Cell(flexcpPicture, i, COL_医嘱内容) = imgSign.ListImages(1).Picture
                End If
                
                rsTmp.MoveNext
            Next
            
            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, COL_医嘱内容, .Rows - 1, COL_医嘱内容) = 0

            Call .AutoSize(COL_医嘱内容)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
        
        If mbln审核 Then mblnNoSave = True
    Else
        mblnRowChange = False
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        mblnRowChange = True
    End If
    
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AdviceSet复制医嘱(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strIDs As String, _
    Optional ByVal blnHistory As Boolean, Optional ByVal strAlter As String)
'功能：复制指定病人的指定医嘱产生成为新医嘱
'参数：strIDs,strAlter=选择及要切换期效的医嘱ID(组ID)
'      blnHistory=复制包括历史医嘱
'说明：可供外部调用,调用之前处于新增医嘱行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln配方 As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim curDate As Date, blnDoIt As Boolean
    Dim lng开嘱科室ID As Long, str开嘱医生 As String
    Dim lng相关ID As Long, lngCopyRow As Long
    Dim lng序号 As Long, intCount As Integer
    Dim lng医生ID As Long, blnFirst As Boolean
    Dim lngRow As Long, i As Long, j As Long
    
    Dim lng西药房ID As Long, lng成药房ID As Long, lng中药房ID As Long
    Dim str药房IDs As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.相关ID,Nvl(A.婴儿,0) as 婴儿,A.序号,A.医嘱期效," & _
        " A.医嘱状态,A.诊疗类别,A.诊疗项目ID,B.名称,A.标本部位,A.收费细目ID," & _
        " A.开始执行时间,A.医嘱内容,A.医生嘱托,A.单次用量,A.天数,A.总给予量,B.计算单位," & _
        " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,B.计算方式,B.执行频率,B.操作类型," & _
        " B.计价性质,A.执行时间方案,A.执行性质,A.执行科室ID,B.录入限量,C.处方限量,C.处方职务," & _
        " C.毒理分类,C.药品剂型,D.剂量系数,D.住院包装,D.住院单位,D.可否分零,A.申请ID" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,药品特性 C,药品规格 D" & _
        " Where A.诊疗项目ID=B.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
        " And A.收费细目ID=D.药品ID(+) And A.病人ID=[1] And A.主页ID=[2]" & _
        " And Instr([3],','||Nvl(A.相关ID,A.ID)||',')>0" & _
        " Order by 婴儿,序号"
    If blnHistory Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, "," & strIDs & ",")
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        intCount = 0 '已经设置的行数
        mblnRowChange = False
        lngBegin = vsAdvice.Row '开始新增行
        lng序号 = GetCurRow序号(lngBegin) '起始序号
        curDate = zlDatabase.Currentdate
        
        With vsAdvice
            '取开嘱医生和科室
            If mbln护士站 Then
                lngCopyRow = GetPreRow(lngBegin)
                If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngBegin)
                If lngCopyRow <> -1 Then
                    If Val(.TextMatrix(lngCopyRow, COL_状态)) <> 1 Then lngCopyRow = -1
                End If
                If lngCopyRow <> -1 Then
                    str开嘱医生 = .TextMatrix(lngCopyRow, COL_开嘱医生)
                End If
                '缺省为病人的住院医师或病人科室的第一个医生
                If str开嘱医生 = "" Then
                    str开嘱医生 = Get住院医师
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str开嘱医生, lng医生ID)
                Else
                    Call Get开嘱医生(mlng病人科室id, mbln护士站, str开嘱医生, lng医生ID, , , True)
                End If
                lng开嘱科室ID = Get开嘱科室ID(lng医生ID, mlng病人科室id)
            Else
                str开嘱医生 = UserInfo.姓名
                lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
            End If
            
            '复制数据
            .Redraw = flexRDNone
            For i = lngBegin To rsTmp.RecordCount + lngBegin - 1
                If i > lngBegin Then .AddItem "", i

                bln配方 = False
                
                .RowData(i) = -1 * rsTmp!ID
                If Not IsNull(rsTmp!相关ID) Then
                    .TextMatrix(i, COL_相关ID) = -1 * rsTmp!相关ID
                End If
                .TextMatrix(i, COL_序号) = lng序号 + intCount
                
                .TextMatrix(i, COL_EDIT) = 1 '新增
                .Cell(flexcpData, i, COL_EDIT) = CStr(lng病人ID & "," & lng主页ID) '记录相关的复制项目
                .TextMatrix(i, COL_期效) = IIF(Nvl(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .TextMatrix(i, COL_状态) = 1 '新开
                .TextMatrix(i, COL_婴儿) = cbo婴儿.ListIndex
                .TextMatrix(i, COL_类别) = Nvl(rsTmp!诊疗类别, "*")
                .TextMatrix(i, COL_诊疗项目ID) = Nvl(rsTmp!诊疗项目ID)
                .TextMatrix(i, COL_名称) = Nvl(rsTmp!名称)
                .TextMatrix(i, COL_标本部位) = Nvl(rsTmp!标本部位)
                .TextMatrix(i, COL_收费细目ID) = Nvl(rsTmp!收费细目ID)
                .TextMatrix(i, COL_医嘱内容) = Nvl(rsTmp!医嘱内容)
                .TextMatrix(i, COL_医生嘱托) = Nvl(rsTmp!医生嘱托)
                
                .TextMatrix(i, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
                .TextMatrix(i, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                .TextMatrix(i, COL_操作类型) = Nvl(rsTmp!操作类型)
                .TextMatrix(i, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(i, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!处方限量)
                Else
                    .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!录入限量)
                End If
                .TextMatrix(i, COL_处方职务) = Nvl(rsTmp!处方职务)
                .TextMatrix(i, COL_剂量系数) = Nvl(rsTmp!剂量系数)
                .TextMatrix(i, COL_住院包装) = Nvl(rsTmp!住院包装)
                .TextMatrix(i, COL_住院单位) = Nvl(rsTmp!住院单位)
                If Not IsNull(rsTmp!剂量系数) Then
                    .TextMatrix(i, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
                End If
                
                If IsDate(txt开始时间.Text) Then
                    .TextMatrix(i, COL_开始时间) = Format(txt开始时间.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开始时间) = txt开始时间.Text
                End If
                
                '临嘱可选频率的可能被设置为了一次性
                If Nvl(rsTmp!医嘱期效, 0) = 1 And Nvl(rsTmp!执行频率, 0) = 0 _
                    And Nvl(rsTmp!频率次数, 0) = 0 And Nvl(rsTmp!频率间隔, 0) = 0 Then
                    .TextMatrix(i, COL_频率性质) = 1
                Else
                    .TextMatrix(i, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                End If
                .TextMatrix(i, COL_频率) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, COL_频率次数) = Nvl(rsTmp!频率次数)
                .TextMatrix(i, COL_频率间隔) = Nvl(rsTmp!频率间隔)
                .TextMatrix(i, COL_间隔单位) = Nvl(rsTmp!间隔单位)
                .TextMatrix(i, COL_执行时间) = Nvl(rsTmp!执行时间方案)
                .TextMatrix(i, COL_执行性质) = Nvl(rsTmp!执行性质, 0)
                
                '处理执行科室
                If Nvl(rsTmp!诊疗类别) = "Z" Then
                    .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID)
                ElseIf InStr(",0,5,", Nvl(rsTmp!执行性质, 0)) = 0 Then
                    If Nvl(rsTmp!执行科室ID, 0) <> 0 Then
                        If InStr(",5,6,7,", Nvl(rsTmp!诊疗类别, "*")) > 0 Then
                            str药房IDs = Get可用药房IDs(rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), mlng病人科室id, 2)
                            If InStr("," & str药房IDs & ",", "," & rsTmp!执行科室ID & ",") > 0 Then
                                .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_执行性质)) = 4 Then
                            '4-指定科室时才取,其它的固定生成
                            .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                        End If
                    End If
                    If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                        '药品类的整个成套相同
                        If Nvl(rsTmp!诊疗类别) = "5" Then
                            If lng西药房ID = 0 Then
                                lng西药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, Nvl(rsTmp!医嘱期效, 0), 2, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng西药房ID
                        ElseIf Nvl(rsTmp!诊疗类别) = "6" Then
                            If lng成药房ID = 0 Then
                                lng成药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, Nvl(rsTmp!医嘱期效, 0), 2, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng成药房ID
                        ElseIf Nvl(rsTmp!诊疗类别) = "7" Then
                            If lng中药房ID = 0 Then
                                lng中药房ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, Nvl(rsTmp!医嘱期效, 0), 2, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng中药房ID
                        Else
                            .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, Nvl(rsTmp!诊疗类别, "*"), _
                                Nvl(rsTmp!诊疗项目ID, 0), 0, Nvl(rsTmp!执行性质, 0), mlng病人科室id, lng开嘱科室ID, Nvl(rsTmp!医嘱期效, 0), 2)
                        End If
                    End If
                End If
                
                If Nvl(rsTmp!诊疗类别) = "E" Then
                    If Nvl(rsTmp!相关ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_相关ID)) = -1 * rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是成药的给药途径,可能是一并给药的
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = -1 * rsTmp!ID Then
                                    '显示给药途径
                                    .TextMatrix(j, COL_用法) = rsTmp!名称
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是中药配方的用法,即配方显示行
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                            bln配方 = True
                        ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                        End If
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        '当前记录是中药配方煎法行
                        bln配方 = True
                    End If
                ElseIf Nvl(rsTmp!诊疗类别) = "7" Then
                    bln配方 = True
                End If
                
                '单量
                .TextMatrix(i, COL_单量) = FormatEx(Nvl(rsTmp!单次用量), 5)
                If InStr(",5,6,7,", Nvl(rsTmp!诊疗类别, "*")) > 0 Or Nvl(rsTmp!计算方式, 0) <> 3 Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!计算单位)
                End If
                
                '天数
                .TextMatrix(i, COL_天数) = Nvl(rsTmp!天数, 0)
                
                '总量
                If InStr(",5,6,", Nvl(rsTmp!诊疗类别, "*")) > 0 Then
                    '成药临嘱有总量,以零售单位存放,住院单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!住院包装) Then
                        .TextMatrix(i, COL_总量) = FormatEx(rsTmp!总给予量 / rsTmp!住院包装, 5)
                    End If
                    If Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!住院单位)
                    End If
                Else
                    '其它情况有中药和其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then
                        .TextMatrix(i, COL_总量) = rsTmp!总给予量
                    End If
                    If bln配方 Then
                        .TextMatrix(i, COL_总量单位) = "付" '中药配方总量单位为"付"
                        
                        '终止时间:长嘱时根据开始时间、付数、频次,执行时间自动确定
                        If Nvl(rsTmp!医嘱期效, 0) = 0 And rsTmp!诊疗类别 = "E" And IsNull(rsTmp!相关ID) Then
                            .Cell(flexcpData, i, COL_终止时间) = Calc终止时间(i)
                            .TextMatrix(i, COL_终止时间) = Format(.Cell(flexcpData, i, COL_终止时间), "MM-dd HH:mm")
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = -1 * rsTmp!ID And InStr(",7,E,", .TextMatrix(j, COL_类别)) > 0 Then
                                    .Cell(flexcpData, j, COL_终止时间) = .Cell(flexcpData, i, COL_终止时间)
                                    .TextMatrix(j, COL_终止时间) = .TextMatrix(i, COL_终止时间)
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If
                
                .TextMatrix(i, COL_标志) = IIF(mbln补录, 2, 0)
                .TextMatrix(i, COL_开嘱医生) = str开嘱医生
                .TextMatrix(i, COL_开嘱科室ID) = lng开嘱科室ID
                If mbln补录 And IsDate(.Cell(flexcpData, i, COL_开始时间)) Then
                    '补录医嘱时,项目的开嘱时间缺省为开始时间
                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(i, COL_开始时间)
                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, i, COL_开始时间)
                Else
                    .TextMatrix(i, COL_开嘱时间) = Format(curDate, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                End If
                .TextMatrix(i, COL_校对护士) = ""
                
                '紧急标志
                blnFirst = True
                If InStr(",5,6,", rsTmp!诊疗类别) > 0 Then
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Val(.TextMatrix(i, COL_标志)) = 2 Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
                        .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
                    End If
                End If
                
                '毒麻精药品标识:中药配方及组成味中药不处理
                If InStr(",5,6,", Nvl(rsTmp!诊疗类别, "*")) > 0 And Not IsNull(rsTmp!毒理分类) Then
                    If InStr(",麻醉药,毒性药,精神药,", rsTmp!毒理分类) > 0 Then
                        .Cell(flexcpFontBold, i, COL_医嘱内容) = True
                    End If
                End If
                
                lngEnd = i
                intCount = intCount + 1
                
                rsTmp.MoveNext
            Next
            
            '显示/隐藏行
            lngRow = 0
            For i = lngBegin To lngEnd
                blnDoIt = False
                If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                        blnDoIt = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    blnDoIt = True
                End If
                                
                .RowHidden(i) = blnDoIt
                If Not blnDoIt And lngRow = 0 Then lngRow = i
                
                '处理医嘱内容的变化
                If Not .RowHidden(i) Then
                    '复制时开始时间变化
                    txt开始时间.Tag = "1"
                    If AdviceTextChange(i) Then
                        .TextMatrix(i, COL_医嘱内容) = AdviceTextMake(i)
                    End If
                    txt开始时间.Tag = ""
                End If
            Next
            
            '调整医嘱的期效
            If strAlter <> "" Then
                For i = lngBegin To lngEnd
                    lng相关ID = -1 * Val(.TextMatrix(i, COL_相关ID))
                    If lng相关ID = 0 Then lng相关ID = -1 * .RowData(i)
                    If InStr("," & strAlter & ",", "," & lng相关ID & ",") > 0 Then
                        blnDoIt = True
                        If .RowHidden(i) Then blnDoIt = False
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                blnDoIt = False
                            End If
                        End If
                        If blnDoIt Then Call AdviceAlterType(i)
                    End If
                Next
            End If
            
            '产生新的医嘱ID
            For i = lngBegin To lngEnd
                lng相关ID = .RowData(i)
                .RowData(i) = zlDatabase.GetNextId("病人医嘱记录")
                For j = i - 1 To lngBegin Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngEnd
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
            
            '调整受影响行的序号
            Call AdviceSet医嘱序号(lngEnd + 1, intCount)
            
            '图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            
            .Row = lngRow: .Col = COL_医嘱内容
            
            Call .AutoSize(COL_医嘱内容)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
        mblnNoSave = True '标记为未保存
    End If

    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check开始时间(ByVal strStart As String, ByVal strEnd As String, ByVal bln补录 As Boolean, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法
'说明：
'1.开始时间不能小于病人的入院时间
'2.开始时间必须小于终止时间
'3.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
'4.补录的医嘱开始时间不能大于当前时间
    Dim strInDate As String
    
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
        
    strInDate = Format(lvwPati.SelectedItem.ListSubItems(4).Tag, "yyyy-MM-dd HH:mm")
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "医嘱的开始执行时间不能小于病人的入院时间 " & strInDate & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    If IsDate(strEnd) Then
        If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(strEnd, "yyyy-MM-dd HH:mm") Then
            strMsg = "医嘱的开始执行时间必须小于执行终止时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Not bln补录 Then
        If DateDiff("n", CDate(strStart), zlDatabase.Currentdate) > TIME_LIMIT Then
            strMsg = "不是补录医嘱时，开始执行时间不能太早于当前时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Format(strStart, "yyyy-MM-dd HH:mm") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") Then
            strMsg = "补录医嘱的开始执行时间不能大于当前时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check开始时间 = True
End Function

Private Function Check终止时间(ByVal strWrite As String, ByVal strStart As String, ByVal strEnd As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的终止时间是否合法
'说明：
'1.终止时间必须大于开始时间,开嘱时间
'2.##如果有频率,长嘱终止时间至少应该在一个频率周期之后
    If Not IsDate(strEnd) Then
        strMsg = "输入的医嘱执行终止时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If

    If IsDate(strStart) Then
        If Format(strEnd, "yyyy-MM-dd HH:mm") <= Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "医嘱的执行终止时间必须大于开始执行时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If IsDate(strWrite) Then
        If Format(strEnd, "yyyy-MM-dd HH:mm") < Format(strWrite, "yyyy-MM-dd HH:mm") Then
            strMsg = "医嘱的执行终止时间不应小于开嘱时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If

    Check终止时间 = True
End Function

Private Function Check开嘱时间(ByVal strDate As String, ByVal strStart As String, ByVal bln补录 As Boolean, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查开嘱时间是否有效
'说明：
'1.补录医嘱时,开嘱时间不应大于医嘱开始时间和当前时间
'2.不应小于病人入院时间
    Dim strInDate As String
    
    If Not IsDate(strDate) Then
        strMsg = "输入的开嘱时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    If bln补录 And IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") > Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "补录的开嘱时间不应该大于医嘱的开始执行时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
        If Format(strDate, "yyyy-MM-dd HH:mm") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") Then
            strMsg = "补录的开嘱时间不应该大于当前时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strInDate = Format(lvwPati.SelectedItem.ListSubItems(4).Tag, "yyyy-MM-dd HH:mm")
    If Format(strDate, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "开嘱时间不能小于病人的入院时间 " & strInDate & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    Check开嘱时间 = True
End Function

Private Function Check配伍禁忌(ByVal str药品IDs As String) As Boolean
'功能：检查西成药,中成药的配伍禁忌;中药配方不在这里检查
'参数：str药品IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr慎用 As Variant, arr禁用 As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng项目ID As Long, str名称 As String, bln未编辑 As Boolean
    Dim lng组编号 As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
    
    arr慎用 = Array(): arr禁用 = Array()
    
    strSQL = "Select 组编号 From 诊疗互斥项目" & _
        " Where 项目ID IN(" & str药品IDs & ") Group by 组编号 Having Count(*)>1"
    Call zlDatabase.OpenRecordset(rsMain, strSQL, Me.Caption) 'In
    For k = 1 To rsMain.RecordCount
        strSQL = "Select A.组编号,A.类型,A.项目ID,B.名称" & _
            " From 诊疗互斥项目 A,诊疗项目目录 B" & _
            " Where A.项目ID=B.ID And A.组编号=" & rsMain!组编号 & _
            " And A.项目ID IN(" & str药品IDs & ")" & _
            " Order by A.组编号,B.编码"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In:项目ID是索引
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    ReDim Preserve arr慎用(UBound(arr慎用) + 1)
                Else
                    ReDim Preserve arr禁用(UBound(arr禁用) + 1)
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                arr慎用(UBound(arr慎用)) = arr慎用(UBound(arr慎用)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            Else
                arr禁用(UBound(arr禁用)) = arr禁用(UBound(arr禁用)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    
    '先检查禁用部份(禁止继续)
    If UBound(arr禁用) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr禁用) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(Mid(arr禁用(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) '每项目
                lng项目ID = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "，" & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目ID), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & "● " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下药品互相禁用：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查慎用部份(提醒是否继续)
    If UBound(arr慎用) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr慎用) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(Mid(arr慎用(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) '每项目
                lng项目ID = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "，" & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目ID), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & "● " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("在病人医嘱中发现以下药品互相慎用：" & strMsg & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check配伍禁忌 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check诊疗互斥(ByVal str诊疗IDs As String) As Boolean
'功能：检查非药品(成药,中药)的互斥
'参数：str诊疗IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr提醒 As Variant, arr禁止 As Variant, arr停止 As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng项目ID As Long, str名称 As String, bln未编辑 As Boolean
    Dim lng组编号 As Long, lngRow As Long, lngSeekRow As Long
    Dim bln临嘱 As Boolean
    
    On Error GoTo errH
        
    arr提醒 = Array(): arr禁止 = Array(): arr停止 = Array()
    
    strSQL = "Select 组编号 From 诊疗互斥项目" & _
        " Where 项目ID IN(" & str诊疗IDs & ") Group by 组编号 Having Count(*)>1"
    Call zlDatabase.OpenRecordset(rsMain, strSQL, Me.Caption) 'In
    For k = 1 To rsMain.RecordCount
        strSQL = "Select A.组编号,A.组名称,A.类型,A.项目ID,B.名称" & _
            " From 诊疗互斥项目 A,诊疗项目目录 B" & _
            " Where A.项目ID=B.ID And A.组编号=" & rsMain!组编号 & _
            " And A.项目ID IN(" & str诊疗IDs & ")" & _
            " Order by A.组编号,B.编码"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In:项目ID是索引
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    ReDim Preserve arr提醒(UBound(arr提醒) + 1)
                    arr提醒(UBound(arr提醒)) = rsTmp!组名称
                ElseIf rsTmp!类型 = 2 Then
                    ReDim Preserve arr禁止(UBound(arr禁止) + 1)
                    arr禁止(UBound(arr禁止)) = rsTmp!组名称
                Else
                    ReDim Preserve arr停止(UBound(arr停止) + 1)
                    arr停止(UBound(arr停止)) = rsTmp!组名称
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                arr提醒(UBound(arr提醒)) = arr提醒(UBound(arr提醒)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            ElseIf rsTmp!类型 = 2 Then
                arr禁止(UBound(arr禁止)) = arr禁止(UBound(arr禁止)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            Else
                arr停止(UBound(arr停止)) = arr停止(UBound(arr停止)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    
    '先检查禁止继续部份
    If UBound(arr禁止) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr禁止) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(arr禁止(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目ID = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目ID), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then 'bln已校对,COL_状态
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下内容互相排斥：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查自动停止部份,如果在一组中包含临嘱则禁止,否则在校对时自动停止
    If UBound(arr停止) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr停止) '每组
            strTmp = "": bln未编辑 = True: bln临嘱 = False
            arrItems = Split(arr停止(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目ID = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目ID), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False ': Exit Do
                    End If
                    If vsAdvice.TextMatrix(lngRow, COL_期效) = "临嘱" Then bln临嘱 = True
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                If bln临嘱 Then '包含临嘱项时,在程序中就禁止了。
                    strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
                End If
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下内容互相排斥：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查提醒是否继续部份
    If UBound(arr提醒) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr提醒) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(arr提醒(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目ID = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目ID), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("在病人医嘱中发现以下内容互相排斥：" & strMsg & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check诊疗互斥 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lngRow As Long) As String
'功能：检查指定药品行的库存情况
'返回：空=表示通过
    Dim dbl总量 As Double, strMsg As String
    Dim lng执行科室ID As Long, i As Integer
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If GetStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID))) <> 0 Then
                If .TextMatrix(lngRow, COL_期效) = "临嘱" And .TextMatrix(lngRow, COL_库存) <> "" Then
                    '成药临嘱直接检查总量
                    dbl总量 = Val(.TextMatrix(lngRow, COL_总量))
                    If dbl总量 > 0 Then
                        If dbl总量 > Val(.TextMatrix(lngRow, COL_库存)) Then
                            strMsg = """" & .TextMatrix(lngRow, COL_医嘱内容) & """库存提醒：" & _
                                vbCrLf & vbCrLf & Get部门名称(Val(.TextMatrix(lngRow, COL_执行科室ID))) & _
                                "当前可用库存为 " & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & _
                                .TextMatrix(lngRow, COL_住院单位) & "，不足 " & _
                                FormatEx(dbl总量, 5) & .TextMatrix(lngRow, COL_住院单位) & "。"
                        End If
                    End If
                ElseIf Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 And .TextMatrix(lngRow, COL_库存) <> "" Then
                    '成药长嘱(按规格下时)检查一个频率周期的总量
                    If Val(.TextMatrix(lngRow, COL_单量)) <> 0 And .TextMatrix(lngRow, COL_频率) <> "" Then
                        dbl总量 = Calc缺省药品总量(Val(.TextMatrix(lngRow, COL_单量)), 1, _
                            Val(.TextMatrix(lngRow, COL_频率次数)), _
                            Val(.TextMatrix(lngRow, COL_频率间隔)), _
                            .TextMatrix(lngRow, COL_间隔单位), _
                            .TextMatrix(lngRow, COL_执行时间), _
                            Val(.TextMatrix(lngRow, COL_剂量系数)), _
                            Val(.TextMatrix(lngRow, COL_住院包装)), _
                            Val(.TextMatrix(lngRow, COL_可否分零)))
                        If dbl总量 > 0 Then
                            If dbl总量 > Val(.TextMatrix(lngRow, COL_库存)) Then
                                strMsg = """" & .TextMatrix(lngRow, COL_医嘱内容) & """库存提醒：" & _
                                    vbCrLf & vbCrLf & Get部门名称(Val(.TextMatrix(lngRow, COL_执行科室ID))) & _
                                    "当前可用库存为 " & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & _
                                    .TextMatrix(lngRow, COL_住院单位) & "，不足一个频率周期所需用量 " & _
                                    FormatEx(dbl总量, 5) & .TextMatrix(lngRow, COL_住院单位) & "。"
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf RowIn配方行(lngRow) And Val(.TextMatrix(lngRow, COL_总量)) <> 0 Then
            '根据付数计算总量,长嘱临嘱都要输入付数
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" And .TextMatrix(i, COL_库存) <> "" Then
                        '总量=住院包装(单味剂量*付数)
                        '中药药房单位按不可分零处理:每付
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            dbl总量 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                        Else
                            dbl总量 = Val(.TextMatrix(i, COL_总量)) * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                        End If
                        If dbl总量 > Val(.TextMatrix(i, COL_库存)) Then
                            lng执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                            If GetStockCheck(lng执行科室ID) = 0 Then Exit For
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, COL_医嘱内容) & _
                                "：所需总量 " & FormatEx(dbl总量, 5) & .TextMatrix(i, COL_住院单位) & _
                                "，可用库存 " & FormatEx(Val(.TextMatrix(i, COL_库存)), 5) & .TextMatrix(i, COL_住院单位)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            If strMsg <> "" Then
                strMsg = "中药配方库存提醒，" & Get部门名称(lng执行科室ID) & "中以下味药库存不足：" & vbCrLf & strMsg
            End If
        End If
    End With
    CheckStock = strMsg
End Function

Private Function CheckMoney() As Boolean
'功能：费用报警检查
'说明：病区有累计费用报警方式时,只提醒。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng病区ID As Long
    Dim bln医保 As Boolean, cur余额 As Currency
    
    '费用余额
    strSQL = _
        " Select Nvl(预交余额,0)-Nvl(费用余额,0) as 余额 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
        " Union ALL " & _
        " Select Sum(金额) as 余额 From 保险模拟结算 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 is Not Null And A.病人ID=[1] And A.主页ID=[2]"
    strSQL = "Select Sum(余额) as 余额 From (" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then cur余额 = Nvl(rsTmp!余额, 0)
        
    '是否医保
    strSQL = "Select B.编码 From 病案主页 A,医疗付款方式 B" & _
        " Where A.医疗付款方式=B.名称(+) And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    bln医保 = Nvl(rsTmp!编码) = "1"
        
    '报警值:NULL与0当作不同意义处理
    lng病区ID = Get病区ID(mlng病人ID, mlng主页ID)
    strSQL = "Select 报警值 From 记帐报警线 Where 报警方法=1 And 病区ID=[1] And 报警值 is Not NULL And 适用病人=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病区ID, IIF(bln医保, 2, 1))
    If Not rsTmp.EOF Then
        If cur余额 < Nvl(rsTmp!报警值, 0) Then
            If MsgBox("病人当前剩余款 " & FormatEx(cur余额, 2) & " 低于报警值 " & FormatEx(Nvl(rsTmp!报警值, 0), 2) & "，要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    CheckMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAdvice() As Boolean
'功能：检查当前病人(婴儿)的医嘱输入是否合法
'说明：如果有不合法的地方，在本函数中提示及定位
    Dim blnValid As Boolean
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    Dim dbl总量 As Double, strMsg As String
    Dim str药品IDs As String, str诊疗IDs As String
    Dim lngCount As Long, lngRow As Long, i As Long
    Dim blnSkipStock As Boolean, blnSkipTotal As Boolean
    Dim vMsg As VbMsgBoxResult, sng天数 As Single
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            '本次新增或修改药品行的处方职务检查
            If .RowData(i) <> 0 _
                And InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 _
                And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                strMsg = CheckOneDuty(.TextMatrix(i, COL_医嘱内容), .TextMatrix(i, COL_处方职务), .TextMatrix(i, COL_开嘱医生), InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> "")
                If strMsg <> "" Then
                    .Col = COL_医嘱内容
                    If .TextMatrix(i, COL_类别) = "7" Then
                        lngRow = .FindRow(CLng(.TextMatrix(i, COL_相关ID)), i + 1)
                        If lngRow <> -1 Then .Row = lngRow
                    Else
                        .Row = i
                    End If
                    Call .ShowCell(.Row, .Col)
                    MsgBox strMsg, vbInformation, gstrSysName
                    .Refresh
                    If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
                    Exit Function
                End If
            End If
            
            '其它输入合法性检查
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                bln配方行 = RowIn配方行(i)
                bln检验行 = RowIn检验行(i)
                lngRow = i
                If bln配方行 Then '得到配方的第一药品行
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                ElseIf bln检验行 Then '得到检验医嘱行
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                End If
                
                '未校对的医嘱行
                '------------------------------------
                If InStr(",1,2,", .TextMatrix(i, COL_状态)) > 0 Then
                    If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                        lngCount = lngCount + 1
                    
                        '临嘱规格判断
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            If .TextMatrix(i, COL_期效) = "临嘱" And Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then
                                strMsg = "没有对应的药品规格信息。"
                                .Col = COL_医嘱内容: Exit For
                            End If
                        End If
                        
                        '必须录入单量
                        If .TextMatrix(i, COL_期效) = "长嘱" Then
                            '长嘱：成药或计时,计量项目需要录入
                            If InStr(",1,2,", Val(.TextMatrix(i, COL_计算方式))) > 0 Or InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                                If Not IsNumeric(.TextMatrix(i, COL_单量)) Or Val(.TextMatrix(i, COL_单量)) <= 0 Then
                                    strMsg = "没有录入正确的单次用量。"
                                    .Col = COL_单量: Exit For
                                End If
                            End If
                        Else
                            '临嘱:成药或可选择频率的计时,计量项目可以录入(也可不录)
                            If Val(.TextMatrix(i, COL_频率性质)) = 0 And InStr(",1,2,", Val(.TextMatrix(i, COL_计算方式))) > 0 Then
                                If .TextMatrix(i, COL_单量) <> "" Then
                                    If Not IsNumeric(.TextMatrix(i, COL_单量)) Or Val(.TextMatrix(i, COL_单量)) <= 0 Then
                                        strMsg = "没有录入正确的单次用量。"
                                        .Col = COL_单量: Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                        '必须录入总量:配方,临嘱(药品或其它)
                        If .TextMatrix(i, COL_期效) = "临嘱" Or bln配方行 Then
                            If Not IsNumeric(.TextMatrix(i, COL_总量)) Or Val(.TextMatrix(i, COL_总量)) <= 0 Then
                                If bln配方行 Then
                                    strMsg = "没有录入正确的中药配方付数。"
                                ElseIf InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                                    strMsg = "没有录入正确的药品总给予量。"
                                Else
                                    strMsg = "没有录入正确的总量。"
                                End If
                                .Col = COL_总量: Exit For
                            End If
                        End If
                                            
                        '必须录入频率:临嘱也要检查,用于指导使用
                        If Val(.TextMatrix(i, COL_频率性质)) = 0 Or bln配方行 Then
                            If .TextMatrix(i, COL_频率) = "" Then
                                strMsg = "没有确定执行频率。"
                                .Col = COL_频率: Exit For
                            End If
                            
                            '执行时间判断:可选频率的必须输入(对临嘱将来可能允许不录入,要注意发送等地方的处理)
                            If .TextMatrix(i, COL_执行时间) = "" Then
                                If Not bln检验行 Or .TextMatrix(i, COL_期效) = "长嘱" Then '检验组合显示行的采集方法为可选频率,但检验项目为一次性
                                    strMsg = "没有录入执行时间方案。"
                                    .Col = COL_执行时间: Exit For
                                End If
                            End If
                        End If
                        
                        '必须录入执行科室:非叮嘱和院外执行时(配方以药品行进行判断)
                        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                            If .TextMatrix(lngRow, COL_类别) = "Z" And Val(.TextMatrix(lngRow, COL_操作类型)) = 3 Then
                                strMsg = "没有确定转科医嘱的转入科室。"
                                .Col = COL_执行科室ID: Exit For
                            ElseIf .TextMatrix(lngRow, COL_类别) = "Z" And Val(.TextMatrix(lngRow, COL_操作类型)) = 7 Then
                                strMsg = "没有确定会诊医嘱的会诊科室。"
                                .Col = COL_执行科室ID: Exit For
                            ElseIf InStr(",0,5,", .TextMatrix(lngRow, COL_执行性质)) = 0 Then
                                strMsg = "没有确定执行科室。"
                                .Col = COL_执行科室ID: Exit For
                            End If
                        End If
                        If lngRow <> i And Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            If InStr(",0,5,", .TextMatrix(i, COL_执行性质)) = 0 Then
                                strMsg = "没有确定执行科室。"
                                .Col = COL_执行科室ID: Exit For
                            End If
                        End If
                        
                        
                        '终止时间判断
                        If .TextMatrix(i, COL_期效) = "长嘱" And Not bln配方行 And .Cell(flexcpData, i, COL_终止时间) <> "" Then
                            If Not Check终止时间(.Cell(flexcpData, i, COL_开嘱时间), .Cell(flexcpData, i, COL_开始时间), .Cell(flexcpData, i, COL_终止时间), False, strMsg) Then
                                .Col = COL_终止时间: Exit For
                            End If
                        End If
                    End If
                    
                    '必须录入开嘱医生:护士使用时
                    If mbln护士站 And .TextMatrix(i, COL_开嘱医生) = "" Then
                        strMsg = "没有确定开嘱医生。"
                        .Col = COL_开嘱医生: Exit For
                    End If
                    
                    '开嘱时间判断
                    If Not Check开嘱时间(.Cell(flexcpData, i, COL_开嘱时间), .Cell(flexcpData, i, COL_开始时间), Val(.TextMatrix(i, COL_标志)) = 2, False, strMsg) Then
                        .Col = COL_开嘱时间: Exit For
                    End If
                End If
                
                '本次新增或修改的行
                '---------------------------------------------------
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    '开始时间判断:只对新增的医嘱或补录的医嘱进行判断,因为否则是不准修改开始时间的(不好判断被修改的非补录医嘱开始时间的相对有效性)
                    If Val(.TextMatrix(i, COL_标志)) = 2 Or .TextMatrix(i, COL_EDIT) = "1" Then
                        If Not Check开始时间(.Cell(flexcpData, i, COL_开始时间), .Cell(flexcpData, i, COL_终止时间), Val(.TextMatrix(i, COL_标志)) = 2, False, strMsg) Then
                            .Col = COL_开始时间: Exit For
                        End If
                    End If
                    
                    If Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                        If .TextMatrix(i, COL_医嘱内容) = "" Then
                            strMsg = "没有录入医嘱内容。"
                            .Col = COL_用法: Exit For
                        End If
                    Else
                        '给药途径，中药用法，采集方法设置检查
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_诊疗项目ID)) = 0 Then
                                strMsg = "没有设置对应的给药途径。"
                                .Col = COL_用法: Exit For
                            End If
                        End If
                        If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                            If .RowData(i) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                If InStr(",7,E,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                                    strMsg = "中药配方没有设置对应的用法。"
                                ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                                    strMsg = "没有设置对应的标本采集方法。"
                                End If
                                .Col = COL_用法: Exit For
                            End If
                        End If
                        
                        '最少总量检查:至少要满足一个频次周期的用量
                        If ((InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And .TextMatrix(i, COL_期效) = "临嘱") Or bln配方行) Then
                            If Not blnSkipTotal And .TextMatrix(i, COL_频率) <> "" Then
                                strMsg = ""
                                If bln配方行 Then '判断
                                    dbl总量 = Calc缺省药品总量(1, 1, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                                    If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                        strMsg = .TextMatrix(i, COL_医嘱内容) & vbCrLf & vbCrLf & _
                                            "在按""" & .TextMatrix(i, COL_频率) & """执行时,至少需要 " & dbl总量 & "付。"
                                    End If
                                ElseIf Val(.TextMatrix(i, COL_剂量系数)) <> 0 Then
                                    If Val(.TextMatrix(i, COL_频率性质)) = 1 Then '临嘱成药可能为一次性
                                        dbl总量 = Calc缺省药品总量(Val(.TextMatrix(i, COL_单量)), 1, 1, 1, "天", "", Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_住院包装)), Val(.TextMatrix(i, COL_可否分零)))
                                    Else
                                        sng天数 = Val(.TextMatrix(i, COL_天数))
                                        If sng天数 = 0 Then sng天数 = 1
                                        dbl总量 = Calc缺省药品总量(Val(.TextMatrix(i, COL_单量)), sng天数, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_住院包装)), Val(.TextMatrix(i, COL_可否分零)))
                                    End If
                                    If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                        strMsg = .TextMatrix(i, COL_医嘱内容) & vbCrLf & vbCrLf & _
                                            "在按每次 " & .TextMatrix(i, COL_单量) & .TextMatrix(i, COL_单量单位) & "," & _
                                            .TextMatrix(i, COL_频率) & IIF(mbln天数, ",用药 " & sng天数 & " 天", "") & _
                                            "执行时,至少需要 " & dbl总量 & .TextMatrix(i, COL_总量单位) & "。"
                                    End If
                                End If
                                If strMsg <> "" Then '提示
                                    .Row = i: .Col = COL_总量: Call .ShowCell(.Row, .Col)
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^要继续吗？", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If txt总量.Enabled And txt总量.Visible Then txt总量.SetFocus
                                        Exit Function
                                    ElseIf vMsg = vbIgnore Then
                                        blnSkipTotal = True
                                    End If
                                End If
                            End If
                        End If
                        
                        '药品库存检查:只提醒,所以也只对本次编辑的才判断
                        If (InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Or bln配方行) And Not blnSkipStock Then
                            strMsg = CheckStock(i)
                            If strMsg <> "" Then
                                .Row = i: .Col = COL_医嘱内容: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^要继续吗？", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipStock = True
                                End If
                            End If
                        End If
                        
                        '执行时间合法性检查
                        If .TextMatrix(i, COL_执行时间) <> "" And .TextMatrix(i, COL_频率) <> "" Then
                            blnValid = ExeTimeValid(.TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                            If Not blnValid Then
                                If .TextMatrix(i, COL_间隔单位) = "周" Then
                                    strMsg = COL_按周执行
                                ElseIf .TextMatrix(i, COL_间隔单位) = "天" Then
                                    strMsg = COL_按天执行
                                ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                                    strMsg = COL_按时执行
                                End If
                                strMsg = "录入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strMsg
                                .Col = COL_执行时间: Exit For
                            End If
                        End If
                        
                        '医保对码检查:以一组医保第一可见行为准
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) = 0 _
                            Or Val(.TextMatrix(i - 1, COL_相关ID)) <> Val(.TextMatrix(i, COL_相关ID)) Then
                            If gint医保对码 = 2 Then mbln提醒对码 = True
                            strMsg = CheckAdviceInsure(i)
                            If strMsg <> "" Then
                                .Row = i: .Col = COL_医嘱内容: Call .ShowCell(.Row, .Col)
                                If gint医保对码 = 1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                    If vMsg = vbIgnore Then mbln提醒对码 = False
                                ElseIf gint医保对码 = 2 Then
                                    MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                                strMsg = "" '防止后面再作处理
                            End If
                        End If
                    End If
                End If
                                
                '互斥数据收集:在所有有效医嘱中,因为可能已校对的与未校对的互斥
                If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        '用于药品配伍禁忌检查:不分期效
                        str药品IDs = str药品IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                    ElseIf Not bln配方行 Then
                        '不管检查组合与手术附加内部之间及内部与其它项目之间
                        str诊疗IDs = str诊疗IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '中间退出的错误提示
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln配方行 Then
                    strMsg = "该中药配方" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, COL_医嘱内容) & """" & strMsg
                End If
                MsgBox strMsg, vbInformation, gstrSysName
                .Refresh
            End If
            If .Col = COL_医嘱内容 Then
                If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
            Else
                Call vsAdvice_KeyPress(13)
            End If
            Exit Function
        End If
        
        '检查药品配伍禁忌
        If str药品IDs <> "" Then
            If Not Check配伍禁忌(Mid(str药品IDs, 2)) Then Exit Function
        End If
        '检查诊疗项目互斥
        If str诊疗IDs <> "" Then
            If Not Check诊疗互斥(Mid(str诊疗IDs, 2)) Then Exit Function
        End If
    End With
    
    '费用报警:有未校对医嘱时
    If lngCount > 0 Then
        If Not CheckMoney Then Exit Function
    End If
    
    CheckAdvice = True
End Function

Private Function SeekNextControl() As Boolean
'功能：定位到下一个焦点的控件上,并根据情况决定是否自动新增一行医嘱
'返回：如果通过SetFocus强制定位的,则返回True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange '强制更新医嘱内容
                                vsAdvice.Row = i
                                Call zlCommFun.PressKey(vbKeyTab)
                                blnDo = vsAdvice.RowData(i) <> 0 '无内容则再跳入编辑
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            Call tbr_ButtonClick(tbr.Buttons("增加"))
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        blnDo = True: objNext.SetFocus
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'功能：获取输入医嘱过程中，回车光标应跳过的控件
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        '一并给药中的药品输入时应跳过的内容
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                '给药途径,附加执行
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_诊疗项目ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt用法.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_执行科室ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo附加执行.Name
                        End If
                    End If
                End If
                '频率
                If .TextMatrix(lngRow, COL_频率) <> "" Then strSkip = strSkip & ";" & Me.txt频率.Name
                '执行时间
                If .TextMatrix(lngRow, COL_执行时间) <> "" Then strSkip = strSkip & ";" & Me.cbo执行时间.Name
                '终止时间:因为一般为空，所以后面输入时固定跳过
                strSkip = strSkip & ";" & Me.txt终止时间.Name
                '开嘱医生
                If .TextMatrix(lngRow, COL_开嘱医生) <> "" Then strSkip = strSkip & ";" & Me.cbo医生.Name
                '开嘱时间
                If .TextMatrix(lngRow, COL_开嘱时间) <> "" Then strSkip = strSkip & ";" & Me.txt开嘱时间.Name
            End If
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'功能：当医嘱卡片输入内容变化时，判断医嘱内容文本是否应该重新组织
    Dim str类别 As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" Then '中药配方或一组检验
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If lngRow <> -1 Then str类别 = .TextMatrix(lngRow, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '检查内容变动
        If blnDefine Then '公共字段部份或可以公共处理的部份
            If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" And InStr(strText, "[开始时间]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cbo医生嘱托.Tag <> "" And InStr(strText, "[医生嘱托]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                If InStr(strText, "[中文频率]") > 0 Or InStr(strText, "[英文频率]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cbo执行时间.Tag <> "" And InStr(strText, "[执行时间]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" And InStr(strText, "[单量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[总量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str类别 '不同的类别检查
        Case "5", "6" '中西成药
            If Not blnDefine Then
                
            Else
                '[输入名][通用名][商品名][英文名][规格][产地]是输入或修改整个药品时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[给药途径]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '中药配方
            If Not blnDefine Then
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmd频率.Tag <> "" And txt频率.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[配方组成][煎法]是输入或修改整个配方时变化
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[付数]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[用法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '检验
            If Not blnDefine Then
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[检验项目][检验标本]是输入或修改整个项目时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[采集方法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '检查
            If Not blnDefine Then
                
            Else
                '[检查项目][检查部位]是输入或修改整个项目时变化
            End If
        Case "F" '手术
            If Not blnDefine Then
                If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[主要手术][附加手术][麻醉方法]是输入或修改整个项目时变化
            End If
        Case Else '其他
            If Not blnDefine Then
                
            Else
                '[诊疗项目]是输入或修改整个项目时变化
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'功能：获取医嘱内容文本
'参数：lngRow=已有医嘱数据的可见行
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean, str类别 As String
    Dim strText As String, strSQL As String
    Dim strField As String, int频率范围 As Integer
    Dim i As Long, k As Long
    
    Dim str中药 As String, str煎法 As String
    Dim str麻醉 As String, str附术 As String
    Dim str检验 As String, str标本 As String
    Dim str部位 As String
    
    On Error GoTo errH
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" Then '中药配方或一组检验
            k = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If k <> -1 Then str类别 = .TextMatrix(k, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '用于按定义公式计算失败，重新按缺省规则进行组织
        strText = ""
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '产生医嘱内容
        Select Case str类别
        Case "C" '检验-------------------------------------------------------------
            str检验 = "": str标本 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    str检验 = .TextMatrix(i, COL_医嘱内容) & "," & str检验
                    str标本 = .TextMatrix(i, COL_标本部位)
                Else
                    Exit For
                End If
            Next
            If str检验 = "" Then '老的方式
                str检验 = .TextMatrix(lngRow, COL_名称)
            Else
                str检验 = Left(str检验, Len(str检验) - 1)
            End If
            
            If Not blnDefine Then
                strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
            Else
                If InStr(strText, "[检验项目]") > 0 Then
                    strField = str检验
                    strText = Replace(strText, "[检验项目]", """" & strField & """")
                End If
                If InStr(strText, "[检验标本]") > 0 Then
                    strField = str标本
                    strText = Replace(strText, "[检验标本]", """" & strField & """")
                End If
                If InStr(strText, "[采集方法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[采集方法]", """" & strField & """")
                End If
            End If
        Case "D" '检查-------------------------------------------------------------
            str部位 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_标本部位) <> "" Then
                        str部位 = str部位 & "," & .TextMatrix(i, COL_标本部位)
                    End If
                Else
                    Exit For
                End If
            Next
            str部位 = Mid(str部位, 2) '检查组合项目的部位
            If str部位 = "" Then '独立检查项目的部位
                str部位 = .TextMatrix(lngRow, COL_标本部位)
            End If
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称) & IIF(str部位 <> "", "(" & str部位 & ")", "")
            Else
                If InStr(strText, "[检查项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[检查项目]", """" & strField & """")
                End If
                If InStr(strText, "[检查部位]") > 0 Then
                    strField = str部位
                    strText = Replace(strText, "[检查部位]", """" & strField & """")
                End If
            End If
        Case "F" '手术-------------------------------------------------------------
            str麻醉 = "": str附术 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        str麻醉 = .TextMatrix(i, COL_医嘱内容)
                    Else
                        str附术 = str附术 & "," & .TextMatrix(i, COL_医嘱内容)
                    End If
                Else
                    Exit For
                End If
            Next
            str附术 = Mid(str附术, 2)
            
            If Not blnDefine Then
                strText = Format(.Cell(flexcpData, lngRow, COL_开始时间), "MM月dd日HH:mm")
                If str麻醉 <> "" Then
                    strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_名称)
                If str附术 <> "" Then
                    strText = strText & " 及 " & str附术
                End If
            Else
                If InStr(strText, "[主要手术]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[主要手术]", """" & strField & """")
                End If
                If InStr(strText, "[附加手术]") > 0 Then
                    strField = str附术
                    strText = Replace(strText, "[附加手术]", """" & strField & """")
                End If
                If InStr(strText, "[麻醉方法]") > 0 Then
                    strField = str麻醉
                    strText = Replace(strText, "[麻醉方法]", """" & strField & """")
                End If
            End If
        Case "8" '中药配方---------------------------------------------------------
            str中药 = "": str煎法 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        str中药 = RTrim(.TextMatrix(i, COL_医嘱内容) & _
                            " " & .TextMatrix(i, COL_单量) & .TextMatrix(i, COL_单量单位) & _
                            " " & .TextMatrix(i, COL_医生嘱托)) & "," & str中药
                    ElseIf .TextMatrix(i, COL_类别) = "E" Then
                        str煎法 = .TextMatrix(i, COL_医嘱内容)
                    End If
                Else
                    Exit For
                End If
            Next
            If str中药 <> "" Then
                str中药 = Mid(str中药, 1, Len(str中药) - 1)
            End If
            If Not blnDefine Then
                '数字后加了空格在文本框中会自动换行
                strText = "中药" & .TextMatrix(lngRow, COL_总量) & "付," & _
                    .TextMatrix(lngRow, COL_频率) & "," & str煎法 & "," & _
                    .TextMatrix(lngRow, COL_用法) & ":" & str中药
            Else
                If InStr(strText, "[付数]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_总量)
                    strText = Replace(strText, "[付数]", """" & strField & """")
                End If
                If InStr(strText, "[配方组成]") > 0 Then
                    strField = str中药
                    strText = Replace(strText, "[配方组成]", """" & strField & """")
                End If
                If InStr(strText, "[用法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[用法]", """" & strField & """")
                End If
                If InStr(strText, "[煎法]") > 0 Then
                    strField = str煎法
                    strText = Replace(strText, "[煎法]", """" & strField & """")
                End If
            End If
        Case "5", "6" '西成药，中成药---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                '性质:0-正名,1-英文名,3-商品名
                strSQL = "Select Nvl(B.名称,A.名称) as 名称,A.规格,A.产地,B.性质" & _
                    " From 收费项目目录 A,收费项目别名 B Where A.ID=B.收费细目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
            ElseIf blnDefine Then
                '性质:0-正名,1-英文名
                strSQL = "Select Nvl(B.名称,A.名称) as 名称,Null as 规格,Null as 产地,B.性质" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B Where A.ID=B.诊疗项目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_标本部位)
                If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                    If strText = "" Then
                        If gbln商品名 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!名称
                    End If
                    If Not IsNull(rsTmp!产地) Then
                        strText = strText & "(" & rsTmp!产地 & ")"
                    End If
                    If Not IsNull(rsTmp!规格) Then
                        strText = strText & " " & rsTmp!规格
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_名称)
                    End If
                End If
            Else
                If InStr(strText, "[输入名]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_标本部位)
                    If strField = "" Then
                        If gbln商品名 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[输入名]", """" & strField & """")
                End If
                If InStr(strText, "[通用名]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!名称
                    strText = Replace(strText, "[通用名]", """" & strField & """")
                End If
                If InStr(strText, "[商品名]") > 0 Then
                    rsTmp.Filter = "性质=3"
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = rsTmp!名称
                    strText = Replace(strText, "[商品名]", """" & strField & """")
                End If
                If InStr(strText, "[英文名]") > 0 Then
                    rsTmp.Filter = "性质=2"
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = rsTmp!名称
                    strText = Replace(strText, "[英文名]", """" & strField & """")
                End If
                If InStr(strText, "[规格]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!规格)
                    strText = Replace(strText, "[规格]", """" & strField & """")
                End If
                If InStr(strText, "[产地]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!产地)
                    strText = Replace(strText, "[产地]", """" & strField & """")
                End If
                If InStr(strText, "[给药途径]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[给药途径]", """" & strField & """")
                End If
            End If
        Case Else '其它所有类别-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称)
            Else
                If InStr(strText, "[诊疗项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[诊疗项目]", """" & strField & """")
                End If
            End If
            '术后医嘱特殊显示
            If .TextMatrix(lngRow, COL_类别) = "Z" And Val(.TextMatrix(lngRow, COL_操作类型)) = 4 Then
                strText = "━━━" & strText & "━━━"
            End If
        End Select
        
        '公共字段或可以公共处理的字段-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[开始时间]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_开始时间)
                strText = Replace(strText, "[开始时间]", """" & strField & """")
            End If
            If InStr(strText, "[医生嘱托]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_医生嘱托)
                strText = Replace(strText, "[医生嘱托]", """" & strField & """")
            End If
            If InStr(strText, "[中文频率]") > 0 Then
                strField = .TextMatrix(lngRow, COL_频率)
                strText = Replace(strText, "[中文频率]", """" & strField & """")
            End If
            If InStr(strText, "[英文频率]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_频率) <> "" Then
                    int频率范围 = Get频率范围(lngRow)
                    strSQL = "Select 英文名称 From 诊疗频率项目 Where 名称=[1] And 适用范围=[2]"
                    Set rsTmp = New ADODB.Recordset '清除Filter
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, COL_频率), int频率范围)
                    If Not rsTmp.EOF Then strField = Nvl(rsTmp!英文名称)
                End If
                strText = Replace(strText, "[英文频率]", """" & strField & """")
            End If
            If InStr(strText, "[单量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_单量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_单量) & .TextMatrix(lngRow, COL_单量单位)
                End If
                strText = Replace(strText, "[单量]", """" & strField & """")
            End If
            If InStr(strText, "[总量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_总量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_总量) & .TextMatrix(lngRow, COL_总量单位)
                End If
                strText = Replace(strText, "[总量]", """" & strField & """")
            End If
            If InStr(strText, "[执行时间]") > 0 Then
                strField = .TextMatrix(lngRow, COL_执行时间)
                strText = Replace(strText, "[执行时间]", """" & strField & """")
            End If
        End If
                
        '计算医嘱内容
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                Err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceCopyCurr(ByVal lngRow As Long)
'功能：根据指定行的医嘱复制产生新的医嘱
    Dim lngBegin As Long, lngEnd As Long
    Dim lngStart As Long, lngOver As Long
    Dim lng序号 As Long, lng相关ID As Long
    Dim lngShow As Long, curDate As Date
    Dim blnFirst As Boolean, i As Long, j As Long
    
    Dim lng开嘱科室ID As Long, str开嘱医生 As String
    
    curDate = zlDatabase.Currentdate
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    
    With vsAdvice
        '取开嘱医生和科室
        If mbln护士站 Then
            str开嘱医生 = .TextMatrix(lngRow, COL_开嘱医生)
            lng开嘱科室ID = .TextMatrix(lngRow, COL_开嘱科室ID)
        Else
            str开嘱医生 = UserInfo.姓名
            lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng病人科室id)
        End If
    
        .Redraw = flexRDNone
        For i = lngBegin To lngEnd
            If .RowData(i) <> 0 Then
                '添加新行
                If .RowData(.Rows - 1) <> 0 Then .AddItem "", .Rows
                lngRow = .Rows - 1
                If lngStart = 0 Then lngStart = lngRow
                lngOver = lngRow
                
                '医嘱起始序号
                If lng序号 = 0 Then
                    lng序号 = GetCurRow序号(lngRow)
                End If
                
                '复制内容
                '--------------------------------------------------------------------------------
                For j = 0 To .Cols - 1
                    .TextMatrix(lngRow, j) = .TextMatrix(i, j)
                    .Cell(flexcpData, lngRow, j) = .Cell(flexcpData, i, j)
                    .Cell(flexcpFontBold, lngRow, j) = .Cell(flexcpFontBold, i, j) '毒麻药品标识
                Next
                .RowData(lngRow) = .RowData(i)
                .RowHidden(lngRow) = .RowHidden(i)
                .RowHeight(lngRow) = .RowHeight(i)
                
                '调整内容
                '--------------------------------------------------------------------------------
                .RowData(lngRow) = -1 * .RowData(lngRow)
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    .TextMatrix(lngRow, COL_相关ID) = -1 * Val(.TextMatrix(lngRow, COL_相关ID))
                End If
                
                .TextMatrix(lngRow, COL_状态) = 1 '新开
                .TextMatrix(lngRow, COL_EDIT) = 1 '新增
                .TextMatrix(lngRow, COL_序号) = lng序号
                lng序号 = lng序号 + 1
                
                .TextMatrix(lngRow, COL_开始时间) = Format(curDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                '配方长嘱终止时间:根据开始时间、付数、频次,执行时间自动确定
                If .TextMatrix(lngRow, COL_期效) = "长嘱" And .TextMatrix(lngRow, COL_类别) = "E" _
                    And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                    If RowIn配方行(lngRow) Then
                        .Cell(flexcpData, lngRow, COL_终止时间) = Calc终止时间(lngRow)
                        .TextMatrix(lngRow, COL_终止时间) = Format(.Cell(flexcpData, lngRow, COL_终止时间), "MM-dd HH:mm")
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngRow) And InStr(",7,E,", .TextMatrix(j, COL_类别)) > 0 Then
                                .TextMatrix(j, COL_终止时间) = .TextMatrix(lngRow, COL_终止时间)
                                .Cell(flexcpData, j, COL_终止时间) = .Cell(flexcpData, lngRow, COL_终止时间)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                .TextMatrix(lngRow, COL_开嘱时间) = Format(curDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                .TextMatrix(lngRow, COL_开嘱医生) = str开嘱医生
                .TextMatrix(lngRow, COL_开嘱科室ID) = lng开嘱科室ID
                .TextMatrix(lngRow, COL_校对护士) = ""
                
                '紧急标志
                .TextMatrix(lngRow, COL_标志) = IIF(mbln补录, 2, 0)
                blnFirst = True
                If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                    If Val(.TextMatrix(lngRow, COL_相关ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                        Set .Cell(flexcpPicture, lngRow, COL_F标志) = imgFlag.ListImages("补录").Picture
                        .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
                    End If
                End If
            End If
        Next
        
        '产生新的医嘱ID
        '--------------------------------------------------------------------------------
        For i = lngStart To lngOver
            lng相关ID = .RowData(i)
            .RowData(i) = zlDatabase.GetNextId("病人医嘱记录")
            For j = i - 1 To lngStart Step -1
                If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            For j = i + 1 To lngOver
                If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            
            '定位到第一个可见行
            If Not .RowHidden(i) And lngShow = 0 Then lngShow = i
            
            '处理医嘱内容的变化
            If Not .RowHidden(i) Then
                '复制时开始时间变化
                txt开始时间.Tag = "1"
                If AdviceTextChange(i) Then
                    .TextMatrix(i, COL_医嘱内容) = AdviceTextMake(i)
                End If
                txt开始时间.Tag = ""
            End If
        Next
        
        '图标对齐:设置为中对齐,不然擦边框时可能有问题
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
        
        '复制后刷新
        '--------------------------------------------------------------------------------
        .Row = lngShow
        Call .ShowCell(.Row, .Col): .TopRow = .Row
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
        
        mblnNoSave = True '标记为未保存
    End With
End Sub

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱是否可以切换期效
'参数：lngRow=可见的医嘱行
'说明：允许切换期效的条件：
'   1.成长嘱：执行频率=0(可选频率),2(持续性)
'   2.成临嘱：执行频率=0(可选频率),1(一次性);药品必须指定了规格
    Dim rsMore As New ADODB.Recordset
    Dim strSQL As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由输入的可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn配方行(lngRow) Then
            '中药配方固定可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn检验行(lngRow) Then
            '检验以检验行为准判断
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", "临嘱", "长嘱")
        
        '以原始频率为准判断:因为可选择频率的可能已缺成一次性
        strSQL = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
        
        If strType = "长嘱" Then
            If InStr(",0,2,", Nvl(rsMore!执行频率, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", Nvl(rsMore!执行频率, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceAlterType(ByVal lngRow As Long)
'功能：在尽量保持数据的情况下，切换指定行医嘱的期效(长期<->临时)
'参数：lngRow=可见的医嘱行
'说明：执行该函数时应保证已用CanAlterType函数进行了判断
    Dim rsMore As New ADODB.Recordset
    Dim strType As String, strSQL As String
    Dim int频率性质 As Integer, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng用法ID As Long, blnToNormal As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim lngCopyRow As Long, i As Long
    
    With vsAdvice
        '最终要转换为的期效
        strType = IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", "临嘱", "长嘱")
        
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) <> 0 Then
            '取上一或下一有效行,某些内容缺省与该行相同
            lngCopyRow = GetPreRow(lngRow)
            If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
            '获取一组医嘱的操作行范围
            Call GetRowScope(lngRow, lngBegin, lngEnd)
        End If
        
        '针对不同类别的医嘱进行转换-----------------------------------------
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由录入的医嘱直接处理
            .TextMatrix(lngRow, COL_期效) = strType
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2
                .TextMatrix(lngRow, COL_状态) = 1
            End If
        ElseIf RowIn配方行(lngRow) Then '中药配方
            '药品长嘱不能为离院带药
            If strType = "长嘱" And .TextMatrix(lngEnd, COL_类别) = "E" _
                And .RowData(lngEnd) = Val(.TextMatrix(lngBegin, COL_相关ID)) Then
                If Val(.TextMatrix(lngBegin, COL_执行性质)) <> 5 And Val(.TextMatrix(lngEnd, COL_执行性质)) = 5 Then
                    lng用法ID = Val(.TextMatrix(lngEnd, COL_诊疗项目ID))
                    blnToNormal = True '表示给药执行应恢复成正常值
                End If
            End If
            
            For i = lngBegin To lngEnd
                '期效值
                .TextMatrix(i, COL_期效) = strType

                '终止时间
                If strType = "长嘱" Then
                    If i = lngBegin Then
                        .Cell(flexcpData, i, COL_终止时间) = Calc终止时间(i)
                        .TextMatrix(i, COL_终止时间) = Format(.Cell(flexcpData, i, COL_终止时间), "MM-dd HH:mm")
                    Else
                        .TextMatrix(i, COL_终止时间) = .TextMatrix(lngBegin, COL_终止时间)
                        .Cell(flexcpData, i, COL_终止时间) = .Cell(flexcpData, lngBegin, COL_终止时间)
                    End If
                Else
                    .TextMatrix(i, COL_终止时间) = ""
                    .Cell(flexcpData, i, COL_终止时间) = Empty
                End If
                
                '执行性质:药品长嘱不能为"离院带药"
                If i = lngEnd And blnToNormal Then
                    strSQL = "Select 执行科室 From 诊疗项目目录 Where ID=[1]"
                    Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng用法ID)
                    
                    .TextMatrix(i, COL_执行性质) = Nvl(rsMore!执行科室, 0)
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                        .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", lng用法ID, 0, _
                            Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), IIF(strType = "长嘱", 0, 1))
                    Else
                        .TextMatrix(i, COL_执行科室ID) = 0
                    End If
                End If
                
                If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    .TextMatrix(i, COL_EDIT) = 2
                    .TextMatrix(i, COL_状态) = 1
                End If
            Next
        Else '其它诊断项目,包括药品,检查(组合),手术(组合)；检验组合因代码处理部份相同,因此一起处理
            '获取给药途径ID
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 _
                And .TextMatrix(lngEnd, COL_类别) = "E" And .RowData(lngEnd) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lng用法ID = Val(.TextMatrix(lngEnd, COL_诊疗项目ID))
                
                '药品长嘱不能为离院带药
                If strType = "长嘱" Then
                    If Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 And Val(.TextMatrix(lngEnd, COL_执行性质)) = 5 Then
                        blnToNormal = True '表示给药执行应恢复成正常值
                    End If
                End If
            End If
            
            '------------------------------------------------------------------------------------------------------
            '同时处理一组医嘱的相关行
            For i = lngBegin To lngEnd
                '期效值
                .TextMatrix(i, COL_期效) = strType
                
                '获取当前项目的附加信息
                If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And i = lngBegin Then
                    '第一药品行才取这些信息
                    strSQL = "Select 项目ID,频次,疗程 From 诊疗用法用量 Where Nvl(性质,0)>0 And 项目ID=[1] And 用法ID=[2]"
                    strSQL = "Select A.执行科室,A.执行频率,A.计算方式,A.计算单位,B.频次,B.疗程" & _
                        " From 诊疗项目目录 A,(" & strSQL & ") B Where A.ID=B.项目ID(+) And A.ID=[1]"
                Else
                    strSQL = "Select 执行科室,执行频率,计算方式,计算单位,Null as 频次,Null as 疗程 From 诊疗项目目录 Where ID=[1]"
                End If
                Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_诊疗项目ID)), lng用法ID)
                If Not rsMore.EOF Then '给药途径没有指定的情况
                    '总量(单位)
                    If strType = "临嘱" Then
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            '中、西成药临嘱的总量单位就是住院单位
                            .TextMatrix(i, COL_总量单位) = .TextMatrix(i, COL_住院单位)
                        Else
                            '其它临嘱要输入总量
                            .TextMatrix(i, COL_总量单位) = Nvl(rsMore!计算单位)
                            
                            '如果为一次性或计次临嘱缺省总量为1
                            If i = lngBegin Then
                                If Nvl(rsMore!执行频率, 0) = 1 Or Nvl(rsMore!计算方式, 0) = 3 Then
                                    .TextMatrix(i, COL_总量) = 1
                                End If
                            ElseIf Not (lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))) Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngBegin, COL_总量)
                            End If
                        End If
                    Else
                        .TextMatrix(i, COL_总量) = ""
                        .TextMatrix(i, COL_总量单位) = ""
                    End If
                
                    '频率性质,执行频率,执行时间
                    If i = lngBegin Then '以第一行为准
                        int频率性质 = Val(.TextMatrix(i, COL_频率性质))
                        If strType = "临嘱" And Nvl(rsMore!执行频率, 0) = 0 And mbln一次性 Then
                            .TextMatrix(i, COL_频率性质) = 1 '住院可选择频率的临嘱缺省为一次性
                        Else
                            .TextMatrix(i, COL_频率性质) = Nvl(rsMore!执行频率, 0)
                        End If
        
                        '执行频率:当适用范围有所变化时
                        If Val(.TextMatrix(i, COL_频率性质)) <> int频率性质 Then
                            '标记为重新取
                            .TextMatrix(i, COL_频率) = ""
                            .TextMatrix(i, COL_执行时间) = ""
                            
                            '药品设置的缺省频率优先
                            If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 _
                                And Not IsNull(rsMore!频次) And Val(.TextMatrix(i, COL_频率性质)) <> 1 Then
                                Call Get频率信息_编码(rsMore!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                                .TextMatrix(i, COL_频率) = str频率
                                .TextMatrix(i, COL_频率次数) = int频率次数
                                .TextMatrix(i, COL_频率间隔) = int频率间隔
                                .TextMatrix(i, COL_间隔单位) = str间隔单位
                            End If
                            '缺省与上一新增行相同
                            If .TextMatrix(i, COL_频率) = "" And lngCopyRow <> -1 Then
                                If .TextMatrix(i, COL_期效) = .TextMatrix(lngCopyRow, COL_期效) _
                                    And Val(.TextMatrix(i, COL_频率性质)) = Val(.TextMatrix(lngCopyRow, COL_频率性质)) Then
                                    If Val(.TextMatrix(lngCopyRow, COL_EDIT)) = 1 And .TextMatrix(lngCopyRow, COL_频率) <> "" _
                                        And Not (.TextMatrix(i, COL_类别) = "7" And Not RowIn配方行(lngCopyRow)) _
                                        And Not (.TextMatrix(i, COL_类别) <> "7" And RowIn配方行(lngCopyRow)) Then
                                        .TextMatrix(i, COL_频率) = .TextMatrix(lngCopyRow, COL_频率)
                                        .TextMatrix(i, COL_频率次数) = .TextMatrix(lngCopyRow, COL_频率次数)
                                        .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngCopyRow, COL_频率间隔)
                                        .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngCopyRow, COL_间隔单位)
                                    End If
                                End If
                            End If
                            '或取缺省频率
                            If .TextMatrix(i, COL_频率) = "" Then
                                Call Get缺省频率(Get频率范围(i), str频率, int频率次数, int频率间隔, str间隔单位)
                                .TextMatrix(i, COL_频率) = str频率
                                .TextMatrix(i, COL_频率次数) = int频率次数
                                .TextMatrix(i, COL_频率间隔) = int频率间隔
                                .TextMatrix(i, COL_间隔单位) = str间隔单位
                            End If
                            
                            '执行时间:可选频率的项目
                            If Val(.TextMatrix(i, COL_频率性质)) = 0 Then
                                If lngCopyRow <> -1 Then '与上一行相同
                                    If .TextMatrix(i, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                                        .TextMatrix(i, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                                    End If
                                End If
                                If .TextMatrix(i, COL_执行时间) = "" Then  '缺省时间方案
                                    .TextMatrix(i, COL_执行时间) = Get缺省时间(1, .TextMatrix(i, COL_频率), lng用法ID)
                                End If
                            End If
                        End If
                    Else
                        .TextMatrix(i, COL_频率) = .TextMatrix(lngBegin, COL_频率)
                        .TextMatrix(i, COL_频率次数) = .TextMatrix(lngBegin, COL_频率次数)
                        .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngBegin, COL_频率间隔)
                        .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngBegin, COL_间隔单位)
                        .TextMatrix(i, COL_频率性质) = .TextMatrix(lngBegin, COL_频率性质)
                        .TextMatrix(i, COL_执行时间) = .TextMatrix(lngBegin, COL_执行时间)
                    End If
                    
                    '终止时间:长嘱才有
                    If strType = "临嘱" Then
                        .TextMatrix(i, COL_终止时间) = ""
                        .Cell(flexcpData, i, COL_终止时间) = Empty
                    End If
                    
                    '药品临嘱天数和总量
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And strType = "临嘱" Then
                        '确定临嘱用药天数：
                        '1.最少为一个频率周期天数
                        '2-有疗程则为疗程天数(应大于一个频率周期天数)
                        If i = lngBegin Then '以第一行为准
                            sng天数 = Val(.TextMatrix(i, COL_天数)) '如果以前有则保持
                            If sng天数 = 0 Then sng天数 = msng天数
                            If mbln天数 Then
                                If .TextMatrix(i, COL_间隔单位) = "周" Then
                                    If 7 > sng天数 Then sng天数 = 7
                                ElseIf .TextMatrix(i, COL_间隔单位) = "天" Then
                                    If Val(.TextMatrix(i, COL_频率间隔)) > sng天数 Then
                                        sng天数 = Val(.TextMatrix(i, COL_频率间隔))
                                    End If
                                ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                                    If Val(.TextMatrix(i, COL_频率间隔)) \ 24 > sng天数 Then
                                        sng天数 = Val(.TextMatrix(i, COL_频率间隔)) \ 24
                                    End If
                                End If
                            End If
                            If Nvl(rsMore!疗程, 1) > sng天数 Then sng天数 = Nvl(rsMore!疗程, 1)
                            If sng天数 = 0 Then sng天数 = 1
                        End If
                        
                        '天数
                        If mbln天数 And Val(.TextMatrix(i, COL_频率性质)) <> 1 Then
                            .TextMatrix(i, COL_天数) = sng天数
                        End If
                        
                        '总量
                        If .TextMatrix(i, COL_频率) <> "" And Val(.TextMatrix(i, COL_单量)) <> 0 _
                            And Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_住院包装)) <> 0 Then
                            If Val(.TextMatrix(i, COL_频率性质)) = 1 Then '临嘱药品可能缺省为一次性
                                '仅按疗程算改为按最少用药天数算
                                .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                        Val(.TextMatrix(i, COL_单量)), 1, 1, 1, "天", "", Val(.TextMatrix(i, COL_剂量系数)), _
                                        Val(.TextMatrix(i, COL_住院包装)), Val(.TextMatrix(i, COL_可否分零))), 5)
                            Else
                                '仅按疗程算改为按最少用药天数算
                                .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                        Val(.TextMatrix(i, COL_单量)), sng天数, Val(.TextMatrix(i, COL_频率次数)), _
                                        Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位), _
                                        .TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_剂量系数)), _
                                        Val(.TextMatrix(i, COL_住院包装)), Val(.TextMatrix(i, COL_可否分零))), 5)
                            End If
                        End If
                    End If
                    
                    '执行性质:药品长嘱不能为"离院带药"
                    If i = lngEnd And blnToNormal Then
                        .TextMatrix(i, COL_执行性质) = Nvl(rsMore!执行科室, 0)
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                            .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, "E", lng用法ID, 0, _
                                Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), IIF(strType = "长嘱", 0, 1))
                        Else
                            .TextMatrix(i, COL_执行科室ID) = 0
                        End If
                    End If
                    
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_状态) = 1
                    End If
                End If
            Next
        End If
    End With
    
    Call ReSetColor(lngRow)
    mblnNoSave = True '标记为未保存
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'功能：调用Pass系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        1-保存自动审查,2-提交自动审查,3-手工调用审查
'        6-单药警告,12-用药研究,22-病生状态/过敏史管理(编辑)
'      lngRow=当前药品医嘱的行号，lngCmd=0,6时需要
'返回：本次审核返回的最高级别警示值,为-1,-2,-3表示没有进行审查
'      检测PASS菜单时，返回>=0表示可以弹出菜单
'说明：用药审查：涉及当天下的临嘱(包括已执行)，和未停止的长嘱
'      用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String, str频率 As String
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long
    
    lngMaxWarn = -1
    AdviceCheckWarn = lngMaxWarn
    
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
    If mlng病人ID <> mlngPassPati Then
        strSQL = _
            " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
            " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
            " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
            " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
            " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    
        Call PassSetPatientInfo(mlng病人ID, mlng主页ID, rsTmp!姓名, Nvl(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
            rsTmp!科室码 & "/" & rsTmp!科室名, IIF(Not IsNull(rsTmp!医生名), Nvl(rsTmp!医生码) & "/" & Nvl(rsTmp!医生名), ""), _
            IIF(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        mlngPassPati = mlng病人ID
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With vsAdvice
            If .RowData(lngRow) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COL_类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                '取药品名称
                str药品 = .TextMatrix(lngRow, COL_医嘱内容)
                If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
                If InStr(str药品, "(") > 0 Then str药品 = Left(str药品, InStr(str药品, "(") - 1)
                '取药品给药途径
                str用法 = ""
                k = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                If k <> -1 Then str用法 = .TextMatrix(k, COL_医嘱内容)
                
                '传入查询药品信息
                Call PassSetQueryDrug(.TextMatrix(lngRow, COL_收费细目ID), str药品, .TextMatrix(lngRow, COL_单量单位), str用法)
                    
                '设置菜单可用状态
                Call SetPassMenuState
                
                AdviceCheckWarn = 1 '表示可以弹出菜单
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    
    '过敏史/病生状态编辑
    '-------------------------------------------------------------
    If lngCmd = 22 Then
        'lngCmd=21-只读,22-非强制编辑,23-强制编辑
        If PassDoCommand(lngCmd) = 2 Then
            '如果返回值为2表示"过敏史/病生状态编辑"管理发生变化，需要重新自动审查
            lngCmd = 2 '转为自动调用审查,继续执行
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    
    '传入病人医嘱信息
    '-------------------------------------------------------------
    With vsAdvice
        If lngCmd = 6 Then
            Call PassSetWarnDrug(.RowData(lngRow)) '单药警告(已警告的医嘱唯一码)
        Else
            '用药审核或用药研究
            lngCount = 0
            curDate = zlDatabase.Currentdate
            str药品 = "": str用法 = "": str频率 = ""
            For i = .FixedRows To .Rows - 1
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex And Val(.TextMatrix(i, COL_收费细目ID)) <> 0
                blnDo = blnDo And (lngCmd = 12 Or .TextMatrix(i, COL_期效) = "长嘱" _
                    Or .TextMatrix(i, COL_期效) = "临嘱" And Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                If blnDo Then
                    '取药品名称
                    str药品 = .TextMatrix(i, COL_医嘱内容)
                    If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
                    If InStr(str药品, "(") > 0 Then str药品 = Left(str药品, InStr(str药品, "(") - 1)
                    
                    '取药品给药途径
                    If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then str用法 = "" '一并给药不重复取
                    If str用法 = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, COL_相关ID)), i + 1)
                        If k <> -1 Then str用法 = .TextMatrix(k, COL_医嘱内容)
                    End If
                    
                    '取用药频率(次/天),都为整数四舍五入
                    If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then str频率 = "" '一并给药不重复取
                    If str频率 = "" Then
                        If .TextMatrix(i, COL_间隔单位) = "天" Then
                            str频率 = .TextMatrix(i, COL_频率次数) & "/" & .TextMatrix(i, COL_频率间隔)
                        ElseIf .TextMatrix(i, COL_间隔单位) = "周" Then
                            str频率 = .TextMatrix(i, COL_频率次数) & "/7"
                        ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                            If Val(.TextMatrix(i, COL_频率间隔)) <= 24 Then
                                str频率 = Format(24 / Val(.TextMatrix(i, COL_频率间隔)) * Val(.TextMatrix(i, COL_频率次数)), "0") & "/1"
                            Else
                                str频率 = Val(.TextMatrix(i, COL_频率次数)) & "/" & Format(Val(.TextMatrix(i, COL_频率间隔)) / 24, "0")
                            End If
                        End If
                    End If
                    
                    '传入医嘱信息
                    Call PassSetRecipeInfo(.RowData(i), .TextMatrix(i, COL_收费细目ID), str药品, _
                        .TextMatrix(i, COL_单量), .TextMatrix(i, COL_单量单位), str频率, _
                        Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd"), _
                        Format(.Cell(flexcpData, i, COL_终止时间), "yyyy-MM-dd"), _
                        str用法, .TextMatrix(i, COL_相关ID), IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1), _
                        UserInfo.编号 & "/" & UserInfo.姓名)
                    lngCount = lngCount + 1
                End If
            Next
            '无可审查的药品
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With
    
    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    
    '获取医嘱审查结果,并填写警示灯
    '-------------------------------------------------------------
    If lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3 Then
        '返回值顺：0-蓝灯,1-黄灯,2-红灯,3-黑灯,4-橙灯
        '警示级顺：0-蓝灯,1-黄灯,4-橙灯,2-红灯,3-黑灯(因为PASS升级的原因)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex And Val(.TextMatrix(i, COL_收费细目ID)) <> 0
                blnDo = blnDo And (.TextMatrix(i, COL_期效) = "长嘱" _
                    Or .TextMatrix(i, COL_期效) = "临嘱" And Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                If blnDo Then
                    k = PassGetWarn(.RowData(i))
                    strOld = .Cell(flexcpData, i, COL_警示)

                    '设置警示灯
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, COL_警示) = CStr(k)
                        Set .Cell(flexcpPicture, i, COL_警示) = imgPass.ListImages(k + 1).Picture
                    Else
                        .Cell(flexcpData, i, COL_警示) = ""
                        Set .Cell(flexcpPicture, i, COL_警示) = Nothing
                    End If
                    
                    '标记审查结果变化,以备更新数据库
                    If CStr(.Cell(flexcpData, i, COL_警示)) <> strOld Then
                        .Cell(flexcpData, i, COL_序号) = 1
                        mblnNoSave = True '标记为未保存
                    End If
                                        
                    '记录最高级别警示值
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    '返回审查结果
    AdviceCheckWarn = lngMaxWarn
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
    If Button = 2 And gblnPass And InStr(mstrPrivs, "合理用药监测") > 0 And Not mbln护士站 Then
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
    '系统设置
    mnuPassItem(14).Enabled = PassGetState("SYS-SET") = 1
    '用药研究
    mnuPassItem(16).Enabled = PassGetState("DISQUISITION") = 1
    '警告:有警示值(不为空),且大于0-蓝灯
    mnuPassItem(18).Enabled = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_警示)) > 0
    '审查
    'mnuPassItem(19).Enabled = PassGetState("") = 1
    
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
    Case 14 '系统设置
        Call PassDoCommand(11)
    Case 16 '用药研究
        Call AdviceCheckWarn(12)
    Case 18 '警告
        Call AdviceCheckWarn(6, vsAdvice.Row)
    Case 19 '审查
        Call AdviceCheckWarn(3)
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
