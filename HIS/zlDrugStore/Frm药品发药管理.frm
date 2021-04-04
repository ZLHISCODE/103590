VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm药品发药管理 
   Caption         =   "药品处方发药"
   ClientHeight    =   7560
   ClientLeft      =   3465
   ClientTop       =   1845
   ClientWidth     =   11400
   DrawMode        =   12  'Nop
   Icon            =   "Frm药品发药管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11400
   Begin VB.CheckBox Chk显示退药待发单据 
      Appearance      =   0  'Flat
      Caption         =   "显示退药待发单据"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7440
      TabIndex        =   45
      Top             =   720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox cbo病区 
      Height          =   300
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf列表 
      Height          =   5415
      Left            =   30
      TabIndex        =   14
      Top             =   990
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicBackGroud 
      Height          =   5415
      Left            =   2280
      ScaleHeight     =   5355
      ScaleWidth      =   7245
      TabIndex        =   17
      Top             =   980
      Width           =   7305
      Begin VB.ComboBox cbo配药人 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   48
         Text            =   "cbo配药人"
         Top             =   4860
         Width           =   1215
      End
      Begin VB.PictureBox picRecipeColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   460
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   1095
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         Begin VB.Label lblRecipeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "普通"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Top             =   105
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdAlley 
         Caption         =   "过敏史/病生状态"
         Height          =   350
         Left            =   90
         TabIndex        =   37
         Top             =   630
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "发药(&S)"
         Height          =   350
         Left            =   5910
         TabIndex        =   13
         ToolTipText     =   "热键：F2"
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CheckBox Chk全退 
         Appearance      =   0  'Flat
         Caption         =   "全退(&A)"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6120
         TabIndex        =   30
         Top             =   4560
         Width           =   1005
      End
      Begin VB.TextBox Txt收费员 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         TabIndex        =   31
         Top             =   4860
         Width           =   885
      End
      Begin VB.ComboBox Txt开单医生 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4860
         Width           =   1335
      End
      Begin ZL9BillEdit.BillEdit Bill处方明细 
         Height          =   2655
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4683
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.ComboBox TxtNo 
         Height          =   315
         ItemData        =   "Frm药品发药管理.frx":030A
         Left            =   4860
         List            =   "Frm药品发药管理.frx":030C
         TabIndex        =   2
         Top             =   720
         Width           =   2325
      End
      Begin VB.PictureBox PicState 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6510
         ScaleHeight     =   375
         ScaleWidth      =   675
         TabIndex        =   21
         Top             =   90
         Visible         =   0   'False
         Width           =   675
         Begin VB.Label LblState 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "作废"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.TextBox Txt科室 
         Enabled         =   0   'False
         Height          =   300
         Left            =   570
         TabIndex        =   4
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox Txt床号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6390
         TabIndex        =   12
         Top             =   1080
         Width           =   795
      End
      Begin VB.TextBox Txt住院号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4860
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txt年龄 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         TabIndex        =   8
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox Txt性别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2310
         TabIndex        =   6
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox txt原始付数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4440
         Width           =   525
      End
      Begin VB.TextBox txt中药煎法 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4440
         Width           =   2955
      End
      Begin VB.Label Lbl收费员 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收费员"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4200
         TabIndex        =   32
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label Lbl配药人 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "配药人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   18
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label Lbl科室 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4200
         TabIndex        =   1
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Lbl标题 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   120
         TabIndex        =   20
         Top             =   90
         Width           =   7140
      End
      Begin VB.Label Lbl床号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5970
         TabIndex        =   11
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Lbl住院号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标识号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4260
         TabIndex        =   9
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Lbl年龄 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   7
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Lbl性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1860
         TabIndex        =   5
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Lbl开单医生 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开单医生"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lbl中药煎法 
         AutoSize        =   -1  'True
         Caption         =   "中药煎法"
         Height          =   180
         Left            =   1980
         TabIndex        =   35
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl原始付数 
         AutoSize        =   -1  'True
         Caption         =   "原始付数"
         Height          =   180
         Left            =   150
         TabIndex        =   34
         Top             =   4500
         Width           =   720
      End
   End
   Begin VB.Frame fraFind 
      Height          =   480
      Left            =   120
      TabIndex        =   40
      Top             =   6600
      Width           =   3975
      Begin VB.CommandButton cmdIC 
         Caption         =   "读卡"
         Height          =   300
         Left            =   3410
         TabIndex        =   43
         Top             =   135
         Width           =   495
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   690
         TabIndex        =   0
         Top             =   150
         Width           =   2325
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   3480
         Picture         =   "Frm药品发药管理.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "处方定位(F2)"
         Top             =   135
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgFilter 
         Height          =   240
         Left            =   3075
         Picture         =   "Frm药品发药管理.frx":0458
         Top             =   150
         Width           =   240
      End
      Begin VB.Label lblFind 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡↓"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   20
         TabIndex        =   42
         ToolTipText     =   "病人定位(F3)"
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.PictureBox PicToolbar 
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   9510
      ScaleHeight     =   720
      ScaleWidth      =   1830
      TabIndex        =   38
      Top             =   15
      Width           =   1830
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "划价人员"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   39
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.CheckBox Chk清单 
      Appearance      =   0  'Flat
      Caption         =   "显示所有过程单据"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   720
      Width           =   1845
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   5100
      Top             =   150
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5520
      Top             =   150
   End
   Begin VB.PictureBox PicCloseConsignment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   215
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   7440
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   6840
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1164
      BandCount       =   1
      _CBWidth        =   11400
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "Tbar1"
      MinHeight1      =   600
      Width1          =   3000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar1 
         Height          =   600
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Find"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "Cancel"
               Object.ToolTipText     =   "取消发药"
               Object.Tag             =   "取消"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "划价"
               Key             =   "Charge"
               Object.ToolTipText     =   "划价"
               Object.Tag             =   "划价"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发料"
               Key             =   "Stuff"
               Object.ToolTipText     =   "发料"
               Object.Tag             =   "发料"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   7
            EndProperty
         EndProperty
         Begin VB.Timer TimePrint 
            Enabled         =   0   'False
            Left            =   4560
            Top             =   120
         End
         Begin MSComctlLib.ImageList imgPass 
            Left            =   8100
            Top             =   30
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
                  Picture         =   "Frm药品发药管理.frx":6CAA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm药品发药管理.frx":6F64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm药品发药管理.frx":721E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm药品发药管理.frx":74D8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm药品发药管理.frx":7792
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   7200
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsfPrint 
      Height          =   2985
      Left            =   390
      TabIndex        =   25
      Top             =   2550
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5265
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   345
      Left            =   0
      TabIndex        =   28
      Top             =   630
      Width           =   3950
      _ExtentX        =   6959
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "待配药(&1)"
      TabPicture(0)   =   "Frm药品发药管理.frx":7A4C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "已配药(&2)"
      TabPicture(1)   =   "Frm药品发药管理.frx":7A68
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "待发药(&3)"
      TabPicture(2)   =   "Frm药品发药管理.frx":7A84
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "退药(&4)"
      TabPicture(3)   =   "Frm药品发药管理.frx":7AA0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin VB.Image img病区 
      Height          =   240
      Left            =   4320
      Picture         =   "Frm药品发药管理.frx":7ABC
      ToolTipText     =   "选择病区"
      Top             =   6720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgLeftRight_S 
      Height          =   5385
      Left            =   3720
      MousePointer    =   9  'Size W E
      Top             =   990
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MnuFileSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu MnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu MnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu MnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileBillprint 
         Caption         =   "打印配药单(&B)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MnuFileRePrint 
         Caption         =   "打印处方签(&D)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileReport 
         Caption         =   "打印发药清单(&W)"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "打印退药通知单(&R)"
      End
      Begin VB.Menu mnuFileLable 
         Caption         =   "打印药品标签(&L)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileBack 
         Caption         =   "打印退费单据(T)"
      End
      Begin VB.Menu MnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFilePara 
         Caption         =   "参数设置(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu MnuEditDosage 
         Caption         =   "配药模式(&D)"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuEditAbolish 
         Caption         =   "取消模式(&A)"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuEditConsignment 
         Caption         =   "发药模式(&C)"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditHandback 
         Caption         =   "退药模式(&H)"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditBatch 
         Caption         =   "批量发药(&B)"
      End
      Begin VB.Menu MnuEditSendOther 
         Caption         =   "发其它药房的处方(&F)"
      End
      Begin VB.Menu MnuEditHandbackBatch 
         Caption         =   "退其它药房的处方(&T)"
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "按票据号发药(&I)"
      End
      Begin VB.Menu mnuEditBillRestore 
         Caption         =   "按票据号退药(&R)"
      End
      Begin VB.Menu mnuline9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlag 
         Caption         =   "停止发药标记(&S)"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "取消发药(&Q)"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuCharge 
         Caption         =   "门诊划价(&M)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuStuff 
         Caption         =   "卫材发料(@W)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "切换配药人(&E)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu MnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu MnuViewToolS 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu sdfsdfsd 
            Caption         =   "-"
         End
         Begin VB.Menu MnuViewToolT 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MnuViewState 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&O)"
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "小字体(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "中字体(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "大字体(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLocate 
         Caption         =   "定位方式(&S)"
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "就诊卡(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "单据号(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "门诊号(&3)"
            Index           =   2
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "姓名(&4)"
            Index           =   3
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "身份证(&5)"
            Index           =   4
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "IC卡(&6)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuView4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "过滤(&F)"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuHelpWeb 
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu MnuHelpWebM 
            Caption         =   "发送反馈(&E)..."
         End
      End
      Begin VB.Menu MnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
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
Attribute VB_Name = "Frm药品发药管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--注册表相关变量--
Private intFont As Integer                              '字体
Private IntShowCol As Integer                           '在处方明细中是否显示付数(0)

Private mintShowBill收费 As Integer                     '显示收费处方范围
Private mintShowBill记帐 As Integer                     '显示记帐处方范围
Private mstrShowBill As String                          '查询SQL子串
Private mstrShowSendedBill As String                    '查询SQL子串：仅对于已发药单据

Private IntAutoPrint As Integer                         '发药后打印处方单(1)
Private int校验配药人 As Integer                        '配药时是否校验配药人
Private int校验发药人 As Integer                        '发药时是否校验发药人
Private int药品名称 As Integer                          '药品名称显示格式：0-编码与名称;1-仅编码;2-仅名称

Private intPrint As Integer                             '不打印未配药单据(0)
Private mbln记帐单 As Boolean                           '打印配药单时是否包含记帐单
Private strPrintWindow As String                        '打印未配药单据为3时有效
Private mbln就诊卡 As Boolean                           '是否自动定位到就诊卡
Private mlng待发单据 As Long                            '是否显示退药待发单据
Private mint离院带药 As Integer
Private mint自动配药 As Integer                         '是否使用自动配药功能：0-不使用；1-使用
Private mint自动配药时限 As Integer                     '超过该时限就需要验证配药人：默认为始终不验证配药人
Private mint输入模式 As Integer

'0-不打印未配药单据
'1-打印本部门所有未配药单据
'2-打印本窗口所有未配药单据
'3-选择打印(发药窗口)
Private mlngRefresh As Long                             '刷新间隔(0)
Private mlngPrintInterval As Long                       '打印配药单间隔(0)
Private mIntPrintDelay As Integer                       '延迟打印(60)
Private mIntPrintHandbackNO As Integer                  '打印退费单据号(0)
Private mintPrintDrugLable  As Integer                  '打印药品标签
Private lng药房ID As Long                               '药房(设置本机所对应的药房)
Private Str配药人 As String                             '设置配药人
Private mstr自动配药人 As String                        '用于自动配药功能中
Private Str窗口 As String                               '发药窗口(设置本机所对应的发药窗口)
Private IntTimes As Integer                             '已延迟
Private intVerify As Integer                            '是否需要校验处方
Private BlnEnterCell As Boolean                         '是否允许激法ENTERCELL()事件
Private str序号 As String                               '保存当前待发药单据明细序号集

Private mstrOracleMoneyForamt As String                 'ORACLE中金额格式
Private mstrVBMoneyForamt As String                     'VB中金额格式

'--系统参数--
Private StrFindStyle As String                          '匹配串
Private IntCheckStock As Integer                        '检查库存
Private IntSendAfterDosage As Integer                   '是否必须经过配药过程(0)
Private mblnStarPass As Boolean                         '启用合理用药(PASS)
Public Int允许未审核处方发药 As Integer                 '未审核是否允许发药
Public mint允许未收费处方发药 As Integer                '未收费是否允许发药
Private bln医嘱作废 As Boolean                          '是否允许未作废医嘱退药
Private int金额保留位数 As Integer                      '费用金额保留位数
Private int审核划价单 As Integer                        '执行后自动审核划价单
Private mint自动销帐 As Integer
Private mbln报警包含划价费用 As Boolean
Private mbln显示大小单位 As Boolean

'--常规变量--
Private BlnStartUp As Boolean                           '启动成功
Private BlnFirstStart As Boolean                        '第一次启动
Private LngSendRow As Long                              '待发
Private BlnInRefresh As Boolean                         '是否处于刷新状态
Private BlnInOper As Boolean                            '是否输入NO号
Private mstrFilter  As String                           '界面过滤条件
Private mrsBatchSend As ADODB.Recordset                 '用于批量发药
Private mblnFilterRefresh   As Boolean
Private mbln允许取消发药 As Boolean                     '是否允许对未退药品进行取消发药处理
Private mstr操作员 As String
Private mstr配药人 As String
Private mdate上次校验时间 As Date
Private mblnIsFirst As Boolean                          '未校验
Private mblnAuto As Boolean

Private mstrStartDate As String
Private mstrEndDate As String

Private mstrPrintRecipe As String                       '用于发药后打印，记录单据号、单据类型：单据号1,单据类型1|单据号2,单据类型2......

Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private mblnStateTimeRefresh As Boolean
Private mblnStateTimePrint As Boolean

'--本程序使用记录集--
Private RecPhysic As New ADODB.Recordset                '药品记录
Private RecPart As New ADODB.Recordset                  '部门表

Private mrsPASS As New ADODB.Recordset                  'PASS用数据集

'--变量--
Private BlnAllowClick As Boolean                        '允许执行Click事件
Private strUnit As String                               '单位名称
Private str单位串 As String                             '单位串
Private mInt单据 As Integer                             '单据类型  0-门诊及住院所有单据 8-门诊划价及门诊记帐 9-住院记帐
Private IntBillStyle As Integer                         '单据
Private mstrNo As String                                'NO
Private mint门诊标志 As Integer                         '当前单据的门诊标志 1-门诊;2-住院
Private mint记录性质 As Integer                         '当前单据的记录性质 1-收费记录;2-记帐记录
Private StrLastNo As String                             '上次选择或输入的NO
Private IntLastBill As Integer                          '上次选择或输入的单据
Private strLastData As String                           '上次选择或输入的单据的填制或审核日期
Private mintLastSequence As Integer                     '处方明细列表的上次选择的序号
Private StrFind_1 As String                             '未配药处方查找串
Private StrFind_2 As String                             '已配药处方查找串
Private StrFind_3 As String                             '未发药处方查找串
Private StrFind_4 As String                             '已发药处方查找串
Private StrDate As String                               '当前系数日期
Private strBill As String                               '记录上张已发药处方号及单据类型
Private mblnAllBack As Boolean                          '是否全退
Private mblnCard As Boolean                             '是否刷就诊卡
    
Private mblnIs中药处方 As Boolean                        '当前处方是否为中药处方
Private mstr毒麻类提示 As String
Private mstr价格失效提示 As String
Private mbln显示重量 As Boolean

'PASS
Private mstr单量单位 As String
Private mlng病人ID As Long
Private mlngPassPati As Long
Private mlng主页ID As Long
Private mstr挂号单 As String

Private Const mlng紫色 As Long = &HC000C0
'--排序方式--
Private strOrder_1 As String                            '未配药处方排序串
Private strOrder_2 As String                            '已配药处方排序串
Private strOrder_3 As String                            '未发药处方排序串
Private strOrder_4 As String                            '已发药处方排序串

'--返回参数--
Private mstrSourceDep As String                         '来源科室串
Public BlnSetParaSuccess As Boolean                     '设置成功与否
Public Int模式 As Integer
Private mlngMode As Long
Private mstrPrivs As String                              '权限串
Private strChargePrivs As String                        '门诊划价权限串
Private strStuffPrivs As String                         '卫材发放管理权限串
Private BlnRefresh As Boolean
Private mbln发病区处方 As Boolean

Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private Enum 处方列名
    颜色 = 0
    处方类型 = 1
    选择 = 2
    标志 = 3
    类型 = 4
    单据 = 5
    收费 = 6
    配药人 = 7
    NO = 8
    姓名 = 9
    金额 = 10
    日期 = 11
    可操作 = 12
    说明 = 13
    就诊卡号 = 14
    门诊号 = 15
    身份证 = 16
    IC卡 = 17
    病人ID = 18

    '非退药才有
    未审核 = 19
    实收金额 = 20
        
    '退药才有
    门诊标志 = 19
    记录性质 = 20
        
    '发药、退药列数
    发药列数 = 21
    退药列数 = 21
End Enum

Private Enum 列名
    审查结果 = 0
    顺序号 = 1       '表单中的序号
    药品名称 = 2
    其它名 = 3
    英文名 = 4
    序号 = 5
    规格 = 6
    批号 = 7
    Id = 8
    药品ID = 9
    批次 = 10
    单位 = 11
    单价 = 12
    付数 = 13
    数量 = 14
    金额 = 15
    重量 = 16
    单量 = 17
    用法 = 18
    频次 = 19
    医生嘱托 = 20
    费别 = 21
    库存数 = 22
    货位 = 23
    已退数 = 24
    准退数 = 25
    准退数大 = 26
    准退数小 = 27
    退药数 = 28
    退药数大 = 29
    单位大 = 30
    退药数小 = 31
    单位小 = 32
    分批 = 33
    新批号 = 34
    新效期 = 35
    新产地 = 36
    备注 = 37
    医嘱id = 38
    实际数量 = 39
    包装 = 40
    列数 = 41
End Enum

Private Type Type_SQLCondition
    date开始日期 As Date
    date结束日期 As Date
    str开始NO As String
    str结束NO As String
    str姓名 As String
    str就诊卡 As String
    str标识号 As String
    lng科室ID As Long
    str填制人 As String
    str审核人 As String
    lng药品ID As Long
    str当前NO As String
    str门诊号 As String
    str身份证 As String
    strIC卡 As String
    str医保号 As String
End Type

Private SQLCondition As Type_SQLCondition

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object

Private Enum FindType
    就诊卡 = 0
    单据号 = 1
    门诊号 = 2
    姓名 = 3
    身份证 = 4
    IC卡 = 5
End Enum

Private Const cstLocate As Integer = 0
Private Const cstFilter As Integer = 1

'处方类型：普通、儿科、急诊、精二、精一、麻醉
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

'处方类型名称，按顺序，用;分隔
Private Const mconstrRecipeType = "普通;儿科;急诊;精二;精一;麻醉"

'默认处方颜色：普通－白色；急诊－淡黄色；儿科－淡绿色；麻醉、精一－淡红色；精二－白色
Private Const mconlng普通 = &HFFFFFF
Private Const mconlng儿科 = &HC0FFC0
Private Const mconlng急诊 = &HC0FFFF
Private Const mconlng精二 = &HFFFFFF
Private Const mconlng精一 = &HC0C0FF
Private Const mconlng麻醉 = &HC0C0FF

'用户定义的处方颜色，从注册表取的字符串，用;分隔
Private mstrUserRecipeColor As String

Private Function CheckBatchRecipe() As Boolean
    Dim n As Integer
    Dim rsTemp As ADODB.Recordset
    Dim BlnFirst As Boolean
    Dim lngRow As Long, lng药品ID As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    Dim blnBatchSend As Boolean
    Dim i As Integer
    
    On Error GoTo ErrHand
       
    '检查病人费用余额
    If Not CheckSendBillMoney(True) Then Exit Function
    
    For n = 1 To Msf列表.Rows - 1
        If Val(Msf列表.TextMatrix(n, 处方列名.标志)) = 1 Then
            Msf列表.Row = n
            Call Msf列表_EnterCell
            DoEvents
            
            '检查药品存储库房
            If CheckDrugStock = False Then Exit Function
            
            '检测是否允许
            If CheckBill(3, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
            
            '检查是否收费(发药处理)
            gstrSQL = " Select Decode(配药人,Null,'','部门发药','',配药人) 配药人,已收费 From 未发药品记录" & _
                     " Where No=[1] And (库房ID=[3] Or 库房ID Is NULL) And 单据=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex), lng药房ID)
            
            With rsTemp
                If IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
                    If IntSendAfterDosage = 0 Then
                        If IsNull(!配药人) Then
                            MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If Trim(!配药人) = "" Then
                            MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                mstr配药人 = NVL(!配药人)
                
                If mint允许未收费处方发药 = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) = 8 Then
                    If !已收费 = 0 Then
                        MsgBox "该处方还未收费，不能执行发药操作！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                If Int允许未审核处方发药 = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 8 Then
                    If !已收费 = 0 Then
                        MsgBox "该处方还未审核，不能执行发药操作！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                Call GetBillSequence
                If str序号 = "" Then Exit Function
                If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str序号) Then Exit Function
                If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Function
                If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf列表.TextMatrix(n, 处方列名.金额)) Then Exit Function
                
                '校验发药人
                If int校验发药人 = 1 And Not BlnFirst Then
                    mstr操作员 = zlDatabase.UserIdentify(Me, "校验发药人", glngSys, 1341, "发药")
                    BlnFirst = True
                Else
                    mstr操作员 = gstrUserName
                End If
                If mstr操作员 = "" Then Exit Function
                    
                If Not CheckSpec(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
                
                If mstr毒麻类提示 <> "" Then
                    If MsgBox("单号为[" & TxtNo & "]" & "的处方中含有以下毒麻类药品，确定发药吗？" & mstr毒麻类提示, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    mstr毒麻类提示 = ""
                End If
                
                If Not CheckStock(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
            End With
        End If
    Next
    
    CheckBatchRecipe = True
    Exit Function
ErrHand:
    CheckBatchRecipe = False
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Integer
    
    For lngRow = 1 To Bill处方明细.Rows - 2
        gstrSQL = "Select 收费细目id From 收费执行科室 Where 执行科室id = [1] And 收费细目id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查药品存储库房", lng药房ID, Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID)))
        
        If rsTmp.EOF Then
            MsgBox Bill处方明细.TextMatrix(lngRow, 列名.药品名称) & "未设置存储库房，不能发药！", vbInformation, gstrSysName
            Exit Function
        End If
    Next

    CheckDrugStock = True
End Function
Private Function CheckBillExist(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select ID From 药品收发记录 " & _
             " Where 单据=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查单据是否存在", int单据, strNo)
    CheckBillExist = Not rsTemp.EOF
End Function
Private Function CheckIsSended(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    '检查是否已退药
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Count(Id) From 药品收发记录 Where 单据 = [1] And NO = [2] And 记录状态 <> 1 And 审核日期 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否允许取消发药", int单据, strNo)
    
    CheckIsSended = (rsTemp.RecordCount > 0)
End Function

Private Function CheckRecipe() As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long, lng药品ID As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    
    On Error GoTo ErrHand
    
    '检查病人费用余额
    If Not CheckSendBillMoney(False) Then Exit Function
    
    '检查药品存储库房
    If CheckDrugStock = False Then Exit Function
    
    '检测是否允许
    If CheckBill(3, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
    '检查是否收费(发药处理)
    gstrSQL = " Select Decode(配药人,Null,'','部门发药','',配药人) 配药人,已收费 From 未发药品记录" & _
             " Where No=[1] And (库房ID=[3] Or 库房ID Is NULL) And 单据=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex), lng药房ID)
    
    With rsTemp
        If IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            If IntSendAfterDosage = 0 Then
                If IsNull(!配药人) Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(!配药人) = "" Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        mstr配药人 = NVL(!配药人)
        
        If mint允许未收费处方发药 = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) = 8 Then
            If !已收费 = 0 Then
                MsgBox "该处方还未收费，不能执行发药操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If Int允许未审核处方发药 = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 8 Then
            If !已收费 = 0 Then
                MsgBox "该处方还未审核，不能执行发药操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If intVerify = 1 And Txt开单医生.ListIndex = 0 Then Txt开单医生.Enabled = True: Txt开单医生.SetFocus: Exit Function
        Call GetBillSequence
        If str序号 = "" Then Exit Function
        If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str序号) Then Exit Function
        If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Function
        If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf列表.TextMatrix(Msf列表.Row, 处方列名.金额)) Then Exit Function
        
        '校验发药人
        If int校验发药人 = 1 Then
            mstr操作员 = zlDatabase.UserIdentify(Me, "校验发药人", glngSys, 1341, "发药")
        Else
            mstr操作员 = gstrUserName
        End If
        If mstr操作员 = "" Then Exit Function
        
        If Not CheckSpec(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
        
        If mstr毒麻类提示 <> "" Then
            If MsgBox("单号为[" & TxtNo & "]" & "的处方中含有以下毒麻类药品，确定发药吗？" & mstr毒麻类提示, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("你确定单号为[" & TxtNo & "]" & "的处方发药吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
        
        If Not CheckStock(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
    End With
    
    CheckRecipe = True
    Exit Function
ErrHand:
    CheckRecipe = False
End Function

Private Function CheckSendBillMoney(ByVal blnBatch As Boolean) As Boolean
    '发药检查－检查病人费用余额，并根据记帐报警设置作相应处理
    'blnBatch：True-批量发药;False-单处方发药
    '主要算法：
    '1、系统参数"执行后自动审核"有效时才检查
    '2、只对记帐划价单
    '3、按病人ID计算单据汇总金额
    '4、根据记帐报警设置作相应处理
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rs费用类别 As ADODB.Recordset
    Dim strNo As String
    Dim lng病人ID As Long
    Dim str病人id As String
    Dim strFirstNo As String
    
    Dim cur处方金额 As Currency
    
    Dim str费用类别 As String
    Dim str费用类别名 As String
    
    On Error GoTo errH
    
    '系统参数"执行后自动审核"有效时才检查
    If int审核划价单 = 0 Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    If blnBatch Then
        With mrsBatchSend
            '只对记帐划价单才检查
            .Filter = "单据=9 And 未审核=1"
            
            '按病人ID计算单据汇总金额
            .Sort = "病人ID"
            
            If .RecordCount = 0 Then
                CheckSendBillMoney = True
                Exit Function
            End If
            
            .MoveFirst
            
            '根据记帐报警设置作相应处理
            Do While Not .EOF
                If lng病人ID <> Val(!病人ID) Then
                    If lng病人ID <> 0 Then
                        '判断是住院还是门诊病人
                        gstrSQL = "Select Distinct Decode(B.门诊标志, 1, '门诊', 4, '门诊', '住院') As 来源, " & _
                            " B.病人id,nvl(B.主页id,0) 主页id,Decode(B.门诊标志, 1, 0, 4, 0, B.病人病区id) 病人病区id, C.姓名 " & _
                            " From 药品收发记录 A,病人费用记录 B,病人信息 C " & _
                            " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                            " And A.单据=9 And A.no=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                        
                        '取费用类别
                        gstrSQL = " Select Distinct b.编码, b.名称 " & _
                            " From 病人费用记录 a, 收费项目类别 b, 药品收发记录 c " & _
                            " Where a.收费类别 = b.编码 And a.Id = c.费用id And c.单据 = 9 And c.No In([1]) "
                        Set rs费用类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                        
                        Do While Not rs费用类别.EOF
                            str费用类别 = str费用类别 & rs费用类别!编码
                            str费用类别名 = str费用类别名 & "," & rs费用类别!名称
                            rs费用类别.MoveNext
                        Loop
                                            
                        '检查费用余额
                        If Not FinishBillingWarn(rsTmp, cur处方金额, str费用类别, str费用类别名) Then
                            CheckSendBillMoney = False
                            Exit Function
                        End If
                    End If
                    
                    strNo = !NO
                    cur处方金额 = Val(!金额)
                    strFirstNo = !NO
                    lng病人ID = Val(!病人ID)
                Else
                    strNo = strNo & "," & !NO
                    cur处方金额 = cur处方金额 + Val(!金额)
                End If
                
                .MoveNext
                
                If .EOF Then
                    '判断是住院还是门诊病人
                    gstrSQL = "Select Distinct Decode(B.门诊标志, 1, '门诊', 4, '门诊', '住院') As 来源, " & _
                        " B.病人id,nvl(B.主页id,0) 主页id,Decode(B.门诊标志, 1, 0, 4, 0, B.病人病区id) 病人病区id, C.姓名 " & _
                        " From 药品收发记录 A,病人费用记录 B,病人信息 C " & _
                        " Where A.费用id=B.Id And b.病人id = c.病人id " & _
                        " And A.单据=9 And A.no=[1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                    
                    '取费用类别
                    gstrSQL = " Select Distinct b.编码, b.名称 " & _
                        " From 病人费用记录 a, 收费项目类别 b, 药品收发记录 c " & _
                        " Where a.收费类别 = b.编码 And a.Id = c.费用id And c.单据 = 9 And c.No In([1]) "
                    Set rs费用类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                    
                    Do While Not rs费用类别.EOF
                        str费用类别 = str费用类别 & rs费用类别!编码
                        str费用类别名 = str费用类别名 & "," & rs费用类别!名称
                        rs费用类别.MoveNext
                    Loop
                                        
                    '检查费用余额
                    If Not FinishBillingWarn(rsTmp, cur处方金额, str费用类别, str费用类别名) Then
                        CheckSendBillMoney = False
                        Exit Function
                    End If
                End If
            Loop
        End With
    Else
        If Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 9 Or Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.未审核)) <> 1 Then
            CheckSendBillMoney = True
            Exit Function
        End If
        
        strNo = Mid(TxtNo.Text, 1, 8)
        
        cur处方金额 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.实收金额))
        
        '判断是住院还是门诊病人
        gstrSQL = "Select Distinct Decode(B.门诊标志, 1, '门诊', 4, '门诊', '住院') As 来源, " & _
            " B.病人id,nvl(B.主页id,0) 主页id,Decode(B.门诊标志, 1, 0, 4, 0, B.病人病区id) 病人病区id, C.姓名 " & _
            " From 药品收发记录 A,病人费用记录 B,病人信息 C " & _
            " Where A.费用id=B.Id And b.病人id = c.病人id " & _
            " And A.单据=9 And A.no=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
        
        '取费用类别
        gstrSQL = " Select Distinct b.编码, b.名称 " & _
            " From 病人费用记录 a, 收费项目类别 b, 药品收发记录 c " & _
            " Where a.收费类别 = b.编码 And a.Id = c.费用id And c.单据 = 9 And c.No In([1]) "
        Set rs费用类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
        
        Do While Not rs费用类别.EOF
            str费用类别 = str费用类别 & rs费用类别!编码
            str费用类别名 = str费用类别名 & "," & rs费用类别!名称
            rs费用类别.MoveNext
        Loop
                            
        '检查费用余额
        If Not FinishBillingWarn(rsTmp, cur处方金额, str费用类别, str费用类别名) Then
            CheckSendBillMoney = False
            Exit Function
        End If
    End If
    CheckSendBillMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub GetDosagePeople()
    Dim rsTemp As ADODB.Recordset
    '配药人
    gstrSQL = " Select 简码||'-'||姓名 As 姓名 From 人员表  Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in " & _
             " (Select Distinct 人员ID From 人员性质说明 Where 人员性质='药房发药人' " & _
             " And 人员ID IN (Select 人员ID From 部门人员 Where 部门ID=[1]))" & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
    
    With rsTemp
        Me.cbo配药人.Clear
        Do While Not .EOF
            cbo配药人.AddItem !姓名
            .MoveNext
        Loop
        If cbo配药人.ListCount = 0 Then
            cbo配药人.Enabled = False
        End If
    End With
End Sub

Private Sub GetRecipeColor()
    mstrUserRecipeColor = zlDatabase.GetPara("处方颜色", glngSys, 1341)

    If mstrUserRecipeColor = "" Then
        Call GetDefaultRecipeColor
    End If
End Sub

Private Sub GetDefaultRecipeColor()
    mstrUserRecipeColor = CStr(mconlng普通) & ";" & _
                    CStr(mconlng急诊) & ";" & _
                    CStr(mconlng儿科) & ";" & _
                    CStr(mconlng麻醉) & ";" & _
                    CStr(mconlng精一) & ";" & _
                    CStr(mconlng精二)

End Sub
Private Function GetSumMoney(ByVal rsRecipt As ADODB.Recordset) As String
    Dim rsTemp As ADODB.Recordset
    Dim dblSum As Double
    
    Set rsTemp = rsRecipt.Clone
    
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            dblSum = dblSum + Val(.Fields("金额").Value)
            .MoveNext
        Loop
    End With
    
    GetSumMoney = FormatEx(dblSum, mintMoneyDigit)
End Function
Private Sub GetSysParms()
    Int允许未审核处方发药 = gtype_UserSysParms.P6_未审核记帐处方发药
    mint允许未收费处方发药 = gtype_UserSysParms.P148_未收费处方发药
    
    mbln允许取消发药 = (gtype_UserSysParms.P15_门诊收费与发药分离 = 1 Or gtype_UserSysParms.P16_住院记帐与发药分离 = 1)
    
    bln医嘱作废 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0)          '为零表示允许退药
    
    '获取金额小数位数
    int金额保留位数 = gtype_UserSysParms.P9_费用金额保留位数
    
    '判断划价单发药后是否自动审核为记帐单
    int审核划价单 = gtype_UserSysParms.P81_执行后自动审核划价单
    
    '记帐报警包含划价费用
    mbln报警包含划价费用 = gtype_UserSysParms.P98_记帐报警包含划价费用 <> 0

End Sub

Private Sub IniRecord()
    Set mrsBatchSend = New ADODB.Recordset
    With mrsBatchSend
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "未审核", adDouble, 18, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
Private Sub PrintRecipe()
    '打印处方
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    
    If mstrPrintRecipe = "" Then Exit Sub
    
    If IntAutoPrint < 2 Then
        blnPrint = IIf(IntAutoPrint = 1, True, False)
        If IntAutoPrint = 0 Then
            If MsgBox("打印该处方单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
        
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                strRecipeNo = Split(arrRecipe(n), ",")(0)
                intBillType = Val(Split(arrRecipe(n), ",")(1))
            
                If Not BillHaveHerial(strRecipeNo, intBillType) Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                        "NO=" & strRecipeNo, _
                        "性质=" & IIf(intBillType = 8, 1, 2), _
                        "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
                        "ReportFormat=1", "PrintEmpty=0", 2)
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                        "NO=" & strRecipeNo, _
                        "性质=" & IIf(intBillType = 8, 1, 2), _
                        "ReportFormat=1", "PrintEmpty=0", 2)
                End If
            Next
        End If
    End If
    
    Me.MnuFileRePrint.Caption = "重打已发药处方-" & strRecipeNo & "(&D)"
    
    mstrPrintRecipe = ""
End Sub

Private Sub Select病区()
    Dim rsTmp As ADODB.Recordset
    
    If cbo病区.ListCount > 0 Then Exit Sub
    
    '病区
    gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
             " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
             " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order By 编码||'-'||名称 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病区")
    
    With cbo病区
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!科室
            .ItemData(.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If .ListIndex <> -1 Then
            .ListIndex = 0
        End If
    End With
End Sub

Private Function SendBatchRecipe() As Boolean
    Dim n As Integer
    Dim lngRow As Long, lng药品ID As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim rsSendRecipeDetail As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "配药人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "填制人", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsSendRecipeDetail = New ADODB.Recordset
    With rsSendRecipeDetail
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    For n = 1 To Msf列表.Rows - 1
        If Val(Msf列表.TextMatrix(n, 处方列名.标志)) = 1 Then
            Msf列表.Row = n
            Call Msf列表_EnterCell
            DoEvents
            
            With rsSendRecipeByNo
                .AddNew
                !NO = Mid(TxtNo.Text, 1, 8)
                !单据 = TxtNo.ItemData(TxtNo.ListIndex)
                !配药人 = cbo配药人.Text
                !填制人 = IIf(Txt开单医生.ListIndex = 0, "", Mid(Txt开单医生, InStr(1, Txt开单医生, "-") + 1))
                .Update
            End With
            
            With rsSendRecipeDetail
                For lngRow = 1 To Bill处方明细.Rows - 2
                    .AddNew
                    !NO = Mid(TxtNo.Text, 1, 8)
                    !收发ID = Val(Bill处方明细.TextMatrix(lngRow, 列名.Id))
                    !药品ID = Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID))
                    !批次 = Val(Bill处方明细.TextMatrix(lngRow, 列名.批次))
                    .Update
                  Next
            End With
            
'            '先更新批次
'            For lngRow = 1 To Bill处方明细.Rows - 2
'                LngID = Val(Bill处方明细.TextMatrix(lngRow, 列名.Id))
'                lng药品ID = Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID))
'                lng批次 = Val(Bill处方明细.TextMatrix(lngRow, 列名.批次))
'                gstrSQL = "zl_药品收发记录_更新批次(" & LngID & "," & lng药品ID & "," & lng批次 & ")"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新批次")
'            Next
'
'            If IntSendAfterDosage = 0 Then
'                '必须经过配药过程，则配药人不填
'                gstrSQL = "zl_药品收发记录_处方发药(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
'                                  "','" & mstr操作员 & "'" & ",NULL," & IIf(Txt开单医生.ListIndex = 0, "NULL", _
'                                  "'" & Mid(Txt开单医生, InStr(1, Txt开单医生, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "', " & int金额保留位数 & "," & int审核划价单 & ")"
'            Else
'                gstrSQL = "zl_药品收发记录_处方发药(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
'                                  "','" & mstr操作员 & "'" & ",'" & cbo配药人.Text & "'," & IIf(Txt开单医生.ListIndex = 0, "NULL", _
'                                  "'" & Mid(Txt开单医生, InStr(1, Txt开单医生, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "'," & int金额保留位数 & "," & int审核划价单 & ")"
'            End If
'            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品发药")
'
'            '如果已启用了电子签名，则需要对配药人进行电子签名处理
'            If gbln药品使用电子签名 = True Then
'                If SaveSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo.Text, 1, 8), lng药房ID) = False Then
'                    Exit Function
'                End If
'            End If
'
'            '记录该处方号及单据类型
'            strBill = Mid(TxtNo.Text, 1, 8) & "|" & TxtNo.ItemData(TxtNo.ListIndex)
'            mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & Mid(TxtNo.Text, 1, 8) & "," & TxtNo.ItemData(TxtNo.ListIndex)
        End If
    Next
    
    '按处方号排序后批量发药
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For n = 1 To rsSendRecipeByNo.RecordCount
        rsSendRecipeDetail.Filter = "NO='" & rsSendRecipeByNo!NO & "'"
        rsSendRecipeDetail.MoveFirst
        For lngRow = 1 To rsSendRecipeDetail.RecordCount
            gstrSQL = "zl_药品收发记录_更新批次(" & rsSendRecipeDetail!收发ID & "," & rsSendRecipeDetail!药品ID & "," & rsSendRecipeDetail!批次 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新批次")
            
            rsSendRecipeDetail.MoveNext
        Next
        
        gstrSQL = "zl_药品收发记录_处方发药("
        '库房ID
        gstrSQL = gstrSQL & lng药房ID
        '单据
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '发药人(审核人)
        gstrSQL = gstrSQL & ",'" & mstr操作员 & "'"
        '配药人(必须经过配药过程时，则配药人不填)
        gstrSQL = gstrSQL & "," & IIf(IntSendAfterDosage = 0, "Null", IIf(rsSendRecipeByNo!配药人 = "", "NULL", "'" & rsSendRecipeByNo!配药人 & "'")) & ""
        '校验人（开单医生）
        gstrSQL = gstrSQL & "," & IIf(rsSendRecipeByNo!填制人 = "", "NULL", "'" & rsSendRecipeByNo!填制人 & "'") & ""
        '发药方式
        gstrSQL = gstrSQL & ",1"
        '发药时间
        gstrSQL = gstrSQL & ",Null"
        '操作员编码
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '操作员名称
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '金额保留位数
        gstrSQL = gstrSQL & "," & int金额保留位数
        '自动审核记账单
        gstrSQL = gstrSQL & "," & int审核划价单
        gstrSQL = gstrSQL & ")"
       
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品发药")

        '如果已启用了电子签名，则需要对配药人进行电子签名处理
        If gbln药品使用电子签名 = True Then
            If SaveSignatureRecored(EsignTache.send, rsSendRecipeByNo!单据, rsSendRecipeByNo!NO, lng药房ID) = False Then
                Exit Function
            End If
        End If

        '记录该处方号及单据类型
        strBill = rsSendRecipeByNo!NO & "|" & rsSendRecipeByNo!单据
        mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsSendRecipeByNo!NO & "," & rsSendRecipeByNo!单据
        
        rsSendRecipeByNo.MoveNext
    Next
    
    mstr操作员 = ""
    mstr配药人 = ""
    Txt开单医生.Enabled = False
    Me.TxtNo.SetFocus
            
    SendBatchRecipe = True
    Exit Function
ErrHand:
    SendBatchRecipe = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SendRecipe() As Boolean
    Dim lngRow As Long, lng药品ID As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    
    On Error GoTo ErrHand
    
    '先更新批次
    For lngRow = 1 To Bill处方明细.Rows - 2
        LngID = Val(Bill处方明细.TextMatrix(lngRow, 列名.Id))
        lng药品ID = Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID))
        lng批次 = Val(Bill处方明细.TextMatrix(lngRow, 列名.批次))
        gstrSQL = "zl_药品收发记录_更新批次(" & LngID & "," & lng药品ID & "," & lng批次 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新批次")
    Next
    If IntSendAfterDosage = 0 Then
        '必须经过配药过程，则配药人不填
        gstrSQL = "zl_药品收发记录_处方发药(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
                          "','" & mstr操作员 & "'" & ",NULL," & IIf(Txt开单医生.ListIndex = 0, "NULL", _
                          "'" & Mid(Txt开单医生, InStr(1, Txt开单医生, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "', " & int金额保留位数 & "," & int审核划价单 & ")"
    Else
        gstrSQL = "zl_药品收发记录_处方发药(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
                          "','" & mstr操作员 & "'" & ",'" & cbo配药人.Text & "'," & IIf(Txt开单医生.ListIndex = 0, "NULL", _
                          "'" & Mid(Txt开单医生, InStr(1, Txt开单医生, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "'," & int金额保留位数 & "," & int审核划价单 & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品发药")
    
    '如果已启用了电子签名，则需要对配药人进行电子签名处理
    If gbln药品使用电子签名 = True Then
        If SaveSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo.Text, 1, 8), lng药房ID) = False Then
            Exit Function
        End If
    End If
    
    '记录该处方号及单据类型
    strBill = Mid(TxtNo.Text, 1, 8) & "|" & TxtNo.ItemData(TxtNo.ListIndex)
    mstrPrintRecipe = Mid(TxtNo.Text, 1, 8) & "," & TxtNo.ItemData(TxtNo.ListIndex)
    
    mstr操作员 = ""
    mstr配药人 = ""
    
    Txt开单医生.Enabled = False
    Me.TxtNo.SetFocus
    
    SendRecipe = True
    Exit Function
ErrHand:
    SendRecipe = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetBatchSendRecord()
    Dim n As Integer
    
    Call IniRecord
    With mrsBatchSend
        For n = 1 To Msf列表.Rows - 1
            If Val(Msf列表.TextMatrix(n, 处方列名.标志)) = 1 And Msf列表.TextMatrix(n, 处方列名.NO) <> "" Then
                .AddNew
                !NO = Msf列表.TextMatrix(n, 处方列名.NO)
                !单据 = Msf列表.TextMatrix(n, 处方列名.单据)
                !病人ID = Val(Msf列表.TextMatrix(n, 处方列名.病人ID))
                !未审核 = Val(Msf列表.TextMatrix(n, 处方列名.未审核))
                !金额 = Val(Msf列表.TextMatrix(n, 处方列名.实收金额))
                .Update
            End If
        Next
    End With
End Sub
Private Sub SetCheckBox(Optional ByVal intRow As Integer = -1)
    'intRow = -1 对所有行重新做标志
    'intRow = 0  点击标题行时对所有行重新做标志
    'intRow > 0  对指定行做标志
    
    Dim n As Integer
    Dim strFlagName As String
    Dim i As Integer
    
    With Msf列表
        If .Rows <= 1 Then Exit Sub
         
        .Redraw = False
         
        .Col = 处方列名.选择
         
        If intRow = -1 Then
            For n = 0 To .Rows - 1
                .Row = n
                If Val(.TextMatrix(n, 处方列名.标志)) = 0 Then
                    strFlagName = "checked"
                    .TextMatrix(n, 处方列名.标志) = 1
                Else
                    strFlagName = "unchecked"
                    .TextMatrix(n, 处方列名.标志) = 0
                End If
                Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
            Next
        ElseIf intRow = 0 Then
            If Val(.TextMatrix(intRow, 处方列名.标志)) = 0 Then
                strFlagName = "checked"
            Else
                strFlagName = "unchecked"
            End If
            
            For n = 0 To .Rows - 1
                .Row = n
                .TextMatrix(n, 处方列名.标志) = Abs(Val(.TextMatrix(n, 处方列名.标志)) - 1)
                Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
            Next
        Else
            .Row = intRow
            If Val(.TextMatrix(intRow, 处方列名.标志)) = 0 Then
                strFlagName = "checked"
            Else
                strFlagName = "unchecked"
            End If
            .TextMatrix(intRow, 处方列名.标志) = Abs(Val(.TextMatrix(intRow, 处方列名.标志)) - 1)
            Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
        End If
        
        Call SetBatchSendRecord
        
        .Redraw = True
    End With
End Sub

Private Sub SetColHide()
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '根据用户的列设置，禁止显示部分列
    strSave = zlDatabase.GetPara("列设置", glngSys, 1341)
    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|规格,0|批号,0|单位,0|单价,0|数量,0|金额,0|重量,0|用法,0|用量,0|频次,0|医生嘱托,0|费别,0|库存数,0|库房货位,0|已退数,0|准退数,0|退药数,0|备注"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    mbln显示重量 = False
    With Bill处方明细
        For intRow = 0 To intRows
            intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1))
            If intCol > -1 Then
                If Split(arrColumn(intRow), "|")(1) = "药品名称" Then
                    int药品名称 = Val(Split(arrColumn(intRow), "|")(0))
                Else
                    If Val(Split(arrColumn(intRow), "|")(0)) = 1 Then
                        .ColWidth(intCol) = 0
                    ElseIf .ColWidth(intCol) = 0 Then
                        Select Case Split(arrColumn(intRow), "|")(1)
                        Case "其它名"
                            .ColWidth(列名.其它名) = 2000
                        Case "英文名"
                            .ColWidth(列名.英文名) = 2000
                        Case "规格"
                            .ColWidth(列名.规格) = 1500
                        Case "批号"
                            .ColWidth(列名.批号) = 1500
                        Case "单位"
                            .ColWidth(列名.单位) = IIf(mbln显示大小单位 = True, 0, 500)
                        Case "单价"
                            .ColWidth(列名.单价) = 1000
                        Case "数量"
                            .ColWidth(列名.数量) = 1200
                        Case "金额"
                            .ColWidth(列名.金额) = 1200
                        Case "用量"
                            .ColWidth(列名.单量) = 1200
                        Case "用法"
                            .ColWidth(列名.用法) = 1500
                        Case "频次"
                            .ColWidth(列名.频次) = 1500
                        Case "备注"
                            .ColWidth(列名.备注) = 1200
                        Case "费别"
                            .ColWidth(列名.费别) = 1000
                        Case "库房货位"
                            .ColWidth(列名.货位) = IIf(MnuEditHandback.Checked, 0, 1200)
                        Case "重量"
                            mbln显示重量 = True
                            If mblnIs中药处方 Then
                                .ColWidth(列名.重量) = 1200
                            End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '如果是退药状态，这些列必须显示
        .ColWidth(列名.已退数) = IIf(MnuEditHandback.Checked, 1200, 0)
        .ColWidth(列名.准退数) = IIf(MnuEditHandback.Checked, 1200, 0)
        .ColWidth(列名.退药数) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = False, 1200, 0)
        .ColWidth(列名.准退数大) = 0
        .ColWidth(列名.准退数小) = 0
        .ColWidth(列名.退药数大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
        .ColWidth(列名.退药数小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
        .ColWidth(列名.单位大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
        .ColWidth(列名.单位小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
    End With
End Sub
Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub


Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号，lngCmd=0时需要
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If
    
    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
    strSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,病人费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, IntBillStyle)
    
    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If
    
    mlng病人ID = rsTmp!病人ID
    mstr挂号单 = NVL(rsTmp!挂号单)
    mlng主页ID = rsTmp!主页id
    
    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    If mlng病人ID <> mlngPassPati Then
        If mstr挂号单 <> "" Then               '门诊病人
            strSQL = "Select 病人ID,Count(Distinct Trunc(登记时间)) as 就诊次数 From 病人挂号记录 Where 病人ID=[1] Group by 病人ID"
            strSQL = "Select D.就诊次数,A.姓名,A.性别,A.出生日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,(" & strSQL & ") D,人员表 E" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=D.病人ID" & _
                " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlng病人ID, rsTmp!就诊次数, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), "")
        Else                                    '住院病人
            strSQL = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlng病人ID, mlng主页ID, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), _
                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        End If
        mlngPassPati = mlng病人ID
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With Bill处方明细
            '取药品名称
            str药品 = .TextMatrix(lngRow, 列名.药品名称)
            If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
            If InStr(str药品, "(") > 0 Then str药品 = Left(str药品, InStr(str药品, "(") - 1)
            '取药品给药途径
            str用法 = .TextMatrix(lngRow, 列名.用法)
            
            '传入查询药品信息
            Call PassSetQueryDrug(.TextMatrix(lngRow, 列名.药品ID), str药品, mstr单量单位, str用法)
                
            '设置菜单可用状态
            Call SetPassMenuState
            
            AdviceCheckWarn = 1 '表示可以弹出菜单
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


Private Sub SetFilter(ByVal blnState As Boolean)
    Dim strFind As String
    
    strFind = Trim(txtFind.Text)
    
    If strFind = "" Then Exit Sub
    mstrFilter = ""
    
    Select Case lblFind.Tag
        Case FindType.就诊卡
            If blnState = False Then
                mstrFilter = mstrFilter & " And Upper(A.就诊卡号) = [6] "
            Else
                mstrFilter = mstrFilter & " And Upper(B.就诊卡号) = [6] "
            End If
            SQLCondition.str就诊卡 = strFind
        Case FindType.门诊号
            If blnState = False Then
                mstrFilter = mstrFilter & " And Upper(A.门诊号) = [14] "
            Else
                mstrFilter = mstrFilter & " And Upper(B.门诊号) = [14] "
            End If
            SQLCondition.str门诊号 = strFind
        Case FindType.单据号
            If IsNumeric(strFind) Then
                strFind = GetFullNO(strFind, 13)
                txtFind.Text = strFind
            End If
            strFind = UCase(strFind)
            mstrFilter = mstrFilter & " And A.NO Between [3] And [4] "
            SQLCondition.str开始NO = strFind
            SQLCondition.str结束NO = strFind
        Case FindType.姓名
            If mblnCard = True Then
                If blnState = False Then
                    mstrFilter = mstrFilter & " And Upper(A.就诊卡号) = [6] "
                Else
                    mstrFilter = mstrFilter & " And Upper(B.就诊卡号) = [6] "
                End If
                SQLCondition.str就诊卡 = strFind
            Else
                If blnState = False Then
                    mstrFilter = mstrFilter & " And Upper(A.姓名) Like Upper([5]) "
                Else
                    mstrFilter = mstrFilter & " And Upper(B.姓名) Like Upper([5]) "
                End If
                SQLCondition.str姓名 = strFind & "%"
            End If
        Case FindType.身份证
            If blnState = False Then
                mstrFilter = mstrFilter & " And A.身份证号 = [15] "
            Else
                mstrFilter = mstrFilter & " And B.身份证号 = [15] "
            End If
            SQLCondition.str身份证 = strFind
        Case FindType.IC卡
            If blnState = False Then
                mstrFilter = mstrFilter & " And A.IC卡号 = [16] "
            Else
                mstrFilter = mstrFilter & " And B.IC卡号 = [16] "
            End If
            SQLCondition.strIC卡 = strFind
    End Select

    Call mnuViewRefresh_Click
'    mstrFilter = ""
End Sub
Private Sub SetPosition()
    If mbln发病区处方 Then
        img病区.Top = PicBackGroud.Top - img病区.Height - 50
        img病区.Left = PicBackGroud.Left
        
        If img病区.BorderStyle = 1 Then
            cbo病区.Visible = True
            cbo病区.Top = img病区.Top - 20
            cbo病区.Left = img病区.Left + img病区.Width + 50
            
            Chk清单.Top = img病区.Top + 20
            
            If Chk清单.Visible Then
                Chk清单.Left = cbo病区.Left + cbo病区.Width + 200
            End If
            
            If Chk显示退药待发单据.Visible Then
                Chk显示退药待发单据.Left = cbo病区.Left + cbo病区.Width + 200
            End If
            
            Call Select病区
        Else
            cbo病区.Visible = False
            
            Chk清单.Top = img病区.Top + 20
            
            If Chk清单.Visible Then
                Chk清单.Left = img病区.Left + img病区.Width + 200
            End If
            
            If Chk显示退药待发单据.Visible Then
                Chk显示退药待发单据.Left = img病区.Left + img病区.Width + 200
            End If
        End If
    Else
        Chk清单.Top = PicBackGroud.Top - Chk清单.Height - 50
        Chk清单.Left = PicBackGroud.Left
    End If
    
    Chk显示退药待发单据.Top = Chk清单.Top
End Sub

Private Sub SetRecipeColor()
    '标记处方颜色
    Dim lngRow As Integer
    
    Msf列表.Redraw = False
'    Msf列表.TextMatrix(0, 处方列名.颜色) = ""
    For lngRow = 1 To Msf列表.Rows - 1
        Msf列表.Row = lngRow
        Msf列表.Col = 处方列名.颜色
        Msf列表.CellBackColor = Split(mstrUserRecipeColor, ";")(Val(Msf列表.TextMatrix(lngRow, 处方列名.处方类型)))
    Next
    Msf列表.Redraw = True
End Sub

Private Sub SetTimerState(ByVal BlnSet As Boolean)
    '关闭和启用Timer控件，有弹出窗口时调用
    'blnSet：True-开启；False-关闭
    
    If BlnSet Then
        '开启时恢复原来的状态
        TimeRefresh.Enabled = mblnStateTimeRefresh
        TimePrint.Enabled = mblnStateTimePrint
    Else
        '关闭时先记录原来的状态
        mblnStateTimeRefresh = TimeRefresh.Enabled
        mblnStateTimePrint = TimePrint.Enabled
        
        If mblnStateTimeRefresh Then TimeRefresh.Enabled = False
        If mblnStateTimePrint Then TimePrint.Enabled = False
    End If
End Sub

Private Function 判断是否中药处方(ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    '通过药品id判断是否是中药
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim DblWidth As Double

    strSQL = "Select a.类别 as 类别 From 收费项目目录 a ,药品收发记录 b Where b.药品id=a.Id And b.单据=[2] and b.No=[1] And (b.记录状态=1 Or Mod(b.记录状态,3)=0) and (b.库房ID+0=[3] OR b.库房ID IS NULL) " _
        & " union all " _
        & "Select a.类别 as 类别 From 收费项目目录 a ,H药品收发记录 b Where b.药品id=a.Id And b.单据=[2] and b.No=[1] And (b.记录状态=1 Or Mod(b.记录状态,3)=0) and (b.库房ID+0=[3] OR b.库房ID IS NULL) "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[判断是否中药处方]", BillNo, BillType, lng药房ID)
    
    mblnIs中药处方 = IIf(rs!类别 = 7, True, False)
    rs.Close
    
    On Error Resume Next
    
    DblWidth = Me.ScaleWidth - (ImgLeftRight_S.Left + ImgLeftRight_S.Width)
    If mblnIs中药处方 Then
        With Bill处方明细
            .Top = Txt床号.Top + Txt床号.Height + 50
            .Height = IIf(txt原始付数.Top - .Top - 50 < 0, .Height, txt原始付数.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    Else
        With Bill处方明细
            .Top = Txt床号.Top + Txt床号.Height + 50
            .Height = IIf(cbo配药人.Top - .Top - 50 < 0, .Height, cbo配药人.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
            If .ColWidth(列名.重量) <> 0 Then
                .ColWidth(列名.重量) = 0
            End If
        End With
    End If
    
    判断是否中药处方 = mblnIs中药处方
    
End Function

Private Sub 中药处方特别处理(ByVal BillStyle As Integer, ByVal BillNo As String)
    '中药处方显示原始付数和中药煎法
    Dim strSQL As String
    Dim rs As New ADODB.Recordset

    strSQL = "Select a.外观,b.付数 From 药品收发记录 a ,病人费用记录 b Where a.费用id=b.Id " _
        & " And a.单据=[2] And a.No=[1] " _
        & " And (a.记录状态=1 Or Mod(a.记录状态,3)=0) and (a.库房ID+0=[3] OR a.库房ID IS NULL) " _
        & " union all " _
        & " Select a.外观,b.付数 From H药品收发记录 a ,H病人费用记录 b Where a.费用id=b.Id " _
        & " And a.单据=[2] And a.No=[1] " _
        & " And (a.记录状态=1 Or Mod(a.记录状态,3)=0) and (a.库房ID+0=[3] OR a.库房ID IS NULL) "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[中药处方特别处理]", BillNo, BillStyle, lng药房ID)
    
    txt原始付数.Text = CStr(IIf(IsNull(rs!付数), 1, rs!付数))
    txt中药煎法.Text = IIf(IsNull(rs!外观), "", rs!外观)
    
    rs.Close
    
End Sub

Private Sub Bill处方明细_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill处方明细_cboClick(ListIndex As Long)
    With Bill处方明细
        If Not .Active Then Exit Sub
        If .ListCount = 0 Then Exit Sub
        .TextMatrix(.Row, 列名.批号) = .CboText
        .TextMatrix(.Row, 列名.批次) = .ItemData(.ListIndex)
    End With
End Sub
Private Sub Bill处方明细_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Call Bill处方明细_cboClick(Bill处方明细.ListIndex)
End Sub

Private Sub Bill处方明细_EnterCell(Row As Long, Col As Long)
    Dim lng批次 As Long, lng药品ID As Long, Dbl数量 As Double, blnAllow As Boolean
    Dim strNo As String, int单据 As Integer, strUnit As String, str包装 As String
    Dim rs批号 As New ADODB.Recordset
    
    If Not BlnEnterCell Then Exit Sub
    If TxtNo.ListIndex = -1 Or BlnRefresh = False Then Exit Sub
    
    '检查单据是否存在
    If Not CheckBillExist(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)), Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    Call ShowStock
    With Bill处方明细
        '设置当前行的颜色
        Call .SetRowColor(Row, &H8000000F, True)
        
        If .CboVisible Or .TxtVisible Then Exit Sub
        .ColData(.Col) = 0
        .Clear
        .Active = False
        .TxtVisible = False
        .CboVisible = False
        If .Row = .Rows - 1 Then
            If mblnAuto = False Then
                mintLastSequence = 0
            End If
            Exit Sub
        End If
        
        If Val(.TextMatrix(Row, 列名.药品ID)) = 0 Then Exit Sub    '药品ID为空，则退出
        
        mintLastSequence = Val(.TextMatrix(Row, 列名.序号))
        
        strUnit = GetUnit(lng药房ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
        Select Case strUnit
        Case "售价单位"
            str包装 = "1"
        Case "门诊单位"
            str包装 = "门诊包装"
        Case "住院单位"
            str包装 = "住院包装"
        Case "药库单位"
            str包装 = "药库包装"
        End Select
        
        If (MnuEditDosage.Checked Or MnuEditConsignment.Checked) Then
            If Val(.TextMatrix(Row, 列名.批次)) = 0 Then Exit Sub    '药品批次为空，则退出
            If Not (.Col = 列名.批号) Then Exit Sub
            lng批次 = Val(.TextMatrix(Row, 列名.批次))
            lng药品ID = Val(.TextMatrix(Row, 列名.药品ID))
            Dbl数量 = FormatEx(Val(.TextMatrix(Row, 列名.数量)), 5)
            strNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
            int单据 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据))

            '如果存在发药记录且部分退药，则不允许修改批次信息
            blnAllow = False

            gstrSQL = " Select count(*) Records From 药品收发记录 " & _
                " Where (Mod(记录状态,3)=0 or 记录状态=1) And 审核人 Is Not NULL " & _
                " And NO=[1] And 库房ID=[3] And 单据=[2] " & _
                " And 药品ID=[4] And Nvl(批次,0)=[5]"
            Set rs批号 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, int单据, lng药房ID, lng药品ID, lng批次)
            
            With rs批号
                blnAllow = (!Records = 0)
            End With
            
            '提取所有批次信息
            gstrSQL = " SELECT B.上次批号 批号,B.批次,ROUND(B.实际数量/" & str包装 & ",2) 数量" & _
                " FROM 药品规格 A,药品库存 B,收费价目 C,收费项目目录 F" & _
                " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID AND B.库房ID = [1] AND B.药品ID=[2] AND A.药品ID = C.收费细目ID" & _
                " AND ((SYSDATE BETWEEN C.执行日期 AND C.终止日期) OR C.终止日期 IS NULL)" & _
                " AND NVL(批次,0)<>0 AND NVL(实际数量,0)<>0 AND 性质=1" & _
                " AND ROUND(DECODE(F.是否变价,NULL,C.现价,0,C.现价,B.实际金额/B.实际数量),2)=" & _
                "     (SELECT ROUND(DECODE(F.是否变价,NULL,C.现价,0,C.现价,B.实际金额/B.实际数量),2) 单价" & _
                "     FROM 药品规格 A,药品库存 B,收费价目 C,收费项目目录 F" & _
                "     WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID AND B.库房ID = [1] AND B.药品ID=[2] AND A.药品ID = C.收费细目ID" & _
                "     AND ((SYSDATE BETWEEN C.执行日期 AND C.终止日期) OR C.终止日期 IS NULL)" & _
                "     AND NVL(批次,0)<>0 AND NVL(实际数量,0)<>0 AND 性质=1 AND NVL(批次,0)=[3])" & _
                " AND ROUND(B.实际数量/" & str包装 & ",2)>=[4] AND (NVL(A.药房分批,0)=0 OR (NVL(A.药房分批,0)=1 AND (效期 IS NULL OR 效期>TRUNC(SYSDATE))))" & _
                " ORDER BY B.批次"
            Set rs批号 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, lng药品ID, lng批次, Dbl数量)
            
            With rs批号
                Do While Not .EOF
                    If (!批次 <> lng批次 And blnAllow) Or !批次 = lng批次 Then
                        Bill处方明细.AddItem IIf(IsNull(!批号), "", !批号) & "(" & !批次 & ")"
                        Bill处方明细.ItemData(Bill处方明细.NewIndex) = !批次
                    End If
                    .MoveNext
                Loop
            End With
        ElseIf MnuEditHandback.Checked Then
            If mbln显示大小单位 = True Then
                .Tag = Val(.TextMatrix(.Row, 列名.准退数大)) * Val(.TextMatrix(.Row, 列名.包装)) + Val(.TextMatrix(.Row, 列名.准退数小))
                If Not (.Col = 列名.退药数大 Or .Col = 列名.退药数小) Then Exit Sub
            Else
                .Tag = Val(.TextMatrix(.Row, 列名.准退数))
                If Not (.Col = 列名.退药数) Then Exit Sub
            End If
        Else
            Exit Sub
        End If
        
        If (MnuEditDosage.Checked Or MnuEditConsignment.Checked) Then
            .ColData(.Col) = IIf(.ListCount = 0, 0, 3)
            .Active = IIf((.ListCount > 0), True, False)
        ElseIf MnuEditHandback.Checked Then
            '如果该处方已转出，则不允许操作
            If Not zlDatabase.NOMoved("药品收发记录", Mid(TxtNo.Text, 1, 8), "单据=", TxtNo.ItemData(TxtNo.ListIndex)) Then
                .ColData(.Col) = 4
                .Active = CmdSend.Enabled
            End If
        End If
    End With
End Sub

Private Sub Bill处方明细_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim blnUnValid As Boolean
    Dim dblCount As Double
    Dim dblSumCount As Double
    Dim rsTemp As New ADODB.Recordset
    
    With Bill处方明细
        If KeyCode = vbKeyReturn Then
            If mbln显示大小单位 = True Then
                If Not (.TxtVisible And (.Col = 列名.退药数大 Or .Col = 列名.退药数小)) Then Exit Sub
            Else
                If Not (.TxtVisible And .Col = 列名.退药数) Then Exit Sub
            End If
            
            blnUnValid = False
            .Text = Trim(.Text)
            
            blnUnValid = (.Text = "")
            If Not blnUnValid Then blnUnValid = Not IsNumeric(.Text)
            If Not blnUnValid Then
                If mbln显示大小单位 = True Then
                    If .Col = 列名.退药数大 Then
                        dblSumCount = Val(.Text) * Val(.TextMatrix(.Row, 列名.包装)) + Val(.TextMatrix(.Row, 列名.退药数小))
                    Else
                        dblSumCount = Val(.TextMatrix(.Row, 列名.退药数大)) * Val(.TextMatrix(.Row, 列名.包装)) + Val(.Text)
                    End If
                Else
                    dblSumCount = Val(.Text)
                End If
                blnUnValid = Not ((Abs(dblSumCount) <= Abs(.Tag)) And ((Val(dblSumCount) >= 0 And Val(.Tag) >= 0) Or (Val(dblSumCount) <= 0 And Val(.Tag) <= 0)))
            End If
            
            If blnUnValid Then
                If mbln显示大小单位 = True Then
                    If .Col = 列名.退药数大 Then
                        .Text = Val(.TextMatrix(.Row, 列名.准退数大))
                    Else
                        .Text = Val(.TextMatrix(.Row, 列名.准退数小))
                    End If
                Else
                    .Text = Val(.Tag)
                End If
            End If
            
            '先检查是否是医嘱产生的药品记录
            '如果不是则不管
            '如果是，检查系统参数是否允许未作废医嘱退药，如果不允许，退药数为零
            '如果允许则不管
            dblCount = Val(FormatEx(.Text, 5))
            If dblCount <> 0 And bln医嘱作废 = False Then
                gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", Val(Bill处方明细.TextMatrix(Bill处方明细.Row, 列名.Id)))
                
                If (rsTemp!扣率 Like "1*") Then       '临嘱
                    gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 病人费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(Bill处方明细.TextMatrix(Bill处方明细.Row, 列名.Id)))
                    
                    If Not rsTemp.EOF Then
                        If (rsTemp!门诊标志 = 1 Or rsTemp!门诊标志 = 4) And rsTemp!医嘱序号 <> 0 Then
                            gstrSQL = "Select decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where ID=[1]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rsTemp!医嘱序号))

                            If rsTemp!作废 = 0 Then
                                dblCount = 0
                                MsgBox "该笔医嘱还未作废，不能退药！", vbInformation, gstrSysName
                            End If
                        End If
                    End If
                End If
            End If
            
            .Text = FormatEx(dblCount, 5)
            
            If mbln显示大小单位 = True Then
                If .Col = 列名.退药数大 Then
                    .TextMatrix(.Row, 列名.退药数大) = FormatEx(.Text, 5)
                Else
                    .TextMatrix(.Row, 列名.退药数小) = FormatEx(.Text, 5)
                End If
                .TextMatrix(.Row, 列名.退药数) = FormatEx(dblSumCount, 5) / Val(.TextMatrix(.Row, 列名.包装))
                
                If Val(.TextMatrix(.Row, 列名.退药数)) <> Val(.TextMatrix(.Row, 列名.实际数量)) / Val(.TextMatrix(.Row, 列名.包装)) Then
                    mblnAllBack = False
                End If
            Else
                .TextMatrix(.Row, 列名.退药数) = FormatEx(.Text, 5)
                
                If Val(.TextMatrix(.Row, 列名.退药数)) <> Val(.TextMatrix(.Row, 列名.准退数)) Then
                    mblnAllBack = False
                End If
            End If
        End If
    End With
End Sub

Private Sub Bill处方明细_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    If Button = 2 Then
        With Bill处方明细
            lngRow = .MouseRow
            If lngRow >= .MsfObj.FixedRows And lngRow < .Rows - 1 Then
                .Row = lngRow
            End If
        End With
    End If
    
End Sub

Private Sub Bill处方明细_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim str药品 As String
    
    'Pass
    If Button = 2 And gblnPass And tabShow.Tab = 2 And Len(Bill处方明细.TextMatrix(Bill处方明细.Row, 列名.医嘱id)) > 0 Then
            With Bill处方明细
            If .Rows > 1 And .Row < .Rows - 1 Then
                '检查Pass状态
                If AdviceCheckWarn(0, .Row) >= 0 Then PopupMenu mnuPass, 2
            End If
        End With
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


Private Sub Cbar_Resize()
    Form_Resize
End Sub

Private Sub cbo病区_Click()
    If cbo病区.ListIndex = -1 Then Exit Sub
    
    If cbo病区.ItemData(cbo病区.ListIndex) <> Val(cbo病区.Tag) Then
        cbo病区.Tag = cbo病区.ItemData(cbo病区.ListIndex)
        Call mnuViewRefresh_Click
    End If
End Sub


Private Sub cbo配药人_Click()
    '    Exit Sub
End Sub

Private Sub cbo配药人_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo配药人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cbo配药人_KeyPress(KeyAscii As Integer)
Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo配药人.Text)
        If cbo配药人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo配药人.List(cbo配药人.ListIndex) Then Call zlControl.CboSetIndex(cbo配药人.hWnd, -1)
        End If
        If strText = "" Then
            cbo配药人.ListIndex = -1
        ElseIf cbo配药人.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo配药人.ListCount - 1
                If Mid(cbo配药人.List(i), 1, InStr(1, cbo配药人.List(i), "-") - 1) = strText _
                    Or Mid(cbo配药人.List(i), InStr(1, cbo配药人.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo配药人.ListCount - 1
                    If UCase(cbo配药人.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo配药人.ListIndex = intIdx
            SendMessage cbo配药人.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo配药人_Click
            Exit Sub
        End If
        If cbo配药人.ListIndex = -1 Then
            cbo配药人.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo配药人_Click
            ElseIf intIdx <> cbo配药人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo配药人.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo配药人_Click
            End If
        End If
    End If
End Sub


Private Sub cbo配药人_LostFocus()
    Call cbo配药人_Validate(True)
End Sub
Private Sub cbo配药人_Validate(Cancel As Boolean)
    Dim n As Integer
    Dim blnFind As Boolean
    
    cbo配药人.Text = Trim(cbo配药人.Text)
    If InStr(cbo配药人.Text, "-") > 0 Then
        cbo配药人.Text = Mid(cbo配药人.Text, InStr(cbo配药人.Text, "-") + 1)
    End If
    If cbo配药人.Text <> "" Then
        For n = 0 To cbo配药人.ListCount - 1
            If cbo配药人.Text = Mid(cbo配药人.List(n), InStr(cbo配药人.List(n), "-") + 1) Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            cbo配药人.Text = ""
            Exit Sub
        End If
    End If
           
End Sub


Private Sub Chk清单_Click()
    Call mnuViewRefresh_Click
End Sub

Private Sub Chk全退_Click()
    Dim intRow As Integer
    Dim lng整数量 As Long
    Dim dbl小数量 As Double
    
    If Not Chk全退.Enabled Then Exit Sub
    With Bill处方明细
        For intRow = 1 To .Rows - 2
            If mbln显示大小单位 = True Then
                If Chk全退.Value = 1 Then
                    .TextMatrix(intRow, 列名.退药数大) = .TextMatrix(intRow, 列名.准退数大)
                    .TextMatrix(intRow, 列名.退药数小) = .TextMatrix(intRow, 列名.准退数小)
                    
                    .TextMatrix(intRow, 列名.退药数) = FormatEx(Val(.TextMatrix(intRow, 列名.实际数量)) / Val(.TextMatrix(intRow, 列名.包装)), mintNumberDigit)
                Else
                    .TextMatrix(intRow, 列名.退药数) = ""
                    .TextMatrix(intRow, 列名.退药数大) = ""
                    .TextMatrix(intRow, 列名.退药数小) = ""
                End If
            Else
                .TextMatrix(intRow, 列名.退药数) = IIf(Chk全退.Value = 1, .TextMatrix(intRow, 列名.准退数), "")
            End If
        Next
        mblnAllBack = (Chk全退.Value = 1)
    End With
End Sub
Private Sub Chk显示退药待发单据_Click()
    mlng待发单据 = Chk显示退药待发单据.Value
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdAlley_Click()
    '功能：对病人过敏史/病生状态进行管理
    'Pass
    Call AdviceCheckWarn(21)
End Sub

Private Sub cmdFind_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

Private Sub cmdIC_Click()
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        txtFind.Text = mobjICCard.Read_Card()
        If txtFind.Text <> "" Then Call txtFind_KeyPress(vbKeyReturn)
    End If
End Sub
Private Sub CmdSend_Click()
    Dim lngRow As Long, lng药品ID As Long, LngID As Long, lng批次 As Long, lng分批 As Long
    Dim blnInput As Boolean, strShow As String, strReturn As String, str操作员 As String, strTmp As String
    Dim rsTemp As New ADODB.Recordset, blnInTrans As Boolean
    Dim intUnit As Integer
    Dim bln是否有退药 As Boolean
    Dim str序号串 As String
    Dim n As Integer
    Dim BlnFirst As Boolean
    Dim strSignInfo As String
    
    blnInTrans = False
    On Error Resume Next
    
    mstr毒麻类提示 = ""
    
    err = 0
    
    If TxtNo.ListIndex = -1 Then   '无效属性
        MsgBox "请先选择处方！", vbInformation, gstrSysName
        If TxtNo.Enabled Then TxtNo.SetFocus
        Exit Sub
    End If
    
    
    On Error GoTo ErrHand
    
    '检查单据是否存在
    If Not CheckBillExist(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lng药房ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    '--配药操作--wq
    If MnuEditDosage.Checked Then
        '如果不需经过配药过程，本操作相当于发药
        If Not IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            GoTo SendBill
        End If
        
        '检测是否允许
        If CheckBill(1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        '校验配药人，如果启用电子签名则不使用
        If gbln药品使用电子签名 = False Then
            If int校验配药人 = 1 Then
                str操作员 = zlDatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
            Else
                str操作员 = Str配药人
            End If
            If str操作员 = "" Then Exit Sub
        End If
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '先更新批次
        For lngRow = 1 To Bill处方明细.Rows - 2
            LngID = Val(Bill处方明细.TextMatrix(lngRow, 列名.Id))
            lng药品ID = Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID))
            lng批次 = Val(Bill处方明细.TextMatrix(lngRow, 列名.批次))
            gstrSQL = "zl_药品收发记录_更新批次(" & LngID & "," & lng药品ID & "," & lng批次 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新批次")
        Next
        
        '再设置配药人
        gstrSQL = "zl_药品收发记录_设置配药人(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo, 1, 8) & "','" & IIf(gbln药品使用电子签名 = True, gstrUserName, IIf(int校验配药人 = 1, str操作员, IIf(Str配药人 = "|当前操作员|", gstrUserName, str操作员))) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-设置配药人")
        
        '如果已启用了电子签名，则需要对配药人进行电子签名处理
        If gbln药品使用电子签名 = True Then
            If SaveSignatureRecored(EsignTache.Dosage, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lng药房ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gcnOracle.CommitTrans
        blnInTrans = False
    End If
    '--取消操作--
    If MnuEditAbolish.Checked Then
        If Not IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            MsgBox "不需经过配药过程，因此不允许执行取消配药操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检测是否允许
        If CheckBill(2, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '如果已启用了电子签名，则取消配药人电子签名
        If gbln药品使用电子签名 = True Then
            If DelSignatureRecored(EsignTache.Dosage, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lng药房ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gstrSQL = "zl_药品收发记录_设置配药人(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo, 1, 8) & "',Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-设置配药人")
        
        gcnOracle.CommitTrans
        blnInTrans = False
    End If
    
    '--发药操作--
    If MnuEditConsignment.Checked Then
SendBill:
        '过滤状态时为批量发药模式
        If imgFilter.BorderStyle = cstFilter Then
            '批量检查处方
            If Not CheckBatchRecipe Then Exit Sub
            
            gcnOracle.BeginTrans
            blnInTrans = True
        
            '批量处方发药
            If Not SendBatchRecipe Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        Else
            '检查处方
            If Not CheckRecipe Then Exit Sub
            
            gcnOracle.BeginTrans
            blnInTrans = True
            
            '处方发药
            If Not SendRecipe Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gcnOracle.CommitTrans
        
        blnInTrans = False
        mblnFilterRefresh = True
        
        '打印处方
        Call PrintRecipe
    End If
    
    '--退药操作--
    If MnuEditHandback.Checked Then
        Dim str日期 As String, sig退药数 As Single, strSubSql As String
        '已转出的数据不允许操作
        If zlDatabase.NOMoved("药品收发记录", Mid(TxtNo.Text, 1, 8), "单据 = ", TxtNo.ItemData(TxtNo.ListIndex)) Then
            MsgBox "该处方已被转出，不允许进行退药操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        '检测是否允许
        If CheckBill(4, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), True) <> 0 Then Exit Sub
        Call GetBillSequence
        If str序号 = "" Then Exit Sub
        If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str序号) Then Exit Sub
        If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Sub
        If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf列表.TextMatrix(Msf列表.Row, 处方列名.金额)) Then Exit Sub
        '下面被注注释于20020905 Modified by zyb
        'If ReadBillData(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) = False Then Exit Sub
        If MsgBox("你确定单号为[" & TxtNo & "]" & "的处方退药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        For lngRow = 1 To Bill处方明细.Rows - 2
            lng分批 = Val(Bill处方明细.TextMatrix(lngRow, 列名.分批))
            lng批次 = Val(Bill处方明细.TextMatrix(lngRow, 列名.批次))
            '如果原来不分批而现在分批
            If lng批次 = 0 And lng分批 = 1 Then
                '如果批号或效期为空，则提取供用户输入
                blnInput = IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新批号)) = "", True, False)
                If blnInput Then
                    strShow = Txt科室.Text & "|" & Txt床号.Text & "|" & Msf列表.TextMatrix(lngRow, 处方列名.姓名) & _
                    "|" & Bill处方明细.TextMatrix(lngRow, 列名.药品名称) & "|" & Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID))
                    strReturn = Frm退药设置.ShowME(Me, strShow)
                    If strReturn = "" Then Exit Sub
                    '更新批号、效期及产地
                    Bill处方明细.TextMatrix(lngRow, 列名.新批号) = Split(strReturn, "|")(0)
                    Bill处方明细.TextMatrix(lngRow, 列名.新效期) = Split(strReturn, "|")(1)
                    Bill处方明细.TextMatrix(lngRow, 列名.新产地) = Split(strReturn, "|")(2)
                End If
            End If
        Next
        str日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        bln是否有退药 = False
        gcnOracle.BeginTrans
        blnInTrans = True
        For lngRow = 1 To Bill处方明细.Rows - 2
            'modified.by.zyb 门诊单位与住院单位不一致时，退药未退完 2003-01-10
            Select Case strUnit
            Case "售价单位"
                strSubSql = "*1"
            Case "门诊单位"
                strSubSql = "*Decode(门诊包装,Null,1,0,1,门诊包装)"
            Case "住院单位"
                strSubSql = "*Decode(住院包装,Null,1,0,1,住院包装)"
            Case "药库单位"
                strSubSql = "*Decode(药库包装,Null,1,0,1,药库包装)"
            End Select
            sig退药数 = Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数))

            gstrSQL = " Select round(" & sig退药数 & strSubSql & ",5) 数量 From 药品规格" & _
                         " Where 药品ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill处方明细.TextMatrix(lngRow, 列名.药品ID)))
                         
            With rsTemp
                sig退药数 = !数量
            End With
            
            If mbln显示大小单位 = True Then
                If (Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数大)) = Val(Bill处方明细.TextMatrix(lngRow, 列名.准退数大)) And _
                    Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数小)) = Val(Bill处方明细.TextMatrix(lngRow, 列名.准退数小))) Or _
                    (Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数)) = Val(Bill处方明细.TextMatrix(lngRow, 列名.准退数大)) * Val(Bill处方明细.TextMatrix(lngRow, 列名.包装)) + Val(Bill处方明细.TextMatrix(lngRow, 列名.准退数小))) Then
                    
                    sig退药数 = Val(Bill处方明细.TextMatrix(lngRow, 列名.实际数量))
                End If
            Else
                If Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数)) = Val(Bill处方明细.TextMatrix(lngRow, 列名.准退数)) Then
                    sig退药数 = Val(Bill处方明细.TextMatrix(lngRow, 列名.实际数量))
                End If
            End If
            
            If sig退药数 <> 0 Then
                '检查价格
                If CheckPrice(Val(Bill处方明细.TextMatrix(lngRow, 列名.Id)), mstr价格失效提示) = False Then
                    If MsgBox("药品[" & Bill处方明细.TextMatrix(lngRow, 列名.药品名称) & "]" & mstr价格失效提示, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gstrSQL = "zl_药品收发记录_部门退药(" & Val(Bill处方明细.TextMatrix(lngRow, 列名.Id)) & ",'" & gstrUserName & "'," & _
                            "to_date('" & str日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
                            IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新批号)) = "", "NULL", "'" & Bill处方明细.TextMatrix(lngRow, 列名.新批号) & "'") & "," & _
                            "" & IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新效期)) = "", "NULL", "to_date('" & Bill处方明细.TextMatrix(lngRow, 列名.新效期) & "','yyyy-MM-dd')") & "," & _
                            IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新产地)) = "", "NULL", "'" & Trim(Bill处方明细.TextMatrix(lngRow, 列名.新产地)) & "'") & "," & _
                            sig退药数 & ",NULL,NULL," & int金额保留位数 & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药")
                        bln是否有退药 = True
                    End If
                Else
                    gstrSQL = "zl_药品收发记录_部门退药(" & Val(Bill处方明细.TextMatrix(lngRow, 列名.Id)) & ",'" & gstrUserName & "'," & _
                        "to_date('" & str日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
                        IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新批号)) = "", "NULL", "'" & Bill处方明细.TextMatrix(lngRow, 列名.新批号) & "'") & "," & _
                        "" & IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新效期)) = "", "NULL", "to_date('" & Bill处方明细.TextMatrix(lngRow, 列名.新效期) & "','yyyy-MM-dd')") & "," & _
                        IIf(Trim(Bill处方明细.TextMatrix(lngRow, 列名.新产地)) = "", "NULL", "'" & Trim(Bill处方明细.TextMatrix(lngRow, 列名.新产地)) & "'") & "," & _
                        sig退药数 & ",NULL,NULL," & int金额保留位数 & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药")
                    bln是否有退药 = True
                End If
            End If
        Next
        
        '如果本地参数设置了自动销帐，并且当前退费单据是记帐单，那么执行门诊/住院销帐
        If mint自动销帐 = 1 And mint记录性质 = 2 And bln是否有退药 = True Then
            For lngRow = 1 To Bill处方明细.Rows - 2
                If Val(Bill处方明细.TextMatrix(lngRow, 列名.退药数)) <> 0 Then
                    str序号串 = str序号串 & IIf(str序号串 = "", Bill处方明细.TextMatrix(lngRow, 列名.序号), "," & Bill处方明细.TextMatrix(lngRow, 列名.序号))
                End If
            Next
            If mint门诊标志 = 1 Or mint门诊标志 = 4 Then
                gstrSQL = "Zl_门诊记帐记录_Delete('" & mstrNo & "','" & str序号串 & "','" & gstrUserCode & "','" & gstrUserName & "')"
            Else
                gstrSQL = "Zl_住院记帐记录_Delete('" & mstrNo & "','" & str序号串 & "','" & gstrUserCode & "','" & gstrUserName & "'," & mint记录性质 & ")"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-退药销帐")
        End If
        
        gcnOracle.CommitTrans
        blnInTrans = False
        
        '打印退费通知单
        Dim int单据 As Integer, strNo As String
        Dim Str发药时间 As String, Int包装系数 As Integer
        
        If bln是否有退药 Then
            int单据 = TxtNo.ItemData(TxtNo.ListIndex)
            strNo = Mid(TxtNo.Text, 1, 8)
            Str发药时间 = str日期
            Int包装系数 = IIf(int单据 = 8, 1, 2)
            
            If MsgBox("你需要打印退药通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
                Me, "No=" & strNo, "单据=" & int单据, "包装系数=" & IIf(Int包装系数 = 1, "D.门诊包装", "D.住院包装"), "退药时间=" & Str发药时间, 2)
            End If
            
            '提示停用药品
            Call CheckStopMedi(int单据 & "|" & strNo)
        Else
            MsgBox "本次没有退药。"
        End If
    End If
    
    BlnInOper = False
    Call mnuViewRefresh_Click
    
    If txtFind.Text <> "" Then
        txtFind.SetFocus
        Call GetFocus(txtFind)
    End If
    Exit Sub
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
    ByVal cur当日金额 As Currency, ByVal Cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
    intWarn As Integer) As Integer
'功能:对病人记帐进行报警提示
'参数:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
'     str收费类别=当前要检查的类别,用于分类报警
'     str类别名称=类别名称,用于提示
'     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
'返回:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
'     intWarn=本次询问性提示中的选择结果,0=为否,1-为是
'     0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
    Dim bln已报警 As Boolean, byt标志 As Byte
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim ArrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str担保 As String, i As Long
    
    BillingWarn = 0
    
    '报警参数检查:NULL是没有设置,0是设置了的
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    '示例："-" 或 ",ABC,567,DEF"
    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
    
    If bln已报警 Then '当intWarn = -1时,也可强行再报警
        If byt标志 = 2 Then
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                ArrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(ArrTmp)
                    If InStr(ArrTmp(i), str收费类别) > 0 Then
                        byt已报方式 = IIf(Right(ArrTmp(i), 1) = "②", 2, 1)
                        'Exit For '取消说明见住院记帐模块
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
    cur剩余款额 = cur剩余款额 + cur担保金额 - Cur记帐金额
    cur当日金额 = cur当日金额 + Cur记帐金额
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur剩余款额 < 0 Then
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf cur剩余款额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur剩余款额 < 0 Then
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";强制记帐;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gtype_UserSysParms.P9_费用金额保留位数) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
            End If
        End If
    End If
End Function

Private Function FinishBillingWarn(ByVal rsTmp As ADODB.Recordset, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
'功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
'参数：objRecord=包含要完成执行的病人信息的数据行
'      str类别="CDE..."，报警金额涉及到的收费类别
'      str类别名="检查,检验,..."，对应的类别名用于提示
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur当日 As Currency
    
    On Error GoTo errH
    
    If rsTmp!来源.Value = "住院" Then
        '住院病人报警
        strSQL = _
            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
            " Union ALL" & _
            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
        strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
        
        strSQL = "Select zl_PatiWarnScheme(A.病人ID,B.主页ID) As 适用病人,C.剩余款," & _
            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
            " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
            " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!病人ID), Val(rsTmp!主页id))
    Else
        '其他按门诊报警
        strSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
        strSQL = "Select zl_PatiWarnScheme(A.病人ID) As 适用病人,A.担保额," & _
            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0)+Nvl(E.帐户余额,0) as 剩余款" & _
            " From 病人信息 A,(" & strSQL & ") B,医保病人关联表 D,医保病人档案 E" & _
            " Where A.病人ID=B.病人ID(+) " & _
            " And A.病人id = D.病人id(+) And A.险类=D.险类(+) And D.险类=E.险类(+) And D.医保号=E.医保号(+) And D.标志(+)=1" & _
            " And A.病人ID=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!病人ID))
    End If
    
    intWarn = -1 '记帐报警时缺省要提示
    '执行报警:门诊病人病区ID=0
    strSQL = "Select Nvl(报警方法,1) as 报警方法," & _
        " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
        " Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!病人病区ID), CStr(NVL(rsPati!适用病人)))
    If Not rsWarn.EOF Then
        If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(Val(rsTmp!病人ID))
        str类别名 = Mid(str类别名, 2)
        For i = 1 To Len(str类别)
            intR = BillingWarn(Me, mstrPrivs, rsWarn, rsTmp!姓名, NVL(rsPati!剩余款, 0), cur当日, cur金额, NVL(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim BlnFirst As Boolean
    
    If KeyCode = vbKeyF2 Then
        If CmdSend.Enabled And CmdSend.Visible Then CmdSend_Click
    End If
    
    If KeyCode = vbKeyF3 Then
        If imgFilter.BorderStyle = cstLocate Then
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call txtFind_Validate(False)
                Call zlControl.TxtSelAll(txtFind)
                Call FindNextPati(txtFind.Tag <> txtFind.Text)
            End If
        Else
            Call SetFilter(MnuEditHandback.Checked)
        End If
    End If
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtFind.SetFocus
        End If
    End If
    
    'Ctrl+F4  读IC卡
    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then
            If lblFind.Tag = FindType.IC卡 Then
                Call cmdIC_Click
            End If
        End If
    End If
End Sub

Private Sub FindNextPati(ByVal BlnFirst As Boolean)
    Static intStar As Integer
    Dim n As Integer
    Dim strFind As String
    Dim blnDo As Boolean
    
    If BlnFirst Then intStar = 1
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    strFind = Trim(txtFind.Text)
    
    With Msf列表
        If .Rows < 2 Then Exit Sub
        
        For n = intStar To .Rows - 1
            Select Case lblFind.Tag
                Case FindType.就诊卡
                    If Trim(.TextMatrix(n, 处方列名.就诊卡号)) = strFind Then blnDo = True
                Case FindType.门诊号
                    If Trim(.TextMatrix(n, 处方列名.门诊号)) = strFind Then blnDo = True
                Case FindType.单据号
                    If Trim(.TextMatrix(n, 处方列名.NO)) = strFind Then blnDo = True
                Case FindType.姓名
                    If mblnCard = True Then
                        If Trim(.TextMatrix(n, 处方列名.就诊卡号)) = strFind Then blnDo = True
                    Else
                        If gbytCode = 1 Then
                            If Trim(.TextMatrix(n, 处方列名.姓名)) Like "*" & strFind & "*" Or mWBX(Trim(.TextMatrix(n, 处方列名.姓名)), 1) Like "*" & UCase(strFind) & "*" Then blnDo = True
                        Else
                            If Trim(.TextMatrix(n, 处方列名.姓名)) Like "*" & strFind & "*" Or mPinYin(Trim(.TextMatrix(n, 处方列名.姓名))) Like "*" & UCase(strFind) & "*" Then blnDo = True
                        End If
                    End If
                Case FindType.身份证
                    If Trim(.TextMatrix(n, 处方列名.身份证)) = strFind Then blnDo = True
                Case FindType.IC卡
                    If Trim(.TextMatrix(n, 处方列名.IC卡)) = strFind Then blnDo = True
            End Select
            
            If blnDo Then
                txtFind.Tag = txtFind.Text
                .Row = n
                Call Msf列表_EnterCell
                .TopRow = n
                intStar = n + 1
                If intStar > .Rows - 1 Then intStar = .Rows - 1
                Exit Sub
            End If
        Next
    End With
    intStar = 1
    txtFind.SetFocus
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    
    zlDatabase.SetPara "显示病区处方", img病区.BorderStyle, glngSys, 1341
    
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & Me.Name, "界面定位", imgFilter.BorderStyle)
    Call SaveSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & Me.Name, "显示退药待发单据", Chk显示退药待发单据.Value)
    
    '保存排序串
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未配药处方排序串", strOrder_1)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已配药处方排序串", strOrder_2)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未发药处方排序串", strOrder_3)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已发药处方排序串", strOrder_4)
    
    '保存输入模式
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "输入模式", mint输入模式)
        
    If Not InDesign And glngOld > 0 Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    Call SaveFlexState(Bill处方明细.MsfObj, Me.Name & "\" & tabShow.Tab)
    Call SaveFlexState(Msf列表, Me.Name & "\" & tabShow.Tab)
    SaveWinState Me, App.ProductName
    
    '卸载身份证刷卡接口
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    
    '卸载IC卡刷卡接口
    Set mobjICCard = Nothing
    
    '卸载电子签名接口
    Set gobjESign = Nothing
End Sub





Private Sub imgFilter_Click()
    imgFilter.BorderStyle = Abs(imgFilter.BorderStyle - 1)
    If imgFilter.BorderStyle = cstFilter Then
        Msf列表.ColWidth(处方列名.选择) = IIf(MnuEditConsignment.Checked, 300, 0)
    Else
        Msf列表.ColWidth(处方列名.选择) = 0
    End If
    
    txtFind.Text = ""
    mstrFilter = ""
    Call mnuViewRefresh_Click
End Sub

Private Sub img病区_Click()
    With img病区
        .BorderStyle = Abs(.BorderStyle - 1)
    End With
    Call SetPosition
    Call mnuViewRefresh_Click
End Sub


Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuViewLocate, 2, fraFind.Left + lblFind.Left - 30, fraFind.Top + lblFind.Top + lblFind.Height + 30
    End If
End Sub


Private Sub mnuCancel_Click()
    Dim blnInTrans As Boolean
        
    On Error GoTo errHandle
    
    If mstrNo <> "" Or IntBillStyle <> 0 Then
        If CheckBill(5, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '如果已启用了电子签名，则取消配药人电子签名
        If gbln药品使用电子签名 = True Then
            If DelSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lng药房ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gstrSQL = "Zl_药品收发记录_取消发药(" & lng药房ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-取消发药")
        
        gcnOracle.CommitTrans
        blnInTrans = False
        
        mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuChange_Click()
    Dim strName As String
    
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    
    strName = zlDatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
    
    TimeRefresh.Enabled = True
    TimePrint.Enabled = True
    
    If Trim(strName) = "" Then Exit Sub
    
    mstr自动配药人 = strName
    
    mdate上次校验时间 = zlDatabase.Currentdate

End Sub
Private Sub mnuCharge_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    blnOK = gobjCharge.Charge(Me, gcnOracle, glngSys, gstrDbUser, 1, 0)
    Call GlobalDeleteAtom(intAtom)
    
    '完成划价
    '刷新未发药处方
    mnuViewRefresh_Click
End Sub

Private Sub MnuEditSendOther_Click()
    With Frm药品批量发药
        .In_单据 = mInt单据
        .In_发药窗口 = Str窗口
        .In_药房ID = lng药房ID
        .In_库存检查 = IntCheckStock
        .In_校验处方 = intVerify
        .In_允许未配药发药 = IntSendAfterDosage
        .IN_允许未审核发药 = Int允许未审核处方发药
        .IN_允许未收费发药 = mint允许未收费处方发药
        .In_权限 = mstrPrivs
        .Str配药人 = IIf(Str配药人 = "|当前操作员|", gstrUserName, Str配药人)
        .In_金额保留位数 = int金额保留位数
        .IN_审核划价单 = int审核划价单
        .In_发其他药房处方 = True
        .Show 1, Me
    End With
    mnuViewRefresh_Click
End Sub

Private Sub mnuFileBack_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "药房=" & lng药房ID)
End Sub

Private Sub mnuFileLable_Click()
    Dim int单据 As Integer, strNo As String
    
    If Trim(Msf列表.TextMatrix(Msf列表.Row, 处方列名.类型)) = "" Then Exit Sub
    
    int单据 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据))
    strNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
    
    '检查单据是否存在
    If Not CheckBillExist(int单据, strNo) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lng药房ID, int单据, strNo)

    If Not BillHaveHerial(strNo, int单据) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "性质=" & IIf(int单据 = 8, 1, 2), "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "药房=" & lng药房ID, 2)
    End If
End Sub
Private Sub mnuStuff_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    
    If Msf列表.Rows = 1 Or Msf列表.TextMatrix(Msf列表.Rows - 1, 处方列名.NO) = "" Then
        mstrNo = ""
        lng病人ID = 0
    End If
    
    If mstrNo <> "" Or IntBillStyle <> 0 Then
        gstrSQL = "Select Nvl(病人id,0) 病人ID From 病人费用记录 Where Id=(Select 费用id From 药品收发记录 Where 单据=[1] And No=[2] And Rownum=1)"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取病人ID]", IntBillStyle, mstrNo)
        If Not rsTmp.EOF Then
            lng病人ID = rsTmp!病人ID
        End If
        rsTmp.Close
    End If
            
    On Error Resume Next
    If gobjStuff Is Nothing Then
        Set gobjStuff = CreateObject("zl9Stuff.clsStuff")
        If gobjStuff Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    Call gobjStuff.TransStuff(Me, gcnOracle, glngSys, gstrDbUser, lng病人ID, mstrNo, lng药房ID, mstrStartDate, mstrEndDate)
    Call GlobalDeleteAtom(intAtom)

End Sub
Private Sub MnuEditAbolish_Click()
    '--表示配药--
    If MnuEditAbolish.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = True
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditBill_Click()
    With Frm按票据号批量发药
        .In_单据 = mInt单据
        .In_发药窗口 = Str窗口
        .In_药房ID = lng药房ID
        .In_库存检查 = IntCheckStock
        .In_校验处方 = intVerify
        .In_允许未配药发药 = IntSendAfterDosage
        .IN_允许未审核发药 = Int允许未审核处方发药
        .IN_允许未收费发药 = mint允许未收费处方发药
        .In_权限 = mstrPrivs
        .Str配药人 = IIf(Str配药人 = "|当前操作员|", gstrUserName, Str配药人)
        .In_金额保留位数 = int金额保留位数
        .IN_审核划价单 = int审核划价单
        .Show 1, Me
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditBillRestore_Click()
    frm按票据号批量退药.In_权限 = mstrPrivs
    If Not frm按票据号批量退药.ShowEditor(Me, lng药房ID, int金额保留位数) Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub MnuEditConsignment_Click()
    If MnuEditConsignment.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = True
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditDosage_Click()
    '--表示配药--
    If MnuEditDosage.Checked = True Then Exit Sub
    MnuEditDosage.Checked = True
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditHandback_Click()
    If MnuEditHandback.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = True
    
    SetButtonState
End Sub

Private Sub mnuEditHandbackBatch_Click()
    frm批量退药.In_权限 = mstrPrivs
    If Not frm批量退药.ShowEditor(Me, lng药房ID, True, int金额保留位数) Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub MnuFileBillprint_Click()
    Dim int单据 As Integer, strNo As String
    
    If Trim(Msf列表.TextMatrix(Msf列表.Row, 处方列名.类型)) = "" Then Exit Sub
    
    int单据 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据))
    strNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
    
    '检查单据是否存在
    If Not CheckBillExist(int单据, strNo) Then
        MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lng药房ID, int单据, strNo)

    If Not BillHaveHerial(strNo, int单据) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "性质=" & IIf(int单据 = 8, 1, 2), "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=2", "PrintEmpty=0", 1)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "性质=" & IIf(int单据 = 8, 1, 2), "ReportFormat=2", "PrintEmpty=0", 1)
    End If
End Sub
Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub MnuFilePara_Click()
    BlnSetParaSuccess = False
    BlnRefresh = False
    
    '关闭Timer
    Call SetTimerState(False)
    
    With Frm发药参数设置
        Set .RecPart = RecPart.Clone
        .mstrPrivs = mstrPrivs
        .Show 1, Me
    End With
    
    If Not BlnSetParaSuccess Then
        '参数无变化时
    
        '开启Timer
        Call SetTimerState(True)
    Else
        '参数有变化时
        Call ReadFromReg
        
        '设置时间控件
        If mlngRefresh > 0 Then
            If mlngRefresh > 60 Then
                mlngRefresh = 60
            End If
            With TimeRefresh
                .Enabled = True
                .Interval = mlngRefresh * 1000
            End With
        Else
            TimeRefresh.Enabled = False
        End If
        
        If mlngPrintInterval > 0 Then
            If mlngPrintInterval > 60 Then
                mlngPrintInterval = 60
            End If
            With TimePrint
                .Enabled = True
                .Interval = mlngPrintInterval * 1000
            End With
        Else
            TimePrint.Enabled = False
        End If
        
        IntTimes = 0
        
        If mIntPrintHandbackNO <> 0 Then
            With TimePrintCancelBill
                .Enabled = False
                .Enabled = True
            End With
        Else
            TimePrintCancelBill.Enabled = False
        End If
        
        If CheckAnother = False Then Exit Sub
        Call SetFormat(2, True)
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileReport_Click()
    Dim str药房 As String
    Dim rsPart As New ADODB.Recordset
    
    gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
    Set rsPart = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取当前药房的名称]", lng药房ID)
    
    str药房 = rsPart!名称
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "库房=" & str药房 & "|" & lng药房ID, "包装系数=" & IIf(mintUnit = mconint门诊单位, "D.门诊包装", "D.住院包装"))
End Sub
Private Sub MnuFileRePrint_Click()
    Dim strPrintNO As String, intBillType As Integer
    
    If Not MnuEditHandback.Checked Then
        If strBill = "" Then Exit Sub
        
        strPrintNO = Split(strBill, "|")(0)
        intBillType = Val(Split(strBill, "|")(1))
    Else
        If Trim(Msf列表.TextMatrix(Msf列表.Row, 处方列名.类型)) = "" Then Exit Sub
        
        intBillType = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据))
        strPrintNO = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
    End If
    
    strUnit = GetUnit(lng药房ID, intBillType, strPrintNO)
    
    If Not BillHaveHerial(strPrintNO, intBillType) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strPrintNO, _
            "性质=" & IIf(intBillType = 8, 1, 2), _
            "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
            "ReportFormat=1", "PrintEmpty=0", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strPrintNO, _
            "性质=" & IIf(intBillType = 8, 1, 2), _
            "ReportFormat=1", "PrintEmpty=0", 2)
    End If
End Sub

Private Sub mnuFileRestore_Click()
    '打印退费通知单
    Dim int单据 As Integer, strNo As String
    Dim Str发药时间 As String, Int包装系数 As Integer
    If Trim(Msf列表.TextMatrix(Msf列表.Row, 处方列名.类型)) = "" Then Exit Sub
    If Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) <> 3 Then Exit Sub
    
    int单据 = Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)
    strNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
    Str发药时间 = Msf列表.TextMatrix(Msf列表.Row, 处方列名.日期)
    strUnit = GetUnit(lng药房ID, int单据, strNo)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
    Me, "No=" & strNo, "单据=" & int单据, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "退药时间=" & Str发药时间, 2)
End Sub

Private Sub mnuFileset_Click()
    zlPrintSet
End Sub

Private Sub mnuFlag_Click()
    Dim frmFlag As New Frm不再发药处方标志
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    mnuViewRefresh_Click
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub MnuHelpWebM_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub MnuEditBatch_Click()
    With Frm药品批量发药
        .In_单据 = mInt单据
        .In_发药窗口 = Str窗口
        .In_药房ID = lng药房ID
        .In_库存检查 = IntCheckStock
        .In_校验处方 = intVerify
        .In_允许未配药发药 = IntSendAfterDosage
        .IN_允许未审核发药 = Int允许未审核处方发药
        .IN_允许未收费发药 = mint允许未收费处方发药
        .In_权限 = mstrPrivs
        .Str配药人 = IIf(Str配药人 = "|当前操作员|", gstrUserName, Str配药人)
        .In_金额保留位数 = int金额保留位数
        .IN_审核划价单 = int审核划价单
        .In_发其他药房处方 = False
        .Show 1, Me
    End With
    mnuViewRefresh_Click
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


Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：药品=药品id，药房=药房id，NO=处方NO，单据类型=药品收发记录.单据，病人ID=病人ID
    Dim lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    If Split(mnuReportItem(Index).Tag, ",")(1) = "ZL1_INSIDE_1341" Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1341", "ZL8_INSIDE_1341"), Me)
    Else
        If mstrNo <> "" Or IntBillStyle <> 0 Then
            strSQL = "Select Nvl(病人id, 0) 病人id From 病人费用记录 Where Id=(Select 费用id From 药品收发记录 Where 单据=[1] And No=[2] And Rownum=1)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[取病人ID]", IntBillStyle, mstrNo)
            If Not rsTmp.EOF Then
                lng病人ID = rsTmp!病人ID
            End If
            rsTmp.Close
        End If
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "药品=" & IIf(SQLCondition.lng药品ID = 0, "", SQLCondition.lng药品ID), _
            "药房=" & IIf(lng药房ID = 0, "", lng药房ID), _
            "NO=" & mstrNo, _
            "单据类型=" & IIf(IntBillStyle = 0, "", IntBillStyle), _
            "病人ID=" & IIf(lng病人ID = 0, "", lng病人ID))
    End If
End Sub



Private Sub MnuViewFind_Click()
    Dim strReturn As String, IntOper As Integer
    
    If MnuEditDosage.Checked Then
        IntOper = 1
    ElseIf MnuEditAbolish.Checked Then
        IntOper = 2
    ElseIf MnuEditConsignment.Checked Then
        IntOper = 3
    Else
        IntOper = 4
    End If
    
    With Frm药品发药查找
        strReturn = .ShowME(Me, lng药房ID, mInt单据, IntOper, mstrPrivs, mbln就诊卡, _
            SQLCondition.date开始日期, _
            SQLCondition.date结束日期, _
            SQLCondition.str开始NO, _
            SQLCondition.str结束NO, _
            SQLCondition.str姓名, _
            SQLCondition.str就诊卡, _
            SQLCondition.str标识号, _
            SQLCondition.lng科室ID, _
            SQLCondition.str填制人, _
            SQLCondition.str审核人, _
            SQLCondition.lng药品ID, _
            SQLCondition.str医保号, _
            mint离院带药)
        If strReturn = "" Then Exit Sub
    End With
    
    mstrStartDate = Format(SQLCondition.date开始日期, "yyyy-mm-dd hh:mm:ss")
    mstrEndDate = Format(SQLCondition.date结束日期, "yyyy-mm-dd hh:mm:ss")
    
    Select Case IntOper
    Case 1
        StrFind_1 = strReturn
    Case 2
        StrFind_2 = strReturn
    Case 3
        StrFind_3 = strReturn
    Case 4
        StrFind_4 = strReturn
    End Select
    
    If imgFilter.BorderStyle = cstFilter Then
        Call txtFind_KeyPress(13)
    Else
        Call mnuViewRefresh_Click
    End If
End Sub

Private Function DataRefresh() As Boolean
    Dim lngRow As Long, lngColor As Long     '循环查找变量
    Dim IntBillThis As Integer, StrNoThis As String
    Dim LngSelectRow As Long, intCol As Integer             '当前选择行
    Dim strCond As String
    Dim strSendType As String
    Dim str待发单据 As String
    Dim strCon病区 As String
    Dim bln医保号 As Boolean
    Dim strSqlCon医保号 As String
    
    '--根据当前状态刷新数据--
    On Error Resume Next
    err = 0
    DataRefresh = True
    If BlnInOper Then Exit Function
    If BlnInRefresh Then Exit Function

    Call zlCommFun.ShowFlash
    stbThis.Panels(2) = "正在刷新数据,请稍候..."
    
    '清除控件原内容
    ClearCons
    BlnInRefresh = True
    DataRefresh = False
    Chk全退.Enabled = False
    
    If imgFilter.BorderStyle = cstFilter And Trim(txtFind.Text) = "" Then
        mstrFilter = " And 1 = 2 "
    End If
    
    strCon病区 = ""
    If mbln发病区处方 Then
        If img病区.BorderStyle = 0 Then
            '不显示病区处方
            strCon病区 = " And (D.门诊标志 <> 2 Or (D.门诊标志 = 2 And D.病人病区id <> D.开单部门id)) "
        End If
        If img病区.BorderStyle = 1 And cbo病区.ListIndex <> -1 Then
            '要显示病区处方，并且病人病区等于当前选择的病区
            strCon病区 = " And D.病人病区id = " & cbo病区.ItemData(cbo病区.ListIndex)
        End If
    End If
    
    If mInt单据 = 0 Then
        strCond = " And A.单据 In (8,9)" '门诊及住院所有单据
    Else
        If mInt单据 = 8 Then
            strCond = " And A.单据 In (8,9) And A.主页ID Is NULL " '门诊划价及门诊记帐
        Else
            strCond = " And A.单据 IN (8,9) And A.主页ID Is Not NULL " '住院记帐
        End If
    End If
    
    If mlng待发单据 = 0 Then
        str待发单据 = " And C.记录状态=1 "
    Else
        str待发单据 = " And MOD(C.记录状态,3)=1 "
    End If
    
    bln医保号 = (SQLCondition.str医保号 <> "")
    strSqlCon医保号 = IIf(bln医保号 = True, " And B.医保号=[17] ", "")
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
    If mint离院带药 = 0 Then
    ElseIf mint离院带药 = 1 Then
        strSendType = " And Not Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    ElseIf mint离院带药 = 2 Then
        strSendType = " And Ltrim(To_Char(Nvl(C.扣率,0),'00')) Like '_3'"
    End If
    
    Lbl配药人.Caption = "配药人"
    
    CmdSend.Visible = True
        
    '所有的查询都增加一个条件，排除已标记为不发药的记录  by lyq 20050416
    If MnuEditDosage.Checked Then
        '读取数据
        gstrSQL = " Select '' As 颜色, 处方类型,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作,说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID, 记录状态 As 未审核,Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额 " & _
                  " From (" & _
                  "     Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID, Decode(D.记录状态, 0, 1, 0) 记录状态 ,d.实收金额, A.处方类型 " & _
                  "     From " & _
                  "         (Select B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.住院号,A.优先级,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型,A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, Nvl(A.处方类型, 0) 处方类型 " & _
                  "         From 未发药品记录 A,病人信息 B" & _
                  "         Where A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & " ANd (A.库房ID=[13] " & IIf(Str窗口 = "", "", " And (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & " Or A.库房ID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                  "         And A.配药人 Is Null " & strSqlCon医保号 & " ) A,药品收发记录 C, 病人费用记录 D" & _
                  "     Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & str待发单据 & strSendType & " And (C.库房id=[13] Or C.库房id Is null) " & IIf(mstrSourceDep = "", "", " And C.对方部门id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_1 = "", " And A.填制日期 " & StrDate, StrFind_1) & mstrFilter & strCon病区 & ") A" & _
                  "     GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.记录状态, A.处方类型"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '配药
        
        With Msf列表
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
                If tabShow.Tab = 2 And mblnStarPass Then
                    cmdAlley.Visible = True
                End If
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
                If cmdAlley.Visible = True Then cmdAlley.Visible = False
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        CmdSend.Caption = "配药(&V)"
        
        If mint自动配药 = 1 Then
            CmdSend.Visible = False
            MnuEditDosage.Visible = False
        Else
            MnuEditDosage.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "配药"))
            CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "配药")
        End If
    End If
    If MnuEditAbolish.Checked Then
        gstrSQL = " Select '' As 颜色, 处方类型,'' As 选择,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作,说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID, 记录状态 As 未审核,Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额 " & _
                  " From (" & _
                  "     Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID, Decode(D.记录状态, 0, 1, 0) 记录状态 ,d.实收金额, A.处方类型 " & _
                  "     From " & _
                  "         (Select B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.住院号,A.优先级,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型,A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, Nvl(A.处方类型, 0) 处方类型 " & _
                  "         From 未发药品记录 A,病人信息 B" & _
                  "         Where A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & " ANd (A.库房ID=[13] " & IIf(Str窗口 = "", "", " And (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & " Or A.库房ID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                  "         And A.配药人 Is Not Null " & strSqlCon医保号 & ") A,药品收发记录 C, 病人费用记录 D" & _
                  "     Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & str待发单据 & strSendType & " And (C.库房id=[13] Or C.库房id Is null) " & IIf(mstrSourceDep = "", "", " And C.对方部门id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_2 = "", " And A.填制日期 " & StrDate, StrFind_2) & mstrFilter & strCon病区 & ") A" & _
                  "     GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.记录状态, A.处方类型"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '未发药品记录
        
        With Msf列表
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        CmdSend.Caption = "取消配药(&C)"
        CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "配药")
    End If
    If MnuEditConsignment.Checked Then
        gstrSQL = " Select '' As 颜色, 处方类型,'' As 选择,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,to_Char(Sum(Round(零售金额," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS 金额,日期,可操作,说明,就诊卡号,门诊号,身份证号,IC卡号,病人ID, 记录状态 As 未审核,Sum(Round(实收金额," & mintMoneyDigit & ")) 实收金额 " & _
                  " From (" & _
                  "     Select A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID, Decode(D.记录状态, 0, 1, 0) 记录状态,d.实收金额, A.处方类型 " & _
                  "     From " & _
                  "         (Select B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.住院号,A.优先级,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型,A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, Nvl(A.处方类型, 0) 处方类型 " & _
                  "         From 未发药品记录 A,病人信息 B" & _
                  "         Where A.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & " ANd (A.库房ID=[13] " & IIf(Str窗口 = "", "", " And (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & " Or A.库房ID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                        IIf(IntSendAfterDosage = 0, " And A.配药人 Is Not Null", "") & strSqlCon医保号 & _
                  "     ) A,药品收发记录 C, 病人费用记录 D" & _
                  "     Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & str待发单据 & strSendType & " And (C.库房id=[13] Or C.库房id Is null) " & IIf(mstrSourceDep = "", "", " And C.对方部门id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_3 = "", " And A.填制日期 " & StrDate, StrFind_3) & mstrFilter & strCon病区 & ") A" & _
                  "     GROUP BY A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.记录状态, A.记录状态, A.处方类型"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '读取所有未发药品记录
        
        With Msf列表
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
            Call SetCheckBox(-1)
        End With
        
        CmdSend.Caption = "发药(&S)"
        CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "发药")
    End If
    If MnuEditHandback.Checked Then
        strCon病区 = Replace(strCon病区, "D.", "H.")
    
        Lbl配药人.Caption = "发药人"
        strCond = Replace(strCond, "A.主页ID", "H.主页ID")
        
        Dim strCond1 As String, strCond2 As String, strTemp As String
        Dim intRight As Integer, intLeft As Integer
        '在嵌套查询中，没有连接病人费用记录表，而条件中存在姓名字段时，需去掉该条件，因它用到病人费用记录表
        strCond1 = ""
        StrFind_4 = UCase(StrFind_4)
        strCond2 = StrFind_4
        intLeft = InStr(1, strCond2, " AND UPPER(H.姓名)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, StrFind_4, " AND")
            strTemp = Mid(StrFind_4, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(StrFind_4, intRight)
            Else
                strCond1 = Mid(StrFind_4, intLeft)
                strCond2 = strTemp
            End If
        End If
        intLeft = InStr(1, strCond2, " AND UPPER(H.标识号)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, strCond2, " AND")
            strTemp = Mid(strCond2, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(strCond2, intRight)
            Else
                strCond1 = strCond1 & Mid(strCond2, intLeft)
                strCond2 = strTemp
            End If
        End If
        intLeft = InStr(1, strCond2, " AND UPPER(B.就诊卡号)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, strCond2, " AND")
            strTemp = Mid(strCond2, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(strCond2, intRight)
            Else
                strCond1 = strCond1 & Mid(strCond2, intLeft)
                strCond2 = strTemp
            End If
        End If
        
        '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药
        If mint离院带药 = 0 Then
        ElseIf mint离院带药 = 1 Then
            strSendType = " And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
        ElseIf mint离院带药 = 2 Then
            strSendType = " And Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
        End If
        
        
        '针对任何一张药品处方，不会存在一部分明细分别在线与后备中存在的情况，因此，可直接通过在线UNION后备的方式解决
        '由于病人费用记录在最外层，而且无主要条件，通过病人费用记录的在线与后备联接后，其效果是全表扫描，因此，只能通过整个在线SQL UNION 整个后备SQL的方式解决
        If Chk清单.Value = 0 Then
            gstrSQL = " SELECT DISTINCT '' As 颜色, '' As 处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') 类型,A.单据,1 已收费,A.审核人 配药人," & _
                     "      A.NO,H.姓名,trim(to_char(sum(A.零售金额),'" & mstrOracleMoneyForamt & "')) AS 金额,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,1 可操作,' ' 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,H.门诊标志, H.记录性质 " & _
                     " FROM " & _
                     "      (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                     "          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                     "          A.零售价,round(B.零售金额," & mintMoneyDigit & ") 零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID " & _
                     "      FROM" & _
                     "          (SELECT *" & _
                     "          FROM 药品收发记录 A" & _
                     "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                     "          AND A.库房ID+0=[13] " & strSendType & _
                     "      " & IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量,SUM(A.零售金额) 零售金额" & _
                     "          FROM 药品收发记录 A" & _
                     "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL" & strSendType & _
                     "          AND A.库房ID+0=[13] " & IIf(mstrSourceDep = "", "", " And A.对方部门id+0 in(" & mstrSourceDep & ") ") & _
                     "      " & IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
            gstrSQL = gstrSQL & _
                     "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0" & _
                     "     ) A,病人费用记录 H,病人信息 B" & _
                     " WHERE A.库房ID+0=[13] " & IIf(Str窗口 = "", "", " AND (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & _
                     " " & strCond & mstrShowSendedBill & strCond1 & mstrFilter & strCon病区 & _
                     " AND (A.记录状态=1 OR MOD(A.记录状态,3)=0) AND A.审核人 IS NOT NULL AND A.费用ID=H.ID AND A.实际数量<>0 AND H.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & strSqlCon医保号 & _
                     " GROUP BY Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐'),A.单据,1,A.审核人,A.NO,H.姓名,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID, H.门诊标志, H.记录性质 "
        Else
            gstrSQL = " SELECT DISTINCT '' As 颜色, '' As 处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') 类型,A.单据,1 已收费,A.审核人 配药人," & _
                     "      A.NO,H.姓名,trim(to_char(sum(A.零售金额),'" & mstrOracleMoneyForamt & "')) AS 金额,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,A.可操作," & _
                     "      DECODE(A.记录状态,1,'第1次发药',DECODE(MOD(A.记录状态,3),0,'第1次发药',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发药',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退药')) 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,H.门诊标志, H.记录性质 " & _
                     " FROM " & _
                     "      (SELECT * FROM" & _
                     "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                     "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                     "              A.零售价 , round(A.零售金额," & mintMoneyDigit & ") 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作 " & _
                     "          FROM" & _
                     "              (SELECT *" & _
                     "              FROM 药品收发记录 A" & _
                     "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                     "              AND A.库房ID+0=[13] " & strSendType & _
                     "          " & IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "              ) A," & _
                     "              (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                     "              FROM 药品收发记录 A" & _
                     "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL " & strSendType & _
                     "              AND A.库房ID+0=[13] " & IIf(mstrSourceDep = "", "", " And A.对方部门id+0 in(" & mstrSourceDep & ") ") & _
                     "          " & IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "              GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
            gstrSQL = gstrSQL & _
                     "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号)" & _
                     "          UNION" & _
                     "          SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                     "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态,A.发药窗口," & _
                     "          A.零售价 , round(A.零售金额," & mintMoneyDigit & ") 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                     "          DECODE(记录状态,1,1,DECODE(MOD(记录状态,3),0,1,MOD(记录状态,3)+1)) 可操作" & _
                     "          FROM 药品收发记录 A" & _
                     "          WHERE nvl(a.发药方式,-999)<>-1 and NOT (记录状态=1 OR MOD(记录状态,3)=0)" & IIf(mstrSourceDep = "", "", " And A.对方部门id+0 in(" & mstrSourceDep & ") ") & strSendType & _
                     "          " & IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2)
            gstrSQL = gstrSQL & _
                     "     ) A,病人费用记录 H,病人信息 B" & _
                     " WHERE A.库房ID+0=[13] " & IIf(Str窗口 = "", "", " AND (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & _
                     " " & strCond & mstrShowSendedBill & strCond1 & mstrFilter & strCon病区 & _
                     " AND A.审核人 IS NOT NULL AND A.费用ID=H.ID AND H.病人ID=B.病人ID" & IIf(bln医保号 = True, "", "(+)") & strSqlCon医保号 & _
                     " GROUP BY Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 8, '收费', 9, '记帐') ,A.单据,1,A.审核人," & _
                     "      A.NO,H.姓名,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),A.可操作," & _
                     "      DECODE(A.记录状态,1,'第1次发药',DECODE(MOD(A.记录状态,3),0,'第1次发药',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发药',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退药')),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID, H.门诊标志, H.记录性质 "
        End If
        
        Dim blnMoved As Boolean
        Dim str开始日期 As String, strSQL As String
       
        str开始日期 = Format(SQLCondition.date开始日期, "yyyy-mm-dd hh:mm:ss")
        
        '判断从开始日期后，是否存在转出的处方数据
        blnMoved = zlDatabase.DateMoved(str开始日期)
        
        '如果存在数据转出，则需要同时从后备表中提取数据
        If blnMoved Then
            strSQL = gstrSQL
            strSQL = Replace(strSQL, "药品收发记录", "H药品收发记录")
            strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
            gstrSQL = gstrSQL & " UNION ALL " & strSQL
        End If
        
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '读取所有未发药品记录
        
        With Msf列表
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        '上色
        Msf列表.Redraw = False
        For lngRow = 1 To Msf列表.Rows - 1
            Msf列表.Row = lngRow
            lngColor = IIf(Val(Msf列表.TextMatrix(lngRow, 处方列名.可操作)) = 1, glng正常, IIf(Val(Msf列表.TextMatrix(lngRow, 处方列名.可操作)) = 2, glng发药, glng退药))
            For intCol = 处方列名.选择 To Msf列表.Cols - 1
                Msf列表.Col = intCol
                Msf列表.CellForeColor = lngColor
            Next
        Next
        Msf列表.Redraw = True
        
        CmdSend.Caption = "退药(&R)"
        CmdSend.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1) And IsHavePrivs(mstrPrivs, "退药")
    End If
    
    Call SetFormat(2)
        
    '设置处方颜色
    Call SetRecipeColor
        
    '定位原来选择的处方，如果失败，则定位到第一行
    Msf列表.Row = ReLocateRow
    Msf列表_EnterCell
    '绑定记录集后，须重新调整按钮位置
    ResizePicClose
    Call zlCommFun.StopFlash
    
    BlnInRefresh = False
    DataRefresh = True
End Function

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub mnuViewFontSet_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSET(i).Checked = False
    Next
    Me.mnuViewFontSET(Index).Checked = True

    Select Case Index
    Case 0
        Me.Msf列表.Font.Size = 9
        Bill处方明细.Font.Size = 9
     Case 1
        Me.Msf列表.Font.Size = 11
        Bill处方明细.Font.Size = 11
    Case 2
        Me.Msf列表.Font.Size = 15
        Bill处方明细.Font.Size = 15
    End Select
    intFont = Index
    
    zlDatabase.SetPara "字体", Index, glngSys, 1341
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewLocateItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuViewLocateItem.UBound
        mnuViewLocateItem(i).Checked = i = Index
    Next
    strItem = Split(mnuViewLocateItem(Index).Caption, "(")(0)
    lblFind.Caption = strItem & "↓"
    lblFind.Tag = Index
    mint输入模式 = Index
    
    If Index <> FindType.IC卡 Then
        cmdIC.Visible = False
        imgFilter.Left = fraFind.Width - imgFilter.Width - 80
        txtFind.Width = imgFilter.Left - txtFind.Left - 80
    End If
    
    txtFind.Text = "": txtFind.Tag = ""
    txtFind.PasswordChar = ""
    txtFind.MaxLength = 0
    
    Select Case Index
        Case FindType.就诊卡
            If gtype_UserSysParms.P12_就诊卡是否密文显示 Then
                txtFind.PasswordChar = "*"
            End If
            txtFind.MaxLength = gtype_UserSysParms.P20_就诊卡号长度
        Case FindType.IC卡
            cmdIC.Visible = True
            cmdIC.Left = fraFind.Width - cmdIC.Width - 80
            imgFilter.Left = cmdIC.Left - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
    End Select
        
    If Visible Then txtFind.SetFocus
End Sub

Private Sub mnuViewRefresh_Click()
    If Not BlnStartUp Then Exit Sub

    StrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    StrDate = " Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_Date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss') "
    
    'Modified by ZYB 2002-11-19 保存用户设置
    Call SaveFlexState(Bill处方明细.MsfObj, Me.Name & "\" & tabShow.Tab)
    Call SaveFlexState(Msf列表, Me.Name & "\" & tabShow.Tab)
    '重新读取数据
    DoEvents
    Call DataRefresh
    DoEvents
    
    '恢复设置
    Call RestoreFlexState(Msf列表, Me.Name & "\" & tabShow.Tab)
    Call RestoreFlexState(Bill处方明细.MsfObj, Me.Name & "\" & tabShow.Tab)
    Bill处方明细.ColWidth(列名.审查结果) = IIf(Not mblnStarPass, 0, 240)
    Bill处方明细.ColWidth(列名.重量) = IIf(mbln显示重量 And mblnIs中药处方, 1200, 0)
    Call SetColHide
    
    If imgFilter.BorderStyle = cstFilter Then
        Msf列表.ColWidth(处方列名.选择) = IIf(MnuEditConsignment.Checked, 300, 0)
    Else
        Msf列表.ColWidth(处方列名.选择) = 0
    End If
    
    mblnFilterRefresh = False
End Sub

Private Sub MnuViewState_Click()
    MnuViewState.Checked = MnuViewState.Checked Xor True
    stbThis.Visible = MnuViewState.Checked
    Form_Resize
End Sub

Private Sub MnuViewToolS_Click()
    MnuViewToolS.Checked = MnuViewToolS.Checked Xor True
    Cbar.Visible = MnuViewToolS.Checked
    MnuViewToolT.Enabled = MnuViewToolS.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolT_Click()
    MnuViewToolT.Checked = MnuViewToolT.Checked Xor True
    If MnuViewToolT.Checked Then
        Tbar1.Buttons("Preview").Caption = "预览"
        Tbar1.Buttons("Print").Caption = "打印"
        Tbar1.Buttons("Find").Caption = "过滤"
        Tbar1.Buttons("Help").Caption = "帮助"
        Tbar1.Buttons("Exit").Caption = "退出"
    Else
        Tbar1.Buttons("Preview").Caption = ""
        Tbar1.Buttons("Print").Caption = ""
        Tbar1.Buttons("Find").Caption = ""
        Tbar1.Buttons("Help").Caption = ""
        Tbar1.Buttons("Exit").Caption = ""
    End If
    
    Cbar.Bands(1).MinHeight = Tbar1.Height
End Sub

Private Sub Msf列表_DblClick()
    Msf列表_KeyDown vbKeyReturn, 0
End Sub

Private Sub Msf列表_GotFocus()
    With Msf列表
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf列表_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)) = "" Then TxtNo.Clear: Exit Sub
    If KeyCode = vbKeyReturn Then TxtNo_Click
End Sub

Private Sub Msf列表_LostFocus()
    With Msf列表
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub SetFormat(ByVal IntStyle As Integer, Optional ByVal BlnSetHead As Boolean = True)
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    '--设置各列表控件的格式--
    
    Select Case IntStyle
    Case 1
        With Msf列表
            If BlnSetHead Then
                .Cols = IIf(MnuEditHandback.Checked, 处方列名.退药列数, 处方列名.发药列数)
                .TextMatrix(0, 处方列名.颜色) = "颜色"
                .TextMatrix(0, 处方列名.处方类型) = ""
                .TextMatrix(0, 处方列名.选择) = ""
                .TextMatrix(0, 处方列名.标志) = "0"
                .TextMatrix(0, 处方列名.类型) = "类型"
                .TextMatrix(0, 处方列名.单据) = "单据"
                .TextMatrix(0, 处方列名.收费) = "收费"
                .TextMatrix(0, 处方列名.配药人) = "配药人"
                .TextMatrix(0, 处方列名.NO) = "NO"
                .TextMatrix(0, 处方列名.姓名) = "姓名"
                .TextMatrix(0, 处方列名.金额) = "金额"
                .TextMatrix(0, 处方列名.日期) = "日期"
                .TextMatrix(0, 处方列名.可操作) = "可操作"
                .TextMatrix(0, 处方列名.说明) = "说明"
                .TextMatrix(0, 处方列名.就诊卡号) = "就诊卡号"
                .TextMatrix(0, 处方列名.门诊号) = "门诊号"
                .TextMatrix(0, 处方列名.身份证) = "身份证号"
                .TextMatrix(0, 处方列名.IC卡) = "IC卡号"
                .TextMatrix(0, 处方列名.病人ID) = "病人ID"
             End If
            .TextMatrix(0, 处方列名.选择) = ""
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(处方列名.颜色) = 500
                .ColWidth(处方列名.处方类型) = 0
                .ColWidth(处方列名.选择) = 300
                .ColWidth(处方列名.标志) = 0
                .ColWidth(处方列名.类型) = 1000
                .ColWidth(处方列名.单据) = 0
                .ColWidth(处方列名.收费) = 0
                .ColWidth(处方列名.配药人) = 0
                .ColWidth(处方列名.NO) = 800
                .ColWidth(处方列名.姓名) = 800
                .ColWidth(处方列名.金额) = 1200
                .ColWidth(处方列名.日期) = 1500
                .ColWidth(处方列名.可操作) = 0
                .ColWidth(处方列名.说明) = 1500
                .ColWidth(处方列名.就诊卡号) = 1000
                .ColWidth(处方列名.门诊号) = 1000
                .ColWidth(处方列名.身份证) = 1600
                .ColWidth(处方列名.IC卡) = 1600
                .ColWidth(处方列名.病人ID) = 0
                .Row = 1
            End If
            
            If imgFilter.BorderStyle = cstFilter Then
                .ColWidth(处方列名.选择) = IIf(MnuEditConsignment.Checked, 300, 0)
            Else
                .ColWidth(处方列名.选择) = 0
            End If
            
            .ColAlignment(处方列名.选择) = 4
            .ColAlignment(处方列名.金额) = 7
            .ColAlignment(处方列名.就诊卡号) = 7
            .ColAlignment(处方列名.门诊号) = 7
            .ColAlignment(处方列名.身份证) = 7
            .ColAlignment(处方列名.IC卡) = 7
            Call RestoreFlexState(Msf列表, Me.Name & "\" & tabShow.Tab)
            .ColWidth(处方列名.单据) = 0
            If MnuEditHandback.Checked Then
                .ColWidth(处方列名.颜色) = 0
                .ColWidth(处方列名.门诊标志) = 0
                .ColWidth(处方列名.记录性质) = 0
            Else
                .ColWidth(处方列名.颜色) = 500
                .ColWidth(处方列名.未审核) = 0
                .ColWidth(处方列名.实收金额) = 0
            End If
        End With
    Case 2
        With Bill处方明细
            .Active = False
            .Rows = 2
            .Cols = 列名.列数
            
            .TextMatrix(0, 列名.审查结果) = "警"
            .TextMatrix(0, 列名.顺序号) = "序号"
            .TextMatrix(0, 列名.药品名称) = "药品名称"
            .TextMatrix(0, 列名.其它名) = "其它名"
            .TextMatrix(0, 列名.英文名) = "英文名"
            .TextMatrix(0, 列名.序号) = "序号"
            .TextMatrix(0, 列名.规格) = "规格"
            .TextMatrix(0, 列名.批号) = "批号"
            .TextMatrix(0, 列名.Id) = "ID"
            .TextMatrix(0, 列名.药品ID) = "药品ID"
            .TextMatrix(0, 列名.批次) = "批次"
            .TextMatrix(0, 列名.单位) = "单位"
            .TextMatrix(0, 列名.单价) = "单价"
            .TextMatrix(0, 列名.付数) = "付数"
            .TextMatrix(0, 列名.数量) = "数量"
            .TextMatrix(0, 列名.金额) = "金额"
            .TextMatrix(0, 列名.重量) = "重量"
            .TextMatrix(0, 列名.单量) = "单量"
            .TextMatrix(0, 列名.用法) = "用法"
            .TextMatrix(0, 列名.频次) = "频次"
            .TextMatrix(0, 列名.医生嘱托) = "医生嘱托"
            .TextMatrix(0, 列名.已退数) = "已退数"
            .TextMatrix(0, 列名.准退数) = "准退数"
            .TextMatrix(0, 列名.准退数大) = "准退数大"
            .TextMatrix(0, 列名.准退数小) = "准退数小"
            .TextMatrix(0, 列名.退药数) = "退药数"
            .TextMatrix(0, 列名.退药数大) = "退药数(大包装)"
            .TextMatrix(0, 列名.单位大) = "单位"
            .TextMatrix(0, 列名.退药数小) = "退药数(小包装)"
            .TextMatrix(0, 列名.单位小) = "单位"
            .TextMatrix(0, 列名.库存数) = "库存数"
            .TextMatrix(0, 列名.货位) = "库房货位"
            .TextMatrix(0, 列名.分批) = "分批"
            .TextMatrix(0, 列名.新批号) = "新批号"
            .TextMatrix(0, 列名.新效期) = "新效期"
            .TextMatrix(0, 列名.新产地) = "新产地"
            .TextMatrix(0, 列名.备注) = "备注"
            .TextMatrix(0, 列名.医嘱id) = "医嘱ID"
            .TextMatrix(0, 列名.实际数量) = "实际数量"
            .TextMatrix(0, 列名.费别) = "费别"
            .TextMatrix(0, 列名.包装) = "包装"
            
            .ColWidth(列名.审查结果) = IIf(Not mblnStarPass, 0, 240)
            .ColWidth(列名.顺序号) = 450
            .ColWidth(列名.药品名称) = 2500
            .ColWidth(列名.其它名) = 2000
            .ColWidth(列名.英文名) = 2000
            .ColWidth(列名.序号) = 0
            .ColWidth(列名.规格) = 1500
            .ColWidth(列名.批号) = 1500
            .ColWidth(列名.Id) = 0
            .ColWidth(列名.药品ID) = 0
            .ColWidth(列名.批次) = 0
            .ColWidth(列名.单位) = IIf(mbln显示大小单位 = True, 0, 500)
            .ColWidth(列名.单价) = 1000
            .ColWidth(列名.付数) = IIf(IntShowCol = 1, 800, 0)
            .ColWidth(列名.数量) = 1200
            .ColWidth(列名.金额) = 1200
            .ColWidth(列名.重量) = 1200
            .ColWidth(列名.单量) = 1200
            .ColWidth(列名.用法) = 1500
            .ColWidth(列名.频次) = 1500
            .ColWidth(列名.医生嘱托) = IIf(MnuEditHandback.Checked, 0, 1500)
            .ColWidth(列名.库存数) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(列名.货位) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(列名.已退数) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(列名.准退数) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(列名.准退数大) = 0
            .ColWidth(列名.准退数小) = 0
            .ColWidth(列名.退药数) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = False, 1200, 0)
            .ColWidth(列名.退药数大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
            .ColWidth(列名.退药数小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
            .ColWidth(列名.单位大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
            .ColWidth(列名.单位小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
            
            .ColWidth(列名.分批) = 0
            .ColWidth(列名.新批号) = 0
            .ColWidth(列名.新效期) = 0
            .ColWidth(列名.新产地) = 0
            .ColWidth(列名.备注) = 1200
            .ColWidth(列名.医嘱id) = 0
            .ColWidth(列名.实际数量) = 0
            .ColWidth(列名.费别) = 1000
            .ColWidth(列名.包装) = 0
            
            .ColAlignment(0) = 1
            .ColAlignment(2) = 1
            .ColAlignment(列名.药品名称) = 1
            .ColAlignment(列名.其它名) = 1
            .ColAlignment(列名.英文名) = 1
            .ColAlignment(列名.规格) = 1
            .ColAlignment(列名.批号) = 1
            .ColAlignment(列名.单位) = 1
            .ColAlignment(列名.用法) = 1
            .ColAlignment(列名.备注) = 1
            
            'Modified by ZYB 2002-11-19 恢复用户设置
            Call RestoreFlexState(.MsfObj, Me.Name & "\" & tabShow.Tab)
            .ColWidth(列名.审查结果) = IIf(Not mblnStarPass, 0, 240)
            .ColWidth(列名.付数) = IIf(IntShowCol = 1, 800, 0)
            .ColWidth(列名.医生嘱托) = IIf(MnuEditHandback.Checked, 0, 1500)
            .ColWidth(列名.库存数) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(列名.货位) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(列名.已退数) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(列名.准退数) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(列名.退药数) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = False, 1200, 0)
            .ColWidth(列名.退药数大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
            .ColWidth(列名.退药数小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 1500, 0)
            .ColWidth(列名.单位大) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
            .ColWidth(列名.单位小) = IIf(MnuEditHandback.Checked And mbln显示大小单位 = True, 500, 0)
            .ColWidth(列名.顺序号) = 450
            .ColWidth(列名.序号) = 0
            .ColWidth(列名.实际数量) = 0
            .ColWidth(列名.包装) = 0
            .ColWidth(列名.单位) = IIf(mbln显示大小单位 = True, 0, 500)
            .ColWidth(列名.准退数大) = 0
            .ColWidth(列名.准退数小) = 0
        End With
    End Select
    
    Call SetColHide
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    Form_Resize
    BlnFirstStart = True
    
    If Me.Tag = "" Then
        Call tabShow_Click(tabShow.Tab)
        Me.Tag = "Refresh"
    End If
End Sub

Private Sub Form_Load()
    BlnEnterCell = False
    BlnStartUp = False
    
    lblUserName.Caption = gstrUserName
    lblUserName.Left = 0
    PicToolbar.Width = lblUserName.Width + 10
    PicToolbar.Height = Tbar1.Height
    lblUserName.Top = (PicToolbar.Height - lblUserName.Height) / 2 + 20
            
    cmdIC.Visible = False
          
    fraFind.Width = IIf(IntSendAfterDosage = 1, 3000, 3950)
    lblFind.Left = 90
    txtFind.Left = lblFind.Left + lblFind.Width + 80
    imgFilter.Left = fraFind.Width - imgFilter.Width - 80
    txtFind.Width = imgFilter.Left - txtFind.Left - 80
    
    '初始化窗体最大最小边界
    glngMinW = 9555
    glngMinH = 6675
    glngMaxW = Screen.Width
    glngMaxH = Screen.Height
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mint离院带药 = 0
    mblnIsFirst = True
    mdate上次校验时间 = zlDatabase.Currentdate
    mstr自动配药人 = ""
    
    strChargePrivs = GetPrivFunc(glngSys, 1120)
    strStuffPrivs = GetPrivFunc(glngSys, 1723)
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If

    If gstrUserName = "" Then
        MsgBox "请为当前用户设置对应的操作员后再使用本模块！", vbInformation, gstrSysName
        Exit Sub
    End If
   
    '取系统参数
    Call GetSysParms
    
    '取金额位数
    mintMoneyDigit = GetDigit(0, 1, 4)
    '设置金额格式
    Call GetMoneyFormat
    
    Call TradeName
    '为各控件装入图标
    If LoadInIcon = False Then Exit Sub
    '依赖数据检测
    If DependOnCheck = False Then Exit Sub
    '从注册表中取出用户设置
    Call ReadFromReg
    '检查相关设置
    If CheckAnother = False Then Exit Sub
    Call mnuViewFontSet_Click(intFont)
    Lbl标题.Caption = GetUnitName & Lbl标题.Caption
    
    Set mobjIDCard = New clsIDCard
    
    '电子签名接口控制
    If gbln药品使用电子签名 = True Then
        On Error Resume Next
        gbln药品使用电子签名 = False
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gbln药品使用电子签名 = False
            End If
        End If
        gbln药品使用电子签名 = True
    End If
    
    Call mnuViewLocateItem_Click(mint输入模式)
    
    '初始化相关
    StrLastNo = ""
    IntLastBill = 0
    strLastData = ""
    mintLastSequence = 1
    StrFindStyle = "%"
    LngSendRow = 0
    Int模式 = 1
    BlnFirstStart = False
    BlnInOper = False
    BlnAllowClick = True
    BlnInRefresh = False
    
    StrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    mstrStartDate = StrDate & " 00:00:00"
    mstrEndDate = StrDate & " 23:59:59"
    
    StrDate = " Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_Date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss') "
    
    StrFind_1 = "": StrFind_2 = "": StrFind_3 = "": StrFind_4 = ""
    
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    Call SetFormat(2, True)
    If IntSendAfterDosage = 0 Then
        If IsHavePrivs(mstrPrivs, "配药") Then
            MnuEditDosage_Click
        End If
    Else
        If IsHavePrivs(mstrPrivs, "发药") Then
            MnuEditConsignment_Click
        ElseIf IsHavePrivs(mstrPrivs, "退药") Then
            MnuEditHandback_Click
        Else
            MnuEditConsignment_Click
        End If
    End If
    
    If glngSys \ 100 = 1 Then
        Me.Caption = "药品处方发药"
    Else
        Me.Caption = "药店处方发药"
        Me.Lbl科室.Caption = "姓名"
        Me.Lbl床号.Visible = False
        Me.Txt床号.Visible = False
        Me.Txt住院号.Visible = False
        Me.Lbl住院号.Visible = False
    End If
    Call SetFormat(1, True)
    Call SetFormat(2, True)
    Call 权限控制
    
    Call mnuViewRefresh_Click
    Call RestoreWinState(Me, App.ProductName)
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1341")
    
    Call SetFormat(1, True)
    Call SetFormat(2, True)
    StrLastNo = ""
    Call Msf列表_EnterCell
    
    MnuEditDosage.Enabled = IsHavePrivs(mstrPrivs, "配药")
    MnuEditAbolish.Enabled = IsHavePrivs(mstrPrivs, "配药")
    MnuEditBatch.Enabled = IsHavePrivs(mstrPrivs, "发药")
    mnuEditBillRestore.Enabled = IsHavePrivs(mstrPrivs, "退药")

    '设置时间控件
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    If mlngRefresh > 0 Then
        If mlngRefresh > 60 Then
            mlngRefresh = 60
        End If
        With TimeRefresh
            .Enabled = True
            .Interval = mlngRefresh * 1000
        End With
    End If
    
    If mlngPrintInterval > 0 Then
        If mlngPrintInterval > 60 Then
            mlngPrintInterval = 60
        End If
        With TimePrint
            .Enabled = True
            .Interval = mlngPrintInterval * 1000
        End With
    End If
    IntTimes = 0
    If mIntPrintHandbackNO <> 0 Then
        With TimePrintCancelBill
            .Enabled = False
            .Enabled = True
        End With
    Else
        TimePrintCancelBill.Enabled = False
    End If
    
    BlnStartUp = True
    BlnEnterCell = True
    
    ImgLeftRight_S.Left = IIf(IntSendAfterDosage = 1, 3100, 4500)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    Dim DblWidth As Double, DblHeight As Double
    
    If IntSendAfterDosage = 0 Then
        If Me.Width < 13000 Then
            Me.Width = 13000
        End If
    Else
        If Me.Width < 11500 Then
            Me.Width = 11500
        End If
    End If
    If Me.Height < 8250 Then
        Me.Height = 8250
    End If
    
    tabShow.Width = IIf(IntSendAfterDosage = 1, 2500, 3930)
    Msf列表.ZOrder 0
    
    With ImgLeftRight_S
        If .Left < IIf(IntSendAfterDosage = 1, 3100, 4500) Then .Left = IIf(IntSendAfterDosage = 1, 3100, 4500)
        If .Left > IIf(IntSendAfterDosage = 1, 3950, 5500) Then .Left = IIf(IntSendAfterDosage = 1, 3950, 5500)
    End With
    
    PicToolbar.Top = Tbar1.Top
    PicToolbar.Left = Me.Width - PicToolbar.Width - 200
        
    With Cbar
        .Align = 1
        If BlnFirstStart = False Then
            'Set .Bands(1).Child = Tbar1
            .Bands(1).MinHeight = Tbar1.Height
        End If
    End With
    
    With fraFind
        .Top = IIf(Cbar.Visible, Cbar.Height, 0)
        .Left = 10
        
        .Width = ImgLeftRight_S.Left - .Left - 80
        
        If cmdIC.Visible = True Then
            cmdIC.Left = .Width - cmdIC.Width - 80
            imgFilter.Left = cmdIC.Left - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
        Else
            imgFilter.Left = .Width - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
        End If
    End With
    
    With ImgLeftRight_S
        .Top = fraFind.Top + fraFind.Height - 20
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        DblHeight = .Height
    End With
    
    With tabShow
        .Top = fraFind.Top + fraFind.Height + 50
        .Left = 0
    End With
    
    With Msf列表
        .Top = ImgLeftRight_S.Top + tabShow.Height
        .Height = ImgLeftRight_S.Height - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = ImgLeftRight_S.Left
        .Left = 0
    End With
    With PicCloseConsignment
        .Top = Msf列表.Top + 30
    End With
    
    '调整单据大小
    DblWidth = Me.ScaleWidth - (ImgLeftRight_S.Left + ImgLeftRight_S.Width)
    With PicBackGroud
        .Left = ImgLeftRight_S.Left + ImgLeftRight_S.Width
        .Top = ImgLeftRight_S.Top
        .Width = DblWidth
        .Height = ImgLeftRight_S.Height
        .ZOrder 0
    End With
    
    '几个界面切换按钮、下拉框、选择框的位置设置
    Call SetPosition
    
    With Lbl标题
        .Width = DblWidth
    End With
    With PicState
        .Left = DblWidth - .Width - 80
    End With
    With TxtNo
        .Left = DblWidth - .Width - 80
    End With
    With LblNo
        .Left = TxtNo.Left - 80 - .Width
    End With
    
    With Txt床号
        .Left = DblWidth - .Width - 80
    End With
    With Lbl床号
        .Left = Txt床号.Left - .Width - 80
    End With
    
    With Txt开单医生
        .Top = DblHeight - .Height - 100
    End With
    With Lbl开单医生
        .Top = Txt开单医生.Top + 60
    End With
    
    With Lbl配药人
        .Top = Lbl开单医生.Top
    End With
    With cbo配药人
        .Top = Txt开单医生.Top
    End With
    
    With Txt收费员
        .Top = Txt开单医生.Top
    End With
    With Lbl收费员
        .Top = Lbl开单医生.Top
    End With
    
    With CmdSend
        .Left = DblWidth - .Width - 50
        .Top = Txt开单医生.Top - 25
    End With
    
    With Chk全退
        .Top = Lbl开单医生.Top
        .Left = CmdSend.Left - Chk全退.Width - 150
    End With
    
    With txt原始付数
        .Top = DblHeight - .Height - 450
    End With
    
    With lbl原始付数
        .Top = txt原始付数.Top + 60
    End With
    
    With txt中药煎法
        .Top = txt原始付数.Top
    End With
        
    With lbl中药煎法
        .Top = lbl原始付数.Top
    End With

    If mblnIs中药处方 Then
        With Bill处方明细
            .Top = Txt床号.Top + Txt床号.Height + 50
            .Height = IIf(txt原始付数.Top - .Top - 50 < 0, .Height, txt原始付数.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    Else
        With Bill处方明细
            .Top = Txt床号.Top + Txt床号.Height + 50
            .Height = IIf(cbo配药人.Top - .Top - 50 < 0, .Height, cbo配药人.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    End If
    
    '调整单据头上的控件
    With Txt年龄
        If glngSys \ 100 = 1 Then
            .Left = DblWidth / 2 - .Width / 2
        Else
            .Left = DblWidth - .Width - 100
        End If
    End With
    With Lbl年龄
        .Left = Txt年龄.Left - .Width - 50
    End With
    With Txt性别
        If glngSys \ 100 = 1 Then
            .Left = DblWidth / 3 - .Width / 2
        Else
            .Left = DblWidth / 2 - .Width / 2
        End If
    End With
    With Lbl性别
        .Left = Txt性别.Left - .Width - 50
    End With
    With Txt住院号
        .Left = (Txt床号.Left - (Txt年龄.Left + Txt年龄.Width) / 2) + Txt住院号.Width / 2
    End With
    With Lbl住院号
        .Left = Txt住院号.Left - .Width
    End With
    
    ResizePicClose

End Sub

Private Sub ImgLeftRight_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight_S
        .Move .Left + x
    End With
    Form_Resize
End Sub
Private Function LoadInIcon() As Boolean
    '--为各控件装入图标--
    On Error Resume Next
    err = 0
    LoadInIcon = False
    
    '工具栏
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BCHARGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTUFF", vbResIcon)
    End With
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CCHARGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTUFF", vbResIcon)
    End With
    With Tbar1
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        
        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Cancel").Image = 3
        .Buttons("Find").Image = 7
        .Buttons("Help").Image = 8
        .Buttons("Exit").Image = 9
        .Buttons("Charge").Image = 10
        .Buttons("Stuff").Image = 11
    End With
    Cbar.Bands(1).MinHeight = Tbar1.Height
    
    RaisEffect PicCloseConsignment, 2
    
    If err <> 0 Then
        MsgBox "相关资源文件丢失，请与软件开发商联系！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function ReadData(ByVal StrQuery As String, Optional ByVal IntStyle As Integer = 0) As Boolean
    Dim strOrder As String
    '--读取数据，并按用户的要求进行排序--
    'IntStyle:0-未配处方;1-已配处方;2-未发处方;3-已发处方
    
    On Error Resume Next
    err = 0
    ReadData = False
    
    gstrSQL = StrQuery
    
    Set RecPhysic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date开始日期, _
            SQLCondition.date结束日期, _
            SQLCondition.str开始NO, _
            SQLCondition.str结束NO, _
            SQLCondition.str姓名, _
            SQLCondition.str就诊卡, _
            SQLCondition.str标识号, _
            SQLCondition.lng科室ID, _
            SQLCondition.str填制人, _
            SQLCondition.str审核人, _
            SQLCondition.lng药品ID, _
            SQLCondition.str当前NO, _
            lng药房ID, _
            SQLCondition.str门诊号, _
            SQLCondition.str身份证, _
            SQLCondition.strIC卡, _
            SQLCondition.str医保号)
            
   With RecPhysic
        'Add By ZYB 2002-11-27
        '取各页面对应的排序串
        If MnuEditDosage.Checked Then
            strOrder = strOrder_1
        ElseIf MnuEditAbolish.Checked Then
            strOrder = strOrder_2
        ElseIf MnuEditConsignment.Checked Then
            strOrder = strOrder_3
        Else
            strOrder = strOrder_4
        End If
        
        If strOrder <> "" And RecPhysic.RecordCount <> 0 Then
            strOrder = GetOrder(strOrder)
            RecPhysic.Sort = strOrder
        End If
    End With
    
    If err <> 0 Then
        MsgBox "读取" & IIf(IntStyle = 0, "未配药单据", IIf(IntStyle = 1, "已配药单据", IIf(IntStyle = 2, "未发药单据", "已发药单据"))) & "时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
End Function

Private Sub Msf列表_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    Dim lngColor As Long
    Dim bln配药 As Boolean
    Dim rsTmp As ADODB.Recordset
            
    picRecipeColor.Visible = False
    lblRecipeType.Visible = False
                
    mnuFileRestore.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 3)
    
    If RecPhysic.State = 1 Then
        If RecPhysic.RecordCount > 0 Then
            stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
        Else
            stbThis.Panels(2) = ""
        End If
    End If
    
    With Msf列表
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngSendRow > 0 And LngSendRow < .Rows Then
            .Row = LngSendRow       '清除上次选中行
            lngColor = Val(Msf列表.TextMatrix(LngSendRow, 处方列名.可操作))
            lngColor = IIf(tabShow.Tab <> 3 Or lngColor = 0, &H80000008, IIf(lngColor = 1, glng正常, IIf(lngColor = 2, glng发药, glng退药)))
            For intCol = 处方列名.选择 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = lngColor
            Next
            .Col = 处方列名.选择
        End If
        
        LngSendRow = LngSelectRow
        .Row = LngSendRow       '设置当前选中行
        lngColor = Val(Msf列表.TextMatrix(LngSendRow, 处方列名.可操作))
        lngColor = IIf(tabShow.Tab <> 3 Or lngColor <= 1, glng正常, IIf(lngColor = 2, glng发药, glng退药))
        For intCol = 处方列名.选择 To .Cols - 1
                .Col = intCol
                .CellBackColor = &HC0C0C0
                .CellForeColor = lngColor
        Next
        .Col = 处方列名.选择
        .Redraw = True:
        
        '读取数据
        If .TextMatrix(.Row, 处方列名.单据) = "" Then TxtNo.Clear: Exit Sub
        
        BlnInOper = False
        With TxtNo
            If Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO) & "--" & Msf列表.TextMatrix(Msf列表.Row, 处方列名.姓名) <> .Text Then
                .Clear
                .AddItem Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO) & "--" & Msf列表.TextMatrix(Msf列表.Row, 处方列名.姓名)
                .ItemData(.NewIndex) = Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)
                
                BlnAllowClick = False
                .ListIndex = 0
                BlnAllowClick = True
            End If
        End With
        StrLastNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
        IntLastBill = Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)
        strLastData = Msf列表.TextMatrix(Msf列表.Row, 处方列名.日期)
        mstrNo = Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO)
        IntBillStyle = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据))
        
        mnuCancel.Enabled = False
        Tbar1.Buttons("Cancel").Enabled = False
            
        '退药状态时取当前单据的记录性质和门诊标志，并判断是否可以取消发药
        If MnuEditHandback.Checked Then
            mint门诊标志 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.门诊标志))
            mint记录性质 = Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.记录性质))
            
            '设置取消发药模式是否可用
            If mbln允许取消发药 Then
                If (((mint门诊标志 = 1 Or mint门诊标志 = 4) And gtype_UserSysParms.P15_门诊收费与发药分离 = 1) Or _
                    (mint门诊标志 = 2 And gtype_UserSysParms.P16_住院记帐与发药分离 = 1)) And CheckIsSended(IntLastBill, StrLastNo) = False Then
                    mnuCancel.Enabled = True
                    Tbar1.Buttons("Cancel").Enabled = True
                End If
            End If
        End If
        
        '检查单据是否存在
        If Not CheckBillExist(IntBillStyle, mstrNo) Then
            MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
            Call mnuViewRefresh_Click
            Exit Sub
        End If
        
        '设置cmdAlley按钮状态
        If tabShow.Tab = 2 And mblnStarPass Then
            '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就不显示cmdAlley按钮
            gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,病人费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, IntBillStyle)
            If rsTmp.RecordCount = 0 Then
                If cmdAlley.Visible Then cmdAlley.Visible = False
            Else
                If Not cmdAlley.Visible Then cmdAlley.Visible = True
            End If
        End If
        
        Call ReadBillData(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据), Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO))
        
        If tabShow.Tab = 0 Then
            bln配药 = IsDosage(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)), Msf列表.TextMatrix(Msf列表.Row, 处方列名.NO))
            If bln配药 Then
                CmdSend.Caption = "配药(&V)"
            Else
                CmdSend.Caption = "发药(&S)"
            End If
        End If
        
        '设置处方明细栏中的标签颜色和说明
        If tabShow.Tab = 3 Then
            picRecipeColor.Visible = False
            lblRecipeType.Visible = False
        Else
            picRecipeColor.Visible = True
            lblRecipeType.Visible = True
            picRecipeColor.BackColor = Val(Split(mstrUserRecipeColor, ";")(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.处方类型))))
            lblRecipeType.Caption = Split(mconstrRecipeType, ";")(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.处方类型)))
        End If
    End With

End Sub

Private Sub Msf列表_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String, strOrder As String
    Dim lngRow As Long, lngColor As Long, intCol As Integer
    Dim intMouseRow As Integer, intMouseCol As Integer
    
    'Add by ZYB 2002-11-27
    '增加点击列排序的功能
    If Button <> 1 Then Exit Sub
    intMouseRow = Msf列表.MouseRow
    intMouseCol = Msf列表.MouseCol
    If intMouseRow = 0 Then
        '取列名
        Select Case intMouseCol
        Case 处方列名.选择
            '全部打勾或者全部不打勾
            Call SetCheckBox(0)
            Exit Sub
        Case 处方列名.类型
            strColumn = "类型"
        Case 处方列名.NO
            strColumn = "NO"
        Case 处方列名.姓名
            strColumn = "病人ID,姓名"
        Case 处方列名.金额
            strColumn = "金额"
        Case 处方列名.日期
            strColumn = "日期"
        Case 处方列名.说明
            strColumn = "说明"
        Case Else
            Exit Sub
        End Select
        
        '取排序串
        If MnuEditDosage.Checked Then
            strOrder = strOrder_1
        ElseIf MnuEditAbolish.Checked Then
            strOrder = strOrder_2
        ElseIf MnuEditConsignment.Checked Then
            strOrder = strOrder_3
        Else
            strOrder = strOrder_4
        End If
        
        '如果列名相同，则改变排序方式；否则按升序方式
        If strOrder Like "*" & strColumn & "*" Then
            strOrder = ExchangeOrder(strOrder)
        Else
            strOrder = strColumn & strAsc
        End If
        
        '对全局变量赋值
        If MnuEditDosage.Checked Then
            strOrder_1 = strOrder
        ElseIf MnuEditAbolish.Checked Then
            strOrder_2 = strOrder
        ElseIf MnuEditConsignment.Checked Then
            strOrder_3 = strOrder
        Else
            strOrder_4 = strOrder
        End If
        strOrder = GetOrder(strOrder)
        
        '对记录集进行排序并重新绑定
        If RecPhysic.RecordCount = 0 Then Exit Sub
        RecPhysic.Sort = strOrder
        With Msf列表
            .Redraw = False
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                stbThis.Panels(2) = "共有" & RecPhysic.RecordCount & "张处方；" & "合计金额" & GetSumMoney(RecPhysic) & "元"
            End If
            DoEvents
            Call SetFormat(1, RecPhysic.EOF)
            Call SetCheckBox(-1)
            DoEvents
        
            '上色
            If MnuEditHandback.Checked Then
                .Redraw = False
                For lngRow = 1 To .Rows - 1
                    .Row = lngRow
                    lngColor = IIf(Val(.TextMatrix(lngRow, 处方列名.可操作)) = 1, glng正常, IIf(Val(.TextMatrix(lngRow, 处方列名.可操作)) = 2, glng发药, glng退药))
                    For intCol = 处方列名.选择 To .Cols - 1
                        .Col = intCol
                        .CellForeColor = lngColor
                    Next
                Next
                .Redraw = True
            End If
            .Redraw = True
            .Row = 1
        End With
        '设置处方颜色
        Call SetRecipeColor
        Call Msf列表_EnterCell
    Else
        '点击其他行时
        If intMouseCol = 处方列名.选择 Then
            Call SetCheckBox(intMouseRow)
            Exit Sub
        End If
    End If
End Sub

Private Sub PicCloseConsignment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    RaisEffect PicCloseConsignment, -2     '下凹
End Sub

Private Sub PicCloseConsignment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    RaisEffect PicCloseConsignment, 2      '外凸
    
    '如果不是在本控件上，则退出
    If x < 0 Or x > PicCloseConsignment.Width Then Exit Sub
    If y < 0 Or y > PicCloseConsignment.Height Then Exit Sub
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    'Modified by ZYB 2002-11-19 保存用户设置
    Call SaveFlexState(Bill处方明细.MsfObj, Me.Name & "\" & PreviousTab)
    Call SaveFlexState(Msf列表, Me.Name & "\" & PreviousTab)
    '恢复设置
    Call RestoreFlexState(Msf列表, Me.Name & "\" & tabShow.Tab)
    Call RestoreFlexState(Bill处方明细.MsfObj, Me.Name & "\" & tabShow.Tab)
    Bill处方明细.ColWidth(列名.审查结果) = IIf(Not mblnStarPass, 0, 240)
    
    txtFind.Text = ""
    Call SetMenuState
    
    BlnInOper = False
    Call mnuViewRefresh_Click
End Sub
Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreView_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Find"
        MnuViewFind_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnufileexit_Click
    Case "Charge"
        mnuCharge_Click
    Case "Stuff"
        mnuStuff_Click
    Case "Cancel"
        mnuCancel_Click
    End Select
End Sub

Private Sub Tbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu MnuViewTool, 2
End Sub

Private Sub TimePrint_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If tabShow.Tab = 3 Then
        If Chk全退.Value = 0 Or mblnAllBack = False Then Exit Sub
    End If
    
    TimePrint.Enabled = False
    DoEvents
    '调用打印程序
    Call AutoPrint
    DoEvents
    TimePrint.Enabled = True
    
    If mint自动配药 = 1 Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub txtFind_Change()
    If txtFind.Text = "" Then txtFind.Tag = ""
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtFind.Text = "" And Me.ActiveControl Is txtFind)
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Tag = "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
    txtFind.Tag = ""
    
    If Not mobjIDCard Is Nothing And txtFind.Text = "" Then
        mobjIDCard.SetEnabled (True)
    End If
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    mblnCard = False
    If imgFilter.BorderStyle = cstLocate Then
        If KeyAscii = 13 Then
             Call Form_KeyDown(vbKeyF3, 0)
             Exit Sub
        End If
             
        If lblFind.Tag = FindType.姓名 Then
            mblnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, glngSys)
        ElseIf lblFind.Tag = FindType.就诊卡 Then
            mblnCard = (KeyAscii <> 8 And Len(txtFind.Text) = gtype_UserSysParms.P20_就诊卡号长度 - 1 And txtFind.SelLength <> Len(txtFind.Text))
        End If
        
        If mblnCard Or KeyAscii = 13 Then
            If KeyAscii <> 13 Then
                txtFind.Text = txtFind.Text & Chr(KeyAscii)
                txtFind.SelStart = Len(txtFind.Text)
            End If
            KeyAscii = 0
            Call Form_KeyDown(vbKeyF3, 0)
        Else
            Select Case lblFind.Tag
                Case FindType.就诊卡
                    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                Case FindType.门诊号
                    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
                Case FindType.单据号
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    If Not (txtFind.Text = "" Or txtFind.SelLength = Len(txtFind.Text)) _
                        And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                Case FindType.姓名
                    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                Case FindType.身份证
                Case FindType.IC卡
            End Select
        End If
    Else
        If KeyAscii = 13 Then
            Call SetFilter(MnuEditHandback.Checked)
            Call zlControl.TxtSelAll(txtFind)
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFind_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    txtFind.MaxLength = 0
    lblFind.Tag = FindType.身份证
    lblFind.Caption = "身份证↓"
    txtFind.Text = strID
    Call txtFind_KeyPress(vbKeyReturn)

    DoEvents

    txtFind.Text = ""

End Sub
Private Sub txtFind_Validate(Cancel As Boolean)
    If Val(lblFind.Tag) = FindType.单据号 Then
        If IsNumeric(txtFind.Text) Then
            txtFind.Text = GetFullNO(txtFind.Text, 13)
        End If
    End If
End Sub
Private Sub TxtNo_Click()
    Dim LngLocate As Long, blnFind As Boolean
    
    On Error GoTo ErrHand
    If BlnAllowClick = False Then Exit Sub
    If TxtNo.ListIndex = -1 Then
        Exit Sub
    End If
    '--为显示数据做准备--
    ClearCons
    
    '--读取单据并显示--
    blnFind = False
    StrLastNo = Mid(TxtNo.Text, 1, 8)
    IntLastBill = TxtNo.ItemData(TxtNo.ListIndex)
    TxtNo.Tag = TxtNo.Text
    '定位表格
    With Msf列表
        For LngLocate = 1 To .Rows - 1
            If Trim(.TextMatrix(LngLocate, 处方列名.单据)) <> "" Then
                If .TextMatrix(LngLocate, 处方列名.单据) = TxtNo.ItemData(TxtNo.ListIndex) And .TextMatrix(LngLocate, 处方列名.NO) = Mid(TxtNo.Text, 1, 8) Then
                    .Row = LngLocate
'                    StrLastNo = ""
                    Msf列表_EnterCell
                    blnFind = True
                    Exit For
                End If
            End If
        Next
    End With
    If Not blnFind Then If Not ReadBillData(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), blnFind) Then Exit Sub
    BlnInOper = False
    
    If CmdSend.Enabled Then Me.CmdSend.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub TxtNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intYear  As Integer, strYear As String, strCond As String
    Dim bln处方号 As Boolean            '假表明是病人标识号
    Dim RecRecord As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(TxtNo) = "" Then Exit Sub
    '--如果不满八位,则按规则产生--
    '--以A打头表示病人的标识号，其它情况认为是NO号
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    bln处方号 = Not ((Mid(TxtNo, 1, 1) = "B") Or (Mid(TxtNo, 1, 1) = "+"))
    If Not bln处方号 Then
        If MnuEditDosage.Checked Then
'            StrFind_1 = GetDateSQL(StrFind_1) & _
'            " And Upper(DECODE(A.单据,8,A.门诊号,A.住院号)) Like '" & Mid(txtNo.Text, 2) & "%'"
            StrFind_1 = GetDateSQL(StrFind_1) & _
            " And Upper(DECODE(A.单据,8,A.门诊号,A.住院号)) Like [12] "
            SQLCondition.str当前NO = Mid(TxtNo.Text, 2) & "%"
        ElseIf MnuEditAbolish.Checked Then
            StrFind_2 = GetDateSQL(StrFind_2) & _
            " And Upper(DECODE(A.单据,8,A.门诊号,A.住院号)) Like [12] "
            SQLCondition.str当前NO = Mid(TxtNo.Text, 2) & "%"
        ElseIf MnuEditConsignment.Checked Then
            StrFind_3 = GetDateSQL(StrFind_3) & _
            " And Upper(DECODE(A.单据,8,A.门诊号,A.住院号)) Like [12] "
            SQLCondition.str当前NO = Mid(TxtNo.Text, 2) & "%"
        Else
            StrFind_4 = GetDateSQL(StrFind_4) & _
            " And Upper(H.标识号) Like [12] "
            SQLCondition.str当前NO = Mid(TxtNo.Text, 2) & "%"
        End If
        Call DataRefresh
        Exit Sub
    End If
    
    TxtNo.Text = GetFullNO(TxtNo.Text, 13)
    
    If mInt单据 = 0 Then
        strCond = " And 单据 In (8,9)" '门诊及住院所有单据
    Else
        If mInt单据 = 8 Then
            strCond = " And 单据 In (8,9) And 主页ID Is NULL " '门诊划价及门诊记帐
        Else
            strCond = " And 单据 = 9 And 主页ID Is Not NULL " '住院记帐
        End If
    End If

    '--如果有两条记录,则提出让用户选择(如果不等于上次NO号则重新提取)--
    With RecRecord
        If .State = 1 Then .Close
        If MnuEditHandback.Checked = False Then
            gstrSQL = "Select A.No,A.单据,A.姓名 " & _
                " From 未发药品记录 A" & _
                " Where (Nvl(A.库房ID,0)=0 Or A.库房ID+0=[13] )" & strCond & _
                " And A.No =[12] "
            SQLCondition.str当前NO = Mid(TxtNo, 1, 8)
        Else
            strCond = Replace(strCond, "单据", "A.单据")
            strCond = Replace(strCond, "主页ID", "H.主页ID")
            
            Dim strCond2 As String
            strCond2 = 转换退药串
            gstrSQL = " Select Distinct A.No,A.单据,H.姓名 " & _
                     " From " & _
                     "     (SELECT A.ID,A.No,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                     "          DECODE(SIGN((A.实际数量*NVL(A.付数,1))-B.已发数量),0,A.付数,1) 付数," & _
                     "          DECODE(SIGN((A.实际数量*NVL(A.付数,1))-B.已发数量),0,A.实际数量,B.已发数量) 实际数量,A.记录状态," & _
                     "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.填制人,A.填制日期,A.配药人,A.对方部门ID,A.库房ID" & _
                     "      From" & _
                     "          (SELECT *" & _
                     "          From 药品收发记录 A" & _
                     "          WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                     "          And A.库房ID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " And A.审核日期 " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                     "          From 药品收发记录 A" & _
                     "          Where A.审核人 Is Not Null" & _
                     "          And A.库房ID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " And A.审核日期 " & StrDate & "", strCond2) & _
                     "          GROUP BY A.no,A.单据,A.药品ID,A.序号) B" & _
                     "      Where A.no = B.no And A.单据 = B.单据 And A.药品ID+0 = B.药品ID And A.序号 = B.序号" & _
                     "     ) A,病人费用记录 H" & _
                     " Where A.库房ID+0=[13] " & strCond & _
                     " And A.No ='" & Mid(TxtNo, 1, 8) & "'" & _
                     " And A.费用ID=H.ID And (Mod(A.记录状态,3)=0 Or A.记录状态=1) And A.实际数量<>0 "
        
            '一张处方不可能同时存在于在线与后备表中，因此，如果数据移出，就直接从后备表中提取，否则原SQL不变
            '药品处方发药可同时对单据 IN (8,9)的单据，因此不排除可能8在线而9后备中的情况
            Dim blnMoved As Boolean
            Dim strSQL As String
            
            blnMoved = zlDatabase.NOMoved("药品收发记录", Mid(TxtNo, 1, 8), " 单据 IN ", " (8,9)")
            
            '如果存在数据转出，则需要同时从后备表中提取数据（可能存在不同类型的单据分别在线与后备表中）
            If blnMoved Then
                strSQL = gstrSQL
                strSQL = Replace(strSQL, "药品收发记录", "H药品收发记录")
                strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
                gstrSQL = gstrSQL & " UNION ALL " & strSQL
            End If
        End If
    End With

   Set RecRecord = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date开始日期, _
            SQLCondition.date结束日期, _
            SQLCondition.str开始NO, _
            SQLCondition.str结束NO, _
            SQLCondition.str姓名, _
            SQLCondition.str就诊卡, _
            SQLCondition.str标识号, _
            SQLCondition.lng科室ID, _
            SQLCondition.str填制人, _
            SQLCondition.str审核人, _
            SQLCondition.lng药品ID, _
            SQLCondition.str当前NO, _
            lng药房ID)

    With RecRecord
        TxtNo.Clear
        If .EOF Then
            MsgBox "未找到指定处方，请重新输入！", vbInformation, gstrSysName
            Msf列表_EnterCell
            Exit Sub
        End If
        Do While Not .EOF
            TxtNo.AddItem !NO & "--" & !姓名
            TxtNo.ItemData(TxtNo.NewIndex) = !单据
            .MoveNext
        Loop
        
        If TxtNo.ListCount = 0 Then Exit Sub
        
        If MnuEditDosage.Checked Then
            '配药
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "配药") And mint自动配药 = 0
        ElseIf MnuEditAbolish.Checked Then
            '取消
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "配药")
        ElseIf MnuEditConsignment.Checked Then
            '发药
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "发药")
        Else
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "退药")
        End If
        
        TxtNo.ListIndex = 0
        StrLastNo = Mid(TxtNo, 1, 8)
        IntLastBill = TxtNo.ItemData(TxtNo.ListIndex)
        
        If .RecordCount > 1 Then
            MsgBox "发现多张相同单号的处方单据，请选择！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
End Sub

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, Optional ByVal blnExist As Boolean = True) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
    On Error Resume Next
    err = 0
    ReadBillData = False
  
    strUnit = GetUnit(lng药房ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    Select Case strUnit
    Case "售价单位"
        strSubSql = "1"
    Case "门诊单位"
        strSubSql = "Decode(门诊包装,Null,1,0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "Decode(住院包装,Null,1,0,1,住院包装)"
    Case "药库单位"
        strSubSql = "Decode(药库包装,Null,1,0,1,药库包装)"
    End Select
    Call Get单位串
    
    '得到药品名称串
    Select Case int药品名称
    Case 0  '药品编码与名称
        strName = "'['||C.编码||']'||" & IIf(mblnTradeName, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "C.编码 As 品名,"
    Case 2  '药品名称
        strName = IIf(mblnTradeName, "NVL(E.名称,C.名称)", "C.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(E.名称,'')", "Decode(E.名称,Null,'',C.名称)") & " As 其它名, "
    
    If MnuEditHandback.Checked = False Then
        'Modified By 朱玉宝 2003-12-10 地区：泸州 增加库存数
        gstrSQL = " SELECT DISTINCT B.NO,H.序号,T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.床号,H.开单人,B.ID," & _
            " B.药品ID,DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号," & _
            " NVL(B.批次,0) 批次,NVL(D.药房分批,0) 分批," & strName & _
            " DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & str单位串 & ",K.实际数量/" & strSubSql & " 库存数," & _
            " NVL(B.付数,1) 付数,B.单量,B.用法,B.频次,B.填制人,B.填制日期,H.操作员姓名," & IIf(MnuEditHandback.Checked = False, "B.配药人", "B.审核人") & " 配药人,L.库房货位,M.医生嘱托,M.id 医嘱id,nvl(M.审查结果,-1) 审查结果,I.计算单位,round(B.零售金额," & mintMoneyDigit & ") 零售金额,H.费别,P.毒理分类, " & _
            " B.实际数量*D.剂量系数* Nvl(B.付数, 1) 重量,Decode(Sign(Nvl(J.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名 " & _
            " FROM 药品收发记录 B,药品规格 D,药品特性 P,收费项目目录 C,收费项目别名 E," & _
            " 病人费用记录 H,病人医嘱记录 M,部门表 S,部门表 T,药品库存 K,药品储备限额 L,诊疗项目目录 I,诊疗项目别名 Z ," & _
            " (Select 库房id, 药品id, Nvl(Sum(实际数量), 0) 库存数量 From 药品库存 Where 性质 = 1 And 库房id = [13] Group By 库房id, 药品id) J" & _
            " WHERE D.药品ID=C.ID And D.药名ID=P.药名ID And H.医嘱序号=M.ID(+) AND C.ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
            " And B.药品ID=L.药品ID(+) And Nvl(B.库房ID,[13])=L.库房ID(+)" & _
            " AND H.开单部门ID=T.ID(+) AND B.药品ID=D.药品ID AND MOD(B.记录状态,3)=1" & _
            " AND S.ID=NVL(B.库房ID,[13]) AND B.费用ID=H.ID AND B.NO=[14] AND B.单据=[15] AND NVL(B.库房ID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.摘要,'小宝')))<>'拒发'" & _
            " AND B.药品ID=K.药品ID(+) AND K.性质(+)=1 AND NVL(B.库房ID,[13])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0) AND B.审核人 IS NULL And D.药名id=I.id " & _
            " And Nvl(B.库房id, [13]) + 0 = J.库房id(+) And B.药品id = J.药品id(+) And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 "
     Else
        '汇总显示单据内容
        '不可能存在一张处方同时在线与后备表中都存在
        blnMoved = zlDatabase.NOMoved("药品收发记录", BillNo, " 单据 = ", BillStyle)
        gstrSQL = " SELECT DISTINCT B.NO,H.序号,T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.床号,H.开单人,B.ID,B.药品ID," & _
                 " DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号," & _
                 " NVL(B.批次,0) 批次,NVL(D.药房分批,0) 分批," & strName & _
                 " DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & str单位串 & "," & _
                 " NVL(B.付数,1) 付数," & _
                 " B.已退数量/" & strSubSql & " 已退数量," & _
                 " B.已发数量/" & strSubSql & " 准退数,B.已发数量 实际数量," & _
                 " B.单量,B.用法,B.频次,B.填制人,B.填制日期,H.操作员姓名," & IIf(MnuEditHandback.Checked = False, "B.配药人", "B.审核人") & " 配药人,I.计算单位,round(B.零售金额," & mintMoneyDigit & " ) 零售金额,H.费别,P.毒理分类, "
        If Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist Then    '输入的情况考虑进去
            Dim strCond2 As String
            strCond2 = 转换退药串
            gstrSQL = gstrSQL & " B.已发数量*D.剂量系数 重量,Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名 FROM "
            gstrSQL = gstrSQL & "   (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                     "          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                     "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.填制人,A.填制日期,A.审核人,A.审核日期,A.对方部门ID,A.库房ID" & _
                     "      FROM" & _
                     "          (SELECT *" & _
                     "          FROM 药品收发记录 A" & _
                     "          WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                     "          AND A.库房ID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                     "          FROM 药品收发记录 A" & _
                     "          WHERE A.审核人 IS NOT NULL" & _
                     "          AND A.库房ID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                     "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
                     "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 " & _
                     "      )"
        Else
            gstrSQL = gstrSQL & " B.实际数量*D.剂量系数 重量,Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) 库存下限,Z.名称 As 英文名 FROM "
            gstrSQL = gstrSQL & "(Select 0 已发数量,0 已退数量,0 准退数量,A.* From 药品收发记录 A)"
        End If
        gstrSQL = gstrSQL & _
                 "       B,药品规格 D,药品特性 P,收费项目目录 C,收费项目别名 E,病人费用记录 H,部门表 S,部门表 T,诊疗项目目录 I,诊疗项目别名 Z , " & _
                 "(Select 库房id, 药品id, Nvl(Sum(实际数量), 0) 库存数量 From 药品库存 Where 性质 = 1 And 库房id = [13] Group By 库房id, 药品id) K, 药品储备限额 L " & _
                 " Where H.开单部门ID=T.ID(+) And B.药品ID=D.药品ID And D.药名ID=P.药名ID And C.ID=D.药品ID " & _
                 " And D.药品ID=E.收费细目ID(+) and E.性质(+)=3 And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 " & _
                 " And S.ID=B.库房ID And B.费用ID=H.ID And B.NO=[14] And B.单据=[15] And B.库房ID+0=[13]"
                 
        If IsDate(Msf列表.TextMatrix(Msf列表.Row, 处方列名.日期)) Then
                 gstrSQL = gstrSQL & " And B.审核日期=To_Date('" & Msf列表.TextMatrix(Msf列表.Row, 处方列名.日期) & "','yyyy-MM-dd hh24:mi:ss')"
        End If
        gstrSQL = gstrSQL & " And B.审核人 Is Not Null And D.药名id=I.id " & _
                            " And B.药品id = L.药品id(+) And Nvl(B.库房id, 24) = L.库房id(+) And" & _
                            " D.药名id = I.ID And Nvl(B.库房id, 24) + 0 = K.库房id(+) And B.药品id = K.药品id(+) "
        
        '如果数据转出，则直接从后备表中提取数据
        If blnMoved Then
            gstrSQL = Replace(gstrSQL, "药品收发记录", "H药品收发记录")
            gstrSQL = Replace(gstrSQL, "病人费用记录", "H病人费用记录")
        End If
    End If
    gstrSQL = gstrSQL & " Order by H.序号,B.药品ID,Nvl(B.批次,0)"
     
    Set RecBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date开始日期, _
            SQLCondition.date结束日期, _
            SQLCondition.str开始NO, _
            SQLCondition.str结束NO, _
            SQLCondition.str姓名, _
            SQLCondition.str就诊卡, _
            SQLCondition.str标识号, _
            SQLCondition.lng科室ID, _
            SQLCondition.str填制人, _
            SQLCondition.str审核人, _
            SQLCondition.lng药品ID, _
            SQLCondition.str当前NO, _
            lng药房ID, BillNo, BillStyle)
    
    If WriteDataToBill(RecBill, blnExist) = False Then Exit Function
    
    '增加中药处方的一些处理 by lyq 2005-04-27
    If 判断是否中药处方(BillStyle, BillNo) Then
        Call 中药处方特别处理(BillStyle, BillNo)
    End If
    
    IntStyle = IIf(MnuEditDosage.Checked, 1, IIf(MnuEditAbolish.Checked, 2, IIf(MnuEditConsignment.Checked, 3, 4)))
    '只判断在线数据（未发药、未配药的数据），因移出数据是不允许操作的，而在具体操作处本来也有此判断，判断也就没有意义了
    If Not blnMoved Then
        If CheckBill(IntStyle, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
    End If
    
    '设置按钮状态
    Select Case tabShow.Tab
    Case 0
        CmdSend.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist) And MnuEditDosage.Visible And mint自动配药 = 0
    Case 1
        CmdSend.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist) And MnuEditAbolish.Visible
    Case 2
        CmdSend.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist) And IsHavePrivs(mstrPrivs, "发药")
    Case 3
        CmdSend.Enabled = (Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist) And IsHavePrivs(mstrPrivs, "退药")
        Chk全退.Enabled = CmdSend.Enabled
        Chk全退.Value = IIf(CmdSend.Enabled, 1, 0)
        mblnAllBack = (Chk全退.Value = 1)
        If blnMoved Then Bill处方明细.Active = False
    End Select

    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        CmdSend.Enabled = False
        Chk全退.Enabled = False
        Exit Function
    End If
    ReadBillData = True
End Function

Private Function WriteDataToBill(ByVal RecData As ADODB.Recordset, Optional ByVal blnExist As Boolean = True) As Boolean
    Dim dblMoney As Currency, IntLocate As Integer
    Dim str操作员 As String, str合计 As String
    Dim dbl总重量 As Double
    Dim str重量单位 As String
    Dim lng整数量 As Long
    Dim dbl小数量 As Double
    
    '--向单据控件中写数据--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    Call ClearCons
    
    mblnAuto = True
    dblMoney = 0
    Lbl配药人.Caption = IIf(tabShow.Tab = 3, IIf(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) <> 3, "发药人", "退药人"), "配药人")
    Lbl收费员.Caption = IIf(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)) = 8, "收费员", "记帐员")
    
    '如果不是原始记录，则不显示列（已退数、准退数）
    Bill处方明细.ColWidth(列名.已退数) = 0
    Bill处方明细.ColWidth(列名.准退数) = 0
    Bill处方明细.ColWidth(列名.退药数) = 0
    If tabShow.Tab = 3 Then
        Bill处方明细.ColWidth(列名.已退数) = IIf(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist, 1000, 0)
        Bill处方明细.ColWidth(列名.准退数) = IIf(Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist, 1000, 0)
        Bill处方明细.ColWidth(列名.退药数) = IIf((Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.可操作)) = 1 Or Not blnExist) And mbln显示大小单位 = False, 1000, 0)
    End If
    
    '填充单据内容
    With RecData
        '填充表头
        If Not .EOF Then
            Me.Txt床号 = IIf(IsNull(!床号), "", !床号)
            If Val(Msf列表.TextMatrix(Msf列表.Row, 处方列名.单据)) = 8 Then Me.Txt床号 = ""
            Me.Txt开单医生.ListIndex = 0
            If (intVerify = 0) And IsHavePrivs(mstrPrivs, "医生查询") Then
                str操作员 = IIf(IsNull(!开单人), "", !开单人)
            Else
                If MnuEditHandback.Checked And intVerify = 1 Then
                    str操作员 = IIf(IsNull(!填制人), "", !填制人)
                Else
                    str操作员 = ""
                End If
            End If
            If str操作员 <> "" Then
                '定位医生
                For IntLocate = 1 To Txt开单医生.ListCount
                    If Mid(Txt开单医生.List(IntLocate), InStr(1, Txt开单医生.List(IntLocate), "-") + 1) = str操作员 Then
                        Txt开单医生.ListIndex = IntLocate
                        Exit For
                    End If
                Next
            End If
            If glngSys \ 100 = 1 Then
                If IsHavePrivs(mstrPrivs, "医生查询") Then
                    Me.Txt科室 = IIf(IsNull(!科室), "", !科室)
                End If
            Else
                Me.Txt科室 = IIf(IsNull(!姓名), "", !姓名)
            End If
            Me.Txt年龄 = IIf(IsNull(!年龄), "", !年龄)
            If IIf(IsNull(!配药人), "", !配药人) <> "" Then
                Me.cbo配药人 = IIf(IsNull(!配药人), "", !配药人)
            End If
            Me.Txt收费员 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
            Me.Txt性别 = IIf(IsNull(!性别), "", !性别)
            Me.Txt住院号 = IIf(IsNull(!住院号), "", !住院号)
        End If
            
        Bill处方明细.Rows = 1
        Bill处方明细.Rows = 2
        Bill处方明细.MsfObj.FixedRows = 1
        Bill处方明细.MsfObj.Redraw = False
        
        Do While Not .EOF
            Bill处方明细.MergeRow .AbsolutePosition, False
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.顺序号) = .AbsolutePosition
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.药品名称) = !品名
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.其它名) = IIf(IsNull(!其它名), "", !其它名)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.英文名) = IIf(IsNull(!英文名), "", !英文名)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.序号) = !序号
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.规格) = IIf(IsNull(!规格), "", !规格)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.批号) = IIf(IsNull(!批号), "", !批号)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.Id) = !Id
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.药品ID) = !药品ID
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.批次) = !批次
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.单位) = IIf(IsNull(!单位), "", !单位)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.单价) = GetFormat(!单价, mintPriceDigit)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.付数) = Format(!付数, "#####0;-#####0; ;")
            
            If mbln显示大小单位 = True Then
                '按大小包装显示数量
                lng整数量 = Int(!数量)
                If !售价单位 = !单位 Or lng整数量 = !数量 Then
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.数量) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                Else
                    dbl小数量 = (Val(!数量) - lng整数量) * !包装
                    If lng整数量 = 0 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.数量) = dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                    Else
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.数量) = lng整数量 & IIf(IsNull(!单位), "", !单位) & dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                    End If
                End If
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.包装) = Val(!包装)
            Else
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.数量) = FormatEx(!数量, mintNumberDigit)
            End If
            
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.金额) = GetFormat(Val(!零售金额), mintMoneyDigit)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.重量) = !重量 & !计算单位
            
            dbl总重量 = dbl总重量 + !重量
            str重量单位 = !计算单位
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.频次) = IIf(IsNull(!频次), "", !频次)
            mstr单量单位 = NVL(!计算单位)
            If Not IsNull(!单量) Then
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.单量) = FormatEx(!单量, 5) & NVL(!计算单位)
            End If
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.用法) = NVL(!用法)
            If MnuEditHandback.Checked Then
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.包装) = Val(!包装)
                If mbln显示大小单位 = True Then
                    '按大小包装显示数量：分别处理已退数量、准退数量、退药数量
                    '已退数量、准退数量列显示模式为"大包装数量+大包装单位+小包装数量+售价单位"；退药数分两列显示，且只显示数值
                    lng整数量 = Int(!已退数量)
                    If !售价单位 = !单位 Or lng整数量 = !已退数量 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.已退数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                    Else
                        dbl小数量 = (Val(!已退数量) - lng整数量) * !包装
                        If lng整数量 = 0 Then
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.已退数) = dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        Else
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.已退数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        End If
                    End If
                    
                    lng整数量 = Int(!准退数)
                    If !售价单位 = !单位 Or lng整数量 = !准退数 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                    Else
                        dbl小数量 = (Val(!准退数) - lng整数量) * !包装
                        If lng整数量 = 0 Then
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数) = dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        Else
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        End If
                    End If
                    
                    lng整数量 = Int(!准退数)
                    If !售价单位 = !单位 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数小) = FormatEx(lng整数量, mintNumberDigit)
                    ElseIf lng整数量 = !准退数 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数大) = FormatEx(lng整数量, mintNumberDigit)
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数小) = FormatEx(0, mintNumberDigit)
                    Else
                        dbl小数量 = (Val(!准退数) - lng整数量) * !包装
                        If lng整数量 = 0 Then
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数小) = FormatEx(dbl小数量, mintNumberDigit)
                        Else
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数大) = FormatEx(lng整数量, mintNumberDigit)
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数小) = FormatEx(dbl小数量, mintNumberDigit)
                        End If
                    End If
                    
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.退药数) = FormatEx(!准退数, mintNumberDigit)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.退药数大) = Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数大)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.退药数小) = Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数小)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.单位大) = IIf(IsNull(!单位), "", !单位)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.单位小) = IIf(IsNull(!售价单位), "", !售价单位)
                Else
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.已退数) = FormatEx(!已退数量, mintNumberDigit)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.准退数) = FormatEx(!准退数, mintNumberDigit)
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.退药数) = FormatEx(!准退数, mintNumberDigit)
                End If
            
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.实际数量) = !实际数量
            Else
                If mbln显示大小单位 = True Then
                    '按大小包装显示数量
                    lng整数量 = Int(!库存数)
                    If !售价单位 = !单位 Or lng整数量 = !库存数 Then
                        Bill处方明细.TextMatrix(.AbsolutePosition, 列名.库存数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                    Else
                        dbl小数量 = (Val(!库存数) - lng整数量) * !包装
                        If lng整数量 = 0 Then
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.库存数) = dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        Else
                            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.库存数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & dbl小数量 & IIf(IsNull(!售价单位), "", !售价单位)
                        End If
                    End If
                Else
                    Bill处方明细.TextMatrix(.AbsolutePosition, 列名.库存数) = FormatEx(NVL(!库存数, 0), mintNumberDigit)
                End If
            
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.货位) = NVL(!库房货位)
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.医生嘱托) = NVL(!医生嘱托)
                Bill处方明细.TextMatrix(.AbsolutePosition, 列名.医嘱id) = NVL(!医嘱id)
                If !审查结果 <> -1 Then
                    BlnEnterCell = False
                    Bill处方明细.Row = .AbsolutePosition
                    Bill处方明细.Col = 0
                    Set Bill处方明细.MsfObj.CellPicture = imgPass.ListImages(Val(!审查结果) + 1).Picture
                    Bill处方明细.MsfObj.CellPictureAlignment = 4
'                    Bill处方明细.CellBackColor = &H8000000F
                    BlnEnterCell = True
                End If
            End If
            
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.分批) = IIf(IsNull(!分批), 0, !分批)
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.新批号) = ""
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.新效期) = ""
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.新产地) = ""
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.备注) = ""
            Bill处方明细.TextMatrix(.AbsolutePosition, 列名.费别) = IIf(IsNull(!费别), "", !费别)
            If MnuEditHandback.Checked Then
                dblMoney = dblMoney + IIf(Chk清单.Value = 1, Val(!零售金额), FormatEx(!准退数 / (!数量 * !付数) * Val(!零售金额), mintMoneyDigit))
            Else
                dblMoney = dblMoney + Val(!零售金额)
            End If
            
            '对低于库存下限的药品上色
            Bill处方明细.MsfObj.Redraw = False
            If !库存下限 = 0 Then
'            If IsLowerLimit(lng药房ID, !药品ID) Then
                Call SetForeColor_ROW(.AbsolutePosition, mlng紫色)
            Else
                Call SetForeColor_ROW(.AbsolutePosition, vbBlack)
            End If
                        
            '特殊药品粗体显示
            If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(!毒理分类)) > 0 And NVL(!毒理分类) <> "" Then
                Bill处方明细.Col = 列名.药品名称
                Bill处方明细.Row = .AbsolutePosition
                Bill处方明细.MsfObj.CellFontBold = True
            End If
                        
            If .AbsolutePosition >= Bill处方明细.Rows - 1 Then Bill处方明细.Rows = Bill处方明细.Rows + 1
            .MoveNext
        Loop
        Bill处方明细.MsfObj.Redraw = True
        '取消最后空白行
        '--If Bill处方明细.Rows - 1 >= 2 Then Bill处方明细.Rows = Bill处方明细.Rows - 1
    End With
    
    '最后空白行显示金额合计
    str合计 = zlCommFun.UppeMoney(dblMoney)
    With Bill处方明细
        .TextMatrix(.Rows - 1, 1) = "金额合计：" & Format(dblMoney, mstrVBMoneyForamt)
        .TextMatrix(.Rows - 1, 2) = "金额合计：" & Format(dblMoney, mstrVBMoneyForamt)
        .TextMatrix(.Rows - 1, 3) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 4) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 5) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 6) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 7) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 8) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 9) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 10) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 11) = "大写：" & str合计
        .TextMatrix(.Rows - 1, 12) = "大写：" & str合计
        If mbln显示重量 And mblnIs中药处方 Then
            .TextMatrix(.Rows - 1, 13) = "总重量：" & dbl总重量 & str重量单位
        End If
        .MergeCell (1)
        .MergeRow .Rows - 1, True
        .MsfObj.LeftCol = 0
    End With
    
    mblnAuto = False
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Bill处方明细.Row = ReLocateDetailRow
    
    WriteDataToBill = True
End Function

Private Function DependOnCheck() As Boolean
    Dim strSQL As String
    '依赖数据检测
    DependOnCheck = False
    
    With RecPart
        gstrSQL = " Select A.简码||'-'||A.姓名 医生 From 人员表 A,人员性质说明 B" & _
                 " Where (A.站点 = '" & gstrNodeNo & "' Or A.站点 is Null) And B.人员性质='医生' And A.ID=B.人员ID" & _
                 " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
                 " Order by A.简码"
        Call zlDatabase.OpenRecordset(RecPart, gstrSQL, "依赖数据检测")
        
        If .EOF Then
            MsgBox "请初始化人员表（医生）", vbInformation, gstrSysName
            Exit Function
        End If
        
        Me.Txt开单医生.Clear
        Txt开单医生.AddItem ""
        Do While Not .EOF
            Txt开单医生.AddItem !医生
            .MoveNext
        Loop
        Txt开单医生.ListIndex = 0
    End With
    
    If IsHavePrivs(mstrPrivs, "所有药房") Then
        strSQL = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房')"
    Else
        strSQL = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房')"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In " & strSQL & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecPart = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecPart
        If .EOF Then
            If IsHavePrivs(mstrPrivs, "所有药房") Then
                strSQL = "请初始化药房！（部门管理）"
            Else
                strSQL = "你不是药房人员，不能使用本模块！"
            End If
            MsgBox strSQL, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
End Function

Private Function ReadFromReg()
    Dim RecRead As New ADODB.Recordset
    Dim strSub1 As String, strSub2 As String
    Dim strSave As String
    Dim arrColumn
    Dim int显示病区处方 As Integer
    
    On Error Resume Next
    
    '取公共及私有参数
    strSave = zlDatabase.GetPara("列设置", glngSys, 1341)
    intFont = Val(zlDatabase.GetPara("字体", glngSys, 1341))
    
    mintShowBill收费 = Val(zlDatabase.GetPara("收费处方显示方式", glngSys, 1341))
    mintShowBill记帐 = Val(zlDatabase.GetPara("记帐处方显示方式", glngSys, 1341))
    mbln记帐单 = (Val(zlDatabase.GetPara("打印包含记帐单", glngSys, 1341)) = 1)
    mIntPrintHandbackNO = Val(zlDatabase.GetPara("打印退费单据间隔", glngSys, 1341))
    mIntPrintDelay = Val(zlDatabase.GetPara("打印延迟", glngSys, 1341))
    int显示病区处方 = Val(zlDatabase.GetPara("显示病区处方", glngSys, 1341))
    mlngRefresh = Val(zlDatabase.GetPara("刷新间隔", glngSys, 1341))
    mlngPrintInterval = Val(zlDatabase.GetPara("打印间隔", glngSys, 1341))
    int校验发药人 = Val(zlDatabase.GetPara("校验发药人", glngSys, 1341))
    int校验配药人 = Val(zlDatabase.GetPara("校验配药人", glngSys, 1341))
    mint自动销帐 = Val(zlDatabase.GetPara("自动销帐", glngSys, 1341))
    mbln显示大小单位 = (Val(zlDatabase.GetPara("显示大小单位", glngSys, 1341)) = 1)
    
    IntShowCol = Val(zlDatabase.GetPara("显示付数", glngSys, 1341))
    IntAutoPrint = Val(zlDatabase.GetPara("发药后自动打印", glngSys, 1341))
    
    '界面条件输入框状态：0-定位;1-过滤。默认是定位
    imgFilter.BorderStyle = Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & Me.Name, "界面定位", cstLocate))
    
    '病区过滤开关：默认是0-不显示
    img病区.BorderStyle = int显示病区处方
    
    '显示退药待发单据
    mlng待发单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & Me.Name, "显示退药待发单据", 1)
    Chk显示退药待发单据.Value = mlng待发单据
    
    '0-不打印未配药单据
    '1-打印本部门所有未配药单据
    '2-打印本窗口所有未配药单据
    '3-选择打印(发药窗口)
    intPrint = Val(zlDatabase.GetPara("发现新单据是否打印", glngSys, 1341))
    mintPrintDrugLable = Val(zlDatabase.GetPara("打印药品标签", glngSys, 1341))
    lng药房ID = Val(zlDatabase.GetPara("发药药房", glngSys, 1341))
    Call GetDrugDigit(lng药房ID, Me.Caption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Str窗口 = zlDatabase.GetPara("发药窗口", glngSys, 1341)
    Str配药人 = zlDatabase.GetPara("配药人", glngSys, 1341)
    strPrintWindow = zlDatabase.GetPara("打印指定发药窗口", glngSys, 1341)
    mint自动配药 = Val(zlDatabase.GetPara("自动配药", glngSys, 1341))
    mint自动配药时限 = Val(zlDatabase.GetPara("自动配药时限", glngSys, 1341))
    
    mstrSourceDep = zlDatabase.GetPara("来源科室", glngSys, 1341)
    
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set RecRead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
    
    With RecRead
        If Not .EOF Then
            IntCheckStock = !库存检查
        End If
        
        .Close
        IntSendAfterDosage = 1          '表示不需要经过配药过程
    End With

   gstrSQL = " Select Nvl(配药,0) AS 配药 From 药房配药控制 Where 药房ID=[1] Order by 门诊"
   Set RecRead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
   
    '只要有一项表示需要经过配药过程的，标记为需要配药
    Do While Not RecRead.EOF
        If RecRead!配药 = 1 Then
            IntSendAfterDosage = 0
            Exit Do
        End If
        RecRead.MoveNext
    Loop
    
    If IntSendAfterDosage = 1 And mint自动配药 = 0 Then
        cbo配药人.Enabled = True
        Call GetDosagePeople
    Else
        cbo配药人.Text = ""
        cbo配药人.Enabled = False
    End If
    
    '屏蔽菜单及工具按钮
    MnuEditDosage.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "配药"))
    MnuEditAbolish.Visible = MnuEditDosage.Visible
    tabShow.TabVisible(0) = MnuEditDosage.Visible
    tabShow.TabVisible(1) = MnuEditDosage.Visible
    mnuChange.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "配药") And mint自动配药 = 1)
    mnuLine10.Visible = mnuChange.Visible
        
    '设置显示及自动打印的条件:注意"未发药品记录"的别名为A
    Select Case mintShowBill收费
        Case 0  '不显示处方
            strSub1 = "A.单据<>9 And A.单据<>8"
            mstrShowSendedBill = "A.单据<>9 And A.单据<>8"
        Case 1  '显示未收费
            strSub1 = "A.单据<>9 And Nvl(A.已收费,0)=0 And A.单据=8"
            mstrShowSendedBill = "A.单据<>9 And A.单据=8"
        Case 2  '显示已收费
            strSub1 = "A.单据<>9 And A.已收费=1 And A.单据=8"
            mstrShowSendedBill = "A.单据<>9 And A.单据=8"
        Case 3  '显示所有处方
            strSub1 = "A.单据<>9 And A.单据=8"
            mstrShowSendedBill = "A.单据<>9 And A.单据=8"
    End Select
    Select Case mintShowBill记帐
        Case 0  '不显示处方
            strSub2 = "A.单据<>8 And A.单据<>9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.单据<>8 And A.单据<>9"
        Case 1  '显示未审核
            strSub2 = "A.单据<>8 And Nvl(A.已收费,0)=0 And A.单据=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.单据<>8 And A.单据=9"
        Case 2  '显示已审核
            strSub2 = "A.单据<>8 And A.已收费=1 And A.单据=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.单据<>8 And A.单据=9"
        Case 3  '显示所有处方
            strSub2 = "A.单据<>8 And A.单据=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.单据<>8 And A.单据=9"
    End Select
    mstrShowBill = " And A.单据 IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    mstrShowSendedBill = " And A.单据 IN(8,9) And (" & mstrShowSendedBill & ")"
    
    '取得药品名称的格式方式
    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|规格,0|批号,0|单位,0|单价,0|数量,0|金额,0|重量,0|用法,0|频次,0|用量,0|库存数,0|库房货位,0|已退数,0|准退数,0|退药数,0|备注"
    arrColumn = Split(strSave, ",")
    int药品名称 = Val(Split(arrColumn(0), "|")(0))
    
    '取处方颜色
    Call GetRecipeColor
    
    '取排序
    strOrder_1 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未配药处方排序串", "")
    strOrder_2 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已配药处方排序串", "")
    strOrder_3 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "未发药处方排序串", "")
    strOrder_4 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "已发药处方排序串", "")
    
    '取输入模式
    mint输入模式 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "输入模式", "0"))
    If mint输入模式 < 0 Or mint输入模式 > 5 Then
        mint输入模式 = 0
    End If
End Function

Private Function CheckAnother() As Boolean
    Dim BlnIn药房 As Boolean, bln住院 As Boolean, Bln单据 As Boolean
    Dim BlnSetPeople As Boolean
    Dim RecTestPeople As New ADODB.Recordset
    Dim LngOld药房ID As Long, StrOld配药人 As String
    
    CheckAnother = False
    
    If lng药房ID <> 0 Then
        With RecPart
            .MoveFirst
            .Find "ID=" & lng药房ID
            BlnIn药房 = (RecPart.EOF <> True)
            
            If BlnIn药房 Then   '说明该部门仍属药房
                '取单位
                bln住院 = False

                gstrSQL = "Select nvl(服务对象,1) 服务对象 From 部门性质说明 Where 部门ID+0=[1]"
                Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !服务对象 = 2 Or !服务对象 = 3 Then bln住院 = True: Exit Do
                        .MoveNext
                    Loop
                    Bln单据 = False
                    If bln住院 Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !服务对象 = 3 Then Bln单据 = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If bln住院 = False Then
                    mInt单据 = 8
                Else
                    mInt单据 = IIf(Bln单据, 0, 9)
                End If
            End If
        End With
    End If
    
    '设置对应的药房
    If lng药房ID = 0 Or BlnIn药房 = False Then
        '调设置窗体
        With Frm发药参数设置
            MsgBox IIf(Str配药人 = "", "请设置药房及配药人！", "请设置药房！"), vbInformation, gstrSysName
            Set .RecPart = RecPart.Clone
            .strShow = IIf(Str配药人 = "", "请设置药房及配药人！", "请设置药房！")
            .mstrPrivs = mstrPrivs
            .Show 1, Me
        End With
        Call ReadFromReg

        '仍未设置药房，退出
        If lng药房ID = 0 Then Exit Function
        '重新获取该药房的使用单位
        With RecPart
            .MoveFirst
            .Find "ID=" & lng药房ID
            BlnIn药房 = (RecPart.EOF <> True)
            
            If BlnIn药房 Then   '说明该部门仍属药房
                '取单位
                bln住院 = False

                gstrSQL = "Select nvl(服务对象,1) 服务对象 From 部门性质说明 Where 部门ID+0=[1]"
                Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !服务对象 = 2 Or !服务对象 = 3 Then bln住院 = True: Exit Do
                        .MoveNext
                    Loop
                    Bln单据 = False
                    If bln住院 Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !服务对象 = 3 Then Bln单据 = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If bln住院 = False Then
                    mInt单据 = 8
                Else
                    mInt单据 = IIf(Bln单据, 0, 9)
                End If
            Else
                Exit Function    '非药房，退出
            End If
        End With
    End If
    
    If IntSendAfterDosage = 0 And Str配药人 <> "|当前操作员|" Then
        LngOld药房ID = lng药房ID
        StrOld配药人 = Str配药人
        
        '设置配药人
        BlnSetPeople = False
        If Str配药人 = "" Then
            MsgBox "请设置配药人！", vbInformation, gstrSysName
            With Frm发药参数设置
                Set .RecPart = RecPart.Clone
                .strShow = "请设置配药人！"
                .mstrPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg

            If Str配药人 = "" Then
                MsgBox "需重新设置配药人，请与系统管理员联系！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '如果配药人非本部门,则必须重新设置
        gstrSQL = " Select Count(*) Records From 部门人员 Where 人员ID=" & _
                 " (Select Distinct ID From 人员表 Where 姓名=[2]) And " & _
                 " 部门ID+0 =[1]"
        Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, Str配药人)
        
        With RecTestPeople
            If .EOF Then
                BlnSetPeople = True
            Else
                If IsNull(!Records) Then
                    BlnSetPeople = True
                Else
                    If !Records = 0 Then
                        BlnSetPeople = True
                    End If
                End If
            End If
        End With
        If BlnSetPeople Then
            MsgBox "请设置配药人（原配药人已不属于本药房）！", vbInformation, gstrSysName
            With Frm发药参数设置
                Set .RecPart = RecPart.Clone
                .strShow = "请设置配药人（原配药人已不属于本药房）！"
                .mstrPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg
            If Str配药人 = "" Then
                MsgBox "需重新设置配药人（原配药人已不属于本药房），请与系统管理员联系！", vbInformation, gstrSysName
                Exit Function
            End If
            If StrOld配药人 = Str配药人 And LngOld药房ID = lng药房ID Then Exit Function
        End If
    End If
    
    CheckAnother = True
End Function

Private Function CheckSpec(ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim strNote As String
    Dim rsTemp As New ADODB.Recordset
    '对毒麻类药品进行检查
    gstrSQL = " SELECT Distinct " & _
        " '['||C.编码||']'||" & IIf(mblnTradeName, "NVL(L.名称,C.名称)", "C.名称") & " 品名,X.毒理分类" & _
        " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,收费项目别名 L,药品特性 X " & _
        " WHERE A.药品ID=B.药品ID AND B.药名ID=X.药名ID And B.药品ID=C.ID " & _
        " AND B.药品ID=L.收费细目ID(+) AND L.性质(+)=3 AND L.码类(+)=1 " & _
        " AND A.审核人 IS NULL AND MOD(A.记录状态,3)=1 AND NVL(A.摘要,'小宝')<>'拒发'" & _
        " AND A.NO=[1] AND A.单据=[2] AND (A.库房ID+0=[3] OR A.库房ID IS NULL) " & _
        " And X.毒理分类<>'普通药'" & _
        " Order by X.毒理分类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[对毒麻类药品进行检查]", strNo, IntBillStyle, lng药房ID)
    
    If rsTemp.RecordCount = 0 Then
        CheckSpec = True
        Exit Function
    End If
    
    With rsTemp
        Do While Not .EOF
            strNote = strNote & vbCrLf & Space(4) & !毒理分类 & "-" & !品名
            .MoveNext
        Loop
    End With
'    If MsgBox("是否对以下毒、麻、精神类药品进行发药？" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    mstr毒麻类提示 = strNote
    CheckSpec = True
End Function
Private Function CheckStock(ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim RecCheckStock As New ADODB.Recordset, RecBillData As New ADODB.Recordset
    Dim dblStock As Double, intCheck As Integer
    '--检查库存--
    '0-不检查;1-检查,不足提醒;2-检查,不足禁止
    On Error Resume Next
    err = 0
    CheckStock = False
    intCheck = IntCheckStock
    
    '逐行检查
    If intCheck <> 0 Then
        gstrSQL = " SELECT A.药品ID,SUM(NVL(A.实际数量,0)*NVL(A.付数,1)) 数量," & _
                " '['||C.编码||']'||" & IIf(mblnTradeName, "NVL(L.名称,C.名称)", "C.名称") & " 品名,NVL(A.批次,0) 批次" & _
                " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,收费项目别名 L,病人费用记录 D " & _
                " WHERE A.药品ID=B.药品ID AND B.药品ID=C.ID" & _
                " AND B.药品ID=L.收费细目ID(+) AND L.性质(+)=3 AND L.码类(+)=1 " & _
                " AND A.审核人 IS NULL AND MOD(A.记录状态,3)=1 AND NVL(A.摘要,'小宝')<>'拒发'" & _
                " AND A.费用ID=D.ID AND A.NO=[1] AND A.单据=[2] AND (A.库房ID+0=[3] OR A.库房ID IS NULL) " & _
                " GROUP BY A.药品ID,'['||C.编码||']'||" & IIf(mblnTradeName, "NVL(L.名称,C.名称)", "C.名称") & ",批次"
        Set RecBillData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lng药房ID)
        
        With RecBillData
            Do While Not .EOF
                gstrSQL = " Select nvl(实际数量,0) 数量" & _
                         " From 药品库存 " & _
                         " Where 库房ID+0=[1] And 药品ID=[2] " & _
                         " And 性质=1 And Nvl(批次,0)=[3]"
                Set RecCheckStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, CLng(RecBillData!药品ID), CLng(RecBillData!批次))
                
                With RecCheckStock
                    If .EOF Then
                        dblStock = 0
                    Else
                        dblStock = !数量
                    End If
                    
                    If dblStock < RecBillData!数量 Then
                        Select Case intCheck
                        Case 1
                            If MsgBox(RecBillData!品名 & "的库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox RecBillData!品名 & "的库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End With
                .MoveNext
            Loop
        End With
    End If
    
    If err <> 0 Then
        MsgBox "检查库存时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckStock = True
End Function

Private Function ResizePicClose()
    Dim DblHeight As Double, DblWidth As Double
    Dim intCols As Integer
    '--调整关闭按钮的位置--
    
    With Msf列表
        DblHeight = .CellHeight * .Rows
        DblWidth = 0
        For intCols = 0 To .Cols - 1
            DblWidth = DblWidth + .ColWidth(intCols)
        Next
        
        If DblHeight > .Height - 180 Or (DblHeight > .Height - 420 And DblWidth > .Width - 70) Then
            With PicCloseConsignment
                .Left = Msf列表.Width - .Width - 30 - 250
            End With
        Else
            With PicCloseConsignment
                .Left = Msf列表.Width - .Width - 30
            End With
        End If
    End With
End Function

Private Sub ClearCons()
    Me.Txt床号 = ""
    Me.Txt开单医生.ListIndex = 0
    Me.Txt科室 = ""
    Me.Txt年龄 = ""
'    Me.cbo配药人 = ""
    Me.Txt性别 = ""
    Me.Txt住院号 = ""
    Me.Txt收费员 = ""
    Me.txt原始付数 = ""
    Me.txt中药煎法 = ""
    
    Bill处方明细.ClearBill
End Sub

Private Function SetMenuState()
    MnuEditDosage.Checked = IIf(tabShow.Tab = 0, True, False)
    MnuEditAbolish.Checked = IIf(tabShow.Tab = 1, True, False)
    MnuEditConsignment.Checked = IIf(tabShow.Tab = 2, True, False)
    MnuEditHandback.Checked = IIf(tabShow.Tab = 3, True, False)
    
    mnuCancel.Enabled = False
    Tbar1.Buttons("Cancel").Enabled = False
    
    '取消发药仅用于退药模式，并且在分离模式下
    Chk清单.Visible = (tabShow.Tab = 3)
    Chk显示退药待发单据.Visible = (tabShow.Tab = 0 Or tabShow.Tab = 1 Or tabShow.Tab = 2)
    
    Call SetPosition
    
End Function

Private Function SetButtonState()
    If MnuEditDosage.Checked Then
        tabShow.Tab = 0
    ElseIf MnuEditAbolish.Checked Then
        tabShow.Tab = 1
    ElseIf MnuEditConsignment.Checked Then
        tabShow.Tab = 2
    Else
        tabShow.Tab = 3
    End If
    
    BlnInOper = False
    Call mnuViewRefresh_Click
End Function

Private Sub TimeRefresh_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If tabShow.Tab = 3 Then
        If Chk全退.Value = 0 Or mblnAllBack = False Then Exit Sub
    End If
    
    TimeRefresh.Enabled = False
    DoEvents
    Call mnuViewRefresh_Click
    DoEvents
    TimeRefresh.Enabled = True
End Sub
Private Sub TimePrintCancelBill_Timer()
    Dim curDateBegin As Date
    Dim curDateEnd As Date
    
    '调用打印退费单
    IntTimes = IntTimes + 1
    '不到分钟数退出
    If IntTimes < mIntPrintHandbackNO Then Exit Sub
    IntTimes = 0
    
    curDateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    curDateBegin = DateAdd("n", 0 - mIntPrintHandbackNO, curDateEnd)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "开始时间=" & Format(curDateBegin, "yyyy-MM-dd hh:mm"), "结束时间=" & Format(curDateEnd, "yyyy-MM-dd hh:mm"), "药房=" & lng药房ID, 2)
End Sub
Private Function AutoPrint()
'功能：自动打印单据
    Dim recAutoPrint As New ADODB.Recordset, strErr As String
    Dim datCurr As Date, strRefresh As String, strCond As String
    Dim strUnit As String
    Dim str操作员 As String
    Dim blnInTrans As Boolean
    Dim blnIgnore As Boolean
    Dim strName As String
    
    '根据打印参数组合条件
    '0-不打印未配药单据
    '1-打印本部门所有未配药单据
    '2-打印本窗口所有未配药单据
    '3-选择打印(发药窗口)
    If BlnInRefresh Then Exit Function
    
    If mblnIsFirst = False And mint自动配药 = 1 Then
        If mint自动配药时限 > 0 Then
            If DateDiff("s", mdate上次校验时间, zlDatabase.Currentdate) > mint自动配药时限 * 60 Then
                strName = zlDatabase.UserIdentify(Me, "校验配药人", glngSys, 1341, "配药")
               
                If Trim(strName) = "" Then Exit Function
                mstr自动配药人 = strName
                
                mdate上次校验时间 = zlDatabase.Currentdate
            End If
        End If
    End If
    
    Select Case intPrint
        Case 0
            If mintPrintDrugLable = 0 Then Exit Function
        Case 1
            If Not mbln记帐单 Then strRefresh = " And 单据=8"
        Case 2
            If mbln记帐单 Then
                If Str窗口 <> "" Then
                    strRefresh = " And (单据=8 And 发药窗口 IN(" & Str窗口 & ") Or 单据=9)"
                End If
            Else
                If Str窗口 <> "" Then
                    strRefresh = " And 单据=8 And 发药窗口 IN(" & Str窗口 & ")"
                Else
                    strRefresh = " And 单据=8"
                End If
            End If
        Case 3
            If mbln记帐单 Then
                If strPrintWindow <> "" Then
                    strRefresh = " And (单据=8 And 发药窗口 IN(" & strPrintWindow & ") Or 单据=9)"
                End If
            Else
                If strPrintWindow <> "" Then
                    strRefresh = " And 单据=8 And 发药窗口 IN(" & strPrintWindow & ")"
                Else
                    strRefresh = " And 单据=8"
                End If
            End If
    End Select
    
    If mInt单据 = 0 Then
        strCond = " And A.单据 In (8,9)" '门诊及住院所有单据
    Else
        If mInt单据 = 8 Then
            strCond = " And A.单据 In (8,9) And A.主页ID Is NULL " '门诊划价及门诊记帐
        Else
            strCond = " And A.单据 = 9 And A.主页ID Is Not NULL " '住院记帐
        End If
    End If
            
    On Error GoTo ErrHand
    BlnInRefresh = True
    
    gstrSQL = " Select NO,单据,填制日期" & _
               " From 未发药品记录 A " & _
               " Where 库房ID+0=[13] " & strRefresh & IIf(mint自动配药 = 1, " And 配药人 Is Null ", "") & _
               " " & IIf(StrFind_1 = "", " And 填制日期 " & StrDate, StrFind_1) & _
               " And 打印状态 Not In (1,2) " & strCond & mstrShowBill & _
               " " & IIf(mstrSourceDep = "", "", " And A.对方部门id+0 in(" & mstrSourceDep & ") ") & _
               " Order by 优先级,姓名,No"
    
    Set recAutoPrint = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
         SQLCondition.date开始日期, _
         SQLCondition.date结束日期, _
         SQLCondition.str开始NO, _
         SQLCondition.str结束NO, _
         SQLCondition.str姓名, _
         SQLCondition.str就诊卡, _
         SQLCondition.str标识号, _
         SQLCondition.lng科室ID, _
         SQLCondition.str填制人, _
         SQLCondition.str审核人, _
         SQLCondition.lng药品ID, _
         SQLCondition.str当前NO, _
         lng药房ID)

    datCurr = zlDatabase.Currentdate()
        
    With recAutoPrint
        Do While Not .EOF
            '打印单据
            If DateDiff("s", !填制日期, datCurr) > mIntPrintDelay Then
                If intPrint > 0 Then
                    If mint自动配药 = 1 Then
                        '处理自动配药，在打印前完成
                        blnIgnore = False
                        
                        '检查是否需要配药
                        If Not IsDosage(Val(!单据), !NO) Then
                            blnIgnore = True
                        End If
                        
                        '检测是否允许
                        If CheckBill(1, Val(!单据), !NO) <> 0 Then
                            blnIgnore = True
                        End If
                        
                        If blnIgnore = False Then
                            gcnOracle.BeginTrans
                            blnInTrans = True
        
                            '再设置配药人
                            str操作员 = IIf(mstr自动配药人 <> "", mstr自动配药人, IIf(Str配药人 = "|当前操作员|", gstrUserName, Str配药人))
                            
                            gstrSQL = "zl_药品收发记录_设置配药人(" & lng药房ID & "," & Val(!单据) & ",'" & !NO & "','" & str操作员 & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-设置配药人")
                            
                            gstrSQL = " Update 未发药品记录 Set 打印状态=1 Where 单据=" & !单据 & " And No='" & !NO & "' And 库房ID=" & lng药房ID & " " & IIf(mstrSourceDep = "", "", " And 对方部门id+0 in(" & mstrSourceDep & ") ")
                            Call ExecuteProcedure(Me.Caption & "-更新单据已打印", False)
                            
                            '如果已启用了电子签名，则需要对配药人进行电子签名处理
                            If gbln药品使用电子签名 = True Then
                                If SaveSignatureRecored(EsignTache.Dosage, Val(!单据), !NO, lng药房ID) = False Then
                                    gcnOracle.RollbackTrans
                                    Exit Function
                                End If
                            End If
                            
                            gcnOracle.CommitTrans
                            blnInTrans = False
                            
                            mblnIsFirst = False
                        End If
                    Else
                        gstrSQL = " Update 未发药品记录 Set 打印状态=1 Where 单据=" & !单据 & " And No='" & !NO & "' And 库房ID=" & lng药房ID & " " & IIf(mstrSourceDep = "", "", " And 对方部门id+0 in(" & mstrSourceDep & ") ")
                        Call ExecuteProcedure(Me.Caption & "-更新单据已打印", False)
                    End If

                    strUnit = GetUnit(lng药房ID, !单据, !NO)
                    If Not BillHaveHerial(!NO, !单据) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "ReportFormat=2", "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "ReportFormat=2", "PrintEmpty=0", 2)
                    End If
                End If
                
                '打印药品标签
                If mintPrintDrugLable = 1 Then
                    If Not BillHaveHerial(!NO, !单据) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "性质=" & IIf(!单据 = 8, 1, 2), "药房=" & lng药房ID, "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "药房=" & lng药房ID, 2)
                    End If
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    BlnInRefresh = False
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, Optional ByVal bln提示 As Boolean = False) As Integer
    Dim dblCount As Double
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim RecCheck As New ADODB.Recordset
    '--根据将要执行的操作，判断是否允许--
    'IntOper:1-配药;2-取消配药;3-发药;4-退药;5-取消发药
    '返回:
    '0-允许操作
    '1-未配药
    '2-已配药
    '3-已发药
    '4-已删除
    '5-未发药
    
    '单独处理取消发药时的检查
    If IntOper = 5 Then
        gstrSQL = "Select 审核人 From 药品收发记录 Where No=[1] And 单据=[2] And 库房ID+0=[3] And 记录状态=1 And 审核人 IS Not Null And Rownum=1 "
        Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lng药房ID)
        If RecCheck.EOF Then
            CheckBill = 4
            MsgBox "未找到指定单据，或已被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
     
    gstrSQL = " Select A.配药人,A.审核人 From 药品收发记录 A" & _
        " Where A.No=[1] And A.单据=[2] " & _
        " " & IIf(IntOper <> 4, " And mod(A.记录状态,3)=1", "") & " And Rownum=1 " & _
        " And Nvl(Ltrim(Rtrim(A.摘要)),'小宝')<>'拒发' And (A.库房ID+0=[3] Or A.库房ID Is NULL)"
    
    If IntOper = 4 Then
        gstrSQL = gstrSQL & " And 审核人 IS Not Null"
    Else
        gstrSQL = gstrSQL & " And 审核人 IS Null"
    End If

    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lng药房ID)
    
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            If InStr(1, "123", IntOper) <> 0 Then CheckBill = 3: MsgBox "该处方已被其它操作员发药，" & IIf(IntOper = 1, "配药", IIf(IntOper = 2, "取消配药", IIf(IntOper = 3, "发药", "退药"))) & "操作中止！", vbInformation, gstrSysName: Exit Function
        Else
            If InStr(1, "4", IntOper) <> 0 Then CheckBill = 5: MsgBox "该处方还未发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            If Not IsNull(!配药人) Then
                If InStr(1, "1", IntOper) <> 0 Then CheckBill = 2: MsgBox "该处方已配药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            Else
                If InStr(1, "2", IntOper) <> 0 Then CheckBill = 1: MsgBox "该处方未配药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            End If
        End If
    End With
    
    '如果是退药，检查是否允许未作废医嘱退药
    If bln医嘱作废 = False And bln提示 Then
        intRows = Bill处方明细.Rows - 2
        For intRow = 1 To intRows
            dblCount = Val(Bill处方明细.TextMatrix(intRow, 列名.退药数))
            If dblCount <> 0 Then
                gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", Val(Bill处方明细.TextMatrix(intRow, 列名.Id)))

                If (rsTemp!扣率 Like "1*") Then       '临嘱
                    gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 病人费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(Bill处方明细.TextMatrix(intRow, 列名.Id)))
                    
                    If Not rsTemp.EOF Then
                        If (rsTemp!门诊标志 = 1 Or rsTemp!门诊标志 = 4) And rsTemp!医嘱序号 <> 0 Then
                            gstrSQL = "Select decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where ID=[1]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rsTemp!医嘱序号))
                            
                            If rsTemp!作废 = 0 Then
                                CheckBill = 1
                                MsgBox "第" & intRow & "行的药品记录对应的医嘱还未作废，不允许退药！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    CheckBill = 0
End Function

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim PrintRec As New ADODB.Recordset
    Dim strCond As String, strTemp As String
    Dim strCond2 As String, strCond1 As String, intLeft As Integer, intRight As Integer
    
    If Msf列表.Rows = 2 Then
        If Msf列表.TextMatrix(1, 处方列名.NO) = "" Then Exit Sub
    End If
    
    If mInt单据 = 0 Then
        strCond = " And A.单据 In (8,9)" '门诊及住院所有单据
    Else
        If mInt单据 = 8 Then
            strCond = " And A.单据 In (8,9) " '门诊划价及门诊记帐
        Else
            strCond = " And A.单据 = 9 " '住院记帐
        End If
    End If
    
    '在嵌套查询中，没有连接病人费用记录表，而条件中存在姓名字段时，需去掉该条件，因它用到病人费用记录表
    strCond1 = ""
    StrFind_4 = UCase(StrFind_4)
    strCond2 = StrFind_4
    intLeft = InStr(1, StrFind_4, " AND UPPER(H.姓名)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, StrFind_4, " AND")
        strTemp = Mid(StrFind_4, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(StrFind_4, intRight)
        Else
            strCond1 = Mid(StrFind_4, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(H.标识号)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(B.就诊卡号)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    
    '根据单据来设置报表数据的单位
    Const str售价 As String = "X.计算单位 单位,ltrim(to_char(S.零售价,'999990.00000')) 单价,ltrim(to_char(S.实际数量,'999990.00000')) 数量,LTRIM(TO_CHAR(S.已退数量,'999990.00000')) 已退数量,LTRIM(TO_CHAR(S.已发数量,'999990.00000')) 准退数,"
    Const str门诊 As String = "D.门诊单位 单位,ltrim(to_char(S.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 单价,ltrim(to_char(S.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 数量,LTRIM(TO_CHAR(S.已退数量/DECODE(D.门诊包装,NULL,1,0,1,D.门诊包装),'999990.00000')) 已退数量,LTRIM(TO_CHAR(S.已发数量/DECODE(D.门诊包装,NULL,1,0,1,D.门诊包装),'999990.00000')) 准退数,"
    Const str住院 As String = "D.住院单位 单位,ltrim(to_char(S.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 单价,ltrim(to_char(S.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 数量,LTRIM(TO_CHAR(S.已退数量/DECODE(D.住院包装,NULL,1,0,1,D.住院包装),'999990.00000')) 已退数量,LTRIM(TO_CHAR(S.已发数量/DECODE(D.住院包装,NULL,1,0,1,D.住院包装),'999990.00000')) 准退数,"
    Const str药库 As String = "D.药库单位 单位,ltrim(to_char(S.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 单价,ltrim(to_char(S.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 数量,LTRIM(TO_CHAR(S.已退数量/DECODE(D.药库包装,NULL,1,0,1,D.药库包装),'999990.00000')) 已退数量,LTRIM(TO_CHAR(S.已发数量/DECODE(D.药库包装,NULL,1,0,1,D.药库包装),'999990.00000')) 准退数,"
    Dim str单位串1 As String
    
    Select Case strUnit
    Case "售价单位"
        str单位串1 = str售价
    Case "门诊单位"
        str单位串1 = str门诊
    Case "住院单位"
        str单位串1 = str住院
    Case "药库单位"
        str单位串1 = str药库
    End Select
    
    '初始化打印单据
    With PrintRec
        If .State = 1 Then .Close
        If MnuEditHandback.Checked Then
            '发退药清单
            strCond = Replace(strCond, "A.", "S.")
            strCond1 = Replace(strCond1, "H.", "C.")
            If Chk清单.Value = 0 Then
                '##################汇总显示每笔记录还允许退多少##################
                gstrSQL = " Select 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, LTrim(To_Char(Sum(数量), '999990.00000')) 数量," & _
                         " LTrim(To_Char(Sum(已退数量), '999990.00000')) 已退数量, LTrim(To_Char(Sum(准退数), '999990.00000')) 准退数, " & _
                         " LTrim(To_Char(Sum(金额), '999990.00')) 金额, 发药时间, 发药人" & _
                         " From (SELECT Distinct DECODE(S.单据,8,'收费',9,'记帐') 类型,S.NO,P.名称 科室,C.姓名,C.性别,C.年龄,C.标识号 住院号,C.床号,'['||x.编码||']'||" & IIf(mblnTradeName, "NVL(A.名称,X.名称)", "X.名称") & " 品名," & _
                         " DECODE(x.规格,NULL,x.产地,DECODE(x.产地,NULL,x.规格,x.规格||'|'||x.产地)) 规格," & str单位串1 & _
                         " LTRIM(TO_CHAR(S.零售金额,'999990.00')) 金额,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,S.审核人 发药人,S.序号" & _
                         " FROM " & _
                         "      (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
                         "          NVL(A.付数,1) 付数,A.实际数量 实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                         "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.序号" & _
                         "      FROM" & _
                         "          (SELECT *" & _
                         "          FROM 药品收发记录 A" & _
                         "          WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                         "          AND A.库房ID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                         "          ) A," & _
                         "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                         "          FROM 药品收发记录 A" & _
                         "          WHERE A.审核人 IS NOT NULL" & _
                         "          AND A.库房ID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                         "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
                         "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0) S,"
                gstrSQL = gstrSQL & "" & _
                         "      病人费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,病人信息 B " & _
                         " WHERE S.药品ID=D.药品ID AND D.药品ID=X.ID And C.病人ID=B.病人ID(+) " & _
                         " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
                         " AND S.对方部门ID+0=P.ID " & strCond & strCond1 & _
                         " AND S.费用ID=C.ID  AND (S.记录状态=1 OR MOD(S.记录状态,3)=0)" & _
                         " AND S.审核人 IS NOT NULL AND S.库房ID+0=[13] AND S.实际数量*S.付数>S.已退数量) " & _
                         " Group By 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, 发药时间, 发药人 "
            Else
                '##################清单显示每笔操作过程##################
                gstrSQL = " Select 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, LTrim(To_Char(Sum(数量), '999990.00000')) 数量," & _
                         " LTrim(To_Char(Sum(已退数量), '999990.00000')) 已退数量, LTrim(To_Char(Sum(准退数), '999990.00000')) 准退数, " & _
                         " LTrim(To_Char(Sum(金额), '999990.00')) 金额, 发药时间, 发药人 From " & _
                         " (SELECT DECODE(S.单据,8,'收费',9,'记帐') 类型,S.NO,P.名称 科室,C.姓名,C.性别,C.年龄,C.标识号 住院号,C.床号,'['||X.编码||']'||" & IIf(mblnTradeName, "NVL(A.名称,X.名称)", "X.名称") & " 品名," & _
                         " DECODE(X.规格,NULL,X.产地,DECODE(X.产地,NULL,X.规格,X.规格||'|'||X.产地)) 规格," & str单位串1 & _
                         " LTRIM(TO_CHAR(S.零售金额,'999990.00')) 金额,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,S.审核人 发药人" & _
                         " FROM "
                gstrSQL = gstrSQL & _
                         "      (SELECT * FROM" & _
                         "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
                         "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                         "              A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作" & _
                         "          FROM" & _
                         "              (SELECT *" & _
                         "              FROM 药品收发记录 A" & _
                         "              WHERE A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                         "              AND A.库房ID+0=[13] " & _
                                        IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                         "              ) A," & _
                         "              (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                         "              FROM 药品收发记录 A" & _
                         "              WHERE A.审核人 IS NOT NULL " & _
                         "              AND A.库房ID+0=[13] " & _
                                        IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                         "              GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
                gstrSQL = gstrSQL & _
                         "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号)" & _
                         "          UNION" & _
                         "          SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0)," & _
                         "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态," & _
                         "          A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                         "          DECODE(A.记录状态,1,1,DECODE(MOD(A.记录状态,3),0,1,MOD(A.记录状态,3)+1)) 可操作" & _
                         "          FROM 药品收发记录 A" & _
                         "          WHERE A.审核人 IS NOT NULL AND NOT (记录状态=1 OR MOD(记录状态,3)=0)" & _
                         "          AND A.库房ID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.审核日期 " & StrDate & "", strCond2) & _
                         "          ) S,"
                gstrSQL = gstrSQL & "" & _
                         "      病人费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,病人信息 B " & _
                         " WHERE S.药品ID=D.药品ID AND D.药品ID=X.ID AND S.对方部门ID+0=P.ID " & _
                         " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 And C.病人ID=B.病人ID(+) " & _
                         " AND S.费用ID=C.ID " & strCond & strCond1 & " AND S.审核人 IS NOT NULL)  " & _
                         " Group By 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, 发药时间, 发药人 "
            End If
        
            Dim blnMoved As Boolean
            Dim str开始日期 As String, strSQL As String
            
            str开始日期 = IIf(strCond2 = "", StrFind_4, strCond2)
            '取开始日期:intRight保存单引号的起始位置
            intRight = InStr(1, str开始日期, "'") + 1
            str开始日期 = Mid(str开始日期, intRight, 19)
            '判断从开始日期后，是否存在转出的处方数据
            blnMoved = zlDatabase.DateMoved(str开始日期)
            
            '如果存在数据转出，则需要同时从后备表中提取数据
            If blnMoved Then
                strSQL = gstrSQL
                strSQL = Replace(strSQL, "药品收发记录", "H药品收发记录")
                strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
                gstrSQL = gstrSQL & " UNION ALL " & strSQL
            End If
            
            If Chk清单.Value = 0 Then
                gstrSQL = gstrSQL & " ORDER BY NO,类型"
            Else
                gstrSQL = gstrSQL & " ORDER BY NO,类型,发药时间"
            End If
        Else
            '未发药清单
            Const str售价1 As String = "C.计算单位 单位,ltrim(to_char(B.零售价,'999990.00000')) 单价,ltrim(to_char(B.实际数量,'999990.00000')) 数量,"
            Const str门诊1 As String = "D.门诊单位 单位,ltrim(to_char(B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 数量,"
            Const str住院1 As String = "D.住院单位 单位,ltrim(to_char(B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 数量,"
            Const str药库1 As String = "D.药库单位 单位,ltrim(to_char(B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 数量,"
            
            Select Case strUnit
            Case "售价单位"
                str单位串 = str售价1
            Case "门诊单位"
                str单位串 = str门诊1
            Case "住院单位"
                str单位串 = str住院1
            Case "药库单位"
                str单位串 = str药库1
            End Select
            gstrSQL = "Select 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, LTrim(To_Char(Sum(数量), '999990.00000')) 数量," & _
                     " LTrim(To_Char(Sum(金额), '999990.00')) 金额, 填制人, 填制日期, 配药人 From " & _
                     " (SELECT DECODE(A.单据,8,'收费',9,'记帐') 类型,A.NO," & _
                     " T.名称 科室,H.姓名,H.性别,H.年龄,H.标识号 住院号,H.床号," & _
                     " '['||c.编码||']'||C.名称 品名,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & str单位串 & _
                     " LTRIM(TO_CHAR(B.零售金额,'999990.00')) 金额,B.填制人,B.填制日期,DECODE(B.配药人,'部门发药','',NULL,'',B.配药人) 配药人" & _
                     " FROM 药品收发记录 B,药品规格 D,收费项目目录 C,病人费用记录 H,部门表 S,部门表 T,未发药品记录 A" & _
                     " WHERE D.药品ID=C.ID AND A.库房ID+0=[13] " & IIf(Str窗口 = "", "", " AND (A.发药窗口 IN(" & Str窗口 & ") Or A.发药窗口 Is NULL)") & _
                     " " & IIf(StrFind_1 = "", " AND A.填制日期 " & StrDate, StrFind_1) & _
                     " " & strCond & mstrShowBill & _
                     " AND B.审核人 IS NULL AND LTRIM(RTRIM(NVL(B.摘要,'小宝')))<>'拒发' " & _
                     " AND H.开单部门ID=T.ID AND B.药品ID=D.药品ID AND MOD(B.记录状态,3)=1" & _
                     " AND S.ID=B.库房ID AND B.费用ID=H.ID AND B.NO=A.NO AND B.单据=A.单据 AND B.库房ID+0=[13]) " & _
                     " Group By 类型, NO, 科室, 姓名, 性别, 年龄, 住院号, 床号, 品名, 规格, 单位, 单价, 填制人, 填制日期, 配药人 " & _
                     " ORDER BY 类型, NO"
        End If
    End With
    
    Set PrintRec = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.date开始日期, _
        SQLCondition.date结束日期, _
        SQLCondition.str开始NO, _
        SQLCondition.str结束NO, _
        SQLCondition.str姓名, _
        SQLCondition.str就诊卡, _
        SQLCondition.str标识号, _
        SQLCondition.lng科室ID, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        SQLCondition.lng药品ID, _
        SQLCondition.str当前NO, _
        lng药房ID)
    
    With PrintRec
        If .EOF Then Exit Sub
        Set MsfPrint.DataSource = PrintRec
    End With
    
    With MsfPrint
        .FixedCols = 0
        For intLeft = 0 To .Cols - 1
            .ColAlignmentFixed(intLeft) = 4
        Next
        
        .ColWidth(0) = 500
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
        .ColWidth(4) = 500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 2500
        .ColWidth(9) = 500
        .ColWidth(10) = 600
        '以下
        If MnuEditHandback.Checked Then
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1000
            .ColWidth(16) = 1000
            .ColWidth(17) = 1000
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 7
            .ColAlignment(15) = 7
        Else
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1000
            .ColWidth(16) = 1000
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
        End If
    End With
    
    ObjAppRow.Add "打印人:" & gstrUserName
    ObjAppRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = "药品处方单"
    Set objPrint.Body = MsfPrint
    
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
End Sub

Private Sub 权限控制()
    '配药、发药、退药、参数控制
    If Not IsHavePrivs(mstrPrivs, "配药") Then
        MnuEditDosage.Visible = False
        MnuEditAbolish.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "发药") Then
        mnuEditBill.Visible = False
        MnuEditBatch.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "退药") Then
        mnuEditBillRestore.Visible = False
        If MnuEditBatch.Visible = False Then MnuEdit1.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "打印本次退药明细") Then
        mnuFileRestore.Visible = False
        MnuFile2.Visible = MnuFileBillprint.Visible
    End If
    If Not IsHavePrivs(mstrPrivs, "打印已发药清单") Then
        mnuFileReport.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "退其它药房的处方") Then
        MnuEditHandbackBatch.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "发其它药房的处方") Then
        MnuEditSendOther.Visible = False
    End If
    intVerify = IIf(IsHavePrivs(mstrPrivs, "校验处方"), 1, 0)
    If Not IsHavePrivs(strChargePrivs, "划价") Then
        mnuCharge.Visible = False
        Tbar1.Buttons("Charge").Visible = False
    End If
    If Trim(strStuffPrivs) = "" Then
        mnuStuff.Visible = False
        Tbar1.Buttons("Stuff").Visible = False
    End If
    If gblnPass And IsHavePrivs(mstrPrivs, "合理用药监测") Then
        mblnStarPass = True
    End If
    
    mbln发病区处方 = IsHavePrivs(mstrPrivs, "发病区处方")
    img病区.Visible = mbln发病区处方
    
    mnuCancel.Visible = mbln允许取消发药 And IsHavePrivs(mstrPrivs, "取消发药")
    Tbar1.Buttons("Cancel").Visible = mbln允许取消发药 And IsHavePrivs(mstrPrivs, "取消发药")
    
    Tbar1.Buttons(9).Visible = (mnuCharge.Visible Or mnuStuff.Visible Or mnuCancel.Visible)
End Sub

Private Sub Txt开单医生_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Txt开单医生.ListIndex <> 0 Then Call CmdSend_Click
End Sub

Private Sub Txt开单医生_KeyPress(KeyAscii As Integer)
    Dim IntMatchIdx As Integer
    
    With Txt开单医生
        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
        If IntMatchIdx = -2 Then Exit Sub
        .ListIndex = IntMatchIdx
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub

Private Function 转换退药串() As String
    Dim strCond1 As String, strCond2 As String, strTemp As String
    Dim intRight As Integer, intLeft As Integer
    '在嵌套查询中，没有连接病人费用记录表，而条件中存在姓名字段时，需去掉该条件，因它用到病人费用记录表
    strCond1 = ""
    StrFind_4 = UCase(StrFind_4)
    strCond2 = StrFind_4
    intLeft = InStr(1, strCond2, " AND UPPER(H.姓名)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, StrFind_4, " AND")
        strTemp = Mid(StrFind_4, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(StrFind_4, intRight)
        Else
            strCond1 = Mid(StrFind_4, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(H.标识号)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(B.就诊卡号)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    转换退药串 = strCond2
End Function

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub ShowStock()
    Dim intUnit As Integer
    Dim lng药品ID As Long, lng批次 As Long
    Dim str单位 As String, str包装 As String
    Dim rsStock As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
'    stbThis.Panels(2).Text = ""
    
    If TxtNo.ListIndex < 0 Then Exit Sub
    If Trim(TxtNo.Text) = "" Then Exit Sub
    
    strUnit = GetUnit(lng药房ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    lng药品ID = Val(Bill处方明细.TextMatrix(Bill处方明细.Row, 列名.药品ID))
    lng批次 = Val(Bill处方明细.TextMatrix(Bill处方明细.Row, 列名.批次))
    
    Select Case strUnit
    Case "售价单位"
        str单位 = "C.计算单位"
        str包装 = "/1"
    Case "门诊单位"
        str单位 = "B.门诊单位"
        str包装 = "/B.门诊包装"
    Case "住院单位"
        str单位 = "B.住院单位"
        str包装 = "/B.住院包装"
    Case "药库单位"
        str单位 = "B.药库单位"
        str包装 = "/B.药库包装"
    End Select
    
    gstrSQL = " Select A.实际数量" & str包装 & " 实际数量," & str单位 & " 单位" & _
             " From 药品库存 A,药品规格 B,收费项目目录 C" & _
             " Where A.药品ID=B.药品ID And B.药品ID=C.ID And A.性质=1 " & _
             " And A.药品ID=[2] And Nvl(A.批次,0)=[3] And A.库房ID=[1]"
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取药品库存]", lng药房ID, lng药品ID, lng批次)
    
    If rsStock.EOF Then Exit Sub
    
    If Me.ActiveControl Is Bill处方明细 Then
        stbThis.Panels(2).Text = "当前库存：" & FormatEx(rsStock!实际数量, 5) & rsStock!单位
    End If
End Sub

Private Sub Get单位串()
    Const str售价 As String = "C.计算单位 As 售价单位,C.计算单位 As 单位,1 As 包装,ltrim(to_char(B.零售价,'999990.00000')) 单价,ltrim(to_char(B.实际数量,'999990.00000')) 数量"
    Const str门诊 As String = "C.计算单位 As 售价单位,D.门诊单位 As 单位,D.门诊包装 As 包装,ltrim(to_char(B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装),'999990.00000')) 数量"
    Const str住院 As String = "C.计算单位 As 售价单位,D.住院单位 As 单位,D.住院包装 As 包装,ltrim(to_char(B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装),'999990.00000')) 数量"
    Const str药库 As String = "C.计算单位 As 售价单位,D.药库单位 As 单位,D.药库包装 As 包装,ltrim(to_char(B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 单价,ltrim(to_char(B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装),'999990.00000')) 数量"
    
    Select Case strUnit
    Case "售价单位"
        str单位串 = str售价
    Case "门诊单位"
        str单位串 = str门诊
    Case "住院单位"
        str单位串 = str住院
    Case "药库单位"
        str单位串 = str药库
    End Select
End Sub
Private Function BillHaveHerial(ByVal strNo As String, ByVal int单据 As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH

    gstrSQL = "Select NO From 病人费用记录 Where NO=[1] And 记录状态 IN(0,1,3)" & _
        " And 记录性质=[3] And 收费类别='7' And 执行部门ID+0=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng药房ID, IIf(int单据 = 8, 1, 2))
    
    BillHaveHerial = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDateSQL(ByVal strInput As String) As String
    Dim lngStart As Long
    Dim blnDefault As Boolean
    '分解SQL，保留原来的日期条件
    strInput = Trim(UCase(strInput))
    If strInput = "" Then
        blnDefault = True
    Else
        lngStart = InStr(1, strInput, " AND TO_DATE(")
        If lngStart <> 0 Then
            lngStart = InStr(lngStart + 4, strInput, " AND")
            If lngStart <> 0 Then
                strInput = Mid(strInput, 1, lngStart)
            End If
        Else
            blnDefault = True
        End If
    End If
    If blnDefault Then
        If MnuEditDosage.Checked Then
            GetDateSQL = " And A.填制日期 " & StrDate
        ElseIf MnuEditAbolish.Checked Then
            GetDateSQL = " And A.填制日期 " & StrDate
        ElseIf MnuEditConsignment.Checked Then
            GetDateSQL = " And A.填制日期 " & StrDate
        Else
            GetDateSQL = " And A.审核日期 " & StrDate
        End If
    Else
        GetDateSQL = strInput
    End If
End Function

Private Function ReLocateRow() As Long
    Dim lngRow As Long, lngRows As Long
    On Error GoTo ErrHand
    
    '定位上次选择的处方，失败返回1
    lngRows = Msf列表.Rows - 1
    For lngRow = 1 To lngRows
        If Val(Msf列表.TextMatrix(lngRow, 处方列名.单据)) = IntLastBill And _
            Msf列表.TextMatrix(lngRow, 处方列名.NO) = StrLastNo And _
            Msf列表.TextMatrix(lngRow, 处方列名.日期) = strLastData Then
            ReLocateRow = lngRow
            Exit Function
        End If
    Next
ErrHand:
    ReLocateRow = IIf(LngSendRow > Msf列表.Rows - 1 Or LngSendRow = 0, 1, LngSendRow)
End Function

Private Function ReLocateDetailRow() As Long
    Dim lngRow As Long, lngRows As Long
    On Error GoTo ErrHand
    
    '定位上次选择的处方明细列表，失败返回1
    lngRows = Bill处方明细.Rows - 1
    
    If mintLastSequence = 0 Then
        ReLocateDetailRow = lngRows
        Exit Function
    End If
    
    For lngRow = 1 To lngRows
        If Val(Bill处方明细.TextMatrix(lngRow, 列名.序号)) = mintLastSequence Then
            ReLocateDetailRow = lngRow
            Exit Function
        End If
    Next
ErrHand:
    ReLocateDetailRow = 1
End Function
Private Function GetDetailCol(ByVal strText As String) As Integer
    Dim intCol As Integer, intCols As Integer
    intCols = Bill处方明细.Cols - 1
    If strText = "用量" Then strText = "单量"
    For intCol = 0 To intCols
        If Trim(Bill处方明细.TextMatrix(0, intCol)) = strText Then
            GetDetailCol = intCol
            Exit Function
        End If
    Next
    GetDetailCol = -1
End Function

Private Function IsDosage(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    Dim int门诊 As Integer, int配药 As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '检查当前处方是否需要经过配药过程
    
    If int单据 = 0 Then Exit Function
    If strNo = "" Then Exit Function
    
    '取当前处方的病人来源
    gstrSQL = " Select 门诊标志 From 病人费用记录 " & _
              " Where ID=(" & _
              "     Select 费用ID From 药品收发记录 " & _
              "     Where (Nvl(库房ID,0)=[3] Or Nvl(库房ID,0)=0) And 单据=[2] And NO=[1] And Rownum<2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取病人来源]", strNo, int单据, lng药房ID)
    
    int门诊 = IIf(rsTemp!门诊标志 = 1 Or rsTemp!门诊标志 = 4, 1, 2)
    
    '根据当前单据判断是否需要配药
    gstrSQL = "Select Nvl(配药,0) AS 配药 From 药房配药控制 Where 药房ID=[1] And Nvl(门诊,1)=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[根据当前单据判断是否需要配药]", lng药房ID, int门诊)
        
    If rsTemp.RecordCount = 0 Then Exit Function
    
    IsDosage = (rsTemp!配药 = 1)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetForeColor_ROW(ByVal lngRow As Long, ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer, intRow As Integer
    '设置某行的颜色
    With Bill处方明细
        BlnEnterCell = False
        intCol = .Col
        intRow = .Row
        .Row = lngRow
        For i = 0 To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        .Row = intRow
        BlnEnterCell = True
    End With
End Sub

Private Sub GetBillSequence()
    Dim intRow As Integer, intRows As Integer
    Dim int序号 As Integer
    '获取当前待发药、待退药处方的有效序号
    str序号 = ""
    intRows = Bill处方明细.Rows - 2
    
    If MnuEditHandback.Checked Then
        '退药数不为零表示本次要退的明细，仅统计出这类明细的序号
        For intRow = 1 To intRows
            If Val(Bill处方明细.TextMatrix(intRow, 列名.退药数)) <> 0 Then
                int序号 = Val(Bill处方明细.TextMatrix(intRow, 列名.序号))
                If InStr(1, str序号 & ",", "," & int序号 & ",") = 0 Then
                    str序号 = str序号 & "," & int序号
                End If
            End If
        Next
    Else
        For intRow = 1 To intRows
            int序号 = Val(Bill处方明细.TextMatrix(intRow, 列名.序号))
            If InStr(1, str序号 & ",", "," & int序号 & ",") = 0 Then
                str序号 = str序号 & "," & int序号
            End If
        Next
    End If
    If str序号 <> "" Then str序号 = Mid(str序号, 2)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

