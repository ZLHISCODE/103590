VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm部门发药管理 
   Caption         =   "药品部门发药"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   11760
   DrawMode        =   14  'Copy Pen
   Icon            =   "Frm部门发药管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11760
   Begin VB.TextBox txt留存数 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   225
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   23
      Text            =   "####"
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox Chk显示退药待发单据 
      Appearance      =   0  'Flat
      Caption         =   "显示退药待发单据"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8520
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlley 
      Caption         =   "过敏史/病生状态"
      Height          =   270
      Left            =   6840
      TabIndex        =   13
      Top             =   4440
      Width           =   1530
   End
   Begin VB.CheckBox Chk清单 
      Appearance      =   0  'Flat
      Caption         =   "显示所有过程单据"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1845
   End
   Begin VB.TextBox TxtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Text            =   "####"
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox Cbo批号 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   3990
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "未发药清单(&N)"
      TabPicture(0)   =   "Frm部门发药管理.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl配药人"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl发药单格式"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Dtp查询日期"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Bill未发药清单"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo配药人"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbo发药单格式"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "汇总发药(&T)"
      TabPicture(1)   =   "Frm部门发药管理.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill汇总发药"
      Tab(1).Control(1)=   "Bill退药销帐"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "缺药清单(&Q)"
      TabPicture(2)   =   "Frm部门发药管理.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Bill缺药清单"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "拒发药清单(&D)"
      TabPicture(3)   =   "Frm部门发药管理.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Bill拒发药清单"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "发退药清单(&A)"
      TabPicture(4)   =   "Frm部门发药管理.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Bill已发药清单"
      Tab(4).ControlCount=   1
      Begin VB.ComboBox cbo发药单格式 
         Height          =   300
         Left            =   8300
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1900
      End
      Begin VB.ComboBox cbo配药人 
         Height          =   300
         Left            =   720
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "cbo配药人"
         Top             =   2760
         Width           =   1900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill未发药清单 
         Height          =   2145
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "按空格键或鼠标右键切换单据状态"
         Top             =   360
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   3784
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill汇总发药 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill缺药清单 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   6
         ToolTipText     =   "按空格键或鼠标右键切换单据状态"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill拒发药清单 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   7
         ToolTipText     =   "按空格键或鼠标右键切换单据状态"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill已发药清单 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   8
         ToolTipText     =   "按空格键或鼠标右键切换单据状态"
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker Dtp查询日期 
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   99680259
         CurrentDate     =   36985
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill退药销帐 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl发药单格式 
         AutoSize        =   -1  'True
         Caption         =   "发药单格式"
         Height          =   180
         Left            =   7300
         TabIndex        =   22
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label lbl配药人 
         AutoSize        =   -1  'True
         Caption         =   "配药人"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   540
      End
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1164
      BandCount       =   2
      _CBWidth        =   11760
      _CBHeight       =   660
      _Version        =   "6.7.8988"
      Child1          =   "Tbar"
      MinWidth1       =   4005
      MinHeight1      =   600
      Width1          =   4770
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "发药药房"
      Child2          =   "Cbo发药药房"
      MinWidth2       =   1695
      MinHeight2      =   300
      Width2          =   1695
      NewRow2         =   0   'False
      Begin VB.ComboBox Cbo发药药房 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   9975
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   180
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   600
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   8805
         _ExtentX        =   15531
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
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发药"
               Key             =   "Consignment"
               Object.ToolTipText     =   "发药"
               Object.Tag             =   "发药"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "申领"
               Key             =   "Desire"
               Object.ToolTipText     =   "缺药申领"
               Object.Tag             =   "申领"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "拒发"
               Key             =   "Handback"
               Object.ToolTipText     =   "拒发"
               Object.Tag             =   "拒发"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退药"
               Key             =   "Restore"
               Object.ToolTipText     =   "退药"
               Object.Tag             =   "退药"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "销账"
               Key             =   "ReVerify"
               Object.ToolTipText     =   "销账"
               Object.Tag             =   "销账"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit1"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
            EndProperty
         EndProperty
         Begin VB.Timer TimerAuto 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   6960
            Top             =   240
         End
         Begin MSComctlLib.ImageList imgPass 
            Left            =   5415
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
                  Picture         =   "Frm部门发药管理.frx":1D86
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm部门发药管理.frx":2040
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm部门发药管理.frx":22FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm部门发药管理.frx":25B4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Frm部门发药管理.frx":286E
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
      TabIndex        =   3
      Top             =   7800
      Width           =   11760
      _ExtentX        =   20743
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
            Object.Width           =   15663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "Frm部门发药管理.frx":2B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm部门发药管理.frx":2E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw科室 
      Height          =   465
      Left            =   10680
      TabIndex        =   62
      Top             =   4680
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   820
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw给药途径 
      Height          =   345
      Left            =   10680
      TabIndex        =   63
      Top             =   5400
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw剂型 
      Height          =   345
      Left            =   10680
      TabIndex        =   64
      Top             =   6000
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame fraCondition 
      Height          =   3000
      Left            =   20
      TabIndex        =   15
      Top             =   620
      Visible         =   0   'False
      Width           =   12975
      Begin VB.Frame frmLine1 
         Height          =   30
         Left            =   0
         TabIndex        =   25
         Top             =   2300
         Width           =   12945
      End
      Begin VB.Frame fraConRequest 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   20
         TabIndex        =   57
         Top             =   2280
         Visible         =   0   'False
         Width           =   10770
         Begin MSComCtl2.DTPicker Dtp销帐结束时间 
            Height          =   315
            Left            =   4200
            TabIndex        =   58
            Top             =   255
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp销帐开始时间 
            Height          =   315
            Left            =   1320
            TabIndex        =   59
            Top             =   255
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin VB.Label lblS1 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   3840
            TabIndex        =   61
            Top             =   315
            Width           =   180
         End
         Begin VB.Label lblTimeRequest 
            AutoSize        =   -1  'True
            Caption         =   "销帐申请时间"
            Height          =   180
            Left            =   120
            TabIndex        =   60
            Top             =   315
            Width           =   1080
         End
      End
      Begin VB.Frame frmLine 
         Height          =   30
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Width           =   12945
      End
      Begin VB.Frame fraConNormal 
         BorderStyle     =   0  'None
         Height          =   1320
         Left            =   20
         TabIndex        =   26
         Top             =   100
         Width           =   10800
         Begin VB.CheckBox chkSend 
            Caption         =   "院内用药"
            Height          =   180
            Index           =   0
            Left            =   960
            TabIndex        =   34
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSend 
            Caption         =   "自取药"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   33
            Top             =   960
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.Frame fraTypeLine 
            Height          =   400
            Left            =   4200
            TabIndex        =   32
            Top             =   795
            Width           =   30
         End
         Begin VB.TextBox txt科室 
            Height          =   300
            Left            =   3640
            TabIndex        =   31
            Top             =   545
            Width           =   6705
         End
         Begin VB.TextBox txtPati 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8790
            TabIndex        =   30
            Top             =   180
            Width           =   1905
         End
         Begin VB.CheckBox chkSendType 
            Caption         =   "发药类型，动态增加"
            Height          =   180
            Index           =   0
            Left            =   4320
            TabIndex        =   29
            Top             =   963
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkSend 
            Caption         =   "离院带药"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   28
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton cmd部门类型 
            Caption         =   "…"
            Height          =   300
            Left            =   10320
            TabIndex        =   27
            Top             =   530
            Width           =   375
         End
         Begin MSComCtl2.DTPicker Dtp结束时间 
            Height          =   315
            Left            =   3640
            TabIndex        =   35
            Top             =   180
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker Dtp开始时间 
            Height          =   315
            Left            =   960
            TabIndex        =   36
            Top             =   180
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   99680259
            CurrentDate     =   36985
         End
         Begin MSComctlLib.TabStrip tbsType 
            Height          =   255
            Left            =   960
            TabIndex        =   37
            Top             =   570
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "临床"
                  Key             =   "T1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "医技"
                  Key             =   "T2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "病区"
                  Key             =   "T3"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblPatiInputType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院号↓"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8040
            TabIndex        =   43
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病人信息"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7200
            TabIndex        =   42
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblS 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   3400
            TabIndex        =   41
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblDepType 
            AutoSize        =   -1  'True
            Caption         =   "领药部门"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl发药类型 
            AutoSize        =   -1  'True
            Caption         =   "发药类型"
            Height          =   180
            Left            =   120
            TabIndex        =   39
            Top             =   963
            Width           =   720
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            Caption         =   "时间范围"
            Height          =   180
            Left            =   120
            TabIndex        =   38
            Top             =   247
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   345
         Left            =   10850
         TabIndex        =   18
         Top             =   960
         Width           =   900
      End
      Begin VB.CommandButton cmdOtherCon 
         Caption         =   "全部条件(&C)"
         Height          =   345
         Left            =   11750
         TabIndex        =   17
         Top             =   960
         Width           =   1140
      End
      Begin VB.Frame fraConExpand 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   20
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   10800
         Begin VB.CheckBox chkType 
            Caption         =   "婴儿药品"
            Height          =   180
            Index           =   1
            Left            =   9600
            TabIndex        =   67
            Top             =   675
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkType 
            Caption         =   "病人药品"
            Height          =   180
            Index           =   0
            Left            =   8400
            TabIndex        =   66
            Top             =   675
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton cmd药品剂型 
            Caption         =   "…"
            Height          =   300
            Left            =   10320
            TabIndex        =   52
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmd给药途径 
            Caption         =   "…"
            Height          =   300
            Left            =   4575
            TabIndex        =   51
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txt药品剂型 
            Height          =   300
            Left            =   6120
            TabIndex        =   50
            Top             =   255
            Width           =   4215
         End
         Begin VB.TextBox txt给药途径 
            Height          =   300
            Left            =   960
            TabIndex        =   49
            Top             =   255
            Width           =   3615
         End
         Begin VB.ComboBox Cbo医嘱类型 
            Height          =   300
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   615
            Width           =   1815
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "退药请求"
            Height          =   225
            Index           =   2
            Left            =   3840
            TabIndex        =   47
            Top             =   675
            Width           =   1125
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "发药请求"
            Height          =   225
            Index           =   1
            Left            =   2400
            TabIndex        =   46
            Top             =   675
            Width           =   1125
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "所有请求"
            Height          =   225
            Index           =   0
            Left            =   960
            TabIndex        =   45
            Top             =   675
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Label lbl给药途径 
            AutoSize        =   -1  'True
            Caption         =   "给药途径"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   315
            Width           =   720
         End
         Begin VB.Label lbl药品剂型 
            AutoSize        =   -1  'True
            Caption         =   "药品剂型"
            Height          =   180
            Left            =   5280
            TabIndex        =   55
            Top             =   315
            Width           =   720
         End
         Begin VB.Label lbl处理条件 
            AutoSize        =   -1  'True
            Caption         =   "处理范围"
            Height          =   180
            Left            =   120
            TabIndex        =   54
            Top             =   675
            Width           =   720
         End
         Begin VB.Label Lbl医嘱类型 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "医嘱类型"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5280
            TabIndex        =   53
            Top             =   675
            Width           =   720
         End
      End
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
         Caption         =   "单据打印(&B)"
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintTotal 
         Caption         =   "打印汇总清单(&C)"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "打印退药通知单(&R)"
      End
      Begin VB.Menu mnuFileWait 
         Caption         =   "打印药品摆药单(&W)"
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
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu MnuEditVerify 
         Caption         =   "发药(&V)"
      End
      Begin VB.Menu MnuEditDesire 
         Caption         =   "缺药申领(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditHandback 
         Caption         =   "拒发确认(&H)"
      End
      Begin VB.Menu MnuEditRestore 
         Caption         =   "退药(&R)"
      End
      Begin VB.Menu mnuEditHandbackBatch 
         Caption         =   "退其它药房的处方(&T)"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReVerify 
         Caption         =   "药品退药销账(&B)"
      End
      Begin VB.Menu mnuFlag 
         Caption         =   "停止发药标记(&S)"
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
         Caption         =   "字体(&Z)"
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "小字体(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "中字体(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSet 
            Caption         =   "大字体(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewLocate 
         Caption         =   "查找(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuViewLocateNext 
         Caption         =   "查找下一条(&N)"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuView4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewTotal 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu MnuViewNone 
         Caption         =   "全清(&C)"
      End
      Begin VB.Menu MnuView5 
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
   Begin VB.Menu PopMenu_1 
      Caption         =   "PopMenu未发药"
      Visible         =   0   'False
      Begin VB.Menu Consignment 
         Caption         =   "发药(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu HandBack 
         Caption         =   "拒发(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Lack 
         Caption         =   "缺药(&L)"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu Nop_1 
         Caption         =   "不处理(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Split_1 
         Caption         =   "-"
      End
      Begin VB.Menu ConsignmentALL 
         Caption         =   "全部发药(&S)"
      End
      Begin VB.Menu HandBackALL 
         Caption         =   "全部拒发(&J)"
      End
      Begin VB.Menu Nop_ALL 
         Caption         =   "全部不处理(&B)"
      End
   End
   Begin VB.Menu PopMenu_2 
      Caption         =   "PopMenu已发药"
      Visible         =   0   'False
      Begin VB.Menu Restore 
         Caption         =   "退药(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Nop_2 
         Caption         =   "不处理(&N)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu PopMenu_3 
      Caption         =   "PopMenu拒发药"
      Visible         =   0   'False
      Begin VB.Menu ResumeDo 
         Caption         =   "恢复(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu Nop_3 
         Caption         =   "不处理(&N)"
         Checked         =   -1  'True
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
   Begin VB.Menu mnuColHide 
      Caption         =   "ColHide"
      Visible         =   0   'False
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "药品(编码和名称)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "药品(仅编码)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuDrugCodeName 
         Caption         =   "药品(仅名称)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuColHideLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "其它名"
         Index           =   0
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "英文名"
         Index           =   1
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "科室"
         Index           =   2
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "开单医生"
         Index           =   3
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "状态"
         Index           =   4
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "类型"
         Index           =   5
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "NO"
         Index           =   6
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "记帐员"
         Index           =   7
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "床号"
         Index           =   8
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "姓名"
         Index           =   9
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "住院号"
         Index           =   10
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "规格"
         Index           =   11
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "产地"
         Index           =   12
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "批号"
         Index           =   13
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "付"
         Index           =   14
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "数量"
         Index           =   15
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "已退数"
         Index           =   16
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "准退数"
         Index           =   17
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "退药数"
         Index           =   18
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "单价"
         Index           =   19
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "金额"
         Index           =   20
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "单量"
         Index           =   21
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "频次"
         Index           =   22
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "用法"
         Index           =   23
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "记帐时间"
         Index           =   24
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "说明"
         Index           =   25
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "操作员"
         Index           =   26
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "发药时间"
         Index           =   27
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "领/退药人"
         Index           =   28
      End
      Begin VB.Menu mnuBillItem 
         Caption         =   "库房货位"
         Index           =   29
      End
   End
   Begin VB.Menu mnuPatiInfo 
      Caption         =   "病人信息"
      Visible         =   0   'False
      Begin VB.Menu mnuInfoItem 
         Caption         =   "住院号(&0)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "姓名(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "床号(&2)"
         Index           =   2
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "单据号(&3)"
         Index           =   3
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "病人ID(&4)"
         Index           =   4
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "就诊卡(&5)"
         Index           =   5
      End
   End
   Begin VB.Menu mnuType 
      Caption         =   "分类说明"
      Visible         =   0   'False
      Begin VB.Menu mnuTypeItem 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Frm部门发药管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--常规变量--

'Public strPart As String                                '选择部门,用于显示于自定义报表中
Public BlnSetPara As Boolean                            '参数设置窗体是否确定后退出
Public BlnRefresh As Boolean                            '其他窗体是否处理了数据,是则刷新


'--查询条件变量--
Private mstr开始日期_未发 As String                        '当前时间
Private mstr结束日期_未发 As String                       '结束时间
Private mstr开始日期_已发 As String
Private mstr结束日期_已发 As String
Private mlng病人ID As Long
Private mstr住院号 As String
Private mstr病人姓名 As String
Private mstrSerchNO As String
Private mstr开始NO As String
Private mstr结束NO As String
Private mstrDrug As String                                '药品剂型
Private mstrUse As String                                 '给药途径
Private mstr部门 As String                                '选择部门
Private mstr部门名称 As String
Private mint类型 As Integer                               '选择部门类型
Private mint范围 As String                                '选择发药范围
Private mstr床号 As String                                '床号
Private mstr发药类型 As String
Private mint病人类型 As Integer                           '0-病人;1-婴儿;2－病人和婴儿

'--参数变量--
Private IntCheckStock As Integer                        '检测库存
Private Int允许未审核处方发药 As Integer                '未审核是否允许发药
Private lng药房ID As Long                               '部门发药
Private Lng操作模式 As Long                             '操作处方单、摆药单、兼有
Private Lng医嘱类型 As Long                             '隶属于Lng操作模式参数，当Lng操作模式=处方单或所有时，本参数才生效（所有、长嘱、临嘱、记帐单、所有医嘱）
Private int离院带药 As Integer                          '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
Private Lng汇总显示 As Long                             '是否按科室汇总显示汇总清单
Private Lng自动打印 As Long                             '发药后是否自动打印
Private Lng缺药检查 As Long                             '如果进行缺药检查,则无库存药品不允许发药
Private mlng待发单据 As Long                            '是否显示退药待发单据
Private intDays As Integer                              '查询天数
'Private IntSendAfterDosage As Integer                  '未配药处方是否允许发药(只处理门诊处方)
Private intFont As Integer                              '字体
Private StrFindStyle As String                          '输入匹配
Private lng未发药记录 As Long
Private BlnEnterCell As Boolean                         '是否激活ENTERCELL()事件
Private int药品名称 As Integer                          '药品名称显示格式：0-编码与名称;1-仅编码;2-仅名称
Private Lng领药人签名 As Long
Private Lng退药人签名 As Long
Private str记帐人 As String
Private mblnStarPass As Boolean                         '启用合理用药(PASS)
Private int发药规则 As Integer                          '0-全额实发 1-零实发 2-部分分零满足
Private int金额保留位数 As Integer                      '费用金额保留位数
Private int审核划价单 As Integer                        '执行后自动审核划价单
Private mstr毒理分类 As String
Private mstr价值分类 As String
Private mstr病区发药方式 As String                      '部门性质
Private mint自动刷新未发药清单 As Integer               '0-不自动刷新
Private mdate上次刷新时间 As Date                       '记录上次刷新时系统时间
Private mblnAllConditon As Boolean                      '条件状态：
Private mbln药品储备 As Boolean                         '是否显示库房货位及库存限量提示
Private mbln显示领退药人 As Boolean                     '是否显示领药或退药人
Private mbln汇总发药 As Boolean                         '汇总发药时是否一并处理退药销帐记录

'--本程序使用变量--
Private strUnit As String                               '单位串
Private BlnStartUp As Boolean                           '启动成功
Private BlnFirstStart As Boolean                        '第一次启动
Private mblnFirstSended As Boolean
Private Bln刷新未发药清单 As Boolean                    '决定未发药清单内容是否在显示前刷新
Private Bln检测库存 As Boolean
Private BlnInRefresh As Boolean                         '正处于刷新状态
Private str排序_未发药 As String                        '排序列
Private str排序_发退药 As String                        '排序列
Private bln医嘱作废 As Boolean                          '未作废医嘱是否允许退药
Private str领药人 As String
Private str退药人 As String
Private mstr价格失效提示 As String
Private LngLastRow As Long
Private lngLastCol As Long
Private bln药品留存入出类别 As Boolean
Private mstrDrawDept As String                          '临时记录领药部门
Private mbln低分辨率 As Boolean                         '判断是否是低分辨率（800×600）
Private mlng汇总发药号 As Variant
Private Const mstrAllType As String = "临床,护理,检查,检验,手术,治疗,营养"
Private mbln是否配制中心 As Boolean                      '药房是否具有‘配制中心’性质
Private mblnCard As Boolean                             '是否刷就诊卡

Private mstrNo As String
Private mInt单据 As Integer

Private mdblConditonHeight As Double
Private mintLastTab As Integer
Private mintLastDeptType As Integer
Private mstrSendDrugId As String

Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private mlngMyWindow As Long

'PASS
Private mlngPatiID As Long
Private mlngPassPati As Long
Private mlng主页ID As Long
Private mstr挂号单 As String
'--本程序使用的记录集--
Private RecBillData As New adodb.Recordset              '未发处方记录（已发处方记录）

Private mrsPASS As New adodb.Recordset                  'PASS用数据集

'--内部记录集--
Private RecChangeData As adodb.Recordset                '用于显示各页内容(未发)
Private RecChangeSendedData As adodb.Recordset          '用于显示已发药清单页面的内容
Private RecRefreshCompare As adodb.Recordset            '用于刷新时使用（恢复上次未发药清单中各记录的设定状态）
Private rs序号 As New adodb.Recordset
Private mrsRequest As New adodb.Recordset               '用于显示销帐申请记录
Private mrsRequestMain As New adodb.Recordset

'--查找记录集--
Private strFind As String
Private Rec未发 As adodb.Recordset
Private Rec已发 As adodb.Recordset

'--常量--
Private Const mlng紫色 As Long = &HC000C0
Private Const gInt未发药清单缺药 As Integer = 0
Private Const gInt未发药清单发药 As Integer = 1
Private Const gInt未发药清单拒发 As Integer = 2
Private Const gInt未发药清单不处理 As Integer = 3
Private Const gInt已发药清单退药 As Integer = 3
Private Const gInt已发药清单不处理 As Integer = 1
Private Const gstr排序列名 As String = "|科室|NO|姓名|床号|药品名称|"

'背景色
Private Const glngOtherBlkColor As Long = &H80000005        '一般状态：白色
Private Const glngSendBlkColor As Long = &HFFC0C0           '发药状态：浅蓝色
Private Const glngSelectBlkColor As Long = &HC0C0C0         '当前选择：灰色

Private mstrPrivs As String                              '权限串
Private mlngMode As Long

Private lng汇总清单行数 As Long

Private Enum PatiInfo
    住院号 = 0
    姓名 = 1
    床号 = 2
    单据号 = 3
    病人ID = 4
    就诊卡 = 5
End Enum

Private Type PrivDetail
    Priv_医生查询 As Boolean
End Type

Private UserPrivDetail As PrivDetail

Private Type CellInfo
    Col As Long
    Row As Long
    CellLeft As Single
    CellTop As Single
    CellHeight As Single
    CellWidth As Single
End Type
Private CurCell As CellInfo

'医保接口
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Const mconstRequest = "单据,4,0|NO,7,1200|药品ID,7,0|申请时间,1,2000|收发序号,7,0|产地,1,2000|批号,1,1000|效期,1,1500|准退数量,7,1000|销帐数量,7,1000|包装,1,0|单位,1,1000"

Private Enum 列名_未发药清单
    审查结果 = 0
    分组符 = 1
    科室 = 2
    开单医生 = 3
    状态 = 4
    类型 = 5
    NO = 6
    记帐员 = 7
    床号 = 8
    姓名 = 9
    住院号 = 10
    药品名称 = 11
    其它名 = 12
    英文名 = 13
    规格 = 14
    产地 = 15
    批号 = 16
    付 = 17
    数量 = 18
    单价 = 19
    金额 = 20
    单量 = 21
    频次 = 22
    用法 = 23
    记帐时间 = 24
    说明 = 25
    单据 = 26
    医嘱id = 27
    退药人 = 28
    库房货位 = 29
    相关ID = 30
    药品ID = 31
    单量单位 = 32
    领药部门 = 33
    领药部门id = 34
    
    列数 = 35
End Enum

Private Enum 列名_已发药清单
    审查结果 = 0
    分组符 = 1
    科室 = 2
    状态 = 3
    类型 = 4
    NO = 5
    床号 = 6
    姓名 = 7
    住院号 = 8
    药品名称 = 9
    其它名 = 10
    英文名 = 11
    规格 = 12
    产地 = 13
    批号 = 14
    付 = 15
    数量 = 16
    已退数 = 17
    准退数 = 18
    退药数 = 19
    单价 = 20
    金额 = 21
    单量 = 22
    频次 = 23
    用法 = 24
    操作员 = 25
    发药时间 = 26
    单据 = 27
    医嘱id = 28
    领药人 = 29
    库房货位 = 30
    相关ID = 31
    药品ID = 32
    单量单位 = 33
    
    列数 = 34
End Enum

Private Enum 列名_汇总清单
    药品名称 = 0
    规格 = 1
    产地 = 2
    批号 = 3
    数量 = 4
    单位 = 5
    单价 = 6
    金额 = 7
        
    列数 = 8
End Enum

Private Enum 列名_科室汇总清单
    科室 = 0
    药品名称 = 1
    规格 = 2
    产地 = 3
    批号 = 4
    应发数量 = 5
    留存数量 = 6
    销帐数量 = 7
    实发数量 = 8
    单位 = 9
    单价 = 10
    金额 = 11
    批次 = 12
    科室ID = 13
    药品ID = 14
    领药部门 = 15
    领药部门id = 16
    
    列数 = 17
End Enum

Private Enum 销帐列表
    单据 = 0
    NO = 1
    药品ID = 2
    申请时间 = 3
    收发序号 = 4
    产地 = 5
    批号 = 6
    效期 = 7
    准退数量 = 8
    销帐数量 = 9
    包装 = 10
    单位 = 11
  
    列数 = 12
End Enum

Private Function CheckGroupSend(ByVal lng相关ID As Long) As Boolean
    '检查同组药品是否能够发送
    '前提是药房具有配制中心属性
    '同组药品，只有当所有都是发药状态（其它包括缺药、拒发、不处理）才能发药
    Dim rsGroupRec As adodb.Recordset     '为发药数据集RecChangeData的一个副本
    Dim i As Integer

    '默认是允许发
    CheckGroupSend = True
    
    '不是配制中心则无该规则
    If mbln是否配制中心 = False Then Exit Function
    
    '无分组的不管
    If lng相关ID = 0 Then Exit Function
    
    '创建发药数据集的副本
    Set rsGroupRec = RecChangeData.Clone
    
    '根据传入的NO，相关ID号判断是否该组药品都能发药
    With rsGroupRec
        .Filter = "相关ID=" & lng相关ID
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            '只要存在执行状态不为1，就不能发药
            If !执行状态 <> 1 Then
                CheckGroupSend = False
                Exit Function
            End If
            .MoveNext
        Loop
    End With
End Function

Private Function CheckIsCenter(ByVal lngStockId As Long) As Boolean
    '返回药房是否具有‘配制中心’性质
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 部门性质说明 Where 工作性质 = '配制中心' And 部门id = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否具有配制中心性质", lngStockId)
    
    If Not rsTmp.EOF Then CheckIsCenter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim int入系数 As Integer, int出系数 As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.系数, b.名称 " _
        & " FROM 药品单据性质 A, 药品入出类别 B " _
        & "Where A.类别id = B.ID " _
      & "AND A.单据 = 27  "
    Call SQLTest(App.Title, "药品部门发药", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")

    Call SQLTest
    
    If rsDepend.EOF Then
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiInfo(ByVal intType As Integer, ByVal strInfo As String) As String
    'intType：PatiInfo的项目值
    '返回病人信息：当前病区（ID和部门名称），病人信息（ID和姓名）
    '格式：13,一病区|1,张三
    Dim rsTemp As adodb.Recordset
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim lngH As Long
    Dim blnCancel As Boolean
    
    On Error GoTo errHandle
    If intType = PatiInfo.住院号 Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And B.住院号 = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
    ElseIf intType = PatiInfo.病人ID Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And A.病人id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", Val(strInfo))
    ElseIf intType = PatiInfo.单据号 Then
        gstrSQL = "Select Distinct A.病人病区id As 病区id, B.编码 || '-' || B.名称 As 部门名称, A.病人id, A.姓名 As 病人姓名 " & _
            " From 病人费用记录 A, 部门表 B " & _
            " Where A.病人病区id = B.ID And NO = [1] "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
    ElseIf intType = PatiInfo.床号 Then
        gstrSQL = "Select A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And B.当前床号 = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
    ElseIf intType = PatiInfo.姓名 Then
        If mblnCard = True Then
            gstrSQL = "Select A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And B.就诊卡号 = [1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
        Else
            '病人名称可能会有重复，返回列表供选择
            gstrSQL = "Select Rownum As ID, 病人姓名, 病区id, 部门名称, 病人id" & _
                " From (Select Distinct B.姓名 As 病人姓名, B.病人id, A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And B.姓名 Like [1])"
            
            vRect = GetControlRect(txtPati.hWnd)
            lngH = txtPati.Height
            sngX = vRect.Left - 15
            sngY = vRect.Top
            
            Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "取病人信息", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, "%" & strInfo & "%")
            If blnCancel = True Then Exit Function
        End If
    ElseIf intType = PatiInfo.就诊卡 Then
        gstrSQL = "Select A.当前病区id As 病区id, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.当前病区id = C.ID And B.就诊卡号 = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取病人信息", UCase(strInfo))
    End If
    
    If rsTemp.EOF Then Exit Function
    
    GetPatiInfo = rsTemp!病区id & "," & rsTemp!部门名称 & "|" & rsTemp!病人ID & "," & rsTemp!病人姓名
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPrivs()
    With UserPrivDetail
        .Priv_医生查询 = IsHavePrivs(mstrPrivs, "医生查询")
    End With
End Sub

Private Function GetSumSended(ByVal int单据 As Integer, ByVal strNo As String, ByVal lng药品ID As Long, ByVal int序号 As Integer)
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(Nvl(付数, 1) * 实际数量) 已发数量 From 药品收发记录 Where 单据 = [1] And NO = [2] And 药品ID+0 = [3] And 序号 = [4]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "计算已发数量", int单据, strNo, lng药品ID, int序号)
    
    If Not rsTmp.EOF Then
        GetSumSended = rsTmp!已发数量
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Get发药单格式()
    Dim rsTemp As adodb.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select 说明 From zltools.zlRPTFMTs Where 报表id = (Select ID From zltools.zlReports Where 编号 = 'ZL1_BILL_1342') Order By 序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取摆药单格式")
    
    If rsTemp.RecordCount > 0 Then
        For n = 0 To rsTemp.RecordCount - 1
            cbo发药单格式.AddItem rsTemp!说明
            rsTemp.MoveNext
        Next
          
        cbo发药单格式.ListIndex = 0
        
        If rsTemp.RecordCount = 1 Then
            cbo发药单格式.Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Get剂型()
    Dim bln中药库房 As Boolean
    Dim rsTmp As adodb.Recordset
    
    On Error GoTo errHandle
    '提取所有剂型
    bln中药库房 = False
    gstrSQL = "Select 1 From 部门性质说明 " & _
         " Where 工作性质 Like '中药%' And 部门ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", lng药房ID)
    
    If Not rsTmp.EOF Then bln中药库房 = True
    
    gstrSQL = "Select Distinct J.编码||'-'||J.名称 剂型" & _
         " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
         " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
         " And A.执行科室ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng药房ID)
    
    With rsTmp
        Lvw剂型.ListItems.Clear
        Lvw剂型.ListItems.Add , "_" & Lvw剂型.ListItems.Count + 1, "所有药品剂型", 1, 1
        Lvw剂型.ListItems(Lvw剂型.ListItems.Count).Checked = True
        Do While Not .EOF
            Lvw剂型.ListItems.Add , "_" & Lvw剂型.ListItems.Count + 1, !剂型, 1, 1
            Lvw剂型.ListItems(Lvw剂型.ListItems.Count).Checked = True
            .MoveNext
        Loop
        If bln中药库房 Then
           Lvw剂型.ListItems.Add , "_" & Lvw剂型.ListItems.Count + 1, "0-方剂", 1, 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get销帐数量(ByVal lng部门id As Long, ByVal lng药品ID As Long) As Double
    Dim dblSum As Double
    
    With mrsRequest
        .Filter = "领药部门id=" & lng部门id & " And 药品ID=" & lng药品ID & " And 审核标志 = 1"
        If .EOF Then Exit Function
        
        Do While Not .EOF
            dblSum = dblSum + !销帐数量 / !包装
            .MoveNext
        Loop
    End With
    
    Get销帐数量 = dblSum
End Function

Private Sub IniConditon()
    Dim dateCurDate As Date
    Dim rsTmp As New adodb.Recordset
    Dim n As Integer
    Const cst每个字节宽度 = 128
    
    On Error GoTo errHandle
    dateCurDate = zldatabase.Currentdate()
    Me.Dtp开始时间.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
    Me.Dtp结束时间.Value = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\已发药清单条件", "时间范围", Me.Dtp开始时间.Value & ";" & Me.Dtp结束时间.Value
    
    Me.Dtp销帐开始时间.Value = Me.Dtp开始时间.Value
    Me.Dtp销帐结束时间.Value = Me.Dtp结束时间.Value
    
    '默认的领药部门类型是病区
    mintLastDeptType = 2
    tbsType.Tabs(3).Selected = True
    
    '提取剂型
    Call Get剂型
    
    '提取发药类型，并动态增加发药类型选择框
    gstrSQL = "Select 名称 From 发药类型 Order By 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取发药类型]")
    
    If rsTmp.RecordCount > 0 Then
        chkSendType(0).Visible = True
        chkSendType(0).Caption = rsTmp!名称
        chkSendType(0).Width = 150 + LenB(chkSendType(0).Caption) * 128
        If rsTmp.RecordCount > 1 Then
            rsTmp.MoveNext
            For n = 2 To rsTmp.RecordCount
                Load chkSendType(n - 1)
                chkSendType(n - 1).Visible = True
                chkSendType(n - 1).Caption = rsTmp!名称
                chkSendType(n - 1).Width = 150 + LenB(chkSendType(n - 1).Caption) * 128
                rsTmp.MoveNext
            Next
        End If
        
        Call ResizeCheckControl
    End If
    
    '设置医嘱类型
    With Cbo医嘱类型
        .Clear
        .AddItem "0-包含所有单据"
        .AddItem "1-仅含长期医嘱"
        .AddItem "2-仅含临时医嘱"
        .AddItem "3-普通记帐单据"
        .AddItem "4-包含所有医嘱"
        .ListIndex = Lng医嘱类型
    End With
    
    '提取所有给药途径
    gstrSQL = "Select 名称 as 用法 ,标本部位 As 分类 From 诊疗项目目录 Where 类别='E' And 操作类型='2'And (服务对象=2 Or 服务对象=3) " & _
            " And (撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or 撤档时间 Is Null) Order by 编码 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    With rsTmp
        Lvw给药途径.ListItems.Add , "_" & Lvw给药途径.ListItems.Count + 1, "所有给药途径", 1, 1
        Lvw给药途径.ListItems(Lvw给药途径.ListItems.Count).Checked = True
        Do While Not .EOF
            Lvw给药途径.ListItems.Add , "_" & Lvw给药途径.ListItems.Count + 1, !用法, 1, 1
            Lvw给药途径.ListItems(Lvw给药途径.ListItems.Count).Checked = True
            Lvw给药途径.ListItems(Lvw给药途径.ListItems.Count).Tag = !分类
            .MoveNext
        Loop
    End With
    
    '设置给药途径分类
    gstrSQL = "Select Distinct 标本部位 As 分类 From 诊疗项目目录 Where 类别 = 'E' And 操作类型 = '2' And 标本部位 Is Not Null"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "取给药途径分类")
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    mnuTypeItem.Item(0).Caption = rsTmp!分类
    
    If rsTmp.RecordCount > 1 Then
        rsTmp.MoveNext
        For n = 2 To rsTmp.RecordCount
            Load mnuTypeItem.Item(n - 1)
            mnuTypeItem.Item(n - 1).Caption = rsTmp!分类
            mnuTypeItem.Item(n - 1).Visible = True
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectDept(ByVal intType As Integer, ByVal strInput As String) As adodb.Recordset
    Dim strSQL As String
    Dim dblX As Double
    Dim dblY As Double
    Dim DblHeight As Double
    
    dblX = fraCondition.Left + fraConNormal.Left + txt科室.Left
    dblY = fraCondition.Top + fraConNormal.Top + txt科室.Top + txt科室.Height
    DblHeight = 5000
    
    If intType = 0 Then
        strSQL = " Select ID, 编码||'-'||名称 部门 From 部门表 " & _
                 " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质='临床' And 服务对象 IN(2,3))" & _
                 " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
    ElseIf intType = 1 Then
        strSQL = " Select ID, 编码||'-'||名称 部门 From 部门表 " & _
                 " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术') And 服务对象 IN(2,3))" & _
                 " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
    Else
        strSQL = " Select ID, 编码||'-'||名称 部门 From 部门表 " & _
                 " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
                 " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) "
    End If
    
    strSQL = strSQL & " And (Upper(编码) Like '" & UCase(strInput) & "%'" & _
            " Or Upper(名称) Like '" & StrFindStyle & UCase(strInput) & "%'" & _
            " Or Upper(简码) Like '" & StrFindStyle & UCase(strInput) & "%')"
            
    strSQL = strSQL & " Order By 编码||'-'||名称"
    
    Set SelectDept = zldatabase.ShowSelect(Me, strSQL, 0, "部门列表", , , , , True, , dblX, dblY, DblHeight)
End Function
Private Sub Get配药人()
    Dim strSQL As String
    Dim rsTemp As New adodb.Recordset
    
    On Error GoTo errHandle
    '设置记帐人
    gstrSQL = "Select Distinct A.简码||'-'||A.姓名 As 姓名" & _
             " From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
             " Where (A.站点 = '" & gstrNodeNo & "' Or A.站点 is Null) And A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id And D.人员性质 = '药房发药人' " & _
             " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND B.部门id=[1] " & _
             " ORDER BY 姓名 "

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取药房人员", lng药房ID)
    
    cbo配药人.Clear
    Do While Not rsTemp.EOF
        cbo配药人.AddItem rsTemp!姓名
        rsTemp.MoveNext
    Loop
    
    cbo配药人.Text = gstrUserAbbr & "-" & gstrUserName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LoadCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strTemp As String
    Dim strBegin As String
    Dim strEnd As String
    Dim dateCurDate As Date
    Dim n As Integer
    Dim i As Integer
    Dim arrStr
    
    If BlnFirstStart = False And gblnMyStyle = False Then
        dateCurDate = zldatabase.Currentdate()
        Me.Dtp开始时间.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
        Me.Dtp结束时间.Value = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
        
        Me.Dtp销帐开始时间.Value = Me.Dtp开始时间.Value
        Me.Dtp销帐结束时间.Value = Me.Dtp结束时间.Value
        
        Cbo医嘱类型.ListIndex = Lng医嘱类型
        
        '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
        If int离院带药 = 1 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf int离院带药 = 2 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf int离院带药 = 3 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf int离院带药 = 4 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf int离院带药 = 5 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 0
        ElseIf int离院带药 = 6 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        Else
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        End If
        Exit Sub
    End If
    
    If intType = 0 Then
        strPath = "未发药清单条件"
    Else
        strPath = "已发药清单条件"
    End If
    
    '条件
    mblnAllConditon = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "全部条件", "0")) = 1)
    If mblnAllConditon = True Then
        cmdOtherCon.Caption = "简要条件(&C)"
    Else
        cmdOtherCon.Caption = "全部条件(&C)"
    End If
    Call ResizeCondition
    
    '时间范围
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "时间范围", "")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        dateCurDate = zldatabase.Currentdate()
        strTemp = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00") & ";" & Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
    Else
        strBegin = Split(strTemp, ";")(0)
        strEnd = Split(strTemp, ";")(1)
        
        If Not IsDate(strBegin) Then
            dateCurDate = zldatabase.Currentdate()
            strBegin = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
        End If
        
        If Not IsDate(strEnd) Then
            dateCurDate = zldatabase.Currentdate()
            strEnd = Format(dateCurDate, "yyyy-MM-dd") & " 23:59:59"
        End If
        
        strTemp = strBegin & ";" & strEnd
    End If
    
    Dtp开始时间.Value = Split(strTemp, ";")(0)
    Dtp结束时间.Value = Split(strTemp, ";")(1)
        
    '病人信息
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "病人信息", "0;")
    If Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 4 Then
        Call mnuInfoItem_Click(0)
    Else
        Call mnuInfoItem_Click(Val(Split(strTemp, ";")(0)))
    End If
    txtPati.Text = Split(strTemp, ";")(1)
    
    '科室
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "科室", "")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        tbsType.Tabs(3).Selected = True
        txt科室.Text = ""
    Else
        If Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 2 Then
            tbsType.Tabs(3).Selected = True
        Else
            tbsType.Tabs(Val(Split(strTemp, ";")(0)) + 1).Selected = True
        End If
        txt科室.Tag = Split(strTemp, ";")(1)
        txt科室.Text = Split(strTemp, ";")(2)
    End If
    
    '发药类型
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "发药类型", "0;")
    If strTemp = "" Or InStr(strTemp, ";") = 0 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
        
        If chkSendType(0).Visible = True Then
            For n = 0 To chkSendType.UBound
                chkSendType(n).Value = 0
            Next
        End If
    ElseIf Val(Split(strTemp, ";")(0)) < 0 Or Val(Split(strTemp, ";")(0)) > 5 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    Else
        '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药（不包括离院带药和自取药）
        If Val(Split(strTemp, ";")(0)) = 1 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf Val(Split(strTemp, ";")(0)) = 2 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 3 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 4 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 0
            chkSend(2).Value = 1
        ElseIf Val(Split(strTemp, ";")(0)) = 5 Then
            chkSend(0).Value = 1
            chkSend(1).Value = 0
            chkSend(2).Value = 0
        ElseIf Val(Split(strTemp, ";")(0)) = 6 Then
            chkSend(0).Value = 0
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        Else
            chkSend(0).Value = 1
            chkSend(1).Value = 1
            chkSend(2).Value = 1
        End If
        
        If Split(strTemp, ";")(1) <> "" Then
            arrStr = Split(Split(strTemp, ";")(1), ",")
            For n = 0 To UBound(arrStr)
                For i = 0 To chkSendType.UBound
                    If arrStr(n) = chkSendType(i).Caption Then
                        chkSendType(i).Value = 1
                    End If
                Next
            Next
        End If
    End If
    
    '给药途径
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "给药途径", "所有给药途径")
    txt给药途径.Text = strTemp
    
    '药品剂型
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "药品剂型", "所有药品剂型")
    txt药品剂型.Text = strTemp
    
    '处理范围
    strTemp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "处理范围", "0"))
    If Val(strTemp) < 0 Or Val(strTemp) > 2 Then
        strTemp = "0"
    End If
    opt范围(Val(strTemp)).Value = True
    
    '医嘱类型
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "医嘱类型", "0")
    If Val(strTemp) < 0 Or Val(strTemp) > 4 Then
        strTemp = "0"
    End If
    Cbo医嘱类型.ListIndex = Val(strTemp)
    
    
End Sub


Private Function Get销帐清单() As Boolean
    Dim strSubUnit As String
    Dim rsTemp As adodb.Recordset
    Dim strCon As String
    Dim strTmpCon As String
    Dim str申请时间 As String
    Dim lng领药部门ID As Long
    Dim lng药品ID As Long
    Dim lng费用id As Long
    Dim dbl准退数量 As Double

    '单位，包装换算
    On Error GoTo errHandle
    Select Case strUnit
    Case "售价单位"
        strSubUnit = "X.计算单位 单位,1 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "门诊单位"
        strSubUnit = "D.门诊单位 单位,D.门诊包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "住院单位"
        strSubUnit = "D.住院单位 单位,D.住院包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "药库单位"
        strSubUnit = "D.药库单位 单位,D.药库包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    End Select
    
    If mint类型 = 0 Then
        strCon = " And H.Id = B.病人科室id "
    ElseIf mint类型 = 1 Then
        strCon = " And H.Id = B.开单部门id "
    Else
        strCon = " And H.Id = B.病人病区ID "
    End If

    If mstrSerchNO <> "" Then
    ElseIf mstr住院号 <> "" Then
        strCon = strCon & " And B.标识号=[4] "
    ElseIf mstr病人姓名 <> "" Then
        strCon = strCon & " And B.姓名 Like [5] "
    ElseIf mlng病人ID <> 0 Then
        strCon = strCon & " And B.病人ID=[6] "
    ElseIf mstr床号 <> "" Then
        strCon = strCon & " And B.床号 = [7] "
    End If
    
    gstrSQL = "Select /*+rule*/ Distinct '['||X.编码||']'||" & IIf(mblnTradeName, "NVL(K.名称,X.名称)", "X.名称") & " As 药品名称, " & _
        " C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室,H.名称 As 领药部门,H.Id As 领药部门Id, " & _
        " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, " & strSubUnit & " " & _
        " From 病人费用销帐 A, 病人费用记录 B," & _
        " (Select A.ID, A.单据, A.NO, A.序号, A.药品id, A.产地, A.批号, A.效期, A.费用id, B.实际数量 " & _
            " From 药品收发记录 A, " & _
            " (Select C.单据, C.NO, C.序号, C.药品id, Sum(Nvl(C.付数, 1) * C.实际数量) As 实际数量 " & _
            " From 药品收发记录 C, 病人费用销帐 A, 病人费用记录 B " & _
            " Where A.费用id = B.ID And B.NO = C.NO And B.ID = C.费用id And A.状态 = 0 " & _
            " And C.单据 In (9, 10) And C.审核日期 Is Not Null And C.库房id = [1] And Instr([3], ',' || A.收费细目id || ',') > 0 " & strTmpCon & _
            " Group By C.单据, C.NO, C.序号, C.药品id " & _
            " Having Sum(Nvl(C.付数, 1) * C.实际数量) > 0) B" & _
            " Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号 And A.审核人 Is Not Null " & _
            " And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0))C, " & _
        " 药品规格 D, 收费项目目录 X, 收费项目别名 K, 部门表 P, 病案主页 F, 部门表 E,部门表 H " & _
        " Where A.费用id = B.ID And B.NO = C.NO And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID And B.病人id = F.病人id And B.主页id = F.主页id  And F.出院日期 Is Null And A.申请部门id = E.ID " & strCon & _
        " And X.Id = K.收费细目ID(+) AND K.性质(+)=3  And B.执行部门id = [1] And Instr([2], ',' || A.申请部门id || ',') > 0 And A.审核人 Is Null And A.状态 = 0 " & _
        " Order By A.申请时间, C.单据, C.NO, C.序号 Desc "
    
    'And Instr([3], ',' || A.收费细目id || ',') > 0 " & _
    '" And A.申请时间 Between [3] And [4] "
    'Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取批次明细", lng药房ID, "," & mstrDrawDept & ",", Dtp销帐开始时间.Value, Dtp销帐结束时间.Value)
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取批次明细", lng药房ID, "," & mstrDrawDept & ",", "," & mstrSendDrugId & ",", mstr住院号, mstr病人姓名, mlng病人ID, mstr床号)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    Do While Not rsTemp.EOF
        With mrsRequest
            .AddNew
            !药品名称 = rsTemp!药品名称
            !领药部门 = rsTemp!领药部门
            !领药部门id = rsTemp!领药部门id
            !单据 = rsTemp!单据
            !NO = rsTemp!NO
            !药品ID = rsTemp!药品ID
            !申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            !收发序号 = rsTemp!收发序号
            !产地 = rsTemp!产地
            !批号 = rsTemp!批号
            !效期 = rsTemp!效期
            
            If gtype_UserSysParms.P149_效期显示方式 = 1 And NVL(!效期) <> "" Then
                '换算为有效期
                !效期 = Format(DateAdd("D", -1, !效期), "yyyy-mm-dd")
            End If
            
            !准退数量 = rsTemp!准退数量
            !销帐数量 = rsTemp!销帐数量
            !包装 = rsTemp!包装
            !单位 = rsTemp!单位
            !收发ID = rsTemp!收发ID
            !主页id = IIf(IsNull(rsTemp!主页id), 0, rsTemp!主页id)
            !费用序号 = rsTemp!费用序号
            !险类 = rsTemp!险类
            !费用id = rsTemp!费用id
            !记录性质 = rsTemp!记录性质
            !审核标志 = 0
            .Update
        End With
        
        With mrsRequestMain
            dbl准退数量 = dbl准退数量 + rsTemp!准退数量
            If lng领药部门ID <> rsTemp!领药部门id And str申请时间 <> Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss") And lng费用id <> rsTemp!费用id Then
                .AddNew
                !领药部门id = rsTemp!领药部门id
                !药品ID = rsTemp!药品ID
                !申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                !费用id = rsTemp!费用id
                !准退数量 = dbl准退数量
                !销帐数量 = rsTemp!销帐数量
                
                .Update
                
                dbl准退数量 = 0
            End If
            lng领药部门ID = rsTemp!领药部门id
            str申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            lng药品ID = rsTemp!药品ID
            lng费用id = rsTemp!费用id
        End With
        
        rsTemp.MoveNext
    Loop
    
    '只处理发药清单对应的药品（按领药部门ID，药品ID为准）
    mrsRequest.MoveFirst
    Do While Not mrsRequest.EOF
        RecBillData.MoveFirst
        Do While Not RecBillData.EOF
            If mrsRequest!领药部门id = RecBillData!领药部门id And mrsRequest!药品ID = RecBillData!药品ID Then
                mrsRequest!审核标志 = 1
                mrsRequest.Update
            End If
            RecBillData.MoveNext
        Loop
        mrsRequest.MoveNext
    Loop
    
    RecBillData.MoveFirst
    mrsRequest.MoveFirst
    
    Call AutoExpendQuantity
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function LoadDataInBill销帐清单(ByVal lng领药部门ID As Integer, ByVal lng药品ID As Long) As Boolean
    Dim dblSumNum As Double
    
    With mrsRequest
        Call ClearBill(Bill退药销帐)
        
        .Filter = "领药部门id=" & lng领药部门ID & " And 药品ID=" & lng药品ID & " And 审核标志 = 1"
        .Sort = "NO,收发序号 Desc"
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.单据) = !单据
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.NO) = !NO
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.药品ID) = !药品ID
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.申请时间) = Format(!申请时间, "yyyy-mm-dd hh:mm:ss")
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.产地) = IIf(IsNull(!产地), "", !产地)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.批号) = IIf(IsNull(!批号), "", !批号)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.效期) = Format(!效期, "yyyy-mm-dd")
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.准退数量) = FormatEx(!准退数量 / !包装, 5)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.销帐数量) = FormatEx(!销帐数量 / !包装, 5)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.包装) = IIf(IsNull(!包装), "", !包装)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.单位) = IIf(IsNull(!单位), "", !单位)
            Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.收发序号) = IIf(IsNull(!收发序号), "", !收发序号)
            Bill退药销帐.rows = Bill退药销帐.rows + 1
            
            dblSumNum = dblSumNum + !销帐数量 / !包装
            
           .MoveNext
        Loop
        
        Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.NO) = "合计"
        Bill退药销帐.TextMatrix(Bill退药销帐.rows - 1, 销帐列表.销帐数量) = FormatEx(dblSumNum, 5)
        
        Bill退药销帐.Row = Bill退药销帐.rows - 1
        Bill退药销帐.Col = 销帐列表.NO
        Bill退药销帐.CellForeColor = glng发药
        
        Bill退药销帐.Col = 销帐列表.销帐数量
        Bill退药销帐.CellForeColor = glng发药
    End With
    
    LoadDataInBill销帐清单 = True
End Function
Private Sub AutoExpendQuantity()
    '考虑到同一费用ID对应多个收发ID的情况，需要将销帐数量分解到多个收发记录上
    '分解的原则是按序号大的优先分配（已按序号降序排序）
    Dim n As Integer
    Dim dbl准退数量 As Double
    Dim dbl剩余数量 As Double
    Dim int收发序号 As Integer
    Dim lng费用id As Long
    Dim lng药品ID As Long
    Dim str申请时间 As String
    
    With mrsRequest
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl准退数量 = !准退数量

            If lng费用id = !费用id And lng药品ID = !药品ID And str申请时间 = !申请时间 Then

            Else
                dbl剩余数量 = !销帐数量
            End If

            If dbl剩余数量 >= dbl准退数量 Then
                dbl剩余数量 = dbl剩余数量 - dbl准退数量
                !销帐数量 = dbl准退数量
            Else
                !销帐数量 = dbl剩余数量
                dbl剩余数量 = 0
            End If

            lng费用id = !费用id
            lng药品ID = !药品ID
            str申请时间 = !申请时间

            .Update
            .MoveNext
        Next
    End With
    
    '销帐数量大于了准退数量，则标志为拒绝审核
    With mrsRequestMain
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mrsRequest.Filter = "药品ID=" & !药品ID & _
                " And 费用ID=" & !费用id & _
                " And 申请时间='" & !申请时间 & "'"
            If mrsRequest.RecordCount > 0 Then
                If !准退数量 < !销帐数量 Then
                    Do While Not mrsRequest.EOF
                        mrsRequest!审核标志 = 2
                        mrsRequest.Update
                        mrsRequest.MoveNext
                    Loop
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ResizeCheckControl()
    '调整发药类型选择框位置
    Dim n As Integer
    Dim dbl最大宽度 As Double
    Dim dblTmp As Double
    Dim dblSumTmp As Double
    Dim int行数 As Integer
    Const cst间隔宽度 = 50
    Const cst行距 = 50
    
    If chkSendType.UBound > 0 Then
        dbl最大宽度 = fraConNormal.Width - fraTypeLine.Left - 150
        
        int行数 = 0
        dblSumTmp = chkSendType(0).Width + cst间隔宽度
        For n = 1 To chkSendType.UBound
            dblTmp = chkSendType(n).Width + dblSumTmp
            
            If dblTmp <= dbl最大宽度 Then
                chkSendType(n).Top = chkSendType(n - 1).Top
                chkSendType(n).Left = chkSendType(n - 1).Left + chkSendType(n - 1).Width + cst间隔宽度
                dblSumTmp = dblSumTmp + chkSendType(n).Width + cst间隔宽度
            Else
                '换新行，并调整其他控件位置
                int行数 = int行数 + 1
                chkSendType(n).Left = chkSendType(0).Left
                chkSendType(n).Top = chkSendType(0).Top + (chkSendType(0).Height + cst行距) * int行数
                dblSumTmp = chkSendType(n).Width + cst间隔宽度
                
                fraTypeLine.Height = fraTypeLine.Height + chkSendType(0).Height * int行数
                
                fraConNormal.Height = fraConNormal.Height + chkSendType(0).Height + cst行距
                fraCondition.Height = fraCondition.Height + chkSendType(0).Height + cst行距
                
                frmLine.Top = frmLine.Top + chkSendType(0).Height + cst行距
                
                fraConExpand.Top = fraConExpand.Top + chkSendType(0).Height + cst行距
                fraConRequest.Top = fraConRequest.Top + chkSendType(0).Height + cst行距
                
                If mblnAllConditon = True Then
                    cmdRefresh.Top = cmdRefresh.Top + chkSendType(0).Height + cst行距
                    cmdOtherCon.Top = cmdRefresh.Top
                End If

                TabShow.Height = TabShow.Height - (chkSendType(0).Height + cst行距)
            End If
        Next
    End If
End Sub


Private Sub ResizeCondition()
    Dim dblDistance As Double
    Dim n As Integer
    Dim DblHeight As Double, DblWidth As Double
    
    fraCondition.Top = IIf(Cbar.Visible, Cbar.Height, 0)
    fraCondition.Width = Me.ScaleWidth - 20
    fraCondition.Visible = True
    
    fraConExpand.Visible = False
    fraConRequest.Visible = False
    
    frmLine.Visible = False
    frmLine1.Visible = False
    
    fraConNormal.Top = 100
    fraConNormal.Left = 20
    fraConExpand.Left = 20
    fraConRequest.Left = 20
    
    frmLine.Width = fraCondition.Width
    frmLine1.Width = fraCondition.Width
    frmLine.ZOrder 0
    frmLine1.ZOrder 0
    
    cmdRefresh.Top = fraConNormal.Top + (fraConNormal.Height - cmdRefresh.Height) - 50
    cmdOtherCon.Top = cmdRefresh.Top
    cmdRefresh.Left = fraConNormal.Left + fraConNormal.Width + 10
    cmdOtherCon.Left = cmdRefresh.Left + cmdRefresh.Width + 10
    
    If mblnAllConditon = True Then
        fraConExpand.Visible = True
        frmLine.Visible = True

        frmLine.Top = fraConNormal.Top + fraConNormal.Height + 20

        fraConExpand.Top = frmLine.Top - 120
    Else
        fraConExpand.Visible = False
        frmLine.Visible = False
    End If

    If fraConExpand.Visible = True Then
        frmLine1.Top = fraConExpand.Top + fraConExpand.Height
    Else
        frmLine1.Top = fraConNormal.Top + fraConNormal.Height
    End If
    fraCondition.Height = frmLine1.Top + 50
    
    With TabShow
        .Top = IIf(Cbar.Visible, Cbar.Height + fraCondition.Height, fraCondition.Height)
        .Left = 0
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = Me.ScaleWidth
    End With
    
    DblHeight = TabShow.Height - TabShow.TabHeight - 120
    DblWidth = TabShow.Width - 150
    With Bill未发药清单
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill汇总发药
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill拒发药清单
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill缺药清单
        .Height = DblHeight
        .Width = DblWidth
    End With
    With Bill已发药清单
        .Height = DblHeight
        .Width = DblWidth
    End With
    
    If Bill退药销帐.Visible = True Then
        Bill汇总发药.Height = DblHeight - Bill退药销帐.Height - 25
        Bill退药销帐.Top = Bill汇总发药.Top + Bill汇总发药.Height + 25
    End If
    
    '调整配药人和发药单打印格式
    If TabShow.Tab = 0 Then
        lbl配药人.Top = TabShow.Height - lbl配药人.Height - 120
        cbo配药人.Top = lbl配药人.Top - 60
        Bill未发药清单.Height = TabShow.Height - TabShow.TabHeight - 120 - lbl配药人.Height - 150
        
        lbl发药单格式.Top = lbl配药人.Top
        cbo发药单格式.Top = cbo配药人.Top
        cbo发药单格式.Left = TabShow.Width - cbo发药单格式.Width - 50
        lbl发药单格式.Left = cbo发药单格式.Left - 50 - lbl发药单格式.Width
    End If
    
    '其他控件调整
    Chk清单.Top = TabShow.Top + 70
    Chk显示退药待发单据.Top = Chk清单.Top
    cmdAlley.Top = TabShow.Top + 30
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function
Private Sub SaveCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strBegin As String
    Dim strEnd As String
    
    '保存查找条件到注册表
    
    If intType = 0 Then
        strPath = "未发药清单条件"
        strBegin = mstr开始日期_未发
        strEnd = mstr结束日期_未发
    Else
        strPath = "已发药清单条件"
        strBegin = mstr开始日期_已发
        strEnd = mstr结束日期_已发
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "全部条件", IIf(mblnAllConditon = True, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "时间范围", strBegin & ";" & strEnd
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "病人信息", Val(lblPatiInputType.Tag) & ";" & Trim(txtPati.Text)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "科室", mint类型 & ";" & mstr部门 & ";" & mstr部门名称
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "发药类型", int离院带药 & ";" & mstr发药类型
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "给药途径", mstrUse
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "药品剂型", mstrDrug
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "处理范围", mint范围
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "医嘱类型", Lng医嘱类型
End Sub


Private Sub ClearCondition(ByVal intType As Integer)
    Dim strPath As String
    Dim strBegin As String
    Dim strEnd As String
    
    '清除条件：根据需要，如果要保存部分条件，在删除后添加
    On Error Resume Next
    
    DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\未发药清单条件"
    DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\已发药清单条件"
    
    '保存查找条件到注册表
    If intType = 0 Then
        strPath = "未发药清单条件"
        strBegin = mstr开始日期_未发
        strEnd = mstr结束日期_未发
    Else
        strPath = "已发药清单条件"
        strBegin = mstr开始日期_已发
        strEnd = mstr结束日期_已发
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\" & strPath, "发药类型", int离院带药 & ";" & mstr发药类型
End Sub
Private Sub SetCondition(ByVal intType As Integer)
    Dim n As Integer
    
    '时间范围
    If intType = 1 Then
        mstr开始日期_已发 = Format(Dtp开始时间.Value, "yyyy-MM-dd hh:mm:ss")
        mstr结束日期_已发 = Format(Dtp结束时间.Value, "yyyy-MM-dd hh:mm:ss")
    Else
        mstr开始日期_未发 = Format(Dtp开始时间.Value, "yyyy-MM-dd hh:mm:ss")
        mstr结束日期_未发 = Format(Dtp结束时间.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    '病人信息
    mstr住院号 = ""
    mstr病人姓名 = ""
    mstr床号 = ""
    mstrSerchNO = ""
    mlng病人ID = 0
        
    If Trim(txtPati.Text) <> "" Then
        Select Case Val(lblPatiInputType.Tag)
            Case PatiInfo.住院号
                If InStr(txtPati.Text, "-") > 0 Then
                    mstr住院号 = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstr住院号 = Trim(txtPati.Text)
                End If
            Case PatiInfo.姓名
                If mblnCard = True Then
                    mlng病人ID = Val(txtPati.Tag)
                Else
                    mstr病人姓名 = Trim(txtPati.Text)
                End If
            Case PatiInfo.床号
                If InStr(txtPati.Text, "-") > 0 Then
                    mstr床号 = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstr床号 = Trim(txtPati.Text)
                End If
            Case PatiInfo.单据号
                If InStr(txtPati.Text, "-") > 0 Then
                    mstrSerchNO = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mstrSerchNO = Trim(txtPati.Text)
                End If
            Case PatiInfo.病人ID
                If InStr(txtPati.Text, "-") > 0 Then
                    mlng病人ID = Mid(Trim(txtPati.Text), 1, InStr(txtPati.Text, "-") - 1)
                Else
                    mlng病人ID = Val(Trim(txtPati.Text))
                End If
            Case PatiInfo.就诊卡
                mlng病人ID = Val(txtPati.Tag)
        End Select
    End If
    
    '部门类型和部门ID
    mstr部门 = ""
    mstr部门名称 = ""
    mint类型 = (tbsType.SelectedItem.Index - 1)
    If Trim(txt科室.Text) <> "" Then
        mstr部门 = txt科室.Tag
        mstr部门名称 = txt科室.Text
    End If
        
    '剂型
    If Trim(txt药品剂型.Text) = "" Or InStr(Trim(txt药品剂型.Text), "所有药品剂型") > 0 Then
        mstrDrug = ""
    Else
        mstrDrug = Trim(txt药品剂型.Text)
    End If
    
    '给药途径
    If Trim(txt给药途径.Text) = "" Or InStr(Trim(txt给药途径.Text), "所有给药途径") > 0 Then
        mstrUse = ""
    Else
        mstrUse = Trim(txt给药途径.Text)
    End If
    
    '处理范围
    If Me.opt范围(1).Value = True Then
        mint范围 = 1
    ElseIf Me.opt范围(2).Value = True Then
        mint范围 = 2
    Else
        mint范围 = 0
    End If
    
    '发药类型
    '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        int离院带药 = 0
    ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
        int离院带药 = 1
    ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        int离院带药 = 3
    ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        int离院带药 = 6
    ElseIf chkSend(0).Value = 1 Then
        int离院带药 = 5
    ElseIf chkSend(1).Value = 1 Then
        int离院带药 = 2
    ElseIf chkSend(2).Value = 1 Then
        int离院带药 = 4
    End If
    
    mstr发药类型 = ""
    If chkSendType(0).Visible = True Then
        For n = 0 To chkSendType.UBound
            If chkSendType(n).Value = 1 Then
                mstr发药类型 = IIf(mstr发药类型 = "", "", mstr发药类型 & ",") & chkSendType(n).Caption
            End If
        Next
    End If
    
    '医嘱类型
    Lng医嘱类型 = Cbo医嘱类型.ListIndex
    
    '病人类型
    If chkType(0).Value = 1 And chkType(1).Value = 1 Then
        mint病人类型 = 2
    ElseIf chkType(1).Value = 1 Then
        mint病人类型 = 1
    Else
        mint病人类型 = 0
    End If
    
End Sub

Private Sub SetGroup(ByVal Bill As MSHFlexGrid, ByVal bln是否分组 As Boolean)
    Dim n As Integer
    Dim lng上行相关ID As Long
    Dim lng本行相关ID As Long
    Dim lng下行相关ID As Long
    Dim int列名_相关ID As Integer
    Dim int列名_分组符 As Integer
    Dim bln是否存在分组 As Boolean
    Dim bln汇总行分组符 As Boolean
    
    '制表符：└ ┌ │
    
    '总行数小于三行时没有必要分组
    If Bill.rows < 3 Then Exit Sub
    
    lng上行相关ID = -1
        
    '按相关ID分组
    With Bill
        Select Case .Name
        Case "Bill未发药清单"
            int列名_相关ID = 列名_未发药清单.相关ID
            int列名_分组符 = 列名_未发药清单.分组符
        Case "Bill已发药清单"
            int列名_相关ID = 列名_已发药清单.相关ID
            int列名_分组符 = 列名_已发药清单.分组符
        End Select
        
        .Redraw = False
        For n = 1 To .rows - 1
             .TextMatrix(n, int列名_分组符) = ""
             .RowHeight(n) = 220
        Next
                
        If Not bln是否分组 Then
            .ColWidth(int列名_分组符) = 0
            .Redraw = True
            Exit Sub
        Else
            .ColWidth(int列名_分组符) = 250
        End If
        
        For n = 1 To .rows - 1
            .Row = n
            .Col = int列名_分组符
            If .TextMatrix(n, int列名_相关ID) <> "" Then
                lng本行相关ID = .TextMatrix(n, int列名_相关ID)
                If n + 1 <= .rows - 1 Then
                    If .TextMatrix(n + 1, int列名_相关ID) <> "" Then    '如果下行为记录行时
                        lng下行相关ID = IIf(.TextMatrix(n + 1, int列名_相关ID) = 0, -1, .TextMatrix(n + 1, int列名_相关ID))
                    ElseIf n + 2 <= .rows - 1 Then  '如果下行为汇总行行时
                        If .TextMatrix(n + 2, int列名_相关ID) <> "" Then    '如果下下行为记录行时
                            lng下行相关ID = IIf(.TextMatrix(n + 2, int列名_相关ID) = 0, -1, .TextMatrix(n + 2, int列名_相关ID))
                        Else
                            lng下行相关ID = -1
                        End If
                    Else
                        lng下行相关ID = -1
                    End If
                Else
                    lng下行相关ID = -1
                End If
                
                If lng本行相关ID = lng上行相关ID Then
                    If lng本行相关ID = lng下行相关ID Then
                        .TextMatrix(n, int列名_分组符) = "│"
                        .RowHeight(n) = 220
                    Else
                        .TextMatrix(n, int列名_分组符) = "└"
                        .CellAlignment = flexAlignLeftTop
                    End If
                ElseIf lng本行相关ID = lng下行相关ID Then
                    .TextMatrix(n, int列名_分组符) = "┌"
                    .CellAlignment = flexAlignLeftBottom
                    bln是否存在分组 = True
                End If
            
                lng上行相关ID = IIf(lng本行相关ID = 0, -1, lng本行相关ID)
            Else
                '如果该行是汇总行，则要根据下行的相关ID判断分组符号
                If n + 1 <= .rows - 1 Then
                    If .TextMatrix(n + 1, int列名_相关ID) <> "" Then
                        If lng上行相关ID <> -1 And lng上行相关ID = IIf(.TextMatrix(n + 1, int列名_相关ID) = 0, -1, .TextMatrix(n + 1, int列名_相关ID)) Then
                            .TextMatrix(n, int列名_分组符) = "│"
                            .RowHeight(n) = 220
                        End If
                    End If
                End If
            End If
        Next
        
        If Not bln是否存在分组 Then .ColWidth(int列名_分组符) = 0
        
        .Redraw = True

    End With
    
End Sub

Private Sub Bill退药销帐_EnterCell()
    Call SetSelectColor(Bill退药销帐)
End Sub


Private Sub Bill退药销帐_GotFocus()
    Call Bill退药销帐_EnterCell
End Sub
Private Sub Bill未发药清单_Scroll()
    Cbo批号.Visible = False
End Sub

Private Sub Bill已发药清单_Scroll()
    TxtInput.Visible = False
End Sub

Private Sub Cbo发药药房_Click()
    lng药房ID = Cbo发药药房.ItemData(Cbo发药药房.ListIndex)
    
    str记帐人 = "所有记帐人"

    If lng药房ID <> Val(Cbo发药药房.Tag) Then
        Cbo发药药房.Tag = lng药房ID
        strUnit = GetSpecUnit(lng药房ID, gint住院药房)
        mbln是否配制中心 = CheckIsCenter(lng药房ID)
        
        Call Get配药人
        Call Get剂型
        
        DoEvents
        
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

Private Sub chkSend_Click(Index As Integer)
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    If chkSend(Index).Value = 0 Then
        blnAllUnCheck = True
        For i = 0 To chkSend.Count - 1
            If chkSend(i).Value = 1 Then
                blnAllUnCheck = False
                Exit For
            End If
        Next
        If blnAllUnCheck = True Then chkSend(Index).Value = 1
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    If Index = 0 Then
        If chkType(1).Value <> 1 Then chkType(0).Value = 1
    Else
        If chkType(0).Value <> 1 Then chkType(1).Value = 1
    End If
End Sub
Private Sub Chk显示退药待发单据_Click()
    mlng待发单据 = Chk显示退药待发单据.Value
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdOtherCon_Click()
    If cmdOtherCon.Caption = "全部条件(&C)" Then
        cmdOtherCon.Caption = "简要条件(&C)"
        mblnAllConditon = True
    Else
        cmdOtherCon.Caption = "全部条件(&C)"
        mblnAllConditon = False
    End If
    
    Call ResizeCondition
End Sub
Private Sub cmdRefresh_Click()
    ''''刷新
    BlnInRefresh = False
    
    If TabShow.Tab = 4 Then
        mblnFirstSended = False
    End If
    
    Call mnuViewRefresh_Click
End Sub

Private Sub cmd部门类型_Click()
    Dim rsTemp As adodb.Recordset
    Dim rsCount As adodb.Recordset
    Dim str科室id() As String
    Dim n As Integer
    Dim i As Integer
    Dim strCond发药类型 As String
    
    On Error GoTo errHandle
    If Me.Lvw科室.Tag <> "" Then
        If Me.Lvw科室.Tag <> tbsType.TabIndex Then
            Me.txt科室.Tag = ""
            Me.txt科室.Text = ""
        End If
    End If
    
    If TabShow.Tab = 4 Then
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
                     " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质='临床' And 服务对象 IN(2,3))" & _
                     " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By 编码||'-'||名称 "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
                     " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术','营养') And 服务对象 IN(2,3))" & _
                     " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By 编码||'-'||名称 "
        Else
            gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
                     " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
                     " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                     " Order By 编码||'-'||名称 "
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取部门科室")
        
                
        With rsTemp
            If .EOF Then
                MsgBox "没有设置该类部门！（部门管理）", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.Lvw科室.ListItems.Clear
            Me.Lvw科室.Tag = tbsType.TabIndex
            Do While Not .EOF
                Me.Lvw科室.ListItems.Add , "_" & !Id, !科室, 1, 1
                .MoveNext
            Loop
        End With
    Else
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "Select Distinct A.编码 || '-' || A.名称 科室, A.ID " & _
                " From 部门表 A, 部门性质说明 B, 未发药品记录 C, 病人费用记录 D " & _
                " Where (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null) And B.工作性质 ='临床' And B.服务对象 In (2, 3) And A.ID = B.部门id And " & _
                " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.库房id = [1] And C.单据 In (9,10) And " & _
                " C.填制日期 Between [2] And [3] And C.NO = D.NO And C.库房id = D.执行部门id And A.ID = D.开单部门id And D.病人科室id = D.开单部门id " & _
                " Order By A.编码 || '-' || A.名称 "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "Select Distinct A.编码 || '-' || A.名称 科室, A.ID " & _
                " From 部门表 A, 部门性质说明 B, 未发药品记录 C, 病人费用记录 D " & _
                " Where (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null) And B.工作性质 In ('检查','检验','治疗','手术','营养') And B.服务对象 In (2, 3) And A.ID = B.部门id And " & _
                " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.库房id = [1] And C.单据 In (9,10) And " & _
                " C.填制日期 Between [2] And [3] And C.NO = D.NO And C.库房id = D.执行部门id And A.ID = D.开单部门id And D.病人科室id <> D.开单部门id " & _
                " Order By A.编码 || '-' || A.名称 "
        Else
            gstrSQL = "Select Distinct A.编码 || '-' || A.名称 科室, A.ID " & _
                " From 部门表 A, 部门性质说明 B, 未发药品记录 C, 病人费用记录 D " & _
                " Where (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null) And B.工作性质 = '护理' And B.服务对象 In (2, 3) And A.ID = B.部门id And " & _
                " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.库房id = [1] And C.单据 In (9,10) And " & _
                " C.填制日期 Between [2] And [3] And C.NO = D.NO And C.库房id = D.执行部门id And A.ID = D.病人病区id "
                
            If mstr病区发药方式 = "" Then
                gstrSQL = gstrSQL & " And D.病人科室id = D.开单部门id "
            End If
            
            gstrSQL = gstrSQL & " Order By A.编码 || '-' || A.名称 "
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取部门科室", lng药房ID, CDate(Format(Dtp开始时间.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(Dtp结束时间.Value, "yyyy-MM-dd hh:mm:ss")))
        
        With rsTemp
            If .EOF Then
                Exit Sub
            End If
            Me.Lvw科室.ListItems.Clear
            Me.Lvw科室.Tag = tbsType.TabIndex
            
            Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
            '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
            If int离院带药 = 0 Then
                strCond发药类型 = ""
            ElseIf int离院带药 = 1 Then
                strCond发药类型 = " And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
            ElseIf int离院带药 = 2 Then
                strCond发药类型 = " And Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3'"
            ElseIf int离院带药 = 3 Then
                strCond发药类型 = " And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_4'"
            ElseIf int离院带药 = 4 Then
                strCond发药类型 = " And Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_4'"
            ElseIf int离院带药 = 5 Then
                strCond发药类型 = " And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_4'"
            ElseIf int离院带药 = 6 Then
                strCond发药类型 = " And (Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(A.扣率,0),'00')) Like '_4')"
            End If
            
'            IIf(mstr发药类型 = "", "", " And Instr([15],',' || D.发药类型 || ',') > 0")
            If mstr发药类型 <> "" Then mstr发药类型 = "," & mstr发药类型 & ","
                
            Do While Not .EOF
                gstrSQL = "Select Count(Distinct A.药品id) As 药品 " & _
                    " From 药品收发记录 A, 未发药品记录 B, 病人费用记录 C " & IIf(mstr发药类型 = "", "", " ,药品规格 D") & _
                    " Where A.单据 = B.单据 And A.NO = B.NO And A.审核人 Is Null And A.NO = C.NO And B.库房id = C.执行部门id " & strCond发药类型 & IIf(mstr发药类型 = "", "", " And A.药品ID = D.药品ID And Instr([5],',' || D.发药类型 || ',') > 0") & _
                    " And B.库房id = [2] And B.单据 In (9,10) And B.填制日期 Between [3] And [4] "
                    
                If tbsType.SelectedItem.Index - 1 = 0 Then
                    gstrSQL = gstrSQL & " And C.开单部门id = [1] And C.病人科室id=C.开单部门id "
                ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
                    gstrSQL = gstrSQL & " And C.开单部门id = [1] And C.病人科室id<>C.开单部门id "
                Else
                    If mstr病区发药方式 = "" Then
                        gstrSQL = gstrSQL & " And C.病人病区id = [1] And C.病人科室id=C.开单部门id "
                    Else
                        gstrSQL = gstrSQL & " And C.病人病区id = [1] "
                    End If
                End If
                
                Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, "取部门科室", CLng(!Id), lng药房ID, CDate(Format(Dtp开始时间.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(Dtp结束时间.Value, "yyyy-MM-dd hh:mm:ss")), mstr发药类型)
                
                Me.Lvw科室.ListItems.Add , "_" & !Id, !科室 & "(" & rsCount!药品 & "种药品待发）", 1, 1
                .MoveNext
            Loop
        End With
    End If
    
    Lvw科室.Move fraCondition.Left + fraConNormal.Left + txt科室.Left - 10, fraCondition.Top + txt科室.Top + txt科室.Height + 60, txt科室.Width, 4000
    Lvw科室.Visible = True
    Lvw科室.SetFocus
    Lvw科室.ZOrder 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

    
Private Sub cmd给药途径_Click()
    Lvw给药途径.Move fraCondition.Left + fraConExpand.Left + txt给药途径.Left - 10, fraCondition.Top + fraConExpand.Top + txt给药途径.Top + txt给药途径.Height + 60, txt给药途径.Width, 3000
    Lvw给药途径.Visible = True
    Lvw给药途径.SetFocus
    Lvw给药途径.ZOrder 0
End Sub


Private Sub cmd药品剂型_Click()
    Lvw剂型.Move fraCondition.Left + fraConExpand.Left + txt药品剂型.Left - 10, fraCondition.Top + fraConExpand.Top + txt药品剂型.Top + txt药品剂型.Height + 60, txt药品剂型.Width, 3000
    Lvw剂型.Visible = True
    Lvw剂型.SetFocus
    Lvw剂型.ZOrder 0
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuPatiInfo, 2, fraCondition.Left + fraConNormal.Left + lblPatiInputType.Left - 30, fraCondition.Top + fraConNormal.Top + lblPatiInputType.Top + lblPatiInputType.Height + 30
    End If
End Sub
Private Sub Lvw给药途径_DblClick()
    Dim n As Integer
    
    With Lvw给药途径
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt给药途径.Tag = ""
        Me.txt给药途径.Text = ""
        
        '如果选择了全选，则不用取所有给药途径了
        If .ListItems(1).Checked Then
            Me.txt给药途径.Tag = ""
            Me.txt给药途径.Text = "所有给药途径"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txt给药途径.Tag = IIf(Me.txt给药途径.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt给药途径.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt给药途径.Text = IIf(Me.txt给药途径.Text = "", .ListItems(n).Text, Me.txt给药途径.Text & "," & .ListItems(n).Text)
            End If
        Next
    
        '如果当前双击的给药途径未被选上，将当前双击的给药途径也加入到编辑框中
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txt给药途径.Tag = IIf(Me.txt给药途径.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt给药途径.Tag & "," & Mid(.SelectedItem.Key, 2))
            Me.txt给药途径.Text = IIf(Me.txt给药途径.Text = "", .SelectedItem.Text, Me.txt给药途径.Text & "," & .SelectedItem.Text)
        End If
        .Visible = False
    End With
End Sub

Private Sub Lvw给药途径_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw给药途径
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有给药途径" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub Lvw给药途径_LostFocus()
    Lvw给药途径.Visible = False
End Sub

Private Sub Lvw给药途径_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If mnuTypeItem.Item(0).Caption <> "-" Then
            PopupMenu mnuType, 2
        End If
    End If
End Sub

Private Sub Lvw剂型_DblClick()
    Dim n As Integer
    
    With Lvw剂型
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt药品剂型.Text = ""
        
        '如果选择了全选，则不用取所有给药途径了
        If .ListItems(1).Checked Then
             Me.txt药品剂型.Text = "所有药品剂型"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
            End If
        Next
    
        '如果当前双击的给药途径未被选上，将当前双击的给药途径也加入到编辑框中
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
        End If
        .Visible = False
    End With
End Sub


Private Sub Lvw剂型_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw剂型
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有药品剂型" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub


Private Sub Lvw剂型_LostFocus()
    Lvw剂型.Visible = False
End Sub
Private Sub Lvw科室_DblClick()
    Dim n As Integer
    
    With Me.Lvw科室
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt科室.Tag = ""
        Me.txt科室.Text = ""
        For n = 1 To .ListItems.Count
            If .ListItems(n).Checked Then
                Me.txt科室.Tag = IIf(Me.txt科室.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt科室.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt科室.Text = IIf(Me.txt科室.Text = "", .ListItems(n).Text, Me.txt科室.Text & "," & .ListItems(n).Text)
            End If
        Next
    
        '如果当前双击的科室未被选上，将当前双击的科室也加入到对方科室编辑框中
        If .SelectedItem.Checked = False Then
            .SelectedItem.Checked = True
            Me.txt科室.Tag = IIf(Me.txt科室.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt科室.Tag & "," & Mid(.SelectedItem.Key, 2))
            Me.txt科室.Text = IIf(Me.txt科室.Text = "", .SelectedItem.Text, Me.txt科室.Text & "," & .SelectedItem.Text)
        End If
        .Visible = False
        txt科室.SetFocus
    End With
End Sub





Private Sub Lvw科室_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    
    For n = 1 To Lvw科室.ListItems.Count
        Lvw科室.ListItems(n).Selected = False
    Next
    
    Item.Selected = True
End Sub


Private Sub Lvw科室_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.Lvw科室.SelectedItem Is Nothing Then Exit Sub
        Call Lvw科室_DblClick
    End Select
End Sub


Private Sub Lvw科室_LostFocus()
    Me.Lvw科室.Visible = False
    txt科室.SetFocus
End Sub
Private Sub mnuBillItem_Click(Index As Integer)
    mnuBillItem(Index).Checked = Not mnuBillItem(Index).Checked
    
    If (Me.mnuBillItem(Index).Caption = "领/退药人" Or Me.mnuBillItem(Index).Caption = "退药人") Then
        mbln显示领退药人 = (Me.mnuBillItem(Index).Checked)
    End If
                        
    Call SetColHideByMenu(mnuBillItem(Index), IIf(TabShow.Tab = 0, Bill未发药清单, Bill已发药清单))
End Sub
Private Sub mnuDrugCodeName_Click(Index As Integer)
    Dim n As Integer
    Dim strSave As String
    
    If mnuDrugCodeName(Index).Checked = True Then Exit Sub
    
    For n = 0 To mnuDrugCodeName.Count - 1
        mnuDrugCodeName(n).Checked = False
    Next
    
    mnuDrugCodeName(Index).Checked = True
    
    '保存设置
    int药品名称 = Index
    
    strSave = Index & "|" & "药品名称"
    
    For n = 0 To mnuBillItem.Count - 1
        strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & mnuBillItem(n).Caption
    Next
    
    zldatabase.SetPara "列设置", strSave, glngSys, 1342
    
    '更新数据
    mnuViewRefresh_Click
End Sub

Private Sub mnuInfoItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuInfoItem.UBound
        mnuInfoItem(i).Checked = (i = Index)
    Next
    
    strItem = Split(mnuInfoItem(Index).Caption, "(")(0)
    lblPatiInputType.Caption = strItem & "↓"
    lblPatiInputType.Tag = Index
    
    txtPati.Text = ""
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    If Val(lblPatiInputType.Tag) = PatiInfo.就诊卡 Then
        If gtype_UserSysParms.P12_就诊卡是否密文显示 Then
            txtPati.PasswordChar = "*"
        End If
        txtPati.MaxLength = gtype_UserSysParms.P20_就诊卡号长度
    End If
    
'    txtPati.SetFocus
    
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

Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号，lngCmd=0时需要
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New adodb.Recordset
    Dim str药品 As String, str用法 As String, lng药品ID As Long, str单量单位 As String
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If mstrNo = "" Then Exit Function
        
        
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
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, mInt单据)
    
    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If
    
    mlngPatiID = rsTmp!病人ID
    mstr挂号单 = NVL(rsTmp!挂号单)
    mlng主页ID = rsTmp!主页id
    
    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    If mlngPatiID <> mlngPassPati Then
        If mstr挂号单 <> "" Then               '门诊病人
            strSQL = "Select 病人ID,Count(Distinct Trunc(登记时间)) as 就诊次数 From 病人挂号记录 Where 病人ID=[1] Group by 病人ID"
            strSQL = "Select D.就诊次数,A.姓名,A.性别,A.出生日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,(" & strSQL & ") D,人员表 E" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=D.病人ID" & _
                " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mstr挂号单)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlngPatiID, rsTmp!就诊次数, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), "")
        Else                                    '住院病人
            strSQL = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngPatiID, mlng主页ID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlngPatiID, mlng主页ID, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), _
                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        End If
        mlngPassPati = mlngPatiID
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        If TabShow = 0 Then
           '取药品名称
            str药品 = Bill未发药清单.TextMatrix(lngRow, 列名_未发药清单.药品名称)
            lng药品ID = Bill未发药清单.TextMatrix(lngRow, 列名_未发药清单.药品ID)
            str单量单位 = Bill未发药清单.TextMatrix(lngRow, 列名_未发药清单.单量单位)
            '取药品给药途径
            str用法 = Bill未发药清单.TextMatrix(lngRow, 列名_未发药清单.用法)
        Else
            '取药品名称
            str药品 = Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.药品名称)
            lng药品ID = Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.药品ID)
            str单量单位 = Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.单量单位)
            '取药品给药途径
            str用法 = Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.用法)
        End If
        
        If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
        If InStr(str药品, "(") > 0 Then str药品 = Left(str药品, InStr(str药品, "(") - 1)
        '传入查询药品信息
        Call PassSetQueryDrug(lng药品ID, str药品, str单量单位, str用法)
            
        '设置菜单可用状态
        Call SetPassMenuState
        
        AdviceCheckWarn = 1 '表示可以弹出菜单

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
Private Sub LoadPASS(ByVal BillStyle As Integer, ByVal BillNo As String)
    Dim strSQL As String
    Dim rs As New adodb.Recordset
    Dim n As Integer
    Dim strCondition As String
    On Error GoTo errHandle

    strSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,0) 挂号单,B.医嘱序号 " & _
        " From 药品收发记录 A,病人费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle)

    If rs!挂号单 <> 0 Then
        strSQL = "Select A.ID,B.名称 as 用法 From 病人医嘱记录 A,诊疗项目目录 B" & _
        " Where A.诊疗类别='E' And B.操作类型 IN('2','4') And A.诊疗项目ID=B.ID And A.病人ID=[1] And A.挂号单=[2] "
        strSQL = _
            " Select A.ID,A.相关ID,Nvl(A.婴儿,0) as 婴儿,A.收费细目ID,A.诊疗类别,A.医嘱内容," & _
            " A.单次用量,B.计算单位,C.用法,A.频率次数,A.开嘱医生,A.开嘱时间,A.执行终止时间," & _
            " nvl(A.审查结果,-1) 审查结果,nvl(A.主页id,0) 主页id,nvl(A.挂号单,'') 挂号单,A.病人id " & _
            " From 病人医嘱记录 A,诊疗项目目录 B,(" & strSQL & ") C" & _
            " Where A.诊疗项目ID=B.ID And A.相关ID=C.ID And A.诊疗类别 IN('5','6','7') And A.收费细目ID is Not Null" & _
            " And A.医嘱状态<>4 And (A.医嘱状态 Not IN(8,9) Or A.医嘱期效=1) " & _
            " And A.病人ID=[1] And A.挂号单=[2] " & _
            " And A.开始执行时间 is Not NULL" & _
            " Order by Nvl(A.婴儿,0),A.序号"
        Set mrsPASS = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rs!病人ID), CStr(rs!挂号单))
        
    ElseIf rs!主页id <> 0 Then
        strSQL = "Select A.ID,B.名称 as 用法 From 病人医嘱记录 A,诊疗项目目录 B" & _
        " Where A.诊疗类别='E' And B.操作类型 IN('2','4') And A.诊疗项目ID=B.ID And A.病人ID=[1] And A.主页ID=[2] "
        strSQL = _
            " Select A.ID,A.相关ID,Nvl(A.婴儿,0) as 婴儿,A.收费细目ID,A.诊疗类别,A.医嘱内容," & _
            " A.单次用量,B.计算单位,C.用法,A.频率次数,A.开嘱医生,A.开嘱时间,A.执行终止时间," & _
            " nvl(A.审查结果,-1) 审查结果,nvl(A.主页id,0) 主页id,nvl(A.挂号单,'') 挂号单,A.病人id " & _
            " From 病人医嘱记录 A,诊疗项目目录 B,(" & strSQL & ") C" & _
            " Where A.诊疗项目ID=B.ID And A.相关ID=C.ID And A.诊疗类别 IN('5','6','7') And A.收费细目ID is Not Null" & _
            " And A.医嘱状态<>4 And (A.医嘱状态 Not IN(8,9) Or A.医嘱期效=1) " & _
            " And A.病人ID=[1] And A.主页ID=[2] " & _
            " And A.开始执行时间 is Not NULL" & _
            " Order by Nvl(A.婴儿,0),A.序号"
        Set mrsPASS = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rs!病人ID), CLng(rs!主页id))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetDrugFormat()
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '取得药品名称的格式方式
    strSave = zldatabase.GetPara("列设置", glngSys, 1342)
    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|科室,0|开单医生,0|状态,0|类型,0|NO,0|记帐员,0|床号,0|姓名,0|住院号,0|规格,0|产地,0|批号,0|付,0|数量,0|已退数,0|准退数,0|退药数,0|单价,0|金额,0|单量,0|频次,0|用法,0|记帐时间,0|说明,0|操作员,0|发药时间,0|领/退药人,0|库房货位"
    arrColumn = Split(strSave, ",")
    int药品名称 = Val(Split(arrColumn(0), "|")(0))
End Sub

Private Function GetColDefaultWidth(ByVal Bill As MSHFlexGrid, ByVal Col As Integer) As Integer
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '返回指定表格指定列的默认宽度
    strSave = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & Me.Name, Bill.Name & "列默认宽度", "")
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    For intRow = 0 To intRows
'        intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1), Bill)
        If Split(arrColumn(intRow), "|")(0) = Bill.TextMatrix(0, Col) Then
            GetColDefaultWidth = Split(arrColumn(intRow), "|")(1)
            Exit For
        End If
    Next
End Function

Private Sub SaveColDefaultWidth(ByVal Bill As MSHFlexGrid)
    '保存列的默认宽度
    Dim strSave As String
    Dim i As Integer
    
    For i = 0 To Bill.Cols - 1
        strSave = strSave & Bill.TextMatrix(0, i) & "|" & Bill.ColWidth(i) & ","
    Next
    SaveSetting "ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "frm部门发药管理", Bill.Name & "列默认宽度", strSave
    
End Sub

Private Sub SetColHide(ByVal Bill As MSHFlexGrid)
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '根据用户的本地参数设置，显示或者隐藏部分列
    strSave = zldatabase.GetPara("列设置", glngSys, 1342)

    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|科室,0|开单医生,0|状态,0|类型,0|NO,0|记帐员,0|床号,0|姓名,0|住院号,0|规格,0|产地,0|批号,0|付,0|数量,0|已退数,0|准退数,0|退药数,0|单价,0|金额,0|单量,0|频次,0|用法,0|记帐时间,0|说明,0|操作员,0|发药时间,0|领/退药人,0|库房货位"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    For intRow = 0 To intRows
        intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1), Bill)
        If intCol > -1 Then
            If Split(arrColumn(intRow), "|")(1) = "药品名称" Then
                int药品名称 = Val(Split(arrColumn(intRow), "|")(0))
            Else
                If Val(Split(arrColumn(intRow), "|")(0)) = 1 Then
                    Bill.ColWidth(intCol) = 0
                ElseIf Bill.ColWidth(intCol) = 0 Then       '如果要显示的列宽为0，则取默认的列宽
                    Bill.ColWidth(intCol) = GetColDefaultWidth(Bill, intCol)
                End If
            End If
        End If
    Next
    
    '部分列要受权限影响，这时要根据权限来确定是否显示
    If Bill.Name = "Bill未发药清单" Then
        If UserPrivDetail.Priv_医生查询 = False Then
            Bill.ColWidth(列名_未发药清单.开单医生) = 0
        Else
            Bill.ColWidth(列名_未发药清单.开单医生) = 1100
        End If
    End If
End Sub


Private Sub SetColHideByMenu(ByVal MenuObj As Menu, ByVal Bill As MSHFlexGrid)
    Dim intCol As Integer
    Dim strSave As String
    Dim n As Integer
        
    intCol = GetDetailCol(MenuObj.Caption, Bill)
    If intCol > -1 Then
        If MenuObj.Checked = False Then
            Bill.ColWidth(intCol) = 0
        Else
            Bill.ColWidth(intCol) = GetColDefaultWidth(Bill, intCol)
        End If
    End If
    
    '保存设置
    strSave = int药品名称 & "|" & "药品名称"
    
    For n = 0 To mnuBillItem.Count - 1
        If mnuBillItem(n).Caption = "退药人" Then
            strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & "领/退药人"
        Else
            strSave = strSave & "," & IIf(mnuBillItem(n).Checked, "0", "1") & "|" & mnuBillItem(n).Caption
        End If
    Next
    
    zldatabase.SetPara "列设置", strSave, glngSys, 1342
End Sub
Private Sub SetColMenu()
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    Dim n As Integer
    
    '取本地注册表来设置列项目控制菜单
    strSave = zldatabase.GetPara("列设置", glngSys, 1342)
    
    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|科室,0|开单医生,0|状态,0|类型,0|NO,0|记帐员,0|床号,0|姓名,0|住院号,0|规格,0|产地,0|批号,0|付,0|数量,0|已退数,0|准退数,0|退药数,0|单价,0|金额,0|单量,0|频次,0|用法,0|记帐时间,0|说明,0|操作员,0|发药时间,0|领/退药人,0|库房货位"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    
    For n = 0 To Me.mnuDrugCodeName.Count - 1
        Me.mnuDrugCodeName(n).Checked = False
    Next
    
    For n = 0 To Me.mnuBillItem.Count - 1
        Me.mnuBillItem(n).Checked = False
    Next
    
    mbln显示领退药人 = False
    
    For intRow = 0 To intRows
        If Split(arrColumn(intRow), "|")(1) = "药品名称" Then
            Me.mnuDrugCodeName(Val(Split(arrColumn(intRow), "|")(0))).Checked = True
        Else
            For n = 0 To Me.mnuBillItem.Count - 1
                If Me.mnuBillItem(n).Caption = Split(arrColumn(intRow), "|")(1) And Me.mnuBillItem(n).Visible = True Then
                    If Val(Split(arrColumn(intRow), "|")(0)) = 0 Then
                        Me.mnuBillItem(n).Checked = True
                        If Me.mnuBillItem(n).Caption = "领/退药人" Or Me.mnuBillItem(n).Caption = "退药人" Then
                            mbln显示领退药人 = True
                        End If
                    End If
                End If
            Next
        End If
    Next
    If UserPrivDetail.Priv_医生查询 = False Then
        mnuBillItem(1).Visible = False
    Else
        mnuBillItem(1).Visible = True
    End If
    
    If mbln药品储备 = False Then
        Me.mnuBillItem(29).Visible = False
    End If
End Sub
Private Sub Bill汇总发药_EnterCell()
    Dim Col As Integer
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    Bill退药销帐.Visible = False
    With Bill汇总发药
        .Height = TabShow.Height - TabShow.TabHeight - 120
        .Width = TabShow.Width - 150
    End With
    
    With Bill退药销帐
        .Visible = False
        .Left = Bill汇总发药.Left
        .Height = 1400
        .Width = TabShow.Width - 150
    End With
    
    Call SetSelectColor(Bill汇总发药)
    
    If Lng汇总显示 = 0 Then Exit Sub
    
    If txt留存数.Visible Then
        txt留存数_LostFocus
        txt留存数.Visible = False
    End If
    If CurCell.Row = 0 Or CurCell.Row >= Bill汇总发药.rows - 2 Or Bill汇总发药.TextMatrix(CurCell.Row, 0) = "小计" Then
        Exit Sub
    End If
    
    DoEvents
    
    If mbln汇总发药 = True Then
        If LoadDataInBill销帐清单(Val(Bill汇总发药.TextMatrix(CurCell.Row, 列名_科室汇总清单.领药部门id)), Val(Bill汇总发药.TextMatrix(CurCell.Row, 列名_科室汇总清单.药品ID))) = True Then
            Bill汇总发药.Height = Bill汇总发药.Height - Bill退药销帐.Height - 25
            
            Bill退药销帐.Visible = True
            Bill退药销帐.Top = Bill汇总发药.Top + Bill汇总发药.Height + 25
        End If
    End If
        
    DoEvents
    
    '处理留存
    If IsHavePrivs(mstrPrivs, "修改留存数量") = False Then Exit Sub
    
    If CurCell.Col <> 列名_科室汇总清单.留存数量 And CurCell.Col <> 列名_科室汇总清单.实发数量 Then
        Exit Sub
    End If
    
    If Val(Bill汇总发药.TextMatrix(CurCell.Row, 列名_科室汇总清单.实发数量)) < 0 Then
        Exit Sub
    End If
    
    LngLastRow = CurCell.Row
    lngLastCol = CurCell.Col
    
    lngLeft = TabShow.Left + Bill汇总发药.Left + CurCell.CellLeft - 20
    lngTop = TabShow.Top + Bill汇总发药.Top + CurCell.CellTop + 20
    
    lngWidth = CurCell.CellWidth - 20

    With txt留存数
        If .Visible = False Then
            .Alignment = 1
            .Move lngLeft, lngTop, lngWidth
            .Visible = True
            .ZOrder 0
            .SetFocus
            .Text = FormatEx(Val(Bill汇总发药.TextMatrix(CurCell.Row, CurCell.Col)), 5)
        End If
    End With
    Call SelAll(txt留存数)
End Sub

Private Sub Bill汇总发药_GotFocus()
    Bill汇总发药_EnterCell
End Sub

Private Sub Bill拒发药清单_DblClick()
    Call Bill拒发药清单_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Bill拒发药清单_EnterCell()
    Call SetSelectColor(Bill拒发药清单)
End Sub

Private Sub Bill拒发药清单_GotFocus()
    Bill拒发药清单_EnterCell
End Sub

Private Sub Bill拒发药清单_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    With Bill拒发药清单
        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
        
        Select Case Trim(.TextMatrix(.Row, 1))
        Case "恢复"
            Call UpdateRsByMenu(Nop_3, 3)
        Case "不处理"
            Call UpdateRsByMenu(ResumeDo, 3)
        End Select
    End With
End Sub

Private Sub Bill拒发药清单_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim MenuDefault As Menu
        With Bill拒发药清单
            If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
            If Trim(.TextMatrix(.Row, 1)) = "合计" Then Exit Sub
            
            Set MenuDefault = SetMenuCheck(PopMenu_3)
            PopupMenu PopMenu_3, 2, , , MenuDefault
        End With
    End If
End Sub

Private Sub Bill缺药清单_EnterCell()
    Call SetSelectColor(Bill缺药清单)
End Sub

Private Sub Bill缺药清单_GotFocus()
    Bill缺药清单_EnterCell
End Sub

Private Sub Bill未发药清单_DblClick()
    Call Bill未发药清单_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Bill未发药清单_EnterCell()
    Dim rsTmp As adodb.Recordset
    Dim rs批号 As New adodb.Recordset
    Dim lng批次 As Long, lng药品ID As Long, Dbl数量 As Double, blnAllow As Boolean
    Dim ArrayPhysic
    
    On Error GoTo errHandle
    If Not BlnEnterCell Then Exit Sub
    Call SetSelectColor(Bill未发药清单)
    Cbo批号.Clear
    Cbo批号.Visible = False
    
    mstrNo = Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.NO)
    mInt单据 = Val(Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.单据))
            
    '设置cmdAlley按钮状态
    If mblnStarPass Then
        '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就不显示cmdAlley按钮
        gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
            " From 药品收发记录 A,病人费用记录 B,病人医嘱记录 C " & _
            " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
            " And A.单据=[2] And A.no=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mInt单据)
        If rsTmp.RecordCount = 0 Then
            If cmdAlley.Visible Then cmdAlley.Visible = False
        Else
            If Not cmdAlley.Visible Then cmdAlley.Visible = True
        End If
    End If
    
    
    '如果该药品药房分批核算，则提取该药品所有批次供用户选择
    '售价相同(指时价药品)且库存充足，则允许药房人员调整批次
    With Bill未发药清单
        If Not (Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "发药") Then Exit Sub
    End With
    
    If CurCell.Col = 列名_未发药清单.批号 Then
        RecChangeData.MoveFirst
        RecChangeData.Find "位置=" & CurCell.Row
        If RecChangeData.EOF Then Exit Sub
        If RecChangeData!分批 = 0 Then Exit Sub
        lng批次 = RecChangeData!批次
        lng药品ID = RecChangeData!药品ID
        Dbl数量 = FormatEx(RecChangeData!实际数量, 5)
        ArrayPhysic = Split(GetPhysicDict(lng药品ID), "^")        '获取该药品的相关信息
        
        '如果存在发药记录且部分退药，则不允许修改批次信息
        blnAllow = False
        
        gstrSQL = " Select count(*) Records From 药品收发记录 A,药品收发记录 B " & _
        " Where (Mod(A.记录状态,3)=0 or A.记录状态=1) And A.审核人 Is Not NULL And B.ID=[1] " & _
        " And A.NO=B.NO And A.单据=B.单据 And A.药品ID=B.药品ID And Nvl(A.批次,0)=Nvl(B.批次,0)"
        Set rs批号 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(RecChangeData!Id))
        
        
        blnAllow = (rs批号!Records = 0)
        
        '提取所有批次信息
        gstrSQL = " SELECT B.上次批号 批号,B.批次,ROUND(B.实际数量/" & ArrayPhysic(3) & ",2) 数量" & _
         " FROM 药品规格 A,药品库存 B,收费价目 C,收费项目目录 F" & _
         " WHERE A.药品ID = B.药品ID AND b.药品ID=F.ID" & _
         " AND B.库房ID = [1] AND B.药品ID=[2] AND A.药品ID = C.收费细目ID" & _
         " AND ((SYSDATE BETWEEN C.执行日期 AND C.终止日期) OR C.终止日期 IS NULL)" & _
         " AND NVL(批次,0)<>0 AND NVL(实际数量,0)<>0 AND 性质=1" & _
         " AND ROUND(DECODE(F.是否变价,NULL,C.现价,0,C.现价,B.实际金额/B.实际数量),2)=" & _
         "     (SELECT ROUND(DECODE(F.是否变价,NULL,C.现价,0,C.现价,B.实际金额/B.实际数量),2) 单价" & _
         "     FROM 药品规格 A,药品库存 B,收费价目 C,收费项目目录 F" & _
         "     WHERE A.药品ID = B.药品ID AND b.药品ID=f.ID " & _
         "     AND B.库房ID = [1] AND B.药品ID=[2] AND A.药品ID = C.收费细目ID" & _
         "     AND ((SYSDATE BETWEEN C.执行日期 AND C.终止日期) OR C.终止日期 IS NULL)" & _
         "     AND NVL(批次,0)<>0 AND NVL(实际数量,0)<>0 AND 性质=1 AND NVL(批次,0)=[3])" & _
         " AND ROUND(B.实际数量/" & ArrayPhysic(3) & ",2)>=[4] AND (NVL(A.药房分批,0)=0 OR (NVL(A.药房分批,0)=1 AND (效期 IS NULL OR 效期>TRUNC(SYSDATE))))" & _
         " ORDER BY B.批次"
        Set rs批号 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, lng药品ID, lng批次, Dbl数量)
        
        With rs批号
            Do While Not .EOF
                If (!批次 <> lng批次 And blnAllow) Or !批次 = lng批次 Then
                    Cbo批号.AddItem IIf(IsNull(!批号), "", !批号) & "(" & !批次 & ")"
                    Cbo批号.ItemData(Cbo批号.NewIndex) = !批次
                End If
                .MoveNext
            Loop
        End With
        Call LocateCboItemData(Cbo批号, lng批次)
        Call ShowCbo
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Bill未发药清单_GotFocus()
    Bill未发药清单_EnterCell
End Sub

Private Sub Bill未发药清单_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    
    CurCell.Col = 0
    Cbo批号.Visible = False
    
    With Bill未发药清单
        If Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "" Then Exit Sub
        If Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "缺药" Then Exit Sub
            
        RecChangeData.MoveFirst
        RecChangeData.Find "位置=" & Bill未发药清单.Row
        If RecChangeData.EOF Then Exit Sub
        If RecChangeData!审核人 = "" And Int允许未审核处方发药 = 0 Then
            Select Case Trim(.TextMatrix(.Row, 列名_未发药清单.状态))
            Case "拒发"
                Call UpdateRsByMenu(Nop_1, 1)
            Case "不处理"
                Call UpdateRsByMenu(HandBack, 1)
            End Select
            Exit Sub
        End If
        
        Select Case Trim(.TextMatrix(.Row, 列名_未发药清单.状态))
        Case "发药"
            Call UpdateRsByMenu(HandBack, 1)
        Case "拒发"
            Call UpdateRsByMenu(Lack, 1)
        Case "缺药"
            Call UpdateRsByMenu(Nop_1, 1)
        Case "不处理"
            Call UpdateRsByMenu(Consignment, 1)
        End Select
    End With
End Sub

Private Sub Bill未发药清单_LostFocus()
    Call Cbo批号_LostFocus
End Sub

Private Sub Bill未发药清单_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim str药品 As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
    intCurRow = Bill未发药清单.MouseRow
    intCurCol = Bill未发药清单.MouseCol
    
    If Button = 2 Then
        If intCurRow = 0 Then
            PopupMenu mnuColHide, 2
            Exit Sub
        End If
        If intCurCol > 0 Then
            Dim MenuDefault As Menu
        
            CurCell.Col = 1
            Cbo批号.Visible = False
            
            With Bill未发药清单
                Consignment.Enabled = True
                If Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "" Then Exit Sub
                If Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "合计" Then Exit Sub
                If Trim(.TextMatrix(.Row, 列名_未发药清单.状态)) = "缺药" Then Exit Sub
                If RecChangeData.RecordCount = 0 Then Exit Sub
                
                RecChangeData.MoveFirst
                RecChangeData.Find "位置=" & Bill未发药清单.Row
                If RecChangeData.EOF Then Exit Sub
                If RecChangeData!审核人 = "" And Int允许未审核处方发药 = 0 Then Consignment.Enabled = False
                
                Set MenuDefault = SetMenuCheck(PopMenu_1)
                PopupMenu PopMenu_1, 2, , , MenuDefault
            End With
        ElseIf intCurCol = 0 Then
            mstrNo = Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.NO)
            mInt单据 = Val(Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.单据))
            
            '检查Pass状态
            If AdviceCheckWarn(0, Bill未发药清单.Row) >= 0 Then PopupMenu mnuPass, 2
        End If
            
    End If
End Sub

Private Sub Bill未发药清单_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String
    Dim bln是否显示分组 As Boolean
        
    '保存排序列
    With Bill未发药清单
        If Button <> 1 Then Exit Sub
        If .MouseRow <> 0 Then Exit Sub

        strColumn = .TextMatrix(.MouseRow, .MouseCol)
        If InStr(1, gstr排序列名, "|" & strColumn & "|") = 0 Then Exit Sub
        If strColumn = "药品名称" Then strColumn = "品名"
        
        '只有按NO排序时才显示分组
        If strColumn = "NO" Then bln是否显示分组 = True

        '如果列名相同，则改变排序方式；否则按升序方式
        If str排序_未发药 Like "*" & strColumn & "*" Then
            str排序_未发药 = ExchangeOrder(str排序_未发药)
        Else
            str排序_未发药 = strColumn & strAsc
        End If
    End With

    '重新显示未发药清单
    Call ClearCons
    Call LoadDataInBill未发药清单
    Call SetGroup(Bill未发药清单, bln是否显示分组)
    
End Sub

Private Sub Bill已发药清单_DblClick()
    Call Bill已发药清单_KeyDown(vbKeySpace, 0)
End Sub

Private Sub Bill已发药清单_EnterCell()
    Dim rsTmp As adodb.Recordset
    Dim rs批号 As New adodb.Recordset
    Dim lng批次 As Long, lng药品ID As Long, Dbl数量 As Double
    Dim ArrayPhysic
    
    On Error GoTo errHandle
    mnuFileRestore = False
    If TxtInput.Visible Then
        Call TxtInput_LostFocus
        TxtInput.Visible = False
    End If
    
    mstrNo = Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.NO)
    mInt单据 = Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.单据))
            
    '设置cmdAlley按钮状态
    If mblnStarPass Then
        '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就不显示cmdAlley按钮
        gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
            " From 药品收发记录 A,病人费用记录 B,病人医嘱记录 C " & _
            " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
            " And A.单据=[2] And A.no=[1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mInt单据)
        If rsTmp.RecordCount = 0 Then
            If cmdAlley.Visible Then cmdAlley.Visible = False
        Else
            If Not cmdAlley.Visible Then cmdAlley.Visible = True
        End If
    End If
    
    If Not IsHavePrivs(mstrPrivs, "退药") Then Exit Sub
    
    '显示退药数文本框，缺省为当前单位格内容，允许用户修改。
    '如果输入值非法（零、空格、非法串、大于全部可退数量）则缺省为全退
    With Bill已发药清单
        .Col = 列名_已发药清单.退药数
        Call SetSelectColor(Bill已发药清单)
    End With
    
    '强制设定焦点为退药数列
    With RecChangeSendedData
        If CurCell.Col = 列名_已发药清单.退药数 And Bill已发药清单.ColWidth(列名_已发药清单.退药数) > 0 Then
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            .Find "位置=" & CurCell.Row
            If .EOF Then Exit Sub
            If !可操作 = 0 Then Exit Sub        '表示该记录是否是原始记录

            '保证每次EnterCell事件激活时，都设置了菜单"打印退药通知单"
            mnuFileRestore = (!可操作 = 3)
            If Not (Trim(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.状态)) = "退药") Then Exit Sub

            TxtInput.Tag = Val(!准退数)
            TxtInput.Text = FormatEx(Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.退药数)), 5)
            Call ShowTxt
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Bill已发药清单_GotFocus()
    Bill已发药清单_EnterCell
End Sub

Private Sub Bill已发药清单_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = vbKeySpace) Then Exit Sub
    If Not IsHavePrivs(mstrPrivs, "退药") Then Exit Sub
    CurCell.Col = 0
    TxtInput.Visible = False
    
    With Bill已发药清单
        If Trim(.TextMatrix(.Row, 列名_已发药清单.状态)) = "" Then Exit Sub
        With RecChangeSendedData
            If .RecordCount = 0 Then
                MsgErr "数据有变化，请刷新后再试！"
                Exit Sub
            End If
            .MoveFirst
            .Find "位置=" & Bill已发药清单.Row
            If .EOF Then Exit Sub
            If !可操作 <> 1 Then Exit Sub
        End With
        
        Select Case Trim(.TextMatrix(.Row, 列名_已发药清单.状态))
        Case "退药"
            Call UpdateRsByMenu(Nop_1, 2)
        Case "不处理"
            If .ColWidth(列名_已发药清单.退药数) = 0 Then Exit Sub
            Call UpdateRsByMenu(Restore, 2)
        End Select
    End With
    Call Bill已发药清单_EnterCell
End Sub

Private Sub Bill已发药清单_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim str药品 As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
    intCurRow = Bill已发药清单.MouseRow
    intCurCol = Bill已发药清单.MouseCol

    If CurCell.Col <> 列名_已发药清单.退药数 Then
        CurCell.Col = 0
        TxtInput.Visible = False
    End If

    If Button = 2 Then
        If intCurRow = 0 Then
            PopupMenu mnuColHide, 2
            Exit Sub
        End If
        If intCurCol > 0 Then
            Dim MenuDefault As Menu
            With Bill已发药清单
                If Trim(.TextMatrix(.Row, 列名_已发药清单.状态)) = "" Then Exit Sub
                If Trim(.TextMatrix(.Row, 列名_已发药清单.状态)) = "合计" Then Exit Sub
                With RecChangeSendedData
                    If .RecordCount = 0 Then
                        MsgErr "数据有变化，请刷新后再试！"
                        Exit Sub
                    End If
                    .MoveFirst
                    .Find "位置=" & Bill已发药清单.Row
                    If .EOF Then Exit Sub
                    If !可操作 <> 1 Then Exit Sub
                End With
                
                Set MenuDefault = SetMenuCheck(PopMenu_2)
                PopupMenu PopMenu_2, 2, , , MenuDefault
            End With
        ElseIf intCurCol = 0 Then
            mstrNo = Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.NO)
            mInt单据 = Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.单据))

            '检查Pass状态
            If AdviceCheckWarn(0, Bill已发药清单.Row) >= 0 Then PopupMenu mnuPass, 2
        End If
    End If
End Sub

Private Sub Bill已发药清单_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String
    Dim bln是否显示分组 As Boolean
    
    '保存排序列
    With Bill已发药清单
        If Button <> 1 Then Exit Sub
        If .MouseRow <> 0 Then Exit Sub
'        If Chk清单.Value = 1 Then Exit Sub
        
        strColumn = .TextMatrix(.MouseRow, .MouseCol)
        If InStr(1, gstr排序列名, "|" & strColumn & "|") = 0 Then Exit Sub
        If strColumn = "药品名称" Then strColumn = "品名"
        
        '只有不显示过程单据并且是按NO排序时才显示分组
        If Chk清单.Value = 0 And strColumn = "NO" Then bln是否显示分组 = True
        
        '如果列名相同，则改变排序方式；否则按升序方式
        If str排序_发退药 Like "*" & strColumn & "*" Then
            str排序_发退药 = ExchangeOrder(str排序_发退药)
        Else
            str排序_发退药 = strColumn & strAsc
        End If
    End With
    
    '重新显示未发药清单
    Call ClearBill(Bill已发药清单)
    Call LoadDataInBill已发药清单
    Call SetGroup(Bill已发药清单, bln是否显示分组)
End Sub

Private Sub Cbo批号_Click()
    RecChangeData.MoveFirst
    RecChangeData.Find "位置=" & CurCell.Row
    If RecChangeData.EOF Then Exit Sub
    
    With RecChangeData
        If !分批 = 0 Then Exit Sub
        If !批次 = Cbo批号.ItemData(Cbo批号.ListIndex) Then Exit Sub
        !批次 = Cbo批号.ItemData(Cbo批号.ListIndex)
        !批号 = Cbo批号.Text
        .Update
    End With
    With Bill未发药清单
        .TextMatrix(.Row, 列名_未发药清单.批号) = Cbo批号
    End With
End Sub

Private Sub Cbo批号_LostFocus()
    On Error Resume Next
    
    If InStr(1, "Bill未发药清单,Cbo批号", ActiveControl.Name) = 0 Then
        CurCell.Col = 0
        Cbo批号.Visible = False
    End If
End Sub

Private Sub Chk清单_Click()
    '记录当前页的CHECK框的状态，在切换后可以恢复原来的状态
    If TabShow.Tab = 1 Then
        Chk清单.Tag = Chk清单.Value & Mid(Chk清单.Tag, 2, 1)
    ElseIf TabShow.Tab = 4 Then
        Chk清单.Tag = Mid(Chk清单.Tag, 1, 1) & Chk清单.Value
    End If
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdAlley_Click()
    '功能：对病人过敏史/病生状态进行管理
    'Pass
    Call AdviceCheckWarn(21)
End Sub

Private Sub Consignment_Click()
    Call UpdateRsByMenu(Consignment, 1)
End Sub

Private Sub ConsignmentALL_Click()
    '全部发药
    Dim Str执行状态 As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    Bill未发药清单.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !执行状态 <> 0 Then
                If Not (!审核人 = "" And Int允许未审核处方发药 = 0) Then
                    !执行状态 = 1
                    Str执行状态 = IIf(!执行状态 = 0, "缺药", IIf(!执行状态 = 1, "发药", IIf(!执行状态 = 2, "拒发", "不处理")))
                    !状态 = Str执行状态
                    .Update
                
                    '如果该记录已填充到表格，则连锁更新
                    With Bill未发药清单
                        If .rows - 1 >= RecChangeData!位置 Then .TextMatrix(RecChangeData!位置, 列名_未发药清单.状态) = Str执行状态
                        
                        lngColor = IIf(Str执行状态 = "发药", glngSendBlkColor, glngOtherBlkColor)
                        
                        .Row = RecChangeData!位置
                        For intCol = 0 To .Cols - 1
                            .Col = intCol
                            .CellBackColor = lngColor
                        Next
                    End With
                End If
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    Bill未发药清单.Redraw = True
    
    '设置菜单及工具按钮的状态
    Call SetMenuAndToolbarState
End Sub

Private Sub Form_Activate()
    Dim dateCurDate As Date
    
    On Error Resume Next
    
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    mblnFirstSended = True
    
    If BlnFirstStart = False Then
'        mnuViewRefresh_Click
    End If
   
    Form_Resize
    BlnFirstStart = True
    TimerAuto.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If TabShow.Tab = 4 And ActiveControl.Name = "TxtInput" Then
        If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then
            If KeyCode = vbKeySpace Then
                Call Bill已发药清单_KeyDown(KeyCode, 0)
            Else
                If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
                    If Bill已发药清单.Row + 1 < Bill已发药清单.rows - 1 Then Bill已发药清单.Row = Bill已发药清单.Row + 1
                Else
                    If Bill已发药清单.Row - 1 > 0 Then Bill已发药清单.Row = Bill已发药清单.Row - 1
                End If
            End If
            Call Bill已发药清单_EnterCell
            Bill已发药清单.SetFocus
        End If
    End If
    
    If Lvw给药途径.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw给药途径)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw给药途径)
            End If
        End If
    End If
    
    If Lvw剂型.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw剂型)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw剂型)
            End If
        End If
    End If
    
    
    If Lvw科室.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw科室)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw科室)
            End If
        End If
        
        If KeyCode = vbKeyEscape Then
            Call Lvw科室_LostFocus
        End If
    End If
    
    err = 0
End Sub

Private Sub SelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.Count
        UserListView.ListItems(n).Checked = True
    Next
End Sub

Private Sub UnSelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.Count
        UserListView.ListItems(n).Checked = False
    Next
End Sub
Private Sub Form_Load()
    Dim dblAdjustWidth As Double
    Dim dblAdjustWidth1 As Double
    
    BlnEnterCell = False
    str排序_未发药 = "NO " & strAsc
    str排序_发退药 = "NO " & strAsc
    
    '初始化变量
    BlnStartUp = False
    BlnFirstStart = False
    Bln刷新未发药清单 = True
    Bln检测库存 = True
    mdblConditonHeight = 3000
    
    If Screen.Width \ Screen.TwipsPerPixelX <= 800 Then
        mbln低分辨率 = True
    End If
    
    '低分辨率时调整部分控件的宽度或位置
    If mbln低分辨率 Then
        dblAdjustWidth = lblInfo.Left - Dtp结束时间.Left - Dtp结束时间.Width - 100
        
        fraConNormal.Width = fraConNormal.Width - dblAdjustWidth
        fraConExpand.Width = fraConNormal.Width
        fraConRequest.Width = fraConNormal.Width
        
        lblInfo.Left = lblInfo.Left - dblAdjustWidth
        lblPatiInputType.Left = lblPatiInputType.Left - dblAdjustWidth
        txtPati.Left = txtPati.Left - dblAdjustWidth
        
        txt科室.Width = txt科室.Width - dblAdjustWidth
        cmd部门类型.Left = cmd部门类型.Left - dblAdjustWidth
        
        
        dblAdjustWidth1 = lbl药品剂型.Left - cmd给药途径.Left - 100
        
        If dblAdjustWidth > dblAdjustWidth1 Then
            txt给药途径.Width = txt给药途径.Width - (dblAdjustWidth - dblAdjustWidth1) / 2 - 150
            txt药品剂型.Width = txt药品剂型.Width - (dblAdjustWidth - dblAdjustWidth1) / 2 - 150
        End If
        
        cmd给药途径.Left = txt给药途径.Left + txt给药途径.Width + 10
        
        lbl药品剂型.Left = cmd给药途径.Left + cmd给药途径.Width + 100
        txt药品剂型.Left = lbl药品剂型.Left + lbl药品剂型.Width + 100
        cmd药品剂型.Left = txt药品剂型.Left + txt药品剂型.Width + 10
        Lbl医嘱类型.Left = lbl药品剂型.Left
        Cbo医嘱类型.Left = txt药品剂型.Left
        opt范围(1).Left = opt范围(1).Left - 150
        opt范围(2).Left = opt范围(2).Left - 150
    End If
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    If gstrUserName = "" Then
        MsgBox "请为当前用户设置对应的操作员后再使用本模块！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    bln药品留存入出类别 = GetDepend
    
    Call GetSysParms
    
    Call GetPrivs
    
    Call TradeName
    
    '为各控件装入图标
    If LoadInIcon = False Then Exit Sub
    '依赖数据检测
    If DependOnCheck = False Then Exit Sub
    
    '初始化条件栏
    Call IniConditon
    
      
    Call LoadCondition(IIf(TabShow.Tab = 4, 1, 0))
    
    '初始化记录集
    Call InitRec
    Call InitRefreshRec
    '设置各控件的样式
    Call SetFormat
    
    Call 权限控制
    
    BlnStartUp = True
    BlnEnterCell = True
    RestoreWinState Me, App.ProductName
    '恢复个性化设置后，有几列始终不能隐藏
    If Bill未发药清单.ColWidth(列名_未发药清单.状态) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.状态) = 700
    If Bill未发药清单.ColWidth(列名_未发药清单.批号) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.批号) = 1500
    If Bill已发药清单.ColWidth(列名_已发药清单.状态) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.状态) = 700
    If Bill已发药清单.ColWidth(列名_已发药清单.退药数) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.退药数) = 1000
    '警示列根据参数来决定是否显示
    Bill未发药清单.ColWidth(列名_未发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
    Bill已发药清单.ColWidth(列名_已发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1342_1")
    
    '取通过本地参数设置的是否显示参数值
    Call SetColMenu
    Call SetColHide(Bill未发药清单)
    Call SetColHide(Bill已发药清单)
    
    '取药房人员
    Call Get配药人
    
    '取配药单格式
    Call Get发药单格式
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    mlngMyWindow = 0
    
    If TabShow.Tab = 0 Then
        SaveSetting "ZLSOFT", "公共模块\操作\" & App.ProductName & "\Frm部门发药管理", "显示退药待发单据", mlng待发单据
    End If
    
'    '保存发药条件
    Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call SaveCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call ClearCondition(IIf(TabShow.Tab = 4, 1, 0))
    
    mintLastTab = 0
    
    '如果是未发药清单或发退药清 单，则保存其设置
    If Bill未发药清单.ColWidth(列名_未发药清单.状态) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.状态) = 700
    If Bill未发药清单.ColWidth(列名_未发药清单.批号) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.批号) = 1500
    If Bill已发药清单.ColWidth(列名_已发药清单.状态) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.状态) = 700
    If Bill已发药清单.ColWidth(列名_已发药清单.退药数) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.退药数) = 1000
    
    Bill未发药清单.Tag = "": Bill已发药清单.Tag = ""
    Bill缺药清单.Tag = "": Bill拒发药清单.Tag = "": Bill汇总发药.Tag = ""
    Call SaveFlexState(Bill未发药清单, "未发药清单")
    Call SaveFlexState(Bill已发药清单, "已发药清单")
    Call SaveFlexState(Bill汇总发药, "汇总发药" & Lng汇总显示)
    Call SaveFlexState(Bill缺药清单, "缺药清单")
    Call SaveFlexState(Bill拒发药清单, "拒发药清单")
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim DblHeight As Double, DblWidth As Double
    Dim dblMaxWidth As Double
    
    On Error Resume Next
    
    dblMaxWidth = IIf(mbln低分辨率, 12240, 13275)
    
    If Me.WindowState = 1 Then Exit Sub
    
    If BlnFirstStart = False Then
        Cbar.Align = 1
        With Cbar
            Set .Bands(1).Child = Tbar
            .Bands(1).MinHeight = Tbar.Height
        End With
    End If
    
    If Me.Height < 8500 Then Me.Height = 8500
    If Me.Width < dblMaxWidth Then Me.Width = dblMaxWidth
    
        
    '调整条件栏
    Call ResizeCondition
    
End Sub

Private Sub HandBack_Click()
    Call UpdateRsByMenu(HandBack, 1)
End Sub

Private Sub HandBackALL_Click()
    Dim Str执行状态 As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    '全部拒发
    Bill未发药清单.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !执行状态 <> 0 Then
                !执行状态 = 2
                Str执行状态 = IIf(!执行状态 = 0, "缺药", IIf(!执行状态 = 1, "发药", IIf(!执行状态 = 2, "拒发", "不处理")))
                !状态 = Str执行状态
                .Update
                
                '如果该记录已填充到表格，则连锁更新
                With Bill未发药清单
                    If .rows - 1 >= RecChangeData!位置 Then .TextMatrix(RecChangeData!位置, 列名_未发药清单.状态) = Str执行状态
                    
                    lngColor = IIf(Str执行状态 = "发药", glngSendBlkColor, glngOtherBlkColor)
                    
                    .Row = RecChangeData!位置
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellBackColor = lngColor
                    Next
                End With
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Bill未发药清单.Redraw = True
    
    '设置菜单及工具按钮的状态
    Call SetMenuAndToolbarState
    
End Sub

Private Sub Lack_Click()
    Call UpdateRsByMenu(Lack, 1)
End Sub

Private Sub mnuEditHandbackBatch_Click()
    TimerAuto.Enabled = False
    If Not frm批量退药.ShowEditor(Me, lng药房ID, False, int金额保留位数) Then Exit Sub
    mnuViewRefresh_Click
    
    DoEvents
    TimerAuto.Enabled = True
End Sub
Private Sub mnuFilePrintTotal_Click()
    Dim str药房 As String, str科室 As String
    Dim rsTmp As New adodb.Recordset
    Dim str显示 As String
    Dim n As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码,名称 From 部门表 Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取当前药房的名称]", lng药房ID)
    
    If Not rsTmp.RecordCount <= 0 Then str药房 = "(" & rsTmp!编码 & ")" & rsTmp!名称
    
    str显示 = ""
    If InStr(mstr部门, ",") > 0 Then
        gstrSQL = "Select ID,名称 From 部门表 Where ID In(" & mstr部门 & ") Order by 编码"
        Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "读取科室名称")
    Else
        gstrSQL = "Select ID,名称 From 部门表 Where ID = [1] Order by 编码"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "读取科室名称", mstr部门)
    End If
    
    If Not rsTmp.RecordCount <= 0 Then
        For n = 1 To rsTmp.RecordCount
            str显示 = str显示 & "," & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If
    
    str显示 = Mid(str显示, 2)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
        "发药库房=" & str药房 & "|" & lng药房ID, _
        "部门性质=" & IIf(mint类型 = 0, "临床科室", IIf(mint类型 = 1, "医技科室", "病人病区")) & "|" & mint类型, _
        "领药部门=" & str显示 & "|" & " IN (" & mstr部门 & ")", "包装系数=" & IIf(strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), "ReportFormat=" & IIf(cbo发药单格式.ListIndex = -1, 1, cbo发药单格式.ListIndex + 1))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileRestore_Click()
'功能：打印退药通知单
    Dim StrDate As String
    
    If Trim(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.药品ID)) = "" Then Exit Sub
    With RecChangeSendedData
        If .RecordCount <> 0 Then
            .MoveFirst
            .Find "位置=" & Bill已发药清单.Row
        End If
        If .EOF Then Exit Sub
        If !可操作 <> 3 Then Exit Sub
        StrDate = Format(!发药时间, "yyyy-MM-dd HH:mm:ss")
    End With
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
End Sub

Private Sub mnuFileWait_Click()
    Dim rsTmp As New adodb.Recordset
    Dim str显示 As String, str绑定 As String
    Dim str药房 As String, i As Long
    Dim n As Integer
    
    On Error GoTo errHandle
    If glngSys \ 100 = 1 Then
        '库房条件
        gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取当前药房的名称]", lng药房ID)
        
        str药房 = rsTmp!名称 & "|" & lng药房ID
            
        str显示 = ""
        If InStr(mstr部门, ",") > 0 Then
            gstrSQL = "Select ID,名称 From 部门表 Where ID In(" & mstr部门 & ") Order by 编码"
            Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "读取科室名称")
        Else
            gstrSQL = "Select ID,名称 From 部门表 Where ID = [1] Order by 编码"
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "读取科室名称", mstr部门)
        End If
        If Not rsTmp.RecordCount <= 0 Then
            For n = 1 To rsTmp.RecordCount
                str显示 = str显示 & "," & rsTmp!名称
                rsTmp.MoveNext
            Next
        End If
        str显示 = Mid(str显示, 2)
        str绑定 = mstr部门
    
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1342_1", Me, _
            "住院药局=" & str药房, "住院科室=" & str显示 & "|" & " IN (" & str绑定 & ")", _
            "开始时间=" & mstr开始日期_未发, "结束时间=" & mstr结束日期_未发, 1)

    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFlag_Click()
    Dim frmFlag As New Frm不再发药处方标志
    
    TimerAuto.Enabled = False
    BlnRefresh = False
    
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    
    If BlnRefresh Then
        Call mnuViewRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
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
    '默认参数：药品=药品id，药房=药房id，病人ID=病人id，住院号=住院号，NO=处方NO，单据类型=药品收发记录.单据
    Dim lng药品ID As Long
    
    If TabShow.Tab = 0 Then
        If Bill未发药清单.Row > 0 Then
            If Val(Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.药品ID)) > 0 Then
                lng药品ID = Val(Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.药品ID))
            End If
        End If
    ElseIf TabShow.Tab = 4 Then
        If Bill已发药清单.Row > 0 Then
            If Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.药品ID)) > 0 Then
                lng药品ID = Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.药品ID))
            End If
        End If
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "药品=" & IIf(lng药品ID = 0, "", lng药品ID), _
        "药房=" & IIf(lng药房ID = 0, "", lng药房ID), _
        "病人ID=" & IIf(mlng病人ID = 0, "", mlng病人ID), _
        "住院号=" & mstr住院号, _
        "NO=" & mstrNo, _
        "单据类型=" & IIf(mInt单据 = 0, "", mInt单据))
        
End Sub

Private Sub mnuReVerify_Click()
    TimerAuto.Enabled = False
    BlnRefresh = False
    
    Frm药品销账.ShowForm Me, lng药房ID, strUnit, int药品名称, int金额保留位数
    
    If BlnRefresh Then
        Call mnuViewRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub mnuTypeItem_Click(Index As Integer)
    Dim n As Integer
    Dim strType As String
    
    With mnuTypeItem
        .Item(Index).Checked = Not .Item(Index).Checked
        For n = 0 To .Count - 1
            If .Item(n).Checked = True Then
                strType = strType & ";" & .Item(n).Caption & ";"
            End If
        Next
    End With
    
    With Lvw给药途径
        For n = 1 To .ListItems.Count
            If InStr(1, strType, ";" & .ListItems(n).Tag & ";") > 0 Then
                .ListItems(n).Checked = True
            Else
                .ListItems(n).Checked = False
            End If
        Next
    End With
End Sub

Private Sub mnuViewFontSet_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSet(i).Checked = False
    Next
    Me.mnuViewFontSet(Index).Checked = True
    
    Me.Bill未发药清单.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill缺药清单.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill已发药清单.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill拒发药清单.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    Me.Bill汇总发药.Font.Size = IIf(Index = 0, 9, IIf(Index = 1, 11, 15))
    
    zldatabase.SetPara "字体", Index, glngSys, 1342
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub MnuViewLocate_Click()
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    TimerAuto.Enabled = False
    strFind = Frm部门发药定位.ShowME(lng药房ID, Me, mstrPrivs)
    If strFind = "" Then
        TimerAuto.Enabled = True
        Exit Sub
    End If
    
    '初始化记录集
    Set Rec未发 = New adodb.Recordset
    Set Rec已发 = New adodb.Recordset
    With Rec未发
        If .State = 1 Then .Close
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    With Rec已发
        If .State = 1 Then .Close
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '处理未发药品记录
    With RecChangeData
        If .RecordCount <> 0 Then
            .Filter = strFind
            Do While Not .EOF
                Rec未发.AddNew
                Rec未发!位置 = !位置
                Rec未发!科室 = !科室
                Rec未发!类型 = !类型
                Rec未发!NO = !NO
                Rec未发!床号 = !床号
                Rec未发!姓名 = !姓名
                Rec未发!药品ID = !药品ID
                Rec未发.Update
                .MoveNext
            Loop
            .Filter = 0
        End If
    End With
    '处理已发药品记录
    With RecChangeSendedData
        If .RecordCount <> 0 Then
            .Filter = strFind
            Do While Not .EOF
                Rec已发.AddNew
                Rec已发!位置 = !位置
                Rec已发!科室 = !科室
                Rec已发!类型 = !类型
                Rec已发!NO = !NO
                Rec已发!床号 = !床号
                Rec已发!姓名 = !姓名
                Rec已发!药品ID = !药品ID
                Rec已发.Update
                .MoveNext
            Loop
            .Filter = 0
        End If
    End With
    
    Call FindRecord
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub MnuViewLocateNext_Click()
    Call FindRecord(False)
End Sub

Private Sub MnuViewNone_Click()
    Call UpdateState(False, False)
End Sub

Private Sub MnuViewTotal_Click()
    Call UpdateState(False, True)
End Sub

Private Sub Nop_1_Click()
    Call UpdateRsByMenu(Nop_1, 1)
End Sub

Private Sub Nop_2_Click()
    Call UpdateRsByMenu(Nop_2, 2)
End Sub

Private Sub Nop_3_Click()
    Call UpdateRsByMenu(Nop_3, 3)
End Sub

Private Sub Nop_ALL_Click()
    Dim Str执行状态 As String
    Dim intCol As Integer
    Dim lngColor As Long
    
    '全部不处理
    Bill未发药清单.Redraw = False
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !执行状态 <> 0 Then
                !执行状态 = 3
                Str执行状态 = IIf(!执行状态 = 0, "缺药", IIf(!执行状态 = 1, "发药", IIf(!执行状态 = 2, "拒发", "不处理")))
                !状态 = Str执行状态
                .Update
                
                '如果该记录已填充到表格，则连锁更新
                With Bill未发药清单
                    If .rows - 1 >= RecChangeData!位置 Then .TextMatrix(RecChangeData!位置, 列名_未发药清单.状态) = Str执行状态
                    
                    lngColor = IIf(Str执行状态 = "发药", glngSendBlkColor, glngOtherBlkColor)

                    .Row = RecChangeData!位置
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellBackColor = lngColor
                    Next
                End With
            End If
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Bill未发药清单.Redraw = False
    
    '设置菜单及工具按钮的状态
    Call SetMenuAndToolbarState
    
End Sub

Private Sub Restore_Click()
    Call UpdateRsByMenu(Restore, 2)
End Sub

Private Sub ResumeDo_Click()
    Call UpdateRsByMenu(ResumeDo, 3)
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreView_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Consignment"
        MnuEditVerify_Click
    Case "Desire"
        MnuEditDesire_Click
    Case "Handback"
        MnuEditHandback_Click
    Case "Restore"
        MnuEditRestore_Click
    Case "ReVerify"
        mnuReVerify_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnufileexit_Click
    End Select
End Sub

Private Sub Cbar_Resize()
    Form_Resize
End Sub

Private Sub MnuEditDesire_Click()
    '
End Sub

Private Sub MnuEditHandback_Click()
    Dim IntSet As Integer
    On Error GoTo ErrHand
    '按药品ID顺序更新
    
    '先恢复拒发药处方记录为正常记录
    gcnOracle.BeginTrans
    With Bill拒发药清单
        For IntSet = 1 To .rows - 1
            If Trim(.TextMatrix(IntSet, 1)) = "恢复" Then
                gstrSQL = "zl_药品收发记录_部门恢复(" & .RowData(IntSet) & ")"
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-恢复拒发药品")
            End If
        Next
    End With
    
    '根据用户设置当前正常记录为拒发药
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
            .Sort = "药品ID Asc"
        End If
        Do While Not .EOF
            If !执行状态 = 2 Then
                If CheckBill(0, !Id) <> 0 Then gcnOracle.RollbackTrans: Exit Sub
                gstrSQL = "zl_药品收发记录_部门拒发(" & !Id & ")"
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-设置拒发药品")
            End If
            .MoveNext
        Loop
    End With
    
    '刷新
    gcnOracle.CommitTrans
    Set RecRefreshCompare = CopyNewRec(RecChangeData)
    mnuViewRefresh_Click
    Call InitRefreshRec
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    If RecChangeData.RecordCount <> 0 Then RecChangeData.Sort = "NO Asc"
End Sub

Private Sub MnuEditRestore_Click()
    Dim StrDate As String
    Dim lng分批 As Long, lng批次 As Long, lngRow As Long
    Dim strShow As String, strReturn As String, blnInput As Boolean, strSubSql As String
    Dim sig退药数 As Single
    Dim RecRecord As New adodb.Recordset
    Dim rsTemp As New adodb.Recordset
    Dim bln是否有退药 As Boolean
    Dim str药品id As String
     
    On Error GoTo ErrHand
    
    If TxtInput.Visible Then
        Call TxtInput_LostFocus
        TxtInput.Visible = False
    End If
    
    '按药品ID顺序更新
    StrDate = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Sub
        If .EOF Then Exit Sub
        
        Call BuildRecord(False)
        If Not CheckCorrelation Then Exit Sub
        
        If MsgBox("你确定要退药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '退药人签名
        str退药人 = ""
        If Lng退药人签名 = 1 Then
            str退药人 = zldatabase.UserIdentify(Me, "退药人签名", glngSys, 1342, "退药")
            If str退药人 = "" Then
                Exit Sub
            End If
        End If
        
        .Sort = "药品ID Asc"
        
        Do While Not .EOF
            If !执行状态 = 3 Then
                '先检查是否允许退药（医嘱）
                If bln医嘱作废 = False Then
                    gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1]"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", CLng(!Id))
                    
                    If (rsTemp!扣率 Like "1*") Then       '临嘱
                        gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 病人费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", CLng(!Id))
                        
                        If Not rsTemp.EOF Then
                            If (rsTemp!门诊标志 = 1 Or rsTemp!门诊标志 = 4) And rsTemp!医嘱序号 <> 0 Then
                                gstrSQL = "Select decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where ID=[1]"
                                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rsTemp!医嘱序号))
                                
                                If rsTemp!作废 = 0 Then
                                    MsgBox "第" & !位置 & "行药品对应的医嘱还未作废，不能退药！", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                
                lngRow = !位置
                lng分批 = IIf(IsNull(!分批), 0, !分批)
                lng批次 = IIf(IsNull(!批次), 0, !批次)
                '如果原来不分批而现在分批
                If lng批次 = 0 And lng分批 = 1 Then
                    '如果批号或效期为空，则提取供用户输入
                    blnInput = IIf(IsNull(!批号), True, False)
                    If Not blnInput Then blnInput = (Trim(!批号) = "")
                    If blnInput Then
                        strShow = Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.科室) & "|" & Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.床号) & _
                        "|" & Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.姓名) & "|" & Bill已发药清单.TextMatrix(lngRow, 列名_已发药清单.药品名称) & "|" & !药品ID
                        strReturn = Frm退药设置.ShowME(Me, strShow)
                        If strReturn = "" Then Exit Sub
                        '更新批号、效期及产地
                        !批号 = Split(strReturn, "|")(0)
                        !效期 = Split(strReturn, "|")(1)
                        !产地 = Split(strReturn, "|")(2)
                        .Update
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        gcnOracle.BeginTrans
        Do While Not .EOF
            If !执行状态 = 3 Then
                If CheckBill(2, !Id) <> 0 Then gcnOracle.RollbackTrans: Exit Sub
                
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
                    
                sig退药数 = !退药数
                
                gstrSQL = " Select round(" & sig退药数 & strSubSql & ",5) 数量 From 药品规格" & _
                         " Where 药品ID=[1]"
                Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(RecChangeSendedData!药品ID))
                
                With RecRecord
                    sig退药数 = !数量
                End With
                If Val(!准退数) = Val(!退药数) Then
                    sig退药数 = Val(!实际数量)
                End If
                
                If sig退药数 <> 0 Then
                    If CheckPrice(!Id, mstr价格失效提示) = False Then
                        If MsgBox("药品[" & !品名 & "(" & !规格 & ")]" & mstr价格失效提示, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            gstrSQL = "zl_药品收发记录_部门退药(" & !Id & ",'" & gstrUserName & "',To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                            IIf(IsNull(!批号), "NULL", IIf(Mid(!批号, 1, 1) = "(", "NULL", "'" & Mid(!批号, 1, 8) & "'")) & "," & _
                            IIf(IsNull(!效期), "NULL", IIf(!效期 = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                            IIf(IsNull(!产地), "NULL", "'" & !产地 & "'") & "," & sig退药数 & ",NULL,'" & str退药人 & "'," & int金额保留位数 & ")"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药")
                            bln是否有退药 = True
                            
                            If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                                str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                            End If
                        End If
                    Else
                        gstrSQL = "zl_药品收发记录_部门退药(" & !Id & ",'" & gstrUserName & "',To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                        IIf(IsNull(!批号), "NULL", IIf(Mid(!批号, 1, 1) = "(", "NULL", "'" & Mid(!批号, 1, 8) & "'")) & "," & _
                        IIf(IsNull(!效期), "NULL", IIf(!效期 = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                        IIf(IsNull(!产地), "NULL", "'" & !产地 & "'") & "," & sig退药数 & ",NULL,'" & str退药人 & "'," & int金额保留位数 & ")"
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药")
                        bln是否有退药 = True
                        
                        If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                            str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                        End If
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    gcnOracle.CommitTrans
    
    '打印退药单
    If bln是否有退药 = True Then
        If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
        End If
        
        '提示停用药品
        If str药品id <> "" Then
            Call CheckStopMedi(str药品id)
        End If
    Else
        MsgBox "本次没有退药。"
        Exit Sub
    End If
    
    '刷新
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.Sort = "NO Asc"
End Sub

Private Sub MnuEditVerify_Click()
    Dim StrCurDate As String
    Dim LngLocate As Long
    Dim strRecipeKey As String              '保存本次发药处方的ID
    Dim blnUpdate As Boolean
    Dim str显示 As String
    Dim n As Integer
    Dim rsTmp As New adodb.Recordset
    Dim str期间 As String
    Dim lngPatId As Long
    Dim blnBeginTrans As Boolean
    Dim strId批次 As String
    Dim strID As String
    Dim strDept As String
    Dim lngPre费用id As Long
    Dim strPreNo As String
    Dim lngPre费用序号 As Long
    Dim dblSum As Double
    Dim str药品id As String
    Dim dbl留存数量 As Double
    Dim dblPrice As Double
    Dim strSubSql As String
    Dim RecRecord As adodb.Recordset
        
    On Error GoTo ErrHand
    
    If txt留存数.Visible Then
        txt留存数_LostFocus
        txt留存数.Visible = False
    End If
    
    mlng汇总发药号 = Val(zldatabase.GetNextNo(20))
    
    '按病人ID分批更新
    StrCurDate = Format(zldatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    str期间 = Format(StrCurDate, "yyyy")
    
    '检查存储库房
    If CheckDrugStock = False Then Exit Sub
    
    '以下处理发药------------------------------------------------------------------------------------------------------
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Sub
        If .EOF Then Exit Sub

        Call BuildRecord(True)
        If Not CheckCorrelation Then Exit Sub
        
        If MsgBox("你确定要发药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '领药人签名
        TimerAuto.Enabled = False
        str领药人 = ""
        If Lng领药人签名 = 1 Then
            str领药人 = zldatabase.UserIdentify(Me, "领药人签名", glngSys, 1342, "")
            If str领药人 = "" Then
                TimerAuto.Enabled = True
                Exit Sub
            End If
        End If
        TimerAuto.Enabled = True
        
        '必须按病人ID，药品ID排序
        .Sort = "病人ID Asc ,药品ID Asc"
        
        Do While Not .EOF
            '执行状态为1并且通过单据检查才可以发送
            If !执行状态 = 1 And CheckBill(1, !Id) = 0 And CheckGroupSend(!相关ID) = True Then
                If lngPatId = 0 Then
                    lngPatId = !病人ID
                End If
                
                '病人ID相同时候
                If lngPatId = !病人ID Then
                    '如果传入的字符串大于3950时就提交事务（最大字符串为4000）
                    If LenB(strId批次) > 3950 Then
                        gcnOracle.BeginTrans
                        blnBeginTrans = True
                        
                        gstrSQL = "Zl_药品收发记录_批量发药('" & strId批次 & "'," & lng药房ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str领药人 & "'," & mlng汇总发药号 & "," & int金额保留位数 & ",'" & NeedName(cbo配药人.Text) & "') "
                        gcnOracle.BeginTrans
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-部门批量发药")
                                                
                        If int审核划价单 = 1 Then
                            gstrSQL = "Zl_住院记帐记录_发药审核('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-住院记帐审核")
                        End If
                        gcnOracle.CommitTrans
                                                                    
                        blnBeginTrans = False
                        blnUpdate = True
                        lngPatId = 0
                        strId批次 = !Id & "," & NVL(!批次, 0)
                        strID = !Id
                    Else
                        strId批次 = IIf(strId批次 = "", !Id & "," & NVL(!批次, 0), strId批次 & "|" & !Id & "," & NVL(!批次, 0))
                        strID = IIf(strID = "", !Id, strID & "," & !Id)
                    End If
                Else
                    '如果病人ID不同则提交事务
                    blnBeginTrans = True
                    
                    gstrSQL = "Zl_药品收发记录_批量发药('" & strId批次 & "'," & lng药房ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str领药人 & "'," & mlng汇总发药号 & "," & int金额保留位数 & ",'" & NeedName(cbo配药人.Text) & "') "
                    gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-部门批量发药")
                                        
                    If int审核划价单 = 1 Then
                        gstrSQL = "Zl_住院记帐记录_发药审核('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-住院记帐审核")
                    End If
                    gcnOracle.CommitTrans
                    
                    blnBeginTrans = False
                    blnUpdate = True
                    lngPatId = !病人ID
                    strId批次 = !Id & "," & NVL(!批次, 0)
                    strID = !Id
                End If
            End If
            .MoveNext
            
            '如果后面没有记录并且传入字符串不为空，则提交事务
            If .EOF And strId批次 <> "" Then
                blnBeginTrans = True
                
                gstrSQL = "Zl_药品收发记录_批量发药('" & strId批次 & "'," & lng药房ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ,3,'" & str领药人 & "'," & mlng汇总发药号 & "," & int金额保留位数 & ",'" & NeedName(cbo配药人.Text) & "') "
                gcnOracle.BeginTrans
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-部门批量发药")
                                    
                If int审核划价单 = 1 Then
                    gstrSQL = "Zl_住院记帐记录_发药审核('" & strID & "','" & gstrUserCode & "','" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss'))"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-住院记帐审核")
                End If
                gcnOracle.CommitTrans
                blnUpdate = True
                
                blnBeginTrans = False
            End If
        Loop
    End With
    '以上处理发药----------------------------------------------------------------------------------------------------------
    
    
    '以下处理药品留存-------------------------------------------------------------------------------------------------------
    gcnOracle.BeginTrans
    '前提条件是按科室汇总
    If TabShow.Tab = 1 And Lng汇总显示 = 1 Then
        For n = 1 To Bill汇总发药.rows - 3
            If Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.留存数量)) <> 0 Then
                If Not bln药品留存入出类别 Then
                    gcnOracle.RollbackTrans
                    MsgBox "没有设置药品留存的入出类别，请检查药品入出分类！本次按全实发处理。", vbInformation + vbOKOnly, gstrSysName
                    If blnUpdate Then GoTo RefData
                    Exit Sub
                End If
                
                dbl留存数量 = Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.留存数量))
                dblPrice = Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.单价))
                
                Select Case strUnit
                Case "售价单位"
                    strSubSql = "round(" & dbl留存数量 & ",5) As 数量, round(" & dblPrice & ",5) As 单价"
                Case "门诊单位"
                    strSubSql = "round(" & dbl留存数量 & " * Decode(门诊包装,Null,1,0,1,门诊包装) ,5) As 数量, round(" & dblPrice & " /Decode(门诊包装,Null,1,0,1,门诊包装) ,5) As 单价 "
                Case "住院单位"
                    strSubSql = "round(" & dbl留存数量 & " * Decode(住院包装,Null,1,0,1,住院包装) ,5) As 数量, round(" & dblPrice & " /Decode(住院包装,Null,1,0,1,住院包装) ,5) As 单价 "
                Case "药库单位"
                    strSubSql = "round(" & dbl留存数量 & " * Decode(药库包装,Null,1,0,1,药库包装) ,5) As 数量, round(" & dblPrice & " /Decode(药库包装,Null,1,0,1,药库包装) ,5) As 单价 "
                End Select
                    
                gstrSQL = " Select " & strSubSql & " From 药品规格" & _
                         " Where 药品ID=[1]"
                Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.药品ID)))
                dbl留存数量 = RecRecord!数量
                dblPrice = RecRecord!单价
                
                gstrSQL = "ZL_药品留存记录_INSERT(" & str期间 & "," & mlng汇总发药号 & "," & lng药房ID & "," & Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.科室ID)) & "," & Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.药品ID)) & "," & Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.批次)) & ", " & _
                " " & dbl留存数量 & "," & dblPrice & ", '" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss') ," & Val(Bill汇总发药.TextMatrix(n, 列名_科室汇总清单.领药部门id)) & ") "
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-保存留存")
            End If
        Next
    End If
    gcnOracle.CommitTrans
    '以上处理药品留存-------------------------------------------------------------------------------------------------------
    
    
    '以下处理销帐数据------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int审核标志 As Integer
    Dim bln是否有退药 As Boolean
    Dim str序号数量 As String
    
    '前提条件是汇总销帐记录一并发药
    If TabShow.Tab = 1 And mbln汇总发药 = True Then
        If mrsRequest.State <> 0 Then
            mrsRequest.Filter = ""
            mrsRequest.Sort = "No,费用id,收发id"
            If mrsRequest.RecordCount > 0 Then
                With mrsRequest
                    gcnOracle.BeginTrans
                    blnBeginTrans = True
                    gclsInsure.InitOracle gcnOracle
                    Do While Not .EOF
                        If !审核标志 <> 0 Then
                            If lngPre费用id <> !费用id Then
                                '费用销帐记录处理
                                gstrSQL = "zl_病人费用销帐_Audit(" & !费用id & ",To_Date('" & !申请时间 & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                                               gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')," & !审核标志 & ")"
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新病人费用销帐记录")
                                lngPre费用id = !费用id
                            End If
                        End If
                        
                        '退药处理
                        If !审核标志 = 1 And !销帐数量 <> 0 Then
                            gstrSQL = "zl_药品收发记录_部门退药(" & !收发ID & ",'" & gstrUserName & "',To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')," & _
                                IIf(IsNull(!批号), "NULL", IIf(Mid(!批号, 1, 1) = "(", "NULL", "'" & Mid(!批号, 1, 8) & "'")) & "," & _
                                IIf(IsNull(!效期), "NULL", IIf(!效期 = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & "," & _
                                IIf(IsNull(!产地), "NULL", "'" & !产地 & "'") & "," & !销帐数量 & ",NULL,'" & gstrUserName & "'," & int金额保留位数 & "," & mlng汇总发药号 & ")"
            
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-药品退药销帐")
                            bln是否有退药 = True
                            
                            If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                                str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                            End If
                        
                            '销帐处理
                            strPreNo = !NO
                            lngPre费用序号 = !费用序号
                            dblSum = dblSum + !销帐数量
                            
                            .MoveNext
                            If .EOF Then
                                .MovePrevious
                                str序号数量 = !费用序号 & ":" & dblSum
                
                                gstrSQL = "ZL_住院记帐记录_Delete('" & !NO & "','" & str序号数量 & "','" & gstrUserCode & "','" & gstrUserName & "'," & !记录性质 & ")"
                                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-删除记帐记录")
                
                                '医保处理
                                If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                                    MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                                    MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                                            "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                                End If
                                .MoveNext
                            Else
                                If strPreNo <> !NO Or (strPreNo = !NO And lngPre费用序号 <> !费用序号) Then
                                    .MovePrevious
                                    str序号数量 = !费用序号 & ":" & dblSum
                                    
                                    gstrSQL = "ZL_住院记帐记录_Delete('" & !NO & "','" & str序号数量 & "','" & gstrUserCode & "','" & gstrUserName & "'," & !记录性质 & ")"
                                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-删除记帐记录")
                    
                                    '医保处理
                                    If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                                        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                                        MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                                                "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                                    End If
                                    
                                    dblSum = 0
                                    .MoveNext
                                End If
                            End If
                            .MovePrevious
                        End If
                        .MoveNext
                    Loop
                End With
            
                '医保，记帐作废上传，作废时上传
                If strMCNO <> "" Then
                    arrMCRec = Split(strMCNO, "|")
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                gcnOracle.RollbackTrans
                                GoTo RefData
                            End If
                        End If
                    Next
                End If
                                        
                gcnOracle.CommitTrans
                blnBeginTrans = False
                
                '医保，记帐作废上传，完成后上传
                If strMCNO <> "" Then
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                            End If
                        End If
                    Next
                End If
                
                If bln是否有退药 = True Then
                    If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrCurDate, "包装系数=" & IIf(strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), 2)
                    End If
                End If
            End If
        End If
    End If
    
    '提示停用药品
    If str药品id <> "" Then
        Call CheckStopMedi(str药品id)
    End If
    
    '以上处理销帐数据------------------------------------------------------------------------------------------------------
    
    blnBeginTrans = False
RefData:
    '刷新
    If blnUpdate Then
        strDept = mstrDrawDept
        Set RecRefreshCompare = CopyNewRec(RecChangeData)
        mnuViewRefresh_Click
        Call InitRefreshRec
        
        '打印汇总单据
        If Lng自动打印 = 1 Then
            str显示 = ""
            
            If InStr(strDept, ",") > 0 Then
                gstrSQL = "Select ID,名称 From 部门表 Where ID In(" & strDept & ") Order by 编码"
                Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "读取科室名称")
            Else
                gstrSQL = "Select ID,名称 From 部门表 Where ID = [1] Order by 编码"
                Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "读取科室名称", strDept)
            End If
            
            If Not rsTmp.RecordCount <= 0 Then
                For n = 1 To rsTmp.RecordCount
                    str显示 = str显示 & "," & rsTmp!名称
                    rsTmp.MoveNext
                Next
            End If
            
            str显示 = Mid(str显示, 2)
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & lng药房ID, _
                "部门性质=" & mint类型, _
                "领药部门=" & str显示 & "|" & " IN (" & strDept & ")", _
                "包装系数=" & IIf(strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), _
                "发药号=" & mlng汇总发药号, "ReportFormat=" & IIf(cbo发药单格式.ListIndex = -1, 1, cbo发药单格式.ListIndex + 1), "PrintEmpty=0", 2)

        End If
    End If
    
    blnUpdate = False
        
    Exit Sub
ErrHand:
    '如果已开启事务，并且未提交，则出错时回滚事务
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    If RecChangeData.RecordCount <> 0 Then RecChangeData.Sort = "NO Asc"
    If blnUpdate Then GoTo RefData
End Sub



Private Sub MnuFileBillprint_Click()
    '
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub MnuFilePara_Click()
    Dim intFixedCol As Integer
    Dim dateCurDate As Date
    
    BlnSetPara = False
    TimerAuto.Enabled = False
    With Frm部门发药参数设置
        .strPrivs = mstrPrivs
        .Show 1, Me
    End With
    
    '从注册表中读取相关参数设置
    If BlnSetPara Then
        '重新获取注册表
        Call ReadFromReg
        
        Call Get配药人
        Call Get剂型

        '重新设置汇总清单的格式
        With Bill汇总发药
            .rows = 2
            .Cols = IIf(Lng汇总显示 = 1, 列名_科室汇总清单.列数, 列名_汇总清单.列数)
        
            If Lng汇总显示 = 0 Then
                .TextMatrix(0, 列名_汇总清单.药品名称) = "药品名称"
                .TextMatrix(0, 列名_汇总清单.规格) = "规格"
                .TextMatrix(0, 列名_汇总清单.产地) = "产地"
                .TextMatrix(0, 列名_汇总清单.批号) = "批号"
                .TextMatrix(0, 列名_汇总清单.数量) = "数量"
                .TextMatrix(0, 列名_汇总清单.单位) = "单位"
                .TextMatrix(0, 列名_汇总清单.单价) = "单价"
                .TextMatrix(0, 列名_汇总清单.金额) = "金额"
                            
                .ColWidth(列名_汇总清单.药品名称) = 2000
                .ColWidth(列名_汇总清单.规格) = 1500
                .ColWidth(列名_汇总清单.产地) = 1500
                .ColWidth(列名_汇总清单.批号) = 1200
                .ColWidth(列名_汇总清单.数量) = 1200
                .ColWidth(列名_汇总清单.单位) = 500
                .ColWidth(列名_汇总清单.单价) = 1200
                .ColWidth(列名_汇总清单.金额) = 1200
            Else
                .TextMatrix(0, 列名_科室汇总清单.科室) = "科室"
                .TextMatrix(0, 列名_科室汇总清单.药品名称) = "药品名称"
                .TextMatrix(0, 列名_科室汇总清单.规格) = "规格"
                .TextMatrix(0, 列名_科室汇总清单.产地) = "产地"
                .TextMatrix(0, 列名_科室汇总清单.批号) = "批号"
                .TextMatrix(0, 列名_科室汇总清单.应发数量) = "应发数量"
                .TextMatrix(0, 列名_科室汇总清单.留存数量) = "留存数量"
                .TextMatrix(0, 列名_科室汇总清单.销帐数量) = "销帐数量"
                .TextMatrix(0, 列名_科室汇总清单.实发数量) = "实发数量"
                .TextMatrix(0, 列名_科室汇总清单.单位) = "单位"
                .TextMatrix(0, 列名_科室汇总清单.单价) = "单价"
                .TextMatrix(0, 列名_科室汇总清单.金额) = "金额"
                .TextMatrix(0, 列名_科室汇总清单.批次) = "批次"
                .TextMatrix(0, 列名_科室汇总清单.科室ID) = "科室ID"
                .TextMatrix(0, 列名_科室汇总清单.药品ID) = "药品ID"
                
                .ColWidth(列名_科室汇总清单.科室) = 1200
                .ColWidth(列名_科室汇总清单.药品名称) = 2000
                .ColWidth(列名_科室汇总清单.规格) = 1500
                .ColWidth(列名_科室汇总清单.产地) = 1500
                .ColWidth(列名_科室汇总清单.批号) = 1200
                .ColWidth(列名_科室汇总清单.应发数量) = 1200
                .ColWidth(列名_科室汇总清单.留存数量) = 1200
                .ColWidth(列名_科室汇总清单.销帐数量) = IIf(mbln汇总发药 = True, 1200, 0)
                .ColWidth(列名_科室汇总清单.实发数量) = 1200
                .ColWidth(列名_科室汇总清单.单位) = 500
                .ColWidth(列名_科室汇总清单.单价) = 1200
                .ColWidth(列名_科室汇总清单.金额) = 1200
                .ColWidth(列名_科室汇总清单.批次) = 0
                .ColWidth(列名_科室汇总清单.科室ID) = 0
                .ColWidth(列名_科室汇总清单.药品ID) = 0
            End If
        
            For intFixedCol = 0 To .Cols - 1
                .ColAlignmentFixed(intFixedCol) = 4
            Next
            .ColAlignment(IIf(Lng汇总显示 = 1, 列名_科室汇总清单.规格, 列名_汇总清单.规格)) = 1
            .ColAlignment(IIf(Lng汇总显示 = 1, 列名_科室汇总清单.批号, 列名_汇总清单.批号)) = 1
        End With
        Call RestoreFlexState(Bill汇总发药, "汇总发药" & Lng汇总显示)
        Bill汇总发药.ColWidth(列名_科室汇总清单.销帐数量) = IIf(mbln汇总发药 = True, 1200, 0)
        
        '刷新数据
        Call mnuViewRefresh_Click
    End If
    
    Bill未发药清单.ColWidth(列名_未发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
    Bill已发药清单.ColWidth(列名_已发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
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
        Tbar.Buttons("Preview").Caption = "预览"
        Tbar.Buttons("Print").Caption = "打印"
        Tbar.Buttons("Consignment").Caption = "发药"
        Tbar.Buttons("Desire").Caption = "申领"
        Tbar.Buttons("Handback").Caption = "拒发"
        Tbar.Buttons("Restore").Caption = "退药"
        Tbar.Buttons("ReVerify").Caption = "销帐"
        Tbar.Buttons("Help").Caption = "帮助"
        Tbar.Buttons("Exit").Caption = "退出"
    Else
        Tbar.Buttons("Preview").Caption = ""
        Tbar.Buttons("Print").Caption = ""
        Tbar.Buttons("Consignment").Caption = ""
        Tbar.Buttons("Desire").Caption = ""
        Tbar.Buttons("Handback").Caption = ""
        Tbar.Buttons("Restore").Caption = ""
        Tbar.Buttons("ReVerify").Caption = ""
        Tbar.Buttons("Help").Caption = ""
        Tbar.Buttons("Exit").Caption = ""
    End If
    
    Cbar.Bands(1).MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub MnuHelpWebM_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    If BlnStartUp = False Then Exit Sub
    If BlnInRefresh Then Exit Sub
    
    BlnInRefresh = True
    Bln刷新未发药清单 = True
    Bln检测库存 = True
    
    '避免退药数量写入刷新后的列表中
    If TxtInput.Visible Then
        TxtInput.Visible = False
        Call TxtInput_LostFocus
    End If
    
    '数据刷新后，必须重新设置查找条件
    strFind = ""
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    
    Call AviShow
    ''''设置条件
    Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    Call RefreshData
    Call AviShow(False)
    
    mdate上次刷新时间 = zldatabase.Currentdate
    
    BlnInRefresh = False
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    On Error Resume Next
    
    If TabShow.Tab = 4 Then
        mnuBillItem(2).Visible = False
        mnuBillItem(6).Visible = False
        mnuBillItem(23).Visible = False
        mnuBillItem(24).Visible = False
    ElseIf TabShow.Tab = 0 Then
        mnuBillItem(2).Visible = UserPrivDetail.Priv_医生查询
        mnuBillItem(6).Visible = True
        mnuBillItem(23).Visible = True
        mnuBillItem(24).Visible = True
    End If
    If TabShow.Tab = 0 Or TabShow.Tab = 4 Then
        cmdAlley.Visible = mblnStarPass
    Else
        cmdAlley.Visible = False
    End If
    
    If TabShow.Tab = 0 Then
        Chk显示退药待发单据.Visible = True
    Else
        Chk显示退药待发单据.Visible = False
    End If
    
    Cbo批号.Visible = False
    
    '换页的时候保存和恢复条件处理
    If mintLastTab <> TabShow.Tab Then
        '保存上个页面条件
        Call SetCondition(IIf(mintLastTab = 4, 1, 0))
        
        Call SaveCondition(IIf(mintLastTab = 4, 1, 0))
        
        '恢复当前页面条件
        Call LoadCondition(IIf(TabShow.Tab = 4, 1, 0))
        
        Call SetCondition(IIf(TabShow.Tab = 4, 1, 0))
    End If
    mintLastTab = TabShow.Tab
    
    '如果是未发药清单或发退药清单，则保存其设置
    Dim strTag As String
    If Bill未发药清单.ColWidth(列名_未发药清单.状态) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.状态) = 700
    If Bill未发药清单.ColWidth(列名_未发药清单.批号) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.批号) = 1500
    If Bill已发药清单.ColWidth(列名_已发药清单.状态) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.状态) = 700
    If Bill已发药清单.ColWidth(列名_已发药清单.退药数) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.退药数) = 1000
    
    Bill汇总发药.ColWidth(列名_科室汇总清单.销帐数量) = IIf(mbln汇总发药 = True, 1200, 0)
    
    If PreviousTab = 0 Then
        strTag = Bill未发药清单.Tag: Bill未发药清单.Tag = ""
        Call SaveFlexState(Bill未发药清单, "未发药清单")
        Bill未发药清单.Tag = strTag
    End If
    If PreviousTab = 4 Then
        strTag = Bill已发药清单.Tag: Bill已发药清单.Tag = ""
        Call SaveFlexState(Bill已发药清单, "已发药清单")
        Bill已发药清单.Tag = strTag
    End If
    If PreviousTab = 1 Then
        strTag = Bill汇总发药.Tag: Bill汇总发药.Tag = ""
        Call SaveFlexState(Bill汇总发药, "汇总发药" & Lng汇总显示)
        Bill汇总发药.Tag = strTag
    End If
    If PreviousTab = 2 Then
        strTag = Bill缺药清单.Tag: Bill缺药清单.Tag = ""
        Call SaveFlexState(Bill缺药清单, "缺药清单")
        Bill缺药清单.Tag = strTag
    End If
    If PreviousTab = 3 Then
        strTag = Bill拒发药清单.Tag: Bill拒发药清单.Tag = ""
        Call SaveFlexState(Bill拒发药清单, "拒发药清单")
        Bill拒发药清单.Tag = strTag
    End If
    
    Chk清单.Visible = (TabShow.Tab = 1 Or TabShow.Tab = 4)
    Chk清单.Enabled = Chk清单.Visible
    Chk清单.Caption = IIf(TabShow.Tab = 4, "显示所有过程单据", "按药品批次汇总")
    If Chk清单.Tag <> "" Then
        If Chk清单.Value <> IIf(TabShow.Tab = 1, Mid(Chk清单.Tag, 1, 1), Mid(Chk清单.Tag, 2, 1)) Then
            Chk清单.Value = IIf(TabShow.Tab = 1, Mid(Chk清单.Tag, 1, 1), Mid(Chk清单.Tag, 2, 1))
            Exit Sub
        End If
    Else
        Chk清单.Tag = "00"
    End If
    
    Call RefreshDataBaseOnPage
    
    MnuViewTotal.Enabled = (TabShow.Tab = 4)
    MnuViewNone.Enabled = (TabShow.Tab = 4)
    MnuViewLocate.Enabled = (TabShow.Tab = 0 Or TabShow.Tab = 4)
    MnuViewLocateNext.Enabled = (MnuViewLocate.Enabled And Val(MnuViewLocateNext.Tag))
    
    mnuFileRestore = False
    Select Case TabShow.Tab
    Case 0
        TxtInput.Visible = False
        Bill未发药清单.Col = 列名_未发药清单.状态
        Bill未发药清单.SetFocus
        Call SetMenu(Trim(Bill未发药清单.TextMatrix(1, 列名_未发药清单.科室)) <> "")
    Case 1
        TxtInput.Visible = False
        Bill汇总发药.SetFocus
        Call SetMenu(Trim(Bill汇总发药.TextMatrix(1, 0)) <> "")
    Case 2
        TxtInput.Visible = False
        Bill缺药清单.SetFocus
        Call SetMenu(Trim(Bill缺药清单.TextMatrix(1, 0)) <> "")
    Case 3
        TxtInput.Visible = False
        Bill拒发药清单.SetFocus
        Call SetMenu(Trim(Bill拒发药清单.TextMatrix(1, 0)) <> "")
    Case 4
        If Not BlnInRefresh Then
            '已发处方记录集
            Set RecChangeSendedData = New adodb.Recordset
            With RecChangeSendedData
                If .State = 1 Then .Close
                .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "ID", adDouble, 18, adFldIsNullable
                .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
                .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
                .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
                .Fields.Append "单据", adDouble, 18, adFldIsNullable
                .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
                .Fields.Append "序号", adDouble, 18, adFldIsNullable
                .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
                .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "毒理分类", adLongVarChar, 10, adFldIsNullable
                .Fields.Append "批次", adDouble, 18, adFldIsNullable
                .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "分批", adDouble, 2, adFldIsNullable
                .Fields.Append "付", adDouble, 18, adFldIsNullable
                .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "已退数", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "准退数", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "退药数", adDouble, 18, adFldIsNullable
                .Fields.Append "可操作", adDouble, 2, adFldIsNullable
                .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
                .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
                .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
                .Fields.Append "操作员", adLongVarChar, 20, adFldIsNullable
                .Fields.Append "发药时间", adLongVarChar, 40, adFldIsNullable
                .Fields.Append "位置", adDouble, 18, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .Open
            End With
        End If
        
        Call SetCondition(1)
        If mblnFirstSended = False Then
            Call zlCommFun.ShowFlash
            If RefreshSendedData = False Then Call zlCommFun.StopFlash: Exit Sub
            Bill已发药清单.Col = 1
            Bill已发药清单.SetFocus
            Call SetMenu(Trim(Bill已发药清单.TextMatrix(1, 列名_已发药清单.科室)) <> "")
            Call zlCommFun.StopFlash
        End If
'        mblnFirstSended = False
    End Select
    
    If TabShow.Tab = 0 Then
        Call SetColHide(Bill未发药清单)
    ElseIf TabShow.Tab = 4 Then
        Call SetColHide(Bill已发药清单)
    End If
End Sub

Private Function GetDetailCol(ByVal strText As String, ByVal Bill As MSHFlexGrid) As Integer
    Dim intCol As Integer, intCols As Integer
    intCols = Bill.Cols - 1
    If strText = "用量" Then strText = "单量"
    If strText = "领/退药人" Then
        If TabShow.Tab = 0 Then
            strText = "退药人"
        End If
    End If
            
    For intCol = 0 To intCols
        If Trim(Bill.TextMatrix(0, intCol)) = strText Then
            GetDetailCol = intCol
            Exit Function
        End If
    Next
    GetDetailCol = -1
End Function

Private Sub SetFormat()
    Dim intFixedCol As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    Dim i As Integer
    
    '设置各列表控件的格式
    With Bill未发药清单
        .rows = 2
        .Cols = 列名_未发药清单.列数
        
        .TextMatrix(0, 列名_未发药清单.审查结果) = "警"
        .TextMatrix(0, 列名_未发药清单.分组符) = "组"
        .TextMatrix(0, 列名_未发药清单.科室) = "科室"
        .TextMatrix(0, 列名_未发药清单.开单医生) = "开单医生"
        .TextMatrix(0, 列名_未发药清单.状态) = "状态"
        .TextMatrix(0, 列名_未发药清单.类型) = "类型"
        .TextMatrix(0, 列名_未发药清单.NO) = "NO"
        .TextMatrix(0, 列名_未发药清单.记帐员) = "记帐员"
        .TextMatrix(0, 列名_未发药清单.床号) = "床号"
        .TextMatrix(0, 列名_未发药清单.姓名) = "姓名"
        .TextMatrix(0, 列名_未发药清单.住院号) = "住院号"
        .TextMatrix(0, 列名_未发药清单.药品名称) = "药品名称"
        .TextMatrix(0, 列名_未发药清单.其它名) = "其它名"
        .TextMatrix(0, 列名_未发药清单.英文名) = "英文名"
        .TextMatrix(0, 列名_未发药清单.规格) = "规格"
        .TextMatrix(0, 列名_未发药清单.产地) = "产地"
        .TextMatrix(0, 列名_未发药清单.批号) = "批号"
        .TextMatrix(0, 列名_未发药清单.付) = "付"
        .TextMatrix(0, 列名_未发药清单.数量) = "数量"
        .TextMatrix(0, 列名_未发药清单.单价) = "单价"
        .TextMatrix(0, 列名_未发药清单.金额) = "金额"
        .TextMatrix(0, 列名_未发药清单.单量) = "单量"
        .TextMatrix(0, 列名_未发药清单.频次) = "频次"
        .TextMatrix(0, 列名_未发药清单.用法) = "用法"
        .TextMatrix(0, 列名_未发药清单.记帐时间) = "记帐时间"
        .TextMatrix(0, 列名_未发药清单.说明) = "说明"
        .TextMatrix(0, 列名_未发药清单.单据) = "单据"
        .TextMatrix(0, 列名_未发药清单.医嘱id) = "医嘱id"
        .TextMatrix(0, 列名_未发药清单.退药人) = "退药人"
        .TextMatrix(0, 列名_未发药清单.库房货位) = "库房货位"
        .TextMatrix(0, 列名_未发药清单.相关ID) = ""
        .TextMatrix(0, 列名_未发药清单.药品ID) = ""
        .TextMatrix(0, 列名_未发药清单.单量单位) = ""
        .TextMatrix(0, 列名_未发药清单.领药部门) = "领药部门"
        .TextMatrix(0, 列名_未发药清单.领药部门id) = ""
                
        .ColWidth(列名_未发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
        .ColWidth(列名_未发药清单.分组符) = 0
        .ColWidth(列名_未发药清单.科室) = 1000
        .ColWidth(列名_未发药清单.开单医生) = 1100
        .ColWidth(列名_未发药清单.状态) = 700
        .ColWidth(列名_未发药清单.类型) = 900
        .ColWidth(列名_未发药清单.NO) = 900
        .ColWidth(列名_未发药清单.记帐员) = 800
        .ColWidth(列名_未发药清单.床号) = 600
        .ColWidth(列名_未发药清单.姓名) = 700
        .ColWidth(列名_未发药清单.住院号) = 1200
        .ColWidth(列名_未发药清单.药品名称) = 2000
        .ColWidth(列名_未发药清单.其它名) = 2000
        .ColWidth(列名_未发药清单.英文名) = 2000
        .ColWidth(列名_未发药清单.规格) = 1500
        .ColWidth(列名_未发药清单.产地) = 1500
        .ColWidth(列名_未发药清单.批号) = 1500
        .ColWidth(列名_未发药清单.付) = 300
        .ColWidth(列名_未发药清单.数量) = 1200
        .ColWidth(列名_未发药清单.单价) = 1200
        .ColWidth(列名_未发药清单.单量) = 1200
        .ColWidth(列名_未发药清单.单量) = 500
        .ColWidth(列名_未发药清单.频次) = 500
        .ColWidth(列名_未发药清单.用法) = 800
        .ColWidth(列名_未发药清单.说明) = 1200
        .ColWidth(列名_未发药清单.记帐时间) = 1800
        .ColWidth(列名_未发药清单.单据) = 0
        .ColWidth(列名_未发药清单.医嘱id) = 0
        .ColWidth(列名_未发药清单.退药人) = 1000
        .ColWidth(列名_未发药清单.库房货位) = 1200
        .ColWidth(列名_未发药清单.相关ID) = 0
        .ColWidth(列名_未发药清单.药品ID) = 0
        .ColWidth(列名_未发药清单.单量单位) = 0
        .ColWidth(列名_未发药清单.领药部门) = 1000
        .ColWidth(列名_未发药清单.领药部门id) = 0
        
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(列名_未发药清单.规格) = 1
        .ColAlignment(列名_未发药清单.批号) = 1
        .ColAlignment(列名_未发药清单.记帐时间) = 1
        .ColAlignment(列名_未发药清单.其它名) = 1
        .ColAlignment(列名_未发药清单.英文名) = 1
    End With
    
    With Bill汇总发药
        .rows = 2
        .Cols = IIf(Lng汇总显示 = 1, 列名_科室汇总清单.列数, 列名_汇总清单.列数)
        
        If Lng汇总显示 = 0 Then
            .TextMatrix(0, 列名_汇总清单.药品名称) = "药品名称"
            .TextMatrix(0, 列名_汇总清单.规格) = "规格"
            .TextMatrix(0, 列名_汇总清单.产地) = "产地"
            .TextMatrix(0, 列名_汇总清单.批号) = "批号"
            .TextMatrix(0, 列名_汇总清单.数量) = "数量"
            .TextMatrix(0, 列名_汇总清单.单位) = "单位"
            .TextMatrix(0, 列名_汇总清单.单价) = "单价"
            .TextMatrix(0, 列名_汇总清单.金额) = "金额"
                        
            .ColWidth(列名_汇总清单.药品名称) = 2000
            .ColWidth(列名_汇总清单.规格) = 1500
            .ColWidth(列名_汇总清单.产地) = 1500
            .ColWidth(列名_汇总清单.批号) = 1200
            .ColWidth(列名_汇总清单.数量) = 1200
            .ColWidth(列名_汇总清单.单位) = 500
            .ColWidth(列名_汇总清单.单价) = 1200
            .ColWidth(列名_汇总清单.金额) = 1200
        Else
            .TextMatrix(0, 列名_科室汇总清单.科室) = "开单科室"
            .TextMatrix(0, 列名_科室汇总清单.药品名称) = "药品名称"
            .TextMatrix(0, 列名_科室汇总清单.规格) = "规格"
            .TextMatrix(0, 列名_科室汇总清单.产地) = "产地"
            .TextMatrix(0, 列名_科室汇总清单.批号) = "批号"
            .TextMatrix(0, 列名_科室汇总清单.应发数量) = "应发数量"
            .TextMatrix(0, 列名_科室汇总清单.留存数量) = "留存数量"
            .TextMatrix(0, 列名_科室汇总清单.销帐数量) = "销帐数量"
            .TextMatrix(0, 列名_科室汇总清单.实发数量) = "实发数量"
            .TextMatrix(0, 列名_科室汇总清单.单位) = "单位"
            .TextMatrix(0, 列名_科室汇总清单.单价) = "单价"
            .TextMatrix(0, 列名_科室汇总清单.金额) = "金额"
            .TextMatrix(0, 列名_科室汇总清单.批次) = "批次"
            .TextMatrix(0, 列名_科室汇总清单.科室ID) = "科室ID"
            .TextMatrix(0, 列名_科室汇总清单.药品ID) = "药品ID"
            .TextMatrix(0, 列名_科室汇总清单.领药部门) = "领药部门"
            .TextMatrix(0, 列名_科室汇总清单.领药部门id) = ""
            
            
            .ColWidth(列名_科室汇总清单.科室) = 1200
            .ColWidth(列名_科室汇总清单.药品名称) = 2000
            .ColWidth(列名_科室汇总清单.规格) = 1500
            .ColWidth(列名_科室汇总清单.产地) = 1500
            .ColWidth(列名_科室汇总清单.批号) = 1200
            .ColWidth(列名_科室汇总清单.应发数量) = 1200
            .ColWidth(列名_科室汇总清单.留存数量) = 1200
            .ColWidth(列名_科室汇总清单.销帐数量) = IIf(mbln汇总发药 = True, 1200, 0)
            .ColWidth(列名_科室汇总清单.实发数量) = 1200
            .ColWidth(列名_科室汇总清单.单位) = 500
            .ColWidth(列名_科室汇总清单.单价) = 1200
            .ColWidth(列名_科室汇总清单.金额) = 1200
            .ColWidth(列名_科室汇总清单.批次) = 0
            .ColWidth(列名_科室汇总清单.科室ID) = 0
            .ColWidth(列名_科室汇总清单.药品ID) = 0
            .ColWidth(列名_科室汇总清单.领药部门) = 1200
            .ColWidth(列名_科室汇总清单.领药部门id) = 0
        End If
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(IIf(Lng汇总显示 = 1, 列名_科室汇总清单.规格, 列名_汇总清单.规格)) = 1
        .ColAlignment(IIf(Lng汇总显示 = 1, 列名_科室汇总清单.批号, 列名_汇总清单.批号)) = 1
    End With

    With Bill缺药清单
        .rows = 2
        .Cols = 12
        
        .TextMatrix(0, 0) = "科室"
        .TextMatrix(0, 1) = "NO"
        .TextMatrix(0, 2) = "类型"
        .TextMatrix(0, 3) = "床号"
        .TextMatrix(0, 4) = "姓名"
        .TextMatrix(0, 5) = "药品名称"
        .TextMatrix(0, 6) = "规格"
        .TextMatrix(0, 7) = "产地"
        .TextMatrix(0, 8) = "批号"
        .TextMatrix(0, 9) = "数量"
        .TextMatrix(0, 10) = "单价"
        .TextMatrix(0, 11) = "金额"
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 800
        .ColWidth(2) = 900
        .ColWidth(3) = 800
        .ColWidth(4) = 1000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1000
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(6) = 1
    End With

    With Bill拒发药清单
        .rows = 2
        .Cols = 13
        
        .TextMatrix(0, 0) = "科室"
        .TextMatrix(0, 1) = "状态"
        .TextMatrix(0, 2) = "NO"
        .TextMatrix(0, 3) = "类型"
        .TextMatrix(0, 4) = "床号"
        .TextMatrix(0, 5) = "姓名"
        .TextMatrix(0, 6) = "药品名称"
        .TextMatrix(0, 7) = "规格"
        .TextMatrix(0, 8) = "产地"
        .TextMatrix(0, 9) = "批号"
        .TextMatrix(0, 10) = "数量"
        .TextMatrix(0, 11) = "单价"
        .TextMatrix(0, 12) = "金额"
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 700
        .ColWidth(2) = 800
        .ColWidth(3) = 900
        .ColWidth(4) = 800
        .ColWidth(5) = 1000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(9) = 1500
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
    
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
    End With

    With Bill已发药清单
        .rows = 2
        .Cols = 列名_已发药清单.列数
        
        .TextMatrix(0, 列名_已发药清单.审查结果) = "警"
        .TextMatrix(0, 列名_已发药清单.分组符) = "组"
        .TextMatrix(0, 列名_已发药清单.科室) = "科室"
        .TextMatrix(0, 列名_已发药清单.状态) = "状态"
        .TextMatrix(0, 列名_已发药清单.类型) = "类型"
        .TextMatrix(0, 列名_已发药清单.NO) = "NO"
        .TextMatrix(0, 列名_已发药清单.床号) = "床号"
        .TextMatrix(0, 列名_已发药清单.姓名) = "姓名"
        .TextMatrix(0, 列名_已发药清单.住院号) = "住院号"
        .TextMatrix(0, 列名_已发药清单.药品名称) = "药品名称"
        .TextMatrix(0, 列名_已发药清单.其它名) = "其它名"
        .TextMatrix(0, 列名_已发药清单.英文名) = "英文名"
        .TextMatrix(0, 列名_已发药清单.规格) = "规格"
        .TextMatrix(0, 列名_已发药清单.产地) = "产地"
        .TextMatrix(0, 列名_已发药清单.批号) = "批号"
        .TextMatrix(0, 列名_已发药清单.付) = "付"
        .TextMatrix(0, 列名_已发药清单.数量) = "数量"
        .TextMatrix(0, 列名_已发药清单.已退数) = "已退数"
        .TextMatrix(0, 列名_已发药清单.准退数) = "准退数"
        .TextMatrix(0, 列名_已发药清单.退药数) = "退药数"
        .TextMatrix(0, 列名_已发药清单.单价) = "单价"
        .TextMatrix(0, 列名_已发药清单.金额) = "金额"
        .TextMatrix(0, 列名_已发药清单.单量) = "单量"
        .TextMatrix(0, 列名_已发药清单.频次) = "频次"
        .TextMatrix(0, 列名_已发药清单.用法) = "用法"
        .TextMatrix(0, 列名_已发药清单.操作员) = "操作员"
        .TextMatrix(0, 列名_已发药清单.发药时间) = "发药时间"
        .TextMatrix(0, 列名_已发药清单.单据) = "单据"
        .TextMatrix(0, 列名_已发药清单.医嘱id) = "医嘱id"
        .TextMatrix(0, 列名_已发药清单.领药人) = "领/退药人"
        .TextMatrix(0, 列名_已发药清单.库房货位) = "库房货位"
        .TextMatrix(0, 列名_已发药清单.相关ID) = ""
        .TextMatrix(0, 列名_已发药清单.药品ID) = ""
        .TextMatrix(0, 列名_已发药清单.单量单位) = ""
                
        .ColWidth(列名_已发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
        .ColWidth(列名_已发药清单.分组符) = 0
        .ColWidth(列名_已发药清单.科室) = 1200
        .ColWidth(列名_已发药清单.状态) = 700
        .ColWidth(列名_已发药清单.类型) = 900
        .ColWidth(列名_已发药清单.NO) = 900
        .ColWidth(列名_已发药清单.床号) = 600
        .ColWidth(列名_已发药清单.姓名) = 700
        .ColWidth(列名_已发药清单.住院号) = 1200
        .ColWidth(列名_已发药清单.药品名称) = 2000
        .ColWidth(列名_已发药清单.其它名) = 2000
        .ColWidth(列名_已发药清单.英文名) = 2000
        .ColWidth(列名_已发药清单.规格) = 1500
        .ColWidth(列名_已发药清单.产地) = 1500
        .ColWidth(列名_已发药清单.批号) = 1500
        .ColWidth(列名_已发药清单.付) = 300
        .ColWidth(列名_已发药清单.数量) = 1000
        .ColWidth(列名_已发药清单.已退数) = 1000
        .ColWidth(列名_已发药清单.准退数) = 1000
        .ColWidth(列名_已发药清单.退药数) = 1000
        .ColWidth(列名_已发药清单.单价) = 1000
        .ColWidth(列名_已发药清单.金额) = 1000
        .ColWidth(列名_已发药清单.单量) = 500
        .ColWidth(列名_已发药清单.频次) = 500
        .ColWidth(列名_已发药清单.用法) = 800
        .ColWidth(列名_已发药清单.操作员) = 800
        .ColWidth(列名_已发药清单.发药时间) = 1500
        .ColWidth(列名_已发药清单.单据) = 0
        .ColWidth(列名_已发药清单.医嘱id) = 0
        .ColWidth(列名_已发药清单.领药人) = 1000
        .ColWidth(列名_已发药清单.库房货位) = 1200
        .ColWidth(列名_已发药清单.相关ID) = 0
        .ColWidth(列名_已发药清单.药品ID) = 0
        .ColWidth(列名_已发药清单.单量单位) = 0
                
        For intFixedCol = 0 To .Cols - 1
            .ColAlignmentFixed(intFixedCol) = 4
        Next
        .ColAlignment(列名_已发药清单.规格) = 1
        .ColAlignment(列名_已发药清单.批号) = 1
        .ColAlignment(列名_已发药清单.用法) = 1
        .ColAlignment(列名_已发药清单.已退数) = 7
        .ColAlignment(列名_已发药清单.准退数) = 7
        .ColAlignment(列名_已发药清单.退药数) = 7
        .ColAlignment(列名_已发药清单.发药时间) = 1
        .ColAlignment(列名_已发药清单.其它名) = 1
        .ColAlignment(列名_已发药清单.英文名) = 1
    End With
  
    '初始销帐列表
    strTemp = Split(mconstRequest, "|")
    With Bill退药销帐
        .Redraw = False
        .rows = 2
        .Cols = 销帐列表.列数
        .SelectionMode = flexSelectionByRow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            
            If strArr(0) = "效期" Then
                .TextMatrix(0, i) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
            Else
                .TextMatrix(0, i) = strArr(0)
            End If
            
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = 4
        Next
        .Redraw = True
    End With
  
    '保存初始的列宽
    Call SaveColDefaultWidth(Bill未发药清单)
    Call SaveColDefaultWidth(Bill已发药清单)
  
    Call RestoreFlexState(Bill汇总发药, "汇总发药" & Lng汇总显示)
    Call RestoreFlexState(Bill缺药清单, "缺药清单")
    Call RestoreFlexState(Bill拒发药清单, "拒发药清单")
    Call RestoreFlexState(Bill未发药清单, "未发药清单")
    Call RestoreFlexState(Bill已发药清单, "已发药清单")
    '恢复个性化设置后，有几列始终不能隐藏
    If Bill未发药清单.ColWidth(列名_未发药清单.状态) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.状态) = 700
    If Bill未发药清单.ColWidth(列名_未发药清单.批号) < 200 Then Bill未发药清单.ColWidth(列名_未发药清单.批号) = 1500
    If Bill已发药清单.ColWidth(列名_已发药清单.状态) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.状态) = 700
    If Bill已发药清单.ColWidth(列名_已发药清单.退药数) < 200 Then Bill已发药清单.ColWidth(列名_已发药清单.退药数) = 1000
    '警示列根据参数来决定是否显示
    Bill未发药清单.ColWidth(列名_未发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
    Bill已发药清单.ColWidth(列名_已发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
    
    Bill汇总发药.ColWidth(列名_科室汇总清单.销帐数量) = IIf(mbln汇总发药 = True, 1200, 0)
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
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BBACKSTRICK", vbResIcon)
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
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CBACKSTRICK", vbResIcon)
    End With
    With Tbar
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor

        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Consignment").Image = 3
        .Buttons("Desire").Image = 4
        .Buttons("Handback").Image = 5
        .Buttons("Restore").Image = 6
        .Buttons("Help").Image = 7
        .Buttons("Exit").Image = 8
        .Buttons("ReVerify").Image = 9
    End With
    Cbar.Bands(1).MinHeight = Tbar.Height
    
    If err <> 0 Then
        MsgBox "相关资源文件丢失，请与软件开发商联系！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function InitRefreshRec()
    
    '用于执行功能（发药、拒发）后，将上次设定的非发药及缺药的记录的执行状态恢复
    Set RecRefreshCompare = New adodb.Recordset
    With RecRefreshCompare
        If .State = 1 Then .Close
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "状态", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "付", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adDouble, 18, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "已收费", adDouble, 2, adFldIsNullable
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable            '判断库存用
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "记帐时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "审核人", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Function InitRec()
    '编制人:朱玉宝
    '编制日期:2000-11-02
    
    '未发处方记录集
    If Bln刷新未发药清单 = True Then
        Set RecChangeData = New adodb.Recordset
        With RecChangeData
            If .State = 1 Then .Close
            .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "开单医生", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "状态", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
            .Fields.Append "序号", adDouble, 18, adFldIsNullable
            .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "其它名", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "英文名", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "毒理分类", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "价值分类", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "批次", adDouble, 18, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "分批", adDouble, 2, adFldIsNullable
            .Fields.Append "付", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "记帐员", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
            .Fields.Append "ID", adDouble, 18, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "已收费", adDouble, 2, adFldIsNullable
            .Fields.Append "位置", adDouble, 18, adFldIsNullable
            .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
            .Fields.Append "实际数量", adDouble, 18, adFldIsNullable            '判断库存用
            .Fields.Append "留存数量", adDouble, 18, adFldIsNullable
            .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "记帐时间", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "配药人", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "审核人", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "审查结果", adDouble, 18, adFldIsNullable
            .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
            .Fields.Append "退药人", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "相关id", adDouble, 18, adFldIsNullable
            .Fields.Append "科室ID", adDouble, 18, adFldIsNullable
            .Fields.Append "单量单位", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "库存下限", adDouble, 18, adFldIsNullable
            .Fields.Append "领药部门", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
                        
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set mrsRequest = New adodb.Recordset
        With mrsRequest
            If .State = 1 Then .Close
            .Fields.Append "领药部门", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
            .Fields.Append "单据", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "包装", adDouble, 18, adFldIsNullable
            .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
            .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
            .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
            .Fields.Append "险类", adDouble, 18, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
            .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
            .Fields.Append "药品名称", adLongVarChar, 100, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set mrsRequestMain = New adodb.Recordset
        With mrsRequestMain
            If .State = 1 Then .Close
            .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
            .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
            .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    '已发处方记录集
    Set RecChangeSendedData = New adodb.Recordset
    With RecChangeSendedData
        If .State = 1 Then .Close
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "其它名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "英文名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "毒理分类", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "分批", adDouble, 2, adFldIsNullable
        .Fields.Append "付", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "已退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "退药数", adDouble, 18, adFldIsNullable
        .Fields.Append "可操作", adDouble, 2, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "操作员", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发药时间", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "审查结果", adDouble, 18, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "领药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "实际数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "相关id", adDouble, 18, adFldIsNullable
        .Fields.Append "单量单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "转出", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function
Private Function DependOnCheck() As Boolean
    Dim strSQL As String
    Dim BlnIn药房 As Boolean
    Dim n As Integer
    
    On Error GoTo errHandle
    '依赖数据检测
    DependOnCheck = False
    
   '检测药房设置否(中药房、西药房及成药房)
    If IsHavePrivs(mstrPrivs, "所有药房") Then
        strSQL = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房' And 服务对象 IN (2,3))"
    Else
        strSQL = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房' And B.服务对象 IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In " & strSQL & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecBillData
        If .EOF Then
           If IsHavePrivs(mstrPrivs, "所有药房") Then
               strSQL = "请初始化药房（部门管理）"
           Else
               strSQL = "你不是药房工作人员，不能操作本模块！"
           End If
           MsgBox strSQL, vbInformation, gstrSysName
           Exit Function
        Else
            Cbo发药药房.Clear
            Do While Not .EOF
                Cbo发药药房.AddItem !名称
                Cbo发药药房.ItemData(Cbo发药药房.NewIndex) = !Id
                .MoveNext
            Loop
            Cbo发药药房.ListIndex = 0
       End If
       
       Call ReadFromReg
       Call mnuViewFontSet_Click(intFont)
       
       If lng药房ID <> 0 Then
           .MoveFirst
           .Find "ID=" & lng药房ID
           BlnIn药房 = (.EOF <> True)
       End If
       
       '设置对应的药房
       If lng药房ID = 0 Or BlnIn药房 = False Then
           '调设置窗体
            With Frm部门发药参数设置
                .strPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg
            
            If lng药房ID = 0 Then
                MsgBox "需重新设置药房，请与系统管理员联系！", vbInformation, gstrSysName
                Exit Function
            End If
           
           '仍未设置药房，退出
           If lng药房ID = 0 Then Exit Function
           .MoveFirst
           .Find "ID=" & lng药房ID
           BlnIn药房 = (.EOF <> True)
           If Not BlnIn药房 Then Exit Function
       End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadFromReg() As Boolean
    Dim strTemp As String
    Dim RecRead As New adodb.Recordset
    Dim dateCurDate As Date
    Dim strArr
    Dim n As Integer
    
    On Error GoTo errHandle
    '取公共及私有参数
    '私有模块
    intFont = Val(zldatabase.GetPara("字体", glngSys, 1342))
    StrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    '公共模块
    intDays = Val(zldatabase.GetPara("查询天数", glngSys, 1342)) - 1
    int发药规则 = Val(zldatabase.GetPara("发药规则", glngSys, 1342))
    Lng领药人签名 = Val(zldatabase.GetPara("领药人签名", glngSys, 1342))
    Lng缺药检查 = Val(zldatabase.GetPara("缺药检查", glngSys, 1342))
    Lng退药人签名 = Val(zldatabase.GetPara("退药人签名", glngSys, 1342))
    mint自动刷新未发药清单 = Val(zldatabase.GetPara("自动刷新未发药清单", glngSys, 1342))
    mstr病区发药方式 = zldatabase.GetPara("病区发药方式", glngSys, 1342, "临床,护理,检查,检验,手术,治疗,营养")
    mbln药品储备 = (Val(zldatabase.GetPara("库房货位及库存限量提示", glngSys, 1342, 0)) = 1)
    mbln汇总发药 = (Val(zldatabase.GetPara("发药时汇总退药销帐记录", glngSys, 1342, 0)) = 1)

    Lng操作模式 = Val(zldatabase.GetPara("操作模式", glngSys, 1342))
    Lng医嘱类型 = Val(zldatabase.GetPara("医嘱类型", glngSys, 1342))
    int离院带药 = Val(zldatabase.GetPara("出院带药", glngSys, 1342))
    Lng汇总显示 = Val(zldatabase.GetPara("按科室汇总显示汇总清单", glngSys, 1342))
    str记帐人 = zldatabase.GetPara("记帐人", glngSys, 1342, "所有记帐人")
    mstr毒理分类 = zldatabase.GetPara("毒理分类", glngSys, 1342)
    mstr价值分类 = zldatabase.GetPara("价值分类", glngSys, 1342)
    
    lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1342))
    Lng自动打印 = Val(zldatabase.GetPara("自动打印", glngSys, 1342))
    
    mlng待发单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\Frm部门发药管理", "显示退药待发单据", 1)
    

    '根据参数设置
    Chk显示退药待发单据.Value = mlng待发单据
    
    '[药品出库库存检查]系统参数
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set RecRead = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
    
    If Not RecRead.EOF Then
        IntCheckStock = RecRead!库存检查
    End If
    
    '根据系统参数设定的单位显示数据
    strUnit = GetSpecUnit(lng药房ID, gint住院药房)
    
    '设置当前药房
    If lng药房ID > 0 And Cbo发药药房.ListCount > 0 Then
        Cbo发药药房.Tag = lng药房ID
        For n = 0 To Cbo发药药房.ListCount - 1
            If lng药房ID = Cbo发药药房.ItemData(n) Then
                Cbo发药药房.ListIndex = n
                Exit For
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetSysParms()
    Int允许未审核处方发药 = gtype_UserSysParms.P6_未审核记帐处方发药
    
    bln医嘱作废 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0)          '为零表示允许退药
    
    '获取金额小数位数
    int金额保留位数 = gtype_UserSysParms.P9_费用金额保留位数
    
    '判断划价单发药后是否自动审核为记帐单
    int审核划价单 = gtype_UserSysParms.P81_执行后自动审核划价单
End Sub
Private Function RefreshData() As Boolean
    Dim strCond As String, strSubSql As String
    Dim strName As String
    Dim str过滤记帐人 As String
    Dim strSql退药人 As String
    Dim strSql病人类型 As String
    
    RefreshData = False
    On Error GoTo errHandle
    '必要条件检查
    If mstrSerchNO = "" And mstr部门 = "" Then
'        MsgBox "请选择领药部门！", vbInformation, gstrSysName
        Call ClearCons
        Exit Function
    End If
    
    str过滤记帐人 = IIf(str记帐人 <> "所有记帐人", " AND S.填制人=[1] ", "")
    
    '扣率:bit1=0-长嘱,1-临嘱；bit2:3-离院带药
    '操作模式:0-所有,1-记帐单,2-记帐表
    If Lng操作模式 = 0 Then
        strCond = " And S.单据 IN(9,10)"
    ElseIf Lng操作模式 = 1 Then
        strCond = " And S.单据=9"
    ElseIf Lng操作模式 = 2 Then
        strCond = " And S.单据=10"
    End If
    
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If Lng医嘱类型 = 0 Then
    ElseIf Lng医嘱类型 = 1 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf Lng医嘱类型 = 2 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf Lng医嘱类型 = 3 Then
        strCond = strCond & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf Lng医嘱类型 = 4 Then
        strCond = strCond & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If int离院带药 = 0 Then
    ElseIf int离院带药 = 1 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf int离院带药 = 2 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf int离院带药 = 3 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 4 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 5 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 6 Then
        strCond = strCond & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    If mint类型 = 0 Then
        strCond = strCond & " And H.Id = C.病人科室id "
    ElseIf mint类型 = 1 Then
        strCond = strCond & " And H.Id = C.开单部门id "
    Else
        strCond = strCond & " And H.Id = C.病人病区ID "
    End If
    
    '单位设置
    Select Case strUnit
    Case "售价单位"
        strSubSql = "X.计算单位 单位,1 包装,"
    Case "门诊单位"
        strSubSql = "D.门诊单位 单位,D.门诊包装 包装,"
    Case "住院单位"
        strSubSql = "D.住院单位 单位,D.住院包装 包装,"
    Case "药库单位"
        strSubSql = "D.药库单位 单位,D.药库包装 包装,"
    End Select
    
    '得到药品名称串
    Call GetDrugFormat
    Select Case int药品名称
    Case 0  '药品编码与名称
        strName = "'['||X.编码||']'||" & IIf(mblnTradeName, "NVL(E.名称,X.名称)", "X.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "X.编码" & " As 品名,"
    Case 2  '药品名称
        strName = IIf(mblnTradeName, "NVL(E.名称,X.名称)", "X.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(E.名称,'')", "Decode(E.名称,Null,'',X.名称)") & " As 其它名, "
    
    '病人类型：病人或婴儿
    If mint病人类型 = 0 Then
        strSql病人类型 = " And Nvl(C.婴儿费,0)=0 "
    ElseIf mint病人类型 = 1 Then
        strSql病人类型 = " And Nvl(C.婴儿费,0)>0 "
    End If
    
    gstrSQL = "SELECT A.*, Nvl(C.留存数量,0) As 留存数量 " & IIf(mbln显示领退药人 = True, ", B.退药人", "") & " FROM " & _
             " (SELECT DISTINCT S.ID,S.药品ID,NVL(N.已收费,0) 已收费,P.名称 科室,S.配药人,C.开单人 开单医生,C.操作员姓名 审核人,S.单据,S.扣率," & _
             " S.NO,S.序号,C.病人ID,C.床号,C.姓名,C.门诊标志,C.标识号,C.操作员姓名," & strName & " S.付数 付,S.实际数量 数量," & _
             " NVL(D.药房分批,0) 分批,X.规格,T.毒理分类,T.价值分类,C.登记时间,H.名称 As 领药部门,H.Id As 领药部门Id," & _
             strSubSql & _
             " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,C.医嘱序号,I.计算单位,NVL(S.产地,NVL(X.产地,'')) 产地,nvl(M.审查结果,-1) 审查结果,nvl(C.医嘱序号,-1) 医嘱id," & IIf(mbln药品储备 = True, "L.", "'' ") & "库房货位,M.相关ID,S.对方部门id As 科室ID,C.序号 费用序号," & IIf(mbln药品储备 = True, "Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) ", "0 ") & " 库存下限, Z.名称 As 英文名 " & _
             " FROM 药品收发记录 S,病人费用记录 C,未发药品记录 N,部门表 P,部门表 H,收费项目别名 E,收费项目目录 X,药品规格 D,药品特性 T,诊疗项目目录 I,病人医嘱记录 M," & IIf(mbln药品储备 = True, "药品储备限额 L,", "") & "诊疗项目别名 Z "
             
    If mbln药品储备 = True Then
        gstrSQL = gstrSQL & ",(Select 库房id, 药品id, Nvl(Sum(实际数量), 0) 库存数量 From 药品库存 Where 性质 = 1 And 库房id = [2] Group By 库房id, 药品id) K "
    End If
             
    gstrSQL = gstrSQL & " WHERE S.药品ID=D.药品ID AND d.药品ID=X.ID and D.药名ID=T.药名ID AND D.药名ID=I.ID and C.医嘱序号=M.ID(+) " & _
             " And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 " & IIf(mbln药品储备 = True, " And S.药品ID=L.药品ID(+) And Nvl(S.库房ID,[2])=L.库房ID(+) ", "") & _
             " AND D.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & IIf(mstr床号 = "", "", " And C.床号=[11] ") & _
             " AND S.NO=N.NO AND S.单据=N.单据 AND NVL(S.库房ID,[2])+0=NVL(N.库房ID,[2]) AND S.费用ID=C.ID " & _
             IIf(Val(mlng病人ID) = 0, "", " AND C.病人ID=[3]") & IIf(Trim(mstr住院号) = "", "", " AND C.标识号=[4]") & IIf(mstr病人姓名 = "", "", " AND C.姓名 LIKE [5] ") & _
             " AND S.对方部门ID+0=P.ID AND S.审核人 IS NULL " & IIf(mstr开始NO = "", "", " AND S.NO>=[6] ") & IIf(mstr结束NO = "", "", " AND S.NO<=[7] ") & _
             " AND NVL(S.库房ID,[2])+0=[2] AND N.填制日期 BETWEEN [8] AND [9] " & IIf(mstrDrug = "", "", " And Instr([14],',' || T.药品剂型 || ',') > 0") & IIf(mstr发药类型 = "", "", " And Instr([15],',' || D.发药类型 || ',') > 0") & _
             " AND NVL(LTRIM(RTRIM(S.摘要)),'小宝')<>'拒发' and nvl(S.发药方式,-999)<>-1 " & strSql病人类型
    
    If mbln药品储备 = True Then
        gstrSQL = gstrSQL & " And Nvl(S.库房id, [2]) + 0 = K.库房id(+) And S.药品id = K.药品id(+) "
    End If
             
    Select Case mint范围
    Case 1
        gstrSQL = gstrSQL & " And S.实际数量>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.实际数量<0"
    End Select
    
    If Trim(mstr部门) <> "" Then
        If mint类型 = 0 Then
            gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id=C.开单部门id"
        ElseIf mint类型 = 1 Then
            gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id<>C.开单部门id"
        Else
            If mstr病区发药方式 = "" Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 And C.病人科室id=C.开单部门id"
            Else
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 "
                If mstr病区发药方式 <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([16],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End If
    Else
        If mint类型 = 0 Then
            gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
        ElseIf mint类型 = 1 Then
            gstrSQL = gstrSQL & " And C.病人科室id<>C.开单部门id"
        Else
            If mstr病区发药方式 = "" Then
                gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
            Else
                If mstr病区发药方式 <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([16],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End If
    End If
    
    If mlng待发单据 = 0 Then
        gstrSQL = gstrSQL & " And S.记录状态 = 1"
    Else
        gstrSQL = gstrSQL & " And Mod(S.记录状态,3)=1"
    End If
    gstrSQL = gstrSQL & strCond & IIf(mstrUse = "", "", " And Instr([13],',' || S.用法 || ',') > 0") & str过滤记帐人 & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ") & " Order By S.No,S.单据) A "
    
    gstrSQL = gstrSQL & ", (Select 药品id,库房id,部门id,留存数量 From 药品留存计划  Where 状态=0) C "
    
    '求最后一次退药的退药人
    If mbln显示领退药人 = True Then
        strSql退药人 = ",(Select a.单据 ,a.No,a.序号,a.领用人 退药人 From 药品收发记录 a," & _
                " (Select s.单据,s.No,s.序号, Max(s.记录状态) 记录状态 " & _
                " From 药品收发记录 s, 未发药品记录 n " & _
                " Where s.No = n.No And s.单据 = n.单据 And Nvl(s.库房id, [2]) + 0 = Nvl(n.库房id, [2]) And " & _
                " Nvl(s.库房id, [2]) + 0 = [2] " & _
                " AND N.填制日期 BETWEEN [8] AND [9] And Nvl(s.发药方式, -999) <> -1 And " & _
                " Mod(s.记录状态, 3) = 2 And s.单据 In (9, 10) " & _
                " Group By s.单据,s.No,s.序号) b " & _
                " Where a.单据=b.单据 And a.No=b.No And a.序号=b.序号 And a.记录状态=b.记录状态) B "
        gstrSQL = gstrSQL & strSql退药人
    End If
    
    gstrSQL = gstrSQL & " Where A.领药部门id = C.部门id(+) And C.库房id(+) = [2] And A.药品id = C.药品id(+) "
    
    If mbln显示领退药人 = True Then
        gstrSQL = gstrSQL & " And A.单据 = B.单据(+) And A.No = B.No(+) And A.序号 = B.序号(+) "
    End If
    
    gstrSQL = gstrSQL & "  Order By a.No,a.费用序号 "
        
    '--刷新数据--
'    on error Resume Next
    err = 0
    
    '初始化记录集
    Call InitRec
    
    '未发处方记录

    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, "提取单据信息", _
        str记帐人, _
        lng药房ID, _
        mlng病人ID, _
        mstr住院号, _
        "%" & mstr病人姓名 & "%", _
        mstr开始NO, _
        mstr结束NO, _
        CDate(mstr开始日期_未发), _
        CDate(mstr结束日期_未发), _
        "," & mstr部门 & ",", _
        mstr床号, _
        mstrSerchNO, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr发药类型 & ",", _
        "," & mstr病区发药方式 & ",")
    
    With RecBillData
        MnuFilePreview.Enabled = Not (.EOF)
        MnuFilePrint.Enabled = Not (.EOF)
        MnuFileExcel.Enabled = Not (.EOF)
        Tbar.Buttons("Preview").Enabled = Not (.EOF)
        Tbar.Buttons("Print").Enabled = Not (.EOF)
    End With
    
    '汇总发药时的一些处理
    If Not RecBillData.EOF Then
        mstrDrawDept = ""
        mstrSendDrugId = ""
        
        '取发药清单中的领药部门和药品，作为参数来取退药销帐清单
        RecBillData.MoveFirst
        Do While Not RecBillData.EOF
            If InStr("," & mstrDrawDept & ",", "," & RecBillData!领药部门id & ",") = 0 Then
                mstrDrawDept = IIf(mstrDrawDept = "", "", mstrDrawDept & ",") & RecBillData!领药部门id
            End If
            
            If InStr("," & mstrSendDrugId & ",", "," & RecBillData!药品ID & ",") = 0 Then
                mstrSendDrugId = IIf(mstrSendDrugId = "", "", mstrSendDrugId & ",") & RecBillData!药品ID
            End If
            
            RecBillData.MoveNext
        Loop
        RecBillData.MoveFirst
        
        Call Get销帐清单
    End If
    
    If ProduceInsideRecordset = False Then Exit Function
    If RefreshDataBaseOnPage(True) = False Then Exit Function
    
    Call ClearBill(Bill拒发药清单)
    Call Load拒发
    Call SetMenuAndToolbarState
    
    If err <> 0 Then
        MsgBox "刷新时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call tabShow_Click(TabShow.Tab)
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ProduceInsideRecordset() As Boolean
    Dim ArrayPhysic
    Dim IntArray As Integer, lngState As Long
    
    '--产生内部记录集(未发)--
    On Error GoTo ErrHand
    err = 0
    ProduceInsideRecordset = False
   
    With RecBillData
        Do While Not .EOF
            RecChangeData.AddNew
            RecChangeData!Id = !Id
            RecChangeData!状态 = "发药"
            RecChangeData!科室 = !科室
            RecChangeData!领药部门 = !领药部门
            RecChangeData!领药部门id = !领药部门id
            RecChangeData!开单医生 = !开单医生
            RecChangeData!类型 = IIf(NVL(!医嘱序号, 0) = 0, IIf(!门诊标志 = 1 Or !门诊标志 = 4, "门诊记帐单", IIf(!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(!扣率) = True, "住院记帐单", IIf(!扣率 Like "0*", "长嘱", IIf(!扣率 Like "1*", "临嘱", "记帐表"))))
            RecChangeData!药品ID = !药品ID
            RecChangeData!位置 = .AbsolutePosition
            RecChangeData!NO = !NO
            RecChangeData!单据 = !单据
            RecChangeData!病人ID = !病人ID
            RecChangeData!序号 = !序号
            RecChangeData!床号 = !床号
            RecChangeData!姓名 = IIf(IsNull(!姓名), "", !姓名)
            RecChangeData!住院号 = NVL(!标识号)
            RecChangeData!品名 = !品名
            RecChangeData!其它名 = !其它名
            RecChangeData!英文名 = !英文名
            RecChangeData!规格 = IIf(IsNull(!规格), "", !规格)
            RecChangeData!产地 = IIf(IsNull(!产地), "", !产地)
            RecChangeData!毒理分类 = IIf(IsNull(!毒理分类), "", !毒理分类)
            RecChangeData!价值分类 = IIf(IsNull(!价值分类), "", !价值分类)
            RecChangeData!批次 = IIf(IsNull(!批次), 0, !批次)
            RecChangeData!批号 = IIf(IsNull(!批号), "", !批号)
            RecChangeData!分批 = IIf(IsNull(!分批), 0, !分批)
            RecChangeData!付 = IIf(IsNull(!付), 1, !付)
            RecChangeData!实际数量 = FormatEx(IIf(IsNull(!数量), 1, !数量) / !包装, 5)
            RecChangeData!留存数量 = FormatEx(IIf(IsNull(!留存数量), 0, !留存数量) / !包装, 5)
            RecChangeData!数量 = FormatEx(IIf(IsNull(!数量), 1, !数量) / !包装, 5) & !单位
            RecChangeData!单价 = FormatEx(!单价 * !包装, 5)
            RecChangeData!金额 = Format(!金额, "#####0.00;-#####0.00; ;")
            RecChangeData!记帐员 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
            RecChangeData!单量 = IIf(IsNull(!单量), "", FormatEx(!单量, 5) & NVL(!计算单位))
            RecChangeData!单量单位 = NVL(!计算单位)
            RecChangeData!频次 = IIf(IsNull(!频次), "", !频次)
            RecChangeData!用法 = IIf(IsNull(!用法), "", !用法)
            RecChangeData!说明 = IIf(IsNull(!说明), "", !说明)
            If IsNull(!登记时间) Then
                RecChangeData!记帐时间 = ""
            Else
                RecChangeData!记帐时间 = Format(!登记时间, "yyyy-MM-dd HH:mm:ss")
            End If
            RecChangeData!配药人 = IIf(IsNull(!配药人), "", !配药人)
            RecChangeData!审核人 = IIf(IsNull(!审核人), "", !审核人)
            RecChangeData!已收费 = !已收费                          '未收费或记帐处方，不允许发药
            RecChangeData!审查结果 = !审查结果
            RecChangeData!医嘱id = !医嘱id
            If mbln显示领退药人 = True Then
                RecChangeData!退药人 = !退药人
            Else
                RecChangeData!退药人 = ""
            End If
            RecChangeData!库房货位 = IIf(IsNull(!库房货位), "", !库房货位)
            RecChangeData!相关ID = IIf(IsNull(!相关ID), 0, !相关ID)
            RecChangeData!科室ID = IIf(IsNull(!科室ID), 0, !科室ID)
            RecChangeData!库存下限 = !库存下限
            '检查是否允许发药
            lngState = 1
            If RecChangeData!已收费 = 0 Then lngState = 3
            '20020903 Modified by zyb
            '如果说明是拒发，则表明该药品已拒发，同时设置其执行状态
            '--Begin
            If Not IsNull(!说明) Then
                lngState = IIf(!说明 = "拒发", 2, lngState)
            End If
            '--End
            If Int允许未审核处方发药 = 0 Then
                If IsNull(RecChangeData!审核人) Then
                    lngState = 3
                Else
                    If Trim(RecChangeData!审核人) = "" Then lngState = 3
                End If
            Else
                lngState = 1
            End If
            
            RecChangeData!执行状态 = lngState                        '缺省为发药
            RecChangeData.Update
            If err <> 0 Then GoTo ErrHand
            .MoveNext
        Loop
    End With
    
    If err <> 0 Then
ErrHand:
        MsgBox "产生内部记录集时，发生不可预知的错误！", vbInformation, gstrSysName
        InitRec
        Exit Function
    End If
    
    ProduceInsideRecordset = True
End Function

Private Function ProduceInsideSendedRecordset() As Boolean
    Dim ArrayPhysic
    Dim IntArray As Integer
    Dim dblSumSended As Double '已发数量
    
    '--产生内部记录集(已发)--
'    on error Resume Next
    err = 0
    ProduceInsideSendedRecordset = False
    
    With RecBillData
        Do While Not .EOF
            RecChangeSendedData.AddNew
            RecChangeSendedData!Id = !Id
            RecChangeSendedData!药品ID = !药品ID
            RecChangeSendedData!位置 = .AbsolutePosition
            RecChangeSendedData!科室 = !科室
            RecChangeSendedData!类型 = IIf(NVL(!医嘱序号, 0) = 0, IIf(!门诊标志 = 1 Or !门诊标志 = 4, "门诊记帐单", IIf(!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(!扣率) = True, "住院记帐单", IIf(!扣率 Like "0*", "长嘱", IIf(!扣率 Like "1*", "临嘱", "记帐表"))))
            RecChangeSendedData!执行状态 = 1                        '缺省为不处理
            RecChangeSendedData!NO = !NO
            RecChangeSendedData!单据 = !单据
            RecChangeSendedData!序号 = !序号
            RecChangeSendedData!病人ID = !病人ID
            RecChangeSendedData!床号 = !床号
            RecChangeSendedData!姓名 = IIf(IsNull(!姓名), "", !姓名)
            RecChangeSendedData!住院号 = NVL(!标识号)
            RecChangeSendedData!品名 = !品名
            RecChangeSendedData!其它名 = !其它名
            RecChangeSendedData!英文名 = !英文名
            RecChangeSendedData!规格 = IIf(IsNull(!规格), "", !规格)
            RecChangeSendedData!产地 = IIf(IsNull(!产地), "", !产地)
            RecChangeSendedData!毒理分类 = NVL(!毒理分类)
            RecChangeSendedData!分批 = IIf(IsNull(!分批), 0, !分批)
            RecChangeSendedData!批次 = IIf(IsNull(!批次), 0, !批次)
            RecChangeSendedData!批号 = IIf(IsNull(!批号), "", !批号)
            RecChangeSendedData!效期 = IIf(IsNull(!效期), "", !效期)
            RecChangeSendedData!付 = IIf(IsNull(!付), 1, !付)
            RecChangeSendedData!数量 = FormatEx(IIf(IsNull(!数量), 1, !数量) / !包装, 5) & !单位
            If Chk清单.Value = 0 Or !可操作 <> 1 Then
                RecChangeSendedData!已退数 = FormatEx(IIf(IsNull(!已退数量), 1, !已退数量) / !包装, 5)
                RecChangeSendedData!准退数 = FormatEx(IIf(IsNull(!准退数), 1, !准退数) / !包装, 5)
                RecChangeSendedData!退药数 = FormatEx(IIf(IsNull(!准退数), 1, !准退数) / !包装, 5)
            Else
                dblSumSended = GetSumSended(!单据, !NO, !药品ID, !序号)
                RecChangeSendedData!已退数 = FormatEx((IIf(IsNull(!数量), 1, !数量) - dblSumSended) / !包装, 5)
                RecChangeSendedData!准退数 = FormatEx(dblSumSended / !包装, 5)
                RecChangeSendedData!退药数 = FormatEx(dblSumSended / !包装, 5)
            End If
            RecChangeSendedData!单位 = !单位
            RecChangeSendedData!单价 = FormatEx(!单价 * !包装, 5)
            RecChangeSendedData!金额 = !金额
            RecChangeSendedData!单量 = IIf(IsNull(!单量), "", FormatEx(!单量, 5) & NVL(!计算单位))
            RecChangeSendedData!单量单位 = NVL(!计算单位)
            RecChangeSendedData!频次 = IIf(IsNull(!频次), "", !频次)
            RecChangeSendedData!用法 = IIf(IsNull(!用法), "", !用法)
            RecChangeSendedData!说明 = IIf(IsNull(!说明), "", !说明)
            RecChangeSendedData!操作员 = IIf(IsNull(!审核人), "", !审核人)
            RecChangeSendedData!发药时间 = IIf(IsNull(!发药时间), "", !发药时间)
            If Val(!转出) = 1 Then
                RecChangeSendedData!可操作 = -1
            Else
                RecChangeSendedData!可操作 = IIf(IsNull(!可操作), 0, !可操作)
            End If
            RecChangeSendedData!审查结果 = !审查结果
            RecChangeSendedData!医嘱id = !医嘱id
            RecChangeSendedData!领药人 = !领药人
            RecChangeSendedData!实际数量 = !准退数
            RecChangeSendedData!库房货位 = IIf(IsNull(!库房货位), "", !库房货位)
            RecChangeSendedData!转出 = Val(!转出)
            If Chk清单.Value = 1 Then
                RecChangeSendedData!相关ID = 0
            Else
                RecChangeSendedData!相关ID = IIf(IsNull(!相关ID), 0, !相关ID)
            End If
            
            .MoveNext
        Loop
    End With
    
    If err <> 0 Then
        MsgBox "产生内部记录集时，发生不可预知的错误！", vbInformation, gstrSysName
        Call InitRec
        Exit Function
    End If
    ProduceInsideSendedRecordset = True
End Function

Private Function RefreshDataBaseOnPage(Optional ByVal BlnRefsh As Boolean = False) As Boolean
    Dim lngRows As Long
    Dim strCaption As String
    '--根据用户选择的页面产生显示数据--
    On Error Resume Next
    err = 0
    RefreshDataBaseOnPage = False
    
    '清空将要显示的控件中的内容
    If InStr(1, "1,2,3", TabShow.Tab) <> 0 Then ClearCons
    If Bln刷新未发药清单 And (TabShow.Tab = 0 Or TabShow.Tab = 1) Then        '做一些准备
        If Bln检测库存 Then
            '检测库存够否
            Call CheckStock
        End If
        Bln检测库存 = False
    End If
    
    '根据页面初始化下拉框
    stbThis.Panels(2).Text = ""
    Select Case TabShow.Tab
    Case 0
        If Bln刷新未发药清单 Then
            Call ClearCons
            If LoadDataInBill未发药清单 = False Then GoTo ErrHand
            Call SetGroup(Bill未发药清单, True)
        End If
        lngRows = lng未发药记录
        strCaption = "条未发药品记录"
    Case 1
        If LoadDataInBill汇总清单 = False Then GoTo ErrHand
'        lngRows = IIf(Bill未发药清单.TextMatrix(Bill未发药清单.Rows - 1, 0) = "", Bill未发药清单.Rows - 2, Bill未发药清单.Rows - 1)
'        strCaption = "条未发药品记录"
        lngRows = IIf(Bill汇总发药.TextMatrix(Bill汇总发药.rows - 1, 0) = "", 0, lng汇总清单行数)
        strCaption = "条药品汇总记录"
    Case 2
        If LoadDataInBill缺药清单 = False Then GoTo ErrHand
        lngRows = IIf(Bill缺药清单.TextMatrix(Bill缺药清单.rows - 1, 0) = "", Bill缺药清单.rows - 2, Bill缺药清单.rows - 1)
        strCaption = "条缺药记录"
    Case 3
        If LoadDataInBill拒发清单 = False Then GoTo ErrHand
        lngRows = IIf(Bill拒发药清单.TextMatrix(Bill拒发药清单.rows - 1, 0) = "", Bill拒发药清单.rows - 2, Bill拒发药清单.rows - 1)
        strCaption = "条拒发药品记录"
    Case 4
        '手工预填充
        lngRows = IIf(Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 0) = "", Bill已发药清单.rows - 2, Bill已发药清单.rows - 1)
        strCaption = "条已发药品记录"
    End Select
    
    If err <> 0 Then
ErrHand:
        MsgBox "显示[" & TabShow.TabCaption(TabShow.Tab) & "]页面的数据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    If lngRows <> 0 Then stbThis.Panels(2).Text = "当前共有" & lngRows & strCaption
    If mlng汇总发药号 > 0 Then
        stbThis.Panels(2).Text = stbThis.Panels(2).Text & "[上次发药号：" & mlng汇总发药号 & "]"
    End If
    RefreshDataBaseOnPage = True
End Function

Private Function ClearCons()
    '--根据页面清空相关控件的显示内容--
    Select Case TabShow.Tab
    Case 0
        Call ClearBill(Bill未发药清单)
    Case 1
        Call ClearBill(Bill汇总发药)
        Call ClearBill(Bill退药销帐)
        Bill退药销帐.Visible = False
        Bill汇总发药.Height = TabShow.Height - TabShow.TabHeight - 120
    Case 2
        Call ClearBill(Bill缺药清单)
    Case 3
        '只清除暂选择还未拒发的药品记录
        Dim i As Integer, j As Integer
        With Bill拒发药清单
            For i = 1 To .rows - 1
                If Trim(.TextMatrix(i, 1)) = "" Then
                    If i = .rows - 1 Then
                        For j = 0 To .Cols - 1
                            .TextMatrix(i, j) = ""
                        Next
                    Else
                        .RemoveItem i: i = i - 1
                    End If
                End If
            Next
        End With
    Case 4
        '手工处理
        Call ClearBill(Bill已发药清单)
    End Select
End Function

Private Function LoadDataInBill未发药清单() As Boolean
    Dim blnEnable As Boolean, lngRow As Long, intCol As Long
    Dim strCompare As String, strColumn As String, strValue As String
    Dim dbl合计金额 As Double, dbl小计金额 As Double
    Dim lngColor As Long
    Dim strOrder As String
    
    '--填充未发药清单--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill未发药清单 = False
    
    With Bill未发药清单
        .MousePointer = 11
        .Redraw = False
    End With
    
    dbl小计金额 = 0: dbl合计金额 = 0
    If InStr(1, str排序_未发药, strAsc) <> 0 Then
        strColumn = Mid(str排序_未发药, 1, InStr(1, str排序_未发药, strAsc) - 1)
    Else
        strColumn = Mid(str排序_未发药, 1, InStr(1, str排序_未发药, strDesc) - 1)
    End If
    strColumn = Trim(strColumn)
    
    '根据排序列进行排序
    With RecChangeData
        lng未发药记录 = .RecordCount
        If .RecordCount <> 0 Then .MoveFirst
        strOrder = GetOrder(str排序_未发药)
        '如果按NO排序，则同时按相关ID排序，便于设置分组
        strOrder = IIf(InStr(strOrder, "NO") > 0, strOrder & ",相关ID" & IIf(InStr(strOrder, " ASC") > 0, " ASC", " DESC"), strOrder)
        .Sort = strOrder
        Do While Not .EOF
'            Str执行状态 = IIf(!执行状态 = 0, "缺药", IIf(!执行状态 = 1, "发药", IIf(!执行状态 = 2, "拒发", "不处理")))
            If !说明 <> "拒发" Then
                blnEnable = True
                Bill未发药清单.MergeRow(Bill未发药清单.rows - 1) = False
                
                strValue = IIf(IsNull(.Fields(strColumn).Value), "", .Fields(strColumn).Value)
                If strCompare <> strValue And strCompare <> "" Then
                    '增加合计栏
                    Call AddCollect(dbl小计金额, "小计")
                    dbl小计金额 = 0
                End If
                
                '赋值
                strCompare = IIf(IsNull(.Fields(strColumn).Value), "", .Fields(strColumn).Value)
                
                '填充数据
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.科室) = !科室
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.开单医生) = IIf(IsNull(!开单医生), "", !开单医生)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.状态) = !状态
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.类型) = !类型
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.NO) = !NO
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.记帐员) = !记帐员
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.床号) = IIf(IsNull(!床号), "", !床号)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.姓名) = !姓名
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.住院号) = !住院号
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.药品名称) = !品名
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.其它名) = IIf(IsNull(!其它名), "", !其它名)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.英文名) = IIf(IsNull(!英文名), "", !英文名)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.规格) = IIf(IsNull(!规格), "", !规格)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.产地) = !产地
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.批号) = !批号
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.付) = !付
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.数量) = !数量
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.单价) = !单价
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.金额) = Format(!金额, "#####0.00;-#####0.00; ;")
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.单量) = !单量
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.频次) = !频次
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.用法) = !用法
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.记帐时间) = Format(!记帐时间, "yyyy-MM-dd HH:mm:ss")
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.说明) = !说明
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.单据) = !单据
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.医嘱id) = !医嘱id
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.退药人) = IIf(IsNull(!退药人), "", !退药人)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.库房货位) = IIf(IsNull(!库房货位), "", !库房货位)
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.相关ID) = !相关ID
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.药品ID) = !药品ID
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.领药部门) = !领药部门
                Bill未发药清单.TextMatrix(Bill未发药清单.rows - 1, 列名_未发药清单.领药部门id) = !领药部门id
                                
                Bill未发药清单.ColWidth(列名_未发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
                Bill未发药清单.ColWidth(列名_未发药清单.单据) = 0
                Bill未发药清单.ColWidth(列名_未发药清单.医嘱id) = 0
                
                If !审查结果 <> -1 Then
                    BlnEnterCell = False
                    Bill未发药清单.Row = Bill未发药清单.rows - 1
                    Bill未发药清单.Col = 0
                    Set Bill未发药清单.CellPicture = imgPass.ListImages(Val(!审查结果) + 1).Picture
                    Bill未发药清单.CellPictureAlignment = 4
                    BlnEnterCell = True
                End If

                '特殊药品粗体显示
                BlnEnterCell = False
                If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(!毒理分类)) > 0 And NVL(!毒理分类) <> "" Then
                    Bill未发药清单.Col = 列名_未发药清单.药品名称
                    Bill未发药清单.Row = Bill未发药清单.rows - 1
                    Bill未发药清单.CellFontBold = True
                End If

                If mbln药品储备 = True Then
                    If !库存下限 = 0 Then
                        Bill未发药清单.Row = Bill未发药清单.rows - 1
                        For intCol = 0 To Bill未发药清单.Cols - 1
                            Bill未发药清单.Col = intCol
                            Bill未发药清单.CellForeColor = mlng紫色
                        Next
                        Bill未发药清单.RowData(Bill未发药清单.Row) = mlng紫色
                    End If
                End If
                
                '设置发药状态的背景色
                lngColor = IIf(!状态 = "发药", glngSendBlkColor, glngOtherBlkColor)
                Bill未发药清单.Row = Bill未发药清单.rows - 1
                For intCol = 0 To Bill未发药清单.Cols - 1
                    Bill未发药清单.Col = intCol
                    Bill未发药清单.CellBackColor = lngColor
                Next
                
                BlnEnterCell = True
                
                dbl小计金额 = dbl小计金额 + Val(!金额)
                dbl合计金额 = dbl合计金额 + Val(!金额)
                !位置 = Bill未发药清单.rows - 1
                .Update
                Bill未发药清单.rows = Bill未发药清单.rows + 1
                
                Bill未发药清单.ColAlignment(列名_未发药清单.药品名称) = 1
                
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then
            If strCompare <> "" Then
                '增加合计栏
                Call AddCollect(dbl小计金额, "小计")
                Call AddCollect(dbl合计金额)
            End If
            .MoveFirst
        End If
        
        '合并
        Bill未发药清单.MergeCells = flexMergeFree
        For lngRow = 0 To Bill未发药清单.rows - 1
            If InStr(1, "小计,合计", Bill未发药清单.TextMatrix(lngRow, 列名_未发药清单.科室)) <> 0 Then
                Bill未发药清单.MergeRow(lngRow) = True
            End If
        Next
        
        Call SetMenu(blnEnable)
    End With
    
    With Bill未发药清单
        .MousePointer = 0
        .Redraw = True
        .Row = 1
        .Col = 1
    End With
    
    If err <> 0 Then Exit Function
    
    LoadDataInBill未发药清单 = True
    Bln刷新未发药清单 = False
End Function

Private Function LoadDataInBill汇总清单() As Boolean
    Dim LngFindPhysicID As Long, strPartName As String
    Dim LngLocate As Long, blnEnable As Boolean
    Dim dbl合计金额 As Double, dbl科室合计 As Double
    Dim lng批次 As Long
    Dim n As Integer
    '--填充汇总清单--
'    on error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill汇总清单 = False
    Bill汇总发药.Redraw = False
    
    LngFindPhysicID = 0
    lng批次 = 0
    strPartName = ""
    dbl科室合计 = 0
    dbl合计金额 = 0
    
    With RecChangeData
        If .RecordCount = 0 Then
            LoadDataInBill汇总清单 = True
            Call SetMenu(blnEnable)
            Bill汇总发药.Redraw = True
            Exit Function
        End If
        
        .MoveFirst
        
        If Chk清单.Value = 0 Then   '按药品名称汇总
            .Sort = IIf(Lng汇总显示 = 1, "领药部门 Asc,", "") & "药品ID Asc"
        Else    '按药品批次汇总
            .Sort = IIf(Lng汇总显示 = 1, "领药部门 Asc,", "") & "药品ID Asc" & ",批次 Asc"
        End If
        '手工处理后显示
        Do While Not .EOF
            If Lng汇总显示 = 1 Then
                If !执行状态 = 1 And !领药部门 <> strPartName And CheckGroupSend(!相关ID) = True Then
                    LngLocate = !位置
                    blnEnable = True
                    strPartName = !领药部门
                    If LngFindPhysicID <> 0 And IIf(Chk清单.Value = 0, True, lng批次 <> 0) Then
                        Bill汇总发药.rows = Bill汇总发药.rows + 1
                        Call AddCollect(dbl科室合计, "小计")
                        dbl科室合计 = 0
                        LngFindPhysicID = 0
                        lng批次 = 0
                    End If
                End If
            End If
            lng汇总清单行数 = 0
            If !执行状态 = 1 And (!药品ID <> LngFindPhysicID Or IIf(Chk清单.Value = 0, False, !批次 <> lng批次)) And CheckGroupSend(!相关ID) = True Then '只汇总发药的记录
                LngLocate = !位置
                blnEnable = True
                With Bill汇总发药
                    If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                    .MergeRow(.rows - 1) = False
                    If Lng汇总显示 = 1 Then
                        .Row = .rows - 1: .Col = 1: .CellAlignment = 1
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.科室) = RecChangeData!科室
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.药品名称) = RecChangeData!品名
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.规格) = RecChangeData!规格
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.产地) = RecChangeData!产地
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.批号) = RecChangeData!批号
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.应发数量) = FormatEx(RecChangeData!实际数量 * RecChangeData!付, 5)
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.单位) = Right(RecChangeData!数量, 1)
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.单价) = FormatEx(RecChangeData!单价, 5)
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.金额) = Format(RecChangeData!金额, "#####0.00;-#####0.00; ;")
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.批次) = RecChangeData!批次
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.科室ID) = RecChangeData!科室ID
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.药品ID) = RecChangeData!药品ID
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.领药部门) = RecChangeData!领药部门
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.领药部门id) = RecChangeData!领药部门id
                        .TextMatrix(.rows - 1, 列名_科室汇总清单.留存数量) = FormatEx(RecChangeData!留存数量, 5)
                        
                        If mbln汇总发药 = True Then
                            .TextMatrix(.rows - 1, 列名_科室汇总清单.销帐数量) = FormatEx(Get销帐数量(RecChangeData!领药部门id, RecChangeData!药品ID), 5)
                        End If
                        
                    Else
                        .Row = .rows - 1: .Col = 0: .CellAlignment = 1
                        .TextMatrix(.rows - 1, 列名_汇总清单.药品名称) = RecChangeData!品名
                        .TextMatrix(.rows - 1, 列名_汇总清单.规格) = RecChangeData!规格
                        .TextMatrix(.rows - 1, 列名_汇总清单.产地) = RecChangeData!产地
                        .TextMatrix(.rows - 1, 列名_汇总清单.批号) = RecChangeData!批号
                        .TextMatrix(.rows - 1, 列名_汇总清单.数量) = FormatEx(RecChangeData!实际数量 * RecChangeData!付, 5)
                        .TextMatrix(.rows - 1, 列名_汇总清单.单位) = Right(RecChangeData!数量, 1)
                        .TextMatrix(.rows - 1, 列名_汇总清单.单价) = FormatEx(RecChangeData!单价, 5)
                        .TextMatrix(.rows - 1, 列名_汇总清单.金额) = Format(RecChangeData!金额, "#####0.00;-#####0.00; ;")
                    End If
                    
                    '特殊药品粗体显示
                    If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(RecChangeData!毒理分类)) > 0 And NVL(RecChangeData!毒理分类) <> "" Then
                        .Row = .rows - 1
                        .Col = IIf(Lng汇总显示 = 1, 列名_科室汇总清单.药品名称, 列名_汇总清单.药品名称)
                        .CellFontBold = True
                    End If
                    
                    dbl科室合计 = dbl科室合计 + Val(RecChangeData!金额)
                    dbl合计金额 = dbl合计金额 + Val(RecChangeData!金额)
                    
                End With
                LngFindPhysicID = !药品ID
                lng批次 = !批次
                
                '汇总
                If Not .EOF Then .MoveNext
                Do While Not .EOF
                    If LngFindPhysicID = !药品ID And IIf(Chk清单.Value = 0, True, lng批次 = !批次) And !执行状态 = 1 And CheckGroupSend(!相关ID) = True Then
                        If strPartName <> !领药部门 And Lng汇总显示 = 1 Then Exit Do
                        With Bill汇总发药
                            If Lng汇总显示 = 0 Then
                                .TextMatrix(.rows - 1, 列名_汇总清单.数量) = FormatEx(Val(.TextMatrix(.rows - 1, 列名_汇总清单.数量)) + (RecChangeData!实际数量 * RecChangeData!付), 5)
                                .TextMatrix(.rows - 1, 列名_汇总清单.金额) = Format(Val(.TextMatrix(.rows - 1, 列名_汇总清单.金额)) + Val(RecChangeData!金额), "#####0.00;-#####0.00; ;")
                            Else
                                .TextMatrix(.rows - 1, 列名_科室汇总清单.应发数量) = FormatEx(Val(.TextMatrix(.rows - 1, 列名_科室汇总清单.应发数量)) + (RecChangeData!实际数量 * RecChangeData!付), 5)
                                .TextMatrix(.rows - 1, 列名_科室汇总清单.金额) = Format(Val(.TextMatrix(.rows - 1, 列名_科室汇总清单.金额)) + Val(RecChangeData!金额), "#####0.00;-#####0.00; ;")
                            End If
                            dbl科室合计 = dbl科室合计 + Val(RecChangeData!金额)
                            dbl合计金额 = dbl合计金额 + Val(RecChangeData!金额)
                        End With
                    End If
                    .MoveNext
                Loop
                .MoveFirst
                .Find "位置=" & LngLocate
            End If
            
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        
        '统计实际发药数量
        If Lng汇总显示 = 1 Then
            For n = 1 To Bill汇总发药.rows - 1
                With Bill汇总发药
                    If .TextMatrix(n, 0) <> "小计" Then
                        '应发数量小于了销帐数量，实发为负数（表示科室将实物退药），留存数为0
                        If Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)) < 0 Then
                            .TextMatrix(n, 列名_科室汇总清单.实发数量) = FormatEx(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)), 5)
                            .TextMatrix(n, 列名_科室汇总清单.留存数量) = 0
                        Else
                            '如果留存数量为0，实发数量则按参数规则分配，并根据实际应发数量计算（实际应发＝应发数量－销帐数量）
                            If Val(.TextMatrix(n, 列名_科室汇总清单.留存数量)) = 0 Then
                                If int发药规则 = 0 Then
                                    .TextMatrix(n, 列名_科室汇总清单.实发数量) = FormatEx(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)), 5)
                                ElseIf int发药规则 = 1 Then
                                    .TextMatrix(n, 列名_科室汇总清单.实发数量) = 0
                                Else
                                    .TextMatrix(n, 列名_科室汇总清单.实发数量) = FormatEx(Int(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量))), 5)
                                End If
                                .TextMatrix(n, 列名_科室汇总清单.留存数量) = FormatEx(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.实发数量)), 5)
                            Else
                            '如果留存数量不为0（从药品留存计划取值），根据实际应发数量计算（实际应发＝应发数量－销帐数量）
                                If Val(.TextMatrix(n, 列名_科室汇总清单.留存数量)) > Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)) Then
                                    '留存数量大于了实际应发数量，则留存数量＝实际应发数量
                                    .TextMatrix(n, 列名_科室汇总清单.留存数量) = FormatEx(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量)), 5)
                                End If
                                
                                '实发数量＝应发数量－留存数量－销帐数量
                                .TextMatrix(n, 列名_科室汇总清单.实发数量) = FormatEx(Int(Val(.TextMatrix(n, 列名_科室汇总清单.应发数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.留存数量)) - Val(.TextMatrix(n, 列名_科室汇总清单.销帐数量))), 5)
                            End If
                        End If
                        
                        .Row = n
                        .Col = 列名_科室汇总清单.实发数量
                        .CellFontBold = True
                        If Val(.TextMatrix(n, 列名_科室汇总清单.实发数量)) < 0 Then
                            .CellForeColor = vbRed
                        ElseIf Val(.TextMatrix(n, 列名_科室汇总清单.实发数量)) > 0 Then
                            .CellForeColor = vbBlue
                        End If
                    End If
                End With
            Next
        End If
                
        lng汇总清单行数 = Bill汇总发药.rows - 1
        If Lng汇总显示 = 1 And dbl科室合计 <> 0 Then
            Bill汇总发药.rows = Bill汇总发药.rows + 1
            Call AddCollect(dbl科室合计, "小计")
        End If
        Call SetMenu(blnEnable)
        
        .Sort = "NO Asc"
        
        '如果不是按批次汇总发药，就不显示批号列
        With Bill汇总发药
            .ColWidth(IIf(Lng汇总显示 = 1, 列名_科室汇总清单.批号, 列名_汇总清单.批号)) = IIf(Chk清单.Value = 1, 1200, 0)
        End With
        
    End With
    
    With Bill汇总发药
        If .TextMatrix(.rows - 1, 0) <> "" Then .rows = .rows + 1
        If dbl合计金额 <> 0 Then Call AddCollect(dbl合计金额)
        For LngLocate = 1 To .rows - 1
            If InStr(1, "小计,合计", .TextMatrix(LngLocate, 0)) <> 0 Then
                .MergeCells = flexMergeFree
                .MergeRow(LngLocate) = True
            End If
        Next
        .Row = 1: .Col = 0
    End With
    
    
    Bill汇总发药.Redraw = True
    If err <> 0 Then Exit Function
    LoadDataInBill汇总清单 = True
End Function

Private Function LoadDataInBill缺药清单() As Boolean
    Dim LngRecords As Long, blnEnable As Boolean
    '--填充缺药清单--
    Debug.Print Now
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill缺药清单 = False
    
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
        Else
            LoadDataInBill缺药清单 = True: Call SetMenu(blnEnable): Exit Function
        End If
        
        '手工处理后显示
        LngRecords = 0
        Do While Not .EOF
            If !执行状态 = 0 Then   '只显示缺药记录
                If Not IsNull(!配药人) Then
                    If !配药人 <> "部门发药" Then
                        blnEnable = True
                        With Bill缺药清单
                            If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                            .TextMatrix(.rows - 1, 0) = RecChangeData!科室
                            .TextMatrix(.rows - 1, 1) = RecChangeData!NO
                            .TextMatrix(.rows - 1, 2) = RecChangeData!类型
                            .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecChangeData!床号), "", RecChangeData!床号)
                            .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecChangeData!姓名), "", RecChangeData!姓名)
                            .TextMatrix(.rows - 1, 5) = RecChangeData!品名
                            .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecChangeData!规格), "", RecChangeData!规格)
                            .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecChangeData!产地), "", RecChangeData!产地)
                            .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecChangeData!批号), "", RecChangeData!批号)
                            .TextMatrix(.rows - 1, 9) = FormatEx(RecChangeData!实际数量 * RecChangeData!付, 5) & Right(RecChangeData!数量, 1)
                            .TextMatrix(.rows - 1, 10) = FormatEx(RecChangeData!单价, 5)
                            .TextMatrix(.rows - 1, 11) = Format(RecChangeData!金额, "#####0.00;-#####0.00; ;")
                            
                            '特殊药品粗体显示
                            If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(RecChangeData!毒理分类)) > 0 And NVL(RecChangeData!毒理分类) <> "" Then
                                .Row = .rows - 1: .Col = 5
                                .CellFontBold = True
                            End If
                        End With
                        LngRecords = LngRecords + 1
                    End If
                End If
            End If
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        Call SetMenu(blnEnable)
    End With
    
    If err <> 0 Then Exit Function
    LoadDataInBill缺药清单 = True
End Function

Private Function LoadDataInBill拒发清单() As Boolean
    Dim lngRow As Long, blnEnable As Boolean
    
    '--填充拒发清单--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill拒发清单 = False
    
    '装入设定为拒发的清单(未更新数据库)
    With RecChangeData
        If .RecordCount <> 0 Then
            .MoveFirst
        Else
            LoadDataInBill拒发清单 = True: Call SetMenu(blnEnable): Exit Function
        End If
        
        '手工处理后显示
        Do While Not .EOF
            If !执行状态 = 2 Then   '只显示拒发记录
                blnEnable = True
                With Bill拒发药清单
                    If Trim(.TextMatrix(.rows - 1, 1)) <> "" Then .rows = .rows + 1
                    .TextMatrix(.rows - 1, 0) = RecChangeData!科室
                    .TextMatrix(.rows - 1, 1) = ""
                    .TextMatrix(.rows - 1, 2) = RecChangeData!NO
                    .TextMatrix(.rows - 1, 3) = RecChangeData!类型
                    .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecChangeData!床号), "", RecChangeData!床号)
                    .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecChangeData!姓名), "", RecChangeData!姓名)
                    .TextMatrix(.rows - 1, 6) = RecChangeData!品名
                    .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecChangeData!规格), "", RecChangeData!规格)
                    .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecChangeData!产地), "", RecChangeData!产地)
                    .TextMatrix(.rows - 1, 9) = IIf(IsNull(RecChangeData!批号), "", RecChangeData!批号)
                    .TextMatrix(.rows - 1, 10) = FormatEx(RecChangeData!实际数量 * RecChangeData!付, 5) & Right(RecChangeData!数量, 1)
                    .TextMatrix(.rows - 1, 11) = FormatEx(RecChangeData!单价, 5)
                    .TextMatrix(.rows - 1, 12) = Format(RecChangeData!金额, "#####0.00;-#####0.00; ;")
                    .RowData(.rows - 1) = 0
                    
                    '特殊药品粗体显示
                    If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(RecChangeData!毒理分类)) > 0 And NVL(RecChangeData!毒理分类) <> "" Then
                        .Row = .rows - 1: .Col = 6
                        .CellFontBold = True
                    End If
                    
                    .rows = .rows + 1
                End With
            End If
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        Call SetMenu(blnEnable)
    End With
    lngRow = IIf(Trim(Bill拒发药清单.TextMatrix(Bill拒发药清单.rows - 1, 0)) <> "", Bill拒发药清单.rows - 1, Bill拒发药清单.rows - 2)

    If err <> 0 Then Exit Function
    LoadDataInBill拒发清单 = True
End Function

Private Function LoadDataInBill已发药清单() As Boolean
    Dim Str执行状态 As String, blnEnable As Boolean, lngColor As Long, intCol As Integer
    
    '--填充已发药清单--
    On Error Resume Next
    err = 0
    blnEnable = False
    LoadDataInBill已发药清单 = False
    
    With Bill已发药清单
        .MousePointer = 11
        .Redraw = False
    End With
    
    '根据排序列进行排序
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
'        If Chk清单.Value = 0 Then .Sort = GetOrder(str排序_发退药)
        .Sort = GetOrder(str排序_发退药)
        Do While Not .EOF
            blnEnable = True
            
            '检查该明细是否已转出
            If Val(!转出) = 0 Then
                Str执行状态 = IIf(!执行状态 = 3, "退药", "不处理")
            Else
                Str执行状态 = "不处理"
            End If
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.科室) = !科室
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.状态) = Str执行状态
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.类型) = !类型
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.NO) = !NO
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.床号) = IIf(IsNull(!床号), "", !床号)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.姓名) = !姓名
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.住院号) = !住院号
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.药品名称) = !品名
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.其它名) = IIf(IsNull(!其它名), "", !其它名)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.英文名) = IIf(IsNull(!英文名), "", !英文名)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.规格) = IIf(IsNull(!规格), "", !规格)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.产地) = IIf(IsNull(!产地), "", !产地)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.批号) = IIf(IsNull(!批号), "", !批号)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.付) = !付
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.数量) = !数量
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.已退数) = !已退数
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.准退数) = !准退数
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.退药数) = ""
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.单价) = !单价
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.金额) = Format(!金额, "#####0.00;-#####0.00; ;")
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.单量) = !单量
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.频次) = !频次
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.用法) = !用法
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.操作员) = !操作员
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.发药时间) = !发药时间
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.单据) = !单据
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.医嘱id) = !医嘱id
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.领药人) = IIf(IsNull(!领药人), "", !领药人)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.库房货位) = IIf(IsNull(!库房货位), "", !库房货位)
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.相关ID) = !相关ID
            Bill已发药清单.TextMatrix(Bill已发药清单.rows - 1, 列名_已发药清单.药品ID) = !药品ID
            
            Bill已发药清单.ColWidth(列名_已发药清单.审查结果) = IIf(Not mblnStarPass, 0, 240)
            Bill已发药清单.ColWidth(列名_已发药清单.单据) = 0
            Bill已发药清单.ColWidth(列名_已发药清单.医嘱id) = 0
            
            If !审查结果 <> -1 Then
                BlnEnterCell = False
                Bill已发药清单.Row = Bill已发药清单.rows - 1
                Bill已发药清单.Col = 0
                Set Bill已发药清单.CellPicture = imgPass.ListImages(Val(!审查结果) + 1).Picture
                Bill已发药清单.CellPictureAlignment = 4
                BlnEnterCell = True
            End If
            
            !位置 = Bill已发药清单.rows - 1
            .Update
            
            '根据记录状态的不同，进行着色（可操作：1-原始记录；2-发药；3-退药；-1-数据已转出，不允许操作）
            lngColor = IIf(!可操作 = 2, glng发药, IIf(!可操作 = 3, glng退药, glng正常))
            Bill已发药清单.Row = Bill已发药清单.rows - 1
            For intCol = 0 To Bill已发药清单.Cols - 1
                Bill已发药清单.Col = intCol
                Bill已发药清单.CellForeColor = lngColor
            Next
            Bill未发药清单.RowData(Bill未发药清单.Row) = lngColor
            
            '特殊药品粗体显示
            If InStr(";毒性药;麻醉药;精神I类;精神II类;", NVL(!毒理分类)) > 0 And NVL(!毒理分类) <> "" Then
                Bill已发药清单.Col = 列名_已发药清单.药品名称
                Bill已发药清单.CellFontBold = True
            End If
            
            Bill已发药清单.rows = Bill已发药清单.rows + 1
            
            Bill已发药清单.ColAlignment(列名_已发药清单.药品名称) = 1
            
            
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        Me.stbThis.Panels(2) = "当前共有" & .RecordCount & "条已发药品记录"
        Call SetMenu(blnEnable)
    End With
    
    With Bill已发药清单
        .MousePointer = 0
        .Redraw = True
        .Row = 1
        .Col = 1
    End With
    
    Me.MousePointer = 0
    If err <> 0 Then Exit Function
    LoadDataInBill已发药清单 = True
End Function

Private Function CheckStock(Optional ByVal lng药品ID As Long = 0)
    Dim RecCheckStock As New adodb.Recordset            '检测库存记录集
    Dim dblStock As Double                              '库存数
    Dim LngPhysicID As Long                             '当前药品ID
    Dim DblCompare As Double                            '用于比较的数量
    Dim Str执行状态 As String
    Dim lngState As Long, LngLocate As Long
    Dim BlnSet As Boolean, BlnRestore As Boolean, blnEof As Boolean, blnGetData As Boolean
    Dim strSubSql As String
    Dim rsStock As adodb.Recordset
    Dim blnFlag As Boolean
    Dim intCol As Integer
    
    '--根据库存状态显示缺药并更新记录集--
    '有可能几条记录都发同一种规格的药品，因此需统计出从开始一直统计到当前位置的总数量
    '如果供刷新用记录不为空、有对应记录且执行状态不为发药或缺药，则不必检查库存，直接恢复上次设定的执行状态
    '--Modified by ZYB 20021009
    '--参数lng药品ID用于：当某笔记录的状态发生改变时，只判断使用该药品ID的记录的库存
    
    On Error GoTo errHandle
    '单位设置
    Select Case strUnit
    Case "售价单位"
        strSubSql = "/1"
    Case "门诊单位"
        strSubSql = "/Decode(Nvl(门诊包装,0),0,1,门诊包装)"
    Case "住院单位"
        strSubSql = "/Decode(Nvl(住院包装,0),0,1,住院包装)"
    Case "药库单位"
        strSubSql = "/Decode(Nvl(药库包装,0),0,1,药库包装)"
    End Select
    
    Set rsStock = New adodb.Recordset
    With rsStock
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18
        .Fields.Append "批次", adDouble, 18
        .Fields.Append "变价", adDouble, 18
        .Fields.Append "数量", adDouble, 18
        .Fields.Append "序号", adDouble, 5
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    LngPhysicID = 0
    BlnSet = (RecRefreshCompare.RecordCount <> 0)
    
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If lng药品ID <> 0 Then
            .Filter = "药品ID=" & lng药品ID
            If .RecordCount = 0 Then .Filter = 0: Exit Function
        End If
        
        With Bill未发药清单
            .Redraw = False
        End With
        
        Do While Not .EOF
            If BlnSet Then  '检测是否有对应记录
                With RecRefreshCompare
                    .MoveFirst
                    .Find "ID=" & RecChangeData!Id
                    BlnRestore = (.EOF Xor True)
                    If BlnRestore Then BlnRestore = (!执行状态 >= 2)
                End With
            End If
            
            If BlnSet And BlnRestore Then
                !执行状态 = RecRefreshCompare!执行状态
                .Update
            Else
                If !执行状态 <= 1 Then  '检查缺药与发药的记录
                    blnGetData = False
                    If LngPhysicID <> !药品ID Then
                        LngPhysicID = !药品ID
                        blnEof = True
                        blnGetData = True
                    End If
                    With rsStock
                        If .RecordCount <> 0 Then
                            .Filter = "药品ID=" & LngPhysicID & " And 批次=" & IIf(IsNull(RecChangeData!批次), 0, RecChangeData!批次)
                            blnEof = .EOF
                            blnGetData = blnEof
                            If .RecordCount <> 0 Then LngLocate = !序号
                            .Filter = 0
                        End If
                    End With
                    
                    If blnGetData Then
                        If blnEof Then
                            LngLocate = rsStock.RecordCount + 1

                            gstrSQL = " Select nvl(F.是否变价,0) 变价,nvl(A.实际数量,0)" & strSubSql & " 数量" & _
                                         " From 药品规格 B,收费项目目录 F," & _
                                         "      (Select * From 药品库存 " & _
                                         "      Where 性质=1 And 库房ID=[1] And 药品ID=[2] And nvl(批次,0)=[3]) A" & _
                                         " Where B.药品ID=F.ID And A.药品ID(+)=B.药品ID And B.药品ID=[2]"
                            Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, CLng(RecChangeData!药品ID), CLng(IIf(IsNull(RecChangeData!批次), 0, RecChangeData!批次)))
                            
                           With RecCheckStock
                                If .EOF Then
                                    dblStock = 0
                                Else
                                    dblStock = !数量
                                End If
                                
                                '增加相应的库存记录
                                With rsStock
                                    .AddNew
                                    !药品ID = LngPhysicID
                                    !批次 = IIf(IsNull(RecChangeData!批次), 0, RecChangeData!批次)
                                    !变价 = RecCheckStock!变价
                                    !数量 = dblStock
                                    !序号 = LngLocate
                                    .Update
                                End With
                            End With
                        End If
                    End If
                    rsStock.MoveFirst
                    rsStock.Find "序号=" & LngLocate
                    dblStock = rsStock!数量         '取当前库存
                    DblCompare = !实际数量
                    
                    If dblStock < DblCompare Then
                        '设置该记录为缺药状态(该状态不能由用户修改)
                        !执行状态 = IIf(Lng缺药检查 = 1 Or rsStock!批次 <> 0 Or rsStock!变价 = 1, 0, !执行状态)
                        .Update
                    ElseIf !执行状态 = 0 Then
                        '检测是否允许发药
                        lngState = 1
                        If !已收费 = 0 Then lngState = 3
                        If Int允许未审核处方发药 = 0 Then
                            If IsNull(!审核人) Then
                                lngState = 3
                            Else
                                If Trim(!审核人) = "" Then lngState = 3
                            End If
                        End If
                        !执行状态 = lngState                        '缺省为发药
                        .Update
                    End If
                    
                    '如果执行状态为发药，则减库存
                    If !执行状态 = 1 Then
                        With rsStock
                            !数量 = !数量 - DblCompare
                            .Update
                        End With
                    End If
                End If
            End If
            
            '如果没有指定药品（设置所有药品的状态时），并且当前药品的执行状态不是3时，则根据参数自动设置是否“不处理”
            If lng药品ID = 0 And !执行状态 <> 3 Then
                If mstr毒理分类 <> "" And !毒理分类 <> "" Then
                    If InStr("," & mstr毒理分类 & ",", "," & !毒理分类 & ",") > 0 Then
                        !执行状态 = 3
                    End If
                End If
                If mstr价值分类 <> "" And !价值分类 <> "" Then
                    If InStr("," & mstr价值分类 & ",", "," & !价值分类 & ",") > 0 Then
                        !执行状态 = 3
                    End If
                End If
            End If
            
            Str执行状态 = IIf(!执行状态 = 0, "缺药", IIf(!执行状态 = 1, "发药", IIf(!执行状态 = 2, "拒发", "不处理")))
            !状态 = Str执行状态
            .Update
            
            '如果该记录已填充到表格，则连锁更新
            With Bill未发药清单
                If .rows - 1 >= RecChangeData!位置 Then
                    .TextMatrix(RecChangeData!位置, 列名_未发药清单.状态) = Str执行状态
                    If Str执行状态 = "发药" Then
                        .Row = RecChangeData!位置
                        For intCol = 0 To .Cols - 1
                            .Col = intCol
                            .CellBackColor = glngSendBlkColor
                        Next
                    End If
                End If
            End With
            .MoveNext
        Loop
        
        With Bill未发药清单
            .Redraw = True
        End With
        If lng药品ID <> 0 Then .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    Set rsStock = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As adodb.Recordset
    Dim lngRow As Integer
    Dim lng药品ID As Long
    
    On Error GoTo errHandle
    CheckDrugStock = True
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        .Sort = "药品ID Asc"
        
        Do While Not .EOF
            If lng药品ID <> !药品ID Then
                If !执行状态 = 1 Then
                    gstrSQL = "Select 收费细目id From 收费执行科室 Where 执行科室id = [1] And 收费细目id = [2]"
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查药品存储库房", lng药房ID, Val(!药品ID))
                    
                    If rsTmp.EOF Then
                        MsgBox !品名 & "未设置存储库房，不能发药！", vbInformation, gstrSysName
                        CheckDrugStock = False
                        Exit Function
                    End If
                    
                    lng药品ID = !药品ID
                Else
                    lng药品ID = 0
                End If
            End If
            .MoveNext
        Loop
    End With
    
    CheckDrugStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RefreshSendedData() As Boolean
    Dim strCond As String, strSubSql As String
    Dim strName As String
    Dim str过滤记帐人 As String
    Dim strSql病人类型 As String
    
    On Error GoTo errHandle
    '必要条件检查
    If mstrSerchNO = "" And mstr部门 = "" Then
'        MsgBox "请选择领药部门！", vbInformation, gstrSysName
        Call ClearCons
        Exit Function
    End If
    
    mblnFirstSended = False
    
    str过滤记帐人 = IIf(str记帐人 <> "所有记帐人", " AND A.填制人=[1] ", "")
    
    '扣率:bit1=0-长嘱,1-临嘱
    '操作模式:0-所有,1-记帐单,2-记帐表
    If Lng操作模式 = 0 Then
        strCond = " And S.单据 IN(9,10)"
    ElseIf Lng操作模式 = 1 Then
        strCond = " And S.单据=9"
    ElseIf Lng操作模式 = 2 Then
        strCond = " And S.单据=10"
    End If
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If Lng医嘱类型 = 0 Then
    ElseIf Lng医嘱类型 = 1 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf Lng医嘱类型 = 2 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf Lng医嘱类型 = 3 Then
        strCond = strCond & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf Lng医嘱类型 = 4 Then
        strCond = strCond & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '单位设置
    Select Case strUnit
    Case "售价单位"
        strSubSql = "X.计算单位 单位,1 包装,"
    Case "门诊单位"
        strSubSql = "D.门诊单位 单位,D.门诊包装 包装,"
    Case "住院单位"
        strSubSql = "D.住院单位 单位,D.住院包装 包装,"
    Case "药库单位"
        strSubSql = "D.药库单位 单位,D.药库包装 包装,"
    End Select
    
    '得到药品名称串
    Call GetDrugFormat
    Select Case int药品名称
    Case 0  '药品编码与名称
        strName = "'['||X.编码||']'||" & IIf(mblnTradeName, "NVL(A.名称,X.名称)", "X.名称") & " As 品名,"
    Case 1  '药品编码
        strName = "X.编码 As 品名,"
    Case 2  '药品名称
        strName = IIf(mblnTradeName, "NVL(A.名称,X.名称)", "X.名称") & " As 品名,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(A.名称,'')", "Decode(A.名称,Null,'',X.名称)") & " As 其它名, "
    
    '病人类型：病人或婴儿
    If mint病人类型 = 0 Then
        strSql病人类型 = " And Nvl(C.婴儿费,0)=0 "
    ElseIf mint病人类型 = 1 Then
        strSql病人类型 = " And Nvl(C.婴儿费,0)>0 "
    End If
    
    
    If Chk清单.Value = 0 Then
        '##################汇总显示每笔记录还允许退多少##################
        gstrSQL = " SELECT DISTINCT S.ID,S.单据,S.药品ID,S.NO,S.序号,S.扣率,P.名称 科室,C.门诊标志,C.标识号,C.病人ID,C.床号,C.姓名," & _
            strName & _
            " NVL(D.药房分批,0) 分批,X.规格,T.毒理分类," & _
            strSubSql & _
            " S.付数 付,S.实际数量 数量,S.已退数量,S.已发数量 准退数,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,S.效期," & _
            " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,S.审核人,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,1 可操作,C.医嘱序号,I.计算单位,NVL(S.产地,NVL(X.产地,'')) 产地,nvl(M.审查结果,-1) 审查结果,nvl(C.医嘱序号,-1) 医嘱id,S.领药人," & IIf(mbln药品储备 = True, "L.", "'' ") & "库房货位,M.相关ID,c.序号 费用序号, Z.名称 As 英文名,0 As 转出 " & _
            " FROM " & _
            "      (SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,A.扣率," & _
            "          NVL(A.付数,1) 付数,A.实际数量 实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
            "          A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID,A.产地,decode(NVL(A.领用人,''),'','','(领)'||A.领用人) 领药人 " & _
            "      FROM 药品收发记录 A," & _
            "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
            "          FROM 药品收发记录 A" & _
            "          WHERE A.审核人 IS NOT NULL" & _
            "          AND A.库房ID+0=[2] " & _
            "          AND A.审核日期 BETWEEN [8] AND [9] " & str过滤记帐人 & _
            "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
            "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0 And A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)) S,"
        gstrSQL = gstrSQL & "" & _
            "      病人费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,药品特性 T,诊疗项目目录 I,病人医嘱记录 M," & IIf(mbln药品储备 = True, "药品储备限额 L,", "") & "诊疗项目别名 Z" & _
            " WHERE S.药品ID=D.药品ID AND S.对方部门ID+0=P.ID AND D.药名ID=T.药名ID AND d.药品ID=X.ID AND D.药名ID=I.ID and C.医嘱序号=M.ID(+) " & _
            " And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2" & IIf(mbln药品储备 = True, " And S.药品ID=L.药品ID(+) And Nvl(S.库房ID,[2])=L.库房ID(+) ", "") & _
            " AND D.药品ID=A.收费细目ID(+) AND a.性质(+)=3 " & strCond & IIf(mstr床号 = "", "", " And C.床号=[11] ") & _
            " AND S.费用ID=C.ID " & IIf(Val(mlng病人ID) = 0, "", " AND C.病人ID=[3] ") & IIf(Trim(mstr住院号) = "", "", " AND C.标识号=[4] ") & IIf(mstr病人姓名 = "", "", " AND C.姓名 LIKE [5] ") & _
            " AND (S.记录状态=1 OR MOD(S.记录状态,3)=0)" & _
            " AND S.审核人 IS NOT NULL AND S.库房ID+0=[2] " & IIf(mstrDrug = "", "", " And Instr([14],',' || T.药品剂型 || ',') > 0") & IIf(mstr发药类型 = "", "", " And Instr([15],',' || D.发药类型 || ',') > 0") & strCond & strSql病人类型 & _
            IIf(mstr开始NO = "", "", " AND S.NO>=[6] ") & IIf(mstr结束NO = "", "", " AND S.NO<=[7] ") & " AND Abs(S.实际数量*S.付数)>Abs(S.已退数量) " & IIf(mstrUse = "", "", " And Instr([13],',' || S.用法 || ',') > 0") & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ")
        If Trim(mstr部门) <> "" Then
            If mint类型 = 0 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id=C.开单部门id"
            ElseIf mint类型 = 1 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id<>C.开单部门id"
            Else
                If mstr病区发药方式 = "" Then
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 And C.病人科室id=C.开单部门id"
                Else
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 "
                    If mstr病区发药方式 <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                            " Where Instr([16],',' || 工作性质 || ',') > 0) "
                    End If
                End If
            End If
        Else
            If mint类型 = 0 Then
                gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
            ElseIf mint类型 = 1 Then
                gstrSQL = gstrSQL & " And C.病人科室id<>C.开单部门id"
            Else
                If mstr病区发药方式 = "" Then
                    gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
                Else
                    If mstr病区发药方式 <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                            " Where Instr([16],',' || 工作性质 || ',') > 0) "
                    End If
                End If
            End If
        End If
    Else
        '##################清单显示每笔操作过程##################
        gstrSQL = " SELECT DISTINCT S.ID,S.单据,S.药品ID,S.NO,S.序号,S.扣率,P.名称 科室,C.门诊标志,C.标识号,C.病人ID,C.床号,C.姓名," & strName & _
                 " NVL(D.药房分批,0) 分批,X.规格,T.毒理分类," & _
                 strSubSql & _
                 " S.付数 付,S.实际数量 数量,S.已退数量,S.已发数量 准退数,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,S.效期," & _
                 " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,S.审核人,S.审核日期,可操作,C.医嘱序号,I.计算单位,NVL(S.产地,NVL(X.产地,'')) 产地,nvl(M.审查结果,-1) 审查结果,nvl(C.医嘱序号,-1) 医嘱id,S.领药人," & IIf(mbln药品储备 = True, "L.", "'' ") & "库房货位, Z.名称 As 英文名,0 As 转出 " & _
                 " FROM "
        gstrSQL = gstrSQL & _
                 "          (SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,A.扣率," & _
                 "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                 "              A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作,A.产地," & _
                 "              decode(nvl(A.领用人,''),'','',Decode(A.记录状态,1,'(领)'||A.领用人," & _
                 "              decode(Mod(A.记录状态,3),0,'(领)'||A.领用人,1,'(领)'||A.领用人,2,'(退)'||A.领用人))) 领药人 " & _
                 "          FROM 药品收发记录 A," & _
                 "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE A.审核人 IS NOT NULL" & _
                 "          AND A.库房ID+0=[2] " & _
                 "          AND A.审核日期 BETWEEN [8] AND [9] " & str过滤记帐人 & _
                 "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
                 "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 And A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)"
        gstrSQL = gstrSQL & _
                 "          UNION" & _
                 "          SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,A.扣率," & _
                 "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态," & _
                 "          A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                 "          DECODE(A.记录状态,1,1,DECODE(MOD(A.记录状态,3),0,1,MOD(A.记录状态,3)+1)) 可操作,A.产地," & _
                 "          decode(nvl(A.领用人,''),'','',Decode(A.记录状态,1,'(领)'||A.领用人," & _
                 "          decode(Mod(A.记录状态,3),0,'(领)'||A.领用人,1,'(领)'||A.领用人,2,'(退)'||A.领用人))) 领药人 " & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE A.审核人 IS NOT NULL AND NOT (记录状态=1 OR MOD(记录状态,3)=0)" & _
                 "          AND A.库房ID+0=[2] " & _
                 "          AND A.审核日期 BETWEEN [8] AND [9] " & str过滤记帐人 & _
                 "          ) S,"
        gstrSQL = gstrSQL & "" & _
                 "      病人费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,药品特性 T,诊疗项目目录 I,病人医嘱记录 M," & IIf(mbln药品储备 = True, "药品储备限额 L,", "") & "诊疗项目别名 Z " & _
                 " WHERE S.药品ID=D.药品ID AND D.药名ID=T.药名ID AND d.药品ID=x.ID AND S.对方部门ID+0=P.ID AND D.药名ID=I.ID and C.医嘱序号=M.ID(+) " & _
                 " And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 " & IIf(mbln药品储备 = True, " And S.药品ID=L.药品ID(+) And Nvl(S.库房ID,[2])=L.库房ID(+) ", "") & _
                 " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & IIf(mstr床号 = "", "", " And C.床号=[11] ") & _
                 " AND S.费用ID=C.ID " & strCond & strSql病人类型 & _
                 " AND S.审核人 IS NOT NULL" & IIf(mstrDrug = "", "", " And Instr([14],',' || T.药品剂型 || ',') > 0") & IIf(mstr发药类型 = "", "", " And Instr([15],',' || D.发药类型 || ',') > 0") & _
                 IIf(mstr开始NO = "", "", " AND S.NO>=[6] ") & IIf(mstr结束NO = "", "", " AND S.NO<=[7] ") & IIf(mstrUse = "", "", " And Instr([13],',' || S.用法 || ',') > 0") & IIf(mstrSerchNO = "", "", " AND S.NO=[12] ") & _
                 IIf(Val(mlng病人ID) = 0, "", " AND C.病人ID=[3] ") & IIf(Trim(mstr住院号) = "", "", " AND C.标识号=[4] ") & IIf(mstr病人姓名 = "", "", " AND C.姓名 LIKE [5] ")
        If Trim(mstr部门) <> "" Then
            If mint类型 = 0 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id=C.开单部门id"
            ElseIf mint类型 = 1 Then
                gstrSQL = gstrSQL & " And Instr([10], ',' || C.开单部门id || ',') > 0 And C.病人科室id<>C.开单部门id"
            Else
                If mstr病区发药方式 = "" Then
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 And C.病人科室id=C.开单部门id"
                Else
                    gstrSQL = gstrSQL & " And Instr([10], ',' || C.病人病区ID || ',') > 0 "
                    If mstr病区发药方式 <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                            " Where Instr([16],',' || 工作性质 || ',') > 0) "
                    End If
                End If
            End If
        Else
            If mint类型 = 0 Then
                gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
            ElseIf mint类型 = 1 Then
                gstrSQL = gstrSQL & " And C.病人科室id<>C.开单部门id"
            Else
                If mstr病区发药方式 = "" Then
                    gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
                Else
                    If mstr病区发药方式 <> mstrAllType Then
                        gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                            " Where Instr([16],',' || 工作性质 || ',') > 0) "
                    End If
                End If
            End If
        End If
    End If
    
    Dim blnMoved As Boolean
    Dim strSQL As String
    '判断是否存在部分数据已转出
    blnMoved = zldatabase.DateMoved(mstr开始日期_已发)
    If blnMoved Then
        'SQL按记录序号汇总，因任何一笔明细要么在线，要么后备，因此，以UNION方式处理
        strSQL = gstrSQL
        strSQL = Replace(strSQL, "药品收发记录", "H药品收发记录")
        strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
        strSQL = Replace(strSQL, "0 As 转出", "1 As 转出")
        
        gstrSQL = gstrSQL & " UNION ALL " & strSQL
    End If
    
    If Chk清单.Value = 0 Then
        gstrSQL = gstrSQL & " Order By No,单据,费用序号 "
    Else
        gstrSQL = gstrSQL & " Order By No,单据,审核日期"
    End If
    
    '--刷新已发药清单--
'    on error Resume Next
'    err = 0
    
    '初始化记录集
    Call InitRec
    
    RefreshSendedData = False
    
    '已发处方记录
    Set RecBillData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        str记帐人, _
        lng药房ID, _
        mlng病人ID, _
        mstr住院号, _
        mstr病人姓名, _
        mstr开始NO, _
        mstr结束NO, _
        CDate(mstr开始日期_已发), _
        CDate(mstr结束日期_已发), _
        "," & mstr部门 & ",", _
        mstr床号, _
        mstrSerchNO, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr发药类型 & ",", _
        "," & mstr病区发药方式 & ",")
        
    '手工预填充
    Call ClearBill(Bill已发药清单)
    If ProduceInsideSendedRecordset = False Then Exit Function
    If LoadDataInBill已发药清单 = False Then
        MsgBox "填充已发药清单时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    Call SetGroup(Bill已发药清单, Chk清单.Value = 0)
    
    If err <> 0 Then
        MsgBox "读取已发药处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    RefreshSendedData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPhysicDict(ByVal lng药品ID As String) As String
    Dim str单位 As String, str系数 As String
    Dim rsTemp As New adodb.Recordset
    '--获取指定药品ID的品名、规格、单位及包装--
    On Error GoTo errHandle
    GetPhysicDict = " ^ ^ ^ "
    gstrSQL = " SELECT A.药品ID,A.药名ID,NVL(A.药房分批,0) 分批," & _
              " DECODE(B.规格,NULL,B.产地,DECODE(B.产地,NULL,B.规格,B.规格||'|'||B.产地)) 规格," & _
              " A.门诊单位,A.门诊包装,A.住院单位,A.住院包装,A.药库单位,药库包装,B.计算单位 售价单位,1 售价包装 " & _
              " FROM 药品规格 A,收费项目目录 B" & _
              " WHERE A.药品ID=B.ID AND A.药品ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取指定药品ID的品名、规格、单位及包装]", lng药品ID)
    
    If rsTemp.EOF Then Exit Function
    
    Select Case strUnit
    Case "售价单位"
        str单位 = rsTemp!售价单位
        str系数 = rsTemp!售价包装
    Case "门诊单位"
        str单位 = rsTemp!门诊单位
        str系数 = rsTemp!门诊包装
    Case "住院单位"
        str单位 = rsTemp!住院单位
        str系数 = rsTemp!住院包装
    Case "药库单位"
        str单位 = rsTemp!药库单位
        str系数 = rsTemp!药库包装
    End Select
    
    GetPhysicDict = "小宝"
    GetPhysicDict = GetPhysicDict & "^" & IIf(IsNull(rsTemp!规格), " ", rsTemp!规格) & "^" & _
    str单位 & "^" & str系数 & "^" & rsTemp!分批
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetMenuAndToolbarState()
    Dim LngCurLocate As Long                    '当前位置
    '--设置菜单及工具按钮的状态--
    With RecChangeData
        If .RecordCount <> 0 Then .MoveFirst
        .Find "执行状态=0"                      '缺药
        MnuEditDesire.Enabled = (.EOF Xor True)
        Tbar.Buttons("Desire").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
        
        .Find "执行状态=1"                      '发药
        MnuEditVerify.Enabled = (.EOF Xor True)
        Tbar.Buttons("Consignment").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
        
        .Find "执行状态=2"                      '拒发
        mnuEditHandback.Enabled = (.EOF Xor True)
        Tbar.Buttons("Handback").Enabled = (.EOF Xor True)
        If .RecordCount <> 0 Then .MoveFirst
    End With
    With RecChangeSendedData
        If .RecordCount <> 0 Then .MoveFirst
        .Find "执行状态=3"                      '退药
        MnuEditRestore.Enabled = (.EOF Xor True)
        Tbar.Buttons("Restore").Enabled = (.EOF Xor True)
    End With
    
    If mnuEditHandback.Enabled = False Then
        With Bill拒发药清单
            For LngCurLocate = 1 To .rows - 1
                If Trim(.TextMatrix(LngCurLocate, 1)) = "恢复" Then
                    mnuEditHandback.Enabled = True
                    Tbar.Buttons("Handback").Enabled = True
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Function AviShow(Optional ByVal BlnShow As Boolean = True)
    '控制Flash窗体
    DoEvents
    
    If BlnShow Then
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
    Else
        zlCommFun.StopFlash
    End If
    
    DoEvents
End Function

Private Function CheckBill(ByVal IntOper As Integer, ByVal LngID As Long) As Integer
    Dim RecCheck As New adodb.Recordset
    
    '--根据将要执行的操作，判断是否允许--
    '0-拒发;1-发药;2-退药
    '返回:
    '0-允许操作
    '1-已发药
    '2-已删除
    '3-未发药
    On Error GoTo errHandle
    gstrSQL = " Select A.NO,Nvl(B.记录状态,0) AS 审核标志,A.审核人,Decode(Nvl(A.摘要,'小宝'),'拒发',3,B.执行状态) 执行状态 From 药品收发记录 A,病人费用记录 B " & _
             " Where A.费用ID=B.ID And A.ID=[1] "
    If IntOper = 2 Then
        gstrSQL = gstrSQL & " And 审核人 IS Not Null"
    Else
        gstrSQL = gstrSQL & " And 审核人 IS Null"
    End If
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngID)
    
    With RecCheck
        If .EOF Then CheckBill = 2: MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            If IntOper <> 2 Then CheckBill = 1: MsgBox "该处方[" & !NO & "]已被其它操作员发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        Else
            If IntOper = 2 Then CheckBill = 3: MsgBox "该处方[" & !NO & "]还未发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        End If
        If IntOper = 1 Then
            If !执行状态 = 3 Then CheckBill = 2: MsgBox "该处方[" & !NO & "]已拒发，操作被迫中止！", vbInformation, gstrSysName: Exit Function
            If !审核标志 = 0 And Int允许未审核处方发药 = 0 Then
                CheckBill = 4: MsgBox "该处方[" & !NO & "]还未审核，操作被迫中止！", vbInformation, gstrSysName
            End If
        End If
    End With
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function subPrint(ByVal bytMode As Byte)
    '--打印--
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim intCol As Integer
    
    Select Case TabShow.Tab
    Case 0
        Set ObjThis = Bill未发药清单
    Case 1
        Set ObjThis = Bill汇总发药
    Case 2
        Set ObjThis = Bill缺药清单
    Case 3
        Set ObjThis = Bill拒发药清单
    Case 4
        Set ObjThis = Bill已发药清单
    End Select
    
    '恢复字体前景色
    With ObjThis
        .Redraw = False
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = &H80000008
        Next
        .Col = 0
        .Redraw = True
    End With
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "打印人:" & gstrUserName
    ObjAppRow.Add "打印日期:" & Format(zldatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "开始时间:" & Format(IIf(TabShow.Tab = 4, mstr开始日期_已发, mstr开始日期_未发), "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "结束时间:" & Format(IIf(TabShow.Tab = 4, mstr结束日期_已发, mstr结束日期_未发), "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = TabShow.TabCaption(TabShow.Tab)
    Set objPrint.Body = ObjThis
    
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
    
    '恢复选中状态的字体前景色
    Call SetSelectColor(ObjThis)
End Function

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuViewTool, 2
End Sub

Private Sub 权限控制()
    '参数设置
    '发药
    '拒发
    '退药

    If Not IsHavePrivs(mstrPrivs, "发药") Then
        MnuEditVerify.Visible = False
        Tbar.Buttons("Consignment").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "拒发") Then
        mnuEditHandback.Visible = False
        Tbar.Buttons("Handback").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "退药") Then
        If MnuEditVerify.Visible = False And mnuEditHandback.Visible = False Then
            mnuEdit.Visible = False
            Tbar.Buttons("Edit1").Visible = False
        Else
            MnuEditRestore.Visible = False
        End If
        Tbar.Buttons("Restore").Visible = False
    End If
    mnuFilePrintTotal.Visible = IsHavePrivs(mstrPrivs, "汇总打印")
    mnuFileRestore.Visible = IsHavePrivs(mstrPrivs, "打印本次退药明细")
    If Not mnuFileRestore.Visible Then MnuFile2.Visible = mnuFilePrintTotal.Visible
    If Not IsHavePrivs(mstrPrivs, "退其它药房的处方") Then
        mnuEditHandbackBatch.Visible = False
    End If
    If gblnPass And IsHavePrivs(mstrPrivs, "合理用药监测") Then
        mblnStarPass = True
    End If
    If Not IsHavePrivs(mstrPrivs, "退药销帐") Then
        mnuReVerify.Visible = False
        Tbar.Buttons("ReVerify").Visible = False
    End If
End Sub

Private Sub ClearBill(ByVal MsfObj As MSHFlexGrid)
    '清除控件内容
    Dim i As Long, j As Long
    
    MsfObj.Redraw = False
    For i = 1 To MsfObj.rows - 1
        For j = 0 To MsfObj.Cols - 1
            MsfObj.TextMatrix(i, j) = ""
        Next
    Next
    
    MsfObj.rows = 2
    MsfObj.Row = 1: MsfObj.Col = 0
    MsfObj.Redraw = True
End Sub

Private Sub SetSelectColor(ByVal MsfObj As MSHFlexGrid)
    Dim LngSelectRow As Long, intCol As Integer, lngColor As Long
    Dim strCompare As String
    
    On Error Resume Next
    
    With MsfObj
        '用于下拉框定位
        CurCell.Col = .Col
        CurCell.Row = .Row
        CurCell.CellHeight = .CellHeight
        CurCell.CellLeft = .CellLeft
        CurCell.CellTop = .CellTop - 30
        CurCell.CellWidth = .CellWidth
        
        .Redraw = False
        LngSelectRow = .Row         '保存当前选中行
        If Val(.Tag) <> 0 Then
            .Row = Val(.Tag)        '清除上次选中行
            strCompare = IIf(.TextMatrix(.Row, 0) = "", "小宝", .TextMatrix(.Row, 0))
            
            Select Case .Name
            Case "Bill已发药清单"
                With RecChangeSendedData
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Find "位置=" & Val(MsfObj.Tag)
                    End If
                    If .EOF Then
                        lngColor = &H80000008
                    Else
                        lngColor = IIf(!可操作 = 1, glng正常, IIf(!可操作 = 2, glng发药, glng退药))
                    End If
                End With
            Case "Bill未发药清单"
                With RecChangeData
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Find "位置=" & Val(MsfObj.Tag)
                    End If
                    If .EOF Then
                        lngColor = glngOtherBlkColor
                    Else
                        lngColor = IIf(!执行状态 = 1, glngSendBlkColor, glngOtherBlkColor)
                    End If
                End With
'                lngColor = IIf(InStr(1, "合计,小计", strCompare) <> 0, glng发药, glng正常)
            Case "Bill汇总发药"
                lngColor = IIf(InStr(1, "合计,小计", strCompare) <> 0, glng发药, glng正常)
            Case Else
                lngColor = glng正常
            End Select
            
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(.Name = "Bill未发药清单", lngColor, glngOtherBlkColor)
            Next
            .Col = 0
        End If
        
        .Tag = LngSelectRow
        .Row = .Tag                 '设置当前选中行
        strCompare = IIf(.TextMatrix(.Row, 0) = "", "小宝", .TextMatrix(.Row, 0))
        
        Select Case .Name
        Case "Bill已发药清单"
            With RecChangeSendedData
                If .RecordCount <> 0 Then
                    .MoveFirst
                    .Find "位置=" & LngSelectRow
                End If
                If .EOF Then
                    lngColor = &H8000000D
                Else
                    lngColor = IIf(!可操作 = 1, glng正常, IIf(!可操作 = 2, glng发药, glng退药))
                End If
            End With
        Case "Bill未发药清单"
            lngColor = IIf(InStr(1, "合计,小计", strCompare) <> 0, glng发药, glng正常)
        Case "Bill汇总发药"
            lngColor = IIf(InStr(1, "合计,小计", strCompare) <> 0, glng发药, glng正常)
        Case Else
            lngColor = glng正常
        End Select
        
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &HC0C0C0
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Function SetMenuCheck(ByVal MenuObj As Menu) As Menu
    Dim MenuCheck As Menu, strState As String
    '设置对应菜单的选择状态,并返回
    
    Select Case MenuObj.Name
    Case "PopMenu_1"
        Consignment.Checked = False
        Lack.Checked = False
        HandBack.Checked = False
        Nop_1.Checked = False
        
        strState = Bill未发药清单.TextMatrix(Bill未发药清单.Row, 列名_未发药清单.状态)
        Select Case strState
        Case "发药"
            Set MenuCheck = Consignment
        Case "缺药"
            Set MenuCheck = Lack
        Case "拒发"
            Set MenuCheck = HandBack
        Case "不处理"
            Set MenuCheck = Nop_1
        End Select
        MenuCheck.Checked = True
    Case "PopMenu_2"
        Restore.Checked = False
        Nop_2.Checked = False
        
        strState = Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.状态)
        Select Case strState
        Case "退药"
            Set MenuCheck = Restore
        Case "不处理"
            Set MenuCheck = Nop_2
        End Select
        MenuCheck.Checked = True
    Case "PopMenu_3"
        ResumeDo.Checked = False
        Nop_3.Checked = False
        
        strState = Bill拒发药清单.TextMatrix(Bill拒发药清单.Row, 1)
        Select Case strState
        Case "恢复"
            Set MenuCheck = ResumeDo
        Case "不处理"
            Set MenuCheck = Nop_3
        End Select
        MenuCheck.Checked = True
    End Select
    Set SetMenuCheck = MenuCheck
End Function

Private Sub UpdateRsByMenu(ByVal MenuObj As Menu, Optional ByVal IntStyle As Integer = 1)
    Dim lngFind As Long
    '1:未发药
    '3:拒发药
    '2:已发药
    
    '--更新内部记录集--
    Select Case IntStyle
    Case 1
        With Bill未发药清单
            lngFind = .Row
        End With
        With RecChangeData
            If .RecordCount <> 0 Then .MoveFirst
            .Find "位置=" & lngFind
            If .EOF Then
                MsgBox "未找到该记录！", vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case MenuObj.Name
            Case "Consignment"
                lngFind = 1
            Case "HandBack"
                lngFind = 2
            Case Else
                lngFind = 3
            End Select
            !执行状态 = lngFind
            .Update
        End With
    
        '更新相关记录的执行状态
        Call CheckStock(RecChangeData!药品ID)
    Case 2
        With Bill已发药清单
            lngFind = .Row
        End With
        With RecChangeSendedData
            If .RecordCount <> 0 Then .MoveFirst
            .Find "位置=" & lngFind
            If .EOF Then
                MsgBox "未找到该记录！", vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case MenuObj.Name
            Case "Restore"
                lngFind = 3
                Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.状态) = "退药"
                Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.退药数) = Val(Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.准退数))
            Case Else
                lngFind = 1
                Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.状态) = "不处理"
                Bill已发药清单.TextMatrix(Bill已发药清单.Row, 列名_已发药清单.退药数) = ""
            End Select
            !执行状态 = lngFind
            .Update
        End With
        Call Bill已发药清单_EnterCell
    Case 3
        Select Case MenuObj.Name
        Case "ResumeDo"
            Bill拒发药清单.TextMatrix(Bill拒发药清单.Row, 1) = "恢复"
        Case Else
            Bill拒发药清单.TextMatrix(Bill拒发药清单.Row, 1) = "不处理"
        End Select
    End Select
    
    '设置菜单及工具按钮的状态
    Call SetMenuAndToolbarState
End Sub

Private Function Load拒发()
    Dim ArrayPhysic As Variant
    Dim rsRefuse As New adodb.Recordset
    Dim strCond As String, strSubSql As String
    
    '先装入已实实在在拒发的处方清单

    '扣率:bit1=0-长嘱,1-临嘱；bit2:3-离院带药
    '操作模式:0-所有,1-记帐单,2-记帐表
    On Error GoTo errHandle
    If Lng操作模式 = 0 Then
        strCond = " And S.单据 IN(9,10)"
    ElseIf Lng操作模式 = 1 Then
        strCond = " And S.单据=9"
    ElseIf Lng操作模式 = 2 Then
        strCond = " And S.单据=10"
    End If
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    '用医嘱序号来判定是否为医嘱 by lyq 2005-05-18
    Dim str医嘱序号 As String
    If Lng医嘱类型 = 0 Then
    ElseIf Lng医嘱类型 = 1 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' "
        str医嘱序号 = " And Nvl(C.医嘱序号,0) + 0 > 0 "
    ElseIf Lng医嘱类型 = 2 Then
        strCond = strCond & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' "
        str医嘱序号 = " And Nvl(C.医嘱序号,0) + 0 > 0 "
    ElseIf Lng医嘱类型 = 3 Then
        strCond = strCond
        str医嘱序号 = " And (Nvl(C.医嘱序号,0) + 0 = 0 Or S.扣率 Is Null) "
    ElseIf Lng医嘱类型 = 4 Then
        strCond = strCond & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') "
        str医嘱序号 = " And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If int离院带药 = 0 Then
    ElseIf int离院带药 = 1 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf int离院带药 = 2 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf int离院带药 = 3 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 4 Then
        strCond = strCond & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 5 Then
        strCond = strCond & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf int离院带药 = 6 Then
        strCond = strCond & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    '单位设置
    Select Case strUnit
    Case "售价单位"
        strSubSql = "X.计算单位 单位,1 包装,"
    Case "门诊单位"
        strSubSql = "D.门诊单位 单位,D.门诊包装 包装,"
    Case "住院单位"
        strSubSql = "D.住院单位 单位,D.住院包装 包装,"
    Case "药库单位"
        strSubSql = "D.药库单位 单位,D.药库包装 包装,"
    End Select

    gstrSQL = " SELECT DISTINCT S.ID,S.药品ID,P.名称 科室,S.配药人,C.操作员姓名 审核人,S.单据,NVL(S.扣率,0) 扣率," & _
             " S.NO,C.床号,C.姓名,C.门诊标志,'['||X.编码||']'||" & IIf(mblnTradeName, "NVL(A.名称,X.名称)", "X.名称") & " 品名,S.付数 付,S.实际数量 数量," & _
             " NVL(D.药房分批,0) 分批,X.规格," & _
             strSubSql & _
             " DECODE(S.批号,NULL,'',S.批号) 批号,NVL(S.批次,0) 批次,T.毒理分类," & _
             " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,C.医嘱序号" & _
             " FROM " & _
             "      (SELECT * FROM 药品收发记录 S " & _
             "      WHERE MOD(记录状态,3)=1 AND NVL(LTRIM(RTRIM(摘要)),'小宝')='拒发' " & _
             "      AND 审核人 IS NULL" & _
             "      AND (库房ID+0=[1] OR 库房ID IS NULL) AND 填制日期 BETWEEN [2] AND [3] " & strCond & IIf(mstrUse = "", "", " And Instr([6],',' || S.用法 || ',') > 0")
    gstrSQL = gstrSQL & ") S,病人费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,药品特性 T " & _
             " WHERE S.药品ID=D.药品ID AND D.药品ID=X.ID and D.药名ID=T.药名ID" & _
             " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & IIf(mstrDrug = "", "", " And Instr([7],',' || T.药品剂型 || ',') > 0") & IIf(mstr发药类型 = "", "", " And Instr([8],',' || D.发药类型 || ',') > 0") & _
             " AND S.对方部门ID=P.ID AND S.NO=C.NO AND S.费用ID=C.ID " & str医嘱序号 & IIf(mstr床号 = "", "", " And C.床号=[5] ")
             
    Select Case mint范围
    Case 1
        gstrSQL = gstrSQL & " And S.实际数量>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.实际数量<0"
    End Select
    
    If Trim(mstr部门) <> "" Then
        If mint类型 = 0 Then
            gstrSQL = gstrSQL & " And Instr([4], ',' || C.开单部门id || ',') > 0 And C.病人科室id=C.开单部门id"
        ElseIf mint类型 = 1 Then
            gstrSQL = gstrSQL & " And Instr([4], ',' || C.开单部门id || ',') > 0 And C.病人科室id<>C.开单部门id"
        Else
            If mstr病区发药方式 = "" Then
                gstrSQL = gstrSQL & " And Instr([4], ',' || C.病人病区ID || ',') > 0 And C.病人科室id=C.开单部门id"
            Else
                gstrSQL = gstrSQL & " And Instr([4], ',' || C.病人病区ID || ',') > 0 "
                If mstr病区发药方式 <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([9],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End If
    Else
        If mint类型 = 0 Then
            gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
        ElseIf mint类型 = 1 Then
            gstrSQL = gstrSQL & " And C.病人科室id<>C.开单部门id"
        Else
            If mstr病区发药方式 = "" Then
                gstrSQL = gstrSQL & " And C.病人科室id=C.开单部门id"
            Else
                If mstr病区发药方式 <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([9],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End If
    End If
    gstrSQL = gstrSQL & " Order By S.No,S.单据"
    
    '--填充拒发清单--
'    On Error Resume Next
'    err = 0
    
    Set rsRefuse = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        lng药房ID, _
        CDate(mstr开始日期_未发), _
        CDate(mstr结束日期_未发), _
        "," & mstr部门 & ",", _
        mstr床号, _
        "," & mstrUse & ",", _
        "," & mstrDrug & ",", _
        "," & mstr发药类型 & ",", _
        "," & mstr病区发药方式 & ",")
    
    With rsRefuse
        Do While Not .EOF
            With Bill拒发药清单
                If Trim(.TextMatrix(.rows - 1, 0)) <> "" Then .rows = .rows + 1
                .TextMatrix(.rows - 1, 0) = rsRefuse!科室
                .TextMatrix(.rows - 1, 1) = "不处理"
                .TextMatrix(.rows - 1, 2) = rsRefuse!NO
                .TextMatrix(.rows - 1, 3) = IIf(NVL(rsRefuse!单量, 0) = 0, IIf(rsRefuse!门诊标志 = 1 Or rsRefuse!门诊标志 = 4, "门诊记帐单", IIf(rsRefuse!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(rsRefuse!扣率) = True, "住院记帐单", IIf(rsRefuse!扣率 Like "0*", "长嘱", IIf(rsRefuse!扣率 Like "1*", "临嘱", "记帐表"))))
                .TextMatrix(.rows - 1, 4) = IIf(IsNull(rsRefuse!床号), "", rsRefuse!床号)
                .TextMatrix(.rows - 1, 5) = IIf(IsNull(rsRefuse!姓名), "", rsRefuse!姓名)
                '获取该药品的相关信息
                .TextMatrix(.rows - 1, 6) = rsRefuse!品名
                .TextMatrix(.rows - 1, 7) = IIf(IsNull(rsRefuse!规格), "", rsRefuse!规格)
                .TextMatrix(.rows - 1, 8) = IIf(IsNull(rsRefuse!批号), "", rsRefuse!批号)
                .TextMatrix(.rows - 1, 9) = FormatEx(rsRefuse!数量 * rsRefuse!付 / rsRefuse!包装, 5) & rsRefuse!单位
                .TextMatrix(.rows - 1, 10) = FormatEx(rsRefuse!单价 * rsRefuse!包装, 5)
                .TextMatrix(.rows - 1, 11) = Format(rsRefuse!金额, "#####0.00;-#####0.00; ;")
                .RowData(.rows - 1) = rsRefuse!Id
            End With
            If Not .EOF Then
                .MoveNext
            Else
                Exit Do
            End If
        Loop
        .Close
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetMenu(Optional ByVal blnEnable As Boolean = False)
    MnuFilePrint.Enabled = blnEnable
    MnuFilePreview.Enabled = blnEnable
    MnuFileExcel.Enabled = blnEnable
    Tbar.Buttons("Preview").Enabled = blnEnable
    Tbar.Buttons("Print").Enabled = blnEnable
End Sub

Private Sub LocateCboItemData(ByVal cboObj As ComboBox, ByVal lngItem As Long)
    Dim LngLocate As Long
    With cboObj
        If .ListCount = 0 Then Exit Sub
        For LngLocate = 0 To .ListCount - 1
            If .ItemData(LngLocate) = lngItem Then
                .ListIndex = LngLocate
                Exit Sub
            End If
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub ShowCbo()
    On Error Resume Next
    
    With Cbo批号
        If .ListCount = 0 Then Exit Sub
        .Left = Bill未发药清单.Left + TabShow.Left + CurCell.CellLeft
        .Top = Bill未发药清单.Top + TabShow.Top + CurCell.CellTop
        .Width = CurCell.CellWidth
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub ShowTxt(Optional ByVal 对齐方式 As Integer = 1)
    '0-左对齐;1-右对齐;2-居中对齐
    On Error Resume Next
    With TxtInput
        .Alignment = 对齐方式
        .Left = Bill已发药清单.Left + TabShow.Left + CurCell.CellLeft
        .Top = Bill已发药清单.Top + TabShow.Top + CurCell.CellTop + 20
        .Width = CurCell.CellWidth - 20
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    Call SelAll(TxtInput)
End Sub

Private Sub tbsType_Click()
    If mintLastDeptType <> tbsType.SelectedItem.Index - 1 Then
        txt科室.Tag = ""
        txt科室.Text = ""
        mintLastDeptType = tbsType.SelectedItem.Index - 1
    End If
    
End Sub

Private Sub TimerAuto_Timer()
    '自动刷新只针对未发药品清单
    Dim dateCurr As Date
        
    '如果窗口最小化时退出
    If Me.WindowState = 1 Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If mlngMyWindow = 0 Then
        mlngMyWindow = GetActiveWindow()
    Else
        If mlngMyWindow <> GetActiveWindow() Then Exit Sub
    End If
    
    '如果不是未发药界面或者自动刷新参数为0时退出
    If TabShow.Tab <> 0 Or mint自动刷新未发药清单 = 0 Then Exit Sub
    
    '根据当前时间与上次刷新时间间隔来控制是否刷新
    dateCurr = zldatabase.Currentdate
    If DateDiff("s", mdate上次刷新时间, dateCurr) < mint自动刷新未发药清单 * 60 Then Exit Sub
    
    TimerAuto.Enabled = False
    DoEvents
    Call mnuViewRefresh_Click

'    MsgBox "Ok！" & "[" & Format(dateCurr, "yyyy-mm-dd hh:mm:ss") & "]" & "[" & Format(mdate上次刷新时间, "yyyy-mm-dd hh:mm:ss") & "]"
'    mdate上次刷新时间 = zldatabase.Currentdate
    
    DoEvents
    TimerAuto.Enabled = True
End Sub
Private Sub TxtInput_LostFocus()
    Dim blnUnValid As Boolean, dblCount As Double
    Dim lng医嘱序号 As Long
    Dim rsTemp As New adodb.Recordset
'    On Error Resume Next
    On Error GoTo errHandle
    If Not TxtInput.Visible Then Exit Sub
    blnUnValid = False
    TxtInput = Trim(TxtInput)
    If TxtInput = "" Then TxtInput = 0
    
    blnUnValid = Not IsNumeric(TxtInput)
    If Not blnUnValid Then blnUnValid = Not ((Abs(TxtInput) <= Abs(TxtInput.Tag)) And ((Val(TxtInput) >= 0 And Val(TxtInput.Tag) >= 0) Or (Val(TxtInput) <= 0 And Val(TxtInput.Tag) <= 0)))
    If blnUnValid Then TxtInput = Val(TxtInput.Tag)
    
    With RecChangeSendedData
        .MoveFirst
        .Find "位置=" & CurCell.Row
        If .EOF Then Exit Sub
        
        '先检查是否是医嘱产生的药品记录
        '如果不是则不管
        '如果是，检查系统参数是否允许未作废医嘱退药，如果不允许，退药数为零
        '如果允许则不管
        dblCount = FormatEx(TxtInput.Text, 5)
        If dblCount <> 0 And bln医嘱作废 = False Then
            gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", CLng(!Id))

            If (rsTemp!扣率 Like "1*") Then       '临嘱
                gstrSQL = "Select nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 病人费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", CLng(!Id))
            
                If Not rsTemp.EOF Then
                    If (rsTemp!门诊标志 = 1 Or rsTemp!门诊标志 = 4) And rsTemp!医嘱序号 <> 0 Then
                        gstrSQL = "Select decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where ID=[1]"
                        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rsTemp!医嘱序号))
                        
                        If rsTemp!作废 = 0 Then
                            dblCount = 0
                            'MsgBox "该笔医嘱还未作废，不能退药！", vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
        End If
        
        Bill已发药清单.TextMatrix(CurCell.Row, 列名_已发药清单.退药数) = FormatEx(dblCount, 5)
        !退药数 = Val(TxtInput.Text)
        .Update
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MsgErr(ByVal strMsg As String)
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub AddCollect(ByVal dbl合计金额 As Double, Optional ByVal str汇总 As String = "合计")
    Dim intCol As Integer, str合计 As String
    
    dbl合计金额 = Val(Format(dbl合计金额, "#####0.00;-#####0.00; ;"))
    str合计 = zlCommFun.UppeMoney(dbl合计金额)
    
    Select Case TabShow.Tab
    Case 0
        With Bill未发药清单
            .TextMatrix(.rows - 1, 列名_未发药清单.科室) = str汇总
            .TextMatrix(.rows - 1, 列名_未发药清单.开单医生) = str汇总
            .TextMatrix(.rows - 1, 列名_未发药清单.状态) = Format(dbl合计金额, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, 列名_未发药清单.类型) = Format(dbl合计金额, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, 列名_未发药清单.NO) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.记帐员) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.床号) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.姓名) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.住院号) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.药品名称) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.规格) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.产地) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.批号) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.付) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.数量) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.单价) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.金额) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.单量) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.频次) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.用法) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.记帐时间) = str合计
            .TextMatrix(.rows - 1, 列名_未发药清单.说明) = str合计
            
            .Row = .rows - 1
            .Col = 0: .CellAlignment = 4
'            .Col = 1: .CellAlignment = 4
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellForeColor = glng发药
            Next
            .RowData(.Row) = glng发药
            .rows = .rows + 1
        End With
    Case 1
        If dbl合计金额 = 0 Then Exit Sub
        With Bill汇总发药
            .TextMatrix(.rows - 1, 0) = str汇总
            .TextMatrix(.rows - 1, 1) = Format(dbl合计金额, "#####0.00;-#####0.00; ;")
            .TextMatrix(.rows - 1, 2) = str合计
            .TextMatrix(.rows - 1, 3) = str合计
            .TextMatrix(.rows - 1, 4) = str合计
            If Lng汇总显示 = 1 Then
                .TextMatrix(.rows - 1, 5) = str合计
                .TextMatrix(.rows - 1, 6) = str合计
                .TextMatrix(.rows - 1, 7) = str合计
            End If
            
            .Row = .rows - 1
            .Col = 0: .CellAlignment = 4
            .Col = 1: .CellAlignment = 7
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellForeColor = glng发药
            Next
            .RowData(.Row) = glng发药
        End With
    Case Else
        Exit Sub
    End Select
End Sub

Private Sub UpdateState(ByVal bln未发 As Boolean, Optional ByVal bln全选 As Boolean = True)
    Dim intState As Integer, strState As String, lng位置 As Long
    '更新未发药清单或已发药清单的状态
    
    intState = IIf(bln未发, IIf(bln全选, gInt未发药清单发药, gInt未发药清单不处理), _
                            IIf(bln全选, gInt已发药清单退药, gInt已发药清单不处理))
    strState = IIf(bln未发, IIf(bln全选, "发药", "不处理"), _
                            IIf(bln全选, "退药", "不处理"))
    
    If bln未发 Then
'        With RecChangeData
'            If .RecordCount = 0 Then Exit Sub
'            .MoveFirst
'
'            Do While Not .EOF
'
'                .MoveNext
'            Loop
'        End With
    Else
        With RecChangeSendedData
            If TxtInput.Visible Then TxtInput.Visible = False
            
            If .RecordCount = 0 Then Exit Sub
            .MoveFirst
            
            Do While Not .EOF
                If !可操作 = 1 Then
                    lng位置 = !位置
                    Bill已发药清单.TextMatrix(lng位置, 列名_已发药清单.状态) = strState
                    If intState = 3 Then
                        Bill已发药清单.TextMatrix(lng位置, 列名_已发药清单.退药数) = Bill已发药清单.TextMatrix(lng位置, 列名_已发药清单.准退数)
                    Else
                        Bill已发药清单.TextMatrix(lng位置, 列名_已发药清单.退药数) = ""
                    End If
                    
                    !执行状态 = intState
                    .Update
                End If
                .MoveNext
            Loop
        End With
    End If
    Call SetMenuAndToolbarState
End Sub

Private Sub FindRecord(Optional ByVal BlnFirst As Boolean = True)
    Dim RecObject As adodb.Recordset
    Static lng未发 As Long, lng已发 As Long
    Dim lngRecord As Long
    Dim strMsg As String
    Dim blnExist As Boolean
    'lngLocate:初始化静态变量（当页面发生改变或刷新时）
    
    MnuViewLocateNext.Enabled = False
    MnuViewLocateNext.Tag = 0
    If strFind = "" Then Exit Sub
    
    '查找指定内容的记录
    Select Case TabShow.Tab
    Case 0      '未发药清单
        lngRecord = lng未发
        Set RecObject = Rec未发.Clone
    Case 4      '已发药清单
        lngRecord = lng已发
        Set RecObject = Rec已发.Clone
    End Select
    
    '合法性验证
    If RecObject Is Nothing Then Exit Sub
    If RecObject.State = 0 Then Exit Sub
    If RecObject.RecordCount = 0 Then Exit Sub
    
    RecObject.MoveFirst
    If Not BlnFirst Then
        RecObject.Find "位置=" & lngRecord
        If RecObject.EOF Then RecObject.MoveFirst
        RecObject.MoveNext
    End If
    
    Do While Not RecObject.EOF
        '查找该记录集是否在内部映射记录集中
        Select Case TabShow.Tab
        Case 0
            With RecChangeData
                If .RecordCount = 0 Then Exit Sub
                .Filter = strFind
                If .RecordCount <> 0 Then
                    .Find "位置=" & RecObject!位置
                    blnExist = Not (.EOF)
                End If
                .Filter = 0
            End With
        Case 4
            With RecChangeSendedData
                If .RecordCount = 0 Then Exit Sub
                .Filter = strFind
                If .RecordCount <> 0 Then
                    .Find "位置=" & RecObject!位置
                    blnExist = Not (.EOF)
                End If
                .Filter = 0
            End With
        End Select
        If blnExist Then Exit Do
        RecObject.MoveNext
    Loop
    If Not blnExist Then
        If MsgBox("查找结束，是否从头再找一遍？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call FindRecord(True)
        End If
        Exit Sub
    End If
    
    MnuViewLocateNext.Enabled = True
    MnuViewLocateNext.Tag = 1
    '选择表格的具体行
    Select Case TabShow.Tab
    Case 0
        lng未发 = RecObject!位置
        Bill未发药清单.Row = RecObject!位置
        Bill未发药清单_EnterCell
    Case 4
        lng已发 = RecObject!位置
        Bill已发药清单.Row = RecObject!位置
        Bill已发药清单_EnterCell
    End Select
End Sub

Private Function CheckSpec(ByVal strRecipeKey As String) As Boolean
    Dim strNote As String
    Dim rsTemp As New adodb.Recordset
    '对毒麻类药品进行检查
    On Error GoTo errHandle
    gstrSQL = "SELECT Distinct '['||C.编码||']'||NVL(L.名称,C.名称) 品名,X.毒理分类 " & _
             "   FROM (Select 药名ID,药品ID From 药品规格 Where 药品ID IN (" & strRecipeKey & ")) B, " & _
             "        收费项目目录 C, " & _
             "        收费项目别名 L, " & _
             "        药品特性     X " & _
             "  WHERE X.药名ID = B.药名ID And B.药品ID = C.ID  " & _
             "        AND C.ID = L.收费细目ID(+) AND L.性质(+) = 3 AND L.码类(+) = 1  " & _
             "        AND X.毒理分类 <> '普通药' " & _
             "  Order by X.毒理分类"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "对毒麻类药品进行检查")
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
    If MsgBox("是否对以下毒、麻、精神类药品进行发药？" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    CheckSpec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BuildRecord(Optional ByVal bln发药 As Boolean = True)
    Dim intRow As Integer, intRows As Integer
    Dim strNo As String, lng单据 As Long, str序号 As String, lng病人ID As Long
    Dim blnAdd As Boolean
    
    Call InitCheckRec
    '根据待发药、待退药清单，按单据获取明细序号
    If bln发药 Then
        intRows = RecChangeData.RecordCount
        If RecChangeData.RecordCount <> 0 Then RecChangeData.MoveFirst
        For intRow = 1 To intRows
            If Val(RecChangeData!执行状态) = 1 Then
                strNo = RecChangeData!NO
                lng单据 = Val(RecChangeData!单据)
                lng病人ID = RecChangeData!病人ID
                
                rs序号.Filter = "单据标识='" & strNo & "|" & lng单据 & "'"
                blnAdd = (rs序号.RecordCount = 0)
                If Not blnAdd Then
                    rs序号.Find "病人ID=" & lng病人ID
                    blnAdd = rs序号.EOF
                End If
                
                If blnAdd Then rs序号.AddNew
                rs序号!单据标识 = strNo & "|" & lng单据
                rs序号!病人ID = lng病人ID
                
                str序号 = NVL(rs序号!序号)
                If InStr(1, "," & str序号 & ",", "," & Val(RecChangeData!序号) & ",") = 0 Then
                    If str序号 = "" Then
                        str序号 = Val(RecChangeData!序号)
                    Else
                        str序号 = str序号 & "," & Val(RecChangeData!序号)
                    End If
                    rs序号!序号 = str序号
                End If
                rs序号.Update
                rs序号.Filter = 0
            End If
            RecChangeData.MoveNext
        Next
        If RecChangeData.RecordCount <> 0 Then RecChangeData.MoveFirst
    Else
        intRows = RecChangeSendedData.RecordCount
        If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.MoveFirst
        For intRow = 1 To intRows
            If Val(RecChangeSendedData!执行状态) = 3 Then
                If Val(NVL(RecChangeSendedData!退药数, 0)) <> 0 Then
                    strNo = RecChangeSendedData!NO
                    lng单据 = Val(RecChangeSendedData!单据)
                    lng病人ID = RecChangeSendedData!病人ID
                    
                    rs序号.Filter = "单据标识='" & strNo & "|" & lng单据 & "'"
                    blnAdd = (rs序号.RecordCount = 0)
                    If Not blnAdd Then
                        rs序号.Find "病人ID=" & lng病人ID
                        blnAdd = rs序号.EOF
                    End If
                    
                    If blnAdd Then rs序号.AddNew
                    rs序号!单据标识 = strNo & "|" & lng单据
                    rs序号!病人ID = lng病人ID
                    
                    str序号 = NVL(rs序号!序号)
                    If InStr(1, "," & str序号 & ",", "," & Val(RecChangeSendedData!序号) & ",") = 0 Then
                        If str序号 = "" Then
                            str序号 = Val(RecChangeSendedData!序号)
                        Else
                            str序号 = str序号 & "," & Val(RecChangeSendedData!序号)
                        End If
                        rs序号!序号 = str序号
                    End If
                    rs序号.Update
                    rs序号.Filter = 0
                End If
            End If
            RecChangeSendedData.MoveNext
        Next
        If RecChangeSendedData.RecordCount <> 0 Then RecChangeSendedData.MoveFirst
    End If

    '打印
    intRows = rs序号.RecordCount
    If rs序号.RecordCount <> 0 Then rs序号.MoveFirst
    For intRow = 1 To intRows
        Debug.Print rs序号!单据标识 & "," & rs序号!病人ID & "," & rs序号!序号
        rs序号.MoveNext
    Next
    If rs序号.RecordCount <> 0 Then rs序号.MoveFirst
End Sub

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng单据 As Long, str序号 As String, lng病人ID As Long
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng病人ID = !病人ID
            str序号 = NVL(!序号)
            If Not IsReceiptBalance_Charge(mstrPrivs, lng单据, strNo, str序号) Then Exit Function
            If Not IsOutPatient(mstrPrivs, lng单据, strNo, lng病人ID) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitCheckRec()
    Set rs序号 = New adodb.Recordset
    With rs序号
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub











Private Sub txtPati_GotFocus()
    Call SelAll(txtPati)
    
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    If Val(lblPatiInputType.Tag) = PatiInfo.就诊卡 Then
        If gtype_UserSysParms.P12_就诊卡是否密文显示 Then
            txtPati.PasswordChar = "*"
        End If
        txtPati.MaxLength = gtype_UserSysParms.P20_就诊卡号长度
    End If
End Sub
Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         Call txtPati_Validate(True)
    End If
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    mblnCard = False
    
    If Val(lblPatiInputType.Tag) = PatiInfo.住院号 Or Val(lblPatiInputType.Tag) = PatiInfo.病人ID Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf Val(lblPatiInputType.Tag) = PatiInfo.姓名 Then
        mblnCard = zlCommFun.InputIsCard(txtPati, KeyAscii, glngSys)
    ElseIf Val(lblPatiInputType.Tag) = PatiInfo.就诊卡 Then
        mblnCard = (KeyAscii <> 8 And Len(txtPati.Text) = gtype_UserSysParms.P20_就诊卡号长度 - 1 And txtPati.SelLength <> Len(txtPati.Text))
    End If
End Sub

    
Private Sub txtPati_Validate(Cancel As Boolean)
    Dim strDeptInfo As String
    Dim strInput As String
    
    '取病人名称，病人当前病区，并提取处方记录
    '当取到病人信息后，返回输入框格式：输入信息-病人姓名
    If InStr(Trim(txtPati.Text), "-") > 0 Then
        '取“-”前面的输入信息
        strInput = Mid(Trim(txtPati.Text), 1, InStr(Trim(txtPati.Text), "-") - 1)
    Else
        strInput = Trim(txtPati.Text)
    End If
    
    If strInput = "" Then Exit Sub
    
    If Val(lblPatiInputType.Tag) = PatiInfo.单据号 Then
        If IsNumeric(strInput) Then
            strInput = GetFullNO(strInput, 14)
        End If
    End If
    
    strDeptInfo = GetPatiInfo(Val(lblPatiInputType.Tag), strInput)
    
    If strDeptInfo <> "" Then
        mintLastDeptType = 2
        tbsType.Tabs(3).Selected = True
        
        txt科室.Text = Mid(Split(strDeptInfo, "|")(0), InStr(Split(strDeptInfo, "|")(0), ",") + 1)
        txt科室.Tag = Mid(Split(strDeptInfo, "|")(0), 1, InStr(Split(strDeptInfo, "|")(0), ",") - 1)
        
        Select Case Val(lblPatiInputType.Tag)
        Case PatiInfo.姓名
            If mblnCard = True Then
                txtPati.Text = UCase(strInput)
                txtPati.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            Else
                txtPati.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            End If
        Case PatiInfo.就诊卡
            txtPati.PasswordChar = ""
            txtPati.MaxLength = 0
            txtPati.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtPati.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case Else
            txtPati.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
        End Select
        
        DoEvents
        
        Call cmdRefresh_Click
    End If
End Sub

Private Sub txt给药途径_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub txt科室_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As adodb.Recordset
    If KeyCode = vbKeyReturn Then
        If Trim(txt科室.Text) = "" Then
            txt科室.Tag = ""
            Exit Sub
        End If
        
        Set rsTemp = SelectDept(tbsType.SelectedItem.Index - 1, Trim(txt科室.Text))
        
        If Not rsTemp Is Nothing Then
            txt科室.Tag = rsTemp("ID")
            txt科室.Text = rsTemp("部门")
        End If
    End If
End Sub
Private Sub txt科室_Validate(Cancel As Boolean)
    If Trim(txt科室.Text) = "" Then
        txt科室.Tag = ""
        Exit Sub
    End If
End Sub


Private Sub txt留存数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call txt留存数_LostFocus
        txt留存数.Visible = False
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub








Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub txt留存数_LostFocus()
    Dim dbl留存数 As Double
    Dim dbl应发数 As Double
    Dim dbl实发数 As Double
    
    dbl应发数 = Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.应发数量)) - Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.销帐数量))
    
    If lngLastCol = 列名_科室汇总清单.实发数量 Then
        dbl实发数 = Val(txt留存数.Text)
        If dbl实发数 > dbl应发数 Or dbl实发数 < 0 Then
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量) = FormatEx(dbl应发数, 5)
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.留存数量) = 0
        Else
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量) = FormatEx(dbl实发数, 5)
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.留存数量) = FormatEx(dbl应发数 - Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量)), 5)
        End If
    ElseIf lngLastCol = 列名_科室汇总清单.留存数量 Then
        dbl留存数 = Val(txt留存数.Text)
        If dbl留存数 > dbl应发数 Or dbl留存数 < 0 Then
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量) = FormatEx(dbl应发数, 5)
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.留存数量) = 0
        Else
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量) = FormatEx(dbl应发数 - Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.留存数量)), 5)
            Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.留存数量) = FormatEx(dbl留存数, 5)
        End If
    End If
            
    DoEvents
    
    Bill汇总发药.Row = LngLastRow
    Bill汇总发药.Col = 列名_科室汇总清单.实发数量
    If Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量)) < 0 Then
        Bill汇总发药.CellForeColor = vbRed
    ElseIf Val(Bill汇总发药.TextMatrix(LngLastRow, 列名_科室汇总清单.实发数量)) > 0 Then
        Bill汇总发药.CellForeColor = vbBlue
    End If
End Sub


