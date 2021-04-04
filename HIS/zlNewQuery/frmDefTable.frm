VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDefTable 
   Caption         =   "用户表格定义"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmDefTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Tag             =   "可变化的"
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfPrint 
      Height          =   1290
      Left            =   570
      TabIndex        =   18
      Top             =   4545
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2275
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      MergeCells      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   825
      Left            =   3075
      ScaleHeight     =   825
      ScaleWidth      =   5010
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   750
      Width           =   5010
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   0
         Left            =   1575
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196610
         BuddyIndex      =   1
         OrigLeft        =   1695
         OrigTop         =   375
         OrigRight       =   1935
         OrigBottom      =   660
         Max             =   50
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1125
         MaxLength       =   30
         TabIndex        =   1
         Top             =   105
         Width           =   3300
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   3
         Top             =   450
         Width           =   450
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   6
         Top             =   450
         Width           =   450
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   1
         Left            =   3570
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196610
         BuddyIndex      =   2
         OrigLeft        =   3705
         OrigTop         =   405
         OrigRight       =   3945
         OrigBottom      =   690
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTableName 
         Caption         =   "表格名称(&N)"
         Height          =   240
         Left            =   75
         TabIndex        =   0
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label lblRow 
         Caption         =   "表格行数(&R)"
         Height          =   225
         Left            =   75
         TabIndex        =   2
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label lblCol 
         Caption         =   "表格列数(&C)"
         Height          =   210
         Left            =   2025
         TabIndex        =   5
         Top             =   510
         Width           =   1035
      End
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   1455
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":1582
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1590
      Top             =   4080
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
            Picture         =   "frmDefTable.frx":2B14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   915
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4290
      Left            =   2865
      ScaleHeight     =   4290
      ScaleWidth      =   6420
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2385
      Width           =   6420
      Begin VB.Timer tmr 
         Interval        =   60
         Left            =   2910
         Top             =   3315
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   210
         ScaleHeight     =   3390
         ScaleWidth      =   5505
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   5505
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
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
            Left            =   405
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   1785
            Visible         =   0   'False
            Width           =   2160
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
            Height          =   2805
            Left            =   240
            TabIndex        =   8
            Top             =   135
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            RowHeightMin    =   345
            BackColorSel    =   16771515
            BackColorBkg    =   -2147483628
            GridColor       =   -2147483636
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   0
            GridLinesFixed  =   1
            MergeCells      =   1
            AllowUserResizing=   3
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
      End
   End
   Begin VB.PictureBox picLvwBack 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   135
      ScaleHeight     =   3015
      ScaleWidth      =   2415
      TabIndex        =   13
      Top             =   1035
      Width           =   2415
      Begin MSComctlLib.ListView lvw 
         Height          =   2385
         Left            =   315
         TabIndex        =   17
         Top             =   240
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   4207
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "表格名称"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "行数"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "列数"
            Object.Width           =   1587
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6330
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      SimpleText      =   $"frmDefTable.frx":40A6
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDefTable.frx":40ED
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
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
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   5655
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":4981
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":4BA1
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":4DC1
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":4FE1
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5201
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5421
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":563B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5855
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5A6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5C89
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":5EA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":60BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":67B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":6EB1
            Key             =   "View"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":70CD
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":72ED
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6465
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":750D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":772D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":794D
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":7B6D
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":7D8D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":7FAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":81C7
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":83E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":8601
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":8821
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":8A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":8C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":934F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":9A49
            Key             =   "View"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":9C65
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTable.frx":9E85
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8880
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存"
               Object.Tag             =   "保存"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "取消"
               Object.ToolTipText     =   "取消"
               Object.Tag             =   "取消"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "合并"
               Key             =   "合并"
               Object.ToolTipText     =   "合并"
               Object.Tag             =   "合并"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "撤消"
               Key             =   "撤消"
               Object.ToolTipText     =   "撤消"
               Object.Tag             =   "撤消"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "字体"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "颜色"
               Key             =   "颜色"
               Object.ToolTipText     =   "颜色"
               Object.Tag             =   "颜色"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "水平"
               Key             =   "水平"
               Object.ToolTipText     =   "水平对齐"
               Object.Tag             =   "水平"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "垂直"
               Key             =   "垂直"
               Object.ToolTipText     =   "垂直对齐"
               Object.Tag             =   "垂直"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "查看"
               Object.ToolTipText     =   "表格查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   14
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   16
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image picX 
      Height          =   1530
      Left            =   2385
      MousePointer    =   9  'Size W E
      Top             =   885
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUpdatePage 
         Caption         =   "更新查询页面(&U)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "增加(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "取消(&C)"
      End
   End
   Begin VB.Menu mnuDesign 
      Caption         =   "格式(&O)"
      Begin VB.Menu mnuDesignInsert 
         Caption         =   "插入(&I)"
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "列(在左侧)(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "列(在右侧)(&R)"
            Index           =   1
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "行(在上方)(&A)"
            Index           =   2
         End
         Begin VB.Menu mnuDesignInsertTable 
            Caption         =   "行(在下方)(&B)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuDesignDel 
         Caption         =   "删除(&D)"
         Begin VB.Menu mnuDesignDelTable 
            Caption         =   "列(&C)"
            Index           =   0
         End
         Begin VB.Menu mnuDesignDelTable 
            Caption         =   "行(&R)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDesign_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignAutoMerge 
         Caption         =   "允许合并(&W)"
         Begin VB.Menu mnuAutoMergeCol 
            Caption         =   "列(&C)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAutoMergeRow 
            Caption         =   "行(&R)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDesignMerge 
         Caption         =   "合并单元(&M)"
      End
      Begin VB.Menu mnuDesignMergeCancel 
         Caption         =   "撤消合并(&Z)"
      End
      Begin VB.Menu mnuDesign_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignFont 
         Caption         =   "显示字体(&F)"
      End
      Begin VB.Menu mnuDesignColor 
         Caption         =   "文字颜色(&C)"
      End
      Begin VB.Menu mnuDesignLineColor 
         Caption         =   "表格颜色(&L)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesign_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignHsb 
         Caption         =   "水平对齐(&H)"
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "左边对齐(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "居中对齐(&C)"
            Index           =   1
         End
         Begin VB.Menu mnuHsbAlign 
            Caption         =   "右边对齐(&R)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDesignVsb 
         Caption         =   "垂直对齐(&V)"
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "顶部对齐(&T)"
            Index           =   0
         End
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "居中对齐(&C)"
            Index           =   1
         End
         Begin VB.Menu mnuVsbAlign 
            Caption         =   "底部对齐(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDesign_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignSize 
         Caption         =   "统一尺寸(&S)"
         Begin VB.Menu mnuSize 
            Caption         =   "相同列宽(&W)"
            Index           =   0
         End
         Begin VB.Menu mnuSize 
            Caption         =   "相同行高(&H)"
            Index           =   1
         End
         Begin VB.Menu mnuSize 
            Caption         =   "两者都相同(&B)"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单2"
      Visible         =   0   'False
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "左边对齐(&L)"
         Index           =   0
      End
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "居中对齐(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShort2Hsb 
         Caption         =   "右边对齐(&R)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuShort3 
      Caption         =   "快捷菜单3"
      Visible         =   0   'False
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "顶部对齐(&T)"
         Index           =   0
      End
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "居中对齐(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShort3Vsb 
         Caption         =   "底部对齐(&B)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmDefTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'本模块中所用到的局部变量说明
Private mblnFirst As Boolean                      '是否为初次进入本模块(True:初次进入;False:不是初次进入)
Private mintColumn As Integer

Private mSelStartRow As Long
Private mSelEndRow As Long
Private mSelStartCol As Long
Private mSelEndCol As Long

Private mSvrMouseX As Long
Private mSvrMouseY As Long

Private mSvrRow As Long
Private mSvrCol As Long


Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
'    DoEvents
    
    
    '界面显示后的数据初始化工作
    Call AdjustEnabled
    Call mnuViewRefresh_Click
    Call DrawRuler
End Sub

Private Sub Form_Load()
    '界面显示前的数据初始化工作
    mblnFirst = True
    
    RestoreWinState Me, App.ProductName
        
    picX.Width = 45
                                
    Call mnuViewIcon_Click(lvw.View)
    
    Call ReadRegister
    Call Reset
    Call ModulePrivs
            
End Sub

Private Sub Form_Resize()
    '根据窗体状态,调整窗体中各控件的显示位置
    Dim sglCbrH As Single
    Dim sglStbH As Single
    
    On Error Resume Next
    sglCbrH = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sglStbH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    Call ResizeControl(picLvwBack, 0, sglCbrH, picX.Left, Me.ScaleHeight - sglStbH - sglCbrH)
    
    Call ResizeControl(picEdit, picX.Left + picX.Width, picLvwBack.Top, Me.ScaleWidth - picX.Left - picX.Width, picEdit.Height)
    Call ResizeControl(txt(0), lblTableName.Left + lblTableName.Width, txt(0).Top, picEdit.ScaleWidth - lblTableName.Left - lblTableName.Width - 60, txt(0).Height)
    
    Call ResizeControl(picBack, picEdit.Left, picEdit.Top + picEdit.Height + 30, Me.ScaleWidth - picEdit.Left, Me.ScaleHeight - sglStbH - picEdit.Top - picEdit.Height - 30)
    
    picX.Top = picLvwBack.Top
    picX.Height = picLvwBack.Height
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If picBack.Tag = "1" Then
        If MsgBox("修改后的表格要保存才生效，确认不保存就退出吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call zlCommFun.OpenIme
    Call WriteRegister
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw.SortKey = mintColumn
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '
    Call Reset
    Call ShowTable(Val(Mid(lvw.SelectedItem.Key, 2)))
    picBack.Tag = ""
    Call AdjustEnabled
    
    Call LoadStatus
    
End Sub

Private Sub CellAlign()
    '设置指定单元格的对齐方式,可以一次指定多个单元格
    Dim i As Long
    Dim j As Long
    Dim Index As Long
    
    For i = 0 To mnuHsbAlign.UBound
        If mnuHsbAlign(i).Checked Then Index = i * 3
    Next
    For i = 0 To mnuVsbAlign.UBound
        If mnuVsbAlign(i).Checked Then Index = Index + i
    Next
        
    msf.Redraw = False
    Call SaveRowCol
    
    For i = 1 To msf.Rows - 1
        msf.Row = i
        For j = 1 To msf.Cols - 1
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then
                msf.CellAlignment = Index
                picBack.Tag = "1"
            End If
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
End Sub

Private Sub mnuAutoMergeCol_Click()
    Dim j As Long
    
    mnuAutoMergeCol.Checked = Not mnuAutoMergeCol.Checked
    
    msf.Redraw = False
    Call SaveRowCol
            
    For j = 1 To msf.Cols - 1
        msf.Col = j
        If msf.CellBackColor = msf.BackColorSel Then
            msf.MergeCol(j) = mnuAutoMergeCol.Checked
            picBack.Tag = "1"
        End If
    Next

    Call RestoreRowCol
    msf.Redraw = True
    
    'msf.MergeCol(msf.Col) = mnuAutoMergeCol.Checked
    Call AdjustEnabled
    
End Sub

Private Sub mnuAutoMergeRow_Click()
    Dim j As Long
    
    mnuAutoMergeRow.Checked = Not mnuAutoMergeRow.Checked
    
    msf.Redraw = False
    Call SaveRowCol
            
    For j = 1 To msf.Rows - 1
        msf.Row = j
        If msf.CellBackColor = msf.BackColorSel Then
            msf.MergeRow(j) = mnuAutoMergeRow.Checked
            picBack.Tag = "1"
        End If
    Next
    Call RestoreRowCol
    msf.Redraw = True
    
'    msf.MergeRow(msf.Row) = mnuAutoMergeRow.Checked
    Call AdjustEnabled
'    picBack.Tag = "1"
End Sub

Private Sub mnuDesignColor_Click()
    '设置指定单元格的文字颜色,可以一次指定多个单元格
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H1
    dlg.Color = msf.CellForeColor
    dlg.ShowColor
    If Err.Number = 0 Then
        msf.Redraw = False
        Call SaveRowCol
            
        For i = 1 To msf.Rows - 1
            For j = 1 To msf.Cols - 1
                msf.Row = i
                msf.Col = j
                If msf.CellBackColor = msf.BackColorSel Then
                    msf.CellForeColor = dlg.Color
                    picBack.Tag = "1"
                End If
            Next
        Next
        Call RestoreRowCol
        msf.Redraw = True
    Else
        Err.Clear
    End If
End Sub

Private Sub mnuDesignDelTable_Click(Index As Integer)
    '删除指定的行或列，可以一次指定多行或多列要删除
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    
    msf.Redraw = False
    Call SaveRowCol
    
    If mSelStartCol > mSelEndCol Then Call ExChange(mSelStartCol, mSelEndCol)
    If mSelStartRow > mSelEndRow Then Call ExChange(mSelStartRow, mSelEndRow)
    
    Select Case Index
    Case 0
        For j = 1 To msf.Rows - 1
            For i = mSelStartCol To msf.Cols - 1 - (mSelEndCol - mSelStartCol)
                msf.TextMatrix(j, i) = msf.TextMatrix(j, i + 1 + mSelEndCol - mSelStartCol)
            Next
        Next
        msf.Cols = IIf((msf.Cols - mSelEndCol + mSelStartCol - 1) < 2, 2, msf.Cols - mSelEndCol + mSelStartCol - 1)
    Case 1
        For j = 1 To msf.Cols - 1
            For i = mSelStartRow To msf.Rows - 1 - (mSelEndRow - mSelStartRow)
                msf.TextMatrix(i, j) = msf.TextMatrix(i + 1 + mSelEndRow - mSelStartRow, j)
            Next
        Next
        msf.Rows = IIf((msf.Rows - mSelEndRow + mSelStartRow - 1) < 2, 2, msf.Rows - mSelEndRow + mSelStartRow - 1)
    End Select
    
    Call AdjustNo
    picBack.Tag = "1"
    
    Call RestoreRowCol
    msf.Redraw = True
    
End Sub

Private Sub mnuDesignFont_Click()
    '设置指定单元格的字体属性,可以一次指定多个单元格
    Dim i As Long
    Dim j As Long
            
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H3 Or &H100 Or &H400 Or &H200 Or &H10000
    
    dlg.FontName = msf.CellFontName
    dlg.FontSize = msf.CellFontSize
    dlg.FontBold = msf.CellFontBold
    dlg.FontItalic = msf.CellFontItalic
    dlg.FontStrikethru = msf.CellFontStrikeThrough
    dlg.FontUnderline = msf.CellFontUnderline
    dlg.Color = msf.CellForeColor
    dlg.ShowFont
    If Err.Number = 0 Then
        msf.Redraw = False
        Call SaveRowCol
                        
        For i = 1 To msf.Rows - 1
            For j = 1 To msf.Cols - 1
                msf.Row = i
                msf.Col = j
                If msf.CellBackColor = msf.BackColorSel Then
                    msf.CellFontName = dlg.FontName
                    msf.CellFontSize = dlg.FontSize
                    msf.CellFontBold = dlg.FontBold
                    msf.CellFontItalic = dlg.FontItalic
                    msf.CellFontStrikeThrough = dlg.FontStrikethru
                    msf.CellFontUnderline = dlg.FontUnderline
                    msf.CellForeColor = dlg.Color
                    picBack.Tag = "1"
                End If
            Next
        Next
        Call RestoreRowCol
        msf.Redraw = True
    Else
        Err.Clear
    End If

End Sub

Private Sub mnuDesignInsertTable_Click(Index As Integer)
    Dim i As Long
    Dim intRow As Long
    Dim intCol As Long
    
    Select Case Index
    Case 0
        msf.Cols = msf.Cols + 1
        Call MoveColData(msf.Col)
    Case 1
        msf.Cols = msf.Cols + 1
        Call MoveColData(msf.Col + 1)
    Case 2
        msf.Rows = msf.Rows + 1
        Call MoveRowData(msf.Row)
    Case 3
        msf.Rows = msf.Rows + 1
        Call MoveRowData(msf.Row + 1)
    End Select
    
    Call AdjustNo
    
    txt(1).Text = msf.Rows - 1
    txt(2).Text = msf.Cols - 1
    
    picBack.Tag = "1"
    
End Sub

Private Sub mnuDesignLineColor_Click()
    '设置表格的网格线颜色
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H1
    dlg.Color = msf.CellForeColor
    dlg.ShowColor
    If Err.Number = 0 Then
        msf.Redraw = False
        msf.GridColor = dlg.Color
        msf.Redraw = True
        picBack.Tag = "1"
    Else
        Err.Clear
    End If

End Sub

Private Sub mnuDesignMerge_Click()
    Dim strText As String
    
    '1.检查是否可以合并行或列
    If CheckIsMerge = False Then Exit Sub
    
    '2.合并行或列
    strText = msf.TextMatrix(msf.Row, msf.Col)
    If frmDefTableMerge.ShowMergeBox(Me, strText) Then
        Call MergeCell(strText)
        picBack.Tag = "1"
    End If
End Sub

Private Sub mnuDesignMergeCancel_Click()
    '撤消合并单元格
    
    '1.检查是否为合并行或列
    If CheckIsMerge = False Then Exit Sub
    
    '2.撤消行或列
    Call CancelMergeCell
    picBack.Tag = "1"
End Sub

Private Sub mnuEditCancel_Click()
    '取消对表格的修改或增加
    'picLvwBack.Tag=1表示新增表格;picLvwBack.Tag=2表示修改表格
    
    If picBack.Tag = "1" Then
        If MsgBox("修改后的表格要保存才生效，确认不保存就退出吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    picBack.Tag = ""
    Call Reset
    If picLvwBack.Tag <> "" And Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
        
    picLvwBack.Tag = ""
    picLvwBack.Enabled = True
    picEdit.Enabled = False
    
    Call AdjustEnabled
    
End Sub

Private Sub mnuEditDelete_Click()
    '删除选定的表格
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("确认要删除表格[" & lvw.SelectedItem.Text & "]吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errHand
    
    gstrSQL = "zl_咨询表格元素_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    lvw.ListItems.Remove lvw.SelectedItem.Index
    If Not (lvw.SelectedItem Is Nothing) Then
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        Call Reset
    End If
    Call AdjustEnabled
    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditModify_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    picLvwBack.Enabled = False
    picLvwBack.Tag = "2"
    picEdit.Enabled = True
    
    Call AdjustEnabled
    Call LoadStatus
    picBack.Tag = ""
    
    txt(0).SetFocus
End Sub

Private Sub mnuEditNew_Click()
    '新增加表格
    
    Call Reset
    
    picLvwBack.Enabled = False
    picLvwBack.Tag = "1"
    picEdit.Enabled = True
    
    Call CreateDefaultTable
    Call AdjustNo
    msf.Visible = True
    
    Call AdjustEnabled
    Call LoadStatus
    picBack.Tag = ""
    
    txt(0).SetFocus
    
End Sub

Private Sub mnuEditSave_Click()
    '保存表格元素及内容设置
    Dim lng序号 As Long
    Dim strSQL(4) As String
    Dim vRowHeight As String
    Dim vColWidth As String
    Dim vMergeRow As String
    Dim vMergeCol As String
    Dim i As Long
    Dim j As Long
    Dim Itmx As ListItem
    
    If StrIsValid(txtInput.Text, txtInput.MaxLength) = False Then Exit Sub
    If txt(0).Text = "" Then
        MsgBox "必须要输入当前表格的名称！", vbInformation, gstrSysName
        Exit Sub
    End If
    For i = 1 To msf.Rows - 1
        vRowHeight = vRowHeight & ";" & msf.RowHeight(i)
        If msf.MergeRow(i) Then vMergeRow = vMergeRow & ";" & i
    Next
    vRowHeight = Mid(vRowHeight, 2)
    If vMergeRow <> "" Then vMergeRow = Mid(vMergeRow, 2)
    
    For i = 1 To msf.Cols - 1
        vColWidth = vColWidth & ";" & msf.ColWidth(i)
        If msf.MergeCol(i) Then vMergeCol = vMergeCol & ";" & i
    Next
    vColWidth = Mid(vColWidth, 2)
    If vMergeCol <> "" Then vMergeCol = Mid(vMergeCol, 2)
    
    If picLvwBack.Tag = "1" Then
        lng序号 = Val(MaxValue("咨询表格元素", "序号")) + 1
        strSQL(0) = "zl_咨询表格元素_insert(" & lng序号 & ",'" & txt(0).Text & "'," & msf.Cols - 1 & ",'" & vColWidth & "'," & msf.Rows - 1 & ",'" & vRowHeight & "','" & vMergeRow & "','" & vMergeCol & "')"
    Else
        lng序号 = Val(Mid(lvw.SelectedItem.Key, 2))
        strSQL(1) = "zl_咨询表格元素_update(" & lng序号 & ",'" & txt(0).Text & "'," & msf.Cols - 1 & ",'" & vColWidth & "'," & msf.Rows - 1 & ",'" & vRowHeight & "','" & vMergeRow & "','" & vMergeCol & "')"
        strSQL(2) = "zl_咨询表格内容_delete(" & lng序号 & ")"
    End If
    
    msf.Redraw = False
    Call SaveRowCol
    
    On Error GoTo errHand
    gcnOracle.BeginTrans
    For i = 0 To 2
        If strSQL(i) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        End If
    Next
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            gstrSQL = "zl_咨询表格内容_insert(" & lng序号 & "," & i & "," & j & ",'" & msf.TextMatrix(i, j) & "'," & msf.CellAlignment & "," & msf.CellForeColor & ",'" & msf.CellFontName & ";" & msf.CellFontSize & ";" & msf.CellFontBold & ";" & msf.CellFontItalic & ";" & msf.CellFontStrikeThrough & ";" & msf.CellFontUnderline & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    Next
    
    picBack.Tag = ""
    
    gcnOracle.CommitTrans
    Call RestoreRowCol
    msf.Redraw = True
            
    If picLvwBack.Tag = "1" Then
        Set Itmx = lvw.ListItems.Add(, "K" & lng序号, txt(0).Text, 1, 1)
        Itmx.SubItems(1) = msf.Rows - 1
        Itmx.SubItems(2) = msf.Cols - 1
        Itmx.Selected = True
    Else
        Call mnuViewRefresh_Click
    End If
    
    picLvwBack.Tag = ""
    picLvwBack.Enabled = True
    picEdit.Enabled = False
    
    Call AdjustEnabled
    
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditSelectAll_Click()
    If txtInput.Visible Then Exit Sub
    Call SelectRect(1, 1, msf.Rows - 1, msf.Cols - 1)
End Sub

Private Sub mnuFileExcel_Click()
    Call PrintObject(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFilePreView_Click()
    Call PrintObject(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintObject(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileUpdatePage_Click()
    
    Call gfrmMain.FrameDefault.RefreshPage
    
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuHsbAlign_Click(Index As Integer)
    
    Call SelectHsbAlign(Index)
    
    Call CellAlign
End Sub

Private Sub mnuShort2Hsb_Click(Index As Integer)
    Call mnuHsbAlign_Click(Index)
End Sub

Private Sub mnuShort3Vsb_Click(Index As Integer)
    Call mnuVsbAlign_Click(Index)
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim i As Long
            
    msf.Redraw = False
    Call SaveRowCol
    Select Case Index
    Case 0          '相同列宽
        For i = 1 To msf.Cols - 1
            msf.Col = i
            If msf.CellBackColor = msf.BackColorSel Then
                msf.ColWidth(i) = msf.ColWidth(mSvrCol)
                picBack.Tag = "1"
            End If
        Next
    Case 1          '相同行高
        For i = 1 To msf.Rows - 1
            msf.Row = i
            If msf.CellBackColor = msf.BackColorSel Then
                msf.RowHeight(i) = msf.RowHeight(mSvrRow)
                picBack.Tag = "1"
            End If
        Next
    Case 2          '两者都相同
        For i = 1 To msf.Cols - 1
            msf.Col = i
            If msf.CellBackColor = msf.BackColorSel Then
                msf.ColWidth(i) = msf.ColWidth(mSvrCol)
                picBack.Tag = "1"
            End If
        Next
        For i = 1 To msf.Rows - 1
            msf.Row = i
            If msf.CellBackColor = msf.BackColorSel Then
                msf.RowHeight(i) = msf.RowHeight(mSvrRow)
                picBack.Tag = "1"
            End If
        Next
    End Select
    Call RestoreRowCol
    msf.Redraw = True
    
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
    
End Sub

Private Sub mnuViewRefresh_Click()
    '
    Dim svrKey As String
        
    svrKey = SaveLvwItem(lvw)
    Call LoadDefTable
    Call RestoreLvwItem(lvw, svrKey)
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call lvw_ItemClick(lvw.SelectedItem)
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub


Private Sub mnuViewToolText_Click()
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mnuVsbAlign_Click(Index As Integer)
    
    Call SelectVsbAlign(Index)
    Call CellAlign
    
End Sub

Private Sub msf_DblClick()
    If picLvwBack.Tag = "" Then Exit Sub
    If msf.Row > 0 And msf.Col > 0 Then Call LocationTxt(msf.Row, msf.Col)
End Sub

Private Sub msf_EnterCell()
    Dim vIndex As Long
    If picLvwBack.Tag = "" Then Exit Sub

    '
    vIndex = msf.CellAlignment
    Select Case vIndex
    Case 0, 1, 2
        Call SelectHsbAlign(0)
        Call SelectVsbAlign(vIndex)
    Case 3, 4, 5
        Call SelectHsbAlign(1)
        Call SelectVsbAlign(vIndex - 3)
    Case 6, 7, 8
        Call SelectHsbAlign(2)
        Call SelectVsbAlign(vIndex - 6)
    End Select
    
    mnuAutoMergeRow.Checked = msf.MergeRow(msf.Row)
    mnuAutoMergeCol.Checked = msf.MergeCol(msf.Col)

    Dim X1 As Long
    Dim X2 As Long

    Call ClearCellBackColor
    
    msf.Redraw = False
    Call SaveRowCol

    msf.Row = 0
    msf.CellForeColor = &HFF0000
    msf.Row = mSvrRow
    msf.Col = 0
    msf.CellForeColor = &HFF0000

    Call RestoreRowCol
    msf.Redraw = True
    
    Call AdjustEnabled
    
    Select Case CalcMergeArea(msf.Row, msf.Col, X1, X2)
    Case 0
        msf.CellBackColor = msf.BackColorSel
    Case 1
        Call SelectRect(msf.Row, X1, msf.Row, X2)
    Case 2
        Call SelectRect(X1, msf.Col, X2, msf.Col)
    End Select
    
    
End Sub

Private Sub msf_GotFocus()
    If picLvwBack.Tag = "" Then Exit Sub
    zlCommFun.OpenIme True
End Sub

Private Sub msf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim j As Long
    If picLvwBack.Tag = "" Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        msf.Redraw = False
        Call SaveRowCol
            
        For i = 1 To msf.Rows - 1
            For j = 1 To msf.Cols - 1
                msf.Row = i
                msf.Col = j
                If msf.CellBackColor = msf.BackColorSel Then
                    msf.TextMatrix(i, j) = ""
                    picBack.Tag = "1"
                End If
            Next
        Next
        Call RestoreRowCol
        msf.Redraw = True
    Else
        If KeyCode = vbKeyReturn Then Exit Sub
        If KeyCode = vbKeyRight Then Exit Sub
        If KeyCode = vbKeyDelete Then Exit Sub
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyNumlock Then Exit Sub
        If (KeyCode >= vbKeyF1 And KeyCode < vbKeyF12) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyMultiply Or Shift = vbCtrlMask Or Shift = vbShiftMask Or Shift = vbAltMask Then Exit Sub

        Call msf_DblClick
        picBack.Tag = "1"

    End If
End Sub

Private Sub MSF_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim j As Long
    
    If picLvwBack.Tag = "" Then Exit Sub
    If KeyAscii = 32 Then
        Call msf_DblClick
    ElseIf KeyAscii = 13 Then
        Call msf_LeaveCell
        Call NextCell(msf.Row, msf.Col)
        Call msf_EnterCell
    Else
        If CheckIsInclude(UCase(Chr(KeyAscii)), "可打印字符") Then
            Call msf_DblClick

            txtInput.Text = Chr(KeyAscii)
            picBack.Tag = "1"
            SendKeys "{END}"
        End If
    End If

End Sub

Private Sub msf_LeaveCell()
    If picLvwBack.Tag = "" Then Exit Sub
    If txtInput.Visible Then
        txtInput.Visible = False
        Call MergeCell(txtInput.Text)
    End If
    
    msf.Redraw = False
    Call SaveRowCol
    
    msf.Row = 0
    msf.CellForeColor = 0
    msf.Row = mSvrRow
    msf.Col = 0
    msf.CellForeColor = 0
    
    Call RestoreRowCol
    msf.Redraw = True
    
End Sub


Private Sub msf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picLvwBack.Tag = "" Then Exit Sub
    If Button <> 1 Then Exit Sub
    mSelStartRow = msf.MouseRow
    mSelStartCol = msf.MouseCol
End Sub

Private Sub msf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picLvwBack.Tag = "" Then Exit Sub
    If Button = 1 Then
        If mSvrMouseX <> msf.MouseRow Or mSvrMouseY <> msf.MouseCol Then
            If mSelStartRow <> msf.MouseRow Or mSelStartCol <> msf.MouseCol Then Call SelectRect(mSelStartRow, mSelStartCol, msf.MouseRow, msf.MouseCol)
            mSvrMouseX = msf.MouseRow
            mSvrMouseY = msf.MouseCol
        End If
    End If
End Sub

Private Sub msf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picLvwBack.Tag = "" Then Exit Sub
    If Button = 1 Then
        mSelEndRow = msf.MouseRow
        mSelEndCol = msf.MouseCol
        Call AdjustEnabled
    End If
    If Button = 2 Then Me.PopupMenu mnuDesign
End Sub

Private Sub picBack_Resize()
    '
    On Error Resume Next
    Call ResizeControl(picDraw, 0, 0, picBack.Width, picBack.Height)
    Call ResizeControl(msf, 0, 0, picDraw.Width, picDraw.Height)
End Sub

Private Sub picEdit_Paint()
    RaisEffect picEdit, -1, "", 0
End Sub

Private Sub picLvwBack_Resize()
    '调整此图片框中的各控件的显示比例
    On Error Resume Next
    
    Call ResizeControl(lvw, 0, 0, picLvwBack.Width, picLvwBack.Height)
    
End Sub

Private Sub picX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picX.Left = picX.Left + X
    If picX.Left < 1500 Then picX.Left = 1500
    If Me.Width - picX.Left - picX.Width < 3000 Then picX.Left = Me.Width - picX.Width - 3000
    
    Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "预览"
        Call mnuFilePreView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "增加"
        Call mnuEditNew_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "查看"
        If lvw.View < 3 Then
            Call mnuViewIcon_Click(lvw.View + 1)
        Else
            Call mnuViewIcon_Click(0)
        End If
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "保存"
        Call mnuEditSave_Click
    Case "取消"
        Call mnuEditCancel_Click
    Case "合并"
        Call mnuDesignMerge_Click
    Case "撤消"
        Call mnuDesignMergeCancel_Click
    Case "字体"
        Call mnuDesignFont_Click
    Case "颜色"
        Call mnuDesignColor_Click
    Case "水平"
        Me.PopupMenu mnuShort2
    Case "垂直"
        Me.PopupMenu mnuShort3
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "大图标"
        Call mnuViewIcon_Click(0)
    Case "小图标"
        Call mnuViewIcon_Click(1)
    Case "列表"
        Call mnuViewIcon_Click(2)
    Case "详细资料"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub tmr_Timer()
    '行高或列宽有变动时，自动调整显示控件的位置
    Dim i As Long
    Dim blnChange As Boolean
    
       
    blnChange = False
    For i = 0 To msf.Rows - 1
        If msf.RowHeight(i) <> msf.RowData(i) Then
            If i = 0 Then
                msf.RowHeight(i) = 300
                msf.RowData(i) = msf.RowHeight(i)
                blnChange = True
            Else
                msf.RowData(i) = msf.RowHeight(i)
                blnChange = True
            End If
        End If
    Next
    
    For i = 0 To msf.Cols - 1
        If msf.ColWidth(i) <> msf.ColData(i) Then
            If i = 0 Then
                msf.ColWidth(i) = 600
                msf.ColData(i) = msf.ColWidth(i)
                blnChange = True
            Else
                msf.ColData(i) = msf.ColWidth(i)
                blnChange = True
            End If
        End If
    Next
    
    If txtInput.Visible = False Then Exit Sub
    If blnChange = True Then Call LocationTxt(msf.Row, msf.Col)
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 0 Then Call LoadStatus
    picBack.Tag = "1"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
    If Index = 0 Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    If Index <> 0 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "正整数") = True Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 Then zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim i As Long
    Dim j As Long
    Dim vNewStartRow As Long
    Dim vNewStartCol As Long
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Index <> 0 And Cancel = False Then
        If Val(txt(Index).Text) <= 0 Then
            MsgBox "行列数不能为0，至少需要一行或列！", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        If Val(txt(Index).Text) > 50 And Index = 1 Then
            MsgBox "行数不能大于50行！", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        If Val(txt(Index).Text) > 30 And Index = 2 Then
            MsgBox "列数不能大于30行！", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        Select Case Index
        Case 1
            If msf.Rows < Val(txt(1).Text) + 1 Then vNewStartRow = msf.Rows
            If Val(txt(1).Text) + 1 <> msf.Rows Then msf.Rows = Val(txt(1).Text) + 1
        Case 2
            If msf.Cols < Val(txt(2).Text) + 1 Then vNewStartCol = msf.Cols
            If msf.Cols <> Val(txt(2).Text) + 1 Then msf.Cols = Val(txt(2).Text) + 1
        End Select
        
        If vNewStartCol > 1 Then
            Call SaveRowCol
            msf.Redraw = False
            For i = 1 To msf.Rows - 1
                msf.Row = i
                For j = vNewStartCol To msf.Cols - 1
                    msf.Col = j
                    msf.CellAlignment = 1
                Next
            Next
            msf.Redraw = True
            Call RestoreRowCol
        End If
        
        Call AdjustNo
        If Index = 2 Then Call CheckColWidth
    End If
End Sub

Private Sub txtInput_Change()
    picBack.Tag = "1"
    msf.TextMatrix(msf.Row, msf.Col) = txtInput.Text
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If StrIsValid(txtInput.Text, txtInput.MaxLength) = False Then Exit Sub
        
        Call msf_LeaveCell
        msf.SetFocus
        Call MSF_KeyPress(13)
    End If
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txtInput.Text, txtInput.MaxLength)
    
    '检查是否是合并区，如果是合并区，单元格内容不能为空，至少要有一空格，以保证合并区的存在
    Dim i As Long
    Dim j As Long
    Dim k As Boolean
    Dim intRow As Long
    Dim intCol As Long
    
    msf.Redraw = False
    Call SaveRowCol
    
    intRow = msf.Row
    intCol = msf.Col
    
    If intRow - 1 > 0 Then
        msf.Row = intRow - 1
        If msf.CellBackColor = msf.BackColorSel Then k = True
    End If
    If intRow + 1 < msf.Rows Then
        msf.Row = intRow + 1
        If msf.CellBackColor = msf.BackColorSel Then k = True
    End If
    If intCol - 1 > 0 Then
        msf.Col = intCol - 1
        If msf.CellBackColor = msf.BackColorSel Then k = True
    End If
    If intCol + 1 > msf.Cols Then
        msf.Col = intCol + 1
        If msf.CellBackColor = msf.BackColorSel Then k = True
    End If
        
    Call RestoreRowCol
    msf.Redraw = True
    If k And txtInput.Text = "" Then
        MsgBox "合并区内容不能为空，至少要有一空格，以保证合并区有效！", vbOKOnly + vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub udn_Change(Index As Integer)
    If Index = 0 Then msf.Rows = Val(txt(1).Text) + 1
    If Index = 1 Then msf.Cols = Val(txt(2).Text) + 1
    Call AdjustNo
    If Index = 1 Then Call CheckColWidth
End Sub

'-----------------------------------------------------------------------------------------------------------------
'
'以下是自定义函数或过程部份,仅供本模块所使用
'
'-----------------------------------------------------------------------------------------------------------------
Private Sub ModulePrivs()
    '根据模块权限,处理功能项的隐藏或显示
    '权限有:增删改
    
'    mnuEdit.Visible = True
'    mnuDesign.Visible = True
'
'    If InStr(gstrPrivs, "增删改") = 0 Then
'        mnuEdit.Visible = False
'        mnuDesign.Visible = False
'
'        tbrThis.Buttons("增加").Visible = False
'        tbrThis.Buttons("修改").Visible = False
'        tbrThis.Buttons("删除").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'        tbrThis.Buttons("保存").Visible = False
'        tbrThis.Buttons("取消").Visible = False
'        tbrThis.Buttons("Split_3").Visible = False
'
'        tbrThis.Buttons("合并").Visible = False
'        tbrThis.Buttons("撤消").Visible = False
'        tbrThis.Buttons("字体").Visible = False
'        tbrThis.Buttons("颜色").Visible = False
'        tbrThis.Buttons("水平").Visible = False
'        tbrThis.Buttons("垂直").Visible = False
'        tbrThis.Buttons("Split_4").Visible = False
'    End If
End Sub

Private Sub DrawRuler()
'    '作横向标尺和纵向标尺;1逻辑厘米=567缇(vb单位)
'    Dim i As Long
'    Dim blnDraw As Boolean
'
'    '1.画横向标尺该度
'    i = 0
'    blnDraw = True
'    picVsb.Cls
'    While blnDraw
'        i = i + 1
'        Call DrawLine(picVsb, i * 567, picLR.Top + picLR.Height, i * 567, picLR.Top + picLR.Height + 1000, RGB(255, 255, 255))
'        Call DrawText(picLR, i * 567 - picLR.TextWidth(CStr(i)) / 2, (picLR.Height - picLR.TextHeight(CStr(i))) / 2, i, &HFF0000)
'        If (i + 1) * 567 > picVsb.Width Then blnDraw = False
'    Wend
'
'    '2.画纵向标尺该度;数字要逆时针旋转90度
'    i = 0
'    blnDraw = True
'    picHsb.Cls
'    While blnDraw
'        i = i + 1
'        Call DrawLine(picHsb, picTB.Left + picTB.Width, i * 567, picTB.Left + picTB.Width + 1000, i * 567, RGB(255, 255, 255))
'        Call DrawText(picTB, (picTB.Width - picTB.TextWidth(CStr(i))) / 2 - 60, i * 567 - picTB.TextHeight(CStr(i)) / 2 + 150, i, &HFF0000, 90)
'        If (i + 1) * 567 > picHsb.Height Then blnDraw = False
'    Wend
End Sub

Private Sub CreateDefaultTable()
    '产生缺省的表格数据
    Dim i As Long
    Dim j As Long
    
    With msf
        .Rows = 8
        .Cols = 5
        .ColWidth(0) = 300
        .ColData(0) = 600
        For i = 1 To .Cols - 1
            .ColWidth(i) = 1200
            .ColData(i) = 1200
            .MergeCol(i) = False
        Next
        For i = 0 To .Rows - 1
            .RowData(i) = 300
            .RowHeight(i) = 300
            .MergeRow(i) = False
        Next
        txt(0).Text = "新查询表格"
        txt(1).Text = "7"
        txt(2).Text = "4"
        
        txtInput.FontName = "宋体"
        txtInput.FontSize = 12
        txtInput.FontBold = False
        txtInput.FontItalic = False
        txtInput.FontStrikethru = False
        txtInput.FontUnderline = False
        txtInput.ForeColor = 0
        
        Call SaveRowCol
        For i = 1 To msf.Rows - 1
            msf.Row = i
            For j = 1 To msf.Cols - 1
                msf.Col = j
                msf.CellFontName = "宋体"
                msf.CellFontSize = 12
                msf.CellFontBold = False
                msf.CellFontItalic = False
                msf.CellFontStrikeThrough = False
                msf.CellFontUnderline = False
                msf.CellForeColor = 0
                msf.CellAlignment = 1
            Next
        Next
        Call RestoreRowCol
    End With
End Sub

Private Sub ShowTable(ByVal No As Long)
    '显示表格到界面上
    Dim i As Long
    Dim strTmp As String
    Dim intPos As Long
    
    On Error GoTo errHand
    
    gstrSQL = "select 序号,名称,列数,列宽,行数,行高,合并行,合并列,颜色 from 咨询表格元素 where 序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, No)
    If gRs.BOF = False Then
        txt(0).Text = IIf(IsNull(gRs!名称), "", gRs!名称)
        
        If IIf(IsNull(gRs!行数), 0, gRs!行数) <= 0 Then
            MsgBox "错误的表格行数（行数小于1）！", vbInformation, gstrSysName
            Exit Sub
        End If
        If IIf(IsNull(gRs!列数), 0, gRs!列数) <= 0 Then
            MsgBox "错误的表格列数（行数小于1）！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        txt(1).Text = gRs!行数
        txt(2).Text = gRs!列数
        
        msf.Rows = gRs!行数 + 1
        msf.Cols = gRs!列数 + 1
        
        On Error Resume Next
        For i = 1 To msf.Rows - 1
            msf.RowHeight(i) = 300
            msf.MergeRow(i) = False
        Next
        For i = 1 To msf.Rows - 1
            msf.RowHeight(i) = Split(gRs!行高, ";")(i - 1)
            msf.MergeCol(i) = False
        Next
                        
        For i = 1 To msf.Cols - 1
            msf.ColWidth(i) = 1200
        Next
        For i = 1 To msf.Cols - 1
            msf.ColWidth(i) = IIf(Val(Split(gRs!列宽, ";")(i - 1)) = -1, 1200, Split(gRs!列宽, ";")(i - 1))
        Next
                                
        strTmp = IIf(IsNull(gRs!合并行), "", gRs!合并行 & ";")
        intPos = InStr(strTmp, ";")
        While intPos > 0
            msf.MergeRow(Val(Mid(strTmp, 1, intPos - 1))) = True
            strTmp = Mid(strTmp, intPos + 1)
            intPos = InStr(strTmp, ";")
        Wend

        strTmp = IIf(IsNull(gRs!合并列), "", gRs!合并列 & ";")
        intPos = InStr(strTmp, ";")
        While intPos > 0
            msf.MergeCol(Val(Mid(strTmp, 1, intPos - 1))) = True
            strTmp = Mid(strTmp, intPos + 1)
            intPos = InStr(strTmp, ";")
        Wend
        
        gstrSQL = "select 表号,行号,列号,内容,对齐,颜色,字体 from 咨询表格内容 where 表号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, No)
        If gRs.BOF = False Then
            While Not gRs.EOF
                msf.Row = gRs!行号
                msf.Col = gRs!列号
                msf.TextMatrix(msf.Row, msf.Col) = IIf(IsNull(gRs!内容), "", gRs!内容)
                msf.CellAlignment = IIf(IsNull(gRs!对齐), 9, gRs!对齐)
                msf.CellForeColor = IIf(IsNull(gRs!颜色), 0, gRs!颜色)
                msf.CellFontName = Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(0)
                msf.CellFontSize = Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(1)
                msf.CellFontBold = IIf(Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(2) = True, True, False)
                msf.CellFontItalic = IIf(Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(3) = True, True, False)
                msf.CellFontStrikeThrough = IIf(Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(4) = True, True, False)
                msf.CellFontUnderline = IIf(Split(IIf(IsNull(gRs!字体), "宋体;9;False;False;False;False", gRs!字体), ";")(5) = True, True, False)
                gRs.MoveNext
            Wend
        End If
        Call AdjustNo
        msf.Visible = True
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdjustNo(Optional ByVal blnRow As Boolean = True, Optional ByVal blnCol As Boolean = True)
    Dim i As Long
    
    msf.Redraw = False
    Call SaveRowCol
    If blnRow Then
        msf.Col = 0
        For i = 1 To msf.Rows - 1
            msf.Row = i
            msf.CellFontBold = True
            msf.TextMatrix(i, 0) = i
        Next
    End If
    
    If blnCol Then
        msf.Row = 0
        For i = 1 To msf.Cols - 1
            msf.Col = i
            msf.CellFontBold = True
            msf.TextMatrix(0, i) = i
            msf.ColAlignmentFixed(i) = 4
        Next
    End If
    msf.ColAlignmentFixed(0) = 4
    Call RestoreRowCol
    msf.Redraw = True
End Sub

Private Sub LocationTxt(ByVal Row As Long, ByVal Col As Long)
    Dim svrTag As String
        
    If msf.Visible = False Then Exit Sub
    With txtInput
        svrTag = picBack.Tag
        .Text = msf.TextMatrix(Row, Col)
        .ForeColor = msf.CellForeColor
        .FontName = msf.CellFontName
        .FontSize = msf.CellFontSize
        .FontBold = msf.CellFontBold
        .FontItalic = msf.CellFontItalic
        .FontStrikethru = msf.CellFontStrikeThrough
        .FontUnderline = msf.CellFontUnderline
        .Left = msf.CellLeft + msf.Left
        .Top = msf.CellTop + msf.Top
        .Width = msf.CellWidth
        .Height = msf.CellHeight
        .Visible = True
        picBack.Tag = svrTag
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub NextCell(ByVal Row As Long, ByVal Col As Long)
    '计算下一单元格
    Dim i As Long
    Dim intRow As Long
    Dim intCol As Long
    
    intRow = IIf(Col = msf.Cols - 1, IIf(Row = msf.Rows - 1, Row, Row + 1), Row)
    intCol = IIf(Col = msf.Cols - 1, IIf(Row = msf.Rows - 1, Col, 1), Col + 1)
        
    msf.Row = intRow
    msf.Col = intCol
    
End Sub

Private Sub MoveColData(ByVal intCol As Long)
    '从intCol列开始向后移动整列数据(包括intCol列)
    Dim i As Long
    Dim j As Long
    
    For j = 1 To msf.Rows - 1
        For i = msf.Cols - 1 To intCol + 1 Step -1
            msf.TextMatrix(j, i) = msf.TextMatrix(j, i - 1)
        Next
        msf.TextMatrix(j, intCol) = ""
    Next
End Sub

Private Sub MoveRowData(ByVal intRow As Long)
    '从intRow行开始向下移动整列数据(包括intRow行)
    Dim i As Long
    Dim j As Long
    
    For j = 1 To msf.Cols - 1
        For i = msf.Rows - 1 To intRow + 1 Step -1
            msf.TextMatrix(i, j) = msf.TextMatrix(i - 1, j)
        Next
        msf.TextMatrix(intRow, j) = ""
    Next
End Sub

Private Sub ExChange(X As Long, Y As Long)
    '交换X和Y的值
    Dim Tmp As Long
    
    Tmp = X
    X = Y
    Y = Tmp
End Sub

Private Sub ClearCellBackColor()
    '清除所有选定的单元格;实质是清除背景色
    Dim i As Long
    Dim j As Long
    
    msf.Redraw = False
    Call SaveRowCol
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then msf.CellBackColor = RGB(255, 255, 255)
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
End Sub

Private Sub SelectRect(ByVal Row1 As Long, ByVal Col1 As Long, ByVal Row2 As Long, ByVal Col2 As Long)
    '选定指定区域的单元格;实质是设置这些单元格的背景色
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    
    Dim i As Long
    Dim j As Long
    
    '1.检查行列的大小,进行交换处理
    X1 = Row1
    X2 = Row2
    If Row1 > Row2 Then Call ExChange(X1, X2)
    
    Y1 = Col1
    Y2 = Col2
    If Col1 > Col2 Then Call ExChange(Y1, Y2)
    
    '2.检查当前区域是否已经处于选中状态
    
    
    '3.设置区域的背景色,使之处于选中状态
    msf.Redraw = False
    Call SaveRowCol
    
    '清除已选择的区域
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then msf.CellBackColor = msf.BackColor
        Next
    Next
    
    
    For i = X1 To X2
        For j = Y1 To Y2
            msf.Row = i
            msf.Col = j
            If i > 0 And j > 0 Then
                msf.CellBackColor = msf.BackColorSel
            End If
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
End Sub

Private Sub MergeCell(ByVal strText As String)
    '合并指定区域的单元格,并附以内容
    
    Dim i As Long
    Dim j As Long
    
    msf.Redraw = False
    Call SaveRowCol
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then
                msf.TextMatrix(i, j) = strText
            End If
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
End Sub

Private Sub CancelMergeCell()
    '撤消合并指定区域的单元格
    
    Dim i As Long
    Dim j As Long
    Dim vFirst As Boolean
    
    vFirst = True
    msf.Redraw = False
    Call SaveRowCol
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then
                If vFirst = False Then msf.TextMatrix(i, j) = ""
                vFirst = False
            End If
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
End Sub


Private Function CheckIsMerge() As Boolean
    '检查当前选定的区域是否可以合并
    Dim i As Long
    Dim j As Long
    Dim vCol As Long
    Dim vRow As Long
    
    CheckIsMerge = False
    
    msf.Redraw = False
    Call SaveRowCol
    
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then vRow = vRow + 1
            If vRow > 1 Then Exit For
        Next
        If vRow > 1 Then Exit For
        vRow = 0
    Next
    
    For i = 1 To msf.Cols - 1
        For j = 1 To msf.Rows - 1
            msf.Row = j
            msf.Col = i
            If msf.CellBackColor = msf.BackColorSel Then vCol = vCol + 1
            If vCol > 1 Then Exit For
        Next
        If vCol > 1 Then Exit For
        vCol = 0
    Next
    
    Call RestoreRowCol
    msf.Redraw = True
    
    If vRow > 1 And vCol > 1 Then Exit Function
    If vRow = 0 And vCol = 0 Then Exit Function
    If vRow > 1 And msf.MergeRow(msf.Row) = False Then Exit Function
    If vCol > 1 And msf.MergeCol(msf.Col) = False Then Exit Function
    
    CheckIsMerge = True
End Function

Private Function CheckIsResizeWidth() As Boolean
    '检查当前选定的区域是否可以设置相同的宽度
    Dim i As Long
    Dim j As Long
    Dim vCol As Long
    
    msf.Redraw = False
    Call SaveRowCol
    
    For i = 1 To msf.Rows - 1
        For j = 1 To msf.Cols - 1
            msf.Row = i
            msf.Col = j
            If msf.CellBackColor = msf.BackColorSel Then vCol = vCol + 1
            If vCol > 1 Then Exit For
        Next
        If vCol > 1 Then Exit For
        vCol = 0
    Next
    
    Call RestoreRowCol
    msf.Redraw = True
    
    If vCol < 2 Then Exit Function
    
    CheckIsResizeWidth = True
End Function

Private Function CheckIsResizeHeight() As Boolean
    '检查当前选定的区域是否可以设置相同的高度
    Dim i As Long
    Dim j As Long
    Dim vRow As Long
            
    msf.Redraw = False
    Call SaveRowCol
    
    For i = 1 To msf.Cols - 1
        For j = 1 To msf.Rows - 1
            msf.Row = j
            msf.Col = i
            If msf.CellBackColor = msf.BackColorSel Then vRow = vRow + 1
            If vRow > 1 Then Exit For
        Next
        If vRow > 1 Then Exit For
        vRow = 0
    Next
    Call RestoreRowCol
    msf.Redraw = True
    
    If vRow < 2 Then Exit Function
    
    CheckIsResizeHeight = True
End Function

Private Function CalcMergeArea(ByVal Row As Long, ByVal Col As Long, StartPos As Long, EndPos As Long) As Byte
    '计算当前包含当前单元格的合并区域,如果没有合并,返回的是当前单元格
    Dim i As Long
        
    '1.查找是否有行与当前单元格合并
    If msf.MergeRow(Row) Then
        StartPos = Col
        EndPos = Col
        For i = Col - 1 To 1 Step -1
            If msf.TextMatrix(Row, i) <> msf.TextMatrix(Row, Col) Or msf.TextMatrix(Row, Col) = "" Then Exit For
            StartPos = i
        Next
        For i = Col + 1 To msf.Cols - 1
            If msf.TextMatrix(Row, i) <> msf.TextMatrix(Row, Col) Or msf.TextMatrix(Row, Col) = "" Then Exit For
            EndPos = i
        Next
        If StartPos <> Col Or EndPos <> Col Then
            CalcMergeArea = 1
            Exit Function
        End If
    End If
        
    '2.查找是否有列与当前单元格合并
    If msf.MergeCol(Col) Then
        StartPos = Row
        EndPos = Row
        For i = Row - 1 To 1 Step -1
            If msf.TextMatrix(i, Col) <> msf.TextMatrix(Row, Col) Or msf.TextMatrix(Row, Col) = "" Then Exit For
            StartPos = i
        Next
        For i = Row + 1 To msf.Rows - 1
            If msf.TextMatrix(i, Col) <> msf.TextMatrix(Row, Col) Or msf.TextMatrix(Row, Col) = "" Then Exit For
            EndPos = i
        Next
        If StartPos <> Row Or EndPos <> Row Then
            CalcMergeArea = 2
            Exit Function
        End If
    End If
    CalcMergeArea = 0
End Function

Private Sub AdjustEnabled()
    '调整功能菜单或按钮的可用状态
    mnuFilePreView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileExcel.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditNew.Enabled = True
    mnuEditSelectAll.Enabled = True
    mnuEditSave.Enabled = True
    mnuEditCancel.Enabled = True
    
    mnuDesignHsb.Enabled = True
    mnuDesignVsb.Enabled = True
    
    mnuDesignFont.Enabled = True
    mnuDesignColor.Enabled = True
    mnuDesignLineColor.Enabled = True
    mnuDesignInsert.Enabled = True
    mnuDesignDel.Enabled = True
    mnuDesignAutoMerge.Enabled = True
    mnuDesignMerge.Enabled = True
    mnuDesignMergeCancel.Enabled = True
    mnuDesignSize.Enabled = True
    mnuViewRefresh.Enabled = True
    mnuSize(0).Enabled = True
    mnuSize(1).Enabled = True
    mnuSize(2).Enabled = True
    
    If lvw.ListItems.Count = 0 Then
        mnuFilePreView.Enabled = False
        mnuFilePrint.Enabled = False
        mnuFileExcel.Enabled = False
    End If
    
    If lvw.SelectedItem Is Nothing Then
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
    End If
    
    If picLvwBack.Tag <> "" Then
        mnuEditNew.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuViewRefresh.Enabled = False
        txt(0).Locked = False
    Else
        mnuEditSave.Enabled = False
        mnuEditSelectAll.Enabled = False
        mnuEditCancel.Enabled = False
        txt(0).Locked = True
    End If
    
    If mnuEditSave.Enabled = False Then
        
        mnuDesignHsb.Enabled = False
        mnuDesignVsb.Enabled = False
    
        mnuDesignColor.Enabled = False
        mnuDesignLineColor.Enabled = False
        mnuDesignFont.Enabled = False
        mnuDesignInsert.Enabled = False
        mnuDesignDel.Enabled = False
        mnuDesignAutoMerge.Enabled = False
        mnuDesignMerge.Enabled = False
        mnuDesignMergeCancel.Enabled = False
        mnuDesignSize.Enabled = False
    End If
    
    '1.检查是否可以合并行或列
    If CheckIsMerge = False Then
        mnuDesignMerge.Enabled = False
        mnuDesignMergeCancel.Enabled = False
    End If
    
    If CheckIsResizeWidth = False Then
        mnuSize(0).Enabled = False
    End If
    If CheckIsResizeHeight = False Then
        mnuSize(1).Enabled = False
    End If
    
    If mnuSize(0).Enabled = False Or mnuSize(1).Enabled = False Then
        mnuSize(2).Enabled = False
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePreView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("增加").Enabled = mnuEditNew.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("保存").Enabled = mnuEditSave.Enabled
    tbrThis.Buttons("取消").Enabled = mnuEditCancel.Enabled
    
    tbrThis.Buttons("合并").Enabled = mnuDesignMerge.Enabled
    tbrThis.Buttons("撤消").Enabled = mnuDesignMergeCancel.Enabled
    tbrThis.Buttons("字体").Enabled = mnuDesignFont.Enabled
    tbrThis.Buttons("颜色").Enabled = mnuDesignColor.Enabled
    tbrThis.Buttons("水平").Enabled = mnuDesignHsb.Enabled
    tbrThis.Buttons("垂直").Enabled = mnuDesignVsb.Enabled
        
End Sub

Private Sub Reset()
    '复位,所有用到的一些数据进行初始化
    
    msf.Visible = False
    txtInput.Visible = False
    
    msf.Rows = 2
    msf.Cols = 2
    msf.Row = 1
    msf.Col = 1
    msf.TextMatrix(1, 1) = ""
    msf.MergeCol(1) = False
    msf.MergeRow(1) = False
    msf.RowData(1) = 0
    msf.ColData(1) = 0
    msf.CellAlignment = 9
    msf.CellForeColor = 0
    msf.CellBackColor = &H80000005
        
    mnuAutoMergeCol.Checked = False
    mnuAutoMergeRow.Checked = False
    
End Sub

Private Sub LoadDefTable()
    '装载用户自定义表格项目
    Dim Itmx As ListItem
    
    On Error GoTo errHand
    lvw.ListItems.Clear
    gstrSQL = "select 序号,名称,列数,列宽,行数,行高,合并行,合并列,颜色 from 咨询表格元素"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!序号, IIf(IsNull(gRs!名称), "", gRs!名称), 1, 1)
            Itmx.SubItems(1) = IIf(IsNull(gRs!行数), "", gRs!行数)
            Itmx.SubItems(2) = IIf(IsNull(gRs!列数), "", gRs!列数)
            gRs.MoveNext
        Wend
    End If
    CloseRecord gRs
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadStatus()
    Select Case picLvwBack.Tag
    Case ""
        stbThis.Panels(2).Text = "当前表格:《" & txt(0).Text & "》  处于查看状态，不能编辑。"
    Case "1"
        stbThis.Panels(2).Text = "当前表格:《" & txt(0).Text & "》  处于新增编辑状态。"
    Case "2"
        stbThis.Panels(2).Text = "当前表格:《" & txt(0).Text & "》  处于修改编辑状态。"
    End Select
End Sub

Private Sub ReadRegister()
    '读取注册表信息
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    picX.Left = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "picX位置", 2385)
End Sub

Private Sub WriteRegister()
    '将信息写回注册表
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "picX位置", picX.Left
End Sub

Private Sub PrintObject(ByVal intMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     intMode: 2表示预览 1打印 3输出到EXCEL
    '返回：
    '---------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow

    If lvw.SelectedItem Is Nothing Then Exit Sub

    If UserInfo.姓名 = "" Then Call GetUserInfo
    
    objPrint.Title = "用户表格-" & lvw.SelectedItem.Text
    
'    objPrint.BelowAppItems.Add "打印人:" & UserInfo.姓名
'    objPrint.BelowAppItems.Add "打印时间:" & Format(zlDatabase.Currentdate, "YYYY年MM月DD日")
    objPrint.Footer = "第[页码]页;;"
   
    msf.Redraw = False
    Call SaveRowCol
    msfPrint.Rows = msf.Rows - 1
    msfPrint.Cols = msf.Cols - 1
    For i = 1 To msf.Rows - 1
        msf.Row = i
        msfPrint.Row = i - 1
        msfPrint.RowHeight(i - 1) = msf.RowHeight(i)
        msfPrint.MergeRow(i - 1) = msf.MergeRow(i)
        For j = 1 To msf.Cols - 1
            msf.Col = j
            msfPrint.Col = j - 1
            msfPrint.MergeCol(j - 1) = msf.MergeCol(j)
            msfPrint.ColWidth(j - 1) = msf.ColWidth(j)
            msfPrint.CellFontName = msf.CellFontName
            msfPrint.CellFontSize = msf.CellFontSize
            msfPrint.CellFontBold = msf.CellFontBold
            msfPrint.CellAlignment = msf.CellAlignment
            msfPrint.CellFontStrikeThrough = msf.CellFontStrikeThrough
            msfPrint.CellFontUnderline = msf.CellFontUnderline
            msfPrint.CellForeColor = msf.CellForeColor
            
            msfPrint.TextMatrix(i - 1, j - 1) = msf.TextMatrix(i, j)
        Next
    Next
    Call RestoreRowCol
    msf.Redraw = True
    
    Set objPrint.Body = msfPrint
    
    If intMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, intMode
    End If

End Sub

Private Sub SaveRowCol()
    mSvrRow = msf.Row
    mSvrCol = msf.Col
End Sub

Private Sub RestoreRowCol()
    On Error Resume Next
    msf.Row = mSvrRow
    msf.Col = mSvrCol
End Sub

Private Sub CheckColWidth()
    Dim i As Long
    
    For i = 1 To msf.Cols - 1
        msf.ColWidth(i) = IIf(msf.ColWidth(i) = -1, 1200, msf.ColWidth(i))
    Next
End Sub

Private Sub SelectHsbAlign(ByVal Index As Long)
    Dim i As Long
    
    For i = 0 To mnuHsbAlign.UBound
        mnuHsbAlign(i).Checked = False
        mnuShort2Hsb(i).Checked = False
    Next
    mnuHsbAlign(Index).Checked = True
    mnuShort2Hsb(Index).Checked = True
End Sub

Private Sub SelectVsbAlign(ByVal Index As Long)
    Dim i As Long
    
    For i = 0 To mnuVsbAlign.UBound
        mnuVsbAlign(i).Checked = False
        mnuShort3Vsb(i).Checked = False
    Next
    mnuVsbAlign(Index).Checked = True
    mnuShort3Vsb(Index).Checked = True
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

