VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVBill 
   Caption         =   "所见单设置"
   ClientHeight    =   6855
   ClientLeft      =   -135
   ClientTop       =   240
   ClientWidth     =   10335
   Icon            =   "frmVBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   4320
      TabIndex        =   18
      Top             =   1320
      Width           =   975
      Begin VB.CommandButton cmdTab 
         Caption         =   "标记图"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.TreeView tvwItem 
         Height          =   1995
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   3519
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "iLsTree"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   855
      Left            =   3945
      TabIndex        =   14
      Top             =   4200
      Width           =   120
      _ExtentX        =   212
      _ExtentY        =   1508
      ButtonWidth     =   318
      ButtonHeight    =   1508
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "弹出所见项分类"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3390
      Top             =   4575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   5
      ImageHeight     =   51
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAttr 
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   7560
      TabIndex        =   10
      Top             =   720
      Width           =   2000
      Begin VB.CommandButton cmdHideAttr 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   1680
         TabIndex        =   13
         Top             =   120
         Width           =   220
      End
      Begin VB.ComboBox cmbControl 
         Height          =   300
         ItemData        =   "frmVBill.frx":090C
         Left            =   0
         List            =   "frmVBill.frx":0913
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   280
         Width           =   1455
      End
      Begin zl9CISBase.EGrid grdAttr 
         Height          =   4215
         Left            =   0
         TabIndex        =   27
         Top             =   600
         Width           =   1455
         _extentx        =   2566
         _extenty        =   7435
         font            =   "frmVBill.frx":091A
         fontfixed       =   "frmVBill.frx":093E
         backcolorfixed  =   -2147483643
         forecolor       =   -2147483640
         editforecolor   =   -2147483640
         rows            =   1
         fixedrows       =   0
         rowsizingmode   =   0
         rowheightmin    =   0
         highlight       =   1
      End
      Begin VB.Label lblAttr 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "属性"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ilstTool 
      Left            =   600
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":0964
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":163E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":2318
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":2632
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000005&
      Height          =   2820
      Left            =   1110
      ScaleHeight     =   2760
      ScaleWidth      =   5685
      TabIndex        =   5
      Top             =   840
      Width           =   5745
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   7815
         TabIndex        =   6
         Top             =   0
         Width           =   7815
         Begin VB.PictureBox fraTable 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Index           =   0
            Left            =   960
            MousePointer    =   15  'Size All
            ScaleHeight     =   1335
            ScaleWidth      =   1575
            TabIndex        =   24
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
            Begin TTF160Ctl.F1Book F1Book1 
               Height          =   735
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1296
               _0              =   $"frmVBill.frx":330C
               _1              =   $"frmVBill.frx":3715
               _2              =   $"frmVBill.frx":3B1E
               _3              =   $"frmVBill.frx":3F27
               _4              =   $"frmVBill.frx":4330
               _5              =   $"frmVBill.frx":473A
               _count          =   6
               _ver            =   2
            End
            Begin zl9CISBase.VisItem VisItem 
               Height          =   225
               Index           =   0
               Left            =   480
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1200
               _extentx        =   2328
               _extenty        =   397
               mousepointer    =   15
               font            =   "frmVBill.frx":4911
               enabled         =   0   'False
            End
         End
         Begin zl9CISBase.VisItem VisItem1 
            Height          =   345
            Index           =   0
            Left            =   840
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   1215
            _extentx        =   2328
            _extenty        =   397
            mousepointer    =   15
            font            =   "frmVBill.frx":4935
         End
         Begin VB.PictureBox Line1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   1335
            Index           =   0
            Left            =   360
            ScaleHeight     =   1305
            ScaleWidth      =   0
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   8
         End
         Begin VB.PictureBox shpDot 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   75
            Index           =   0
            Left            =   3360
            MousePointer    =   6  'Size NE SW
            ScaleHeight     =   45
            ScaleWidth      =   45
            TabIndex        =   8
            Top             =   3120
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            MousePointer    =   5  'Size
            TabIndex        =   7
            Text            =   "标签"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Shape shpSelect 
            BorderStyle     =   3  'Dot
            Height          =   855
            Left            =   3975
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1245
      Left            =   6360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Width           =   285
   End
   Begin VB.HScrollBar HScroll 
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   240
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   6840
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10335
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   645
      Width1          =   8370
      FixedBackground1=   0   'False
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "Save"
               Object.ToolTipText     =   "保存"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "设计"
               Key             =   "Design"
               Object.ToolTipText     =   "设计记帐单"
               Object.Tag             =   "设计"
               ImageKey        =   "Design"
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "记帐单查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
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
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "项目"
               Key             =   "Element"
               Object.ToolTipText     =   "设置所见项目"
               Object.Tag             =   "项目"
               ImageKey        =   "Element"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "对齐"
               Key             =   "Align"
               Object.ToolTipText     =   "设置对象之间的对齐方式"
               Object.Tag             =   "对齐"
               ImageKey        =   "Align"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   7
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Left"
                     Text            =   "左对齐"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HAlign"
                     Text            =   "居中对齐"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Right"
                     Text            =   "右对齐"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Top"
                     Text            =   "上对齐"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VAlign"
                     Text            =   "中间对齐"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Bottom"
                     Text            =   "下对齐"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "居中"
               Key             =   "Form"
               Object.ToolTipText     =   "水平居中"
               Object.Tag             =   "居中"
               ImageKey        =   "Form"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HCenter"
                     Text            =   "水平居中"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VCenter"
                     Text            =   "垂直居中"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "横距"
               Key             =   "HDistance"
               Object.ToolTipText     =   "设置对象之间的水平间距"
               Object.Tag             =   "横距"
               ImageIndex      =   14
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HSpace"
                     Text            =   "横间距相同"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HNo"
                     Text            =   "无横间距"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VSpace"
                     Text            =   "竖间距相同"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "VNo"
                     Text            =   "无竖间距"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "纵距"
               Key             =   "VDistance"
               Object.ToolTipText     =   "设置对象之间的垂直间距"
               Object.Tag             =   "纵距"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "尺寸"
               Key             =   "Size"
               Object.ToolTipText     =   "将选择的对象设为统一的大小"
               Object.Tag             =   "尺寸"
               ImageKey        =   "Size"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "相同宽度"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "相同高度"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "两者都相同"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "锁定"
               Key             =   "Lock"
               Object.ToolTipText     =   "锁定"
               Object.Tag             =   "锁定"
               ImageKey        =   "Lock"
               Style           =   1
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   1680
      Top             =   5145
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":4959
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":4B75
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":4D95
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":4FB5
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":51D5
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":53F5
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":5615
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":5D0F
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":6029
            Key             =   "Align"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":6723
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":6E1D
            Key             =   "Distance"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":7517
            Key             =   "Size"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":7C11
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":830B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3960
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":8A05
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":8C21
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":8E41
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":9061
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":9281
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":94A1
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":96C1
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":9DBB
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":A0D5
            Key             =   "Align"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":A7CF
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":AEC9
            Key             =   "Distance"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":B5C3
            Key             =   "Size"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":BCBD
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":C3B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLsTree32 
      Left            =   2250
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":CB31
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":D40B
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":D725
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":D87F
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":D9D9
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":DB33
            Key             =   "Option"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":DC8D
            Key             =   "Combox_NotUse"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":DDE7
            Key             =   "Combox"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   2295
      Top             =   5715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":DF41
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":E09B
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":E3B5
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":EC8F
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":EDE9
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":EF43
            Key             =   "Option"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":F09D
            Key             =   "Combox_NotUse"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBill.frx":F1F7
            Key             =   "Combox"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraToolbox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   735
      Begin VB.PictureBox picTool 
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   22
         Top             =   240
         Width           =   735
         Begin MSComctlLib.Toolbar ControlBar 
            Height          =   2280
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   4022
            ButtonWidth     =   1032
            ButtonHeight    =   1005
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ilstTool"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "指针"
                  ImageIndex      =   1
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "标签"
                  ImageIndex      =   4
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "线条"
                  ImageIndex      =   2
                  Style           =   2
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "附加表"
                  ImageIndex      =   3
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lvwSubItem 
         Height          =   2295
         Left            =   0
         TabIndex        =   16
         Top             =   2985
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iLsTree32"
         SmallIcons      =   "iLsTree"
         ColHdrIcons     =   "iLsTree"
         ForeColor       =   -2147483641
         BackColor       =   -2147483643
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "编码"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "英文名"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "替换域"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "类型"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "长度"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "小数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "单位"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "表示法"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "性别域"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "数值域"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "正常域"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "初始值"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "文字表述"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "空值文字"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "临床意义"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "通用"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   900
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "所见项"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1065
         Width           =   900
      End
   End
   Begin MSComctlLib.ProgressBar prbRefresh 
      Height          =   195
      Left            =   1920
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6495
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   635
      SimpleText      =   $"frmVBill.frx":F351
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmVBill.frx":F398
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
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
   Begin VB.Image imgY 
      Height          =   5115
      Index           =   2
      Left            =   0
      MousePointer    =   9  'Size W E
      Top             =   1080
      Width           =   45
   End
   Begin VB.Image imgY 
      Height          =   5115
      Index           =   1
      Left            =   7320
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   45
   End
   Begin VB.Image imgY 
      Height          =   5115
      Index           =   0
      Left            =   2520
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileDesign 
         Caption         =   "设计(&D)"
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileReload 
         Caption         =   "重新装入(&R)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFile0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImp 
         Caption         =   "导入(&I)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExp 
         Caption         =   "导出(&E)"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "删除(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "全部选择(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "格式(&R)"
      Begin VB.Menu mnuFormatAlign 
         Caption         =   "对齐(&A)"
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "左对齐(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "居中对齐(&C)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "右对齐(&R)"
            Index           =   2
         End
         Begin VB.Menu mnuAlign_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFormatVAlign 
            Caption         =   "顶端对齐(&T)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatVAlign 
            Caption         =   "中间对齐(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatVAlign 
            Caption         =   "底端对齐(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFormat0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatToGrid 
         Caption         =   "对齐到网格(&G)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFormatSizeToGrid 
         Caption         =   "按网格调整大小(&D)"
      End
      Begin VB.Menu mnuFormatForm 
         Caption         =   "在窗口内居中对齐(&W)"
         Visible         =   0   'False
         Begin VB.Menu mnuFormatFormAlign 
            Caption         =   "水平居中(&H)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatFormAlign 
            Caption         =   "垂直居中(&V)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFormat1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatS 
         Caption         =   "统一尺寸(&M)"
         Begin VB.Menu mnuFormatSize 
            Caption         =   "相同宽度(&W)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatSize 
            Caption         =   "相同高度(&H)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatSize 
            Caption         =   "两者都相同(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFomrat2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatVsc 
         Caption         =   "水平间距(&H)"
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "相同(&E)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "增加(&I)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "减少(&D)"
            Index           =   2
         End
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "移除(&R)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFormatHsc 
         Caption         =   "垂直间距(&V)"
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "相同(&E)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "增加(&I)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "减少(&D)"
            Index           =   2
         End
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "移除(&R)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFormat3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatLock 
         Caption         =   "锁定元素(&L)"
         Shortcut        =   ^K
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
      Begin VB.Menu mnuEdit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ViewList 
         Caption         =   "显示所见分类列表(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEdit_ViewAttr 
         Caption         =   "显示属性页(&S)"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "记帐单列表(&L)"
         Checked         =   -1  'True
         Shortcut        =   {F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewAttrib 
         Caption         =   "属性表框(&A)"
         Checked         =   -1  'True
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
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
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmVBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'支持的控件选择方式：鼠标单击√、双击进入所见项设置、Ctrl/Shift多选√、鼠标框及全选√
'复制、粘贴以及控件删除
Option Explicit

Private Const GRIDDISTANCE As Long = 120
Private Const COLOR_GRAY As Long = &HE0E0E0
Private Const COLOR_WHITE As Long = &HFFFFFF
Private Const COLOR_BLUE As Long = &H800000
Private Const COLOR_YELLOW As Long = &HFFFF&
Private Type CtrlPoint
    CtrlName As String
    CtrlIndex As Long
    Visible As Boolean
End Type
Private SelectedCtrls() As CtrlPoint
Private iRangeX As Single, iRangeY As Single, iRangeWidth As Single, iRangeHeight As Single
'拖动的起始点
Private iOrigX As Long, iOrigY As Long
'当前对象
Private CurrObject As Long
'当前选择的控件ID（ToolBar的按钮ID）
Private CurrObjType  As Integer
Private bNotRunCombox_Click As Boolean

Private clsComLib As New zl9ComLib.clsComLib
Private clsDatabase As New zl9ComLib.clsDatabase
Private clsCommfun As New zl9ComLib.clsCommfun
Private clsControl As New zl9ComLib.clsControl
Private iCurrTab As Integer
Private Modified As Boolean
Public FormID As String

Private Sub cmbControl_Click()
    Dim CtrlName As String, CtrlIndex As Long
    Dim iPos1 As Integer, iPos2 As Integer
    
    If bNotRunCombox_Click Then
        Exit Sub
    End If
    
'    On Error Resume Next
    If Len(Trim(cmbControl.Text)) = 0 Then Exit Sub
    iPos1 = InStr(cmbControl.Text, "(")
    iPos2 = InStr(cmbControl.Text, ")")
    
    CtrlName = Mid(cmbControl.Text, 1, iPos1 - 1)
    CtrlIndex = CLng(Mid(cmbControl.Text, iPos1 + 1, iPos2 - iPos1 - 1))
    Select Case CtrlName
        Case "标签"
            CtrlName = "Text1"
        Case "线"
            CtrlName = "Line1"
        Case "附加表"
            CtrlName = "fraTable"
        Case "所见项"
            CtrlName = "VisItem1"
    End Select
    
    SelectControl Me.Controls(CtrlName)(CtrlIndex), False, False
    ShowAttribute False
End Sub

Private Sub cmdHideAttr_Click()
    mnuEdit_ViewAttr_Click
End Sub

Private Sub cmdTab_Click(Index As Integer)
    If iCurrTab = Index Then Exit Sub
    iCurrTab = Index
    If mnuEdit_ViewList.Checked Then
        Form_Resize
    Else
        ShowList 3800, lblItem.Top + lblItem.Height + 50 + fraToolbox.Top
    End If
    If tvwItem(iCurrTab).Nodes.Count > 0 Then
        Set tvwItem(iCurrTab).SelectedItem = tvwItem(iCurrTab).Nodes(1)
        tvwItem(iCurrTab).SetFocus
        tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
    Else
        lvwSubItem.ListItems.Clear
    End If
End Sub

Private Sub ControlBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    CurrObjType = Button.Index
    
    If CurrObjType <> 2 Then
        PicForm.MousePointer = vbCrosshair
    Else
        PicForm.MousePointer = vbDefault
    End If
End Sub

Private Sub ControlBar_DblClick()
'    Dim NewControl As Control
'    If CurrObjType <> 2 Then
'        Select Case CurrObjType
'            Case 3
''                Load Text1(Text1.Count)
''                With Text1(Text1.Count - 1)
'                Set NewControl = LoadNewControl("Text1")
'                With NewControl
'                    .Visible = True
'                End With
'                SelectControl NewControl, False
'
'                AddControlList "Text1", CStr(NewControl.Index), NewControl.Text
'            Case 4
'                Load Line1(Line1.Count)
'                With Line1(Line1.Count - 1)
'                    .Visible = True
'                End With
'                SelectControl Line1(Line1.Count - 1), False
'
'                AddControlList "Line1", CStr(Line1.Count - 1)
'        End Select
'        ControlBar.Buttons(2).Value = tbrPressed
'        CurrObjType = 2
'
'        ShowAttribute
'    End If
End Sub

Private Sub Form_Activate()
    Dim objCtrl As Control
    If Me.Tag = "" Then Exit Sub
    
    Form_Resize
    
    Me.Tag = ""
    Me.MousePointer = vbHourglass
    BeginShowProgress
    ReadForm Me, "PicForm", FormID, , , Me.prbRefresh
    For Each objCtrl In Me.Controls
        If UCase(objCtrl.Name) = "FRATABLE" Then
            Proc_Table_TopLeftChanged F1Book1(objCtrl.Index)
        End If
    Next
    CreateAllCtrlList
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim objCtrl As Control
    ReDim SelectedCtrls(0)
    
    Call RestoreWinState(Me, App.ProductName)
    
    With PicForm
        .Left = 0
        .Top = 0
        .Width = Screen.Width
        .Height = Screen.Height
    End With
    DrawGrid COLOR_GRAY
    
    CurrObjType = 2
    PicForm.MousePointer = vbDefault
    Modified = False
    
    grdAttr.AllowAddNew = False

    CreateItemTree
    With lvwSubItem
        .View = lvwReport
        
        .ColumnHeaders(1).Width = 2000
        For i = 2 To .ColumnHeaders.Count
            .ColumnHeaders(i).Width = 0
        Next
    End With
    
    On Error Resume Next
    iCurrTab = 1
    Set tvwItem(1).SelectedItem = tvwItem(1).Nodes(1)
    
    Me.Tag = "Loading" '要加载所见单
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    With Toolbar2 '弹出菜单
        .Left = -50: .Top = (Me.ScaleHeight - lngStatus - lngTools - .Height) / 2 + lngTools
        .Visible = Not mnuEdit_ViewList.Checked
    End With
    With imgY(0)
        .Top = lngTools
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
    With imgY(1)
        .Left = Me.ScaleWidth - fraAttr.Width - .Width: .Top = lngTools
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
    With imgY(2)
        .Top = lngTools
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
    With VScroll
        .Top = lngTools
        .Left = IIf(fraAttr.Visible, imgY(1).Left, Me.ScaleWidth) - .Width
        .Height = Me.ScaleHeight - lngStatus - HScroll.Height - .Top
        .Min = 0
        .Max = PicForm.Height - .Height
        .SmallChange = .Height / 5
        .LargeChange = .Height
        .Value = PicForm.Top
    End With
    With HScroll
        .Top = Me.ScaleHeight - lngStatus - .Height
        .Left = imgY(0).Left + imgY(0).Width
        .Width = IIf(fraAttr.Visible, imgY(1).Left, Me.ScaleWidth) - VScroll.Width - .Left
        .Min = 0
        .Max = PicForm.Width - .Width
        .SmallChange = .Width / 5
        .LargeChange = .Width
        .Value = PicForm.Left
    End With
    With picMain
        .Left = imgY(0).Left + imgY(0).Width
        .Top = lngTools
        .Width = VScroll.Left - .Left
        .Height = HScroll.Top - .Top
    End With
    
    With fraToolbox
        .Left = IIf(mnuEdit_ViewList.Checked, imgY(2).Left + imgY(2).Width, Toolbar2.Width + Toolbar2.Left)
        .Width = imgY(0).Left - .Left
        .Top = lngTools
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
    With lblTool
        .Left = 0: .Top = 20
        .Width = fraToolbox.Width - .Left
    End With
    With picTool
        .Left = 0: .Top = lblTool.Top + lblTool.Height + 20
        .Width = fraToolbox.Width - .Left
    End With
    With ControlBar
        .Left = 0: .Top = 20
        .Width = fraToolbox.Width - .Left
        picTool.Height = .Top + .Height + 20
    End With
    With lblItem
        .Left = 0: .Top = picTool.Top + picTool.Height + 50
        .Width = fraToolbox.Width - .Left
    End With
    With lvwSubItem
        .Left = 0: .Top = lblItem.Top + lblItem.Height + 20
        .Width = fraToolbox.Width - .Left: .Height = fraToolbox.Height - .Top
        .Refresh
    End With
    
    With fraAttr
        .Left = imgY(1).Left + imgY(1).Width
        .Width = Me.ScaleWidth - .Left
        .Top = lngTools
        .Height = Me.ScaleHeight - lngStatus - .Top
    End With
    With lblAttr
        .Left = 0: .Top = 20
        .Width = fraAttr.Width
    End With
    With cmbControl
        .Left = 0: .Top = lblAttr.Top + lblAttr.Height + 20
        .Width = fraAttr.Width
    End With
    With grdAttr
        .Left = 0: .Width = fraAttr.Width
        .Height = fraAttr.Height - .Top
        .ColWidth(0) = 1000
        .ColWidth(1) = .Width - .ColWidth(0)
        .ColAlignment(1) = flexAlignLeftCenter
    End With
    With cmdHideAttr
        .Left = lblAttr.Width - .Width - 30: .Top = lblAttr.Top
    End With

    '   显示选项卡
    ShowList imgY(2).Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub fraTable_DblClick(Index As Integer)
    With frmAddTable
        .theTableID = "": .TableTitle = "附加表"
        WriteToTable F1Book1(Index)
        Set .theTable = F1Book1(Index)
        .tbrThis.Buttons("打印").Visible = False
        .tbrThis.Buttons("预览").Visible = False
        .tbrThis.Buttons("保存").Visible = False
        .tbrThis.Buttons("Split_2").Visible = False
        
        .mnuFilePrintset.Visible = False
        .mnuFileExcel.Visible = False
        .mnuFilePreview.Visible = False
        .mnuFilePrint.Visible = False
        .mnuFileSplit1.Visible = False
        .mnusplit2.Visible = False
        .mnuEditSave.Visible = False
        
        .Show vbModal
        Set .theTable = Nothing
    End With
    Unload frmAddTable
    
    Modified = True
    
    fraTable_Resize Index
    Me.MousePointer = vbHourglass
    BeginShowProgress
    RefreshObject F1Book1(Index), , False, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
End Sub
''重新刷新表内所见项
'Private Sub RefreshObject(theTable As TTF160Ctl.F1Book)
'    Dim iDecPos As Integer
'    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
'    Dim iCurrRow As Integer, iCurrCol As Integer
'    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
'
'    On Error Resume Next
'    iCurrRow = theTable.Row: iCurrCol = theTable.Col
'    For Each tmpCtrl In Me.Controls
'        If tmpCtrl.Name = "VisItem" Then
'            If tmpCtrl.Container.Index = theTable.Container.Index Then tmpCtrl.Visible = False
'        End If
'    Next
'    For iRow = 1 To theTable.MaxRow
'        For iCol = 1 To theTable.MaxCol
'            theTable.SetActiveCell iRow, iCol
'
'            Set objCellFormat = theTable.GetCellFormat
'            If Len(objCellFormat.ValidationText) > 0 And iRow = theTable.SelStartRow And iCol = theTable.SelStartCol Then
'                aVisItemInfo = Split(objCellFormat.ValidationText, ",")
'
'                objCellFormat.ValidationText = ""
'                theTable.SetCellFormat objCellFormat
'
'                AddObject theTable, iRow, iCol, CLng(aVisItemInfo(0)), False, theTable.TextRC(iRow, iCol), Me
'                With VisItem(VisItem.UBound)
''                    Set .Container = theTable.Container
'                    .Visible = True: .Enabled = False
'                End With
'            End If
'        Next iCol
'    Next iRow
'    For Each tmpCtrl In Me.Controls
'        If tmpCtrl.Name = "VisItem" And Not tmpCtrl.Visible Then Unload tmpCtrl
'    Next
'    theTable.SetActiveCell iCurrRow, iCurrCol
'End Sub
'将所见项的值写入单元格中
Private Sub WriteToTable(theTable As TTF160Ctl.F1Book)
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
    
    On Error Resume Next
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    For iRow = 1 To theTable.MaxRow
        For iCol = 1 To theTable.MaxCol
            theTable.SetActiveCell iRow, iCol

            Set objCellFormat = theTable.GetCellFormat
            If Len(objCellFormat.ValidationText) > 0 And iRow = theTable.SelStartRow And iCol = theTable.SelStartCol Then
                aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                theTable.TextRC(iRow, iCol) = Me.VisItem(aVisItemInfo(1)).Value
            End If
        Next iCol
    Next iRow
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub

Private Sub fraTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    InitMoveControl fraTable(Index), Button, Shift, x, y
    Select Case Button
        Case vbLeftButton
        Case vbRightButton
            Me.PopupMenu Me.mnuFormat
    End Select
End Sub

Private Sub FraTable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ProcMoveControl fraTable(Index), Button, Shift, x, y
End Sub

Private Sub fraTable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    EndMoveControl fraTable(Index), Button, Shift, x, y
End Sub

Private Sub fraTable_Resize(Index As Integer)
    With F1Book1(Index)
        .Left = 0: .Top = 0
        .Width = fraTable(Index).Width: .Height = fraTable(Index).Height
        .Visible = True
    End With
End Sub

Private Sub grdAttr_BeforeColUpdate(ByVal RowIndex As Long, ByVal ColIndex As Long, NewValue As String, ByVal OldValue As String, Cancel As Boolean)
    Dim vSetValue As Variant
    Dim tmpControl As Control
    Dim MinValue As Variant, MaxValue As Variant
    On Error Resume Next
    If InStr("左边距,上边距,宽度,高度,输入顺序", grdAttr.Text(RowIndex, 0)) > 0 Then
        If Len(Trim(NewValue)) = 0 Then
            MsgBox "值不能为空！请输入内容或敲Esc键放弃。", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Not IsNumeric(NewValue) Then
            MsgBox "必须输入数字！", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
        NewValue = CLng(NewValue)
        If CLng(NewValue) > 10000 Then
            MsgBox "无效的属性值！", vbExclamation + vbOKOnly, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End If
    
    vSetValue = NewValue
    Select Case grdAttr.Text(RowIndex, 0)
        Case "左边距"
        Case "上边距"
        Case "宽度"
        Case "高度"
        Case "标题"
            If LenB(StrConv(NewValue, vbFromUnicode)) > 50 Then
                MsgBox "标题不能超过50个字符（汉字占两个字符）！", vbExclamation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
        Case "计量单位"
            If LenB(StrConv(NewValue, vbFromUnicode)) > 10 Then
                MsgBox "计量单位不能超过10个字符（汉字占两个字符）！", vbExclamation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
        Case "对齐"
            vSetValue = grdAttr.ListIndex(RowIndex, ColIndex)
        Case "不可输入"
        Case "缺省值"
            If Len(Trim(NewValue)) = 0 Then
                NewValue = ""
                vSetValue = ""
            Else
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            With tmpControl
                If .Method < 2 And .ItemType <> 3 Then
                    Select Case .ItemType
                        Case 0
                            If Not IsNumeric(NewValue) Then
                                MsgBox "必须输入数字！", vbExclamation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            If .MaxLength > 0 And Abs(CLng(NewValue)) > 0 And Len(CStr(Abs(CLng(NewValue)))) > CLng(.MaxLength) - CLng(.DecLength) Then
                                MsgBox "值超过定义的最大长度！", vbExclamation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            If .ValuesCount > 0 Then
                                If Len(Trim(.Values(0))) > 0 Then
                                    If CDbl(NewValue) < CDbl(.Values(0)) Then
                                        MsgBox "输入值不能小于" & .Values(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                End If
                                If .ValuesCount > 1 Then
                                    If Len(Trim(.Values(1))) > 0 Then
                                    If CDbl(NewValue) > CDbl(.Values(1)) Then
                                        MsgBox "输入值不能大于" & .Values(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    End If
                                End If
                            End If
                            If CInt(.DecLength) > 0 Then
                                NewValue = Format(Round(NewValue, CLng(.DecLength)), "#." + String(CLng(.DecLength), "0"))
                            Else
                                NewValue = Round(NewValue, 0)
                            End If
                        Case 1
                            If .MaxLength > 0 And Len(Trim(NewValue)) > CLng(.MaxLength) Then
                                MsgBox "值超过定义的最大长度！", vbExclamation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            If .ValuesCount > 0 Then
                                If Len(Trim(.Values(0))) > 0 Then
                                    If NewValue < .Values(0) Then
                                        MsgBox "输入值不能小于" & .Values(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                End If
                                If .ValuesCount > 1 Then
                                    If Len(Trim(.Values(1))) > 0 Then
                                    If NewValue > .Values(1) Then
                                        MsgBox "输入值不能大于" & .Values(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    End If
                                End If
                            End If
                        Case 2
                            If Not IsDate(NewValue) Then
                                MsgBox "错误的日期格式！日期格式为XXXX-XX-XX", vbExclamation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            Else
                                NewValue = Format(NewValue, "YYYY-MM-DD")
                            End If
                            If .ValuesCount > 0 Then
                                If Len(Trim(.Values(0))) > 0 Then
                                    If CDate(NewValue) < CDate(.Values(0)) Then
                                        MsgBox "输入值不能小于" & .Values(0) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                End If
                                If .ValuesCount > 1 Then
                                    If Len(Trim(.Values(1))) > 0 Then
                                    If CDate(NewValue) > CDate(.Values(1)) Then
                                        MsgBox "输入值不能大于" & .Values(1) & "！", vbExclamation + vbOKOnly, gstrSysName
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    End If
                                End If
                            End If
                    End Select
                End If
            End With
            End If
            vSetValue = NewValue
    End Select
    
    SetControlAttr grdAttr.Text(RowIndex, 0), vSetValue
    If InStr("左边距,上边距,宽度,高度,输入顺序", grdAttr.Text(RowIndex, 0)) > 0 Then NewValue = vSetValue
End Sub

Private Sub grdAttr_DblClick()
    On Error Resume Next
    If Len(grdAttr.Text(grdAttr.Row, 0)) > 0 Then grdAttr.Edit grdAttr.Row, grdAttr.Col
End Sub

Private Sub grdAttr_LostFocus()
    grdAttr.ValidValue
End Sub

Private Sub HScroll_Change()
    PicForm.Left = -1 * HScroll.Value
End Sub

Private Sub imgY_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim iOldLeft As Long
    If Not Button = vbLeftButton Then Exit Sub
    Select Case Index
        Case 0
            imgY(0).Left = imgY(0).Left + x
            If imgY(0).Left - imgY(2).Left < 600 Then imgY(0).Left = 600 + imgY(2).Left
            If imgY(0).Left - imgY(2).Left > 3000 Then imgY(0).Left = 3000 + imgY(2).Left
            
            Form_Resize
        Case 1
            With fraAttr
                .Width = .Width - x
                If .Width < 2000 Then .Width = 2000
                If .Width > 5000 Then .Width = 5000
                
                Form_Resize
            End With
        Case 2
            iOldLeft = imgY(2).Left
            imgY(2).Left = imgY(2).Left + x
            If imgY(2).Left < 1400 Then imgY(2).Left = 1400
            If imgY(2).Left > 4000 Then imgY(2).Left = 4000
            imgY(0).Left = imgY(0).Left + imgY(2).Left - iOldLeft
            
            Form_Resize
    End Select
End Sub

Private Sub lblItem_Click()
    If Not mnuEdit_ViewList.Checked Then
        If fraList.Visible Then
            fraList.Visible = False
        Else
            ShowList 3800, lblItem.Top + lblItem.Height + 50 + fraToolbox.Top
            If lvwSubItem.ListItems.Count = 0 And Not tvwItem(iCurrTab).SelectedItem Is Nothing Then tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
        End If
    End If
End Sub

Private Sub Line1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case vbLeftButton
            InitMoveControl Line1(Index), Button, Shift, x, y
        Case vbRightButton
    End Select
End Sub

Private Sub Line1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ProcMoveControl Line1(Index), Button, Shift, x, y
End Sub

Private Sub Line1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    EndMoveControl Line1(Index), Button, Shift, x, y
End Sub

Private Sub lvwSubItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CurrObjType = 9
    
    PicForm.MousePointer = vbCrosshair
End Sub

Private Sub mnuEdit_ViewAttr_Click()
    With mnuEdit_ViewAttr
        .Checked = Not .Checked
        fraAttr.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuEdit_ViewList_Click()
    Dim iToolWidth As Long
    With mnuEdit_ViewList
        .Checked = Not .Checked
        fraList.Visible = .Checked
    End With
    
    iToolWidth = imgY(0).Left - imgY(2).Left
    imgY(2).Left = IIf(mnuEdit_ViewList.Checked, 1800, 0)
    imgY(0).Left = imgY(2).Left + iToolWidth
    Form_Resize
    
    If lvwSubItem.ListItems.Count = 0 And Not tvwItem(iCurrTab).SelectedItem Is Nothing Then tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
End Sub

Private Sub mnuEditRemove_Click()
    DeleteCtrls
    
    Modified = True
End Sub

Private Sub mnuEditSelAll_Click()
    SelectAll
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileReload_Click()
    Dim objCtrl As Control
    If Modified Then
        If MsgBox("所见单已修改，重新加载后所作的修改将失效！是否继续", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    DeleteAllCtrls
    
    Me.MousePointer = vbHourglass
    BeginShowProgress
    ReadForm Me, "PicForm", FormID, , , Me.prbRefresh
    For Each objCtrl In Me.Controls
        If UCase(objCtrl.Name) = "FRATABLE" Then
            Proc_Table_TopLeftChanged F1Book1(objCtrl.Index)
        End If
    Next
    CreateAllCtrlList
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Modified = False
End Sub

Private Sub mnuFileSave_Click()
    Me.MousePointer = vbHourglass
    BeginShowProgress
    UnSelectAll
    SaveForm Me, "PicForm", FormID, Me.prbRefresh
    Me.prbRefresh.Visible = False
    Me.MousePointer = vbDefault
    
    Modified = False
End Sub

Private Sub mnuFormatDoAlign_Click(Index As Integer)
    ControlsAlign Index
End Sub

Private Sub mnuFormatHscSpace_Click(Index As Integer)
    Dim i As Long, j As Long, iNum As Long
    Dim tmpControl As Control
    Dim aSelectedPoints() As Long '用于存放SelectedCtrls的指针
    Dim PointsIndex As Long, PointsBuff As Long
    Dim tmpWidth As Long, ObjectDistance As Long, CurrLeft As Long
    Dim BuffIndex As Long
    iNum = UBound(SelectedCtrls)
    
    If SelectedCounts < IIf(Index = 0, 3, 2) Then Exit Sub
    ReDim aSelectedPoints(0)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ReDim Preserve aSelectedPoints(UBound(aSelectedPoints) + 1)
            aSelectedPoints(UBound(aSelectedPoints)) = i
        End If
    Next
    '注意线的处理
    On Error Resume Next
    '按选择对象的位置排序
    For i = 1 To UBound(aSelectedPoints) - 1
        PointsIndex = i
        For j = i + 1 To UBound(aSelectedPoints)
            If Me.Controls(SelectedCtrls(aSelectedPoints(j)).CtrlName)(SelectedCtrls(aSelectedPoints(j)).CtrlIndex).Top < Me.Controls(SelectedCtrls(aSelectedPoints(PointsIndex)).CtrlName)(SelectedCtrls(aSelectedPoints(PointsIndex)).CtrlIndex).Top Then PointsIndex = j
        Next j
        BuffIndex = aSelectedPoints(i)
        aSelectedPoints(i) = aSelectedPoints(PointsIndex)
        aSelectedPoints(PointsIndex) = BuffIndex
    Next i
    Select Case Index
        Case 0
            tmpWidth = 0
            For i = 2 To UBound(aSelectedPoints) - 1
                tmpWidth = tmpWidth + Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex).Height
            Next
            Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(1)).CtrlName)(SelectedCtrls(aSelectedPoints(1)).CtrlIndex)
            ObjectDistance = (Me.Controls(SelectedCtrls(aSelectedPoints(UBound(aSelectedPoints))).CtrlName)(SelectedCtrls(aSelectedPoints(UBound(aSelectedPoints))).CtrlIndex).Top - tmpControl.Top - tmpControl.Height - tmpWidth) / (UBound(aSelectedPoints) - 1)
            
            CurrLeft = tmpControl.Top + tmpControl.Height + ObjectDistance
            For i = 2 To UBound(aSelectedPoints) - 1
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Top = CurrLeft
                
                CurrLeft = CurrLeft + tmpControl.Height + ObjectDistance
            Next
        Case 1, 2
            For i = 2 To UBound(aSelectedPoints)
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Top = IIf(Index = 1, tmpControl.Top + GRIDDISTANCE, tmpControl.Top - GRIDDISTANCE)
            Next
        Case 3
            Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(1)).CtrlName)(SelectedCtrls(aSelectedPoints(1)).CtrlIndex)
            
            CurrLeft = tmpControl.Top + tmpControl.Height
            For i = 2 To UBound(aSelectedPoints)
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Top = CurrLeft
                
                CurrLeft = CurrLeft + tmpControl.Height
            Next
    End Select
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    
    ShowAttribute
End Sub

Private Sub mnuFormatLock_Click()
    mnuFormatLock.Checked = Not mnuFormatLock.Checked
    Toolbar1.Buttons("Lock").Value = IIf(mnuFormatLock.Checked, tbrPressed, tbrUnpressed)

    If CurrObject > 0 Then ShowSelect Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
    
    Me.mnuFormatAlign.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatForm.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatHsc.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatS.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatSizeToGrid.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatToGrid.Enabled = Not mnuFormatLock.Checked
    Me.mnuFormatVsc.Enabled = Not mnuFormatLock.Checked
    Me.Toolbar1.Buttons("Align").Enabled = Not mnuFormatLock.Checked
    Me.Toolbar1.Buttons("HDistance").Enabled = Not mnuFormatLock.Checked
    Me.Toolbar1.Buttons("VDistance").Enabled = Not mnuFormatLock.Checked
    Me.Toolbar1.Buttons("Size").Enabled = Not mnuFormatLock.Checked
End Sub

Private Sub mnuFormatSize_Click(Index As Integer)
    Dim i As Long, iNum As Long, curr
    Dim tmpControl As Control
    Dim CurrWidth As Long, CurrHeight As Long
    iNum = UBound(SelectedCtrls)
    If SelectedCounts < 2 Then Exit Sub
    
    '注意线的处理
    Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
    CurrWidth = tmpControl.Width: CurrHeight = tmpControl.Height
    On Error Resume Next
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
            With tmpControl
                If UCase(tmpControl.Name) = "LINE1" Then
                    If (Index = 0 Or Index = 2) And .Width > 15 And CurrWidth > 15 Then .Width = CurrWidth
                    If (Index = 1 Or Index = 2) And .Height > 15 And CurrHeight > 15 Then .Height = CurrHeight
                Else
                    If Index = 0 Or Index = 2 Then .Width = CurrWidth
                    If Index = 1 Or Index = 2 Then .Height = CurrHeight
                End If
            End With
        End If
    Next
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    ShowAttribute
End Sub

Private Sub mnuFormatSizeToGrid_Click()
    Dim i As Long, iNum As Long
    Dim tmpControl As Control
    iNum = UBound(SelectedCtrls)
    
    '注意线的处理
    On Error Resume Next
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
            With tmpControl
                .Left = CLng(.Left / GRIDDISTANCE) * GRIDDISTANCE
                .Width = CLng(.Width / GRIDDISTANCE) * GRIDDISTANCE
                .Top = CLng(.Top / GRIDDISTANCE) * GRIDDISTANCE
                .Height = CLng(.Height / GRIDDISTANCE) * GRIDDISTANCE
            End With
        End If
    Next
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    ShowAttribute
End Sub

Private Sub mnuFormatToGrid_Click()
    mnuFormatToGrid.Checked = Not mnuFormatToGrid.Checked
End Sub

Private Sub mnuFormatVAlign_Click(Index As Integer)
    ControlsAlign 3 + Index
End Sub

Private Sub mnuFormatVscSpace_Click(Index As Integer)
    Dim i As Long, j As Long, iNum As Long
    Dim tmpControl As Control
    Dim aSelectedPoints() As Long '用于存放SelectedCtrls的指针
    Dim PointsIndex As Long, PointsBuff As Long
    Dim tmpWidth As Long, ObjectDistance As Long, CurrLeft As Long
    Dim BuffIndex As Long
    iNum = UBound(SelectedCtrls)
    
    If SelectedCounts < IIf(Index = 0, 3, 2) Then Exit Sub
    ReDim aSelectedPoints(0)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ReDim Preserve aSelectedPoints(UBound(aSelectedPoints) + 1)
            aSelectedPoints(UBound(aSelectedPoints)) = i
        End If
    Next
    '注意线的处理
    On Error Resume Next
    '按选择对象的位置排序
    For i = 1 To UBound(aSelectedPoints) - 1
        PointsIndex = i
        For j = i + 1 To UBound(aSelectedPoints)
            If Me.Controls(SelectedCtrls(aSelectedPoints(j)).CtrlName)(SelectedCtrls(aSelectedPoints(j)).CtrlIndex).Left < Me.Controls(SelectedCtrls(aSelectedPoints(PointsIndex)).CtrlName)(SelectedCtrls(aSelectedPoints(PointsIndex)).CtrlIndex).Left Then PointsIndex = j
        Next j
        BuffIndex = aSelectedPoints(i)
        aSelectedPoints(i) = aSelectedPoints(PointsIndex)
        aSelectedPoints(PointsIndex) = BuffIndex
    Next i
    Select Case Index
        Case 0
            tmpWidth = 0
            For i = 2 To UBound(aSelectedPoints) - 1
                tmpWidth = tmpWidth + Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex).Width
            Next
            Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(1)).CtrlName)(SelectedCtrls(aSelectedPoints(1)).CtrlIndex)
            ObjectDistance = (Me.Controls(SelectedCtrls(aSelectedPoints(UBound(aSelectedPoints))).CtrlName)(SelectedCtrls(aSelectedPoints(UBound(aSelectedPoints))).CtrlIndex).Left - tmpControl.Left - tmpControl.Width - tmpWidth) / (UBound(aSelectedPoints) - 1)
            
            CurrLeft = tmpControl.Left + tmpControl.Width + ObjectDistance
            For i = 2 To UBound(aSelectedPoints) - 1
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Left = CurrLeft
                
                CurrLeft = CurrLeft + tmpControl.Width + ObjectDistance
            Next
        Case 1, 2
            For i = 2 To UBound(aSelectedPoints)
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Left = IIf(Index = 1, tmpControl.Left + GRIDDISTANCE, tmpControl.Left - GRIDDISTANCE)
            Next
        Case 3
            Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(1)).CtrlName)(SelectedCtrls(aSelectedPoints(1)).CtrlIndex)
            
            CurrLeft = tmpControl.Left + tmpControl.Width
            For i = 2 To UBound(aSelectedPoints)
                Set tmpControl = Me.Controls(SelectedCtrls(aSelectedPoints(i)).CtrlName)(SelectedCtrls(aSelectedPoints(i)).CtrlIndex)
                tmpControl.Left = CurrLeft
                
                CurrLeft = CurrLeft + tmpControl.Width
            Next
    End Select
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    ShowAttribute
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    Me.mnuViewToolButton.Checked = Not Me.mnuViewToolButton.Checked
    Me.CoolBar1.Visible = Me.mnuViewToolButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    Me.mnuViewToolText.Checked = Not Me.mnuViewToolText.Checked
    If Me.mnuViewToolText.Checked Then
        For i = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(i).Caption = Me.Toolbar1.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(i).Caption = ""
        Next
    End If
    Me.CoolBar1.Bands(1).MINHEIGHT = Me.Toolbar1.Height
    Me.CoolBar1.Refresh
    Form_Resize
End Sub

'处理鼠标画虚框
Private Sub PicForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then UnSelectAll: Me.PopupMenu Me.mnuFormat
    
    If Not Button = vbLeftButton Then Exit Sub
    If Not (Shift = vbShiftMask Or Shift = vbCtrlMask) Then UnSelectAll
    With shpSelect
        .Tag = "Start"
        
        iRangeX = x
        iRangeY = y
        If mnuFormatToGrid.Checked Then
            iRangeX = CLng(iRangeX / GRIDDISTANCE) * GRIDDISTANCE
            iRangeY = CLng(iRangeY / GRIDDISTANCE) * GRIDDISTANCE
        End If

        iRangeWidth = 0
        iRangeHeight = 0
    End With
End Sub
'
Private Sub PicForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mnuEdit_ViewList.Checked Then fraList.Visible = False
    If Not Button = vbLeftButton Or shpSelect.Tag <> "Start" Then Exit Sub
    
    With shpSelect
        '重画已画过的虚框以清除之
        DrawRect iRangeX, iRangeY, iRangeWidth, iRangeHeight, , IIf(CurrObjType = 4, 1, 0)
        
        iRangeWidth = x - iRangeX
        iRangeHeight = y - iRangeY
        If mnuFormatToGrid.Checked Then
            iRangeWidth = CLng(iRangeWidth / GRIDDISTANCE) * GRIDDISTANCE
            iRangeHeight = CLng(iRangeHeight / GRIDDISTANCE) * GRIDDISTANCE
        End If
        If CurrObjType = 4 Then
            If Abs(iRangeWidth) >= Abs(iRangeHeight) Then
                iRangeHeight = 0
            Else
                iRangeWidth = 0
            End If
        End If
    
        '画新框
        DrawRect iRangeX, iRangeY, iRangeWidth, iRangeHeight, , IIf(CurrObjType = 4, 1, 0)
    End With
End Sub
'
Private Sub PicForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpItem As MSComctlLib.ListItem
    Dim NewControl As Control, TableControl As Control
    Dim cellFormat As TTF160Ctl.F1CellFormat
    If Not Button = vbLeftButton Or shpSelect.Tag <> "Start" Then Exit Sub
    
    DrawRect iRangeX, iRangeY, iRangeWidth, iRangeHeight, , IIf(CurrObjType = 4, 1, 0)
    With shpSelect
        .Tag = ""
    End With
    If iRangeWidth < 0 Then iRangeX = iRangeX + iRangeWidth
    If iRangeHeight < 0 Then iRangeY = iRangeY + iRangeHeight
    If CurrObjType = 2 Then
        '选择虚框中的对象
        SelectRectCtrls iRangeX, iRangeX + Abs(iRangeWidth), iRangeY, iRangeY + Abs(iRangeHeight)
    Else
        If Abs(iRangeWidth) > 30 Or Abs(iRangeHeight) > 30 Then
            Select Case CurrObjType
                Case 3
                    Set NewControl = LoadNewControl("Text1")
                    With NewControl
                        .Left = iRangeX: .Top = iRangeY
                        .Width = Abs(iRangeWidth): .Height = Abs(iRangeHeight)
                        .Visible = True
                    End With
                    SelectControl NewControl, False
            
                    AddControlList "Text1", CStr(NewControl.Index), NewControl.Text
                Case 4
                    Set NewControl = LoadNewControl("Line1")
                    With NewControl
                        .Left = iRangeX: .Top = iRangeY
                        .Width = Abs(iRangeWidth): .Height = Abs(iRangeHeight)
                        .Visible = True
                    End With
                    SelectControl NewControl, False
            
                    AddControlList "Line1", CStr(NewControl.Index)
                Case 5
'                    With frmSelElement
'                        .pDepartID = ""
'                        .pFileType = -1
'                        .pElementType = "1"
'                        .Show vbModal
'                        If Len(.pElementID) > 0 Then
'                            Set TableControl = LoadNewControl("F1Book1")
'                            TableControl.Tag = .pElementID + ";" + .pElementName
'                            InitTable TableControl
'                            ReadTable TableControl, .pElementID
'
'                            Set NewControl = LoadNewControl("fraTable")
'                            Set TableControl.Container = NewControl
'                            NewControl.Left = iRangeX: NewControl.Top = iRangeY
'                            NewControl.Width = Abs(iRangeWidth): NewControl.Height = Abs(iRangeHeight)
'                            NewControl.Visible = True
'                            SelectControl NewControl, False
'
'                            AddControlList "fraTable", CStr(NewControl.Index)
'                        End If
'                    End With
'                    Unload frmSelElement
                    Set TableControl = LoadNewControl("F1Book1")
                    InitTable TableControl
                    
                    Set NewControl = LoadNewControl("fraTable")
                    Set TableControl.Container = NewControl
                    NewControl.Left = iRangeX: NewControl.Top = iRangeY
                    NewControl.Width = Abs(iRangeWidth): NewControl.Height = Abs(iRangeHeight)
                    NewControl.Visible = True
                    SelectControl NewControl, False
            
                    AddControlList "fraTable", CStr(NewControl.Index)
                Case 9
                    Set tmpItem = lvwSubItem.SelectedItem
                    If Not tmpItem Is Nothing Then
                        Set NewControl = LoadNewControl("VisItem1")
                        With NewControl
                            .Init tmpItem.Text, tmpItem.SubItems(8), tmpItem.SubItems(9), tmpItem.SubItems(5), tmpItem.SubItems(6), tmpItem.SubItems(7), tmpItem.SubItems(11), tmpItem.SubItems(13), Mid(tmpItem.Key, 5), IIf(tmpItem.SubItems(4) = "1", tmpItem.Text, "")
                            .Left = iRangeX: .Top = iRangeY
                            .Width = Abs(iRangeWidth): .Height = Abs(iRangeHeight)
                            .Tag = tmpItem.Text
                            .Visible = True
                        End With
                        SelectControl NewControl, False
                
                        AddControlList "VisItem1", CStr(NewControl.Index), tmpItem.Text
                    End If
            End Select
        End If
        ControlBar.Buttons(2).Value = tbrPressed
        
        CurrObjType = 2
        PicForm.MousePointer = vbDefault
        
        Modified = True
        
        ShowAttribute
    End If
End Sub
'处理选择标记点的拖动，实现已选控件的拖放
Private Sub shpDot_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim SelectIndex As Long, DotIndex As Long
    Dim SelectedCtrl As Control
    Dim LeftOffset As Long, TopOffset As Long, WidthOffset As Long, HeightOffset As Long
    Dim OldWidth As Long, OldHeight As Long
    If Not Button = vbLeftButton Then Exit Sub
    If mnuFormatLock.Checked Then Exit Sub
    
    SelectIndex = Int((Index - 1) / 8) + 1
    Set SelectedCtrl = Me.Controls(SelectedCtrls(SelectIndex).CtrlName)(SelectedCtrls(SelectIndex).CtrlIndex)
    
    If mnuFormatToGrid.Checked Then
        x = CInt(x / GRIDDISTANCE) * GRIDDISTANCE
        y = CInt(y / GRIDDISTANCE) * GRIDDISTANCE
    End If
    DotIndex = (Index - 1) Mod 8
    Select Case DotIndex
        Case 0
            LeftOffset = x: TopOffset = y: WidthOffset = -1 * x: HeightOffset = -1 * y
        Case 1
            If UCase(SelectedCtrl.Name) = "LINE1" And SelectedCtrl.Width > 15 Then
                TopOffset = y: LeftOffset = x
            Else
                TopOffset = y: HeightOffset = -1 * y
            End If
        Case 2
            TopOffset = y: WidthOffset = x: HeightOffset = -1 * y
        Case 3
            If UCase(SelectedCtrl.Name) = "LINE1" And SelectedCtrl.Height > 15 Then
                TopOffset = y: LeftOffset = x
            Else
                LeftOffset = x: WidthOffset = -1 * x
            End If
        Case 4
            If UCase(SelectedCtrl.Name) = "LINE1" And SelectedCtrl.Height > 15 Then
                TopOffset = y: LeftOffset = x
            Else
                WidthOffset = x
            End If
        Case 5
            LeftOffset = x:  WidthOffset = -1 * x: HeightOffset = y
        Case 6
            If UCase(SelectedCtrl.Name) = "LINE1" And SelectedCtrl.Width > 15 Then
                TopOffset = y: LeftOffset = x
            Else
                HeightOffset = y
            End If
        Case 7
            WidthOffset = x: HeightOffset = y
    End Select
    '注意线的处理
    On Error Resume Next
    With SelectedCtrl
        OldWidth = .Width: OldHeight = .Height
        If UCase(SelectedCtrl.Name) = "LINE1" Then
            If ((DotIndex = 1 Or DotIndex = 6) And SelectedCtrl.Width > 15) Or ((DotIndex = 3 Or DotIndex = 4) And SelectedCtrl.Height > 15) Then '移动控件（线）
                If LeftOffset <> 0 Then .Left = .Left + LeftOffset
                If TopOffset <> 0 Then .Top = .Top + TopOffset
            Else
                If .Width > 15 Then .Width = .Width + WidthOffset
                If .Height > 15 Then .Height = .Height + HeightOffset
            End If
        Else
            .Width = .Width + WidthOffset
            .Height = .Height + HeightOffset
        End If
        '对象最小时，不能再拖放
        If LeftOffset <> 0 Then .Left = .Left + OldWidth - .Width
        If TopOffset <> 0 Then .Top = .Top + OldHeight - .Height
    End With
    
    Modified = True
    
    SetCurrObject SelectIndex
    ShowSelect SelectedCtrl
    
    ShowAttribute
    PicForm.Refresh
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case vbLeftButton
            InitMoveControl Text1(Index), Button, Shift, x, y
        Case vbRightButton
    End Select
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ProcMoveControl Text1(Index), Button, Shift, x, y
End Sub

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    EndMoveControl Text1(Index), Button, Shift, x, y
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            If Modified Then
                If MsgBox("所见单已修改，是否保存", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuFileSave_Click
            End If
            mnuFileExit_Click
        Case "Lock"
            mnuFormatLock_Click
        Case "Align"
            Me.PopupMenu mnuFormatAlign
        Case "HDistance"
            Me.PopupMenu mnuFormatVsc
        Case "VDistance"
            Me.PopupMenu mnuFormatHsc
        Case "Size"
            Me.PopupMenu mnuFormatS
        Case "Save"
            mnuFileSave_Click
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub Toolbar2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ShowList 3800, lblItem.Top + lblItem.Height + 50 + fraToolbox.Top
    If lvwSubItem.ListItems.Count = 0 And Not tvwItem(iCurrTab).SelectedItem Is Nothing Then tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
End Sub

Private Sub VisItem1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    InitMoveControl VisItem1(Index), Button, Shift, x, y
    Select Case Button
        Case vbLeftButton
        Case vbRightButton
            Me.PopupMenu Me.mnuFormat
    End Select
End Sub

Private Sub VisItem1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ProcMoveControl VisItem1(Index), Button, Shift, x, y
End Sub

Private Sub VisItem1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    EndMoveControl VisItem1(Index), Button, Shift, x, y
End Sub

Private Sub VScroll_Change()
    PicForm.Top = -1 * VScroll.Value
End Sub
'画网格线
Private Sub DrawGrid(ByVal LineColor As Long)
    Dim iCurrY As Long, i As Long, iNum As Long
    
    For iCurrY = GRIDDISTANCE To PicForm.Height Step GRIDDISTANCE
        PicForm.Line (0, iCurrY)-(PicForm.Width, iCurrY), LineColor
    Next
    For iCurrY = GRIDDISTANCE To PicForm.Width Step GRIDDISTANCE
        PicForm.Line (iCurrY, 0)-(iCurrY, PicForm.Height), LineColor
    Next
End Sub
'显示对象的选择框
Private Sub ShowSelect(ByVal SelectedCtrlID As Control)
    Dim Index As Long
    Dim SelectedCtrl As Control
    Dim BackColor As Long
        
    Index = SeekControl(SelectedCtrlID)
    Set SelectedCtrl = SelectedCtrlID
'    If IsNumeric(SelectedCtrlID) Then
'        Index = Int((SelectedCtrlID - 1) / 8) + 1
'        Set SelectedCtrl = Me.Controls(SelectedCtrls(Index).CtrlName)(SelectedCtrls(Index).CtrlIndex)
'    Else
'        Index = SeekControl(SelectedCtrlID)
'        Set SelectedCtrl = SelectedCtrlID
'    End If
    If Index = CurrObject Then
        If mnuFormatLock.Checked Then
            BackColor = COLOR_YELLOW
        Else
            BackColor = COLOR_BLUE
        End If
    Else
        BackColor = COLOR_WHITE
    End If
    
    '注意线的处理
    With shpDot((Index - 1) * 8 + 1)
        .Left = SelectedCtrl.Left - .Width
        .Top = SelectedCtrl.Top - .Height
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 2)
        .Left = SelectedCtrl.Left + SelectedCtrl.Width / 2 - .Width / 2
        .Top = SelectedCtrl.Top - .Height
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 3)
        .Left = SelectedCtrl.Left + SelectedCtrl.Width
        .Top = SelectedCtrl.Top - .Height
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 4)
        .Left = SelectedCtrl.Left - .Width
        .Top = SelectedCtrl.Top + SelectedCtrl.Height / 2 - .Height / 2
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 5)
        .Left = SelectedCtrl.Left + SelectedCtrl.Width
        .Top = SelectedCtrl.Top + SelectedCtrl.Height / 2 - .Height / 2
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 6)
        .Left = SelectedCtrl.Left - .Width
        .Top = SelectedCtrl.Top + SelectedCtrl.Height
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 7)
        .Left = SelectedCtrl.Left + SelectedCtrl.Width / 2 - .Width / 2
        .Top = SelectedCtrl.Top + SelectedCtrl.Height
        .BackColor = BackColor
        .Visible = True
    End With
    With shpDot((Index - 1) * 8 + 8)
        .Left = SelectedCtrl.Left + SelectedCtrl.Width
        .Top = SelectedCtrl.Top + SelectedCtrl.Height
        .BackColor = BackColor
        .Visible = True
    End With
End Sub
'选择对象
Private Sub SelectControl(ByVal SelectedCtrl As Control, ByVal MultiSelect As Boolean, Optional ByVal ifUnload As Boolean = True)
    Dim iNum As Long, Index As Long, i As Long
    
    Index = SeekControl(SelectedCtrl)
    If Index > 0 Then
        If MultiSelect Then
            SelectedCtrls(Index).Visible = False
            
            For i = (Index - 1) * 8 + 1 To (Index - 1) * 8 + 8
                shpDot(i).Visible = False
            Next
            SetCurrObject NearestObjectIndex(Index)
        Else
            SetCurrObject Index
        End If
    Else
        If Not MultiSelect Then
            UnSelectAll ifUnload
        End If
        iNum = UBound(SelectedCtrls) + 1
        ReDim Preserve SelectedCtrls(iNum)
        With SelectedCtrls(iNum)
            .CtrlName = SelectedCtrl.Name
            .CtrlIndex = SelectedCtrl.Index
            .Visible = True
        End With
        For iNum = 1 To 8
            Load shpDot(shpDot.Count)
        Next
        shpDot(shpDot.Count - 1).MousePointer = vbSizeNWSE
        shpDot(shpDot.Count - 2).MousePointer = vbSizeNS
        shpDot(shpDot.Count - 3).MousePointer = vbSizeNESW
        shpDot(shpDot.Count - 4).MousePointer = vbSizeWE
        shpDot(shpDot.Count - 5).MousePointer = vbSizeWE
        shpDot(shpDot.Count - 6).MousePointer = vbSizeNESW
        shpDot(shpDot.Count - 7).MousePointer = vbSizeNS
        shpDot(shpDot.Count - 8).MousePointer = vbSizeNWSE
        
        SetCurrObject UBound(SelectedCtrls)
    End If
End Sub
'查找对象是否被选择
Private Function SeekControl(ByVal SelectedCtrl As Control) As Long
    Dim i As Long
    For i = 1 To UBound(SelectedCtrls)
        If SelectedCtrls(i).Visible = True And SelectedCtrl.Name = SelectedCtrls(i).CtrlName And SelectedCtrl.Index = SelectedCtrls(i).CtrlIndex Then Exit For
    Next
    If i > UBound(SelectedCtrls) Then i = -1
    
    SeekControl = i
End Function

Private Sub UnSelectAll(Optional ByVal ifUnload As Boolean = True)
    Dim i As Long, iNum As Long
    
    If ifUnload Then
        ReDim SelectedCtrls(0)
        
        iNum = shpDot.UBound
        For i = 1 To iNum
            Unload shpDot(i)
        Next
    Else
        iNum = UBound(SelectedCtrls)
        For i = 1 To iNum
            SelectedCtrls(i).Visible = False
        Next
        iNum = shpDot.UBound
        For i = 1 To iNum
            shpDot(i).Visible = False
        Next
    End If
    
    CurrObject = 0
End Sub
'画一个指定位置、大小的矩形虚框
Private Sub DrawRect(ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single, sngHeight As Single, Optional iPenStyle As Long = PS_DOT, Optional iDrawStyle As Long = 0)
    '下面这4个本地变量用于桌面坐标
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    Dim lngRight As Long, lngBottom As Long
    Dim lngPerX As Long, lngPerY As Long
    
    Dim lngPen As Long, p As POINTAPI, pLT As POINTAPI, pRB As POINTAPI
    Dim lngDC As Long, lngROP As Long
    
    
    lngPerX = Screen.TwipsPerPixelX
    lngPerY = Screen.TwipsPerPixelY
    
    
    '先把值从绨转换成象素
    lngLeft = sngLeft / lngPerX
    lngTop = sngTop / lngPerY
    
    lngWidth = sngWidth / lngPerX
    lngHeight = sngHeight / lngPerY
    '再进行坐标的转换
    p.x = lngLeft: p.y = lngTop
    ClientToScreen PicForm.hwnd, p
    lngLeft = p.x: lngTop = p.y
    With picMain
        pLT.x = .ScaleLeft / lngPerX
        pLT.y = .ScaleTop / lngPerY
        ClientToScreen picMain.hwnd, pLT '从现在开始该处保留最左、最上的值
        
        pRB.x = (.ScaleLeft + .ScaleWidth) / lngPerX
        pRB.y = (.ScaleTop + .ScaleHeight) / lngPerY
        ClientToScreen picMain.hwnd, pRB '从现在开始该处保留最右、最下的值
    End With
    '计算边界超出情况
    With picMain
        If sngLeft + sngWidth + PicForm.Left > .ScaleWidth Then
            lngRight = pRB.x
        Else
            lngRight = lngLeft + lngWidth
        End If
        If sngTop + sngHeight + PicForm.Top > .ScaleHeight Then
            lngBottom = pRB.y
        Else
            lngBottom = lngTop + lngHeight
        End If
        If sngLeft + sngWidth + PicForm.Left < .ScaleLeft Then lngRight = pLT.x
        If sngTop + sngHeight + PicForm.Top < .ScaleTop Then lngBottom = pLT.y
        
        If sngTop + PicForm.Top < .ScaleTop Then lngTop = pLT.y
        If sngLeft + PicForm.Left < .ScaleLeft Then lngLeft = pLT.x
    End With
    
    lngDC = GetDC(0)
    lngPen = SelectObject(lngDC, CreatePen(iPenStyle, 0, RGB(0, 0, 0)))
    lngROP = SetROP2(lngDC, R2_XORPEN)
    
    MoveToEx lngDC, lngLeft, lngTop, p
    If iDrawStyle = 0 Then
        LineTo lngDC, lngRight, lngTop
        LineTo lngDC, lngRight, lngBottom
        LineTo lngDC, lngLeft, lngBottom
        LineTo lngDC, lngLeft, lngTop
    Else
        LineTo lngDC, lngRight, lngBottom
    End If
    
    lngPen = SelectObject(lngDC, lngPen)
    SetROP2 lngDC, lngROP
    DeleteObject lngPen
    ReleaseDC 0, lngDC
End Sub
'选择矩形区域内的控件
Private Sub SelectRectCtrls(ByVal X1 As Long, ByVal X2 As Long, ByVal Y1 As Long, ByVal Y2 As Long)
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    Dim Top As Long, Bottom As Long, Left As Long, Right As Long
    
    For Each tmpCtrl In Me.Controls
        ValidCtrl = True
        On Error Resume Next
        If tmpCtrl.Container.Name <> "PicForm" Or Not tmpCtrl.Visible Then ValidCtrl = False
        On Error GoTo 0
        If ValidCtrl Then
        '注意线的处理
            With tmpCtrl
                Left = .Left
                Top = .Top
                Right = .Left + .Width
                Bottom = .Top + .Height
            End With
            If Not (Left > X2 Or Right < X1 Or Top > Y2 Or Bottom < Y1) Then SelectControl tmpCtrl, True
        End If
    Next
    
    ShowAttribute
End Sub

Private Sub SelectAll()
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    
    UnSelectAll
    For Each tmpCtrl In Me.Controls
        ValidCtrl = True
        On Error Resume Next
        If tmpCtrl.Container.Name <> "PicForm" Or Not tmpCtrl.Visible Then ValidCtrl = False
        On Error GoTo 0
        If ValidCtrl Then
        '注意线的处理
            SelectControl tmpCtrl, True
        End If
    Next
    
    ShowAttribute
End Sub

Private Function ControlsCount() As Integer
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    
    ControlsCount = 0
    For Each tmpCtrl In Me.Controls
        ValidCtrl = True
        On Error Resume Next
        If tmpCtrl.Container.Name <> "PicForm" Or Not tmpCtrl.Visible Then ValidCtrl = False
        On Error GoTo 0
        If ValidCtrl Then
        '注意线的处理
            If InStr("FRATABLE,VISITEM1", UCase(tmpCtrl.Name)) > 0 Then ControlsCount = ControlsCount + 1
        End If
    Next
End Function

Private Sub MoveSelectedCtrl(ByVal x As Long, ByVal y As Long)
    Dim i As Long, iNum As Long
    
    If mnuFormatToGrid.Checked Then
        x = CInt((x - iOrigX) / GRIDDISTANCE) * GRIDDISTANCE + iOrigX
        y = CInt((y - iOrigY) / GRIDDISTANCE) * GRIDDISTANCE + iOrigY
    End If
    
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            '注意线的处理
            With Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                .Left = .Left + x - iOrigX
                .Top = .Top + y - iOrigY
            End With
        End If
    Next
    PicForm.Refresh
End Sub

Private Function NearestObjectIndex(ByVal Index As Long)
    Dim i As Long, iNum As Long
    
    iNum = UBound(SelectedCtrls)
    For i = Index + 1 To iNum
        If SelectedCtrls(i).Visible Then Exit For
    Next
    If i > iNum Then
        For i = Index - 1 To 1 Step -1
            If SelectedCtrls(i).Visible Then Exit For
        Next
    End If
    NearestObjectIndex = i
End Function

Private Sub ShowAllDot(ByVal Visible As Boolean)
    Dim i As Long, iNum As Long, j As Long
    
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            For j = 1 To 8
                '注意线的处理,只有区域对象才存在八个选择点
                shpDot((i - 1) * 8 + j).Visible = Visible
            Next j
        End If
    Next i
End Sub

Private Sub SetCurrObject(ByVal ObjectIndex As Long)
    Dim OldCurrObject As Long
    If CurrObject = ObjectIndex Then Exit Sub
    
    OldCurrObject = CurrObject
    CurrObject = ObjectIndex
    If OldCurrObject > 0 And SelectedCtrls(OldCurrObject).Visible Then ShowSelect Me.Controls(SelectedCtrls(OldCurrObject).CtrlName)(SelectedCtrls(OldCurrObject).CtrlIndex)
    If CurrObject > 0 Then ShowSelect Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
End Sub
'一下三个过程处理对象点击及其移动
'分别对应：MouseDown、MouseMove和MouseUp事件
Private Sub InitMoveControl(ByVal theControl As Control, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = vbShiftMask Or Shift = vbCtrlMask Then
        SelectControl theControl, True
    Else
        SelectControl theControl, False
    End If
    '设置拖动起始点
    iOrigX = x: iOrigY = y
    
    ShowAttribute
End Sub

Private Sub ProcMoveControl(ByVal theControl As Control, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Button = vbLeftButton Then Exit Sub
    If SeekControl(theControl) < 1 Then Exit Sub
    If mnuFormatLock.Checked Then Exit Sub
    
    If x <> iOrigX Or y <> iOrigY Then ShowAllDot False
    MoveSelectedCtrl x, y
    
    Modified = True
End Sub

Private Sub EndMoveControl(ByVal theControl As Control, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long, iNum As Long
    If Not Button = vbLeftButton Then Exit Sub
    If SeekControl(theControl) < 1 Then Exit Sub
    If mnuFormatLock.Checked Then Exit Sub
    
'    MoveSelectedCtrl X, Y
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    ShowAttribute
End Sub

Private Sub ControlsAlign(ByVal AlignMode As Integer)
    Dim i As Long, iNum As Long
    Dim tmpControl As Control
    Dim CurrLeft As Long, CurrWidth As Long, CurrTop As Long, CurrHeight As Long
    iNum = UBound(SelectedCtrls)
    If SelectedCounts < 2 Then Exit Sub
    
    '注意线的处理
    On Error Resume Next
    Select Case AlignMode
        Case 0 '左对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrLeft = tmpControl.Left
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Left = CurrLeft
                End If
            Next
        Case 1 '居中对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrLeft = tmpControl.Left: CurrWidth = tmpControl.Width
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Left = CurrLeft + (CurrWidth - tmpControl.Width) / 2
                End If
            Next
        Case 2 '右对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrLeft = tmpControl.Left: CurrWidth = tmpControl.Width
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Left = CurrLeft + (CurrWidth - tmpControl.Width)
                End If
            Next
        Case 3 '顶端对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrTop = tmpControl.Top
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Top = CurrTop
                End If
            Next
        Case 4 '中间对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrTop = tmpControl.Top: CurrHeight = tmpControl.Height
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Top = CurrTop + (CurrHeight - tmpControl.Height) / 2
                End If
            Next
        Case 5 '底端对齐
            Set tmpControl = Me.Controls(SelectedCtrls(CurrObject).CtrlName)(SelectedCtrls(CurrObject).CtrlIndex)
            CurrTop = tmpControl.Top: CurrHeight = tmpControl.Height
            For i = 1 To iNum
                If SelectedCtrls(i).Visible Then
                    Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
                    tmpControl.Top = CurrTop + (CurrHeight - tmpControl.Height)
                End If
            Next
    End Select
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            ShowSelect Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    ShowAttribute
End Sub

Private Sub DeleteCtrls()
    Dim i As Long, iNum As Long
    
    On Error Resume Next
    iNum = UBound(SelectedCtrls)
        
    bNotRunCombox_Click = True '不要运行cmbControl_Click事件处理
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            Select Case UCase(SelectedCtrls(i).CtrlName)
                Case "TEXT1"
                    cmbControl.Text = "标签(" & SelectedCtrls(i).CtrlIndex & ") " + Trim(Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex).Text)
                Case "LINE1"
                    cmbControl.Text = "线(" & SelectedCtrls(i).CtrlIndex & ") "
                Case "VISITEM1"
                    cmbControl.Text = "所见项(" & SelectedCtrls(i).CtrlIndex & ") " + Trim(Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex).Tag)
                Case "FRATABLE"
                    cmbControl.Text = "附加表(" & SelectedCtrls(i).CtrlIndex & ") "
                    ClearAllObject F1Book1(SelectedCtrls(i).CtrlIndex)
                    Unload F1Book1(SelectedCtrls(i).CtrlIndex)
            End Select
            cmbControl.RemoveItem cmbControl.ListIndex
            
            Unload Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
        End If
    Next
    bNotRunCombox_Click = False
    
    UnSelectAll
    
    ShowAttribute True
End Sub

Private Sub DeleteAllCtrls()
    Dim i As Long, iNum As Long
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    
    On Error Resume Next
    cmbControl.Clear
    
    For Each tmpCtrl In Me.Controls
        ValidCtrl = True
        If UCase(tmpCtrl.Container.Name) <> "PICFORM" Or Not tmpCtrl.Visible Then ValidCtrl = False
        If ValidCtrl Then
            Select Case UCase(tmpCtrl.Name)
                Case "FRATABLE"
                    ClearAllObject F1Book1(tmpCtrl.Index)
                    Unload F1Book1(tmpCtrl.Index)
            End Select
            Unload tmpCtrl
        End If
    Next
    
    UnSelectAll
    ShowAttribute True
End Sub

Private Sub ShowAttribute(Optional ByVal SetCombox As Boolean = True)
'SetCombox:是否设置Combox 的值
    Dim i As Long, iNum As Long
    Dim tmpControl As Control, FirstControlIndex As Long
    iNum = grdAttr.Rows - 1
    For i = 1 To iNum
        grdAttr.RemoveItem 1
    Next
    grdAttr.Clear
    
    On Error Resume Next
    bNotRunCombox_Click = True '不要运行cmbControl_Click事件处理
    If SetCombox Then cmbControl.Text = " "
    bNotRunCombox_Click = False
    If SelectedCounts = 0 Then Exit Sub
    With grdAttr
        .Text(0, 0) = "左边距"
        .AddItem ""
        .Text(1, 0) = "上边距"
        .AddItem ""
        .Text(2, 0) = "宽度"
        .AddItem ""
        .Text(3, 0) = "高度"
    End With
    If SelectedCounts > 1 Then
        Exit Sub
    End If
    
    bNotRunCombox_Click = True '不要运行cmbControl_Click事件处理
    FirstControlIndex = FirstSelectedIndex()
    
    Set tmpControl = Me.Controls(SelectedCtrls(FirstControlIndex).CtrlName)(SelectedCtrls(FirstControlIndex).CtrlIndex)
    Select Case UCase(SelectedCtrls(FirstControlIndex).CtrlName)
        Case "TEXT1"
            With grdAttr
                .AddItem ""
                .Text(4, 0) = "标题"
                .AddItem ""
                .Text(5, 0) = "对齐"
                
                .Text(0, 1) = tmpControl.Left
                .Text(1, 1) = tmpControl.Top
                .Text(2, 1) = tmpControl.Width
                .Text(3, 1) = tmpControl.Height
                .Text(4, 1) = tmpControl.Text
                
                .CellType(5, 1) = editComboBox
                .List_AddItem 5, 1, "左对齐"
                .List_AddItem 5, 1, "右对齐"
                .List_AddItem 5, 1, "居中"
                .ListIndex(5, 1) = tmpControl.Alignment
            End With
            If SetCombox Then cmbControl.Text = "标签(" + Trim(CStr(SelectedCtrls(FirstControlIndex).CtrlIndex)) + ") " + Trim(Me.Controls(SelectedCtrls(FirstControlIndex).CtrlName)(SelectedCtrls(FirstControlIndex).CtrlIndex).Text)
        Case "LINE1", "FRATABLE"
            With grdAttr
                .Text(0, 1) = CStr(tmpControl.Left)
                .Text(1, 1) = tmpControl.Top
                .Text(2, 1) = tmpControl.Width
                .Text(3, 1) = tmpControl.Height
            End With
            If UCase(SelectedCtrls(FirstControlIndex).CtrlName) = "LINE1" Then
                If SetCombox Then cmbControl.Text = "线(" + Trim(CStr(SelectedCtrls(FirstControlIndex).CtrlIndex)) + ") "
            Else
                With grdAttr
                    .AddItem ""
                    .Text(.Rows - 1, 0) = "输入顺序"
                    .Text(.Rows - 1, 1) = tmpControl.TabIndex
                End With
                If SetCombox Then cmbControl.Text = "附加表(" + Trim(CStr(SelectedCtrls(FirstControlIndex).CtrlIndex)) + ") "
            End If
        Case "VISITEM1"
            With grdAttr
                .AddItem ""
                .Text(4, 0) = "不可输入"
                .CellType(4, 1) = editComboBox
                .List_AddItem 4, 1, "否"
                .List_AddItem 4, 1, "是"
                .ListIndex(4, 1) = IIf(tmpControl.Enabled, 0, 1)
            
                .AddItem ""
                .Text(5, 0) = "可屏蔽"
                .CellType(5, 1) = editComboBox
                .List_AddItem 5, 1, "否"
                .List_AddItem 5, 1, "是"
                .ListIndex(5, 1) = IIf(tmpControl.AllowMask, 1, 0)
                
                .AddItem ""
                .Text(6, 0) = "标题"
                .Text(6, 1) = tmpControl.Title
                
                If Len(tmpControl.ExchangeField) = 0 Then
                    .AddItem ""
                    .Text(7, 0) = "缺省值"
                    If tmpControl.ItemType = 3 Then
                        .CellType(7, 1) = editComboBox
                        .List_AddItem 7, 1, ""
                        .List_AddItem 7, 1, "是"
                        .List_AddItem 7, 1, "否"
                    Else
                        If tmpControl.Method > 1 Then
                            .CellType(7, 1) = editComboBox
                            .List_AddItem 7, 1, ""
                            
                            iNum = tmpControl.ValuesCount
                            For i = 0 To iNum - 1
                                .List_AddItem 7, 1, tmpControl.Values(i)
                            Next
                        End If
                    End If
                    .Text(7, 1) = tmpControl.Value
                
                    .AddItem ""
                    .Text(8, 0) = "计量单位"
                    .Text(8, 1) = tmpControl.Unit
                End If
                
                .Text(0, 1) = tmpControl.Left
                .Text(1, 1) = tmpControl.Top
                .Text(2, 1) = tmpControl.Width
                .Text(3, 1) = tmpControl.Height
            
                .AddItem ""
                .Text(.Rows - 1, 0) = "输入顺序"
                .Text(.Rows - 1, 1) = tmpControl.TabIndex
            End With
            If SetCombox Then cmbControl.Text = "所见项(" + Trim(CStr(SelectedCtrls(FirstControlIndex).CtrlIndex)) + ") " + Trim(Me.Controls(SelectedCtrls(FirstControlIndex).CtrlName)(SelectedCtrls(FirstControlIndex).CtrlIndex).Tag)
    End Select
    
    bNotRunCombox_Click = False
End Sub

Private Function SelectedCounts() As Long
    Dim i As Long, iNum As Long
    
    SelectedCounts = 0
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then SelectedCounts = SelectedCounts + 1
    Next
End Function

Private Function FirstSelectedIndex() As Long
    Dim i As Long, iNum As Long
    
    FirstSelectedIndex = 0
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            FirstSelectedIndex = i
            Exit For
        End If
    Next
End Function

Private Sub AddControlList(ByVal CtrlName As String, ByVal CtrlIndex As String, Optional ByVal Caption As String = "")
    With cmbControl
        Select Case UCase(CtrlName)
            Case "TEXT1"
                .AddItem "标签(" + Trim(CtrlIndex) + ") " + Trim(Caption)
            Case "LINE1"
                .AddItem "线(" + Trim(CtrlIndex) + ") " + Trim(Caption)
            Case "VISITEM1"
                .AddItem "所见项(" + Trim(CtrlIndex) + ") " + Trim(Caption)
            Case "FRATABLE"
                .AddItem "附加表(" + Trim(CtrlIndex) + ") " + Trim(Caption)
        End Select
    End With
End Sub

Private Sub SetControlAttr(ByVal AttributeID As String, vNewValue As Variant)
    Dim i As Long, iNum As Long
    Dim tmpControl As Control
    Dim iCounts As Integer
    On Error Resume Next
    
    iNum = UBound(SelectedCtrls)
    For i = 1 To iNum
        If SelectedCtrls(i).Visible Then
            Set tmpControl = Me.Controls(SelectedCtrls(i).CtrlName)(SelectedCtrls(i).CtrlIndex)
            Select Case AttributeID
                Case "左边距"
                    tmpControl.Left = vNewValue
                    vNewValue = tmpControl.Left
                Case "上边距"
                    tmpControl.Top = vNewValue
                    vNewValue = tmpControl.Top
                Case "宽度"
                    If Not (UCase(SelectedCtrls(i).CtrlName) = "LINE1" And tmpControl.Height > 15) Then
                        tmpControl.Width = vNewValue
                    Else
                        vNewValue = tmpControl.Width
                    End If
                    vNewValue = tmpControl.Width
                Case "高度"
                    If Not (UCase(SelectedCtrls(i).CtrlName) = "LINE1" And tmpControl.Width > 15) Then
                        tmpControl.Height = vNewValue
                    Else
                        vNewValue = tmpControl.Height
                    End If
                    vNewValue = tmpControl.Height
                Case "标题"
                    Select Case UCase(SelectedCtrls(i).CtrlName)
                        Case "TEXT1"
                            tmpControl.Text = vNewValue
                            cmbControl.List(cmbControl.ListIndex) = "标签(" + Trim(SelectedCtrls(i).CtrlIndex) + ") " + Trim(vNewValue)
                        Case "VISITEM1"
                            tmpControl.Title = vNewValue
                            'cmbControl.List(cmbControl.ListIndex) = "所见项(" + Trim(SelectedCtrls(i).CtrlIndex) + ") " + Trim(vNewValue)
                    End Select
                Case "对齐"
                    tmpControl.Alignment = vNewValue
                Case "不可输入"
                    tmpControl.Enabled = IIf(vNewValue = "否", True, False)
                Case "可屏蔽"
                    tmpControl.AllowMask = IIf(vNewValue = "否", False, True)
                Case "缺省值"
                    tmpControl.Value = vNewValue
                Case "计量单位"
                    tmpControl.Unit = vNewValue
                Case "输入顺序"
                    iCounts = ControlsCount
                    If vNewValue >= iCounts Then vNewValue = iCounts - 1
                    
                    tmpControl.TabIndex = vNewValue
                    vNewValue = tmpControl.TabIndex
            End Select
            ShowSelect tmpControl
        End If
    Next
    
    Modified = True
End Sub

'创建所见项分类及其项目的TreeView
Private Sub CreateItemTree()
    Dim rsItem As New ADODB.Recordset
    Dim sCurID As String
    Dim iStackPoint As Integer '堆栈指针
    Dim aStack() As String '堆栈
    Dim TmpNode As Node
    Dim i As Integer, AttrID As String
    
    '从诊治所见性质中提取
    clsDatabase.OpenRecordset rsItem, "Select * From 诊治所见性质 Order By 编码", ""
    Do While Not rsItem.EOF
        Load cmdTab(cmdTab.Count)
        With cmdTab(cmdTab.Count - 1)
            .Caption = rsItem("名称") '+ IIf(rsItem("固定") = 1, "（只读）", "")
            .Tag = rsItem("固定") & "-" & rsItem("编码")
            .Visible = True
        End With
        Load tvwItem(tvwItem.Count)
        tvwItem(tvwItem.Count - 1).Visible = True
        
        rsItem.MoveNext
    Loop
    
    For i = 1 To cmdTab.Count - 1
        AttrID = Mid(cmdTab(i).Tag, InStr(cmdTab(i).Tag, "-") + 1)
    
        ReDim aStack(0)
        aStack(0) = ""
        iStackPoint = 0
        
        Do While iStackPoint > -1
            sCurID = aStack(iStackPoint)
            '添加下级所见项分类
            gstrSql = "Select * From 诊治所见分类 A Where A.上级ID" + IIf(sCurID = "", " is null ", "=[1] ") + "And 性质=[2] " + _
                " And EXISTS (SELECT 1 FROM 诊治所见项目 B WHERE B.表示法<=4 AND A.ID=B.分类ID)"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "查询所见项目分类", sCurID, AttrID)
                        
            '该分类的下级已处理，将其从堆栈中弹出
            iStackPoint = iStackPoint - 1
            
            Do While Not rsItem.EOF
                If sCurID = "" Then
                    Set TmpNode = tvwItem(i).Nodes.Add(, , "Key" & rsItem("ID"), rsItem("名称"), "Class")
                Else
                    Set TmpNode = tvwItem(i).Nodes.Add("Key" + sCurID, tvwChild, "Key" & rsItem("ID"), rsItem("名称"), "Class")
                End If
                TmpNode.Tag = rsItem("性质") & "||" & rsItem("编码") & "||" & rsItem("名称") & "||" & rsItem("简码")
                
                '将新分类压入堆栈
                iStackPoint = iStackPoint + 1
                ReDim Preserve aStack(iStackPoint)
                aStack(iStackPoint) = rsItem("ID")
                
                rsItem.MoveNext
            Loop
        Loop
    Next
End Sub

Private Sub ShowSubItem(ByVal NodeID As String, ByVal AttributeID As String)
    Dim rsItem As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim sSQL As String, sItemIcon As String
    lvwSubItem.ListItems.Clear
    '添加下级所见项目
    sSQL = "Select ID,编码,中文名,nvl(英文名,' '),nvl(替换域,0),nvl(类型,0)," + _
       "nvl(长度,10),nvl(小数,0),nvl(单位,' '),nvl(表示法,0),nvl(性别域,0)," + _
       "nvl(数值域,' '),nvl(正常域,' '),nvl(初始值,' '),nvl(文字表述,1),nvl(空值文字,' '),nvl(临床意义,' ') " + _
       "From 诊治所见项目 Where " + IIf(NodeID = "", "性质=[1] And 分类ID is null ", "分类ID=[2] ") + _
       "And 表示法<=4"
    Set rsItem = zldatabase.OpenSQLRecord(sSQL, "查询所见项目", AttributeID, NodeID)
        
    Do While Not rsItem.EOF
        Select Case rsItem(9)
            Case 0, 1
                sItemIcon = "Text"
            Case 2
                sItemIcon = "Combox"
            Case 3
                sItemIcon = "Check"
            Case 4
                sItemIcon = "Option"
        End Select
        Set tmpItem = lvwSubItem.ListItems.Add(, "Item" & rsItem(0), rsItem(2), sItemIcon, sItemIcon)
        tmpItem.SubItems(1) = rsItem(1)
        tmpItem.SubItems(3) = rsItem(3)
        tmpItem.SubItems(4) = rsItem(4)
        tmpItem.SubItems(5) = rsItem(5)
        tmpItem.SubItems(6) = rsItem(6)
        tmpItem.SubItems(7) = rsItem(7)
        tmpItem.SubItems(8) = rsItem(8)
        tmpItem.SubItems(9) = rsItem(9)
        tmpItem.SubItems(10) = rsItem(10)
        tmpItem.SubItems(11) = rsItem(11)
        tmpItem.SubItems(12) = rsItem(12)
        tmpItem.SubItems(13) = rsItem(13)
        tmpItem.SubItems(14) = rsItem(14)
        tmpItem.SubItems(15) = rsItem(15)
        tmpItem.SubItems(16) = rsItem(16)
        
        rsItem.MoveNext
    Loop
End Sub

Private Sub tvwItem_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    If Node Is Nothing Then Exit Sub
    If Node.Key Like "Key_*" Then
        ShowSubItem "", Mid(Node.Key, 5)
    Else
        ShowSubItem Mid(Node.Key, 4), ""
    End If
End Sub

Private Sub ShowList(ByVal Width As Long, Optional ByVal Top As Long = -1)
    Dim i As Integer
    Dim lngTools As Single, lngStatus As Single
    
    lngTools = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    With fraList
        .Left = 0: .Top = IIf(Top = -1, CoolBar1.Top + lngTools, Top)
        .Width = Width
        .Height = Me.ScaleHeight - lngStatus - .Top
        .Visible = True
    End With
    For i = 1 To tvwItem.Count - 1
        tvwItem(i).Visible = IIf(i = iCurrTab, True, False)
        With cmdTab(i)
            If i <= iCurrTab Then
                .Top = (i - 1) * cmdTab(0).Height
            Else
                .Top = fraList.Height - (tvwItem.Count - i) * cmdTab(0).Height
            End If
            
            .Width = fraList.Width
            .Left = 0
            
            .Visible = True
        End With
    Next
    
    With tvwItem(iCurrTab)
        .Left = 0
        .Top = cmdTab(iCurrTab).Top + cmdTab(iCurrTab).Height
        .Width = fraList.Width
        .Height = fraList.Height - (tvwItem.Count - iCurrTab - 1) * cmdTab(0).Height - .Top
    End With
End Sub

Private Function LoadNewControl(ControlName As String) As Control
    Dim i As Integer, bLoop As Boolean
    Dim iControlWidth As Long
    
    On Error Resume Next
    i = 1
    bLoop = True
    iControlWidth = -1
    Do While bLoop
        iControlWidth = Me.Controls(ControlName)(i).Width
        If iControlWidth = -1 Then
            Load Me.Controls(ControlName)(i)
            If InStr("FRATABLE,VISITEM1", UCase(ControlName)) > 0 Then Me.Controls(ControlName)(i).TabIndex = ControlsCount
            Set LoadNewControl = Me.Controls(ControlName)(i)
            bLoop = False
        Else
            iControlWidth = -1
            i = i + 1
        End If
    Loop
End Function
'
'Private Sub CopyTable(srcTable As TTF160Ctl.F1Book, dstTable As TTF160Ctl.F1Book)
'    Dim CellFormat As TTF160Ctl.F1CellFormat
'    With dstTable
'        .MaxRow = srcTable.MaxRow: .MaxCol = srcTable.MaxCol
'        .FixedRows = srcTable.FixedRows: .FixedCols = srcTable.FixedCols
'
'        srcTable.SetSelection 1, 1, srcTable.MaxRow, srcTable.MaxCol
'        Set CellFormat = srcTable.GetCellFormat
'        .SetCellFormat CellFormat
'    End With
'End Sub

Private Sub CreateAllCtrlList()
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    
    UnSelectAll
    For Each tmpCtrl In Me.Controls
        ValidCtrl = True
        On Error Resume Next
        If tmpCtrl.Container.Name <> "PicForm" Or tmpCtrl.Index = 0 Then ValidCtrl = False
        On Error GoTo 0
        If ValidCtrl Then
        '注意线的处理
            Select Case UCase(tmpCtrl.Name)
                Case "TEXT1"
                    AddControlList "Text1", CStr(tmpCtrl.Index), tmpCtrl.Text
                Case "LINE1"
                    AddControlList "Line1", CStr(tmpCtrl.Index)
                Case "FRATABLE"
                    AddControlList "fraTable", CStr(tmpCtrl.Index)
                Case "VISITEM1"
                    AddControlList "VisItem1", CStr(tmpCtrl.Index), tmpCtrl.Tag
            End Select
        End If
    Next
End Sub

Private Sub BeginShowProgress()
    With prbRefresh
        .Left = stbThis.Panels(2).Left + 50
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width - 50
        .Visible = stbThis.Visible
    End With
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

