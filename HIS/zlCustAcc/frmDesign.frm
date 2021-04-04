VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDesign 
   BackColor       =   &H00808080&
   Caption         =   "专项记帐单设置"
   ClientHeight    =   6855
   ClientLeft      =   -135
   ClientTop       =   240
   ClientWidth     =   9090
   Icon            =   "frmDesign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   2220
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCombo 
      Height          =   645
      Left            =   3660
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   29
      Top             =   5730
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   6390
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9090
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "Save"
               Object.ToolTipText     =   "保存"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Caption         =   "新增"
               Key             =   "New"
               Object.ToolTipText     =   "新增记帐单"
               Object.Tag             =   "新增"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除记帐单"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "元素"
               Key             =   "Element"
               Object.ToolTipText     =   "记帐单元素"
               Object.Tag             =   "元素"
               ImageKey        =   "Element"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "对齐"
               Key             =   "Align"
               Object.ToolTipText     =   "左对齐"
               Object.Tag             =   "对齐"
               ImageKey        =   "Align"
               Style           =   5
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
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "间距"
               Key             =   "Distance"
               Object.ToolTipText     =   "横间距相同"
               Object.Tag             =   "间距"
               ImageKey        =   "Distance"
               Style           =   5
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
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "大小"
               Key             =   "Size"
               Object.ToolTipText     =   "相同宽度"
               Object.Tag             =   "大小"
               ImageKey        =   "Size"
               Style           =   5
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
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "锁定"
               Key             =   "Lock"
               Object.ToolTipText     =   "锁定"
               Object.Tag             =   "锁定"
               ImageKey        =   "Lock"
               Style           =   1
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6492
      Width           =   9084
      _ExtentX        =   16034
      _ExtentY        =   635
      SimpleText      =   $"frmDesign.frx":0442
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDesign.frx":0489
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
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
   Begin VB.PictureBox picSplitRight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   5220
      MousePointer    =   9  'Size W E
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   20
      Top             =   2310
      Width           =   60
   End
   Begin VB.PictureBox picSplitLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2310
      MousePointer    =   9  'Size W E
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   19
      Top             =   2640
      Width           =   60
   End
   Begin VB.HScrollBar HScroll 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4590
      Width           =   1245
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1245
      Left            =   4620
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2790
      Width           =   285
   End
   Begin VB.Frame fraAttrib 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   5700
      TabIndex        =   9
      Top             =   1440
      Width           =   2250
      Begin VB.CommandButton cmdEdit 
         Caption         =   "…"
         Height          =   240
         Left            =   1740
         TabIndex        =   15
         Top             =   1530
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.ComboBox cmbControl 
         Height          =   300
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   510
         Width           =   2055
      End
      Begin VB.PictureBox picClose 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   7.5
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1440
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   315
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   1140
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAttrib 
         Height          =   2895
         Left            =   120
         TabIndex        =   17
         Top             =   930
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridColorUnpopulated=   8421504
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "属性表框"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   210
         Width           =   1665
      End
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5280
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":0D1D
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":0F39
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1159
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1379
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1599
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":17B9
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":19D9
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":20D3
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23ED
            Key             =   "Align"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2AE7
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":31E1
            Key             =   "Distance"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":38DB
            Key             =   "Size"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":3FD5
            Key             =   "Lock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4470
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":46CF
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":48EB
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":4B0B
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":4D2B
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":4F4B
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":516B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":538B
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":5A85
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":5D9F
            Key             =   "Align"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":6499
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":6B93
            Key             =   "Distance"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":728D
            Key             =   "Size"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":7987
            Key             =   "Lock"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2085
      Begin MSComctlLib.ImageList ils16 
         Left            =   1260
         Top             =   3000
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
               Picture         =   "frmDesign.frx":8081
               Key             =   "Bill"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ils32 
         Left            =   300
         Top             =   3000
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
               Picture         =   "frmDesign.frx":8ED3
               Key             =   "Bill"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picClose 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   7.5
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1470
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   315
      End
      Begin MSComctlLib.ListView lvwMain 
         Height          =   2205
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "名称"
            Text            =   "名称"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编号"
            Text            =   "编号"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "门诊病人"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "住院统一"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "住院科室"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "医技科室"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "记帐单列表"
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.Frame fraCorner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4500
      TabIndex        =   7
      Top             =   4500
      Width           =   285
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2640
      ScaleHeight     =   3015
      ScaleWidth      =   2655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2655
      Begin VB.CheckBox chk 
         Caption         =   "销"
         Height          =   360
         Index           =   1
         Left            =   1800
         MousePointer    =   5  'Size
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         Index           =   1
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1560
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   270
         Locked          =   -1  'True
         MousePointer    =   5  'Size
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   285
         Index           =   0
         Left            =   540
         MousePointer    =   5  'Size
         TabIndex        =   26
         Top             =   420
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Caption         =   "取消(&C)"
         Height          =   350
         Index           =   2
         Left            =   180
         MousePointer    =   5  'Size
         TabIndex        =   25
         Top             =   2310
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmd 
         Caption         =   "确定(&O)"
         Height          =   350
         Index           =   1
         Left            =   270
         MousePointer    =   5  'Size
         TabIndex        =   24
         Top             =   1950
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   15  'Size All
         TabIndex        =   23
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chk 
         Caption         =   "加班"
         Height          =   360
         Index           =   0
         Left            =   960
         MousePointer    =   5  'Size
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Shape shpSelect 
         BorderStyle     =   3  'Dot
         Height          =   465
         Left            =   1680
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblAdjust 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   90
         Index           =   0
         Left            =   1770
         TabIndex        =   21
         Top             =   1500
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "记帐单"
         Height          =   180
         Index           =   0
         Left            =   210
         MousePointer    =   5  'Size
         TabIndex        =   27
         Top             =   1290
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.Frame fraAdjust 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   0
      Left            =   4470
      TabIndex        =   18
      Top             =   5190
      Visible         =   0   'False
      Width           =   105
      Begin VB.Shape shpAdjust 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   75
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   75
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileDesign 
         Caption         =   "设计(&D)"
         Shortcut        =   ^D
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
      Begin VB.Menu mnuFileNew 
         Caption         =   "新增(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileErase 
         Caption         =   "删除(&R)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImp 
         Caption         =   "导入(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileExp 
         Caption         =   "导出(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditElements 
         Caption         =   "记帐单元素(&E)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditText 
         Caption         =   "增加文本(&T)"
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "增加线条(&L)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制标签(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "删除标签(&R)"
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
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "上对齐(&U)"
            Index           =   4
         End
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "中间对齐(&M)"
            Index           =   5
         End
         Begin VB.Menu mnuFormatDoAlign 
            Caption         =   "下对齐(&D)"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFormatForm 
         Caption         =   "在窗口内居中对齐(&W)"
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
         Caption         =   "大小(&S)"
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
         Caption         =   "竖间距(&V)"
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "相同(&S)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatVscSpace 
            Caption         =   "增加(&A)"
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
         Caption         =   "横间距(&H)"
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "相同(&S)"
            Index           =   0
         End
         Begin VB.Menu mnuFormatHscSpace 
            Caption         =   "增加(&A)"
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
         Begin VB.Menu mnuView1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
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
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "记帐单列表(&L)"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewAttrib 
         Caption         =   "属性表框(&A)"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuView4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
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
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
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
Attribute VB_Name = "frmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mlngWidthAdj As Long = 105  '位置改变方块的边长

Dim msngX As Single, msngY As Single
Dim mblnDown As Boolean           '表示鼠标按下，准备拖动
Dim mlngColumn As Long            '记录上一次ListView的排序列
Dim mstrKey As String             '记录上一次ListView的选择行
Dim mlngRow As Long               '上一次属性表的选中行
Dim mbln新增 As Boolean           '处于新增状态
Dim mblnLoad As Boolean           '处于刚刚装入
Dim mblnChange As Boolean         '当前记帐单是否经过修改

Dim mlngMoveReason As Long         '表示鼠标在窗口中拖动的原因
Dim mintIndex As Integer           '当前选中的ComboBox

Dim mcolBill As Elements        '属于本张单据的控件
Dim mcolSelect As Collection      '当前选中的控件集合。如果一项都没有，则表示选中picForm
Dim mctlSelect As Control         '当前选中的控件集合中参照控件

Dim mstr编号 As String      '本张记帐单的编号
Dim mstr名称 As String      '本张记帐单的名称
Dim mstr适用范围 As String  '本张记帐单的适用范围
Dim mlng数量 As Long        '本张记帐单的收费项目数

Dim mintAlign As Integer        '上次使用的对齐方法
Dim mintForm  As Integer        '上次使用的居中方法
Dim mintDistance As Integer     '上次使用的间距方法
Dim mintSize     As Integer     '上次使用的大小方法

Dim mblnReadonly As Boolean     '是否只有只读的权限

Public Sub SelectSame(ctlSelect As Control)
'双击控件，选中同类型的所有控件
    Dim objTemp As Element
    Dim ctlTemp As Control
    Dim strType As String
    Dim i As Long
    
    
    strType = TypeName(ctlSelect) '得到参考对象的类型
    '首先清除已有的选择控件集
    For i = 1 To mcolSelect.Count
        mcolSelect.Remove 1
    Next
    
    On Error Resume Next
    For Each objTemp In mcolBill
        Set ctlTemp = objTemp.Control
        If ctlTemp.Visible = True And objTemp.Visible = True Then
            '除了要求单据本身的控件，还有一些辅助控件也要排除
            '判断控件是否选中
            If TypeName(ctlTemp) = strType Then
                mcolSelect.Add ctlTemp, ctlTemp.Name & ctlTemp.Index
            End If
        End If
    Next
    Set mctlSelect = ctlSelect '双击对象是所选对象
    
    If mcolSelect.Count = 1 Then
        For i = 0 To cmbControl.ListCount - 1
            If GetFore(cmbControl.List(i)) = mctlSelect.Tag Then
                cmbControl.ListIndex = i
                Exit Sub
            End If
        Next
    Else
        '什么属性也不显示
        If cmbControl.ListIndex = -1 Then
            '以前就是选中各个控件，再赋值也就不会激活事件
            '所以只有手工调用
            Call HideAttach
            Call ShowAttach
            Call ShowAttrib
        Else
            cmbControl.ListIndex = -1
        End If
    End If

End Sub

Private Sub cmd_Click(Index As Integer)
    Static timPrevious As Single
    Dim timNow As Single
    
    timNow = Timer
    If timPrevious <> 0 Then
       
       If timNow - timPrevious < 0.5 Then
            timPrevious = 0
            Call SelectSame(cmd(Index))
            Exit Sub
       End If
    End If
    timPrevious = timNow
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '让焦点在picForm和mshAttrib两个控件之间改变
    
    If (Shift And vbCtrlMask) <> 0 Or (Shift And vbShiftMask) <> 0 Then
        If ActiveControl Is mshAttrib Then
            If picForm.Enabled = True Then
                picForm.SetFocus
            End If
        End If
    Else
        If ActiveControl Is picForm Then
            If fraAttrib.Visible = True And mshAttrib.Enabled = True Then
                mshAttrib.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    Call 权限控制
    mnuViewList.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewList状态", "True") <> "False"
    mnuViewAttrib.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewAttrib状态", "True") <> "False"
    fraList.Visible = mnuViewList.Checked
    picSplitLeft.Visible = mnuViewList.Checked
    fraAttrib.Visible = mnuViewAttrib.Checked
    picSplitRight.Visible = mnuViewAttrib.Checked
    
    '其它初始化操作
    Set mcolSelect = New Collection
    Set mcolBill = New Elements
    mlngRow = -1
    mlng数量 = 0
    mbln新增 = False
    Call LoadAdjustControl
    mblnLoad = True
End Sub

Private Sub LoadAdjustControl()
    Dim lngIndex As Long
    
    '装入调整按钮
    For lngIndex = 0 To 7
        If lngIndex > 0 Then
            Load fraAdjust(lngIndex)
            Load shpAdjust(lngIndex)
        End If
        Set fraAdjust(lngIndex).Container = picForm
        Set shpAdjust(lngIndex).Container = fraAdjust(lngIndex)
        shpAdjust(lngIndex).Left = 15
        shpAdjust(lngIndex).Top = 15
    Next
    fraAdjust(0).MousePointer = vbSizeNWSE '左上角
    fraAdjust(1).MousePointer = vbSizeNS '上边
    fraAdjust(2).MousePointer = vbSizeNESW '右上角
    fraAdjust(3).MousePointer = vbSizeWE '右边
    fraAdjust(4).MousePointer = vbSizeNWSE '右下角
    fraAdjust(5).MousePointer = vbSizeNS '下角
    fraAdjust(6).MousePointer = vbSizeNESW '左下角
    fraAdjust(7).MousePointer = vbSizeWE '左角
    
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    Dim sngWidth As Single, sngHeight As Single, sngTemp As Single
    
    If Me.WindowState = 1 Then Exit Sub '最小化就不处理
    On Error Resume Next
    
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '设置各个控件的宽度
    If fraList.Width = 0 Then fraList.Width = 300
    sngWidth = ScaleWidth - IIf(fraList.Visible = False, 0, fraList.Width + picSplitLeft.Width) _
                     - IIf(fraAttrib.Visible = False, 0, fraAttrib.Width + picSplitRight.Width)
    If sngWidth < 0 Then sngWidth = 0
    If fraList.Visible = True Then
        If fraAttrib.Visible = True Then
            sngTemp = ScaleWidth - sngWidth - picSplitLeft.Width * 2 - fraAttrib.Width
            fraList.Width = IIf(sngTemp < 0, 0, sngTemp)
            sngTemp = ScaleWidth - sngWidth - picSplitLeft.Width * 2 - fraList.Width
            fraAttrib.Width = IIf(sngTemp < 0, 0, sngTemp)
        Else
            sngTemp = ScaleWidth - sngWidth - picSplitLeft.Width
            fraList.Width = IIf(sngTemp < 0, 0, sngTemp)
        End If
    Else
        If fraAttrib.Visible = True Then
            sngTemp = ScaleWidth - sngWidth - picSplitLeft.Width
            fraAttrib.Width = IIf(sngTemp < 0, 0, sngTemp)
        End If
    End If
    sngWidth = ScaleWidth - IIf(fraList.Visible = False, 0, fraList.Width + picSplitLeft.Width) _
                     - IIf(fraAttrib.Visible = False, 0, fraAttrib.Width + picSplitRight.Width)
    If sngWidth < 0 Then sngWidth = 0
    sngHeight = sngBottom - sngTop
    If sngHeight < 0 Then sngHeight = 0
    
    '计算各个控件的位置
    fraList.Left = ScaleLeft
    fraList.Top = sngTop
    fraList.Height = sngBottom - sngTop
    
    lblCaption(0).Top = 120
    lblCaption(0).Left = 60
    lblCaption(0).Width = fraList.Width
    picClose(0).Top = 60
    picClose(0).Left = fraList.Width - picClose(0).Width - 60
    lvwMain.Left = 60
    lvwMain.Top = picClose(0).Top + picClose(0).Height + 60
    lvwMain.Width = fraList.Width - 120
    lvwMain.Height = fraList.Height - lvwMain.Top - 60
    
    picSplitLeft.Top = fraList.Top
    picSplitLeft.Height = fraList.Height
    picSplitLeft.Left = fraList.Left + fraList.Width
    
    fraAttrib.Left = ScaleWidth - fraAttrib.Width
    fraAttrib.Top = sngTop
    fraAttrib.Height = sngBottom - sngTop
    
    lblCaption(1).Top = 120
    lblCaption(1).Left = 60
    lblCaption(1).Width = fraAttrib.Width
    picClose(1).Top = 60
    picClose(1).Left = fraAttrib.Width - picClose(1).Width - 60
    cmbControl.Left = 60
    cmbControl.Top = picClose(1).Top + picClose(1).Height + 60
    cmbControl.Width = fraAttrib.Width - 120
    mshAttrib.Left = 60
    mshAttrib.Top = cmbControl.Top + cmbControl.Height + 30
    mshAttrib.Width = fraAttrib.Width - 120
    mshAttrib.Height = fraAttrib.Height - mshAttrib.Top - 60
    If mshAttrib.Width > 4300 Then
        mshAttrib.ColWidth(0) = 2000
        mshAttrib.ColWidth(1) = mshAttrib.Width - 2000
    Else
        mshAttrib.ColWidth(0) = mshAttrib.Width / 2
        mshAttrib.ColWidth(1) = mshAttrib.Width / 2
    End If
    
    picSplitRight.Top = fraAttrib.Top
    picSplitRight.Height = fraAttrib.Height
    picSplitRight.Left = fraAttrib.Left - picSplitRight.Width
    
    '这里是客户区的各个控件的位置，包括滚动条
    HScroll.Left = IIf(picSplitLeft.Visible = True, picSplitLeft.Left + picSplitLeft.Width, ScaleLeft)
    VScroll.Top = sngTop
    
    sngTemp = picForm.Width + 600 - sngWidth
    If sngTemp > 0 Then
        HScroll.Visible = True
        HScroll.Min = 0
        HScroll.Max = sngTemp
        If picForm.Left > HScroll.Left + mlngWidthAdj Then
            HScroll.Value = 0
        ElseIf picForm.Left < HScroll.Left + mlngWidthAdj - sngTemp Then
            HScroll.Value = sngTemp
        Else
            HScroll.Value = HScroll.Left + mlngWidthAdj - picForm.Left
        End If
        HScroll.SmallChange = 240
        If (HScroll.Max - HScroll.Min) / 5 < 1000 Then
            HScroll.LargeChange = 1000
        Else
            HScroll.LargeChange = HScroll.Max - HScroll.Min
        End If
        picForm.Left = HScroll.Left + mlngWidthAdj - HScroll.Value
    Else
        HScroll.Visible = False
        picForm.Left = HScroll.Left + mlngWidthAdj
    End If
    
    sngTemp = picForm.Height + 600 - sngHeight
    If sngTemp > 0 Then
        VScroll.Visible = True
        VScroll.Min = 0
        VScroll.Max = sngTemp
        If picForm.Top > VScroll.Top + mlngWidthAdj Then
            VScroll.Value = 0
        ElseIf picForm.Top < VScroll.Top + mlngWidthAdj - sngTemp Then
            VScroll.Value = sngTemp
        Else
            VScroll.Value = VScroll.Top + mlngWidthAdj - picForm.Top
        End If
        VScroll.SmallChange = 240
        If (VScroll.Max - VScroll.Min) / 5 < 1000 Then
            VScroll.LargeChange = 1000
        Else
            VScroll.LargeChange = VScroll.Max - VScroll.Min
        End If
        picForm.Top = VScroll.Top + mlngWidthAdj - VScroll.Value
    Else
        VScroll.Visible = False
        picForm.Top = VScroll.Top + mlngWidthAdj
    End If
    fraCorner.Visible = HScroll.Visible And VScroll.Visible
    
    fraCorner.Left = IIf(picSplitRight.Visible = True, picSplitRight.Left, ScaleWidth) - fraCorner.Width
    fraCorner.Top = sngBottom - fraCorner.Height
    HScroll.Width = IIf(picSplitRight.Visible = True, picSplitRight.Left, ScaleWidth) _
        - IIf(fraCorner.Visible = True, VScroll.Width, 0) - HScroll.Left
    VScroll.Height = sngBottom - IIf(fraCorner.Visible = True, HScroll.Height, 0) - VScroll.Top
    HScroll.Top = sngBottom - HScroll.Height
    VScroll.Left = IIf(picSplitRight.Visible = True, picSplitRight.Left, ScaleWidth) - VScroll.Width
    
    If mcolSelect.Count = 0 Then
        Call SetAttach(picForm, Array(-1, -1, -1, 3, 4, 5, -1, -1), fraAdjust)
    End If
    Call ShowCmdEdit
    Me.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        Select Case MsgBox("当前记帐单已被修改，是否保存？", vbYesNoCancel Or vbQuestion Or vbDefaultButton3, gstrSysName)
            Case vbYes
                If SaveBill() = False Then
                    MsgBox "保存失败，不能退出。", vbExclamation, gstrSysName
                    Cancel = 1
                End If
            Case vbNo
            Case vbCancel
                Cancel = 1
        End Select
    End If
    
    If Cancel = 1 Then Exit Sub
    
    '下面的语句在真正退出时才执行
    mstrKey = ""
    Set mcolBill = Nothing
    Set mcolSelect = Nothing
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewAttrib状态", mnuViewAttrib.Checked
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewList状态", mnuViewList.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub lbl_DblClick(Index As Integer)
    Call SelectSame(lbl(Index))
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mlngColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mlngColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mlngColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Static timPrevious As Single
    Dim timNow As Single
    
    timNow = Timer
    If chk(Index).Value = 1 Then
        chk(Index).Value = 0
    Else
        '不这样处理的话每单击一下都要触发两次该事件
        Exit Sub
    End If
    If timPrevious <> 0 Then
       
       If timNow - timPrevious < 0.5 Then
            timPrevious = 0
            Call SelectSame(chk(Index))
            Exit Sub
       End If
    End If
    timPrevious = timNow
End Sub

Private Sub fraAdjust_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuFormatLock.Checked = True Then Exit Sub
    If Button = 1 Then
        mblnDown = True
        msngX = X
        msngY = Y
    End If
End Sub

Private Sub fraAdjust_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
End Sub

Private Sub fraAdjust_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'      0   1   2
'      7       3
'      6   5   4
    Dim lngLeft As Long, lngTop As Long
    
    If Button <> 1 Or mblnDown = False Or mnuFormatLock.Checked = True Then Exit Sub
    
    lngLeft = fraAdjust(Index).Left + X - msngX
    lngTop = fraAdjust(Index).Top + Y - msngY
    
    Select Case Index
        Case 0  '左上方
            If TypeName(mctlSelect) <> "ComboBox" Then
                If lngLeft < fraAdjust(4).Left - mlngWidthAdj * 2 Then
                    mctlSelect.Left = lngLeft + mlngWidthAdj
                    mctlSelect.Width = fraAdjust(4).Left - mctlSelect.Left
                End If
                If lngTop < fraAdjust(4).Top - mlngWidthAdj * 2 Then
                    mctlSelect.Height = fraAdjust(4).Top - (lngTop + mlngWidthAdj)
                    mctlSelect.Top = fraAdjust(4).Top - mctlSelect.Height '因为很多控件都有最小高度
                End If
            Else
                mctlSelect.Left = mctlSelect.Left + X - msngX
                mctlSelect.Top = mctlSelect.Top + Y - msngY
            End If
        Case 1  '正上方
            If TypeName(mctlSelect) <> "ComboBox" Then
                If lngTop < fraAdjust(4).Top - mlngWidthAdj * 2 Then
                    mctlSelect.Height = fraAdjust(4).Top - (lngTop + mlngWidthAdj)
                    mctlSelect.Top = fraAdjust(4).Top - mctlSelect.Height '因为很多控件都有最小高度
                End If
            Else
                mctlSelect.Top = mctlSelect.Top + Y - msngY
            End If
        Case 2  '右上方
            If TypeName(mctlSelect) <> "ComboBox" Then
                If lngLeft > mctlSelect.Left + mlngWidthAdj Then
                    mctlSelect.Width = lngLeft - mctlSelect.Left
                End If
                If lngTop < fraAdjust(4).Top - mlngWidthAdj * 2 Then
                    mctlSelect.Height = fraAdjust(4).Top - (lngTop + mlngWidthAdj)
                    mctlSelect.Top = fraAdjust(4).Top - mctlSelect.Height '因为很多控件都有最小高度
                End If
            Else
                mctlSelect.Left = mctlSelect.Left + X - msngX
                mctlSelect.Top = mctlSelect.Top + Y - msngY
            End If
        Case 3  '右方
            If mcolSelect.Count = 1 Then
                If lngLeft > mctlSelect.Left + mlngWidthAdj Then
                    mctlSelect.Width = lngLeft - mctlSelect.Left
                End If
            Else
                If lngLeft > picForm.Left + mlngWidthAdj Then
                    picForm.Width = lngLeft - picForm.Left
                End If
                Call Form_Resize
            End If
        Case 4  '右下方
            If mcolSelect.Count = 1 Then
                If TypeName(mctlSelect) <> "ComboBox" Then
                    If lngLeft > mctlSelect.Left + mlngWidthAdj Then
                        mctlSelect.Width = lngLeft - mctlSelect.Left
                    End If
                    If lngTop > mctlSelect.Top + mlngWidthAdj Then
                        mctlSelect.Height = lngTop - mctlSelect.Top
                    End If
                Else
                    mctlSelect.Left = mctlSelect.Left + X - msngX
                    mctlSelect.Top = mctlSelect.Top + Y - msngY
                End If
            Else
                If lngLeft > picForm.Left + mlngWidthAdj Then
                    picForm.Width = lngLeft - picForm.Left
                End If
                If lngTop > picForm.Top + mlngWidthAdj Then
                    picForm.Height = lngTop - picForm.Top
                End If
                Call Form_Resize
            End If
        Case 5  '正下方
            If mcolSelect.Count = 1 Then
                If TypeName(mctlSelect) <> "ComboBox" Then
                    If lngTop > mctlSelect.Top + mlngWidthAdj Then
                        mctlSelect.Height = lngTop - mctlSelect.Top
                    End If
                Else
                    mctlSelect.Top = mctlSelect.Top + Y - msngY
                End If
            Else
                If lngTop > picForm.Top + mlngWidthAdj Then
                    picForm.Height = lngTop - picForm.Top
                End If
                Call Form_Resize
            End If
        Case 6  '左下方
            If TypeName(mctlSelect) <> "ComboBox" Then
                If lngLeft < fraAdjust(4).Left - mlngWidthAdj * 2 Then
                    mctlSelect.Left = lngLeft + mlngWidthAdj
                    mctlSelect.Width = fraAdjust(4).Left - mctlSelect.Left
                End If
                If lngTop > mctlSelect.Top + mlngWidthAdj Then
                    mctlSelect.Height = lngTop - mctlSelect.Top
                End If
            Else
                mctlSelect.Left = mctlSelect.Left + X - msngX
                mctlSelect.Top = mctlSelect.Top + Y - msngY
            End If
        Case 7  '左方
            If lngLeft < fraAdjust(4).Left - mlngWidthAdj * 2 Then
                mctlSelect.Left = lngLeft + mlngWidthAdj
                mctlSelect.Width = fraAdjust(4).Left - mctlSelect.Left
            End If
    End Select
    
    If mcolSelect.Count = 1 Then
        Call SetAttach(mctlSelect, Array(0, 1, 2, 3, 4, 5, 6, 7), fraAdjust)
    End If
    mblnChange = True
    Call RefreshPosAttrib
End Sub

Private Sub RefreshPosAttrib()
'功能：更新属性框的有关位置的显示
    Dim lngRow As Long
    
    With mshAttrib
        For lngRow = 0 To .Rows - 1
            Select Case .TextMatrix(lngRow, 0)
                Case "左边距"
                    .TextMatrix(lngRow, 1) = mctlSelect.Left
                Case "顶边距"
                    .TextMatrix(lngRow, 1) = mctlSelect.Width
                Case "宽度"
                    If mcolSelect.Count = 0 Then
                        .TextMatrix(lngRow, 1) = picForm.Width
                    Else
                        .TextMatrix(lngRow, 1) = mctlSelect.Width
                    End If
                Case "高度"
                    If mcolSelect.Count = 0 Then
                        .TextMatrix(lngRow, 1) = picForm.Height
                    Else
                        .TextMatrix(lngRow, 1) = mctlSelect.Height
                    End If
            End Select
        Next
    End With
End Sub

Private Sub mnuEditElements_Click()
    Dim blnReturn As Boolean
    Dim strControl As String
    Dim lngCount As Long, varTemp As Variant
    Dim lngIndex As Long
    
    If mlng数量 = 0 Then
        blnReturn = frmElements.ModifyElement(mcolBill, mlng数量, True)
        If blnReturn = True Then
            '初始化各个控件的TagIndex
            strControl = "NO;销;姓名;性别;年龄;费别;床号;开单部门;病人ID;标识号;入院次数;病人病区;病人科室;"
            For lngCount = 1 To mlng数量
                strControl = strControl & "收费类别_" & lngCount & ";"
                strControl = strControl & "收费细目_" & lngCount & ";"
                strControl = strControl & "细目选择_" & lngCount & ";"
                strControl = strControl & "计算单位_" & lngCount & ";"
                strControl = strControl & "数次_" & lngCount & ";"
                strControl = strControl & "标准单价_" & lngCount & ";"
                strControl = strControl & "应收金额_" & lngCount & ";"
                strControl = strControl & "实收金额_" & lngCount & ";"
                strControl = strControl & "执行部门_" & lngCount & ";"
                strControl = strControl & "附加标志_" & lngCount & ";"
            Next
            strControl = strControl & "加班标志;婴儿费;开单人;发生时间;应收合计;实收合计;确定;取消"
        End If
        varTemp = Split(strControl, ";")
        lngIndex = 0
        For lngCount = LBound(varTemp) To UBound(varTemp)
            If mcolBill(varTemp(lngCount)).Visible = True Then
                lngIndex = lngIndex + 1
                Call SetTabIndex(varTemp(lngCount), lngIndex)
            End If
        Next
    Else
        blnReturn = frmElements.ModifyElement(mcolBill, mlng数量)
    End If
    
    If blnReturn = False Then Exit Sub
    '更新当前系统
    mblnChange = True
    Dim objTemp As Element
    For Each objTemp In mcolBill
        objTemp.Control.Visible = objTemp.Visible '可见性
    Next
    Call LoadControlList
End Sub

Private Function NewBill() As Boolean
    If frmElements.ModifyElement(mcolBill, mlng数量, True) = False Then Exit Function
    
    NewBill = True
End Function

Private Sub mnuEditText_Click()
    mlngMoveReason = 1 '表示画出一条文本标签
    picForm.MousePointer = vbCrosshair
    stbThis.Panels(2).Text = "在记帐单据上画出文本框的位置"
End Sub

Private Sub AddText(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
'在窗口的正中央增加一个文本标签
    Dim i As Long
    
    If lngHeight < 120 Or lngWidth < 120 Then Exit Sub
    
    Load lbl(lbl.UBound + 1)
    Set mctlSelect = lbl(lbl.UBound)
    
    mctlSelect.Caption = "标签"
    mctlSelect.Left = lngLeft
    mctlSelect.Top = lngTop
    mctlSelect.Width = lngWidth
    mctlSelect.Height = lngHeight
    mctlSelect.Visible = True
    '新增控件都与主窗口字体相同
    SetFont mctlSelect, picForm
    mcolBill.Add "标签_" & mctlSelect.Index, mctlSelect, , True
    
    mblnChange = True
    '更新选择集合
    cmbControl.AddItem "标签_" & mctlSelect.Index & "(" & mctlSelect.Caption & ")"
    cmbControl.ListIndex = cmbControl.NewIndex
    picForm.SetFocus
End Sub

Private Sub mnuEditLine_Click()
    mlngMoveReason = 2 '表示画出一条线
    picForm.MousePointer = vbCrosshair
    stbThis.Panels(2).Text = "在记帐单据上画出线条的位置"
End Sub

Private Sub AddLine(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
'在窗口的正中央增加一个文本标签
    Dim i As Long
    
    If lngHeight + lngWidth < 120 Then Exit Sub
    
    Load lbl(lbl.UBound + 1)
    Set mctlSelect = lbl(lbl.UBound)
    
    mctlSelect.Caption = ""
    mctlSelect.BackStyle = 1
    mctlSelect.Appearance = 0
    mctlSelect.BorderStyle = 0
    mctlSelect.BackColor = 0
    mctlSelect.Left = lngLeft
    mctlSelect.Top = lngTop
    mctlSelect.Width = lngWidth
    mctlSelect.Height = lngHeight
    mctlSelect.Visible = True
    '新增控件都与主窗口字体相同
    SetFont mctlSelect, picForm
    mcolBill.Add "标签_" & mctlSelect.Index, mctlSelect, , True
    
    mblnChange = True
    '更新选择集合
    cmbControl.AddItem "标签_" & mctlSelect.Index & "(" & mctlSelect.Caption & ")"
    cmbControl.ListIndex = cmbControl.NewIndex
    picForm.SetFocus
End Sub

Private Sub mnuEditCopy_Click()
'在窗口的正中央增加一个文本标签
    Dim lngCount As Long, i As Long
    Dim ctlCopy As Label
    Dim ctlSource As Control
    Dim colTemp As New Collection
    
    If txtEdit.Visible = True Then Exit Sub
    If mcolSelect.Count < 1 Then Exit Sub
    For Each ctlSource In mcolSelect
        If TypeName(ctlSource) = "Label" Then
            '计算加一
            lngCount = lngCount + 1
        End If
    Next
    '选中控件中没有标签
    If lngCount = 0 Then Exit Sub
    mblnChange = True
    For Each ctlSource In mcolSelect
        If TypeName(ctlSource) = "Label" Then
            
            Load lbl(lbl.UBound + 1)
            Set ctlCopy = lbl(lbl.UBound)
            
            ctlCopy.Caption = ctlSource.Caption
            ctlCopy.BackStyle = ctlSource.BackStyle
            ctlCopy.Appearance = ctlSource.Appearance
            ctlCopy.BorderStyle = ctlSource.BorderStyle
            ctlCopy.BackColor = ctlSource.BackColor
            ctlCopy.ForeColor = ctlSource.ForeColor
            ctlCopy.Width = ctlSource.Width
            ctlCopy.Height = ctlSource.Height
            
            SetFont ctlCopy, ctlSource
            
            If ctlCopy.Width = 15 And lngCount = 1 Then
                '竖线
                ctlCopy.Left = ctlSource.Left + 180
                ctlCopy.Top = ctlSource.Top
            ElseIf ctlCopy.Height = 15 And lngCount = 1 Then
                ctlCopy.Left = ctlSource.Left
                ctlCopy.Top = ctlSource.Top + 180
            Else
                ctlCopy.Left = ctlSource.Left + 30
                ctlCopy.Top = ctlSource.Top + 30
            End If
            ctlCopy.Visible = True
            '
            mcolBill.Add "标签_" & ctlCopy.Index, ctlCopy, , True
            colTemp.Add ctlCopy
            cmbControl.AddItem "标签_" & ctlCopy.Index & "(" & ctlCopy.Caption & ")"
        End If
    Next
    Set mctlSelect = ctlCopy '最后一个
    '更新选择集合
    For i = 1 To mcolSelect.Count
        mcolSelect.Remove 1
    Next
    For Each ctlSource In colTemp
        mcolSelect.Add ctlSource, ctlSource.Name & ctlSource.Index
    Next
    
    If lngCount = 1 Then
        '只有一个，刷新方式不同
        cmbControl.ListIndex = cmbControl.NewIndex
    Else
        '有多个，刷新又不同
        If cmbControl.ListIndex = -1 Then
            Call cmbControl_Click
        Else
            cmbControl.ListIndex = -1
        End If
    End If
    picForm.SetFocus
End Sub

Private Sub mnuEditRemove_Click()
'删除所选控件组中含有的标签
    Dim objElement As Element
    Dim ctlTemp    As Control
    Dim lngCount   As Long
    
    If mcolSelect.Count = 0 Then Exit Sub
    If MsgBox("是否要删除所选对象中的所有标签？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '首先清除已选择列表
    For lngCount = mcolSelect.Count To 1 Step -1
        Set ctlTemp = mcolSelect(lngCount)
        
        If TypeName(ctlTemp) = "Label" Then
            mcolSelect.Remove lngCount
            mcolBill.Remove ctlTemp.Tag
            
            If ctlTemp.Index = 0 Then
                ctlTemp.Visible = False
            Else
                Unload ctlTemp
            End If
        End If
    Next
    '更新列表
    cmbControl.Clear
    cmbControl.AddItem "记帐单"
    For Each objElement In mcolBill
        If objElement.Visible = True Then
            '只允许加可见的控件
            If TypeName(objElement.Control) = "Label" Then
                cmbControl.AddItem objElement.Key & "(" & objElement.Control.Caption & ")"
            Else
                cmbControl.AddItem objElement.Key
            End If
        End If
    Next
    
    If mcolSelect.Count = 1 Then
        Set mctlSelect = mcolSelect(1)
        For lngCount = 0 To cmbControl.ListCount - 1
            If GetFore(cmbControl.List(lngCount)) = mctlSelect.Tag Then
                cmbControl.ListIndex = lngCount
                Exit Sub
            End If
        Next
    ElseIf mcolSelect.Count > 1 Then
        Set mctlSelect = mcolSelect(1)
        '什么属性也不显示
            '以前就是选中各个控件，再赋值也就不会激活事件
            '所以只有手工调用
        Call HideAttach
        Call ShowAttach
        Call ShowAttrib
    Else
        cmbControl.ListIndex = GetIndexOfBill()
    End If
    
    mblnChange = True
End Sub

Private Sub mnuEditSelAll_Click()
    Dim objElement As Element
    Dim lngCount   As Long
    
    '首先清除已选择列表
    For lngCount = 1 To mcolSelect.Count
        mcolSelect.Remove 1
    Next
    
    '再增加所有有控件
    For Each objElement In mcolBill
        If objElement.Visible = True Then
            mcolSelect.Add objElement.Control, objElement.Control.Name & objElement.Control.Index
        End If
    Next
    
    If mcolSelect.Count = 1 Then
        Set mctlSelect = mcolSelect(1)
        For lngCount = 0 To cmbControl.ListCount - 1
            If GetFore(cmbControl.List(lngCount)) = mctlSelect.Tag Then
                cmbControl.ListIndex = lngCount
                Exit Sub
            End If
        Next
    ElseIf mcolSelect.Count > 1 Then
        Set mctlSelect = mcolSelect(1)
        '什么属性也不显示
        If cmbControl.ListIndex = -1 Then
            '以前就是选中各个控件，再赋值也就不会激活事件
            '所以只有手工调用
            Call HideAttach
            Call ShowAttach
            Call ShowAttrib
        Else
            cmbControl.ListIndex = -1
        End If
    Else
        cmbControl.ListIndex = GetIndexOfBill()
    End If
    
End Sub

Private Sub mnuFileExp_Click()
    Dim strFile As String
    Dim lngFile As Long
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    On Error Resume Next
    cdlFile.Filter = "专项记帐单 (*.ZLB)|*.ZLB"
    cdlFile.CancelError = True
    cdlFile.Flags = cdlOFNOverwritePrompt
    cdlFile.ShowSave
    If Err <> 0 Then
        Err.Clear
        Exit Sub
    End If
    strFile = cdlFile.FileName
    
    On Error GoTo errHandle
    
    MousePointer = vbHourglass
    lngFile = FreeFile
    Open strFile For Output Access Write As lngFile
    
    On Error GoTo errHandle
    
    '产生单据头
    strSQL = "select ID,编号,名称,收费项目数,适用范围,宽度,高度,背景色,字体 from 收费记帐单 where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    
    strSQL = "zl_收费记帐单_insert([ID],'[编号]','[名称]'," & _
        GetValue(rsTmp("收费项目数")) & "," & GetValue(rsTmp("适用范围")) & "," & GetValue(rsTmp("宽度")) & "," & _
        GetValue(rsTmp("高度")) & "," & GetValue(rsTmp("背景色")) & "," & GetValue(rsTmp("字体")) & ")"
    Print #lngFile, strSQL
    
    If rsTmp.State = adStateOpen Then rsTmp.Close
    
    '产生单据体
    strSQL = "select 记帐ID,对应字段,序号,类型,定义值,顺序号,左边,顶边,宽度,高度,字体,前景色,背景色,是否显示,外形,边框线,透明 from 收费记帐单定义 where 记帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 2)))
    Do Until rsTmp.EOF
        strSQL = "zl_收费记帐单定义_insert([ID]," & _
            GetValue(rsTmp("对应字段")) & "," & GetValue(rsTmp("序号")) & "," & GetValue(rsTmp("类型")) & "," & _
            GetValue(rsTmp("顺序号")) & "," & GetValue(rsTmp("左边")) & "," & GetValue(rsTmp("顶边")) & "," & _
            GetValue(rsTmp("宽度")) & "," & GetValue(rsTmp("高度")) & "," & GetValue(rsTmp("字体")) & "," & _
            GetValue(rsTmp("是否显示")) & "," & _
            IIf(rsTmp("对应字段") = "开单部门" Or rsTmp("对应字段") = "收费类别" Or rsTmp("对应字段") = "收费细目", IIf(rsTmp("对应字段") = "开单部门" Or rsTmp("对应字段") = "收费类别", "null", "0"), GetValue(rsTmp("定义值"))) & "," & _
            GetValue(rsTmp("前景色")) & "," & GetValue(rsTmp("背景色")) & "," & GetValue(rsTmp("外形")) & "," & _
            GetValue(rsTmp("边框线")) & "," & GetValue(rsTmp("透明")) & ")"
        Print #lngFile, strSQL
        
        rsTmp.MoveNext
    Loop
    
    '及时关闭文件
    Close lngFile
    MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Sub

Private Sub mnuFileImp_Click()
    Dim strFile As String
    Dim lngFile As Long
    Dim strID As String, str编号 As String, str名称 As String, strSQL As String
    
    On Error Resume Next
    cdlFile.Filter = "专项记帐单 (*.ZLB)|*.ZLB"
    cdlFile.CancelError = True
    cdlFile.Flags = cdlOFNFileMustExist
    cdlFile.ShowOpen
    If Err <> 0 Then
        Err.Clear
        Exit Sub
    End If
    strFile = cdlFile.FileName
    
    If frmSaveAs.另存模板(strID, str编号, str名称, False) = False Then
        Exit Sub
    End If
    DoEvents
    
    On Error GoTo errHandle
    
    MousePointer = vbHourglass
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    
    
    
    '产生单据体
    Do Until EOF(lngFile)
        Line Input #lngFile, strSQL
        strSQL = Replace(strSQL, "[ID]", strID)
        strSQL = Replace(strSQL, "[编号]", str编号)
        strSQL = Replace(strSQL, "[名称]", str名称)
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Loop
    
    '及时关闭文件
    Close lngFile
    MousePointer = vbDefault
    
    Call FillList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Sub

Private Sub mnuFileReload_Click()
    Call FillBill
End Sub

Private Sub mnuFileSave_Click()
    Call SaveBill
End Sub

Private Function SaveBill() As Boolean
'保存当前的记帐单
    Dim lngID As Long
    Dim strType As String
    Dim objElement As Element
    Dim ctlTemp As Control
    Dim lngPos As Long
    Dim strTemp As String, lngTemp As Long, strFont As String, strSQL As String
    Dim lst As ListItem
    
    On Error GoTo errHandle
    
    If IsValid = False Then Exit Function
    
    
    MousePointer = 11
    Me.stbThis.Panels(2).Text = "正在保存……"
    gcnOracle.BeginTrans
    
    '处理主记录
    With picForm.Font
        strFont = .Name & "|" & .Size & "|" & IIf(.Bold, "1", "0") & "|" & IIf(.Italic, "1", "0") & "|" & IIf(.Underline, "1", "0")
    End With
    If mbln新增 = True Then
        lngID = zlDatabase.GetNextId("收费记帐单")
        strSQL = "zl_收费记帐单_insert(" & lngID & ",'" & mstr编号 & "','" & mstr名称 & "'," & _
            mlng数量 & ",'" & mstr适用范围 & "'," & picForm.Width & "," & picForm.Height & _
            "," & picForm.BackColor & ",'" & strFont & "')"
    Else
        lngID = Mid(lvwMain.SelectedItem.Key, 2)
        strSQL = "zl_收费记帐单_update(" & lngID & ",'" & mstr编号 & "','" & mstr名称 & "'," & _
            mlng数量 & ",'" & mstr适用范围 & "'," & picForm.Width & "," & picForm.Height & _
            "," & picForm.BackColor & ",'" & strFont & "')"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '处理明细记录
    strSQL = "zl_收费记帐单定义_delete(" & lngID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    For Each objElement In mcolBill
        Set ctlTemp = objElement.Control
        strType = TypeName(ctlTemp)
        Select Case strType
            Case "CheckBox"
                strSQL = ",'" & ctlTemp.Caption & "'," & ctlTemp.ForeColor & "," & ctlTemp.BackColor & "," & _
                    ctlTemp.Appearance & ",0,0)"
            Case "ComboBox"
                strSQL = ",'" & objElement.Value & "'," & ctlTemp.ForeColor & "," & ctlTemp.BackColor & "," & _
                    ctlTemp.Appearance & ",0,0)"
            Case "CommandButton"
                strSQL = ",'" & ctlTemp.Caption & "',0,0,0,0,0)"
            Case "Label"
                strSQL = ",'" & ctlTemp.Caption & "'," & ctlTemp.ForeColor & "," & ctlTemp.BackColor & "," & _
                    ctlTemp.Appearance & "," & ctlTemp.BorderStyle & "," & ctlTemp.BackStyle & ")"
            Case "TextBox"
                strSQL = ",'" & objElement.Value & "'," & ctlTemp.ForeColor & "," & ctlTemp.BackColor & "," & _
                    ctlTemp.Appearance & "," & ctlTemp.BorderStyle & ",0)"
        End Select
        lngPos = InStr(objElement.Key, "_")
        If lngPos = 0 Then
            strTemp = objElement.Key
            lngTemp = 0
        Else
            strTemp = Mid(objElement.Key, 1, lngPos - 1)
            lngTemp = Val(Mid(objElement.Key, lngPos + 1))
        End If
        strSQL = "zl_收费记帐单定义_insert(" & lngID & ",'" & strTemp & "'," & IIf(lngPos = 0, "null", lngTemp) & ",'" & _
            strType & "'," & objElement.TabIndex & "," & ctlTemp.Left & "," & _
            ctlTemp.Top & "," & ctlTemp.Width & "," & ctlTemp.Height & ",'" & _
            ctlTemp.Font.Name & "|" & ctlTemp.Font.Size & "|" & IIf(ctlTemp.Font.Bold, "1", "0") & "|" & IIf(ctlTemp.Font.Italic, "1", "0") & "|" & IIf(ctlTemp.Font.Underline, "1", "0") & "'," & _
            IIf(objElement.Visible, "1", "0") & strSQL
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    Me.stbThis.Panels(2).Text = "保存完毕。"
    MousePointer = 0
    mblnChange = False
    SaveBill = True
    
    On Error Resume Next
    '更新主界面
    If mbln新增 = True Then
        Set lst = lvwMain.ListItems.Add(, "B" & lngID, mstr名称, "Bill", "Bill")
        lst.Selected = True
        mstrKey = lst.Key
    Else
        Set lst = lvwMain.SelectedItem
        lst.Text = mstr名称
    End If
    lst.SubItems(1) = mstr编号
    lst.SubItems(2) = IIf(Mid(mstr适用范围, 1, 1) = "1", "√", "")
    lst.SubItems(3) = IIf(Mid(mstr适用范围, 2, 1) = "1", "√", "")
    lst.SubItems(4) = IIf(Mid(mstr适用范围, 3, 1) = "1", "√", "")
    lst.SubItems(5) = IIf(Mid(mstr适用范围, 4, 1) = "1", "√", "")
    lst.Tag = picForm.Width & "," & picForm.Height & "," & mlng数量 & "," & picForm.BackColor & "," & strFont
    lst.EnsureVisible
    mbln新增 = False
    Exit Function
    
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.stbThis.Panels(2).Text = "保存失败！"
    MousePointer = 0
End Function

Private Function IsValid() As Boolean
'判断当前记帐单的合法性
    If mstr编号 = "" Then
        cmbControl.ListIndex = GetIndexOfBill
        MsgBox "记帐单的编号不能为空。", vbExclamation, gstrSysName
        mshAttrib.Row = 0: mshAttrib.Col = 1
        Call mshAttrib_EnterCell
        If fraAttrib.Visible = False Then
            Call mnuViewAttrib_Click
        End If
        mshAttrib.SetFocus
        Exit Function
    End If
    If mstr名称 = "" Then
        cmbControl.ListIndex = GetIndexOfBill
        MsgBox "记帐单的名称不能为空。", vbExclamation, gstrSysName
        mshAttrib.Row = 1: mshAttrib.Col = 1
        Call mshAttrib_EnterCell
        If fraAttrib.Visible = False Then
            Call mnuViewAttrib_Click
        End If
        mshAttrib.SetFocus
        Exit Function
    End If
    If mlng数量 = 0 Then
        MsgBox "请增加记帐单元素。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    IsValid = True
End Function

Private Sub mnuFileSaveAs_Click()
    Dim strID As String, str编码 As String, str名称 As String
    Dim lst As ListItem
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    strID = Mid(lvwMain.SelectedItem.Key, 2)
    
    frmSaveAs.另存模板 strID, str编码, str名称
    
    If str编码 = "" Then Exit Sub '取消了
    
    Set lst = lvwMain.ListItems.Add(, "B" & strID, str名称, "Bill", "Bill")
    
    With lvwMain.SelectedItem
        lst.SubItems(1) = str编码
        lst.SubItems(2) = .SubItems(2)
        lst.SubItems(3) = .SubItems(3)
        lst.SubItems(4) = .SubItems(4)
        lst.SubItems(5) = .SubItems(5)
        lst.Tag = .Tag
    End With
    
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub picForm_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyLeft
                ResizeControl -15, 1
            Case vbKeyRight
                ResizeControl 15, 1
            Case vbKeyDown
                ResizeControl 15, 2
            Case vbKeyUp
                ResizeControl -15, 2
        End Select
    ElseIf Shift = vbShiftMask Then
        Select Case KeyCode
            Case vbKeyLeft
                ResizeControl -30, 3
            Case vbKeyRight
                ResizeControl 30, 3
            Case vbKeyDown
                ResizeControl 30, 4
            Case vbKeyUp
                ResizeControl -30, 4
        End Select
    End If
End Sub

Private Sub picForm_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ShowAttach
End Sub

Private Sub cmb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyDown KeyCode, Shift
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyDown KeyCode, Shift
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyDown KeyCode, Shift
End Sub

Private Sub cmb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyUp KeyCode, Shift
End Sub

Private Sub cmd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyUp KeyCode, Shift
End Sub

Private Sub chk_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyUp KeyCode, Shift
End Sub

Private Sub picForm_Resize()
    picForm.Cls
    Dim r As RECT
    With picForm
        r.Left = (.ScaleLeft) / Screen.TwipsPerPixelX
        r.Top = (.ScaleTop) / Screen.TwipsPerPixelY
        r.Right = (.ScaleLeft + .ScaleWidth) / Screen.TwipsPerPixelX
        r.Bottom = (.ScaleTop + .ScaleHeight) / Screen.TwipsPerPixelY
        DrawEdge .hdc, r, EDGE_RAISED, BF_RECT
    End With

End Sub

Private Sub txt_DblClick(Index As Integer)
    Call SelectSame(txt(Index))
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    picForm_KeyDown KeyCode, Shift
End Sub

Private Sub cmb_GotFocus(Index As Integer)
    Dim p As POINTAPI
    
    
    '强制移开焦点，以便再次得到GotFocus
    If fraAttrib.Visible = True And mshAttrib.Enabled = True Then
        mshAttrib.SetFocus
    Else
        picForm.SetFocus
    End If
    
    mintIndex = Index
    If CtrlIsPress = True Then
        AddSelect cmb(Index)
    Else
        picCombo.Left = cmb(Index).Left
        picCombo.Top = cmb(Index).Top
        
        SetCapture picCombo.hwnd
        GetCursorPos p
        ScreenToClient picCombo.hwnd, p
        msngX = p.X * Screen.TwipsPerPixelX
        msngY = p.Y * Screen.TwipsPerPixelY
        
        ChangeSelectBefore cmb(Index)
    End If
    '处理双击
    Static timPrevious As Single
    Dim timNow As Single
    
    timNow = Timer
    If timPrevious <> 0 Then
       
       If timNow - timPrevious < 0.5 Then
            timPrevious = 0
            Call SelectSame(cmb(Index))
            Exit Sub
       End If
    End If
    timPrevious = timNow
End Sub

Private Sub picCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then Exit Sub '按了Ctrl键就不处理
    
    If Button = 1 Then
        If Not mctlSelect Is cmb(mintIndex) Then Exit Sub
        Call MoveControl(X, Y, True)
    End If
End Sub

Private Sub picCombo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) = 0 Then
        ChangeSelectAfter
    End If
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) = 0 And Button = 1 Then
        ChangeSelectBefore cmd(Index)
        msngX = X
        msngY = Y
    End If
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then Exit Sub '按了Ctrl键就不处理
    
    If Button = 1 Then
        If Not mctlSelect Is cmd(Index) Then Exit Sub
        Call MoveControl(X, Y)
    End If
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then
        AddSelect cmd(Index)
    ElseIf Button = 1 Then
        ChangeSelectAfter
    End If
End Sub

Private Sub chk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) = 0 And Button = 1 Then
        ChangeSelectBefore chk(Index)
        msngX = X
        msngY = Y
    End If
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then Exit Sub '按了Ctrl键就不处理
    
    If Button = 1 Then
        If Not mctlSelect Is chk(Index) Then Exit Sub
        Call MoveControl(X, Y)
    End If
End Sub

Private Sub chk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then
        AddSelect chk(Index)
    ElseIf Button = 1 Then
        ChangeSelectAfter
    End If
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) = 0 And Button = 1 Then
        ChangeSelectBefore lbl(Index)
        msngX = X
        msngY = Y
    End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then Exit Sub '按了Ctrl键就不处理
    
    If Button = 1 Then
        If Not mctlSelect Is lbl(Index) Then Exit Sub
        Call MoveControl(X, Y)
    End If
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then
        AddSelect lbl(Index)
    ElseIf Button = 1 Then
        ChangeSelectAfter
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) = 0 And Button = 1 Then
        ChangeSelectBefore txt(Index)
        msngX = X
        msngY = Y
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then Exit Sub '按了Ctrl键就不处理
    
    If Button = 1 Then
        If Not mctlSelect Is txt(Index) Then Exit Sub
        Call MoveControl(X, Y)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift And vbCtrlMask) <> 0 Then
        AddSelect txt(Index)
    ElseIf Button = 1 Then
        ChangeSelectAfter
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim str属性 As String
    Dim ctlObject As Control
    
    
    cdg.CancelError = True
    On Error GoTo errExit
    
    If mcolSelect.Count = 0 Then
        Set ctlObject = picForm
    Else
        Set ctlObject = mctlSelect
    End If
    
    str属性 = mshAttrib.TextMatrix(mshAttrib.Row, 0)
    
    With ctlObject
        If str属性 = "字体" Then
            cdg.Flags = cdlCFScreenFonts
            cdg.FontName = .FontName
            cdg.FontSize = .FontSize
            cdg.FontBold = .FontBold
            cdg.FontItalic = .FontItalic
            cdg.ShowFont
            
            If mcolSelect.Count = 0 Then
                .FontName = cdg.FontName
                .FontSize = cdg.FontSize
                .FontBold = cdg.FontBold
                .FontItalic = cdg.FontItalic
            Else
                For Each ctlObject In mcolSelect
                    ctlObject.FontName = cdg.FontName
                    ctlObject.FontSize = cdg.FontSize
                    ctlObject.FontBold = cdg.FontBold
                    ctlObject.FontItalic = cdg.FontItalic
                Next
            End If
            
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = cdg.FontName & "(" & cdg.FontSize & ")"
            Call ShowAttach '改变字体就可能改变其大小
        Else
            cdg.Flags = cdlCCFullOpen Or cdlCCRGBInit
            cdg.Color = IIf(str属性 = "字体色", .ForeColor, .BackColor)
            cdg.ShowColor
            If str属性 = "字体色" Then
                If mcolSelect.Count = 0 Then
                    .ForeColor = cdg.Color
                Else
                    For Each ctlObject In mcolSelect
                        ctlObject.ForeColor = cdg.Color
                    Next
                End If
            Else
                If mcolSelect.Count = 0 Then
                    .BackColor = cdg.Color
                Else
                    For Each ctlObject In mcolSelect
                        ctlObject.BackColor = cdg.Color
                    Next
                End If
            End If
            mshAttrib.Col = 1: mshAttrib.CellForeColor = cdg.Color
        End If
    End With
    mblnChange = True
errExit:
    mshAttrib.SetFocus
End Sub
    
Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = mstrKey Then Exit Sub
    mstrKey = Item.Key
    
    Call FillBill
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
'新增一张记帐单
    mbln新增 = True
    mnuFileDesign.Checked = True
    Toolbar1.Buttons("Design").Value = tbrPressed
    mnuFormatLock.Checked = True
    Toolbar1.Buttons("Lock").Value = tbrPressed
    stbThis.Panels(2).Text = "进入新增状态"
    Call SetMenu
    
    Call FillBill
End Sub

Private Sub mnuFileDesign_Click()
'设计选中记帐单
    If mnuFileDesign.Checked = True Then
        '退出设计状态
        If mblnChange = True Then
            Select Case MsgBox("当前记帐单已被修改，是否保存？", vbYesNoCancel Or vbQuestion Or vbDefaultButton3, gstrSysName)
                Case vbYes
                    If SaveBill() = False Then
                        MsgBox "保存失败，不能退出。", vbExclamation, gstrSysName
                        Toolbar1.Buttons("Design").Value = tbrPressed
                        Exit Sub
                        'Call FillBill
                    End If
                Case vbNo
                    mbln新增 = False
                    Call FillBill
                Case vbCancel
                    Toolbar1.Buttons("Design").Value = tbrPressed
                    Exit Sub
            End Select
        ElseIf mbln新增 = True Then
            '退出新增时的刷新
            mbln新增 = False
            Call FillBill
        End If
        Call HideAttach
    Else
        mnuFormatLock.Checked = True
        Toolbar1.Buttons("Lock").Value = tbrPressed
        stbThis.Panels(2).Text = "进入设计修改状态"
    End If
    mnuFileDesign.Checked = Not mnuFileDesign.Checked
    Toolbar1.Buttons("Design").Value = IIf(mnuFileDesign.Checked, tbrPressed, tbrUnpressed)
    '不管怎么样，都不是处于新增状态
    mbln新增 = False
    Call SetMenu
End Sub

Private Sub mnuFileErase_Click()
'删除选中记帐单
    Dim intIndex As Integer, strSQL As String
    
    On Error GoTo errHandle
    
    If MsgBox("你确认要删除名称为“" & lvwMain.SelectedItem.Text & "”的记帐单吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        strSQL = "zl_收费记帐单_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
        End With
        
        Call FillBill
        Call SetMenu
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFormatDoAlign_Click(Index As Integer)
    Dim ctlTemp As Control
    
    If mcolSelect.Count < 2 Then Exit Sub '没有对齐的必要
    
    Call HideAttach
    Select Case Index
        Case 0 '左对齐
            For Each ctlTemp In mcolSelect
                ctlTemp.Left = mctlSelect.Left
            Next
        Case 1 '横居中
            For Each ctlTemp In mcolSelect
                ctlTemp.Left = mctlSelect.Left + mctlSelect.Width / 2 - ctlTemp.Width / 2
            Next
        Case 2 '右对齐
            For Each ctlTemp In mcolSelect
                ctlTemp.Left = mctlSelect.Left + mctlSelect.Width - ctlTemp.Width
            Next
        Case 4 '上对齐
            For Each ctlTemp In mcolSelect
                ctlTemp.Top = mctlSelect.Top
            Next
        Case 5 '竖居中
            For Each ctlTemp In mcolSelect
                ctlTemp.Top = mctlSelect.Top + mctlSelect.Height / 2 - ctlTemp.Height / 2
            Next
        Case 6 '下对齐
            For Each ctlTemp In mcolSelect
                ctlTemp.Top = mctlSelect.Top + mctlSelect.Height - ctlTemp.Height
            Next
    End Select
    mblnChange = True
    Call ShowAttach
End Sub

Private Sub mnuFormatFormAlign_Click(Index As Integer)
    Dim ctlTemp As Control
    
    If mcolSelect.Count = 0 Then Exit Sub '没有居中的必要
    
    Call HideAttach
    If Index = 0 Then
        '水平居中
        For Each ctlTemp In mcolSelect
            ctlTemp.Left = (picForm.ScaleWidth - ctlTemp.Width) / 2
        Next
    Else
        '垂直居中
        For Each ctlTemp In mcolSelect
            ctlTemp.Top = (picForm.ScaleHeight - ctlTemp.Height) / 2
        Next
    End If
    mblnChange = True
    Call ShowAttach
    '如果只有一个控件，那需要更新其属性表内容
    Call ShowAttrib
End Sub

Private Sub mnuFormatSize_Click(Index As Integer)
    Dim ctlTemp As Control
    
    If mcolSelect.Count < 2 Then Exit Sub '没有比较大小的必要
    
    Call HideAttach
    Select Case Index
        Case 0 '相同宽度
            For Each ctlTemp In mcolSelect
                ctlTemp.Width = mctlSelect.Width
            Next
        Case 1 '相同高度
            For Each ctlTemp In mcolSelect
                If Not TypeName(ctlTemp) = "ComboBox" Then
                    ctlTemp.Height = mctlSelect.Height
                End If
            Next
        Case 2 '相同大小
            For Each ctlTemp In mcolSelect
                ctlTemp.Width = mctlSelect.Width
                If Not TypeName(ctlTemp) = "ComboBox" Then
                    ctlTemp.Height = mctlSelect.Height
                End If
            Next
    End Select
    mblnChange = True
    Call ShowAttach
End Sub

Private Sub mnuFormatVscSpace_Click(Index As Integer)
    Dim ctlArr() As Control
    Dim ctlTemp As Control
    Dim lngCount As Long, lngLoop As Long
    Dim lngTotal As Long, lngSpace As Long
    
    lngCount = mcolSelect.Count
    If lngCount < 2 Then Exit Sub '没有调整间距的必要
    '先把所有选中控件放到一个临时的数组中
    ReDim ctlArr(1 To lngCount)
    For lngLoop = 1 To lngCount
        Set ctlArr(lngLoop) = mcolSelect(lngLoop)
        lngTotal = lngTotal + ctlArr(lngLoop).Height '所有控件的高度之合
    Next
    '接着按顶边大小排序
    For lngCount = 1 To mcolSelect.Count - 1
        For lngLoop = lngCount + 1 To mcolSelect.Count
            If ctlArr(lngLoop).Top < ctlArr(lngCount).Top Then
                '如果发现有比它小的，并交换位置
                Set ctlTemp = ctlArr(lngLoop)
                Set ctlArr(lngLoop) = ctlArr(lngCount)
                Set ctlArr(lngCount) = ctlTemp
            End If
        Next
    Next
    
    '竖间距
    Call HideAttach
    Select Case Index
        Case 0 '相同
            lngSpace = ((ctlArr(mcolSelect.Count).Top + ctlArr(mcolSelect.Count).Height - ctlArr(1).Top) - lngTotal) / (mcolSelect.Count - 1)
            For lngLoop = 2 To mcolSelect.Count - 1
                ctlArr(lngLoop).Top = ctlArr(lngLoop - 1).Top + ctlArr(lngLoop - 1).Height + lngSpace
            Next
        Case 1 '增加
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Top = ctlArr(lngLoop).Top + 30 * (lngLoop - 1)
            Next
        Case 2 '减少
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Top = ctlArr(lngLoop).Top - 30 * (lngLoop - 1)
                If ctlArr(lngLoop).Top < ctlArr(lngLoop - 1).Top + 30 Then
                   ctlArr(lngLoop).Top = ctlArr(lngLoop - 1).Top + 30
                End If
            Next
        Case 3 '移除
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Top = ctlArr(lngLoop - 1).Top + ctlArr(lngLoop - 1).Height
            Next
    End Select
    mblnChange = True
    Call ShowAttach
End Sub

Private Sub mnuFormatHscSpace_Click(Index As Integer)
    Dim ctlArr() As Control
    Dim ctlTemp As Control
    Dim lngCount As Long, lngLoop As Long
    Dim lngTotal As Long, lngSpace As Long
    
    lngCount = mcolSelect.Count
    If lngCount < 2 Then Exit Sub '没有调整间距的必要
    '先把所有选中控件放到一个临时的数组中
    ReDim ctlArr(1 To lngCount)
    For lngLoop = 1 To lngCount
        Set ctlArr(lngLoop) = mcolSelect(lngLoop)
        lngTotal = lngTotal + ctlArr(lngLoop).Width '所有控件的高度之合
    Next
    '接着按左边大小排序
    For lngCount = 1 To mcolSelect.Count - 1
        For lngLoop = lngCount + 1 To mcolSelect.Count
            If ctlArr(lngLoop).Left < ctlArr(lngCount).Left Then
                '如果发现有比它小的，并交换位置
                Set ctlTemp = ctlArr(lngLoop)
                Set ctlArr(lngLoop) = ctlArr(lngCount)
                Set ctlArr(lngCount) = ctlTemp
            End If
        Next
    Next
    
    '横间距
    Call HideAttach
    Select Case Index
        Case 0 '相同
            lngSpace = ((ctlArr(mcolSelect.Count).Left + ctlArr(mcolSelect.Count).Width - ctlArr(1).Left) - lngTotal) / (mcolSelect.Count - 1)
            For lngLoop = 2 To mcolSelect.Count - 1
                ctlArr(lngLoop).Left = ctlArr(lngLoop - 1).Left + ctlArr(lngLoop - 1).Width + lngSpace
            Next
        Case 1 '增加
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Left = ctlArr(lngLoop).Left + 30 * (lngLoop - 1)
            Next
        Case 2 '减少
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Left = ctlArr(lngLoop).Left - 30 * (lngLoop - 1)
                If ctlArr(lngLoop).Left < ctlArr(lngLoop - 1).Left + 30 Then
                   ctlArr(lngLoop).Left = ctlArr(lngLoop - 1).Left + 30
                End If
            Next
        Case 3 '移除
            For lngLoop = 2 To mcolSelect.Count
                ctlArr(lngLoop).Left = ctlArr(lngLoop - 1).Left + ctlArr(lngLoop - 1).Width
            Next
    End Select
    mblnChange = True
    Call ShowAttach
End Sub

Private Sub mnuFormatLock_Click()
'锁定元素
    Dim blnEnable As Boolean
    
    blnEnable = Not mnuFormatLock.Checked
    mnuFormatLock.Checked = blnEnable
    Toolbar1.Buttons("Lock").Value = IIf(blnEnable = True, tbrPressed, tbrUnpressed)
    
    Call SetFormatMenu
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp "ZL9CustAcc", Me.hwnd, Me.Name
End Sub

Private Sub mnuViewAttrib_Click()
    mnuViewAttrib.Checked = Not mnuViewAttrib.Checked
    fraAttrib.Visible = mnuViewAttrib.Checked
    picSplitRight.Visible = mnuViewAttrib.Checked
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
End Sub

Private Sub mnuViewList_Click()
    mnuViewList.Checked = Not mnuViewList.Checked
    fraList.Visible = mnuViewList.Checked
    picSplitLeft.Visible = mnuViewList.Checked
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuHelpAbout_Click()
'    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    ShowAbout Me, "基础数据管理", "zl9BaseCode", App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub HScroll_Change()
    '通过按钮触发
    Call HScroll_Scroll
End Sub

Private Sub HScroll_Scroll()
    '通过拖动触发
    picForm.Left = HScroll.Left + mlngWidthAdj - HScroll.Value
    If mcolSelect.Count = 0 Then
        Call SetAttach(picForm, Array(-1, -1, -1, 3, 4, 5, -1, -1), fraAdjust)
    End If
End Sub

Private Sub picForm_Paint()
    Dim r As RECT
    With picForm
        r.Left = (.ScaleLeft) / Screen.TwipsPerPixelX
        r.Top = (.ScaleTop) / Screen.TwipsPerPixelY
        r.Right = (.ScaleLeft + .ScaleWidth) / Screen.TwipsPerPixelX
        r.Bottom = (.ScaleTop + .ScaleHeight) / Screen.TwipsPerPixelY
        DrawEdge .hdc, r, EDGE_RAISED, BF_RECT
    End With
    
End Sub

Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
End Sub

Private Sub txtEdit_LostFocus()
    '以防万一
    txtEdit.Visible = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        txtEdit.Visible = False
        mshAttrib.SetFocus
    End If
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If AssignValue = False Then
    '赋值未成功
        zlControl.TxtSelAll txtEdit
        txtEdit.SetFocus
        Exit Sub
    End If
    KeyAscii = 0
    txtEdit.Visible = False
    mshAttrib.SetFocus
    
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    Call AssignValue
    '不管赋值成功与否，都要使用输入框不可见
    txtEdit.Visible = False
    mshAttrib.SetFocus
End Sub

Private Function AssignValue() As Boolean
'功能：根据用户在属性框中输入的值，做相应的操作
    Dim strAttrib As String
    Dim lngTemp As Long
    
    strAttrib = mshAttrib.TextMatrix(mshAttrib.Row, 0)
    Select Case strAttrib
        Case "编号"
            If zlCommFun.StrIsValid(txtEdit.Text, 6, True) = False Then
                Exit Function
            End If
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = txtEdit.Text
            mstr编号 = txtEdit.Text
        Case "名称"
            If mcolSelect.Count = 0 Then
                '对记帐单本身
                If zlCommFun.StrIsValid(txtEdit.Text, 50, True) = False Then
                    Exit Function
                End If
                mstr名称 = txtEdit.Text
            Else
                '针对CheckBox、CommandBox控件
                If zlCommFun.StrIsValid(txtEdit.Text, 30, True) = False Then
                    Exit Function
                End If
                mctlSelect.Caption = txtEdit.Text
                Call ShowAttach '改变内容，就可能改变其宽度
            End If
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = txtEdit.Text
        Case "文本"
            '针对标签控件
            If zlCommFun.StrIsValid(txtEdit.Text, 30) = False Then
                Exit Function
            End If
            mctlSelect.Caption = txtEdit.Text
            cmbControl.List(cmbControl.ListIndex) = mctlSelect.Tag & "(" & mctlSelect.Caption & ")"
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = txtEdit.Text
            Call ShowAttach '改变内容，就可能改变其宽度
        Case "左边距", "顶边距", "宽度", "高度"
            If NumIsValid(txtEdit.Text) = False Then
                Exit Function
            End If
            If (strAttrib = "宽度" Or strAttrib = "高度") And Val(txtEdit.Text) < Screen.TwipsPerPixelX Then
                MsgBox strAttrib & "不小于" & Screen.TwipsPerPixelY & "。", vbExclamation, gstrSysName
                Exit Function
            End If
            
            lngTemp = Int(txtEdit.Text)
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = lngTemp
            If strAttrib = "左边距" Then
                mctlSelect.Left = lngTemp
            ElseIf strAttrib = "顶边距" Then
                mctlSelect.Top = lngTemp
            ElseIf strAttrib = "宽度" Then
                If mcolSelect.Count = 0 Then
                    picForm.Width = lngTemp
                Else
                    mctlSelect.Width = lngTemp
                End If
                '重新计算滚动条出现与否
                Call Form_Resize
            Else
                If mcolSelect.Count = 0 Then
                    picForm.Height = lngTemp
                Else
                    mctlSelect.Height = lngTemp
                End If
                '重新计算滚动条出现与否
                Call Form_Resize
            End If
            '改变方块的位置
            Call ShowAttach
        Case "顺序号"
            If Not IsNumeric(txtEdit.Text) Then
                MsgBox "请输入一个数值。", vbExclamation, gstrSysName
                Exit Function
            End If
            If Val(txtEdit.Text) < 1 Then
                MsgBox "请输入一个正数。", vbExclamation, gstrSysName
                Exit Function
            End If
            If Val(txtEdit.Text) > mcolBill.Count Then
                lngTemp = mcolBill.Count '这已经是最大的顺序号了
            Else
                lngTemp = Int(txtEdit.Text)
            End If
            mshAttrib.TextMatrix(mshAttrib.Row, 1) = lngTemp
            
            '更改其它控件的顺序
            
            If mcolBill(mctlSelect.Tag).TabIndex = lngTemp Then
                AssignValue = True
                Exit Function '没有改变
            End If
            Call SetTabIndex(mctlSelect.Tag, lngTemp)
    End Select
    AssignValue = True
    mblnChange = True
End Function

Private Sub SetTabIndex(ByVal strKey As String, ByVal lngIndex As Long)
    '更改其它控件的顺序
    Dim objTemp As Element
    Dim objSelect As Element
    Dim lngPre As Long
    Dim lngCurr As Long
    
    
    Set objSelect = mcolBill(strKey)
    lngPre = objSelect.TabIndex
    
    If lngPre = lngIndex Then
        Exit Sub      '没有改变
    End If
    For Each objTemp In mcolBill
        If objTemp Is objSelect Then
            objTemp.TabIndex = lngIndex
        Else
            lngCurr = objTemp.TabIndex
            If lngCurr > lngPre And lngCurr <= lngIndex Then
                '当前号码是改大了
                objTemp.TabIndex = lngCurr - 1
            ElseIf lngCurr < lngPre And lngCurr >= lngIndex Then
                '当前号码是改小了
                objTemp.TabIndex = lngCurr + 1
            End If
        End If
    Next
End Sub

Private Sub VScroll_Change()
    '通过按钮触发
    Call VScroll_Scroll
End Sub

Private Sub VScroll_Scroll()
    '通过拖动触发
    picForm.Top = VScroll.Top + mlngWidthAdj - VScroll.Value
    If mcolSelect.Count = 0 Then
        Call SetAttach(picForm, Array(-1, -1, -1, 3, 4, 5, -1, -1), fraAdjust)
    End If
End Sub

Private Sub mshAttrib_Scroll()
    Call ShowCmdEdit
End Sub

Private Sub mshAttrib_EnterCell()
'设置当前行的颜色
    Dim blnRedraw As Boolean
    Dim lngRow As Long
    
    With mshAttrib
        blnRedraw = .Redraw
        If .Rows = 1 Then
            mlngRow = 0
            .Col = 0
            .CellForeColor = &H80000008
            .CellBackColor = &H80000005
            .Col = 1
        Else
            If mlngRow = .Row Then
                '为保证不出错，再设置一次
                .Col = 0
                .CellForeColor = &H80000005
                .CellBackColor = &H8000000D
                .Col = 1
            Else
                lngRow = .Row
                '首先还原以前的
                If mlngRow >= 0 And mlngRow < .Rows Then
                    .Row = mlngRow: .Col = 0
                    .CellForeColor = &H80000008
                    .CellBackColor = &H80000005
                End If
                '接着设置当前的
                .Row = lngRow: .Col = 0
                .CellForeColor = &H80000005
                .CellBackColor = &H8000000D
                .Col = 1
                mlngRow = lngRow
            End If
        End If
        .Redraw = blnRedraw
        Call ShowCmdEdit
    End With
End Sub

Private Sub mshAttrib_KeyPress(KeyAscii As Integer)
    Dim strAttrib As String
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then Exit Sub
    
    strAttrib = mshAttrib.TextMatrix(mshAttrib.Row, 0)
    Select Case strAttrib
        Case "字体", "字体色", "背景色"
            '相当于单击按钮
            If KeyAscii = Asc("*") Then cmdEdit_Click
        Case "编号", "名称", "文本", "左边距", "顶边距", "宽度", "高度", "顺序号"
            Call ShowTxtEdit
            DoEvents
            If KeyAscii <> vbKeySpace Then
                txtEdit.Text = Chr(KeyAscii)
                txtEdit.SelStart = Len(txtEdit.Text)
            End If
        Case "门诊病人记帐", "住院统一记帐", "住院科室记帐", "医技科室记帐"
            If KeyAscii = vbKeySpace Then Call Set适用范围(strAttrib)
        Case "3D外观", "边框", "透明"
            If KeyAscii = vbKeySpace Then Call Set风格(strAttrib)
    End Select
End Sub

Private Sub mshAttrib_DblClick()
    Dim strAttrib As String
    
    strAttrib = mshAttrib.TextMatrix(mshAttrib.Row, 0)
    Select Case strAttrib
        Case "字体", "字体色", "背景色"
            '相当于单击按钮
            Call cmdEdit_Click
        Case "编号", "名称", "文本", "左边距", "顶边距", "宽度", "高度", "顺序号"
            
            Call ShowTxtEdit
        Case "门诊病人记帐", "住院统一记帐", "住院科室记帐", "医技科室记帐"
            Call Set适用范围(strAttrib)
        Case "3D外观", "边框", "透明"
            Call Set风格(strAttrib)
    End Select
End Sub

Private Sub Set适用范围(ByVal str范围 As String)
    With mshAttrib
        .TextMatrix(.Row, 1) = IIf(.TextMatrix(.Row, 1) = "允许", "禁止", "允许")
        Select Case str范围
            Case "门诊病人记帐"
                mstr适用范围 = IIf(.TextMatrix(.Row, 1) = "允许", "1", "0") & Mid(mstr适用范围, 2)
            Case "住院统一记帐"
                mstr适用范围 = Mid(mstr适用范围, 1, 1) & IIf(.TextMatrix(.Row, 1) = "允许", "1", "0") & Mid(mstr适用范围, 3)
            Case "住院科室记帐"
                mstr适用范围 = Mid(mstr适用范围, 1, 2) & IIf(.TextMatrix(.Row, 1) = "允许", "1", "0") & Mid(mstr适用范围, 4)
            Case "医技科室记帐"
                mstr适用范围 = Mid(mstr适用范围, 1, 3) & IIf(.TextMatrix(.Row, 1) = "允许", "1", "0")
        End Select
    End With
    mblnChange = True
End Sub

Private Sub Set风格(ByVal str风格 As String)
    Dim ctlTemp As Control
    With mshAttrib
        .TextMatrix(.Row, 1) = IIf(.TextMatrix(.Row, 1) = "√", "", "√")
        Select Case str风格
            Case "3D外观"
                For Each ctlTemp In mcolSelect
                    ctlTemp.Appearance = IIf(.TextMatrix(.Row, 1) = "√", 1, 0)
                Next
                Call ShowAttrib
            Case "边框"
                For Each ctlTemp In mcolSelect
                    ctlTemp.BorderStyle = IIf(.TextMatrix(.Row, 1) = "√", 1, 0)
                Next
            Case "透明"
                For Each ctlTemp In mcolSelect
                    ctlTemp.BackStyle = IIf(.TextMatrix(.Row, 1) = "√", 0, 1)
                Next
        End Select
    End With
    mblnChange = True
End Sub

Private Sub ShowCmdEdit()
'根据属性表选中行的属性，显示或隐藏按钮
    cmdEdit.Visible = False
    txtEdit.Visible = False
    With mshAttrib
        Select Case .TextMatrix(.Row, 0)
            Case "字体", "字体色", "背景色"
                If 250 * .Rows > .Height - mlngWidthAdj Then
                    '会出现纵向滚动条
                    cmdEdit.Left = .Left + .CellLeft + .CellWidth - cmdEdit.Width - 300
                Else
                    cmdEdit.Left = .Left + .CellLeft + .CellWidth - cmdEdit.Width - 60
                End If
                
                cmdEdit.Top = .Top + .CellTop
                cmdEdit.Visible = True
        End Select
    End With
End Sub

Private Sub ShowTxtEdit()
'根据属性表选中行的属性，显示文本编辑框
    cmdEdit.Visible = False
    With mshAttrib
        txtEdit.Left = .Left + .CellLeft + 30
        If 250 * .Rows > .Height - mlngWidthAdj Then
            '会出现纵向滚动条
            txtEdit.Width = .CellWidth - 330
        Else
            txtEdit.Width = .CellWidth - 90
        End If
        
        txtEdit.Top = .Top + .CellTop + 15
        txtEdit.Text = .TextMatrix(.Row, 1)
        txtEdit.Visible = True
        txtEdit.SetFocus
    End With
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If txtEdit.Visible = True Then
            Call AssignValue
            '不管赋值成功与否，都要使用输入框不可见
            txtEdit.Visible = False
            Exit Sub
        End If
        msngX = X
        msngY = Y
        SetSelectRect X, Y
        shpSelect.Tag = "开始"
        SetCapture picForm.hwnd
    End If
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If shpSelect.Tag = "开始" Then SetSelectRect X, Y
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If shpSelect.Tag = "开始" Then
        ReleaseCapture
        SetSelectRect X, Y
        shpSelect.Tag = ""
        Call DrawSelect
        
        Select Case mlngMoveReason
            Case 1 '增加文本
                With shpSelect
                    AddText .Left, .Top, .Width, .Height
                End With
            Case 2 '增加线条
                With shpSelect
                    AddLine .Left, .Top, .Width, .Height
                End With
            Case Else
                '选择控件
                Call SelectControl(Shift)
        End Select
    End If
    shpSelect.Tag = ""
    Call DrawSelect
    picForm.SetFocus
    picForm.MousePointer = 0
    mlngMoveReason = 0 '不管是左键还是右键松开，都不再处于新增项目状态
End Sub

Private Sub picSplitLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mblnDown = True
        msngX = X
    End If
End Sub

Private Sub picSplitLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 And mblnDown = True Then
        sngTemp = picSplitLeft.Left + (X - msngX) * 15
        If sngTemp > 600 And IIf(picSplitRight.Visible = True, picSplitRight.Left, ScaleWidth) - sngTemp > 600 Then
            picSplitLeft.Left = sngTemp
            fraList.Width = picSplitLeft.Left - fraList.Left
            
            Call Form_Resize
        End If
        
    End If
End Sub

Private Sub picSplitLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
End Sub

Private Sub picsplitright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mblnDown = True
        msngX = X
    End If
End Sub

Private Sub picsplitright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 And mblnDown = True Then
        sngTemp = picSplitRight.Left + (X - msngX) * 15
        If ScaleWidth - sngTemp > 600 And sngTemp - IIf(picSplitLeft.Visible = True, picSplitLeft.Left, 0) > 660 Then
            picSplitRight.Left = sngTemp
            fraAttrib.Left = picSplitRight.Left + picSplitRight.Width
            fraAttrib.Width = ScaleWidth - fraAttrib.Left
            
            Call Form_Resize
        End If
    End If
End Sub

Private Sub picSplitRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Design"
            mnuFileDesign_Click
        Case "Save"
            mnuFileSave_Click
        Case "Delete"
            mnuFileErase_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Lock"
            mnuFormatLock_Click
        Case "Element"
            mnuEditElements_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "Align"
            Call mnuFormatDoAlign_Click(mintAlign)
        Case "Form"
            Call mnuFormatFormAlign_Click(mintForm)
        Case "Distance"
            Select Case mintDistance
                Case 0 '横间距相同
                    mnuFormatHscSpace_Click 0
                Case 1 '无横间距
                    mnuFormatHscSpace_Click 3
                Case 2 '竖间距相同
                    mnuFormatVscSpace_Click 0
                Case 3 '无竖间距
                    mnuFormatVscSpace_Click 3
            End Select
        Case "Size"
            Call mnuFormatSize_Click(mintSize)
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    Select Case ButtonMenu.Parent.Key
        Case "View"
            For i = 0 To 3
                mnuViewIcon(i).Checked = False
            Next
            mnuViewIcon(ButtonMenu.Index - 1).Checked = True
            lvwMain.View = ButtonMenu.Index - 1
        Case "Align"
            ButtonMenu.Parent.ToolTipText = GetFore(ButtonMenu.Text)
            mintAlign = ButtonMenu.Index - 1
            Call mnuFormatDoAlign_Click(ButtonMenu.Index - 1)
        Case "Form"
            ButtonMenu.Parent.ToolTipText = GetFore(ButtonMenu.Text)
            mintForm = ButtonMenu.Index - 1
            Call mnuFormatFormAlign_Click(ButtonMenu.Index - 1)
        Case "Distance"
            ButtonMenu.Parent.ToolTipText = GetFore(ButtonMenu.Text)
            Select Case ButtonMenu.Text
                Case "横间距相同"
                    mintDistance = 0
                    mnuFormatHscSpace_Click 0
                Case "无横间距"
                    mintDistance = 1
                    mnuFormatHscSpace_Click 3
                Case "竖间距相同"
                    mintDistance = 2
                    mnuFormatVscSpace_Click 0
                Case "无竖间距"
                    mintDistance = 3
                    mnuFormatVscSpace_Click 3
            End Select
        Case "Size"
            ButtonMenu.Parent.ToolTipText = GetFore(ButtonMenu.Text)
            mintSize = ButtonMenu.Index - 1
            Call mnuFormatSize_Click(ButtonMenu.Index - 1)
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'用于对关闭按钮进行做图
Private Sub picClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As RECT
    With picClose(Index)
        r.Left = .ScaleLeft
        r.Top = .ScaleTop
        r.Right = .ScaleLeft + .ScaleWidth
        r.Bottom = .ScaleTop + .ScaleHeight
        .ForeColor = .BackColor
        DrawEdge .hdc, r, BDR_SUNKENOUTER, BF_RECT
    End With
End Sub

Private Sub picClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As RECT
    With picClose(Index)
        r.Left = .ScaleLeft
        r.Top = .ScaleTop
        r.Right = .ScaleLeft + .ScaleWidth
        r.Bottom = .ScaleTop + .ScaleHeight
        
        If X < .ScaleLeft Or X > .ScaleWidth Or Y < ScaleTop Or Y > .ScaleHeight Then
            Call ReleaseCapture
            .ForeColor = .BackColor
            Rectangle .hdc, .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
        Else
            SetCapture .hwnd
            .ForeColor = .BackColor
            DrawEdge .hdc, r, IIf(Button = 0, BDR_RAISEDINNER, BDR_SUNKENOUTER), BF_RECT
        End If
    End With
End Sub

Private Sub picClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call picClose_MouseMove(Index, Button, Shift, x, y)
    If Index = 0 Then
        Call mnuViewList_Click
    Else
        Call mnuViewAttrib_Click
    End If
End Sub

Private Sub picClose_Paint(Index As Integer)
    With picClose(Index)
        .ForeColor = .BackColor
        Rectangle .hdc, .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
        .CurrentX = (.ScaleWidth - .TextWidth("r")) / 2
        .CurrentY = (.ScaleHeight - .TextHeight("r")) / 2
        .ForeColor = 0
        picClose(Index).Print "r" '用于Marlett字体
    End With
End Sub

Private Sub picSplitLeft_Paint()
    Dim r As RECT
    With picSplitLeft
        r.Left = -3
        r.Top = -3
        r.Right = .ScaleLeft + .ScaleWidth
        r.Bottom = .ScaleTop + .ScaleHeight + 6
        
        DrawEdge .hdc, r, EDGE_RAISED, BF_RECT
    End With
End Sub

Private Sub picSplitRight_Paint()
    Dim r As RECT
    With picSplitRight
        r.Left = .ScaleLeft
        r.Top = .ScaleTop - 3
        r.Right = 3 + .ScaleWidth
        r.Bottom = 6 + .ScaleHeight
        
        DrawEdge .hdc, r, EDGE_RAISED, BF_RECT
    End With
End Sub
Private Sub ChangeSelectBefore(ctlSelect As Control)
'改变当前的焦点控件之前
    Dim i As Long
    
    Call HideAttach
    On Error Resume Next
    Set mctlSelect = mcolSelect(ctlSelect.Name & ctlSelect.Index)
    If Err <> 0 Then
        Err.Clear
        '该控件不在已选的控件中
        For i = 1 To mcolSelect.Count
            mcolSelect.Remove 1
        Next
        Set mctlSelect = ctlSelect
        mcolSelect.Add mctlSelect, mctlSelect.Name & mctlSelect.Index
    End If
End Sub

Private Sub ChangeSelectAfter()
'改变当前的焦点控件之后
    
    Dim i As Long
    If mcolSelect.Count = 1 Then
        If GetFore(cmbControl.Text) = mctlSelect.Tag Then
            ShowAttach
            ShowAttrib
            Exit Sub
        End If
        For i = 0 To cmbControl.ListCount - 1
            If GetFore(cmbControl.List(i)) = mctlSelect.Tag Then
                cmbControl.ListIndex = i
                Exit Sub
            End If
        Next
    Else
        If cmbControl.ListIndex = -1 Then
            Call ShowAttach
            Call ShowAttrib
        Else
            cmbControl.ListIndex = -1
        End If
    End If
End Sub

Private Sub AddSelect(ctlSelect As Control)
'删除或增添选中控件
    Dim i As Long
    On Error Resume Next
    Set mctlSelect = mcolSelect(ctlSelect.Name & ctlSelect.Index)
    If Err <> 0 Then
        Err.Clear
        '该控件不在已选的控件中，则加入
        mcolSelect.Add ctlSelect, ctlSelect.Name & ctlSelect.Index
        Set mctlSelect = ctlSelect
        If mcolSelect.Count = 1 Then
            For i = 0 To cmbControl.ListCount - 1
                If GetFore(cmbControl.List(i)) = mctlSelect.Tag Then
                    cmbControl.ListIndex = i
                    Exit Sub
                End If
            Next
        Else
            '什么属性也不显示
            If cmbControl.ListIndex = -1 Then
                '以前就是选中各个控件，再赋值也就不会激活事件
                '所以只有手工调用
                Call HideAttach
                Call ShowAttach
                Call ShowAttrib
            Else
                cmbControl.ListIndex = -1
            End If
        End If
    Else
        '从选取集合中删去
        mcolSelect.Remove ctlSelect.Name & ctlSelect.Index
        If mcolSelect.Count > 0 Then
            Set mctlSelect = mcolSelect(1)
            If mcolSelect.Count = 1 Then
                For i = 0 To cmbControl.ListCount - 1
                    If GetFore(cmbControl.List(i)) = mctlSelect.Tag Then
                        cmbControl.ListIndex = i
                        Exit Sub
                    End If
                Next
            Else
                '什么属性也不显示
                '以前就是选中各个控件，再赋值也就不会激活事件
                '所以只有手工调用
                Call HideAttach
                Call ShowAttach
                Call ShowAttrib
            End If
        Else
            Set mctlSelect = Nothing
            cmbControl.ListIndex = GetIndexOfBill
        End If
    End If
End Sub

Private Sub MoveControl(X As Single, Y As Single, Optional ByVal blnIndex As Boolean = False)
'参数:blnIndex mintIndex是否有效
    
    Dim ctlTemp As Control
    
    If mnuFormatLock.Checked = True Then Exit Sub
    For Each ctlTemp In mcolSelect
        If blnIndex = True Then
            If ctlTemp Is cmb(mintIndex) Then
                picCombo.Left = picCombo.Left + X - msngX
                picCombo.Top = picCombo.Top + Y - msngY
                ctlTemp.Left = picCombo.Left
                ctlTemp.Top = picCombo.Top
            Else
                ctlTemp.Left = ctlTemp.Left + X - msngX
                ctlTemp.Top = ctlTemp.Top + Y - msngY
            End If
        Else
            ctlTemp.Left = ctlTemp.Left + X - msngX
            ctlTemp.Top = ctlTemp.Top + Y - msngY
        End If
    Next
    mblnChange = True
End Sub

Private Sub ResizeControl(ByVal lngChange As Long, ByVal lngAttrib As Long)
    Dim ctlTemp As Control
    Dim lngRow As Long
    
    If fraAdjust(4).Visible = True Or lblAdjust(0).Visible = True Then
        Call HideAttach
    End If
    '更新位置
    Select Case lngAttrib
        Case 1
            For Each ctlTemp In mcolSelect
                ctlTemp.Left = ctlTemp.Left + lngChange
            Next
        Case 2
            For Each ctlTemp In mcolSelect
                ctlTemp.Top = ctlTemp.Top + lngChange
            Next
        Case 3
            For Each ctlTemp In mcolSelect
                If ctlTemp.Width + lngChange > mlngWidthAdj Then
                    ctlTemp.Width = ctlTemp.Width + lngChange
                End If
            Next
        Case 4
            For Each ctlTemp In mcolSelect
                If TypeName(ctlTemp) <> "ComboBox" Then
                    'ComboBox控件不能改变高度
                    If ctlTemp.Height + lngChange > mlngWidthAdj Then
                        ctlTemp.Height = ctlTemp.Height + lngChange
                    End If
                End If
            Next
    End Select
    
    '更新属性表
    If mcolSelect.Count = 1 Then
        For lngRow = 0 To mshAttrib.Rows - 1
            Select Case mshAttrib.TextMatrix(lngRow, 0)
                Case "左边距"
                    mshAttrib.TextMatrix(lngRow, 1) = mctlSelect.Left
                Case "顶边距"
                    mshAttrib.TextMatrix(lngRow, 1) = mctlSelect.Top
                Case "宽度"
                    mshAttrib.TextMatrix(lngRow, 1) = mctlSelect.Width
                Case "高度"
                    mshAttrib.TextMatrix(lngRow, 1) = mctlSelect.Height
            End Select
        Next
    End If
    
    mblnChange = True
End Sub

Private Function SetSelectRect(X As Single, Y As Single)
'设置选择框的位置
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    
    lngLeft = IIf(X < msngX, X, msngX)
    lngTop = IIf(Y < msngY, Y, msngY)
    lngWidth = Abs(X - msngX)
    lngHeight = Abs(Y - msngY)
    
    If mlngMoveReason <> 2 Then
        shpSelect.Left = lngLeft
        shpSelect.Top = lngTop
        shpSelect.Width = lngWidth
        shpSelect.Height = lngHeight
    Else
        '线条的显示显示效果要单独处理
        If lngWidth > lngHeight Then
            shpSelect.Width = lngWidth
            shpSelect.Height = 15
            '横线，那只能用原始顶边距
            shpSelect.Left = lngLeft
            shpSelect.Top = msngY
        Else
            shpSelect.Width = 15
            shpSelect.Height = lngHeight
            '纵线，那只能用原始左边距
            shpSelect.Left = msngX
            shpSelect.Top = lngTop
        End If
    End If
    '作图
    Call DrawSelect
End Function

Private Sub SelectControl(ByVal Shift As Integer)
'根据选择框的位置，对控件进行选取
    Dim objTemp As Element
    Dim ctlTemp As Control
    Dim i As Long
    
    
    If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
        '首先清除已有的选择控件集
        For i = 1 To mcolSelect.Count
            mcolSelect.Remove 1
        Next
        Set mctlSelect = Nothing
    End If
    
    On Error Resume Next
    For Each objTemp In mcolBill
        Set ctlTemp = objTemp.Control
        If ctlTemp.Visible = True And objTemp.Visible = True Then
            '除了要求单据本身的控件，还有一些辅助控件也要排除
            '判断控件是否选中
            If Not (ctlTemp.Left + ctlTemp.Width < shpSelect.Left Or _
                ctlTemp.Left > shpSelect.Left + shpSelect.Width Or _
                ctlTemp.Top + ctlTemp.Height < shpSelect.Top Or _
                ctlTemp.Top > shpSelect.Top + shpSelect.Height) Then
                '位置适合，加入集合，以"名称+索引"为关键字
                
                If Shift = vbCtrlMask Then
                    '用户在选择的同时按下了Ctrl键
                    mcolSelect.Remove ctlTemp.Name & ctlTemp.Index
                    If Err <> 0 Then
                        Err.Clear
                        '其先是没有选择该控件的，这次把它加上
                        mcolSelect.Add ctlTemp, ctlTemp.Name & ctlTemp.Index
                    End If
                ElseIf Shift = vbShiftMask Then
                    mcolSelect.Add ctlTemp, ctlTemp.Name & ctlTemp.Index
                    If Err <> 0 Then Err.Clear
                Else
                    mcolSelect.Add ctlTemp, ctlTemp.Name & ctlTemp.Index
                End If
            End If
        End If
    Next
    If mcolSelect.Count > 0 Then
        Set mctlSelect = mcolSelect(1)
        If mcolSelect.Count = 1 Then
            For i = 0 To cmbControl.ListCount - 1
                If GetFore(cmbControl.List(i)) = mctlSelect.Tag Then
                    cmbControl.ListIndex = i
                    Exit Sub
                End If
            Next
        Else
            '什么属性也不显示
            If cmbControl.ListIndex = -1 Then
                '以前就是选中各个控件，再赋值也就不会激活事件
                '所以只有手工调用
                Call HideAttach
                Call ShowAttach
                Call ShowAttrib
            Else
                cmbControl.ListIndex = -1
            End If
        End If
    Else
        '显示“记帐单”
        Set mctlSelect = Nothing
        cmbControl.ListIndex = GetIndexOfBill
    End If
End Sub

Private Sub HideAttach()
'隐藏位于选中控件周围的方块
    Dim lngCount As Long

    '首先隐藏所有方块
    For lngCount = fraAdjust.LBound To fraAdjust.UBound
        fraAdjust(lngCount).Visible = False
    Next
    For lngCount = lblAdjust.LBound To lblAdjust.UBound
        lblAdjust(lngCount).Visible = False
    Next
End Sub

Private Sub ShowAttach()
'显示位于选中控件周围的方块
    Dim lngCount As Long, i As Long
    Dim ctlTemp As Control
    
    If mcolSelect.Count = 0 Then
        '一个也没有选中，则显示
        Set fraAdjust(3).Container = Me '需要改变该控件的容器
        fraAdjust(3).ZOrder 1
        Set fraAdjust(4).Container = Me
        fraAdjust(4).ZOrder 1
        Set fraAdjust(5).Container = Me
        fraAdjust(5).ZOrder 1
            
        Call SetAttach(picForm, Array(-1, -1, -1, 3, 4, 5, -1, -1), fraAdjust)
        
        fraAdjust(3).Visible = True
        fraAdjust(4).Visible = True
        fraAdjust(5).Visible = True
        shpAdjust(3).Visible = True
        shpAdjust(4).Visible = True
        shpAdjust(5).Visible = True
    Else
        If mcolSelect.Count = 1 Then
            Set fraAdjust(3).Container = picForm '需要改变该控件的容器
            Set fraAdjust(4).Container = picForm '需要改变该控件的容器
            Set fraAdjust(5).Container = picForm '需要改变该控件的容器
            
            Call SetAttach(mctlSelect, Array(0, 1, 2, 3, 4, 5, 6, 7), fraAdjust)
            For lngCount = 0 To 7
                fraAdjust(lngCount).Visible = True
                shpAdjust(lngCount).Visible = True
                fraAdjust(lngCount).ZOrder       '这次要置上
            Next
        Else
            If lblAdjust.Count < mcolSelect.Count * 8 Then
                '已有的控件数不足显示，再装入新的
                For lngCount = lblAdjust.Count To mcolSelect.Count * 8 - 1
                    Load lblAdjust(lngCount)
                    Set lblAdjust(lngCount).Container = picForm
                    lblAdjust(lngCount).TabIndex = picForm.TabIndex
                Next
            End If
            i = 0
            For Each ctlTemp In mcolSelect
                If ctlTemp Is mctlSelect Then
                    For lngCount = i * 8 To i * 8 + 7
                        '有一个控件的方块颜色不同
                        lblAdjust(lngCount).BackColor = &HFF0000 '深蓝
                    Next
                Else
                    For lngCount = i * 8 To i * 8 + 7
                        lblAdjust(lngCount).BackColor = &HFFFF80 '天蓝
                    Next
                End If
                lngCount = i * 8
                Call SetAttach(ctlTemp, Array(lngCount, lngCount + 1, lngCount + 2, lngCount + 3, _
                    lngCount + 4, lngCount + 5, lngCount + 6, lngCount + 7), lblAdjust)
                '重新显示
                For lngCount = i * 8 To i * 8 + 7
                    lblAdjust(lngCount).Visible = True
                Next
                
                i = i + 1
            Next
        End If
    End If
    
End Sub

Private Sub SetAttach(ctlRefer As Control, varIndex As Variant, ctlSet As Variant)
'功能：设置各个方块的位置
'参数：ctlRefer   参照控件
'      varIndex   按位置顺序0-7得到的控件的索引
'      0   1   2
'      7       3
'      6   5   4
'      ctlSet     要设置的控件数组
    Dim lngIndex As Long

    With ctlRefer
        '0号位的方块
        lngIndex = varIndex(0)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left - mlngWidthAdj
            ctlSet(lngIndex).Top = .Top - mlngWidthAdj
        End If
        
        '1号位的方块
        lngIndex = varIndex(1)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left + (.Width - mlngWidthAdj) / 2
            ctlSet(lngIndex).Top = .Top - mlngWidthAdj
        End If
        
        '2号位的方块
        lngIndex = varIndex(2)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left + .Width
            ctlSet(lngIndex).Top = .Top - mlngWidthAdj
        End If
        
        '3号位的方块
        lngIndex = varIndex(3)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left + .Width
            ctlSet(lngIndex).Top = .Top + (.Height - mlngWidthAdj) / 2
        End If
        
        '4号位的方块
        lngIndex = varIndex(4)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left + .Width
            ctlSet(lngIndex).Top = .Top + .Height
        End If
        
        '5号位的方块
        lngIndex = varIndex(5)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left + (.Width - mlngWidthAdj) / 2
            ctlSet(lngIndex).Top = .Top + .Height
        End If
        
        '6号位的方块
        lngIndex = varIndex(6)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left - mlngWidthAdj
            ctlSet(lngIndex).Top = .Top + .Height
        End If
        
        '7号位的方块
        lngIndex = varIndex(7)
        If lngIndex > -1 Then
            ctlSet(lngIndex).Left = .Left - mlngWidthAdj
            ctlSet(lngIndex).Top = .Top + (.Height - mlngWidthAdj) / 2
        End If
    End With
End Sub

Private Sub ShowComplexAttrib()
'显示复合对象的属性
    '以下逻辑值如果为真，表示该项目被排除
    Dim blnFont As Boolean, strFont As String            '是否显示字体，及共同的字体名称
    Dim blnForeColor As Boolean, lngForeColor  As Long   '是否显示字体色，及共同值
    Dim blnBackColor As Boolean, lngBackColor  As Long   '是否显示背景色，及共同值
    Dim blnAppearance As Boolean, lngAppearance  As Long '是否显示3D外观，及共同值
    Dim blnBorderStyle As Boolean, lngBorderStyle As Long '是否显示边框，及共同值
    Dim blnBackStyle As Boolean, lngBackStyle As Long     '是否显示透明，及共同值
    
    Dim ctlTemp As Control, lngRow As Long
    
    stbThis.Panels(2).Text = "当前共选中" & mcolSelect.Count & "个对象"
    If mcolSelect.Count = 0 Then Exit Sub
    
    strFont = " ": lngForeColor = -1: lngBackColor = -1
    lngAppearance = -1: lngBorderStyle = -1: lngBackStyle = -1
    For Each ctlTemp In mcolSelect
        Select Case TypeName(ctlTemp)
            Case "ComboBox"
                blnAppearance = True
                blnBackStyle = True
                blnBorderStyle = True
            Case "CheckBox"
                blnBackStyle = True
                blnBorderStyle = True
            Case "CommandButton"
                blnForeColor = True
                blnBackColor = True
                blnAppearance = True
                blnBackStyle = True
                blnBorderStyle = True
            Case "Label"
            
            Case "TextBox"
                blnBackStyle = True
        End Select
        
        If blnFont = False Then
            '还需要设置字体的共同值
            If strFont = " " Then
                strFont = ctlTemp.Font.Name & "(" & ctlTemp.Font.Size & ")"
            Else
                If strFont <> ctlTemp.Font.Name & "(" & ctlTemp.Font.Size & ")" Then
                    strFont = ""
                End If
            End If
        End If
        If blnForeColor = False Then
            '还需要设置前景色的共同值
            If lngForeColor = -1 Then
                lngForeColor = ctlTemp.ForeColor
            Else
                If lngForeColor <> ctlTemp.ForeColor Then
                    lngForeColor = 0
                End If
            End If
        End If
        If blnBackColor = False Then
            '还需要设置前景色的共同值
            If lngBackColor = -1 Then
                lngBackColor = ctlTemp.BackColor
            Else
                If lngBackColor <> ctlTemp.BackColor Then
                    lngBackColor = RGB(255, 255, 255)
                End If
            End If
        End If
        If blnAppearance = False Then
            '还需要设置3D的共同值
            If lngAppearance = -1 Then
                lngAppearance = ctlTemp.Appearance
            Else
                If lngAppearance <> ctlTemp.Appearance Then
                    lngAppearance = 0
                End If
            End If
        End If
        If blnBorderStyle = False Then
            '还需要设置边框的共同值
            If lngBorderStyle = -1 Then
                lngBorderStyle = ctlTemp.BorderStyle
            Else
                If lngBorderStyle <> ctlTemp.BorderStyle Then
                    lngBorderStyle = 0
                End If
            End If
        End If
        If blnBackStyle = False Then
            '还需要设置边框的共同值
            If lngBackStyle = -1 Then
                lngBackStyle = ctlTemp.BackStyle
            Else
                If lngBackStyle <> ctlTemp.BackStyle Then
                    lngBackStyle = 1
                End If
            End If
        End If
    Next

    With mshAttrib
        .Rows = 1 '最初只设一行
        '字体是肯定要显示的
        .TextMatrix(0, 0) = "字体"
        .TextMatrix(0, 1) = strFont
        
        If blnForeColor = False Then
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            
            .TextMatrix(lngRow, 0) = "字体色"
            .TextMatrix(lngRow, 1) = "██████"
            .Row = lngRow: .Col = 1: .CellForeColor = lngForeColor
        End If
        If blnBackColor = False Then
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            
            .TextMatrix(lngRow, 0) = "背景色"
            .TextMatrix(lngRow, 1) = "██████"
            .Row = lngRow: .Col = 1: .CellForeColor = lngBackColor
        End If
        If blnAppearance = False Then
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            
            .TextMatrix(lngRow, 0) = "3D外观"
            .TextMatrix(lngRow, 1) = IIf(lngAppearance = 1, "√", "")
        End If
        If blnBorderStyle = False Then
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            
            .TextMatrix(lngRow, 0) = "边框"
            .TextMatrix(lngRow, 1) = IIf(lngBorderStyle = 1, "√", "")
        End If
        If blnBackStyle = False Then
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            
            .TextMatrix(lngRow, 0) = "透明"
            .TextMatrix(lngRow, 1) = IIf(lngBackStyle = 0, "√", "")
        End If
        
    End With
    
End Sub

Private Sub ShowAttrib()
'显示位于选中控件的属性
    Dim strAttrib As String
    Dim lngRow As Long, lngCount As Long
    
With mshAttrib
    .Redraw = False
    
    lngRow = .Row
    strAttrib = .TextMatrix(lngRow, 0)
    Call ClearTable
    If mcolSelect.Count = 0 Then
        stbThis.Panels(2).Text = "当前没有选中对象"
        '显示单据的属性
        .Rows = 10
        .TextMatrix(0, 0) = "编号"
        .TextMatrix(0, 1) = mstr编号
        .TextMatrix(1, 0) = "名称"
        .TextMatrix(1, 1) = mstr名称
        .TextMatrix(2, 0) = "宽度"
        .TextMatrix(2, 1) = picForm.Width
        .TextMatrix(3, 0) = "高度"
        .TextMatrix(3, 1) = picForm.Height
        .TextMatrix(4, 0) = "字体"
        .TextMatrix(4, 1) = picForm.Font.Name & "(" & picForm.Font.Size & ")"
        .TextMatrix(5, 0) = "背景色"
        .TextMatrix(5, 1) = "██████"
        .Row = 5: .Col = 1: .CellForeColor = picForm.BackColor
        .TextMatrix(6, 0) = "门诊病人记帐"
        .TextMatrix(6, 1) = IIf(Mid(mstr适用范围, 1, 1) = "1", "允许", "禁止")
        .TextMatrix(7, 0) = "住院统一记帐"
        .TextMatrix(7, 1) = IIf(Mid(mstr适用范围, 2, 1) = "1", "允许", "禁止")
        .TextMatrix(8, 0) = "住院科室记帐"
        .TextMatrix(8, 1) = IIf(Mid(mstr适用范围, 3, 1) = "1", "允许", "禁止")
        .TextMatrix(9, 0) = "医技科室记帐"
        .TextMatrix(9, 1) = IIf(Mid(mstr适用范围, 4, 1) = "1", "允许", "禁止")
    ElseIf mcolSelect.Count = 1 Then
        stbThis.Panels(2).Text = "当前的选中对象是——“" & mctlSelect.Tag & "”"
        '显示选中控件的属性
        Select Case TypeName(mctlSelect)
            Case "CheckBox"
                .Rows = 10
                .TextMatrix(0, 0) = "名称"
                .TextMatrix(0, 1) = mctlSelect.Caption
                .TextMatrix(1, 0) = "左边距"
                .TextMatrix(1, 1) = mctlSelect.Left
                .TextMatrix(2, 0) = "顶边距"
                .TextMatrix(2, 1) = mctlSelect.Top
                .TextMatrix(3, 0) = "宽度"
                .TextMatrix(3, 1) = mctlSelect.Width
                .TextMatrix(4, 0) = "高度"
                .TextMatrix(4, 1) = mctlSelect.Height
                .TextMatrix(5, 0) = "字体"
                .TextMatrix(5, 1) = mctlSelect.Font.Name & "(" & mctlSelect.Font.Size & ")"
                .TextMatrix(6, 0) = "字体色"
                .TextMatrix(6, 1) = "██████"
                .Row = 6: .Col = 1: .CellForeColor = mctlSelect.ForeColor
                .TextMatrix(7, 0) = "背景色"
                .TextMatrix(7, 1) = "██████"
                .Row = 7: .Col = 1: .CellForeColor = mctlSelect.BackColor
                .TextMatrix(8, 0) = "3D外观"
                .TextMatrix(8, 1) = IIf(mctlSelect.Appearance = 1, "√", "")
                .TextMatrix(9, 0) = "顺序号"
                .TextMatrix(9, 1) = mcolBill(mctlSelect.Tag).TabIndex
            Case "ComboBox"
                .Rows = 7
                .TextMatrix(0, 0) = "左边距"
                .TextMatrix(0, 1) = mctlSelect.Left
                .TextMatrix(1, 0) = "顶边距"
                .TextMatrix(1, 1) = mctlSelect.Top
                .TextMatrix(2, 0) = "宽度"
                .TextMatrix(2, 1) = mctlSelect.Width
                .TextMatrix(3, 0) = "字体"
                .TextMatrix(3, 1) = mctlSelect.Font.Name & "(" & mctlSelect.Font.Size & ")"
                .TextMatrix(4, 0) = "字体色"
                .TextMatrix(4, 1) = "██████"
                .Row = 4: .Col = 1: .CellForeColor = mctlSelect.ForeColor
                .TextMatrix(5, 0) = "背景色"
                .TextMatrix(5, 1) = "██████"
                .Row = 5: .Col = 1: .CellForeColor = mctlSelect.BackColor
                .TextMatrix(6, 0) = "顺序号"
                .TextMatrix(6, 1) = mcolBill(mctlSelect.Tag).TabIndex
            Case "CommandButton"
                .Rows = 7
                .TextMatrix(0, 0) = "名称"
                .TextMatrix(0, 1) = mctlSelect.Caption
                .TextMatrix(1, 0) = "左边距"
                .TextMatrix(1, 1) = mctlSelect.Left
                .TextMatrix(2, 0) = "顶边距"
                .TextMatrix(2, 1) = mctlSelect.Top
                .TextMatrix(3, 0) = "宽度"
                .TextMatrix(3, 1) = mctlSelect.Width
                .TextMatrix(4, 0) = "高度"
                .TextMatrix(4, 1) = mctlSelect.Height
                .TextMatrix(5, 0) = "字体"
                .TextMatrix(5, 1) = mctlSelect.Font.Name & "(" & mctlSelect.Font.Size & ")"
                .TextMatrix(6, 0) = "顺序号"
                .TextMatrix(6, 1) = mcolBill(mctlSelect.Tag).TabIndex
            Case "Label"
                .Rows = 12
                .TextMatrix(0, 0) = "文本"
                .TextMatrix(0, 1) = mctlSelect.Caption
                .TextMatrix(1, 0) = "左边距"
                .TextMatrix(1, 1) = mctlSelect.Left
                .TextMatrix(2, 0) = "顶边距"
                .TextMatrix(2, 1) = mctlSelect.Top
                .TextMatrix(3, 0) = "宽度"
                .TextMatrix(3, 1) = mctlSelect.Width
                .TextMatrix(4, 0) = "高度"
                .TextMatrix(4, 1) = mctlSelect.Height
                .TextMatrix(5, 0) = "字体"
                .TextMatrix(5, 1) = mctlSelect.Font.Name & "(" & mctlSelect.Font.Size & ")"
                .TextMatrix(6, 0) = "字体色"
                .TextMatrix(6, 1) = "██████"
                .Row = 6: .Col = 1: .CellForeColor = mctlSelect.ForeColor
                .TextMatrix(7, 0) = "背景色"
                .TextMatrix(7, 1) = "██████"
                .Row = 7: .Col = 1: .CellForeColor = mctlSelect.BackColor
                .TextMatrix(8, 0) = "3D外观"
                .TextMatrix(8, 1) = IIf(mctlSelect.Appearance = 1, "√", "")
                .TextMatrix(9, 0) = "边框"
                .TextMatrix(9, 1) = IIf(mctlSelect.BorderStyle = 1, "√", "")
                .TextMatrix(10, 0) = "透明"
                .TextMatrix(10, 1) = IIf(mctlSelect.BackStyle = 0, "√", "")
                .TextMatrix(11, 0) = "顺序号"
                .TextMatrix(11, 1) = mcolBill(mctlSelect.Tag).TabIndex
            Case "TextBox"
                .Rows = 10
                .TextMatrix(0, 0) = "左边距"
                .TextMatrix(0, 1) = mctlSelect.Left
                .TextMatrix(1, 0) = "顶边距"
                .TextMatrix(1, 1) = mctlSelect.Top
                .TextMatrix(2, 0) = "宽度"
                .TextMatrix(2, 1) = mctlSelect.Width
                .TextMatrix(3, 0) = "高度"
                .TextMatrix(3, 1) = mctlSelect.Height
                .TextMatrix(4, 0) = "字体"
                .TextMatrix(4, 1) = mctlSelect.Font.Name & "(" & mctlSelect.Font.Size & ")"
                .TextMatrix(5, 0) = "字体色"
                .TextMatrix(5, 1) = "██████"
                .Row = 5: .Col = 1: .CellForeColor = mctlSelect.ForeColor
                .TextMatrix(6, 0) = "背景色"
                .TextMatrix(6, 1) = "██████"
                .Row = 6: .Col = 1: .CellForeColor = mctlSelect.BackColor
                .TextMatrix(7, 0) = "3D外观"
                .TextMatrix(7, 1) = IIf(mctlSelect.Appearance = 1, "√", "")
                .TextMatrix(8, 0) = "边框"
                .TextMatrix(8, 1) = IIf(mctlSelect.BorderStyle = 1, "√", "")
                .TextMatrix(9, 0) = "顺序号"
                .TextMatrix(9, 1) = mcolBill(mctlSelect.Tag).TabIndex
        End Select
    Else
        '显示多个对象共同的属性
        Call ShowComplexAttrib
    End If
    '还原行号
    For lngCount = 0 To .Rows - 1
        If .TextMatrix(lngCount, 0) = strAttrib Then
            '找相同属性的行
            .Row = lngCount
            Exit For
        End If
    Next
    If lngCount = .Rows Then
        '没找到
        '就用以前保存的行号做判断
        If lngRow > .Rows - 1 Then
            .Row = 0
        Else
            .Row = lngRow
        End If
    End If
    
    '只有手工调用
    Call mshAttrib_EnterCell
    .Redraw = True
End With
End Sub

Private Sub ClearTable()
    mshAttrib.Clear
    mshAttrib.Rows = 1
    mshAttrib.ColAlignmentFixed(0) = 1
    mshAttrib.ColAlignment(1) = 1
End Sub

Private Sub cmbControl_Click()
    Dim i As Integer
    
    If cmbControl.ListIndex > -1 Then
        For i = 1 To mcolSelect.Count
            mcolSelect.Remove 1
        Next
        If cmbControl.ListIndex <> GetIndexOfBill Then
            Set mctlSelect = mcolBill(GetFore(cmbControl.Text)).Control
            mcolSelect.Add mctlSelect, mctlSelect.Name & mctlSelect.Index
        End If
    Else
        Call ClearTable
    End If
    Call HideAttach
    Call ShowAttach
    Call ShowAttrib
    If fraAttrib.Visible = True And mshAttrib.Enabled = True Then mshAttrib.SetFocus
End Sub

Private Function GetIndexOfBill() As Long
'功能：得到记帐单控件的属性
    Dim lngCount As Long
    
    For lngCount = 0 To cmbControl.ListCount - 1
        If cmbControl.List(lngCount) = "记帐单" Then
            GetIndexOfBill = lngCount
            Exit Function
        End If
    Next
    GetIndexOfBill = -1
End Function

Private Sub LoadControlList()
'将当前单据上的所有控件都装到列表框中
    Dim objTemp As Element
    
    cmbControl.Clear
    cmbControl.AddItem "记帐单"
    For Each objTemp In mcolBill
        If objTemp.Visible = True Then
            '只允许加可见的控件
            If TypeName(objTemp.Control) = "Label" Then
                cmbControl.AddItem objTemp.Key & "(" & objTemp.Control.Caption & ")"
            Else
                cmbControl.AddItem objTemp.Key
            End If
        End If
    Next
    cmbControl.ListIndex = GetIndexOfBill()
End Sub

Private Sub FillList()
'装入所有的记帐单
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str范围 As String
    Dim lst As ListItem
    Dim strKey As String
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain.SelectedItem.Key
    End If
    
    On Error GoTo errHandle
    
    strSQL = "select ID,编号,名称,收费项目数,适用范围,宽度,高度,字体,背景色 from 收费记帐单"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    lvwMain.ListItems.Clear
    Do Until rsTmp.EOF
        Set lst = lvwMain.ListItems.Add(, "B" & rsTmp("ID"), rsTmp("名称"), "Bill", "Bill")
        lst.SubItems(1) = rsTmp("编号")
        str范围 = IIf(IsNull(rsTmp("适用范围")), "", rsTmp("适用范围"))
        lst.SubItems(2) = IIf(Mid(str范围, 1, 1) = "1", "√", "")
        lst.SubItems(3) = IIf(Mid(str范围, 2, 1) = "1", "√", "")
        lst.SubItems(4) = IIf(Mid(str范围, 3, 1) = "1", "√", "")
        lst.SubItems(5) = IIf(Mid(str范围, 4, 1) = "1", "√", "")
        lst.Tag = rsTmp("宽度") & "," & rsTmp("高度") & "," & rsTmp("收费项目数") _
                     & "," & rsTmp("背景色") & "," & IIf(IsNull(rsTmp("字体")), "宋体|9|0|0|0", rsTmp("字体"))
        rsTmp.MoveNext
    Loop
    If lvwMain.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            Set lst = lvwMain.ListItems(1)
            lst.Selected = True
            lst.EnsureVisible
        Else
            lst.Selected = True
            lst.EnsureVisible
        End If
        mstrKey = lst.Key
    Else
        mstrKey = ""
    End If
    stbThis.Panels(2).Text = "共有" & lvwMain.ListItems.Count & "张自定义记帐单。"
    Call FillBill
    Call SetMenu
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FillBill()
'把记帐单的内容显示出来
    Dim objTemp As Element
    Dim ctlTemp As Control
    Dim lngCount As Long
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim strName As String
    Dim varTemp As Variant
    
    On Error GoTo errHandle
    
    '恢复对齐等菜单组的缺省设置
    mintAlign = 0:    Toolbar1.Buttons("Align").ToolTipText = "左对齐"
    mintForm = 0:     Toolbar1.Buttons("Form").ToolTipText = "水平居中"
    mintDistance = 0: Toolbar1.Buttons("Distance").ToolTipText = "横间距相同"
    mintSize = 0:     Toolbar1.Buttons("Size").ToolTipText = "相同宽度"
    
    LockWindowUpdate picForm.hwnd
    '首先清除现有的集合，并把控件删除或隐藏
    For lngCount = 1 To mcolBill.Count
        Set ctlTemp = mcolBill(1).Control
        Select Case ctlTemp.Name
            Case "cmd"
                If ctlTemp.Index < 3 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                Else
                    Unload ctlTemp
                End If
            Case "chk"
                If ctlTemp.Index < 2 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                    ctlTemp.ForeColor = &H80000012
                    ctlTemp.BackColor = &H8000000F
                    ctlTemp.Appearance = 1
                Else
                    Unload ctlTemp
                End If
            Case "cmb"
                If ctlTemp.Index = 0 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                    ctlTemp.ForeColor = &H80000008
                    ctlTemp.BackColor = &H80000005
                Else
                    Unload ctlTemp
                End If
            Case "lbl"
                If ctlTemp.Index = 0 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                    ctlTemp.Appearance = 1
                    ctlTemp.BorderStyle = 0
                    ctlTemp.BackStyle = 0
                    ctlTemp.ForeColor = &H80000012
                    ctlTemp.BackColor = &H8000000F
                Else
                    Unload ctlTemp
                End If
            Case "txt"
                If ctlTemp.Index = 0 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                    ctlTemp.ForeColor = &H80000008
                    ctlTemp.BackColor = &H80000005
                    ctlTemp.Appearance = 1
                    ctlTemp.BorderStyle = 1
                Else
                    Unload ctlTemp
                End If
            Case Else
                If ctlTemp.Index = 0 Then
                    ctlTemp.Visible = False
                    SetFont ctlTemp, Me
                Else
                    Unload ctlTemp
                End If
        End Select
        mcolBill.Remove 1
    Next
    '刷新
    If lvwMain.SelectedItem Is Nothing Or mbln新增 = True Then
        mstr编号 = ""
        mstr名称 = ""
        mlng数量 = 0
        mstr适用范围 = "0000"
        picForm.Width = 11520
        picForm.Height = 6795
        picForm.BackColor = &H8000000F
        SetFont picForm, Me
    Else
        With lvwMain.SelectedItem
            mstr名称 = .Text
            mstr编号 = .SubItems(1)
            mstr适用范围 = IIf(.SubItems(2) = "", "0", "1") & IIf(.SubItems(3) = "", "0", "1") & _
                           IIf(.SubItems(4) = "", "0", "1") & IIf(.SubItems(5) = "", "0", "1")
                           
            varTemp = Split(.Tag, ",")
            picForm.Width = varTemp(0)
            picForm.Height = varTemp(1)
            mlng数量 = varTemp(2)
            picForm.BackColor = varTemp(3)
            varTemp = Split(varTemp(4), "|")
            picForm.Font.Name = varTemp(0)
            picForm.Font.Size = varTemp(1)
            picForm.Font.Bold = varTemp(2) = "1"
            picForm.Font.Italic = varTemp(3) = "1"
            picForm.Font.Underline = varTemp(4) = "1"
            
            strSQL = "select 对应字段,序号,类型,定义值,顺序号,左边,顶边,宽度,高度,字体,前景色,背景色,是否显示,外形,边框线,透明 " & _
                " from 收费记帐单定义 where 记帐ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(Mid(.Key, 2)))
        End With
        
        Do Until rsTmp.EOF
            strName = rsTmp("对应字段")
            Select Case rsTmp("类型")
                Case "CheckBox"
                    If strName = "加班标志" Then
                        Set ctlTemp = chk(0)
                    ElseIf strName = "销" Then
                        Set ctlTemp = chk(1)
                    Else
                        Load chk(chk.UBound + 1)
                        Set ctlTemp = chk(chk.UBound)
                    End If
                    ctlTemp.Caption = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                    ctlTemp.Height = rsTmp("高度")
                    ctlTemp.ForeColor = rsTmp("前景色")
                    ctlTemp.BackColor = rsTmp("背景色")
                    ctlTemp.Appearance = rsTmp("外形")
                Case "ComboBox"
                    If strName = "NO" Then
                        Set ctlTemp = cmb(0)
                    Else
                        Load cmb(cmb.UBound + 1)
                        Set ctlTemp = cmb(cmb.UBound)
                    End If
                    ctlTemp.ForeColor = rsTmp("前景色")
                    ctlTemp.BackColor = rsTmp("背景色")
                Case "CommandButton"
                    If strName = "确定" Then
                        Set ctlTemp = cmd(1)
                    ElseIf strName = "取消" Then
                        Set ctlTemp = cmd(2)
                    Else
                        If strName = "细目选择" And IIf(IsNull(rsTmp("序号")), "", rsTmp("序号")) = "1" Then
                            Set ctlTemp = cmd(0)
                        Else
                            Load cmd(cmd.UBound + 1)
                            Set ctlTemp = cmd(cmd.UBound)
                        End If
                    End If
                    ctlTemp.Caption = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                    ctlTemp.Height = rsTmp("高度")
                Case "Label"
                    Load lbl(lbl.UBound + 1)
                    Set ctlTemp = lbl(lbl.UBound)
                    ctlTemp.Caption = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                    ctlTemp.Appearance = rsTmp("外形")
                    ctlTemp.BorderStyle = rsTmp("边框线")
                    ctlTemp.BackStyle = rsTmp("透明")
                    ctlTemp.ForeColor = rsTmp("前景色")
                    ctlTemp.BackColor = rsTmp("背景色")
                    ctlTemp.Height = rsTmp("高度")
                Case "TextBox"
                    If strName = "发生时间" Then
                        Set ctlTemp = txt(0)
                    Else
                        Load txt(txt.UBound + 1)
                        Set ctlTemp = txt(txt.UBound)
                    End If
                    ctlTemp.Height = rsTmp("高度")
                    ctlTemp.ForeColor = rsTmp("前景色")
                    ctlTemp.BackColor = rsTmp("背景色")
                    ctlTemp.Appearance = rsTmp("外形")
                    ctlTemp.BorderStyle = rsTmp("边框线")
            End Select
            ctlTemp.Left = rsTmp("左边")
            ctlTemp.Top = rsTmp("顶边")
            ctlTemp.Width = rsTmp("宽度")
            varTemp = Split(rsTmp("字体"), "|")
            ctlTemp.Font.Name = varTemp(0)
            ctlTemp.Font.Size = varTemp(1)
            ctlTemp.Font.Bold = varTemp(2) = "1"
            ctlTemp.Font.Italic = varTemp(3) = "1"
            ctlTemp.Font.Underline = varTemp(4) = "1"
            
            If Nvl(rsTmp("序号"), 0) <> 0 Then
                If rsTmp("类型") = "Label" Then
                    strName = strName & "_" & ctlTemp.Index
                Else
                    strName = strName & "_" & rsTmp("序号")
                End If
            End If
            '最后完成向集合的增加
            Set objTemp = mcolBill.Add(strName, ctlTemp, rsTmp("顺序号"), (rsTmp("是否显示") = 1))
            
            If Mid(strName, 1, 4) = "收费类别" Or Mid(strName, 1, 4) = "收费细目" Or _
               Mid(strName, 1, 2) = "数次" Or strName = "开单部门" Then
                objTemp.Value = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
            End If
            rsTmp.MoveNext
        Loop
    End If
    '显示滚动条
    Call Form_Resize
    
    For Each objTemp In mcolBill
        objTemp.Control.Visible = objTemp.Visible '可见性
    Next
    LockWindowUpdate 0
    Call LoadControlList
    mblnChange = False
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function NumIsValid(ByVal strNumber As String) As Boolean
'功能:分析输入内容是否是有效的数字
'参数:strNumber  输入内容
'返回值:有效返回True,否则为False
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "请输入一个数值。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Abs(Val(strNumber)) > 20000 Then
        MsgBox "数值的绝对值不能超过20000。", vbExclamation, gstrSysName
        Exit Function
    End If
    NumIsValid = True
End Function

Private Sub SetMenu()
    Dim blnItem As Boolean
    Dim lngCount As Long
    
    If mblnReadonly = True Then
        '只读权限
        mnuFileNew.Enabled = False
        mnuFileErase.Enabled = False
        mnuFileSaveAs.Enabled = False
        mnuFileDesign.Enabled = False
        mnuFileDesign.Checked = False
        
        mnuFileSave.Enabled = False
        mnuFileReload.Enabled = False
        Toolbar1.Buttons("New").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        Toolbar1.Buttons("Design").Enabled = False
        Toolbar1.Buttons("Save").Enabled = False
    Else
        '文件菜单
        blnItem = Not lvwMain.SelectedItem Is Nothing
        mnuFileNew.Enabled = Not mnuFileDesign.Checked '如果处理设计状态，就不能再新增
        mnuFileErase.Enabled = mnuFileNew.Enabled And blnItem
        mnuFileSaveAs.Enabled = mnuFileNew.Enabled And blnItem
        
        mnuFileDesign.Enabled = mbln新增 Or blnItem   '只要有记帐单或处于新增状态，就可以进行设计
        If mnuFileDesign.Enabled = False Then mnuFileDesign.Checked = False
        
        mnuFileSave.Enabled = mnuFileDesign.Checked    '设计状态下才能保存
        mnuFileReload.Enabled = mnuFileDesign.Checked    '设计状态下才能重新装入
        Toolbar1.Buttons("New").Enabled = mnuFileNew.Enabled
        Toolbar1.Buttons("Delete").Enabled = mnuFileErase.Enabled
        Toolbar1.Buttons("Design").Enabled = mnuFileDesign.Enabled
        Toolbar1.Buttons("Save").Enabled = mnuFileSave.Enabled
    End If
    mnuFileImp.Enabled = mnuFileNew.Enabled
    mnuFileExp.Enabled = mnuFileSaveAs.Enabled
    
    blnItem = mnuFileDesign.Checked
    '几个特殊的控件
    lvwMain.Enabled = Not blnItem
    mnuViewRefresh.Enabled = Not blnItem
    
    picForm.Enabled = blnItem
    mshAttrib.Enabled = blnItem
    cmbControl.Enabled = blnItem
    If blnItem = False Then
        cmdEdit.Visible = False
        txtEdit.Visible = False
    End If
    For lngCount = 0 To 7
        fraAdjust(lngCount).Enabled = blnItem
    Next
    
    '编辑菜单
    mnuEditElements.Enabled = blnItem
    mnuEditText.Enabled = blnItem
    mnuEditLine.Enabled = blnItem
    mnuEditCopy.Enabled = blnItem
    mnuEditRemove.Enabled = blnItem
    mnuEditSelAll.Enabled = blnItem
    Toolbar1.Buttons("Element").Enabled = blnItem
    
    '格式菜单
    mnuFormatLock.Enabled = blnItem
    Toolbar1.Buttons("Lock").Enabled = blnItem
    If mnuFormatLock.Enabled = False Then mnuFormatLock.Checked = False
    Toolbar1.Buttons("Lock").Value = IIf(mnuFormatLock.Checked, tbrPressed, tbrUnpressed)
    Call SetFormatMenu
End Sub

Private Sub SetFormatMenu()
    
    Dim blnEnable As Boolean
    Dim lngCount As Long
    
    blnEnable = (Not mnuFormatLock.Checked) And mnuFormatLock.Enabled
    Toolbar1.Buttons("Align").Enabled = blnEnable
    Toolbar1.Buttons("Form").Enabled = blnEnable
    Toolbar1.Buttons("Distance").Enabled = blnEnable
    Toolbar1.Buttons("Size").Enabled = blnEnable
    
    For lngCount = mnuFormatDoAlign.LBound To mnuFormatDoAlign.UBound
        If mnuFormatDoAlign(lngCount).Caption <> "-" Then
            mnuFormatDoAlign(lngCount).Enabled = blnEnable
        End If
    Next
    For lngCount = mnuFormatFormAlign.LBound To mnuFormatFormAlign.UBound
        mnuFormatFormAlign(lngCount).Enabled = blnEnable
    Next
    For lngCount = mnuFormatHscSpace.LBound To mnuFormatHscSpace.UBound
        mnuFormatHscSpace(lngCount).Enabled = blnEnable
    Next
    For lngCount = mnuFormatVscSpace.LBound To mnuFormatVscSpace.UBound
        mnuFormatVscSpace(lngCount).Enabled = blnEnable
    Next
    For lngCount = mnuFormatSize.LBound To mnuFormatSize.UBound
        mnuFormatSize(lngCount).Enabled = blnEnable
    Next
    
    If blnEnable = True Then
'      varIndex   按位置顺序0-7得到的控件的索引
'      0   1   2
'      7       3
'      6   5   4
        For lngCount = 0 To 7
            shpAdjust(lngCount).BackColor = &H800000
        Next
        fraAdjust(0).MousePointer = vbSizeNWSE
        fraAdjust(1).MousePointer = vbSizeNS
        fraAdjust(2).MousePointer = vbSizeNESW
        fraAdjust(3).MousePointer = vbSizeWE
        fraAdjust(4).MousePointer = vbSizeNWSE
        fraAdjust(5).MousePointer = vbSizeNS
        fraAdjust(6).MousePointer = vbSizeNESW
        fraAdjust(7).MousePointer = vbSizeWE
    Else
        For lngCount = 0 To 7
            shpAdjust(lngCount).BackColor = RGB(255, 64, 0)
        Next
        fraAdjust(0).MousePointer = vbDefault
        fraAdjust(1).MousePointer = vbDefault
        fraAdjust(2).MousePointer = vbDefault
        fraAdjust(3).MousePointer = vbDefault
        fraAdjust(4).MousePointer = vbDefault
        fraAdjust(5).MousePointer = vbDefault
        fraAdjust(6).MousePointer = vbDefault
        fraAdjust(7).MousePointer = vbDefault
    End If
    
End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    Dim objTemp  As Button
    If InStr(gstrPrivs, "增删改") = 0 Then
        mblnReadonly = True
        mnuEdit.Visible = False
        mnuFormat.Visible = False
        
        mnuFile0.Visible = False
        mnuFile1.Visible = False
        mnuFile2.Visible = False
        mnuFileNew.Visible = False
        mnuFileErase.Visible = False
        mnuFileSaveAs.Visible = False
        mnuFileDesign.Visible = False
        mnuFileSave.Visible = False
        mnuFileReload.Visible = False
        mnuFileImp.Visible = False
        mnuFileExp.Visible = False
        
        For Each objTemp In Toolbar1.Buttons
            If objTemp.Key <> "Help" And objTemp.Key <> "Quit" Then
                objTemp.Visible = False
            End If
        Next
    Else
        mblnReadonly = False
    End If
End Sub

Private Function GetFore(ByVal strFull As String, Optional ByVal strSplit As String = "(") As String
'取给定字符串中某一字符之前的子串
    Dim lngPos As Long
    
    lngPos = InStr(strFull, strSplit)
    If lngPos = 0 Then
        GetFore = strFull
    Else
        GetFore = Mid(strFull, 1, lngPos - 1)
    End If
End Function

Private Sub DrawSelect()
'根据当前的选择在桌面上画出矩形框来
    Static sngLeft As Single, sngTop As Single, sngWidth As Single, sngHeight As Single
    
    If sngWidth <> 0 Or sngHeight <> 0 Then
        '重画一次，以清除上次所画的内容
        DrawRect sngLeft, sngTop, sngWidth, sngHeight
    End If
    
    If shpSelect.Tag = "开始" Then
        With shpSelect
            sngLeft = .Left
            sngTop = .Top
            sngWidth = .Width
            sngHeight = .Height
        End With
        DrawRect sngLeft, sngTop, sngWidth, sngHeight
    Else
        '设置关闭属性
        sngLeft = 0
        sngTop = 0
        sngWidth = 0
        sngHeight = 0
    End If
    
End Sub

Private Sub DrawRect(ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single, sngHeight As Single)
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
    p.X = lngLeft: p.Y = lngTop
    ClientToScreen picForm.hwnd, p
    lngLeft = p.X: lngTop = p.Y
    With picForm
        pLT.X = .ScaleLeft / lngPerX
        pLT.Y = .ScaleTop / lngPerY
        ClientToScreen picForm.hwnd, pLT '从现在开始该处保留最左、最上的值
        
        pRB.X = (.ScaleLeft + .ScaleWidth) / lngPerX
        pRB.Y = (.ScaleTop + .ScaleHeight) / lngPerY
        ClientToScreen picForm.hwnd, pRB '从现在开始该处保留最右、最下的值
    End With
    '计算边界超出情况
    With picForm
        If sngLeft + sngWidth > .ScaleWidth Then
            lngRight = pRB.X
        Else
            lngRight = lngLeft + lngWidth
        End If
        If sngTop + sngHeight > .ScaleHeight Then
            lngBottom = pRB.Y
        Else
            lngBottom = lngTop + lngHeight
        End If
        
        If sngTop < .ScaleTop Then lngTop = pLT.Y
        If sngLeft < .ScaleLeft Then lngLeft = pLT.X
    End With
    
    
    lngDC = GetDC(0)
    lngPen = SelectObject(lngDC, CreatePen(PS_DOT, 1, RGB(0, 0, 0)))
    lngROP = SetROP2(lngDC, R2_XORPEN)
    
    MoveToEx lngDC, lngLeft, lngTop, p
    LineTo lngDC, lngRight, lngTop
    LineTo lngDC, lngRight, lngBottom
    LineTo lngDC, lngLeft, lngBottom
    LineTo lngDC, lngLeft, lngTop
    
    lngPen = SelectObject(lngDC, lngPen)
    SetROP2 lngDC, lngROP
    DeleteObject lngPen
    ReleaseDC 0, lngDC
End Sub
    
Private Function GetValue(varTemp As ADODB.Field) As String
'根据数据库的类型返回相应的值
    If IsNull(varTemp) Then
        '空值
        GetValue = "Null"
    Else
        Select Case varTemp.Type
            Case adNumeric, adVarNumeric
                '数值
                GetValue = varTemp.Value
            Case adVarChar, adChar
                '字符串
                GetValue = "'" & Replace(varTemp.Value, "'", "''") & "'"
            Case adDBTimeStamp
                '日期
                GetValue = "To_date('" & Format(varTemp, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:mi:ss')"
            Case Else
                GetValue = "Null"
        End Select
    End If
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

