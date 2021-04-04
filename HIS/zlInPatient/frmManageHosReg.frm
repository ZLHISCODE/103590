VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageHosReg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "病人入院管理"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   Icon            =   "frmManageHosReg.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFind 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3870
      ScaleHeight     =   420
      ScaleWidth      =   3705
      TabIndex        =   10
      Top             =   810
      Width           =   3705
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   360
         Left            =   570
         TabIndex        =   11
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmManageHosReg.frx":058A
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   2
         ShowSortName    =   -1  'True
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   480
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   3795
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4320
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   840
      Width           =   45
   End
   Begin VB.ComboBox cboNodeList 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   645
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   2430
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6765
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageHosReg.frx":066D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "病人颜色说明"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10980
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   8355
      NewRow1         =   0   'False
      Child2          =   "chkOnly"
      MinWidth2       =   1200
      MinHeight2      =   300
      Width2          =   1065
      NewRow2         =   0   'False
      Begin VB.CheckBox chkOnly 
         Caption         =   "只显示门诊留观预约"
         Height          =   300
         Left            =   8550
         TabIndex        =   7
         Top             =   240
         Width           =   2340
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "入院"
               Key             =   "Add"
               Description     =   "入院"
               Object.ToolTipText     =   "对住院病人进行入院登记"
               Object.Tag             =   "入院"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "留观"
               Key             =   "Keep"
               Description     =   "留观"
               Object.ToolTipText     =   "对留观病人进行登记"
               Object.Tag             =   "留观"
               ImageKey        =   "Keep"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "OutKeep"
                     Object.Tag             =   "门诊留观登记"
                     Text            =   "门诊留观登记"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "InKeep"
                     Object.Tag             =   "住院留观登记"
                     Text            =   "住院留观登记"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Keep_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预约"
               Key             =   "PreAdd"
               Description     =   "预约"
               Object.ToolTipText     =   "预约入院登记"
               Object.Tag             =   "预约"
               ImageKey        =   "PreAdd"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "接收"
               Key             =   "Confirm"
               Description     =   "接收"
               Object.ToolTipText     =   "预约入院接收"
               Object.Tag             =   "接收"
               ImageKey        =   "Confirm"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm0"
                     Text            =   "接收为住院病人"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm1"
                     Text            =   "接收为门诊留观"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm2"
                     Text            =   "接收为住院留观"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Confirm_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改当前入院登记记录"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "Del"
               Description     =   "取消"
               Object.ToolTipText     =   "取消当前入院登记记录"
               Object.Tag             =   "取消"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前入院登记记录"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "设置条件过滤满足条件的病人"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满点条件的病人上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "家属"
               Key             =   "Family"
               Description     =   "家属"
               Object.ToolTipText     =   "家属登记"
               Object.Tag             =   "家属"
               ImageKey        =   "Family"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyAdd"
                     Text            =   "家属登记"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyView"
                     Text            =   "家属信息"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FamilySplit"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvwDist_s 
      Height          =   4290
      Left            =   0
      TabIndex        =   1
      Top             =   1575
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   7567
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   420
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":0F01
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":111B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1335
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":154F
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1769
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1983
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":20FD
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2317
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2531
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":274B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2965
            Key             =   "Keep"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2B7F
            Key             =   "PreAdd"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":3279
            Key             =   "Confirm"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":3973
            Key             =   "Family"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1005
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":A1D5
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A3EF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A609
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A823
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":AA3D
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":AC57
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B3D1
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B5EB
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B805
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BA1F
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BC39
            Key             =   "Keep"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BE53
            Key             =   "PreAdd"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":C54D
            Key             =   "Confirm"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":CC47
            Key             =   "Family"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1650
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":134A9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   1065
      Left            =   30
      TabIndex        =   0
      Top             =   1200
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1879
      TabWidthStyle   =   2
      TabFixedWidth   =   2646
      TabFixedHeight  =   563
      HotTracking     =   -1  'True
      TabMinWidth     =   1235
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "确认登记(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "预约登记(&2)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   4530
      Left            =   3870
      TabIndex        =   2
      Top             =   1290
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7990
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmManageHosReg.frx":13603
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblNode 
      AutoSize        =   -1  'True
      Caption         =   "站点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   9
      Top             =   900
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_PrintMed 
         Caption         =   "打印病案(&M)"
      End
      Begin VB.Menu mnuFile_PrintWristlet 
         Caption         =   "打印腕带(&W)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "收费轧帐(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInsure 
         Caption         =   "保险类别(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "病人入院登记(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditOutKeep 
         Caption         =   "门诊留观登记(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditInKeep 
         Caption         =   "住院留观登记(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreAdd 
         Caption         =   "预约入院登记(&P)"
      End
      Begin VB.Menu mnuEditConfirm 
         Caption         =   "预约入院接收(&C)"
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "接收为住院病人(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "接收为门诊留观(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "接收为住院留观(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改登记(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "取消登记(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditToKeep 
         Caption         =   "撤为留观(&K)"
      End
      Begin VB.Menu mnuEditToIn 
         Caption         =   "住院留观转为住院(&P)"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅登记(&V)"
      End
      Begin VB.Menu mnuEdit_Surety 
         Caption         =   "担保信息(&B)"
      End
      Begin VB.Menu mnuEdit_Family 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_FamilyAdd 
         Caption         =   "家属登记"
      End
      Begin VB.Menu mnuEdit_FamilyView 
         Caption         =   "家属信息"
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
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolDist 
            Caption         =   "病人分布(&D)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInBed 
         Caption         =   "显示入住病人(&I)"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "显示病人方式(&M)"
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按病区显示(&U)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按科室显示(&D)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColmunSet 
         Caption         =   "自定义显示列(&C)"
      End
      Begin VB.Menu mnuView_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageHosReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsPati As ADODB.Recordset
Private mblnMax As Boolean, mblnUnload As Boolean
Private mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mintBedLen As Integer '床号最大长度
Private mcllFilterA As Collection
Private mblnPassShowCard As Boolean '卡号是否密文显示
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限
Private mstrHead As String
Private Enum PATIVSF_COLUMN
    COL_病人性质 = 0
    COL_登记类型 = 1
    COL_病人ID = 2
    COL_门诊号 = 3
    COL_住院号 = 4
    COL_留观号 = 5
    COL_就诊卡 = 6
    COL_床号 = 7
    COL_姓名 = 8
    COL_性别 = 9
    COL_年龄 = 10
    COL_费别 = 11
    COL_医疗付款方式 = 12
    COL_医保号 = 13
    COL_险类 = 14
    COL_入院时间 = 15
    COL_入院病区 = 16
    
    COL_入院科室 = 17
    COL_护理等级 = 18
    COL_次数 = 19
    Col_入院病况 = 20
    COL_入院方式 = 21
    COL_住院目的 = 22
    COL_出生日期 = 23
    COL_国籍 = 24
    COL_民族 = 25
    COL_学历 = 26
    COL_职业 = 27
    COL_身份 = 28
    COL_身份证号 = 29
    COL_手机号 = 30
    COL_婚姻 = 31
    COL_工作单位 = 32
    COL_家庭地址 = 33
    COL_家庭电话 = 34
    COL_门诊诊断 = 35
    COL_备注 = 36
    COL_登记员 = 37
    COL_状态 = 38
    COL_主页ID = 39
    COL_病人类型 = 40
End Enum

'by lesfeng 2010-1-11 性能优化
Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化过滤条件
    '入参:
    '出参:
    '返回:
    '编制:lesfeng
    '日期:2010-01-11 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilterA = New Collection
    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "入院日期"
    mcllFilterA.Add Array("", ""), "住院号"
    '问题17122 by lesfeng 2010-02-02
    mcllFilterA.Add "", "病人姓名"
    mcllFilterA.Add "", "登记人"
    mcllFilterA.Add "", "门诊号"
    mstrFilter = ""
End Sub

Private Sub cboNodeList_Click()
    Call InitUnits
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub chkOnly_Click()
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub mnuEdit_FamilyAdd_Click()
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, 0, 2, mlngModul) '编辑
End Sub

Private Sub mnuEdit_FamilyView_Click()
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID"))) Then
            MsgBox "没有客户信息可以查看家属信息！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
            MsgBox "没有病人信息可以查看家属信息！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, lng病人ID, 1, mlngModul) '查看
End Sub

Private Sub mnuEdit_Surety_Click()
    '56964:刘鹏飞,2013-04-23
    Dim lng病人ID As Long, lngRow As Long
    Dim bln在院病人 As Boolean
    
    lngRow = mshPati.Row
    
    If lngRow >= mshPati.FixedRows And lngRow < mshPati.Rows Then
        lng病人ID = Val(mshPati.TextMatrix(lngRow, GetColNum("病人ID")))
    Else
        lng病人ID = 0
    End If

    frmSurety.mlng病人ID = lng病人ID
    frmSurety.mbln在院病人 = True
    frmSurety.mstrPrivs = mstrPrivs
    frmSurety.Show 1, Me
End Sub

Private Sub mnuEditConfirmType_Click(Index As Integer)
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
        MsgBox "没有预约登记可以接收。", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 2 '接收预约
    frmHosReg.mbytKind = Index '0-正常预约,1-门诊留观,2-住院留观
    frmHosReg.mbytInState = 0 '定为新增
    frmHosReg.mlng病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    frmHosReg.mlng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditInKeep_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 2
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditOutKeep_Click()
    On Error Resume Next
    Err.Clear
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 1
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditPreAdd_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 1 '预约登记
    frmHosReg.mbytKind = 0 '不提供留观的预约
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 2 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditToIn_Click()
'将住院留观病人转为住院病人
    Dim lng病人ID As Long, lng主页ID As Long, intRow As Long
    Dim str住院号 As String, str姓名 As String
    Dim strSQL As String, strNote As String
    Dim lng性质 As Long
    Dim rsTemp As New ADODB.Recordset
    
    intRow = mshPati.Row
    lng性质 = GetColNum("病人性质")
    If Val(mshPati.TextMatrix(intRow, lng性质)) <> 2 Then Exit Sub
        
        
    lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("病人ID")))
    lng主页ID = Val(mshPati.TextMatrix(intRow, GetColNum("主页ID")))
    str住院号 = mshPati.TextMatrix(intRow, GetColNum("住院号"))
            
    strSQL = "Select Nvl(状态,0) 状态 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And 病人性质=2"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        If rsTemp!状态 = 1 Then
            MsgBox "病人当前尚未入科,不能转为住院病人。请先将病人入科后再试。", vbInformation, gstrSysName
            Exit Sub
        ElseIf rsTemp!状态 = 2 Then
            MsgBox "病人当前正在转科,不能转为住院病人。请先将病人转科或取消转科后再试。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("确实要将该住院留观病人转为住院病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '60500:刘鹏飞,2013-05-09,留观登记没有确定住院号，转为住院病人如果是使用统一住院号，应该保持和之前住院一致
    If str住院号 = "" And gbln每次住院新住院号 = False Then
        strSQL = " SELECT Nvl(a.住院号," & vbNewLine & _
            "            (SELECT 住院号" & vbNewLine & _
            "             FROM 病案主页" & vbNewLine & _
            "             WHERE 病人id = a.病人id AND" & vbNewLine & _
            "                   主页id = (SELECT MAX(主页id) FROM 病案主页 WHERE 病人id = a.病人id AND 住院号 IS NOT NULL))) 住院号" & vbNewLine & _
            " FROM 病人信息 a" & vbNewLine & _
            " WHERE 病人id = [1]"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
        If Not rsTemp.EOF Then
            str住院号 = NVL(rsTemp!住院号)
        End If
    End If
    '没有住院号则分配一个
    If str住院号 = "" Or gbln每次住院新住院号 Then
        str住院号 = zlDatabase.GetNextNo(2)
        str姓名 = mshPati.TextMatrix(intRow, GetColNum("姓名"))
        strNote = "在留观病人 " & str姓名 & " 转为住院病人之前，请先为该病人确定一个住院号。"
        If Not frmInput.InputVal(Me, "住院号", strNote, str住院号, 1, 10, False, InStr(mstrPrivs, ";修改住院号;") <> 0) Then Exit Sub
    End If
        
    strSQL = "ZL_病人变动记录_转住院(" & lng病人ID & "," & lng主页ID & "," & str住院号 & ")"
    On Error GoTo errH
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    '行直接处理
    mshPati.TextMatrix(intRow, lng性质) = "0"
    mshPati.TextMatrix(intRow, GetColNum("登记类型")) = "住院病人"
    mshPati.TextMatrix(intRow, GetColNum("住院号")) = str住院号
    
    Call mshPati_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnuEditToKeep_Click()
'将住院病人撤消为住院留观病人
    Dim intRow As Integer, i As Integer
    Dim lng病人ID As Long, lng主页ID As Long, int清除住院号 As Integer
    Dim strSQL As String
    
    intRow = mshPati.Row
    
    If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("病人ID"))) Then
        MsgBox "没有病人可以撤为留观病人！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("病人ID")))
    lng主页ID = Val(mshPati.TextMatrix(intRow, GetColNum("主页ID")))
            
    '去掉了医保连接匹配检查
    
    If MsgBox("确实要将病人""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """撤消为住院留观病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If lng主页ID = 1 Then
        If MsgBox("同时清除该病人的住院号吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then int清除住院号 = 1
    End If
    strSQL = "zl_入院病案主页_DELETE(" & lng病人ID & "," & lng主页ID & ",1," & int清除住院号 & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    '行直接处理
    mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质")) = "2"
    mshPati.TextMatrix(mshPati.Row, GetColNum("登记类型")) = "住院留观"
    If int清除住院号 = 1 Then mshPati.TextMatrix(mshPati.Row, GetColNum("住院号")) = ""
    
    Call mshPati_EnterCell
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFile_PrintMed_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    lng病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    lng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Private Sub mnuFile_PrintWristlet_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    lng病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    lng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetPar.mlngModul = mlngModul
    frmSetPar.mstrPrivs = mstrPrivs
    frmSetPar.Show 1, Me
End Sub

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strTmp As String, str病人ID As String
    
    str病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    If Left(tvwDist_s.Tag, 1) = "U" Then
        strTmp = "病区="
    Else    '未选择时,当成科室
        strTmp = "病人科室="
    End If
    
    If str病人ID <> "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            strTmp & Mid(tvwDist_s.Tag, 2), _
            "病人ID=" & str病人ID, _
            "住院号=" & mshPati.TextMatrix(mshPati.Row, GetColNum("住院号")))
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            strTmp & Mid(tvwDist_s.Tag, 2))
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    Call tbsType_Click
End Sub

Private Sub mnuViewColmunSet_Click()
    Call frmColumnSet.ShowMe(Me, mshPati, mstrHead)

End Sub

Private Sub mnuViewFilter_Click()
    frmHosRegFilter.Show 1, Me
    If gblnOK Then
        mstrFilter = frmHosRegFilter.mstrFilter
        'by lesfeng 2010-1-11 性能优化
        Set mcllFilterA = frmHosRegFilter.mcllFilter
        If mcllFilterA("门诊号") <> "" Then tvwDist_s.Nodes(1).root.Selected = True
        InitNode
        mnuViewreFlash_Click
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmHosRegFind.Show 1, Me
    If gblnOK Then Call SeekPati(frmHosRegFind.optHead)
End Sub

Private Sub mnuViewInBed_Click()
    mnuViewInBed.Checked = Not mnuViewInBed.Checked
    Call ShowPatis(mstrFilter)
End Sub

Private Sub mnuViewToolDist_Click()
    mnuViewToolDist.Checked = Not mnuViewToolDist.Checked
    tbsType.Visible = mnuViewToolDist.Checked
    tvwDist_s.Visible = mnuViewToolDist.Checked
    pic.Visible = tvwDist_s.Visible
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub mshPati_Click()
    If tbsType.SelectedItem.Index = 1 Then Exit Sub
    If mshPati.RowSel = 0 Then Exit Sub
    If (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "住院病人" And InStr(mstrPrivs, "接收住院预约") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "门诊留观" And InStr(mstrPrivs, "接收门诊留观预约") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "住院留观" And InStr(mstrPrivs, "接收住院留观预约") = 0) Then
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
End Sub

Private Sub mshPati_DblClick()
    If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("病人ID")) = "" Then Exit Sub
    mnuEdit_View_Click
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    ElseIf Button = 1 Then
        mblnDown = True
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnload Then
        Unload Me
    Else
        Call InitLocPar(mlngModul)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekPati(False)
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, Curdate As Date
    Dim lngTmp As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    '收费轧帐模块权限
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    On Error GoTo errHandle
    mstrHead = "病人性质,1,0|登记类型,1,1150|病人ID,1,1050|门诊号,1,1050|住院号,1,1050|留观号,1,1050|就诊卡,1,1150|床号,1,800|姓名,1,1100|性别,1,800|年龄,1,800|费别,1,800|医疗付款方式,1,1500|" & _
            "医保号,1,1300|险类,1,1800|入院时间,1,1300|入院病区,1,1850|入院科室,1,1850|护理等级,1,1150|次数,4,800|" & _
            "入院病况,1,1150|入院方式,1,1150|住院目的,1,1150|出生日期,1,1300|" & _
            "国籍,1,800|民族,1,1300|学历,1,800|职业,1,1300|身份,1,1050|身份证号,1,2300|手机号,1,1500|婚姻,1,800|" & _
            "工作单位,1,2300|家庭地址,1,2300|家庭电话,1,1500|门诊诊断,1,4300|备注,1,2300|登记员,1,1050|状态,1,0|主页ID,1,0|病人类型,1,1300|挂号ID,1,0|升级,1,0"
    
    '80509:刘鹏飞,2014-12-09,添加病人查找、过滤
    If Not gobjSquare Is Nothing Then Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
    
    strSQL = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mblnPassShowCard = NVL(rsTemp!卡号密文) <> ""
    End If
    'by lesfeng 2010-1-11 性能优化
    Call InitFilter
    
    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    
    '恢复个性病人清单类型
    RestoreWinState Me, App.ProductName
    mnuViewInBed.Checked = zlDatabase.GetPara("显示入住病人", glngSys, mlngModul, "0")
    '刷新方式
    lngTmp = zlDatabase.GetPara("刷新方式", glngSys, mlngModul, "1")
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = (i = lngTmp)
    Next
    lngTmp = zlDatabase.GetPara("显示病人方式", glngSys, mlngModul, "0")
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
    
    mblnUnload = False
    
    '权限设置
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '初始化站点列表
    Call InitNode
    
    '正常登记
    If InStr(mstrPrivs, ";办理登记;") = 0 Then '包含了门诊留观，住院留观
        mnuEdit_Add.Visible = False
        mnuEditInKeep.Visible = False
        mnuEditOutKeep.Visible = False
        mnuEdit_1.Visible = False
        
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Keep").Visible = False
        tbr.Buttons("Keep_").Visible = False
    End If
            
    '预约和接收
    If InStr(mstrPrivs, ";预约登记;") = 0 Then '不提供留观病人预约登记
        mnuEditPreAdd.Visible = False
        tbr.Buttons("PreAdd").Visible = False
    End If
    If InStr(mstrPrivs, ";接收预约;") = 0 Then '包含了门诊留观，住院留观
        mnuEditConfirm.Visible = False
        tbr.Buttons("Confirm").Visible = False
    Else
        If InStr(mstrPrivs, ";接收住院预约;") = 0 And InStr(mstrPrivs, ";接收门诊留观预约;") = 0 And InStr(mstrPrivs, ";接收住院留观预约;") = 0 Then
            mnuEditConfirm.Enabled = False
            tbr.Buttons("Confirm").Enabled = False
        Else
            If InStr(mstrPrivs, ";接收住院预约;") = 0 Then
                mnuEditConfirmType(0).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(1).Visible = False
            End If
            
            If InStr(mstrPrivs, ";接收门诊留观预约;") = 0 Then
                mnuEditConfirmType(1).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(2).Visible = False
            End If
            
            If InStr(mstrPrivs, ";接收住院留观预约;") = 0 Then
                mnuEditConfirmType(2).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(3).Visible = False
            End If
        End If
    End If
    
    If InStr(mstrPrivs, ";预约登记;") = 0 _
        And InStr(mstrPrivs, ";接收预约;") = 0 Then
        mnuEditPreAdd.Visible = False
        mnuEditConfirm.Visible = False
        mnuEdit_2.Visible = False
        tbr.Buttons("PreAdd").Visible = False
        tbr.Buttons("Confirm").Visible = False
        tbr.Buttons("Confirm_").Visible = False
    End If
                            
    '留观病人子权限:正常登记和预约登记的
    If InStr(mstrPrivs, ";住院病人登记;") = 0 And InStr(mstrPrivs, ";门诊留观登记;") = 0 And InStr(mstrPrivs, ";住院留观登记;") = 0 Then
        mnuEdit_Add.Visible = False
        tbr.Buttons("Add").Visible = False
        mnuEditOutKeep.Visible = False
        tbr.Buttons("Keep").ButtonMenus("OutKeep").Visible = False
        mnuEditInKeep.Visible = False
        tbr.Buttons("Keep").ButtonMenus("InKeep").Visible = False
        mnuEditConfirm.Visible = False
        tbr.Buttons("Confirm").Visible = False
    Else
        If InStr(mstrPrivs, ";住院病人登记;") = 0 Then
            mnuEdit_Add.Visible = False
            tbr.Buttons("Add").Visible = False
            mnuEditConfirmType(0).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm0").Visible = False
        End If
        If InStr(mstrPrivs, ";门诊留观登记;") = 0 Then
            mnuEditOutKeep.Visible = False
            tbr.Buttons("Keep").ButtonMenus("OutKeep").Visible = False
            mnuEditConfirmType(1).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm1").Visible = False
        End If
        If InStr(mstrPrivs, ";住院留观登记;") = 0 Then
            mnuEditInKeep.Visible = False
            tbr.Buttons("Keep").ButtonMenus("InKeep").Visible = False
            mnuEditConfirmType(2).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm2").Visible = False
        End If
    End If
    If InStr(mstrPrivs, ";门诊留观登记;") = 0 _
        And InStr(mstrPrivs, ";住院留观登记;") = 0 Then
        mnuEdit_1.Visible = False
        tbr.Buttons("Keep").Visible = False
        tbr.Buttons("Keep_").Visible = False
    End If
                        
    '修改权限
    If InStr(mstrPrivs, ";办理登记;") = 0 _
        And InStr(mstrPrivs, ";预约登记;") = 0 _
        And InStr(mstrPrivs, ";接收预约;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
                        
    '包含了取消入院,取消预约的功能
    If InStr(mstrPrivs, ";取消入院;") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEditToKeep.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    
    '住院留观转住院
    If InStr(mstrPrivs, ";住院留观转住院;") = 0 Then
        mnuEditToIn.Visible = False
    End If
    Call tbsType_Click
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    Dim DisW As Long '病人分布表宽度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshPati.MousePointer = 0
    
    mshPati.Redraw = False
    
    If mblnMax Then
        tvwDist_s.width = 3780
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    DisW = IIf(tvwDist_s.Visible, tvwDist_s.width + pic.width, 0)
    
    pic.Visible = tvwDist_s.Visible
    lblNode.Visible = cboNodeList.Visible
    
    cboNodeList.Top = Me.ScaleTop + cbrH + 15
    lblNode.Top = cboNodeList.Top
    If cboNodeList.Height - lblNode.Height > 0 Then
        lblNode.Top = lblNode.Top + (cboNodeList.Height - lblNode.Height) \ 2
    End If
    
    With tbsType
        .Top = Me.ScaleTop + cbrH + 15 + IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
        .Left = Me.ScaleLeft + 30
        .width = tvwDist_s.width - 45
    End With
    With tvwDist_s
        .Left = Me.ScaleLeft
        .Top = tbsType.Top + 330
        .Height = Me.ScaleHeight - staH - .Top
    End With
    With pic
        .Left = tvwDist_s.Left + tvwDist_s.width
        .Top = tvwDist_s.Top
        .Height = tvwDist_s.Height
    End With
    With picFind
        .Left = DisW
        .Top = Me.ScaleTop + cbrH
    End With
    With mshPati
        .Left = DisW
        .Top = picFind.Top + picFind.Height ' Me.ScaleTop + cbrH
        .Height = Me.ScaleHeight - cbrH - staH - picFind.Height
        .width = Me.ScaleWidth - DisW
    End With
    cboNodeList.width = tvwDist_s.width - 600
    mshPati.Redraw = True
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, lngTmp As Long
    
    mstrFilter = ""
    
    SaveWinState Me, App.ProductName
    zlDatabase.SetPara "显示入住病人", mnuViewInBed.Checked, glngSys, mlngModul
    
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "刷新方式", lngTmp, glngSys, mlngModul
    
    '显示病人方式
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "显示病人方式", lngTmp, glngSys, mlngModul
    
    
    Unload frmHosRegFind
    Unload frmHosRegFilter
End Sub

Private Sub mnuEdit_Del_Click()
    Dim intRow As Integer, i As Integer
    Dim lng病人ID As Long, lng主页ID As Long, lng挂号ID As Long
    Dim strSQL As String, int险类 As Integer
    Dim rsTmp As ADODB.Recordset
    Dim blnNotCommit As Boolean
    Dim blnTrans As Boolean
    intRow = mshPati.Row
    
    If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("病人ID"))) Then
        MsgBox "没有病人可以取消登记。", vbExclamation, gstrSysName: Exit Sub
    End If
    
    lng病人ID = mshPati.TextMatrix(intRow, GetColNum("病人ID"))
    lng主页ID = Val(mshPati.TextMatrix(intRow, GetColNum("主页ID")))
    lng挂号ID = Val(mshPati.TextMatrix(intRow, GetColNum("挂号ID")))
    '去掉了医保连接匹配检查
    
    If MsgBox("确实要取消病人""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """的登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '问题:31635
    blnNotCommit = False
    int险类 = 0
    On Error GoTo errH
    '问题22073 by lesfeng 2010-08-02  验证是否书写电子病历
    If GetCaseHistory(lng病人ID, lng主页ID) Then
        MsgBox "已经对病人""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """书写电子病历，不能取消入院！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    Set rsTmp = GetMoneyInfo(lng病人ID, , , 2)
    If Not rsTmp Is Nothing Then
        If NVL(rsTmp!预交余额) <> 0 Then '可能没有预交但有费用余额
            If MsgBox("病人""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """有预交款未退，是否要继续办理取消入院？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
        
    If tbsType.SelectedItem.Index = 1 Then
        '医保(取消入院:实际上某些医保是执行出院交易)
        If isYBPati(lng病人ID, , int险类) Then
            If Not gclsInsure.ComeInDelSwap(lng病人ID, lng主页ID, int险类) Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Sub
            End If
        End If
        '问题:31635
        blnNotCommit = True
        
        strSQL = "zl_住院一次费用_Delete(" & lng病人ID & "," & lng主页ID & ")"
       
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    strSQL = "zl_入院病案主页_DELETE(" & lng病人ID & "," & lng主页ID & ",0," & IIf(gbln每次住院新住院号, "1", "0") & ")" '"主页ID=0"表示预约登记
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    gcnOracle.CommitTrans: blnTrans = False
    
    If lng主页ID = 0 Then
        '更新预约系统接口{"挂号id_In": "挂号ID","状态_In": "状态" ---已接收，未接收，已退出}
        Call Sys.NewSystemSvr("预约中心", "入住或入住取消", "{""挂号id_In"": """ & lng挂号ID & """,""状态_In"": ""已退出""}", "")
    End If
     '问题:31635
    If int险类 > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ComeInDelSwap, True, int险类)
    
    On Error GoTo 0
    
    '行直接处理
    If mshPati.Rows > 2 Then
        mshPati.RemoveItem intRow
        Call SetMenu(True)
    Else
        With mshPati
            For i = 0 To .Cols - 1
                .TextMatrix(intRow, i) = ""
            Next
        End With
        Call SetMenu(False)
    End If
    
    If intRow <= mshPati.Rows - 1 Then
        mshPati.Row = intRow
    Else
        mshPati.Row = mshPati.Rows - 1
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    
    Call mshPati_EnterCell
    Call mshPati_Click
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
     '问题:31635
    If int险类 > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ComeInDelSwap, False, int险类)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Modi_Click()
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
        MsgBox "没有病人信息可以修改！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = tbsType.SelectedItem.Index - 1 '正常或预约
    frmHosReg.mbytKind = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质")))
    frmHosReg.mbytInState = 1
    frmHosReg.mlng病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    frmHosReg.mlng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 0
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_View_Click()
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
        MsgBox "没有病人信息可以查看！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = tbsType.SelectedItem.Index - 1 '正常或预约
    frmHosReg.mbytKind = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质")))
    frmHosReg.mbytInState = 2
    frmHosReg.mlng病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    frmHosReg.mlng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    mshPati.Refresh
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewreFlash_Click()
    Call tbsType_Click
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshPati_RowColChange()
    If tbsType.SelectedItem.Index = 1 Then Exit Sub
    If mshPati.RowSel = 0 Then Exit Sub
    If (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "住院病人" And InStr(mstrPrivs, "接收住院预约") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "门诊留观" And InStr(mstrPrivs, "接收门诊留观预约") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("登记类型")) = "住院留观" And InStr(mstrPrivs, "接收住院留观预约") = 0) Then
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim strTag As String
    
    blnCancel = False
    If objHisPati Is Nothing Then blnCancel = True
    If blnCancel = False Then
        If objHisPati.病人ID = 0 Then blnCancel = True
    End If
    
    If tbsType.SelectedItem.Index = 1 Then
        strTag = "在院病人"
    Else
        strTag = "预约病人"
    End If
            
    If blnCancel Then
        MsgBox "没有找到符合条件的病人，请确认要查找的病人是否属于" & strTag & "！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '开始从现有列表中寻找病人
    If FindPatiInfo(objHisPati.病人ID, objHisPati) = True Then Exit Sub
    '提取病人信息数据
    If GetPatiInfo(objHisPati.病人ID, objHisPati) = False Then blnCancel = True: Exit Sub
    If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, strName As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim strTag As String
    Dim lng病人ID As Long
    
    strName = Trim(PatiIdentify.Text)
    
    On Error GoTo ErrHand
    blnCancel = False
    If Not tvwDist_s.SelectedItem Is Nothing And mnuViewByDept(0).Checked = True Then
        PatiIdentify.病人病区ID = Val(Mid(tvwDist_s.SelectedItem.Key, 2))
    Else
        PatiIdentify.病人病区ID = 0
    End If
            
    If objCard.名称 Like "*姓*名*" And blnCard = False And strName <> "" And InStr("-*+/", Left(Trim(PatiIdentify.Text), 1)) = 0 Then
       
        If gblnSeekName = False Then '允许姓名模糊查找
            MsgBox "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。", vbInformation, gstrSysName
            blnCancel = True
            Exit Sub
        End If
        
        If tbsType.SelectedItem.Index = 1 Then
            strIF = " And A.主页ID=C.主页ID And Nvl(C.主页ID,0)<>0"
            strTag = "在院病人"
        Else
            strIF = " And Nvl(C.主页ID,0)=0"
            strTag = "预约病人"
        End If
    
        
        strPati = "Select 1 As 排序id, a.病人id As Id, a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.住院次数, Trunc(c.入院日期, 'dd') As 入院日期, a.出生日期," & vbNewLine & _
            "       a.身份证号, a.家庭地址, a.工作单位, c.病人类型" & vbNewLine & _
            " From 病人信息 a, 病案主页 c" & vbNewLine & _
            " Where a.停用时间 Is Null And a.病人id = c.病人id " & strIF & " And c.出院日期 Is Null  And a.姓名 Like [1] " & _
            IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])") & " And Rownum < 101"
        strPati = strPati & " Order by 排序ID,姓名,入院日期 Desc"
        
        vRect = zlControl.GetControlRect(PatiIdentify.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strName & "%", gintNameDays)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '以病人ID读取
                lng病人ID = NVL(rsTmp!ID)
            End If
        Else '取消选择
            If blnCancel = False Then
                MsgBox "没有找到符合条件的病人，请确认要查找的病人是否属于" & strTag & "！", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If
        
        '开始从现有列表中寻找病人
        If FindPatiInfo(lng病人ID, objCardData) = True Then blnFindPatied = True: blnCancel = True: Exit Sub
        '提取病人信息数据
        If GetPatiInfo(lng病人ID, objCardData) = False Then
            blnCancel = True: Exit Sub
        Else
            blnFindPatied = True: blnCancel = True
        End If
        If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FindPatiInfo(ByVal lngPatiID As Long, objCardData As zlIDKind.PatiInfor) As Boolean
'功能:根据病人ID,定位到病人
    Dim i As Long, blnFind As Boolean
    
    If lngPatiID = 0 Then Exit Function
    
    For i = 1 To mshPati.Rows - 1
        If Val(mshPati.TextMatrix(i, GetColNum("病人ID"))) = lngPatiID Then
            blnFind = True
            Exit For
        End If
    Next i
    
    If objCardData Is Nothing Then
        Set objCardData = New zlIDKind.PatiInfor
    End If
    
    If blnFind = True Then
        mshPati.Row = i: mshPati.TopRow = i
        mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
        objCardData.病人ID = lngPatiID: objCardData.姓名 = mshPati.TextMatrix(i, GetColNum("姓名"))
        FindPatiInfo = True
    End If
End Function

Private Function GetPatiInfo(ByVal lngPatiID As Long, objCardData As zlIDKind.PatiInfor) As Boolean
    Dim i As Long, strSQL As String
    Dim strCard As String, strIF As String
    Dim strNodeNo As String
    Dim rsPati As ADODB.Recordset
    Dim strTag As String
    
    On Error GoTo errH
    '获取站点号
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
    
    strIF = " And A.病人ID=[1]"
    If tbsType.SelectedItem.Index = 1 Then
        strIF = strIF & " And A.主页ID=B.主页ID And Nvl(B.主页ID,0)<>0"
        strTag = "在院病人"
    Else
        strIF = strIF & " And Nvl(B.主页ID,0)=0"
        strTag = "预约病人"
    End If
    
    If mblnPassShowCard = True Then
        strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡,"
    Else
        strCard = "A.就诊卡号 as 就诊卡,"
    End If

    
    mintBedLen = GetMaxBedLen

    strSQL = _
        "Select 病人性质,Decode(B.病人性质,1,'门诊留观',2,'住院留观','住院病人') as 登记类型," & _
        " A.病人ID, A.门诊号, B.住院号,B.留观号," & strCard & "Decode(Nvl(B.状态,0),1,NULL,LPad(B.出院病床," & mintBedLen & ", ' ')) as 床号," & _
        " NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.费别,Nvl(B.医疗付款方式,A.医疗付款方式) 医疗付款方式,Nvl(A.医保号,F.信息值) as 医保号,X.名称 as 险类," & _
        " To_Char(B.入院日期,'YYYY-MM-DD HH24:MI:SS') as 入院时间," & _
        " C.名称 as 入院病区,D.名称 as 入院科室,E.名称 as 护理等级,A.住院次数 as 次数,B.入院病况," & _
        " B.入院方式,B.住院目的,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族," & _
        " A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.婚姻状况 as 婚姻,A.工作单位,A.家庭地址, A.家庭电话, g.门诊诊断, B.备注,B.登记人 as 登记员,B.状态,B.主页ID," & _
        " Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,收费项目目录 E,病案主页从表 F," & _
        "       (Select distinct  病人ID, first_value(记录) OVER (PARTITION BY 病人ID ORDER BY 记录日期 DESC) AS 门诊诊断" & vbNewLine & _
        "       From (SELECT a.病人id,q.诊断描述 记录,q.记录日期" & vbNewLine & _
        "               FROM 病人信息 a, 病案主页 b, 病人挂号记录 p, 病人诊断记录 q" & vbNewLine & _
        "               Where a.病人id = b.病人id AND B.出院日期 is NULL AND b.病人id=p.病人id(+) And p.病人id = q.病人id(+)" & vbNewLine & _
        "               AND p.Id = q.主页id AND p.记录性质=1 and p.记录状态=1 and 记录来源(+) = 3 AND 诊断类型(+) = 1" & vbNewLine & _
        "               AND 诊断次序(+) = 1 " & strIF & ")) g,保险类别 X" & _
        " Where A.病人ID=B.病人ID And B.出院日期 is NULL And B.入院病区ID=C.ID(+)" & _
        " And B.入院科室ID=D.ID " & IIf(cboNodeList.Visible, "And (d.站点=" & strNodeNo & " Or d.站点 Is Null)", "") & " And B.护理等级ID=E.ID(+)" & _
        " And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And B.病人ID = G.病人ID(+)" & _
        " And F.信息名(+)='医保号' And B.险类=X.序号(+)" & strIF
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID)
    
    If rsPati.EOF Then
        MsgBox "没有找到符合条件的病人，请确认要查找的病人是否属于" & strTag & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not tvwDist_s.SelectedItem Is Nothing Then
            For i = 1 To tvwDist_s.Nodes.Count
                If tvwDist_s.Nodes(i).Selected = True Then
                    tvwDist_s.Nodes(i).Selected = False
                End If
            Next i
            tvwDist_s.Tag = ""
            Set tvwDist_s.SelectedItem = Nothing
    End If
    
    mshPati.Clear
    mshPati.ClearStructure
    mshPati.Rows = 2
    
    Set mshPati.DataSource = rsPati
    Call setHeader(mstrHead)          '在其中的enter_cell中已调用SetMenu(false)
    If mnuViewInBed.Checked Then Call SetInBed
    stbThis.Panels(2) = "共 " & rsPati.RecordCount & " 个病人"
    Call SetMenu(True)
    
    mshPati_Click
    
    If objCardData Is Nothing Then Set objCardData = New zlIDKind.PatiInfor
    objCardData.病人ID = lngPatiID: objCardData.姓名 = mshPati.TextMatrix(mshPati.Row, GetColNum("姓名"))
    
    Me.Refresh
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic.Left + X < 1000 Or mshPati.width - X < 2000 Or mshPati.width - X < picFind.width Then Exit Sub
        pic.Left = pic.Left + X
        tbsType.width = tbsType.width + X
        tvwDist_s.width = tvwDist_s.width + X
        picFind.Left = picFind.Left + X
        mshPati.Left = mshPati.Left + X
        mshPati.width = mshPati.width - X
        cboNodeList.width = tvwDist_s.width - 600
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go"
            mnuViewGo_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Add"
            mnuEdit_Add_Click
        Case "Keep"
            If Not Button.ButtonMenus("OutKeep").Visible _
                And Button.ButtonMenus("InKeep").Visible Then
                mnuEditInKeep_Click
            ElseIf Not Button.ButtonMenus("InKeep").Visible _
                And Button.ButtonMenus("OutKeep").Visible Then
                mnuEditOutKeep_Click
            End If
        Case "PreAdd"
            mnuEditPreAdd_Click
        Case "Confirm"
            '根据病人性质决定缺省是那种接收
            Select Case Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质")))
                Case 0
                    If mnuEditConfirmType(0).Enabled And mnuEditConfirmType(0).Visible Then
                        Call mnuEditConfirmType_Click(0)
                    End If
                Case 1
                    If mnuEditConfirmType(1).Enabled And mnuEditConfirmType(1).Visible Then
                        Call mnuEditConfirmType_Click(1)
                    End If
                Case 2
                    If mnuEditConfirmType(2).Enabled And mnuEditConfirmType(2).Visible Then
                        Call mnuEditConfirmType_Click(2)
                    End If
            End Select
        Case "View"
            mnuEdit_View_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Family"
           Call mnuEdit_FamilyAdd_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "OutKeep"
            mnuEditOutKeep_Click
        Case "InKeep"
            mnuEditInKeep_Click
        Case "Confirm0"
            Call mnuEditConfirmType_Click(0)
        Case "Confirm1"
            Call mnuEditConfirmType_Click(1)
        Case "Confirm2"
            Call mnuEditConfirmType_Click(2)
        Case "FamilyAdd"
            Call mnuEdit_FamilyAdd_Click
        Case "FamilyView"
            Call mnuEdit_FamilyView_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tbsType_Click()
    mnuViewInBed.Enabled = tbsType.SelectedItem.Index = 1
    cbr.Bands(2).Visible = tbsType.SelectedItem.Index = 2
    If mnuEdit_Surety.Visible Then mnuEdit_Surety.Enabled = tbsType.SelectedItem.Index = 1
    
    Call InitUnits
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub tvwDist_s_NodeClick(ByVal Node As MSComctlLib.Node)
    '相同点击不再处理
    If tvwDist_s.Tag = Node.Key Then Exit Sub
    tvwDist_s.Tag = Node.Key
    
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshPati.Row
    
    '表头
    objOut.Title.Text = "入院病人清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "部门：" & tvwDist_s.SelectedItem.Text
    objRow.Add "时间：" & Format(frmHosRegFilter.dtp入院B.Value, "yyyy-MM-dd") & " 至 " & Format(frmHosRegFilter.dtp入院E.Value, "yyyy-MM-dd")
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshPati.Redraw = False
    Set objOut.Body = mshPati
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshPati.Row = intRow
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = True
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled Then
        mnuEdit_Del_Click
    ElseIf KeyCode = vbKeyReturn And mnuEdit_View.Enabled Then
        mnuEdit_View_Click
    End If
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub InitNode()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngUnitID As Long
    Dim blnByDept As Boolean
    
    On Error GoTo errHandle
    blnByDept = mnuViewByDept(1).Checked
    
    '加载站点选项
    strSQL = "Select Distinct 站点, c.名称" & vbNewLine & _
            " From (Select Distinct " & IIf(blnByDept, "入院科室id", "入院病区id") & " ID" & vbNewLine & _
            "       From 病案主页" & vbNewLine & _
            "       Where 入院日期 Between [1] And [2] And " & IIf(blnByDept, "入院科室id", "入院病区id") & " Is Not Null) A, 部门表 B, zlnodelist C" & vbNewLine & _
            " Where A.ID = B.ID And B.站点=c.编号 " & vbNewLine & _
            " Order By 站点"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, frmHosRegFilter.dtp入院B, frmHosRegFilter.dtp入院E)
    cboNodeList.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNodeList.AddItem rsTmp!站点 & "-" & rsTmp!名称
            cboNodeList.ItemData(rsTmp.AbsolutePosition - 1) = rsTmp!站点
            rsTmp.MoveNext
        Wend
        Call cbo.Locate(cboNodeList, gstrNodeNo, True)
    Else
        lblNode.Visible = False
        cboNodeList.Visible = False
        Form_Resize
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'功能：初始化病人病区科室分布列表
'说明：以病区-科室分层,所有病区、科室在当前在院病人之中获得
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, i As Integer, lngUnitID As Long
    Dim strPreKey  As String, blnByDept As Boolean
    Dim strNodeNo As String
    Dim strDeptIDs As String
      
    strPreKey = ""
    If Not tvwDist_s.SelectedItem Is Nothing Then strPreKey = tvwDist_s.SelectedItem.Key
    blnByDept = mnuViewByDept(1).Checked
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
        
    tvwDist_s.Nodes.Clear
    Set objNode = tvwDist_s.Nodes.Add(, , "Root", IIf(blnByDept, "所有科室", "所有病区"), 1)
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    
    If tbsType.SelectedItem.Index = 2 And InStr(mstrPrivs, ";全院预约;") = 0 Then
        strDeptIDs = GetDeptOrUnitByUser()
    End If
    
    Set rsTmp = GetInDept(blnByDept, frmHosRegFilter.dtp入院B, frmHosRegFilter.dtp入院E, strNodeNo, strDeptIDs)
    If Not rsTmp.EOF Then
        If blnByDept Then
            lngUnitID = UserInfo.部门ID
        Else
            lngUnitID = Get病区ID(UserInfo.部门ID)
        End If
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvwDist_s.Nodes.Add("Root", tvwChild, IIf(blnByDept, "D", "U") & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, 1)
            If objNode.Key = strPreKey Then objNode.Selected = True
            If rsTmp!ID = lngUnitID And tvwDist_s.SelectedItem Is Nothing Then objNode.Selected = True
            
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
    End If
    If tvwDist_s.SelectedItem Is Nothing Then
        tvwDist_s.Nodes(IIf(tvwDist_s.Nodes.Count > 1, 2, 1)).Selected = True
    End If
        
    InitUnits = True
End Function

Private Sub setHeader(ByVal strHead As String)
    Dim i As Integer, j As Integer
    Dim strWidth As String, strText As String
    Dim arrText As Variant
    
    'gclsBase.GetRegister(私有模块, Me.Name, strPath & "_" & TypeName(vsf(0)) & "_20101228", "")





    With mshPati
        .Redraw = False
        
        
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '恢复列顺序
            '检查是否需要恢复
            strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.EXEName & "\" & Me.Name & "\" & TypeName(mshPati), mshPati.Name & mshPati.Tag & "名称", "")
            arrText = Split(strText, ",")
            
            If strText <> "" Then
                .Cols = UBound(arrText) + 1
                For i = 0 To UBound(arrText)
                    .TextMatrix(0, i) = arrText(i)
                    .ColAlignmentFixed(i) = 4
                    For j = 0 To UBound(Split(strHead, "|"))
                        If (arrText(i) = Split(Split(strHead, "|")(j), ",")(0)) Then
                            .colAlignment(i) = Split(Split(strHead, "|")(j), ",")(1)
                            Exit For
                        End If
                    Next
                Next
            End If

            strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.EXEName & "\" & Me.Name & "\" & TypeName(mshPati), mshPati.Name & mshPati.Tag & "宽度", "")
            If UBound(Split(strWidth, ",")) >= .Cols - 1 Then
                For i = 0 To .Cols - 1
                    .ColWidth(i) = Split(strWidth, ",")(i)
                Next
            End If
        End If
        
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Or .Cols = 0 Or .Cols <> UBound(Split(strHead, "|")) + 1 Then
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
                If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
                .ColAlignmentFixed(i) = 4
            Next
        End If
        
        If Not Visible Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        .ColWidth(0) = 0
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        Call mshPati_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub mshPati_EnterCell()
    If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")) = "" Then Exit Sub
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
    
    Call SetMenu(mnuFile_Print.Enabled)
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = 0
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '双击最大化时会执行
        mblnDown = False
        
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshPati.TextMatrix(1, GetColNum("病人ID")) = "" Then Exit Sub
        
        Set mshPati.DataSource = Nothing

        mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True)
        mshPati_Click
    End If
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    Dim i As Integer, blnHavePrivs As Boolean
    Dim lng性质 As Long
    i = GetColNum("状态")
    lng性质 = GetColNum("病人性质")
    
    '根据权限
    mnuEdit_Modi.Visible = True
    tbr.Buttons("Modi").Visible = True
    If InStr(mstrPrivs, "办理登记") = 0 _
        And InStr(mstrPrivs, "预约登记") = 0 _
        And InStr(mstrPrivs, "接收预约") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    Else
        If tbsType.SelectedItem.Index = 1 Then
            If InStr(mstrPrivs, "办理登记") = 0 Then
                mnuEdit_Modi.Visible = False
                tbr.Buttons("Modi").Visible = False
            End If
        Else
            If InStr(mstrPrivs, "预约登记") = 0 Then
                mnuEdit_Modi.Visible = False
                tbr.Buttons("Modi").Visible = False
            End If
        End If
    End If
            
    '根据可操作性
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    mnuFile_PrintMed.Enabled = blnUsed
    mnuFile_PrintWristlet.Enabled = blnUsed
    
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    If Val(mshPati.TextMatrix(mshPati.Row, i)) = 1 Then
        '刚登记病人
        mnuEdit_Modi.Enabled = blnUsed
        tbr.Buttons("Modi").Enabled = blnUsed
        tbr.Buttons("Del").Enabled = blnUsed
        mnuEdit_Del.Enabled = blnUsed
        
        If tbsType.SelectedItem.Index = 1 Then
            mnuEditConfirm.Enabled = False
            tbr.Buttons("Confirm").Enabled = False
            mnuEditToKeep.Enabled = blnUsed And Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质"))) = 0
        Else
            mnuEditConfirm.Enabled = blnUsed
            tbr.Buttons("Confirm").Enabled = blnUsed
            mnuEditToKeep.Enabled = False
        End If
    Else
        '已入住病人
        mnuEdit_Modi.Enabled = False
        tbr.Buttons("Modi").Enabled = False
        tbr.Buttons("Del").Enabled = False
        mnuEdit_Del.Enabled = False
        mnuEditToKeep.Enabled = False
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
    
    tbr.Buttons("View").Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
    
    '住院留观病人转为住院病人
    mnuEditToIn.Enabled = (Val(mshPati.TextMatrix(mshPati.Row, lng性质)) = 2)
    
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    'tbr.Buttons("轧帐").Visible = blnHavePrivs
    'tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    '担保信息
    mnuEdit_Surety.Visible = InStr(mstrPrivs, ";担保信息;") > 0
    '病人家属
    blnHavePrivs = InStr(";" & GetPrivFunc(glngSys, 9003) & ";", ";病人家属;") > 0
    mnuEdit_Family.Visible = blnHavePrivs
    mnuEdit_FamilyAdd.Visible = blnHavePrivs
    mnuEdit_FamilyView.Visible = blnHavePrivs
    mnuEdit_FamilyView.Enabled = blnUsed And blnHavePrivs
    
    tbr.Buttons("FamilySplit").Visible = blnHavePrivs
    tbr.Buttons("Family").Visible = blnHavePrivs
    tbr.Buttons("Family").ButtonMenus.Item("FamilyView").Enabled = blnUsed And blnHavePrivs
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的病人,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmHosRegFind
            If .txt病人ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
            End If
            If .txt就诊卡.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("就诊卡")) = .txt就诊卡.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If .txt床号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("床号")) Like "*" & .txt床号.Text & "*"
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean)
'功能：根据当前菜单浏览要求(自动生成条件),读取病人信息
'参数：strIF=" And ...."形式的过滤条件
    Dim i As Long, strSQL As String, strDiagnoseSQL As String
    Dim strCard As String, strUnit As String
    Dim blnByDept As Boolean, lngDeptID As Long
    Dim Curdate As Date
    Dim strNodeNo As String
    Dim intDiagDays As Integer
    Dim strPerson As String, strParTable As String, strTable As String, strDiag As String
    Dim varArr As Variant
    Dim rsDiag As New ADODB.Recordset
    Dim j As Integer
    
    'by lesfeng 2010-1-11 性能优化
    On Error GoTo errH
    
    If blnSort = False Then PatiIdentify.Text = ""
    strPerson = ""
    strDiag = ""
    
    '获取站点号
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
    
    If Not blnSort Then
        blnByDept = mnuViewByDept(1).Checked
        
        If strIF = "" Then
            '设置初始条件(本月内入院)
            'strIF = " And B.入院日期 Between trunc(Sysdate,'mm') And Sysdate"
'            strIF = " AND B.入院日期 Between trunc(Sysdate-7) and Sysdate "
            Curdate = zlDatabase.Currentdate
            strIF = ""
            strIF = strIF & " And (B.入院日期  Between [1] And [2]) "
            mcllFilterA.Remove "入院日期"
            mcllFilterA.Add Array(Format(DateAdd("d", -7, Curdate), "yyyy-mm-dd") & " 00:00:00", Format(Curdate, "yyyy-mm-dd") & " 23:59:59"), "入院日期"
        End If
        If tbsType.SelectedItem.Index = 1 Then
            strIF = strIF & " And A.主页ID=B.主页ID And Nvl(B.主页ID,0)<>0"
        Else
            strIF = strIF & " And Nvl(B.主页ID,0)=0"
        End If
        
        '就诊卡号显示
        '55849:刘鹏飞,2012-11-21,将原有Decode判断的方式改为固定提取字段,
        '因为Decode第一个变量使用常量从指标中提取字段数据，可能导致导致查不出结果，或者返回的记录集访问出现E-FAIL错误，估计是ADO和Oracle兼容性的Bug，在特定的Decode和子表查询同时使用时会出现，但没有明确的规律。
        'strCard = "Decode(" & IIf(mblnPassShowCard, 0, 1) & ",1,A.就诊卡号,LPAD('*',Length(A.就诊卡号),'*')) as 就诊卡,"
        If mblnPassShowCard = True Then
            strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡,"
        Else
            strCard = "A.就诊卡号 as 就诊卡,"
        End If
        '当前病区或科室
        If Not tvwDist_s.SelectedItem Is Nothing Then  '如果任何科室或病区没有人,则只显所有病区
            lngDeptID = Val(Mid(tvwDist_s.SelectedItem.Key, 2))
            If blnByDept Then
                If lngDeptID <> 0 Then strUnit = " And B.入院科室ID=[6]"
            Else
                If lngDeptID <> 0 Then strUnit = " And B.入院病区ID=[6]"
            End If
        End If
        
        mintBedLen = GetMaxBedLen(lngDeptID)
        '54179:刘鹏飞,2012-10-12,修改提取病人诊断的sql，诊断显示最后一次门诊诊断（以前为病人历史所有门诊诊断）
        If tbsType.SelectedItem.Index = 1 Then
            If Not (mnuViewInBed.Checked And mnuViewInBed.Enabled) Then
                '等待入科的病人(状态=1)；床号要设置为" ",不然全部为待入病人时排序会出错
                strSQL = _
                    "Select 病人性质,Decode(B.病人性质,1,'门诊留观',2,'住院留观','住院病人') as 登记类型," & _
                    " A.病人ID, A.门诊号, B.住院号,B.留观号," & strCard & "' ' as 床号," & _
                    " NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.费别,Nvl(B.医疗付款方式,A.医疗付款方式) 医疗付款方式,Nvl(A.医保号,F.信息值) as 医保号,X.名称 as 险类," & _
                    " To_Char(B.入院日期,'YYYY-MM-DD HH24:MI:SS') as 入院时间," & _
                    " C.名称 as 入院病区,D.名称 as 入院科室,E.名称 as 护理等级,A.住院次数 as 次数,B.入院病况," & _
                    " B.入院方式,B.住院目的,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族," & _
                    " A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.婚姻状况 as 婚姻,A.工作单位,A.家庭地址,  A.家庭电话, Decode(g.诊断描述, Null, '',g.诊断描述) As 门诊诊断, B.备注,B.登记人 as 登记员,B.状态,B.主页ID," & _
                    " Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型,B.挂号ID " & _
                    " From 病人信息 A,病案主页 B,部门表 C,部门表 D,收费项目目录 E,病案主页从表 F,病人诊断记录 G,保险类别 X" & _
                    " Where A.病人ID=B.病人ID And B.状态 = 1 And B.入院病区ID=C.ID(+)" & _
                    " And B.入院科室ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.站点=" & strNodeNo & " Or d.站点 Is Null)", "") & " And B.护理等级ID=E.ID(+)" & _
                    " And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And B.病人ID = G.病人ID(+) And B.主页ID=G.主页ID(+) And g.记录来源(+) = 2 And g.诊断类型(+) = 1 And g.诊断次序(+) = 1 " & _
                    " And F.信息名(+)='医保号' And B.险类=X.序号(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From 病案主页 Where 病人ID=a.病人id And 主页ID>0 And 病人性质=1 And 入院时间 Is Not Null And 出院时间 Is Null)", "") & _
                    " Order by 入院时间 Desc,住院号 Desc"
            Else
                strSQL = _
                    "Select 病人性质,Decode(B.病人性质,1,'门诊留观',2,'住院留观','住院病人') as 登记类型," & _
                    " A.病人ID, A.门诊号, B.住院号,B.留观号," & strCard & "Decode(Nvl(B.状态,0),1,NULL,LPad(B.出院病床," & mintBedLen & ", ' ')) as 床号," & _
                    " NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.费别,Nvl(B.医疗付款方式,A.医疗付款方式) 医疗付款方式,Nvl(A.医保号,F.信息值) as 医保号,X.名称 as 险类," & _
                    " To_Char(B.入院日期,'YYYY-MM-DD HH24:MI:SS') as 入院时间," & _
                    " C.名称 as 入院病区,D.名称 as 入院科室,E.名称 as 护理等级,A.住院次数 as 次数,B.入院病况," & _
                    " B.入院方式,B.住院目的,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族," & _
                    " A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.婚姻状况 as 婚姻,A.工作单位,A.家庭地址, A.家庭电话, Decode(g.诊断描述, Null, '',g.诊断描述) As 门诊诊断, B.备注,B.登记人 as 登记员,B.状态,B.主页ID," & _
                    " Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型,B.挂号ID" & _
                    " From 病人信息 A,病案主页 B,部门表 C,部门表 D,收费项目目录 E,病案主页从表 F,病人诊断记录 G,保险类别 X" & _
                    " Where A.病人ID=B.病人ID And B.出院日期 is NULL And B.入院病区ID=C.ID(+)" & _
                    " And B.入院科室ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.站点=" & strNodeNo & " Or d.站点 Is Null)", "") & " And B.护理等级ID=E.ID(+)" & _
                    " And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And B.病人ID = G.病人ID(+) And B.主页ID=G.主页ID(+) And g.记录来源(+) = 2 And g.诊断类型(+) = 1 And g.诊断次序(+) = 1 " & _
                    " And F.信息名(+)='医保号' And B.险类=X.序号(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From 病案主页 Where 病人ID=a.病人id And 主页ID>0 And 病人性质=1 And 入院时间 Is Not Null And 出院时间 Is Null)", "") & _
                    " Order by 入院时间 Desc,住院号 Desc"
            End If
        Else
            '查询预约登记病人
            strSQL = _
                    "Select 病人性质,Decode(B.病人性质,1,'门诊留观',2,'住院留观','住院病人') as 登记类型," & _
                    " A.病人ID, A.门诊号, B.住院号,B.留观号," & strCard & "LPad(B.出院病床," & mintBedLen & ", ' ') as 床号," & _
                    " NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.费别,Nvl(B.医疗付款方式,A.医疗付款方式) 医疗付款方式,Nvl(A.医保号,F.信息值) as 医保号,X.名称 as 险类," & _
                    " To_Char(B.入院日期,'YYYY-MM-DD HH24:MI:SS') as 入院时间," & _
                    " C.名称 as 入院病区,D.名称 as 入院科室,E.名称 as 护理等级,A.住院次数 as 次数,B.入院病况," & _
                    " B.入院方式,B.住院目的,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族," & _
                    " A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.婚姻状况 as 婚姻,A.工作单位,A.家庭地址,  A.家庭电话, Null As 门诊诊断, B.备注,B.登记人 as 登记员,B.状态,B.主页ID," & _
                    " Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型,B.挂号ID" & _
                    " From 病人信息 A,病案主页 B,部门表 C,部门表 D,收费项目目录 E,病案主页从表 F,保险类别 X" & _
                    " Where A.病人ID=B.病人ID And B.状态 = 1 And B.入院病区ID=C.ID(+)" & _
                    " And B.入院科室ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.站点=" & strNodeNo & " Or d.站点 Is Null)", "") & " And B.护理等级ID=E.ID(+)" & _
                    " And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) " & _
                    " And F.信息名(+)='医保号' And B.险类=X.序号(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From 病案主页 Where 病人ID=a.病人id And 主页ID>0 And 病人性质=1 And 入院时间 Is Not Null And 出院时间 Is Null)", "") & _
                    " Order by 入院时间 Desc,住院号 Desc"
        End If
        If Not tvwDist_s.SelectedItem Is Nothing Then
            tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
        End If
        
        Call zlCommFun.ShowFlash("正在读取病人清单,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        '问题17122 by lesfeng 2010-02-02
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mcllFilterA("入院日期")(0)), CDate(mcllFilterA("入院日期")(1)), _
        CLng(Val(mcllFilterA("住院号")(0))), CLng(Val(mcllFilterA("住院号")(1))), CStr(mcllFilterA("登记人")), lngDeptID, gstrLike & CStr(mcllFilterA("病人姓名")) & "%", mcllFilterA("门诊号"))
'        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
        If Not mrsPati.EOF And tbsType.SelectedItem.Index = 2 Then
            For i = 0 To mrsPati.RecordCount - 1
                strPerson = strPerson & "," & mrsPati!病人ID
                mrsPati.MoveNext
            Next
            mrsPati.MoveFirst
            strPerson = Mid(strPerson, 2)
            intDiagDays = Val(zlDatabase.GetPara("诊断查找天数", glngSys, glngModul, "3"))
            strParTable = "Select /* +cardinality(a,10) */" & "Column_Value From Table(f_num2List([1]))"
            strTable = strParTable
            
            If Len(strPerson) >= 4000 Then
                varArr = Array()
                varArr = GetParTable(strPerson, strParTable, strTable)
            End If
            strSQL = "Select a.病人id, a.诊断描述, 1 As 序号" & vbNewLine & _
                "From 病案主页 G, 病人诊断记录 A, 病人诊断医嘱 B, 病人医嘱记录 C, 诊疗项目目录 D, 病人挂号记录 E" & vbNewLine & _
                "Where g.病人id = a.病人id And a.Id = b.诊断id And b.医嘱id = c.Id And c.诊疗项目id + 0 = d.Id And a.病人id=e.病人id And a.主页id = e.Id And a.记录来源 = 3 And" & vbNewLine & _
                "      e.记录性质 = 1 And e.记录状态 = 1 And e.登记时间 + 0 > Trunc(Sysdate-" & intDiagDays & ") And c.医嘱状态 In (3, 8) And d.类别 = 'Z' And" & vbNewLine & _
                "      Instr(',1,11,', ',' || a.诊断类型 || ',') > 0 And Instr(',1,2,', d.操作类型) > 0 And g.入院科室id = c.执行科室id And" & vbNewLine & _
                "      Nvl(g.主页id, 0) = 0 And G.病人id in (" & strTable & "A)" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.病人id, a.诊断描述, 2 As 序号" & vbNewLine & _
                "From 病人诊断记录 A, 病人挂号记录 B" & vbNewLine & _
                "Where a.病人id = b.病人id And a.主页id = b.Id And b.记录性质 = 1 And b.记录状态 = 1 And b.登记时间 + 0 > Trunc(Sysdate-" & intDiagDays & ") And" & vbNewLine & _
                "      Instr(',1,11,', ',' || a.诊断类型 || ',') > 0 And a.记录来源 = 3 And b.病人id in (" & strTable & "A)"

            If Len(strPerson) >= 4000 Then
                Set rsDiag = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
            Else
                Set rsDiag = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", strPerson)
            End If
        End If
    End If
    
    mshPati.Clear
    mshPati.ClearStructure
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call setHeader(mstrHead)
        stbThis.Panels(2).Text = "当前设置没有过滤出任何病人"
        Call SetMenu(False)
    Else
        Set mshPati.DataSource = mrsPati
         If tbsType.SelectedItem.Index = 2 And mshPati.Rows > 0 Then
	 If Not rsDiag Is Nothing Then
            If rsDiag.RecordCount > 0 Then
                For i = 1 To mshPati.Rows - 1
                    rsDiag.Filter = "病人ID=" & Val(mshPati.TextMatrix(i, COL_病人ID)) & " And 序号=1"
                    If rsDiag.RecordCount > 0 Then
                        For j = 0 To rsDiag.RecordCount - 1
                            strDiag = strDiag & "," & rsDiag!诊断描述
                            rsDiag.MoveNext
                        Next
                        mshPati.TextMatrix(i, COL_门诊诊断) = Mid(strDiag, 2)
                    Else
                        rsDiag.Filter = "病人ID=" & Val(mshPati.TextMatrix(i, COL_病人ID)) & " And 序号=2"
                         If rsDiag.RecordCount > 0 Then
                            For j = 0 To rsDiag.RecordCount - 1
                                strDiag = strDiag & "," & rsDiag!诊断描述
                                rsDiag.MoveNext
                            Next
                            mshPati.TextMatrix(i, COL_门诊诊断) = Mid(strDiag, 2)
                        End If
                    End If
                    strDiag = ""
                Next
            End If
        End If
	End If
        Call setHeader(mstrHead)          '在其中的enter_cell中已调用SetMenu(false)
        If mnuViewInBed.Checked Then Call SetInBed
        stbThis.Panels(2) = "共 " & mrsPati.RecordCount & " 个病人"
        Call SetMenu(True)
    End If
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub
Public Function GetParTable(ByVal strpar As String, ByVal strParTable As String, ByRef strTableOut As String) As Variant
'功能：对于动态内存表的绑定参数超长的处理
'参数：strPar 参数串，strParTable 内存表形式要传入
'返回：一个字符串数组，10个元素
    Dim n As Long, p As Long
    Dim varPar(0 To 9) As String
    Dim strTable As String, strThis As String
    Dim intNum As Integer '参数号
    
    For n = 0 To 9
        varPar(n) = ""
    Next
    
    p = InStr(strParTable, "[") + 1
    intNum = Mid(strParTable, p, 1)
    
    n = 0
    Do While True
        If Len(strpar) < 4000 Then
            p = Len(strpar) + 1
        Else
            p = InStrRev(Mid(strpar, 1, 4000), ",")
        End If
        
        strThis = Mid(strpar, 1, p - 1)
        
        If n > 9 Then
            strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            If n = 0 Then
                strTable = strParTable
            Else
                strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "[" & (n + intNum) & "]")
            End If
        End If
        
        n = n + 1
        
        strpar = Mid(strpar, p + 1)
        
        If strpar = "" Then Exit Do
    Loop
    
    strTableOut = strTable
    GetParTable = varPar
    
End Function
Private Sub SetInBed()
    Dim i As Integer, j As Integer, k As Integer
    Dim bln As Boolean
    Dim intRow As Integer, intCol As Integer
    
    intRow = mshPati.Row
    bln = mshPati.Redraw
    mshPati.Redraw = False
        
    j = GetColNum("状态")
    k = GetColNum("床号")
    For i = 1 To mshPati.Rows - 1
        '床号不为空的(包括家庭病床)为入住病人
        If Val(mshPati.TextMatrix(i, j)) <> 1 Then
            mshPati.Row = i: mshPati.Col = k
            mshPati.CellBackColor = &HEBFFFF
        End If
    Next
    mshPati.Row = intRow: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = bln
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
'问题22073 by lesfeng 2010-08-02  验证是否书写电子病历
Private Function GetCaseHistory(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：获取指定病人是否存在电子病历记录
'说明：用于获取病人电子病历记录的记录情况
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim int记录数 As Integer
    
    GetCaseHistory = False
    On Error GoTo errH
    
    strSQL = "Select count(病人id) As 计数 From 电子病历记录 " & _
             " Where 病人ID = [1] And 主页ID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!计数) Then
            int记录数 = rsTmp!计数
            If int记录数 > 0 Then GetCaseHistory = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


