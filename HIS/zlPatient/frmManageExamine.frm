VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageExamine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "病人费用审批"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   540
   ClientWidth     =   11040
   Icon            =   "frmManageExamine.frx":0000
   KeyPreview      =   -1  'True
   Picture         =   "frmManageExamine.frx":1601A
   ScaleHeight     =   6480
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TabStrip tbsClass 
      Height          =   340
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   609
      TabFixedWidth   =   2290
      TabFixedHeight  =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "全部(&0)"
            Key             =   "全部"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "手术"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11040
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6195
      MinHeight1      =   720
      Width1          =   4500
      NewRow1         =   0   'False
      Caption2        =   "病人科室"
      Child2          =   "cboDept"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
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
            NumButtons      =   12
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
               Key             =   "Line_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "编辑"
               Key             =   "Edit"
               Description     =   "审批项目编辑"
               Object.ToolTipText     =   "编辑"
               Object.Tag             =   "编辑"
               ImageKey        =   "Billing"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Edit"
                     Object.Tag             =   "编辑"
                     Text            =   "编辑"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤出院病人条件"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Object.ToolTipText     =   "定位病人审批项目"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "Style"
               Description     =   "查看"
               Object.ToolTipText     =   "查看"
               Object.Tag             =   "查看"
               ImageKey        =   "Style"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "大图标"
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "小图标"
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "列表"
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "详细资料"
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8955
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   2670
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5145
      ScaleWidth      =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   1260
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8070
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
      BackColor       =   16777215
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "姓名"
         Text            =   "姓名"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "病员号"
         Text            =   "住院号"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "性别"
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "年龄"
         Text            =   "年龄"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   "入院日期"
         Text            =   "入院日期"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "出院日期"
         Text            =   "出院日期"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "当前科室"
         Text            =   "当前科室"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Key             =   "住院"
         Text            =   "住院"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "险类"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   90
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
            Picture         =   "frmManageExamine.frx":161A8
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":163C2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":165DC
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":167F6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":16F70
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1718A
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":173A4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":175BE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":177D8
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":179F2
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":180EC
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":187E6
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":18EE0
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":190FA
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   90
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
            Picture         =   "frmManageExamine.frx":19314
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1952E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":19748
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":19962
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A0DC
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A2F6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A510
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A72A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A944
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1AB5E
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1B258
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1B952
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1C04C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1C266
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   975
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   635
      TabFixedWidth   =   2290
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "当前在院(&1)"
            Key             =   "InHos"
            Object.ToolTipText     =   "当前在院的病人"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "出院病人(&2)"
            Key             =   "OutHos"
            Object.ToolTipText     =   "期间内出院的病人"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3105
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1C480
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1CD5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3690
      Top             =   90
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
            Picture         =   "frmManageExamine.frx":1D634
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1DF0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   6120
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageExamine.frx":1E7E8
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
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
   Begin VSFlex8Ctl.VSFlexGrid vsExist 
      Height          =   4515
      Left            =   2760
      TabIndex        =   10
      Top             =   1300
      Width           =   8265
      _cx             =   14579
      _cy             =   7964
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
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmManageExamine.frx":1F07A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VB.Label lblMoney 
      BackColor       =   &H00808080&
      Caption         =   " 当前病人需审批的收费项目"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   2775
      TabIndex        =   7
      Top             =   765
      Width           =   6990
   End
   Begin VB.Label lbl_s 
      BackColor       =   &H00808080&
      Caption         =   " 时间:2001-01-01至2001-01-01"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   30
      TabIndex        =   6
      ToolTipText     =   "在该时间范围内的住院病人"
      Top             =   750
      Width           =   2580
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
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_EditItem 
         Caption         =   "审批项目管理(&A)"
      End
      Begin VB.Menu mnuEdit_split 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_EditTemplet 
         Caption         =   "项目模板管理(&B)"
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
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "住院病区(&U)"
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
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤出院病人(&T)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "定位审批项目(&C)"
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位病人(&G)"
      End
      Begin VB.Menu mnuView_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "大图标(&I)"
         Index           =   0
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFindPati 
         Caption         =   "未设置审批项目病人(&W)"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "显示病人方式(&K)"
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按病区显示(&U)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按科室显示(&U)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuView_8 
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
Attribute VB_Name = "frmManageExamine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mlngCurRow As Long, mlngTopRow As Long
Public mdtBegin As Date, mdtEnd As Date

Private mblnFirst As Boolean

Private mlngDeptID As Long

Private mstrPrePati As String

Private mstrPage As String

Public mstrPrivs  As String

Private mintBedLen  As Integer
Private mblnUnLoad As Boolean
Private mrsPati As New ADODB.Recordset
Public mrsExistItem As New ADODB.Recordset

Private Enum ColNum
    类别 = 0: 编码: 名称: 规格: 产地: 单位: 说明: 审批人: 审批时间
End Enum
Private Sub cboDept_Click()
    If cboDept.ItemData(cboDept.ListIndex) = mlngDeptID Then Exit Sub
    mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    
    mstrPage = ""
    vsExist.Rows = 1
    If Visible Then Call tbs_Click
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnLoad = True Then Unload Me: Exit Sub

End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    lbl_s.Left = Me.ScaleLeft
    lbl_s.Top = Me.ScaleTop + cbrH + 45
    lbl_s.Width = pic.Left
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = lbl_s.Top + lbl_s.Height + 45
    tbs.Width = lbl_s.Width
    
    lvw.Left = Me.ScaleLeft
    lvw.Top = tbs.Top + tbs.Height - 75
    lvw.Width = lbl_s.Width
    lvw.Height = Me.ScaleHeight - staH - cbrH - lbl_s.Height - tbs.Height - 15
    
    pic.Top = Me.ScaleTop + cbrH
    pic.Height = Me.ScaleHeight - cbrH - staH
    
    lblMoney.Left = pic.Left + pic.Width
    lblMoney.Top = Me.ScaleTop + cbrH + 45
    lblMoney.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    
    If tbsClass.Visible = True Then
        tbsClass.Left = pic.Left + pic.Width
        tbsClass.Top = lblMoney.Top + lblMoney.Height
        tbsClass.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
        
        vsExist.Left = tbsClass.Left
        vsExist.Top = tbsClass.Top + tbsClass.Height - 50
        vsExist.Width = tbsClass.Width
        vsExist.Height = pic.Height - tbsClass.Height - lblMoney.Height + 50
        vsExist.ZOrder
    Else
         vsExist.Left = pic.Left + pic.Width
         vsExist.Top = lblMoney.Top + lblMoney.Height
         vsExist.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
          vsExist.Height = pic.Height - lblMoney.Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_DblClick()
    If mnuEdit_EditItem.Visible = True And tbs.SelectedItem.Index = 1 Then
        Call mnuEdit_EditItem_Click
    End If
End Sub

Private Sub mnuEdit_EditItem_Click()
    Dim lng病人ID As Long, lng主页ID As Long, lng险类 As Long
    Dim arrTmp As Variant
    
    If tbs.SelectedItem.Index = 2 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    arrTmp = Split(Mid(lvw.SelectedItem.Key, 2), "_")
    lng病人ID = Val(arrTmp(0))
    lng主页ID = Val(arrTmp(1))
    lng险类 = Val(arrTmp(2))

    If InStr(mstrPrivs, "增加审批项目") > 0 Then
        Call frmExamineEdit.ExamineEdit(lng病人ID, lng主页ID, lng险类, False, False)
    ElseIf InStr(mstrPrivs, "删除审批项目") > 0 Then
        Call frmExamineEdit.ExamineEdit(lng病人ID, lng主页ID, lng险类, True, False)
    End If
    If lvw.SelectedItem Is Nothing Then Exit Sub
    mstrPrePati = ""
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub mnuEdit_EditTemplet_Click()
    frmExamineEdit.ExamineEdit 0, 0, 0, False, True
End Sub

Private Sub mnuFile_Excel_Click()
    zlRptPrint 3
End Sub

Private Sub mnuFile_PreView_Click()
    zlRptPrint 2
End Sub

Private Sub mnuFile_Print_Click()
    zlRptPrint 1
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    mlngDeptID = 0
    Call InitUnits
End Sub

Private Sub mnuViewFilter_Click()
    frmSetExamine.EditWhere Me
    If tbs.SelectedItem Is Nothing Then Exit Sub
    If tbs.SelectedItem.Index = 2 Then
        Call LoadPatients
    End If
End Sub


Private Sub mnuViewFind_Click()
    Dim lngRow As Long
    Dim strOld As String
    
    If vsExist.Rows = 1 Then Exit Sub
    If tbsClass.Visible = True Then
        For lngRow = 1 To tbsClass.Tabs.Count
            frmExamineFind.cbo类别.AddItem tbsClass.Tabs.Item(lngRow).Key, lngRow - 1
        Next lngRow
        strOld = tbsClass.SelectedItem.Key
    Else
        frmExamineFind.cbo类别.AddItem vsExist.TextMatrix(1, ColNum.类别), 0
    End If
    If frmExamineFind.cbo类别.ListCount > 0 Then frmExamineFind.cbo类别.ListIndex = 0
    Set frmExamineFind.mrsfind = mrsExistItem
    
    frmExamineFind.Show 1, Me
End Sub



Private Sub FindPati()
    Dim strFilter As String
    Dim strBed As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    With frmDeptGo
        If .txt住院号.Text <> "" Then strFilter = strFilter & " Or 住院号=" & .txt住院号.Text
        If .txt姓名.Text <> "" Then strFilter = strFilter & " Or 姓名 Like '%" & .txt姓名.Text & "%'"
        If .txt床号.Text <> "" Then
            strBed = .txt床号.Text
            If mintBedLen - zlCommFun.ActualLen(strBed) > 0 Then
                strBed = String(mintBedLen - zlCommFun.ActualLen(strBed), " ") & strBed
            End If
            strFilter = strFilter & " Or 床号='" & strBed & "'"
        End If
    End With
    If strFilter = "" Then Exit Sub
    mrsPati.Filter = 0
    mrsPati.Filter = Mid(strFilter, 5)
    
    If mrsPati.EOF Then
        stbThis.Panels(2).Text = "没有发现该病人！"
    Else
        lvw.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID & "_" & mrsPati!险类).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuViewFindPati_Click()
    mnuViewFindPati.Checked = Not mnuViewFindPati.Checked
    Call LoadPatients
End Sub

Private Sub mnuViewGo_Click()
    Dim blnPati As Boolean

    frmDeptGo.Show 1, Me
    If gblnOK = True Then Call FindPati
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    mstrPage = ""
    Call tbs_Click
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub



Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lbl_s.Width + X < 2580 Or vsExist.Width - X < 3500 Then Exit Sub
        pic.Left = pic.Left + X
        Call Form_Resize
        Me.Refresh
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lvw.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Edit"
            mnuEdit_EditItem_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Filter"
             mnuViewFilter_Click
        Case "Go"
            mnuViewFind_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub Form_Load()
    Dim i As Long, datTmp As Date
    mblnUnLoad = False
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call RestoreWinState(Me, App.ProductName)
                
    
    '根据保存列表方式设置菜单
    Call SetView(lvw.View)

    mlngCurRow = 1: mlngTopRow = 1
    mblnFirst = True
    
    '权限设置

    
    If InStr(mstrPrivs, "增加审批项目") > 0 Or InStr(mstrPrivs, "删除审批项目") > 0 Then
        mnuEdit_EditItem.Visible = True
        tbr.Buttons.Item("Edit").Visible = True
    Else
        mnuEdit_EditItem.Visible = False
        tbr.Buttons.Item("Edit").Visible = False
        mnuEdit_split.Visible = False
    End If
    
    If InStr(mstrPrivs, "模板管理") > 0 Then
        mnuEdit_EditTemplet.Visible = True
        mnuEdit_split.Visible = mnuEdit_EditItem.Visible
        tbr.Buttons.Item("Line_2").Visible = mnuEdit_EditItem.Visible
    Else
        
        mnuEdit_EditTemplet.Visible = False
        mnuEdit_split.Visible = False
    End If
    
    '科室(操作员所属病区科室)
    If Not InitUnits Then mblnUnLoad = True: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "没有发现你所属科室,且你不具有所有病区权限,不能使用病人费用审批！", vbInformation, gstrSysName
       mblnUnLoad = True: Exit Sub
    End If
        
    datTmp = zlDatabase.Currentdate
    mdtBegin = Format(DateAdd("m", -1, datTmp), "YYYY-MM-DD")
    mdtEnd = Format(datTmp, "YYYY-MM-DD")
    
    Call LoadPatients '其中已包含Call SetDetail Call SetHeader  Call SetMenu
    
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    If tbs.SelectedItem.Index = 2 Then
         mnuEdit_EditItem.Enabled = False
         tbr.Buttons.Item("Edit").Enabled = False
         mnuViewFilter.Enabled = True
         tbr.Buttons.Item("Filter").Enabled = True
    Else
         mnuEdit_EditItem.Enabled = True
         tbr.Buttons.Item("Edit").Enabled = True
         mnuViewFilter.Enabled = False
         tbr.Buttons.Item("Filter").Enabled = False
    End If
    '读取数据
    mstrPage = tbs.SelectedItem.Key
    Call LoadPatients
    lvw.SetFocus
End Sub

Private Sub ReadExistsItem(lng病人ID As Long, lng主页ID As Long, lng险类 As Long)
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
    
    strSQL = " Select C.名称 类别, A.编码, A.名称, B.使用限量, B.已用数量, A.规格, A.产地, A.计算单位, A.说明,B.审批人,B.审批时间" & _
             " From 收费项目目录 A,病人审批项目 B, 收费项目类别 C" & _
             " Where A.类别 = C.编码 And A.ID=B.项目ID And B.病人ID=[1] AND B.主页ID=[2]" & _
             " Order by 类别,编码"
    
    On Error GoTo errHandle
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        With vsExist
            .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = 4
        End With
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!类别 <> strOld Then
            strClass = strClass & "," & mrsExistItem!类别
            strOld = mrsExistItem!类别
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tbsClass.Tabs.Count To 2 Step -1
        tbsClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tbsClass.Visible = True
        tbsClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '用Alt快捷键焦点无法处理
                Set objTab = tbsClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tbsClass.Tabs.Add(, arrClass(i), arrClass(i), 2)
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tbsClass.Visible = False
    End If
'    If vsExist.Tag <> "" Then
'        Set tbsClass.SelectedItem = tbsClass.Tabs.Item(Int(vsExist.Tag))
'        Call tbsClass_Click
'    End If
'    '恢复列顺序:应放在排序处理之前
'    Call RestoreColPosition
'    '排序处理:先排序,以便后面处理行号
'    Call RestoreColSort
    Call Form_Resize
    If vsExist.Rows > 1 Then
        mnuViewFind.Enabled = True
        tbr.Buttons.Item("Go").Enabled = True
        
        If tbsClass.SelectedItem.Index = 1 Then
            mnuEdit_EditItem.Enabled = True
            tbr.Buttons.Item("Edit").Enabled = True
        End If
    Else
        mnuViewFind.Enabled = False
        tbr.Buttons.Item("Go").Enabled = False
        
        If InStr(mstrPrivs, "增加审批项目") = 0 And InStr(mstrPrivs, "删除审批项目") > 0 Then
            mnuEdit_EditItem.Enabled = False
            tbr.Buttons.Item("Edit").Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String, blnByDept As Boolean, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
    blnByDept = mnuViewByDept(1).Checked
    cbr.Bands(2).Caption = IIf(blnByDept, "病人科室", "病人病区")
    
    '有权则显示包括门诊观察室对应的临床科室,住院留观与住院相同
    cboDept.Clear
    If InStr(mstrPrivs, "所有病区") > 0 Then cboDept.AddItem IIf(blnByDept, "所有科室", "所有病区")
    
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    'by lesfeng 2010-03-08 性能优化
    strSQL = _
         " Select A.ID,A.编码,A.名称" & _
         " From 部门表 A,部门性质说明 B" & _
         " Where B.部门ID = A.ID And B.服务对象 IN(1,2,3) And B.工作性质 IN([1])" & _
         " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
         IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
         " And (A.站点=[2] Or A.站点 is Null)" & _
         " Order by A.编码"
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "临床", "护理"), gstrNodeNo)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
            If UserInfo.部门ID = rsTmp!ID Then cboDept.ListIndex = cboDept.NewIndex
            
            rsTmp.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(mstrPrivs, "所有病区") > 0 Then
        MsgBox "没有发现" & IIf(blnByDept, "临床", "护理") & "部门信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetView(bytStyle As Byte)
'功能：调整床位列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub

Private Function LoadPatients() As Boolean
'功能：读取指定范围内的病人列表
    Dim objItem As ListItem, strSQL As String
    Dim i As Long, j As Long, strCount As String
    Dim blnByDept As Boolean
    
    On Error GoTo errH
    
    Call zlCommFun.ShowFlash("正在读取住院病人清单,请稍候 ...", Me)
    DoEvents
    blnByDept = mnuViewByDept(1).Checked
    Me.Refresh
    
    mintBedLen = GetMaxBedLen(mlngDeptID, blnByDept)
    
    If tbs.SelectedItem.Index = 1 Then
        If blnByDept Then
            strSQL = strSQL & IIf(mlngDeptID > 0, " And E.科室ID=[1]", "")
        Else
            strSQL = strSQL & IIf(mlngDeptID > 0, " And E.病区ID=[1]", "")
        End If
    Else
        If blnByDept Then
            strSQL = strSQL & IIf(mlngDeptID > 0, " And B.出院科室ID=[1]", "")
        Else
            strSQL = strSQL & IIf(mlngDeptID > 0, " And B.当前病区ID=[1]", "")
        End If
    End If
    
    If tbs.SelectedItem.Index = 1 Then
        '当前在院的病人
        '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
        If mnuViewFindPati.Checked = False Then
            strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.医疗付款方式," & _
                " B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
                " C.名称 as 当前科室,B.险类,D.名称 医保名称,B.病人性质,B.状态" & _
                " From 病人信息 A,病案主页 B,部门表 C,保险类别 D,在院病人 E" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.出院科室ID=C.ID And A.病人ID=E.病人ID " & strSQL & _
                " And Nvl(B.主页ID,0)<>0 AND B.险类 Is Not Null And B.险类=D.序号" & _
                IIf(mlngDeptID = 0, " Order by A.住院号 Desc", " Order by 床号")
        Else
             strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.医疗付款方式," & _
                " B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
                " C.名称 as 当前科室,B.险类,D.名称 医保名称,B.病人性质,B.状态" & _
                " From 病人信息 A,病案主页 B,部门表 C,保险类别 D,在院病人 E" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.出院科室ID=C.ID And A.病人ID=E.病人ID " & strSQL & _
                " And Nvl(B.主页ID,0)<>0 AND B.险类 Is Not Null And B.险类=D.序号" & _
                " And NOT Exists(Select D.病人ID,D.主页ID from 病人审批项目 D WHERE B.病人ID=D.病人Id and B.主页id=D.主页id)" & _
                IIf(mlngDeptID = 0, " Order by A.住院号 Desc", " Order by 床号")
           
        End If
    ElseIf tbs.SelectedItem.Index = 2 Then
        '该期间出院的病人
        If mnuViewFindPati.Checked = False Then
            strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.医疗付款方式," & _
                " B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
                " C.名称 as 当前科室,B.险类,D.名称 医保名称,B.病人性质,B.状态" & _
                " From 病人信息 A,病案主页 B,部门表 C,保险类别 D" & _
                " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=C.ID" & strSQL & _
                " And B.险类 Is Not Null And B.险类=D.序号 AND B.入院日期<=[3]" & _
                " And B.出院日期 Between [2] And [3]" & _
                IIf(mlngDeptID = 0, " Order by A.住院号 Desc", " Order by 床号")
        Else
            strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.医疗付款方式," & _
                " B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
                " C.名称 as 当前科室,B.险类,D.名称 医保名称,B.病人性质,B.状态" & _
                " From 病人信息 A,病案主页 B,部门表 C,保险类别 D" & _
                " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=C.ID" & strSQL & _
                " And NOT Exists(Select D.病人ID,D.主页ID from 病人审批项目 D WHERE B.病人ID=D.病人Id and B.主页id=D.主页id)" & _
                " And B.险类 Is Not Null And B.险类=D.序号 AND B.入院日期<=[3]" & _
                " And B.出院日期 Between [2] And [3]" & _
                IIf(mlngDeptID = 0, " Order by A.住院号 Desc", " Order by 床号")
        End If
    End If
    
    mdtBegin = CDate(Format(mdtBegin, "yyyy-MM-dd 00:00:00"))
    mdtEnd = CDate(Format(mdtEnd, "yyyy-MM-dd 23:59:59"))
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, mdtBegin, mdtEnd)
  
    lvw.ListItems.Clear
    
    If Not mrsPati.EOF Then
        For i = 1 To mrsPati.RecordCount
            If IIf(IsNull(mrsPati!病人性质), 0, mrsPati!病人性质) = 0 Then
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!病人ID & "_" & mrsPati!主页ID & "_" & mrsPati!险类, mrsPati!姓名, 1, 1)
            Else
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!病人ID & "_" & mrsPati!主页ID & "_" & mrsPati!险类, mrsPati!姓名, 2, 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(mrsPati!住院号), "", mrsPati!住院号)
            objItem.SubItems(2) = IIf(IsNull(mrsPati!床号) And mrsPati!状态 = 0, "家庭", Nvl(mrsPati!床号, " "))
            objItem.SubItems(3) = IIf(IsNull(mrsPati!性别), "", mrsPati!性别)
            objItem.SubItems(4) = IIf(IsNull(mrsPati!年龄), "", mrsPati!年龄)
            objItem.SubItems(5) = Format(mrsPati!入院日期, "yyyy-MM-dd")
            objItem.SubItems(6) = Format(IIf(IsNull(mrsPati!出院日期), "", mrsPati!出院日期), "yyyy-MM-dd")
            objItem.SubItems(7) = IIf(IsNull(mrsPati!当前科室), "", mrsPati!当前科室)
            objItem.SubItems(8) = mrsPati!主页ID
            objItem.SubItems(9) = Nvl(mrsPati!医保名称)
            objItem.Tag = mrsPati!病人ID
            objItem.ListSubItems(1).Tag = Nvl(mrsPati!状态)
            If objItem.Tag = mstrPrePati Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            
            If InStr(strCount & ",", "," & mrsPati!病人ID & ",") = 0 Then strCount = strCount & "," & mrsPati!病人ID
            mrsPati.MoveNext
        Next
        
        lbl_s.Tag = UBound(Split(Mid(strCount, 2), ",")) + 1
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " 当前在院的病人,人数:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " 时间:" & Format(mdtBegin, "yyyy-MM-dd") & "至" & Format(mdtEnd, "yyyy-MM-dd") & ",人数:" & Val(lbl_s.Tag)
        End If
        
        Me.Refresh
        mstrPrePati = ""
    Else
        lbl_s.Tag = ""
        stbThis.Panels(2).Text = ""
        mstrPrePati = ""
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " 当前在院的病人,人数:0"
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " 时间:" & Format(mdtBegin, "yyyy-MM-dd") & "至" & Format(mdtEnd, "yyyy-MM-dd") & ",人数:0"
        End If
    End If
    Call zlCommFun.StopFlash
    
    If lvw.ListItems.Count > 0 Then
        Set lvw.SelectedItem = lvw.ListItems.Item(1)
        If tbs.SelectedItem.Index = 1 Then
            mnuEdit_EditItem.Enabled = True
            tbr.Buttons.Item("Edit").Enabled = True
        End If
        mnuViewGo.Enabled = True
    Else
        mnuEdit_EditItem.Enabled = False
        tbr.Buttons.Item("Edit").Enabled = False
        mnuViewFind.Enabled = False
        tbr.Buttons.Item("Go").Enabled = False
        mnuViewGo.Enabled = False
        vsExist.Rows = 1
        Set mrsExistItem = Nothing
        For i = tbsClass.Tabs.Count To 2 Step -1
            tbsClass.Tabs.Remove i
        Next
    End If
    If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    Exit Function
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call zlCommFun.ShowFlash("正在读取住院病人清单,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxBedLen(Optional lng部门ID As Long, Optional bln科室 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
'      bln占用=是否只管被占用的床
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln科室 Or lng部门ID = 0 Then
        strSQL = "Select Nvl(Max(Lengthb(床号)),0) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    Else
        strSQL = "Select Nvl(Max(Lengthb(床号)),0) as 长度 From 床位状况记录 Where 状态='占用' And 科室ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng部门ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lng病人ID As Long, lng主页ID As Long, lng险类 As Long
    Dim arrTmp As Variant
    
    If Item.Key = mstrPrePati Then Exit Sub
    
    stbThis.Panels(2).Text = "共" & Val(lbl_s.Tag) & "个病人,当前:" & Item.Text & ",住院号:" & Item.SubItems(1)
        
    arrTmp = Split(Mid(Item.Key, 2), "_")
    lng病人ID = Val(arrTmp(0))
    lng主页ID = Val(arrTmp(1))
    lng险类 = Val(arrTmp(2))
    
    Call ReadExistsItem(lng病人ID, lng主页ID, lng险类)

    mstrPrePati = Item.Key
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw.Sorted = True
    With lvw
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
    lvw.SortKey = ColumnHeader.Index - 1
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub FindNextPati()
    On Error Resume Next
    If mrsPati Is Nothing Then Exit Sub
    If mrsPati.RecordCount = 0 Then Exit Sub
    If mrsPati.Filter = 0 Then Exit Sub
    If mrsPati.EOF Then
        mrsPati.MoveFirst
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    lvw.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID & "_" & mrsPati!险类).Selected = True
    lvw.SelectedItem.EnsureVisible
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub tbsClass_Click()
    If tbsClass.SelectedItem.Index <> 1 Then
        mrsExistItem.Filter = "类别='" & tbsClass.SelectedItem.Tag & "'"
    Else
        mrsExistItem.Filter = 0
    End If
    Set vsExist.DataSource = mrsExistItem
    If tbsClass.SelectedItem.Index <> 1 Then
        vsExist.ColHidden(ColNum.类别) = True
    Else
        vsExist.ColHidden(ColNum.类别) = False
    End If
    
    If InStr("中草药,中成药,西成药,材料", tbsClass.SelectedItem.Tag) = 0 Then
        vsExist.ColHidden(ColNum.产地) = True
    Else
        vsExist.ColHidden(ColNum.产地) = False
    End If
    
    vsExist.Tag = tbsClass.SelectedItem.Index
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.vsExist.Rows = 1 Then Exit Sub
    
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsExist
    
    objPrint.Title.Text = lvw.SelectedItem.SubItems(1) & "-" & lvw.SelectedItem.Text & "费用审核项目清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str病人ID As String, str主页ID As String, str住院号 As String
    Dim blnByDept As Boolean
    
    blnByDept = mnuViewByDept(1).Checked
    If Not lvw.SelectedItem Is Nothing Then
        str病人ID = Val(lvw.SelectedItem.Tag)
        str住院号 = Val(lvw.SelectedItem.SubItems(1))
        str主页ID = Val(lvw.SelectedItem.SubItems(8))
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人ID=" & str病人ID, "主页ID=" & str主页ID, "病区=" & IIf(blnByDept, 0, mlngDeptID), "病人科室=" & IIf(blnByDept, mlngDeptID, 0), "住院号=" & str住院号)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病区=" & IIf(blnByDept, 0, mlngDeptID), "病人科室=" & IIf(blnByDept, mlngDeptID, 0))
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

