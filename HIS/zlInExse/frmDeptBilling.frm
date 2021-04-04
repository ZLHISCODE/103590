VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptBilling 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "科室分散记帐"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9870
   Icon            =   "frmDeptBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmDeptBilling.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2730
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   7140
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3930
      Width           =   7140
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   5835
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeptBilling.frx":0A58
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
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
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmDeptBilling.frx":0DCC
            Text            =   "状态说明"
            TextSave        =   "状态说明"
            Key             =   "state"
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
      TabIndex        =   6
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9870
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
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   7785
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
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
            NumButtons      =   16
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
               Caption         =   "记帐"
               Key             =   "Billing"
               Description     =   "记帐"
               Object.ToolTipText     =   "记帐"
               Object.Tag             =   "记帐"
               ImageKey        =   "Billing"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingBilling"
                     Object.Tag             =   "记帐单"
                     Text            =   "记帐单"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingTable"
                     Object.Tag             =   "记帐表"
                     Text            =   "记帐表"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingSimple"
                     Object.Tag             =   "简单记帐"
                     Text            =   "简单记帐"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "销帐"
               Key             =   "Del"
               Description     =   "销帐"
               Object.ToolTipText     =   "对当前选中单据销帐"
               Object.Tag             =   "销帐"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   2670
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5025
      ScaleWidth      =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   810
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4575
      Left            =   -15
      TabIndex        =   0
      ToolTipText     =   "进入时,默认显示7天以内的病人"
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
      NumItems        =   12
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
         Key             =   "医疗付款方式"
         Text            =   "医疗付款方式"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "当前科室ID"
         Text            =   "当前科室ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "病人类型"
         Text            =   "病人类型"
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
            Picture         =   "frmDeptBilling.frx":0F6A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1184
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":139E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":15B8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1D32
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1F4C
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2380
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":259A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":27B4
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2EAE
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":35A8
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":3CA2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":3EBC
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
            Picture         =   "frmDeptBilling.frx":40D6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":42F0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":450A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":4724
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":4E9E
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":50B8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":52D2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":54EC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":5706
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":5920
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":601A
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":6714
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":6E0E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":7028
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
         NumTabs         =   3
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
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "转出病人(&3)"
            Key             =   "转科"
            Object.Tag             =   "转科"
            Object.ToolTipText     =   "转出病人"
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
            Picture         =   "frmDeptBilling.frx":7242
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":7B1C
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
            Picture         =   "frmDeptBilling.frx":83F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":8CD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2955
      Left            =   2700
      TabIndex        =   1
      Top             =   960
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   5212
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
      MouseIcon       =   "frmDeptBilling.frx":95AA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   2715
      TabIndex        =   2
      Top             =   3960
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3307
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
      MouseIcon       =   "frmDeptBilling.frx":98C4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00808080&
      Caption         =   " 病人费用概况"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   2775
      TabIndex        =   10
      Top             =   765
      Width           =   6990
   End
   Begin VB.Label lbl_s 
      BackColor       =   &H00808080&
      Caption         =   " 时间:2001-01-01至2001-01-01"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   30
      TabIndex        =   9
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
      Begin VB.Menu mnuEditBilling 
         Caption         =   "记帐单(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditTable 
         Caption         =   "记帐表(&T)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSimple 
         Caption         =   "简单记帐(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "自定义记帐单(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(空)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "修改单据(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制单据(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "调整时间(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditAdjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "单据销帐(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelApply 
         Caption         =   "销帐申请(&Q)"
      End
      Begin VB.Menu mnuEditDelAudit 
         Caption         =   "销帐审核(&H)"
      End
      Begin VB.Menu mnuEditDel_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "打印单据(&P)"
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "显示病人方式(&M)"
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
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_5 
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
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
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
Attribute VB_Name = "frmDeptBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private mrsPati As ADODB.Recordset

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    FeeItems As String
    IncomeItems As String
    lngHospNo As Long '住院号
    strPatiName As String
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mintBedLen As Integer
Private mdtBegin As Date, mdtEnd As Date
Private mbln记帐 As Boolean, mbln销帐 As Boolean
Private mstr医嘱期效 As String

Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long

Private mblnFirst As Boolean, mblnMax As Boolean
Private mlngDeptID As Long, mlngUnitID As Long
Private mstrPage As String

Private mstrPrivs As String     '保存当前模块的授权功能
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mlngModul As Long
Private mblnNOMoved As Boolean '记录当前选择的单据是否是在后备数据表中
Private mblnNotClick As Boolean
Private mrsDept As ADODB.Recordset
'刘兴洪 问题:27380 日期:2010-01-22 15:11:16
Private Type Ty_Para
    bln转出病人 As Boolean
    int转出天数 As Integer
End Type
Private mTy_Modul_Para As Ty_Para

Private Sub zlSetPatiPages()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置转出的页面的显示
    '编制:刘兴洪
    '日期:2010-01-27 09:48:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln转出 As Boolean, blnHaveData As Boolean, i As Long
    
    bln转出 = mTy_Modul_Para.bln转出病人
    bln转出 = bln转出 And IIf(mnuViewByDept(0).Checked, mlngUnitID > 0, mlngDeptID > 0)
    
    blnHaveData = False
    For i = 1 To tbs.Tabs.Count
        If tbs.Tabs(i).Key = "转科" Then
            blnHaveData = True: Exit For
        End If
    Next
    If bln转出 Then
        If blnHaveData = False Then
            tbs.Tabs.Add , "转科", "转出病人(&3)"
        End If
    Else
        If blnHaveData Then
            '移出
            If tbs.SelectedItem.Index = i Then
               tbs.Tabs(1).Selected = True
            End If
            tbs.Tabs.Remove i
        End If
    End If
End Sub
Private Sub cboDept_Click()
    Dim strTmp As String
    If mnuViewByDept(0).Checked Then
        If cboDept.ItemData(cboDept.ListIndex) = mlngUnitID Then Exit Sub
        mlngUnitID = cboDept.ItemData(cboDept.ListIndex)
        '当前科室以选定的病人的科室来确定
    Else
        If cboDept.ItemData(cboDept.ListIndex) = mlngDeptID Then Exit Sub
        mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
        If mlngDeptID = 0 Then
            mlngUnitID = 0
        Else
            mlngUnitID = Get病区ID(mlngDeptID)
        End If
    End If
    mstrPage = ""
    Call zlSetPatiPages
        
    If Visible Then Call tbs_Click
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboDept.ListIndex <> -1 Then
        ZLCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    
    
    Dim strRootCaption As String
    strRootCaption = ""
    If InStr(mstrPrivs, ";所有病区;") > 0 Then strRootCaption = IIf(mnuViewByDept(1).Checked, "所有科室", "所有病区")
    
    
    If zlSelectDept(Me, mlngModul, cboDept, mrsDept, cboDept.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    Dim lngID As Long
    
    If cboDept.ListIndex >= 0 Then Exit Sub
    
   If mnuViewByDept(0).Checked Then
        lngID = mlngUnitID
        '当前科室以选定的病人的科室来确定
   Else
       lngID = mlngDeptID
   End If
   zlControl.CboLocate cboDept, lngID, True
   If cboDept.ListIndex < 0 And cboDept.ListCount <> 0 Then cboDept.ListIndex = 0
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    Call InitLocPar(mlngModul)
    If mblnFirst Then
        '问题:29435:主要是要屏蔽转科页面
        Call zlSetPatiPages
        If lvw.Visible And lvw.Enabled Then lvw.SetFocus
        mshList_GotFocus
        
        mblnFirst = False
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2010-01-22 15:12:34
    '问题:27380
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim i As Long, blnHaveData As Boolean
    
    strTemp = zlDatabase.GetPara("最近转出天数", glngSys, mlngModul, "0|3")
    mTy_Modul_Para.int转出天数 = Val(Split(strTemp & "|", "|")(1))
    mTy_Modul_Para.bln转出病人 = IIf(Val(Split(strTemp & "|", "|")(0)) = 1, True, False)
    blnHaveData = False
    For i = 1 To tbs.Tabs.Count
        If tbs.Tabs(i).Key = "转科" Then
            blnHaveData = True: Exit For
        End If
    Next
    If blnHaveData Then
        tbs.Tabs("转科").ToolTipText = "显示" & mTy_Modul_Para.int转出天数 & "天转出的病人"
    End If
End Sub


Private Sub mnuEditAdjust_Click()
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '自动记帐单禁止操作
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) = 3 Then
        MsgBox "该单据为自动记帐单,不能调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '已经冲销过(部分)的单据不允许调整
    If BillExistDelete(strNO, 2) Then
        MsgBox "该单据包含已销帐内容,不允许调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已经结帐
    If HaveBilling(2, strNO) <> 0 Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("该记帐单据包含已经结帐的内容,要调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "该记帐单据包含已经结帐的内容,不能调整！", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
        
    If BillisBatch(strNO) Then '批量记帐
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 2
        frmBillings.mstrInNO = strNO
        frmBillings.mbytUseType = 1
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng病人ID = CLng(lvw.SelectedItem.Tag)
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng病人ID = CLng(lvw.SelectedItem.Tag)
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        Dim lng病人ID  As Long
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 1
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng病人ID = CLng(lvw.SelectedItem.Tag)
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If Not lvw.SelectedItem Is Nothing Then lng病人ID = CLng(lvw.SelectedItem.Tag)
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 1, 2, strNO, mlngUnitID, mlngDeptID, lng病人ID, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuEditBilling_Click()
    Dim cur余额 As Currency, blnOut As Boolean
        
    '出院病人记帐权限
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    
    If blnOut Then
        cur余额 = Get病人余额(CLng(lvw.SelectedItem.Tag), 0)
        If cur余额 = 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur余额 <> 0 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 1
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng病人ID = CLng(lvw.SelectedItem.Tag)
    
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCopy_Click()
    Dim strNO As String, cur余额 As Currency, blnOut As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以复制！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '出院病人记帐权限
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    
    If blnOut Then
        cur余额 = Get病人余额(CLng(lvw.SelectedItem.Tag), 0)
        If cur余额 = 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur余额 <> 0 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytUseType = 1
    frmCharge.mbytInState = 0
    frmCharge.mblnCopyBill = True
    frmCharge.mstrInNO = strNO
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng病人ID = CLng(lvw.SelectedItem.Tag)
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCustBill_Click(Index As Integer)
    '自定义记帐
    Dim lng病人ID As Long, varTemp As Variant
    Dim cur余额 As Currency, blnOut As Boolean
    
    '出院病人记帐权限
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur余额 = Get病人余额(CLng(lvw.SelectedItem.Tag), 0)
        If cur余额 = 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur余额 <> 0 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '参数含义依次是：
    '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs、blnViewCancel
    
    If Not lvw.SelectedItem Is Nothing Then lng病人ID = CLng(lvw.SelectedItem.Tag)
    
    varTemp = Array(mnuEditCustBill(Index).Tag, 1, 0, "", mlngUnitID, mlngDeptID, lng病人ID, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '返回值
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditDelApply_Click()
    Dim lngPatientID As Long
    
    If mlngUnitID = 0 Then
        MsgBox "请先选择病人病区!", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    If Not lvw.SelectedItem Is Nothing Then lngPatientID = Val(lvw.SelectedItem.Tag)
    
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 0
        .mstrPrivs = mstrPrivs
        .mlngPatientID = lngPatientID
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditDelAudit_Click()
    If mlngUnitID = 0 Then
        MsgBox "请先选择病人病区!", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, strTime As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNO = "" Then
        MsgBox "当前没有单据可以打印！", vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) = 3 Then
        MsgBox "该单据为自动记帐单,操作不能继续！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    If mshList.TextMatrix(mshList.Row, GetColNum("符号")) <> 1 Then
        MsgBox "该单据为销帐单据或已被销帐，不能再打印！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1134", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1134", Me, "NO=" & strNO, "登记时间=" & strTime, "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=1", 2)
    End If
End Sub

Private Sub mnuEditSimple_Click()
    Dim cur余额 As Currency, blnOut As Boolean
    
    '出院病人记帐权限
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur余额 = Get病人余额(CLng(lvw.SelectedItem.Tag), 0)
        If cur余额 = 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur余额 <> 0 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mbytUseType = 1
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng病人ID = CLng(lvw.SelectedItem.Tag)
    
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditTable_Click()
    Dim cur余额 As Currency, blnOut As Boolean
    
    '出院病人记帐权限
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur余额 = Get病人余额(CLng(lvw.SelectedItem.Tag), 0)
        If cur余额 = 0 And InStr(mstrPrivsOpt, ";出院结清强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur余额 <> 0 And InStr(mstrPrivsOpt, ";出院未结强制记帐;") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mbytUseType = 1
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng病人ID = CLng(lvw.SelectedItem.Tag)
    
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditModi_Click()
    Dim strNO As String, strInfo As String, strUnitIDs As String
    Dim strInsure As String, arrInsure As Variant
    Dim i As Long
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
                
    '未全部审核或多次审核的不允许修改
    If Not BillIdentical(strNO) Then
        MsgBox "单据中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
        Exit Sub
    End If
                
    '单据修改权限
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))), "修改", strNO) Then Exit Sub
        
    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '出院病人操作权限判断
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "修改") Then Exit Sub
    
    
    '全院销帐
    If InStr(mstrPrivsOpt, ";全院销帐;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("开单部门ID"))) & ",") = 0 Then
            MsgBox "你没有权限对其它科室的单据销帐,不允许修改该单据！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
                    
    '去掉了医保连接匹配检查
                
    '包含分批或时价药品的单据禁止修改
    If Not BillCanModi(strNO, 2) Then
        MsgBox "该张单据中包含分批或时价药品,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
                
    '已经冲销过(部分)的单据不允许修改
    If BillExistDelete(strNO, 2) Then
        MsgBox "该单据包含已销帐费用,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '如果包含部分执行或全部执行的项目,则不一定可以全部冲销,不允许修改
    If HaveExecute(2, strNO, 2) Then
        MsgBox "该单据中包含完全执行或部分执行的项目,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '是否已经结帐单
    If HaveBilling(2, strNO) <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                        '医保病人的单据固定为已结帐就禁止修改
                        MsgBox "该医保记帐单据包含已经结帐的内容,不能修改！", vbExclamation, gstrSysName: Exit Sub
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("该记帐单据包含已经结帐的内容,要修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            MsgBox "该记帐单据包含已经结帐的内容,不能修改！", vbExclamation, gstrSysName: Exit Sub
                    End Select
                End If
            Next
        End If
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
        
    gbytBilling = 0 '记帐修改
    If BillisBatch(strNO) Then '批量记帐
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 0
        frmBillings.mstrInNO = strNO
        frmBillings.mbytUseType = 1
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng病人ID = CLng(lvw.SelectedItem.Tag)
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng病人ID = CLng(lvw.SelectedItem.Tag)
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        Dim lng病人ID  As Long
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 1
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            
            If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng病人ID = CLng(lvw.SelectedItem.Tag)
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If Not lvw.SelectedItem Is Nothing Then lng病人ID = CLng(lvw.SelectedItem.Tag)
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 1, 0, strNO, mlngUnitID, mlngDeptID, lng病人ID, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,修改后的单据号为:[" & gstrModiNO & "],要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        End If
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim bln门诊留观 As Boolean
    Dim bln住院留观 As Boolean
    Dim bln住院单位 As Boolean
    Dim bln转出 As Boolean
    
    bln门诊留观 = gbln门诊留观
    bln住院留观 = gbln住院留观
    bln住院单位 = gbln住院单位
    
    bln转出 = mTy_Modul_Para.bln转出病人
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytUseType = 1
    frmSetExpence.mbytInFun = 0
    frmSetExpence.Show 1, Me
    If gblnOK Then
        '问题:27380
        Call InitPara
        Call zlSetPatiPages
        
        If bln门诊留观 <> gbln门诊留观 Or bln住院留观 <> gbln住院留观 Then
            mlngDeptID = -1: mstrPage = ""
            Call InitUnits
        ElseIf bln住院单位 <> gbln住院单位 Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("单据号")) = "") Then
                Call mnuViewReFlash_Click
                bln转出 = mTy_Modul_Para.bln转出病人
            End If
        End If
        If bln转出 <> mTy_Modul_Para.bln转出病人 And tbs.SelectedItem.Index = 3 Then
             Call mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng病人ID As Long, lng主页ID As Long, str住院号 As String
    Dim strNO As String
    
    If Not lvw.SelectedItem Is Nothing Then
        '问题:29444
        lng病人ID = Val(Split(lvw.SelectedItem.Key, "_")(1))
        lng主页ID = Val(Split(lvw.SelectedItem.Key, "_")(2))
        
        str住院号 = lvw.SelectedItem.SubItems(1)
        
        strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        If strNO = "" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "病区=" & mlngUnitID, "病人科室=" & mlngDeptID, "住院号=" & str住院号)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "病区=" & mlngUnitID, "病人科室=" & mlngDeptID, _
                "NO=" & strNO, "开单人=" & mshList.TextMatrix(mshList.Row, GetColNum("开单人")), "住院号=" & str住院号)
        End If
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病区=" & mlngUnitID, "病人科室=" & mlngDeptID)
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    mlngDeptID = 0: mlngUnitID = 0
    Call InitUnits
End Sub

Private Sub mnuViewFilter_Click()
    With frmDeptFilter
        .mstrPrivs = mstrPrivs
        If .mlngDeptID <> mlngDeptID Then
            .mlngDeptID = mlngDeptID
            .mlngUnitID = mlngUnitID
            .GetOperator    '会隐式调用form_load事件
        End If
        .Show 1, Me
        If gblnOK Then
            mdtBegin = Format(.dtpB.Value, "yyyy-MM-dd 00:00:00")
            mdtEnd = Format(.dtpE.Value, "yyyy-MM-dd 23:59:59")
            
            mstrFilter = .mstrFilter
            mbln记帐 = .chkType(0).Value = 1
            mbln销帐 = .chkType(1).Value = 1
            
             If .chkBill(chkBills.临嘱记帐).Value = 1 And .chkBill(chkBills.长嘱记帐).Value = 1 Then
                If .chkBill(chkBills.普通记帐).Value = 0 And .chkBill(chkBills.自动记帐).Value = 0 Then
                    mstr医嘱期效 = " And D.医嘱期效 In(0,1)"
                ElseIf .chkBill(chkBills.普通记帐).Value = 0 Then
                    mstr医嘱期效 = " And (A.记录性质=2 And D.医嘱期效 In(0,1) Or A.记录性质=3)"
                Else
                    mstr医嘱期效 = ""
                End If
            ElseIf .chkBill(chkBills.临嘱记帐).Value = 1 Then
                If .chkBill(chkBills.普通记帐).Value = 0 And .chkBill(chkBills.自动记帐).Value = 0 Then
                    mstr医嘱期效 = " And D.医嘱期效=1"
                ElseIf .chkBill(chkBills.普通记帐).Value = 0 Then
                    mstr医嘱期效 = " And (A.记录性质=2 And D.医嘱期效=1 Or A.记录性质=3)"
                Else
                    mstr医嘱期效 = " And (D.医嘱期效=1 Or D.医嘱期效 is Null)"
                End If
            ElseIf .chkBill(chkBills.长嘱记帐).Value = 1 Then
                If .chkBill(chkBills.普通记帐).Value = 0 And .chkBill(chkBills.自动记帐).Value = 0 Then
                    mstr医嘱期效 = " And D.医嘱期效=0"
                ElseIf .chkBill(chkBills.普通记帐).Value = 0 Then
                    mstr医嘱期效 = " And (A.记录性质=2 And D.医嘱期效=0 Or A.记录性质=3)"
                Else
                    mstr医嘱期效 = " And (D.医嘱期效=0 Or D.医嘱期效 is Null)"
                End If
            Else
                mstr医嘱期效 = " And D.医嘱期效 is Null"
            End If
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.Operator = zlStr.NeedName(.cbo操作员.Text)
            SQLCondition.FeeItems = .mstrFeeItems
            SQLCondition.IncomeItems = .mstrIncomeItems
            
            '问题号: 51625修改人:刘兴洪,修改时间:2012-12-10 18:21:16
            SQLCondition.lngHospNo = Val(.txtHospitalNO.Text)
            SQLCondition.strPatiName = Trim(.txtName.Text)
            mnuViewReFlash_Click
        End If
    End With
End Sub


Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, strTime As String, blnDel As Boolean
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEditView.Enabled Then mnuEditView_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String
    Dim lng记帐单ID As Long, blnDo As Boolean, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "共" & Val(lbl_s.Tag) & "个病人,当前:" & lvw.SelectedItem.Text & ",住院号:" & _
                lvw.SelectedItem.SubItems(1) & ", 共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2

    mnuEditAdjust.Enabled = Not blnDel
    '自动记帐单和医嘱生成的记帐单不允许修改
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) <> 3 _
                            And mshList.TextMatrix(mshList.Row, GetColNum("单据类型")) = "普通记帐"
    mnuEditDel.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
    
    '设置可否复制单据
    blnDo = True
    If blnDel Then blnDo = False '销帐单据
    If blnDo Then '自动记帐单
        If Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) = 3 Then blnDo = False
    End If
    If blnDo Then '自定义记帐单
        If Val(mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))) <> 0 Then blnDo = False
    End If
    If blnDo Then If BillisBatch(strNO) Then blnDo = False '记帐表
    If blnDo Then If BillisSimple(strNO) Then blnDo = False '简单记帐
    
    mnuEditCopy.Enabled = blnDo
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEditDel.Enabled And mnuEditDel.Visible Then Call mnuEditDel_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then
                If Me.ActiveControl Is lvw Then
                    Call FindNextPati
                Else
                    Call SeekBill(False)
                End If
            End If
        Case vbKeyReturn
            If Not Me.ActiveControl Is cboDept Then
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim strNO As String, strTime As String
    Dim blnBat As Boolean, intTmp As Integer
    Dim str病人IDs As String, strInfo As String, i As Long, intInsure As Integer
    Dim strInsure As String, arrInsure As Variant, bytType As Byte, blnFlagPrint As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以销帐！", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    bytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质")))
    
    '权限判断
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), CDate(strTime), "销帐", strNO, , bytType) Then Exit Sub
        
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, bytType, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
        
    '项目冲销权限
    If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime, bytType) Then Exit Sub
        
    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt, strTime, bytType)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '是否已执行
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) <> 0
    i = BillCanDelete(strNO, bytType, blnBat, strTime, mstrPrivsOpt, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '该单据不存在
                MsgBox "指定单据中的内容不存在,或者你没有相关收费项目的销帐权限！", vbInformation, gstrSysName
            Case 2 '已经全部完全执行
                MsgBox "指定单据中的内容已经全部完全执行！", vbInformation, gstrSysName
            Case 3 '未完全执行部分剩余数量为0
                MsgBox "指定单据中的内容未完全执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '出院病人操作权限判断
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "销帐", strTime, str病人IDs, bytType) Then Exit Sub
    
    '是否已经结帐
    intTmp = HaveBilling(2, strNO, False, strTime, bytType)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                        '医保病人的单据,固定为已结帐的禁止销帐
                        If intTmp = 1 Then
                            MsgBox "该医保记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "该医保记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("该记帐单据包含已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "该记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                                Exit Sub
                            Else
                                MsgBox "该记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName
                            End If
                    End Select
                End If
            Next
        End If
    End If
    
    intInsure = BillExistInsure(strNO, , , bytType) '判断是否含有医保病人记的帐,记帐表检查其中只要有医保病人
    '医保销帐不允许对负数记录进行销帐
    If intInsure <> 0 Then
        If CheckNONegative(strNO, bytType) Then
            MsgBox "该单据存在负数记帐记录,不允许进行医保销帐操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    '是否存在重算冲减记录
    If CheckRecalcRecord(strNO) Then
        MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
            "结帐前请按费别重算费用，否则病人将享受已销帐单据的打折优惠金额！", vbInformation, Me.Caption
    End If
    
    On Error Resume Next
    Err.Clear
        
    If blnBat Then '批量记帐
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 3
        frmBillings.mstrInNO = strNO
        frmBillings.mstrTime = strTime
        frmBillings.mstr病人IDs = str病人IDs
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO, bytType) Then '简单记帐
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long, varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 1
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 3
            frmCharge.mstrInNO = strNO
            frmCharge.mbytNOType = bytType
            frmCharge.mstrTime = strTime
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 1, 3, strNO, 0, 0, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEditView_Click()
    Dim strNO As String, strTime As String, blnDel As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNO = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) = 3 Then
        MsgBox "该单据为自动记帐单,操作不能继续！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
        
    On Error Resume Next
    Err.Clear
    
    If BillisBatch(strNO) Then '批量记帐
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnNOMoved = mblnNOMoved   '是否从后备表中取数
        frmBillings.mstrTime = strTime
        frmBillings.mblnDelete = blnDel
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved   '是否从后备表中取数
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 1
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved   '是否从后备表中取数
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 1, 1, strNO, 0, 0, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuFile_quit_Click()
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
    cbr.Bands(1).minHeight = tbr.ButtonHeight
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
        If lbl_s.Width + X < 2580 Or mshList.Width - X < 3500 Then Exit Sub
        pic.Left = pic.Left + X
        lbl_s.Width = lbl_s.Width + X
        tbs.Width = tbs.Width + X
        lvw.Width = lvw.Width + X
        
        lblMoney.Left = lblMoney.Left + X
        lblMoney.Width = lblMoney.Width + X
        
        mshList.Left = mshList.Left + X
        mshList.Width = mshList.Width - X
        
        mshDetail.Left = mshDetail.Left + X
        mshDetail.Width = mshDetail.Width - X
        
        picHsc.Left = picHsc.Left + X
        picHsc.Width = picHsc.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lvw.SetFocus
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "病人颜色" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEditView_Click
        Case "Billing"
            mnuEditBilling_Click
        Case "Modi"
            mnuEditModi_Click
        Case "Del"
            mnuEditDel_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "BillingBilling"
            mnuEditBilling_Click
        Case "BillingTable"
            mnuEditTable_Click
        Case "BillingSimple"
            mnuEditSimple_Click
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
    
    intRow = mshList.Row
    
    '表头
    objOut.Title.Text = "住院记帐单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmDeptFilter
        objRow.Add "时间：" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " 至 " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    mnuEditCopy.Enabled = blnUsed
    mnuEditAdjust.Enabled = blnUsed
    
    mnuEditDel.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
End Sub

Private Sub SetCustBill()
'设置与自定义记帐单相关的内容
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long, lngSum As Long
    On Error Resume Next
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivsOpt, ";专项记帐;") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '如果创建成功，再读出对应的菜单
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,名称 From 收费记帐单 Where substr(适用范围,3,1)='1' Order by 编号"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '增加到主菜单中
            If lngCount > 1 Then
                Load mnuEditCustBill(lngCount)
            End If
            mnuEditCustBill(lngCount).Caption = rsTmp("名称") & "(&" & lngCount & ")"
            mnuEditCustBill(lngCount).Tag = rsTmp("ID")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditCustBill(1).Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, lngTmp As Long
    
    mstrPrivs = gstrPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call SetCustBill
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels("state").Picture = Me.Picture
        
    lngTmp = Val(zlDatabase.GetPara("显示病人方式", glngSys, mlngModul, 0))
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
    
    i = IIf(zlDatabase.GetPara("页面", glngSys, mlngModul, "1") = "1", 1, 2)
    tbs.Tabs(i).Selected = True
        
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '根据保存列表方式设置菜单
    Call SetView(lvw.View)

    mlngCurRow = 1: mlngTopRow = 1
    mblnFirst = True
    
    '权限设置
    If InStr(mstrPrivsOpt, ";住院记帐;") = 0 Then
        mnuEditBilling.Visible = False
        mnuEditTable.Visible = False
        mnuEditSimple.Visible = False
        mnuEditCust.Visible = False
        mnuEditCopy.Visible = False
        mnuEditBilling_.Visible = False
        
        tbr.Buttons("Billing").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 Then
        mnuEditDel.Visible = False
        If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";销帐审核;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        tbr.Buttons("Del").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
        Or InStr(1, mstrPrivsOpt, ";部分销帐;") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";销帐审核;") = 0 Then
        mnuEditDelAudit.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录调整;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 _
        And InStr(mstrPrivsOpt, ";记录调整;") = 0 _
        And InStr(mstrPrivsOpt, ";住院记帐;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    
    '55380
    If InStr(mstrPrivsOpt, ";住院记帐;") = 0 And InStr(mstrPrivsOpt, ";记录修改;") = 0 _
        And (InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0) Then
        tbr.Buttons("Del_").Visible = False
    End If
    
    Call InitPara
        
    If Not InitUnits Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "没有发现你所属科室,且你不具有所有病区权限,不能使用科室分散记帐！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
        
    mbln记帐 = True
    mbln销帐 = False
    
    mdtEnd = zlDatabase.Currentdate + 7
    mdtBegin = DateAdd("m", -1, mdtEnd)
    mstrPage = tbs.SelectedItem.Key
    
    
    Call LoadPatients '其中已包含Call SetDetail Call SetHeader  Call SetMenu
    
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, sngVsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    
    If mblnMax Then
        sngVsc = 0.3: mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    lbl_s.Left = Me.ScaleLeft
    lbl_s.Top = Me.ScaleTop + cbrH + 45
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = lbl_s.Top + lbl_s.Height + 45
    tbs.Width = lbl_s.Width
    
    lvw.Left = Me.ScaleLeft
    lvw.Top = tbs.Top + tbs.Height - 75
    lvw.Width = lbl_s.Width
    lvw.Height = Me.ScaleHeight - staH - cbrH - lbl_s.Height - tbs.Height - 15
    
    pic.Left = lvw.Left + lvw.Width
    pic.Top = Me.ScaleTop + cbrH
    pic.Height = Me.ScaleHeight - cbrH - staH
    
    lblMoney.Left = pic.Left + pic.Width
    lblMoney.Top = Me.ScaleTop + cbrH + 45
    lblMoney.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    
    mshList.Left = pic.Left + pic.Width
    mshList.Top = lblMoney.Top + lblMoney.Height + 15
    mshList.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    mshList.Height = (Me.ScaleHeight - cbrH - staH - lblMoney.Height - picHsc.Height - 60) * (1 - sngVsc)
    
    picHsc.Left = mshList.Left
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Width = mshList.Width
    
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Left = mshList.Left
    mshDetail.Width = mshList.Width
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - lblMoney.Height - picHsc.Height - mshList.Height - 60
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, lngTmp As Long
    Dim blnHavePara As Boolean
    blnHavePara = InStr(1, mstrPrivs, ";参数设置;") > 0
    mstrFilter = ""
    mlngDeptID = 0
    mlngUnitID = 0
    
    Set mrsPati = Nothing
    Unload frmDeptFilter
    Unload frmDeptGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "页面", tbs.SelectedItem.Index, glngSys, mlngModul, blnHavePara
    
        
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, blnHavePara
            Exit For
        End If
    Next
    
    '显示病人方式
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "显示病人方式", lngTmp, glngSys, mlngModul, blnHavePara
    
End Sub

Private Sub mnuViewGo_Click()
    Dim blnPati As Boolean
    blnPati = Me.ActiveControl Is lvw
    
    If blnPati Then
        '定位病人
        With frmDeptGo
            .fraBill.Visible = False
            .fraPati.Visible = True
            .Height = 2490
            .fraPati.Width = 3100
            .Width = .fraPati.Width + 600
            .cmdCancel.Left = .fraPati.Left + .fraPati.Width - .cmdCancel.Width - 100
            .cmdOk.Left = .cmdCancel.Left - .cmdOk.Width - 100
        End With
    Else
        '定位单据
        With frmDeptGo
            .fraBill.Visible = True
            .fraPati.Visible = False
            .Height = 1770
            .Width = .fraBill.Width + 600
            .cmdCancel.Left = .fraBill.Left + .fraBill.Width - .cmdCancel.Width - 100
            .cmdOk.Left = .cmdCancel.Left - .cmdOk.Width - 100
        End With
    End If
    frmDeptGo.Show 1, Me
    If gblnOK Then
        If blnPati Then
            Call FindPati
        Else
            Call SeekBill(frmDeptGo.optHead)
        End If
    End If
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    Dim strCurNO As String
    
    If frmDeptGo.txtNO.Text = "" Then Exit Sub
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmDeptGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
            Call mshList_EnterCell
            mlngGo = i + 1
            
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

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("单据号")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "单据类型,1,900|单据号,1,850|开单科室,1,850|开单人,1,800|费别,1,900|应收金额,7,850|实收金额,7,850|操作员,1,800|登记时间,1,1850|说明,1,850|符号,1,0|记录性质,1,0|多病人单,1,0|记帐单ID,1,0|开单部门ID,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("符号"): mshList.ColWidth(i) = 0
        i = GetColNum("记录性质"): mshList.ColWidth(i) = 0
        i = GetColNum("多病人单"): mshList.ColWidth(i) = 0
        i = GetColNum("记帐单ID"): mshList.ColWidth(i) = 0
        
        '查看医生的权限
        i = GetColNum("开单人")
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
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
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, lng病人ID As Long, lng主页ID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
        
    If Not blnSort Then
        Call ZLCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        '缺省过滤条件(一天内)
        SQLCondition.Default = (strIF = "")
        If strIF = "" Then
            strIF = " And 登记时间>Sysdate-1 And 记录性质=2 And 记录状态 IN(1,3) And 操作员姓名||''=[5]"
            mstr医嘱期效 = ""   '缺省为普通记帐+长嘱+临嘱
        End If
        
        If lvw.SelectedItem Is Nothing Then
            strIF = strIF & " And Rownum<1"
        Else
            lng病人ID = Val(Split(lvw.SelectedItem.Key, "_")(1))
            lng主页ID = Val(Split(lvw.SelectedItem.Key, "_")(2))
            strIF = strIF & " And 病人ID=[6] And Nvl(主页ID,0)=[7]"
        End If
        
               
        strIF = " Where 门诊标志=2 And 操作员姓名 is Not NULL " & strIF
        
        '筛选时的时间在最后一次转出之前,在院和出院病人都可能存在单据被转出
        If frmDeptFilter.mblnDateMoved Then
            strIF = zlGetFullFieldsTable("住院费用记录", 2, strIF, False)
        Else
            strIF = zlGetFullFieldsTable("住院费用记录", 0, strIF, False)
        End If
        
        '单据号,开单科室,开单人,费别,应收金额,实收金额,操作员,登记时间,说明,符号,记录性质,多病人单,记帐单ID
        'Sign(执行状态):当记帐与销帐时间相同时有必要,如自动记帐
        strSql = _
            "Select Decode(A.记录性质,3,'自动记帐',Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐')) as 单据类型,A.NO as 单据号," & _
            " Decode(Nvl(A.多病人单,0),1,NULL,B.名称) as 开单科室," & _
            " Decode(Nvl(A.多病人单,0),1,NULL,A.开单人) as 开单人," & _
            " A.费别,To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
            " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
            " A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
            " Decode(A.记录性质,3,Decode(Max(A.记录状态),2,'自动销帐','自动记帐'),Decode(Max(A.记录状态),2,'销帐记录','记帐记录')) as 说明," & _
            " Max(A.记录状态) as 符号,A.记录性质,A.多病人单,A.记帐单ID,A.开单部门ID" & _
            " From (" & strIF & ") A,部门表 B,病人医嘱记录 D" & _
            " Where A.开单部门ID=B.ID And A.医嘱序号=D.id(+) " & mstr医嘱期效 & _
            " Group by Sign(Decode(Nvl(A.执行状态,0),0,1,Nvl(A.执行状态,0))),A.NO," & _
            " Decode(A.记录性质,3,'自动记帐',Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐'))," & _
            " Decode(Nvl(A.多病人单,0),1,NULL,B.名称),Decode(Nvl(A.多病人单,0),1,NULL,A.开单人)," & _
            " A.费别,A.操作员姓名,A.登记时间,A.记录性质,A.多病人单,A.记帐单ID,A.开单部门ID" & _
            " Order by A.登记时间 Desc,A.NO Desc"
        With SQLCondition
            If .Default Then .Operator = UserInfo.姓名
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, lng病人ID, lng主页ID, .FeeItems, .IncomeItems)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",当前设置没有过滤出该病人相关的任何单据"
        Call SetMenu(False)
    Else
        '求实收合计金额
        If Not blnSort Then
            strSql = "Select Sum(实收金额) as 金额,Count(Distinct NO) as 单据 From (" & Replace(strIF, "记录状态 IN(1,3)", "记录状态 IN(1,2,3)") & ")"
            With SQLCondition
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, lng病人ID, lng主页ID, .FeeItems, .IncomeItems)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '设置颜色
    If mbln销帐 And Not mbln记帐 Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("符号")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '销帐记录用红色
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '包含销帐的用蓝色
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("单据号")) = "" Then Call SetDetail
        
    mshList.Redraw = True
    
    If Not lvw.SelectedItem Is Nothing And Not blnSort Then
        lng病人ID = Val(lvw.SelectedItem.Tag)
        mrsPati.Filter = "病人id=" & lng病人ID
        Set rsTmp = GetMoneyInfo(CLng(lvw.SelectedItem.Tag), , Not IsNull(mrsPati!险类), 2)
        mrsPati.Filter = ""
        
        If Not rsTmp Is Nothing Then
            lblMoney.Caption = " " & lvw.SelectedItem.Text & "  预交款：" & Format(rsTmp!预交余额, "0.00") & _
                ",未结费用：" & Format(rsTmp!费用余额, gstrDec) & ",剩余款：" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
        Else
            lblMoney.Caption = " " & lvw.SelectedItem.Text & "  预交款：0.00,未结费用：" & gstrDec & ",剩余款：0.00"
        End If
    End If
    
    If Not blnSort Then Call ZLCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    
    '读取数据
    mstrPage = tbs.SelectedItem.Key
    Call LoadPatients
    'lvw.SetFocus
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String, strServiceRange As String
    Dim blnByDept As Boolean
    
    On Error GoTo errH
    blnByDept = mnuViewByDept(1).Checked
    cbr.Bands(2).Caption = IIf(blnByDept, "病人科室", "病人病区")
    cboDept.Clear
    If InStr(mstrPrivs, ";所有病区;") > 0 Then cboDept.AddItem IIf(blnByDept, "所有科室", "所有病区")
    
    '有权则显示门诊观察室对应的临床科室,住院留观与住院相同
    If InStr(mstrPrivsOpt, ";门诊留观记帐;") And gbln门诊留观 Then
        strServiceRange = "1,2,3"
    Else
        strServiceRange = "2,3"
    End If
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";所有病区;") = 0, strServiceRange, IIf(blnByDept, "临床", "护理"), True)
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboDept.AddItem mrsDept!编码 & "-" & mrsDept!名称
            cboDept.ItemData(cboDept.NewIndex) = mrsDept!ID
            If UserInfo.部门ID = mrsDept!ID Then cboDept.ListIndex = cboDept.NewIndex
                
            mrsDept.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(mstrPrivs, ";所有病区;") > 0 Then
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
    Dim objItem As ListItem, strSql As String
    Dim i As Long, j As Long, strCount As String
    Dim blnByDept As Boolean, strWhere As String
    Dim strWhere变动 As String
    Dim blnFind As Boolean
    
    On Error GoTo errH
    
    Call ZLCommFun.ShowFlash("正在读取住院病人清单,请稍候 ...", Me)
    DoEvents
    
    Me.Refresh
    blnByDept = mnuViewByDept(1).Checked
    If blnByDept Then
        mintBedLen = GetMaxBedLen(mlngDeptID, True)
    Else
        mintBedLen = GetMaxBedLen(mlngUnitID, False)
    End If
    
    '留观病人条件
    If InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 _
        And InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strWhere = " And Nvl(B.病人性质,0) IN (0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        strWhere = " And Nvl(B.病人性质,0) IN (0,1)"
    ElseIf InStr(mstrPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strWhere = " And Nvl(B.病人性质,0) IN (0,2)"
    Else
        strWhere = " And Nvl(B.病人性质,0)=0"
    End If
    
    '问题号: 51625修改人:刘兴洪,修改时间:2012-12-10 18:21:16
    If SQLCondition.lngHospNo <> 0 Then
         strWhere = " And  B.住院号 =[6]"
    End If
    If SQLCondition.strPatiName <> "" Then
         strWhere = " And A.姓名 Like [7]"
    End If
    
    strWhere变动 = ""
    If blnByDept Then
        Select Case tbs.SelectedItem.Index
        Case 1, 2
            strWhere = strWhere & IIf(mlngDeptID > 0, " And B.出院科室ID" & IIf(tbs.SelectedItem.Index = 2, "+0", "") & "=[2]", "")
        Case Else
            strWhere = strWhere & IIf(mlngDeptID > 0, " And C.科室ID =[2]", "")
            strWhere变动 = IIf(mlngDeptID > 0, " And 科室ID =[2]", "")
        End Select
    Else
        Select Case tbs.SelectedItem.Index
        Case 1, 2
            strWhere = strWhere & IIf(mlngUnitID > 0, " And B.当前病区ID" & IIf(tbs.SelectedItem.Index = 2, "+0", "") & "=[1]", "")
        Case Else
            strWhere = strWhere & IIf(mlngDeptID > 0, " And C.病区ID =[1]", "")
            strWhere变动 = IIf(mlngDeptID > 0, " And 病区ID =[1]", "")
        End Select
    End If
    
    Select Case tbs.SelectedItem.Index
    Case 1
        '用病案主页的出院科室ID索引较慢,但因为可能有门诊留观病人(留观科室或病区没有床位),所以不能从床位状况记录表去找查
        '当前在院的病人
        strSql = _
            "Select   A.病人ID,B.主页ID,A.住院号, " & _
            "   Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(b.年龄, a.年龄) as 年龄,B.医疗付款方式," & _
            "   B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
            "   C.名称 as 当前科室,B.险类,B.病人性质,B.状态,B.出院科室ID 当前科室ID,B.病人类型" & _
            " From 病人信息 A,病案主页 B,部门表 C,在院病人 ZY " & _
            " Where A.病人ID=B.病人ID And B.病人ID=ZY.病人ID And B.出院科室ID=C.ID" & strWhere & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            " And B.出院日期 is NULL And Nvl(B.主页ID,0)<>0  " & _
             IIf(mlngDeptID = 0, " Order by  A.住院号 Desc", " Order by   LPAD(床号,10,' ')")
    Case 2
        '该期间出院的病人
        strSql = _
            "Select A.病人ID,B.主页ID,A.住院号," & _
            "   Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(b.年龄, a.年龄) as 年龄,B.医疗付款方式," & _
            "   B.入院日期,B.出院日期,LPAD(B.出院病床," & mintBedLen & ",' ') as 床号," & _
            "   C.名称 as 当前科室,B.险类,B.病人性质,B.状态,B.出院科室ID 当前科室ID,B.病人类型" & _
            " From 病人信息 A,病案主页 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=C.ID" & strWhere & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
            " And B.入院日期<=[4] And B.出院日期 Between [3] And [4]" & _
            IIf(mlngDeptID = 0, " Order by A.住院号 Desc", " Order by LPAD(床号,10,' ')")
    Case 3
        '可能存在同一病人一天或范围内的有两条以上的转科,则以最后一条为准.
        'And C.终止时间 =(Select Max(终止时间)  From 病人变动记录  Where 病人ID=C.病人ID And 主页ID=C.主页ID And  终止原因=3  And Nvl(附加床位,0)=0 " & strWhere变动 & ")"
        '问题:29435
 
        strSql = "" & _
        " Select /*+ RULE */ Distinct A.病人ID,B.主页ID,B.住院号, " & _
        "       Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(b.年龄, a.年龄) as 年龄,B.医疗付款方式," & _
        "           B.入院日期,B.出院日期,LPAD(C.床号," & mintBedLen & ",' ') as 床号," & _
        "           D.名称 as 当前科室,B.险类,B.病人性质,B.状态,C.科室ID as 当前科室ID,B.病人类型" & _
        " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D " & _
        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And C.科室ID=D.ID " & _
        "       And Nvl(B.状态,0)<>2 " & IIf(blnByDept, " And B.出院科室ID<>[1] ", " And B.当前病区ID<>[1] ") & _
        "       And B.病人ID=C.病人ID And B.主页ID=C.主页ID  " & _
        "       And C.终止原因=3 And (C.终止时间   Between Sysdate-[5] And Sysdate )  And Nvl(C.附加床位,0)=0" & _
        "       And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL  " & strWhere & _
        "       And C.终止时间 =(Select Max(终止时间)  From 病人变动记录  Where 病人ID=C.病人ID And 主页ID=C.主页ID And  终止原因=3  And Nvl(附加床位,0)=0 " & strWhere变动 & ")"
            
            '不包含审核归档的
            ',收费项目目录 E
            '    "       Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
            '    ",C.经治医师 as 住院医师,B.病案状态," & _
            '    " E.名称 as 护理等级,B.费别,B.病人类型,B.状态,B.险类,A.就诊卡号"
        strSql = "Select * FROM ( " & strSql & ") " & vbCrLf & IIf(mlngDeptID = 0, " Order by   住院号 Desc", " Order by   LPAD(床号,10,' ')")
    Case Else
        Exit Function
    End Select
    
    mdtBegin = CDate(Format(mdtBegin, "yyyy-MM-dd 00:00:00"))
    mdtEnd = CDate(Format(mdtEnd, "yyyy-MM-dd 23:59:59"))
    Set mrsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnitID, mlngDeptID, mdtBegin, _
        mdtEnd, mTy_Modul_Para.int转出天数, SQLCondition.lngHospNo, SQLCondition.strPatiName & "%")
  
    lvw.ListItems.Clear
    
    If Not mrsPati.EOF Then
        For i = 1 To mrsPati.RecordCount
            If IIf(IsNull(mrsPati!病人性质), 0, mrsPati!病人性质) = 0 Then
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!病人ID & "_" & mrsPati!主页ID, mrsPati!姓名, 1, 1)
            Else
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!病人ID & "_" & mrsPati!主页ID, mrsPati!姓名, 2, 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(mrsPati!住院号), "", mrsPati!住院号)
            objItem.SubItems(2) = IIf(IsNull(mrsPati!床号) And mrsPati!状态 = 0, "家庭", Nvl(mrsPati!床号, " "))
            objItem.SubItems(3) = IIf(IsNull(mrsPati!性别), "", mrsPati!性别)
            objItem.SubItems(4) = IIf(IsNull(mrsPati!年龄), "", mrsPati!年龄)
            objItem.SubItems(5) = Format(mrsPati!入院日期, "yyyy-MM-dd")
            objItem.SubItems(6) = Format(IIf(IsNull(mrsPati!出院日期), "", mrsPati!出院日期), "yyyy-MM-dd")
            objItem.SubItems(7) = IIf(IsNull(mrsPati!当前科室), "", mrsPati!当前科室)
            objItem.SubItems(8) = mrsPati!主页ID
            objItem.SubItems(9) = Nvl(mrsPati!医疗付款方式)
            objItem.SubItems(10) = Val("" & mrsPati!当前科室id)
            objItem.SubItems(11) = "" & mrsPati!病人类型
            objItem.Tag = mrsPati!病人ID
            objItem.ListSubItems(1).Tag = Nvl(mrsPati!状态)
            
            objItem.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsPati!病人类型))
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(Nvl(mrsPati!病人类型))
            Next
            
            If InStr(strCount & ",", "," & mrsPati!病人ID & ",") = 0 Then strCount = strCount & "," & mrsPati!病人ID
            mrsPati.MoveNext
        Next
        lbl_s.Tag = UBound(Split(Mid(strCount, 2), ",")) + 1
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " 当前在院的病人,人数:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " 时间:" & Format(mdtBegin, "yyyy-MM-dd") & "至" & Format(mdtEnd, "yyyy-MM-dd") & ",人数:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 3 Then
            lbl_s.Caption = "显示" & mTy_Modul_Para.int转出天数 & "天内转出的病人"
        End If
        Me.Refresh
        Call ClearFeeList
    Else
        lbl_s.Tag = ""
        stbThis.Panels(2).Text = ""
        Call ShowBills '没有病人就没有单据
    End If
    Call ZLCommFun.StopFlash
    Exit Function
errH:
    Call ZLCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call ZLCommFun.ShowFlash("正在读取住院病人清单,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ClearFeeList()
    '清除费用列信息
    With mshDetail
            .Clear
            .Rows = 2
            .Cols = 2
    End With
    With mshList
        .Clear
        .Rows = 2
        .Cols = 2
    End With
    Call SetHeader
End Sub
Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If mblnNotClick Then Exit Sub
    stbThis.Panels(2).Text = "共" & Val(lbl_s.Tag) & "个病人,当前:" & Item.Text & ",住院号:" & Item.SubItems(1)
    If mnuViewByDept(0).Checked Then mlngDeptID = Val(Item.SubItems(10))
    
    '读取单据
    Call ShowBills(mstrFilter)
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

Private Sub FindPati()
    Dim strFilter As String
    Dim strBed As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    With frmDeptGo
        If .txt住院号.Text <> "" Then strFilter = strFilter & " Or 住院号=" & .txt住院号.Text
        If .txt姓名.Text <> "" Then strFilter = strFilter & " Or 姓名 Like '%" & .txt姓名.Text & "%'"
        If .txt床号.Text <> "" Then
            strBed = .txt床号.Text
            If mintBedLen - ZLCommFun.ActualLen(strBed) > 0 Then
                strBed = String(mintBedLen - ZLCommFun.ActualLen(strBed), " ") & strBed
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
        lvw.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
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
    lvw.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID).Selected = True
    lvw.SelectedItem.EnsureVisible
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "开单科室,1,850|开单人,1,800|类别,1,650|名称,1,1600" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|规格,1,1000|单位,4,500|数量,7,850|单价,7,850|应收金额,7,850|实收金额,7,850|统筹金额,7,850|执行科室,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '刘兴洪:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(.Cols - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowDetail(Optional ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)
    Dim strSql As String, i As Long, j As Long
    Dim blnBat As Boolean, bytFlag As Byte
    
    On Error GoTo errH
        
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) <> 0
    bytFlag = mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))
    
    If Not blnSort Then
         If frmDeptFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, , 2, Me.Caption)
        Else
            mblnNOMoved = False   '必须要有这一句
        End If
        
        strSql = _
        " Select E.名称 as 开单科室,A.开单人,C.名称 as 类别,Nvl(F.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
                IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 单位," & _
        "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ",'9999990.00000') as 数量, " & _
        "       To_Char(Sum(A.标准单价)" & _
                IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.统筹金额),'9999999" & gstrDec & "') as 统筹金额, " & _
        "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
        "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行','第'||ABS(A.执行状态)||'次退费') as 说明, A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录  A") & "," & _
        "       收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,收费项目别名 F,药品规格 X" & _
                  IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.收费细目ID=B.ID And A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID" & _
        "       And A.NO=[1] And A.记录性质=[2] And A.门诊标志=2" & _
        "       And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
        "       And A.收费细目ID=X.药品ID(+) And A.病人ID+0=[3]" & _
        "       And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & IIf(strTime <> "", " And A.登记时间=[4]", "") & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Group by Nvl(A.价格父号,A.序号),E.名称,A.开单人,C.名称,Nvl(F.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & "B.规格,A.计算单位," & _
        "       D.名称,Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID,X.住院单位,Nvl(X.住院包装,1)" & _
        " Order by Nvl(A.价格父号,A.序号)"
        
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, Val(lvw.SelectedItem.Tag), CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, Val(lvw.SelectedItem.Tag))
        End If
    End If
        
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '设置颜色
    If blnDel Then
        '退费直接为红色
        mshDetail.ForeColor = &HC0
    Else
        '原始单据退过的为蓝色
        mshDetail.ForeColor = ForeColor
        For i = 1 To mshDetail.Rows - 1
            If Val(mshDetail.TextMatrix(i, mshDetail.Cols - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            End If
        Next
    End If

    Call SetDetail
        
    '记帐表要显示病人信息
    If blnBat Then
        '开单科室,开单人
        mshDetail.ColWidth(0) = 850
        mshDetail.ColWidth(1) = 800
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then mshDetail.ColWidth(1) = 0
    End If
        
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
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

