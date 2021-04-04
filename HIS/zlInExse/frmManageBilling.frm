VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "住院记帐管理"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9960
   Icon            =   "frmManageBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmManageBilling.frx":08CA
   ScaleHeight     =   6225
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9960
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6600
      MinHeight1      =   720
      Width1          =   4995
      NewRow1         =   0   'False
      Caption2        =   "病人病区"
      Child2          =   "cboDept"
      MinWidth2       =   1800
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8070
         TabIndex        =   2
         Text            =   "cboDept"
         Top             =   240
         Width           =   1800
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   6900
         _ExtentX        =   12171
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
            NumButtons      =   19
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
               Caption         =   "划价"
               Key             =   "Price"
               Description     =   "划价"
               Object.ToolTipText     =   "划价"
               Object.Tag             =   "划价"
               ImageKey        =   "Price"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceBilling"
                     Object.Tag             =   "记帐单"
                     Text            =   "记帐单"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceTable"
                     Object.Tag             =   "记帐表"
                     Text            =   "记帐表"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceSimple"
                     Object.Tag             =   "简单记帐"
                     Text            =   "简单记帐"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Auditing"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageKey        =   "Auditing"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingBilling"
                     Object.Tag             =   "记帐单"
                     Text            =   "记帐单"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingTable"
                     Object.Tag             =   "记帐表"
                     Text            =   "记帐表"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingSimple"
                     Object.Tag             =   "简单记帐"
                     Text            =   "简单记帐"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingPati"
                     Object.Tag             =   "按病人审核"
                     Text            =   "按病人审核"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingBatch"
                     Object.Tag             =   "批量审核"
                     Text            =   "批量审核"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Billing_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发药"
               Key             =   "Give"
               Description     =   "发药"
               Object.ToolTipText     =   "发药"
               Object.Tag             =   "发药"
               ImageKey        =   "Give"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Give_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "销帐"
               Key             =   "Del"
               Description     =   "销帐"
               Object.ToolTipText     =   "对当前选中单据销帐"
               Object.Tag             =   "销帐"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   5865
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBilling.frx":0A58
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8731
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
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmManageBilling.frx":0DCC
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
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9855
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3915
      Width           =   9855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2805
      Left            =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   4948
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
      MouseIcon       =   "frmManageBilling.frx":0F6A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3990
      Width           =   9945
      _ExtentX        =   17542
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
      MouseIcon       =   "frmManageBilling.frx":1284
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   270
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
            Picture         =   "frmManageBilling.frx":159E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":17B8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":19D2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1BEC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2366
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2580
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":279A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":29B4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2BCE
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2DE8
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":34E2
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3BDC
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":42D6
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":44F0
            Key             =   "Give"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   270
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
            Picture         =   "frmManageBilling.frx":470A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4B3E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4D58
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":54D2
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":56EC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5906
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5B20
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5D3A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5F54
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":664E
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6D48
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":7442
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":765C
            Key             =   "Give"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   688
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "记帐单据(&1)"
            Key             =   "Auditing"
            Object.ToolTipText     =   "显示直接记帐或划价后审核了的记帐单据"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "划价单据(&2)"
            Key             =   "Price"
            Object.ToolTipText     =   "显示划价后未审核的记帐单据"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
         Caption         =   "住院记帐(&B)"
         Begin VB.Menu mnuEditBillingBilling 
            Caption         =   "记帐单(&B)"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuEditBillingTable 
            Caption         =   "记帐表(&T)"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuEditBillingSimple 
            Caption         =   "简单记帐(&S)"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuEditBillingCust 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "住院划价(&R)"
         Begin VB.Menu mnuEditPriceBilling 
            Caption         =   "记帐单(&B)"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuEditPriceTable 
            Caption         =   "记帐表(&T)"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuEditPriceSimple 
            Caption         =   "简单记帐(&S)"
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnuEditAuditing 
         Caption         =   "记帐审核(&A)"
         Begin VB.Menu mnuEditAuditingBilling 
            Caption         =   "记帐单(&B)"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuEditAuditingTable 
            Caption         =   "记帐表(&T)"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuEditAuditingSimple 
            Caption         =   "简单记帐(&S)"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuEditAuditing_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditAuditingPati 
            Caption         =   "按病人审核(&P)"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuEditAuditingBatch 
            Caption         =   "批量审核(&A)"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGive 
         Caption         =   "单据发药(&G)"
      End
      Begin VB.Menu mnuEditGive_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "修改单据(&M)"
         Shortcut        =   ^M
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
      Begin VB.Menu mnuEditDelBat 
         Caption         =   "批量销帐(&B)"
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
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
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
Attribute VB_Name = "frmManageBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private mstrFilter As String

'护士工作站调用时的过滤条件
'有3种情况：
'1.只过滤显示指定单据号的单据：.单据号,.医嘱ID
'2.过滤显示一组医嘱记录的单据：.医嘱ID,.发送号
'3.过滤显示某次发送的所有单据：.发送号
Private Type TYPE_NurseStation
    Nurse As Boolean '表明当前过滤显示是否强制使用护士工作站调用的条件
    病区ID As Long '护士工作站当前病人的病区
    科室ID As Long '护士工作站当前病人的科室
    发送号 As Long
    医嘱ID As Long '一组医嘱记录的ID：Nvl(相关ID,ID)
    单据号 As String
    划价 As Boolean '缺省定位的页面是否划价(可能两种情况都有,以护士工作站调用时的当前单据的情况为缺省)
    ReLoad As Boolean '是否重新载入(在已显示的情况下)
    Mode As Boolean '是否从模态窗口返回
End Type
Private mvNurseFilter As TYPE_NurseStation

'记帐管理本身的过滤条件
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    InPatientID As Double   '34512
    Patient As String
    FeeItems As String
    IncomeItems As String
End Type
Private SQLCondition As Type_SQLCondition

Private mbln记帐 As Boolean, mbln销帐 As Boolean
Private mstr操作员 As String, mstr医嘱期效 As String

Private mstrPage As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngDeptID As Long, mlngUnitID As Long
Private mblnMax As Boolean
Private mrsDept As ADODB.Recordset

Private mstrPrivs As String
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mlngModul As Long
Private mblnNOMoved As Boolean '记录当前选择的单据是否是在后备数据表中

Public Function ShowMeByNurse(frmMain As Object, ByVal lng病区ID As Long, ByVal lng科室ID As Long, _
    ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, ByVal strNO As String, ByVal bln划价 As Boolean) As Object
'功能：由护士工作站调用并自动过滤显示出指定医嘱条件的费用单据，目的是冲销这些单据
'参数：对应见类型TYPE_NurseStation中的字段定义
'返回：当从非模态窗口返回时,返回记帐管理窗体,用于跟踪关闭事件(非模态显示时的刷新问题)
    With mvNurseFilter
        .Nurse = True
        .病区ID = lng病区ID
        .科室ID = lng科室ID
        .发送号 = lng发送号
        .医嘱ID = lng医嘱ID
        .单据号 = strNO
        .划价 = bln划价
        .ReLoad = False
        .Mode = False
    End With
    
    On Error Resume Next
    If mstrPrivs <> "" Then '已打开
        mvNurseFilter.ReLoad = True
        Call Form_Load
    End If
    Me.Show , frmMain '以非模态显示
    Err.Clear
    
    If Not mvNurseFilter.Mode Then
        Set ShowMeByNurse = Me
    End If
End Function

Private Sub cboDept_Click()
    Dim strTmp As String
    
    If Not mvNurseFilter.Nurse Then
        If tbs.SelectedItem.Key = "Auditing" Then
            If InStr(mstrPrivs, ";查看记帐单;") <= 0 Then Exit Sub
        Else
            If InStr(mstrPrivs, ";查看划价单;") <= 0 Then Exit Sub
        End If
    End If
    
    If cboDept.ItemData(cboDept.ListIndex) = mlngUnitID Then Exit Sub
    mlngUnitID = cboDept.ItemData(cboDept.ListIndex)
    
    If mlngUnitID = 0 Then
        mlngDeptID = 0
    Else
        strTmp = Get科室IDs(mlngUnitID)
        If InStr(1, strTmp, ",") > 0 Then
            mlngDeptID = Split(strTmp, ",")(0)
        Else
            mlngDeptID = Val(strTmp)
        End If
    End If
        
    If Visible Then
        If mvNurseFilter.Nurse Then
            If Not mvNurseFilter.ReLoad Then Call ShowBillsByNurse
        Else
            Call ShowBills(mstrFilter)
        End If
    End If
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
    If InStr(";" & mstrPrivs, ";所有病区") > 0 Then strRootCaption = "所有病区"
    
    If zlSelectDept(Me, mlngModul, cboDept, mrsDept, cboDept.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub

End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
        
    If cboDept.ListIndex >= 0 Then Exit Sub
    zlControl.CboLocate cboDept, mlngUnitID, True
    If cboDept.ListIndex < 0 And cboDept.ListCount <> 0 Then cboDept.ListIndex = 0

End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
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
    
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) = 1 Then '批量记帐
        frmBillings.mbytUseType = 0
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 2
        frmBillings.mlngModule = mlngModul
        frmBillings.mstrInNO = strNO
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mbytUseType = 0
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 0
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mlngModule = mlngModul
            
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 0, 2, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        End If
    End If
End Sub

Private Sub mnuEditAuditingBatch_Click()
    Dim rsWarn As ADODB.Recordset
    Dim blnTrans As Boolean, Curdate As Date
    Dim strSql As String, str审核时间 As String, strNO As String, strInfo As String
    Dim lngCOL审核 As Long, lngCOLNO As Long, lngCOLInsure As Long
    Dim i As Long, j As Long, intInsure As Integer
    
    lngCOL审核 = GetColNum("审核")
    lngCOLNO = GetColNum("单据号")
    lngCOLInsure = GetColNum("险类")
        
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, lngCOL审核) = "√" And mshList.TextMatrix(i, lngCOLNO) <> "" Then
            j = j + 1
        End If
    Next
    If j = 0 Then
        MsgBox "没有选择要审核的划价单据！", vbExclamation, gstrSysName
        Exit Sub
    Else
        If MsgBox("确实要对选择的" & j & "张划价单据进行审核吗！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    Set rsWarn = GetUnitWarn

    '每张单据一个事务进行审核,避免大事务长时间锁定表
    For i = 1 To mshList.Rows - 1
        strNO = mshList.TextMatrix(i, lngCOLNO)
        If mshList.TextMatrix(i, lngCOL审核) = "√" And strNO <> "" Then
            '费用报警
            If AuditingWarn(mstrPrivsOpt, rsWarn, strNO, "") Then
                If str审核时间 = "" Then
                    Curdate = zlDatabase.Currentdate
                    str审核时间 = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                intInsure = Val(mshList.TextMatrix(i, lngCOLInsure))
                strSql = "zl_住院记帐记录_Verify('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL,NULL," & str审核时间 & ")"
                
                gcnOracle.BeginTrans
                    blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                    
                    If intInsure <> 0 Then '医保实时上传,传输费用明细
                        If gclsInsure.GetCapability(support记帐上传, , intInsure) And Not gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                            strInfo = ""
                            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strInfo, , intInsure) Then
                                gcnOracle.RollbackTrans
                                If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                                Call mnuViewReFlash_Click
                                Exit Sub        '只要有一次失败就退出
                            End If
                        End If
                    End If
                gcnOracle.CommitTrans
                blnTrans = False
                
                If intInsure <> 0 Then '医保延后上传,传输费用明细
                    If gclsInsure.GetCapability(support记帐上传, , intInsure) And gclsInsure.GetCapability(support记帐完成后上传, , intInsure) Then
                        strInfo = ""
                        If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strInfo, , intInsure) Then
                            If strInfo <> "" Then
                                MsgBox strInfo, vbInformation, gstrSysName
                            Else
                                MsgBox "单据""" & strNO & """的数据向医保传送失败,该单据已审核！", vbInformation, gstrSysName
                            End If
                            Call mnuViewReFlash_Click
                            Exit Sub '只要有一次失败就退出
                        End If
                    End If
                End If
                                
                If gbln审核打印 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "登记时间=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "药品单位=" & IIf(gbln住院单位, 1, 0), "重打=0", 2)
                End If
            End If
        End If
    Next
    
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call mnuViewReFlash_Click       '执行过程中出错，需要刷新
End Sub

Private Sub mnuEditBillingCust_Click(Index As Integer)
    '自定义记帐
    Dim varTemp As Variant
            
    '参数含义依次是：
    '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs、blnViewCancel
    varTemp = Array(mnuEditBillingCust(Index).Tag, 0, 0, "", mlngUnitID, mlngDeptID, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '返回值
    
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

Private Function GetDelSerial(ByVal strNO As String, strTime As String) As String
'功能：求指定记帐单中未完全执行及有剩余数量的行号,用于批量销帐
'参数：strTime=登记时间,用于部份审核的记帐单
'返回：空=表示没有可以销帐的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    
    strSql = _
        " Select 序号,Sum(Nvl(付数,1)*数次) as 数量" & _
        " From 住院费用记录" & _
        " Where 记录性质=2 And NO=[1] And 登记时间=[2] And Nvl(执行状态,0)<>1 And 价格父号 is NULL" & _
        " Group by 序号 Having Nvl(Sum(Nvl(付数,1)*数次),0)<>0"
    On Error GoTo errH
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    End If
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!序号
        rsTmp.MoveNext
    Loop
    GetDelSerial = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditDelApply_Click()
    Dim strMsg As String
    If mlngUnitID = 0 Then
        If cboDept.Visible Then
            strMsg = "请先选择病人病区!"
            cboDept.SetFocus
        Else
            strMsg = "请先选择病人病区!" & vbCrLf & "(显示病区选择列表:查看-工具栏-住院病区)"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 0
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

Private Sub mnuEditDelBat_Click()
    Dim arrSQL() As Variant, blnTrans As Boolean
    Dim i As Long, j As Long, intInsure As Integer
    Dim blnBat As Boolean, blnBilling As Boolean, blnFlagPrint As Boolean
    Dim strNO As String, blnDo As Boolean, bytType As Byte
    Dim strInfo As String, strTime As String, str序号 As String, strRebateNOS As String, strUnitIDs As String, strUnDelNOs As String
    Dim lngCol单据号 As Long, lngCol登记时间 As Long, lngCol记录性质 As Long, lngCol开单部门ID As Long
    
    If MsgBox("批量销帐操作不可恢复，确实要将当前列表中的单据全部销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If MsgBox("确实要将当前列表中的单据全部销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    lngCol单据号 = GetColNum("单据号")
    lngCol登记时间 = GetColNum("登记时间")
    lngCol记录性质 = GetColNum("记录性质")
    lngCol开单部门ID = GetColNum("开单部门ID")
    
    arrSQL = Array()
    j = 0: blnBilling = True
    For i = 1 To mshList.Rows - 1
        blnDo = True
        strNO = mshList.TextMatrix(i, lngCol单据号)
        strTime = mshList.TextMatrix(i, lngCol登记时间)
        bytType = Val(mshList.TextMatrix(i, lngCol记录性质))
        
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '是否已转入后备数据表中
            '记帐划价单不会在后备表中,记录性质只取2
            '这时不能根据showdetail时的mblnNOMoved来判断,因为没有点击bill行,mblnNOMoved仅是当前所选单据的性质
            '所以必须要现判断
            If zlDatabase.NOMoved("住院费用记录", strNO, , bytType, Me.Caption) Then
                If Not ReturnMovedExes(strNO, bytType, Me.Caption) Then blnDo = False
            End If
        End If
        
        '是否销帐记录
        blnDo = mshList.TextMatrix(i, GetColNum("符号")) <> 2
            
        '权限判断
        If blnDo Then
            If tbs.SelectedItem.Key = "Price" Then
                If Not BillOperCheck(5, mshList.TextMatrix(i, GetColNum("划价人")), CDate(strTime), "销帐", strNO) Then blnDo = False
            Else
                If Not BillOperCheck(5, mshList.TextMatrix(i, GetColNum("记帐人")), CDate(strTime), "销帐", strNO, , bytType) Then blnDo = False
            End If
        End If
        
        '项目冲销权限
        If blnDo Then
            If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime, bytType, 0) Then Screen.MousePointer = 0: Exit Sub  '不再继续，取消再来,否则可能不断弹出提示
        End If
        
        '全院销帐
        If blnDo And InStr(mstrPrivsOpt, ";全院销帐;") = 0 Then
            If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
            
            If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(i, lngCol开单部门ID)) & ",") = 0 Then
                strUnDelNOs = strUnDelNOs & "," & strNO
                blnDo = False
            End If
        End If
            
        '留观病人权限
        If blnDo Then
            strInfo = Check留观病人(strNO, mstrPrivsOpt, strTime, bytType)
            If strInfo <> "" Then
                Screen.MousePointer = 0
                MsgBox "单据""" & strNO & """中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
                Exit Sub '不再继续，取消再来
            End If
        End If
        
        '是否已执行
        If blnDo Then
            blnBat = Val(mshList.TextMatrix(i, GetColNum("多病人单"))) = 1
            If BillCanDelete(strNO, bytType, blnBat, strTime, , blnFlagPrint) <> 0 Then blnDo = False
            If blnFlagPrint Then
                If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        '出院病人操作权限判断
        If blnDo Then
            If Not BillCanBeOperate(strNO, mstrPrivsOpt, "批量销帐", strTime, , bytType) Then Screen.MousePointer = 0: Exit Sub
        End If
        
        '是否已经结帐(有的话只问一次)
        If blnDo Then
            If gbytBillOpt <> 0 Then
                If HaveBilling(2, strNO, True, strTime, bytType) <> 0 Then
                    If gbytBillOpt = 2 Then
                        blnDo = False
                    ElseIf gbytBillOpt = 1 Then
                        If j = 0 Then
                            j = j + 1
                            If MsgBox("对于已经结帐的单据要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                blnDo = False: blnBilling = False
                            End If
                        Else
                            blnDo = blnBilling
                        End If
                    End If
                End If
            End If
        End If
        
        '是否存在打折冲减记录
        If blnDo Then
            If CheckRecalcRecord(strNO) Then
                strRebateNOS = strRebateNOS & strNO & ","
                If (UBound(Split(strRebateNOS, ",")) Mod 8) = 0 Then strRebateNOS = strRebateNOS & vbCrLf
            End If
        End If
                
        '取可销帐的行号(部份审核的单据)
        str序号 = ""
        If blnDo Then
            If Not BillIdentical(strNO) Then            '不用判断自动记帐单
                str序号 = GetDelSerial(strNO, strTime)
                If str序号 = "" Then blnDo = False
            End If
        End If
        
        '医保病人的费用不允许批量销帐
        If blnDo And tbs.SelectedItem.Key = "Auditing" Then '划价销帐时不用
            intInsure = BillExistInsure(strNO, , , bytType) '判断是否医保病人记的帐
            If intInsure > 0 Then
                Screen.MousePointer = 0
                MsgBox "医保病人的记帐费用不允许批量销帐！", vbInformation, gstrSysName
                mshList.Row = i: mshList.TopRow = i
                Call mshList_EnterCell: Exit Sub
            End If
        End If

        '产生SQL
        If blnDo Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_住院记帐记录_Delete('" & strNO & "','" & str序号 & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & bytType & ")"
        End If
    Next
    Screen.MousePointer = 0
    
    If UBound(arrSQL) = -1 Then
        MsgBox "没有可以销帐的记录单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If strRebateNOS <> "" Then
       MsgBox "发现以下单据存在按费别重算的打折冲减记录:" & vbCrLf & Mid(strRebateNOS, 1, InStrRev(strRebateNOS, ",") - 1) & vbCrLf & _
                "结帐前请对这些单据的病人重算费用，否则病人将享受单据销帐前的打折优惠金额！", vbInformation, Me.Caption
    End If
    
    '执行过程
    Call ZLCommFun.ShowFlash("正在执行批量销帐,请稍候 ...", Me)
    DoEvents
    Me.Refresh
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    
    Call ZLCommFun.StopFlash
    Me.Refresh
    
    If strUnDelNOs <> "" Then
        MsgBox "你没有[全院销帐]的权限,以下其它科室的单据未销帐." & vbCrLf & Mid(strUnDelNOs, 2), vbInformation, gstrSysName
    End If
    
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    Call ZLCommFun.StopFlash
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditGive_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSQL() As String, i As Long
    Dim strSql  As String, blnTran As Boolean, bln记帐表 As Boolean
    Dim strNO As String, strTime As String
    Dim str开单科室 As String, strDate As String, str汇总号 As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以发药！", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    str开单科室 = mshList.TextMatrix(mshList.Row, GetColNum("开单科室"))
    bln记帐表 = Len(mshList.TextMatrix(mshList.Row, GetColNum("姓名"))) = 0
    On Error GoTo errH
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '只发放指定时间审核部份的内容
    Set rsTmp = Get待发药清单(strNO, strTime, bln记帐表)
    
    If rsTmp.EOF Then
        MsgBox "单据""" & strNO & """当前内容中没有可以发放的药品！", vbInformation, gstrSysName
        Exit Sub
    Else
        If IsNull(rsTmp!库房ID) Then
            MsgBox "该张单据当前内容未确定执行药房，不能在这里发药。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要对单据""" & strNO & """当前内容发药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        ReDim arrSQL(rsTmp.RecordCount - 1)
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        str汇总号 = zlDatabase.GetNextNo(20)
        
        For i = 0 To rsTmp.RecordCount - 1
            arrSQL(i) = "ZL_药品收发记录_部门发药(" & rsTmp!库房ID & "," & rsTmp!ID & ",'" & UserInfo.姓名 & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null," & str汇总号 & ")"
            rsTmp.MoveNext
        Next
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
        
    Call mshList_EnterCell
    
    '打印发药清单
    '经检查,目前ZL1_BILL_1133_2中并没有用到参数:开单科室,估计是为了某个用户的发药单需求而作的
    If MsgBox("单据""" & strNO & """发药完成，要打印发药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "单据号=" & strNO, "登记时间=" & strTime, str开单科室, 1)
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPriceBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditPriceSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditPriceTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingPati_Click()
    Err.Clear: On Error Resume Next
    If Not frmBillingAuditing.zlCardShow(Me, mlngModul, mstrPrivs, mlngUnitID) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditAuditingSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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
        MsgBox "单据中包含部份未完全审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '权限判断
    If tbs.SelectedItem.Key = "Price" Then
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("划价人")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))), "修改", strNO) Then Exit Sub
    Else
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("记帐人")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))), "修改", strNO) Then Exit Sub
    End If
    
    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '出院病人操作权限判断
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "修改") Then Exit Sub
    
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
    
    '全院销帐
    If InStr(mstrPrivsOpt, ";全院销帐;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("开单部门ID"))) & ",") = 0 Then
            MsgBox "你没有权限对其它科室的单据销帐,不允许修改该单据！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '如果包含部分执行或全部执行的项目,则不一定可以全部冲销,不允许修改
    If HaveExecute(2, strNO, 2) Then
        MsgBox "该单据中包含完全执行或部分执行的项目,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '已结帐单据判断
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
    
    '是否存在重算冲减记录
    If CheckRecalcRecord(strNO) Then
        MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
            "结帐前请按费别重算费用，否则病人将享受单据修改前的打折优惠金额！", vbInformation, Me.Caption
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
        
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '记帐修改
    Else
        gbytBilling = 1 '划价修改
    End If
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) = 1 Then '批量记帐
        frmBillings.mbytUseType = 0
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 0
        frmBillings.mstrInNO = strNO
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mbytUseType = 0
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 0
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 0, 0, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,修改后的单据号为:[" & gstrModiNO & "],要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
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
        
    If InStr(",0,1,", Val(mshList.TextMatrix(mshList.Row, GetColNum("符号")))) = 0 Then
        MsgBox "该单据为销帐单据或已被销帐，不能再打印！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "登记时间=" & strTime, "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=1", 2)
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim bln门诊留观 As Boolean
    Dim bln住院单位 As Boolean
    
    bln门诊留观 = gbln门诊留观
    bln住院单位 = gbln住院单位
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 0
    frmSetExpence.mbytUseType = 0
    frmSetExpence.Show 1, Me
    If gblnOK Then
        If bln门诊留观 <> gbln门诊留观 Then
            '留观病人
            mlngDeptID = -1: mlngUnitID = 0: mstrPage = ""
            Call InitUnits
        ElseIf bln住院单位 <> gbln住院单位 Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("单据号")) = "") Then
                Call mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病区=" & mlngUnitID, "病人科室=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病区=" & mlngUnitID, "病人科室=" & mlngDeptID, "NO=" & strNO, _
                "住院号=" & .TextMatrix(.Row, GetColNum("住院号")), _
                "病人ID=" & .TextMatrix(.Row, GetColNum("病人ID")), _
                "主页ID=" & .TextMatrix(.Row, GetColNum("主页ID")), _
                "开单人=" & .TextMatrix(.Row, GetColNum("开单人")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    With frmBillingFilter
        .mstrPrivs = mstrPrivs
        If .mlngDeptID <> mlngDeptID Then
            .mlngDeptID = mlngDeptID
            .mlngUnitID = mlngUnitID
            .LoadOper
        End If
        
        If tbs.SelectedItem.Key = "Auditing" Then
            .lbl操作员.Caption = "记帐人"
        Else
            .lbl操作员.Caption = "划价人"
        End If
        
        .Show 1, Me
        If gblnOK Then
            mvNurseFilter.Nurse = False '手工过滤后不再使用护士工作站调用的条件
            
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
            
            
            mstr操作员 = ""
            If .cbo操作员.ListIndex <> -1 Then
                If .cbo操作员.ItemData(.cbo操作员.ListIndex) <> 0 Then
                    mstr操作员 = zlStr.NeedName(.cbo操作员.Text)
                End If
            End If
        
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.Operator = mstr操作员
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.InPatientID = Val(.txt住院号.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt姓名.Text) & "%"
            SQLCondition.FeeItems = .mstrFeeItems
            SQLCondition.IncomeItems = .mstrIncomeItems
            
            mnuViewReFlash_Click
        End If
    End With
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
                
        '都需要登记时间(退费或部份审核)
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    Dim lngCOL审核 As Long, i As Long
    Dim lngCOLNO As Long
        
    If tbs.SelectedItem.Key = "Price" Then
        With mshList
            If .MouseRow > 0 Then
                If .TextMatrix(.Row, GetColNum("单据号")) <> "" Then
                    lngCOL审核 = GetColNum("审核")
                    If .MouseCol = lngCOL审核 Then
                        If .TextMatrix(.Row, lngCOL审核) = "√" Then
                            .TextMatrix(.Row, lngCOL审核) = ""
                        Else
                            .TextMatrix(.Row, lngCOL审核) = "√"
                        End If
                    Else
                        If mnuEditView.Enabled Then mnuEditView_Click
                    End If
                End If
            ElseIf .MouseRow = 0 And .Rows > 1 Then
                lngCOL审核 = GetColNum("审核")
                If .MouseCol = lngCOL审核 Then
                    lngCOLNO = GetColNum("单据号")
                    For i = 1 To mshList.Rows - 1
                        If .TextMatrix(i, lngCOLNO) <> "" Then
                            If .MouseCol = lngCOL审核 Then
                                If .TextMatrix(i, lngCOL审核) = "√" Then
                                    .TextMatrix(i, lngCOL审核) = ""
                                Else
                                    .TextMatrix(i, lngCOL审核) = "√"
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
    Else
        If mnuEditView.Enabled Then mnuEditView_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    '都需要登记时间(退费或部份审核)
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
    
    mnuEditAdjust.Enabled = Not blnDel
    '自动记帐单和医嘱生成的记帐单不允许修改
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) <> 3 _
                            And mshList.TextMatrix(mshList.Row, GetColNum("单据类型")) = "普通记帐"
    mnuEditGive.Enabled = Not blnDel And tbs.SelectedItem.Key = "Auditing"
    mnuEditDel.Enabled = Not blnDel
    mnuEditDelBat.Enabled = Not blnDel
    
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Give").Enabled = mnuEditGive.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
        
    If InStr(mstrPrivsOpt, ";住院划价;") = 0 And tbs.SelectedItem.Key <> "Auditing" Then
        mnuEditDel.Enabled = False
        mnuEditDelBat.Enabled = False
        tbr.Buttons("Del").Enabled = False
    End If
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
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
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If Me.ActiveControl Is cboDept Then
            Else
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim strNO As String, strTime As String, blnBat As Boolean, blnFlagPrint As Boolean
    Dim strInfo As String, str病人IDs As String
    Dim strInsure As String, arrInsure As Variant, intInsure As Integer
    Dim intTmp As Integer, i As Long, bytType As Byte   '记录性质
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以销帐！", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    bytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质")))
        
    '权限判断
    If tbs.SelectedItem.Key = "Price" Then
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("划价人")), CDate(strTime), "销帐", strNO) Then Exit Sub
    Else
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("记帐人")), CDate(strTime), "销帐", strNO, , bytType) Then Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '项目冲销权限
    If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime) Then Exit Sub
    
    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt, strTime)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已执行
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) = 1
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
    
    '是否已经结帐:0-未结帐,1=已全部结帐,2-已部分结帐
    intTmp = HaveBilling(2, strNO, False, strTime)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , arrInsure(i)) Then
                        '医保病人的单据,固定为已结帐的禁止销帐。
                        If intTmp = 1 Then
                            MsgBox "该医保记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            '按理说因为医保病人必须全部结帐,所以应该不会出现这种情况
                            '可能出现于医保与普通病人混合的记帐表,未精确处理
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
    
    '销帐模式(在记帐或划价时直接调入单据销帐的模式决定于本身的模式)
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '记帐销帐
    Else
        gbytBilling = 1 '划价销帐
    End If
    If blnBat Then '批量记帐
        frmBillings.mbytUseType = 0
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
        frmSimpleBilling.mbytUseType = 0
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
            If mvNurseFilter.Nurse And mvNurseFilter.医嘱ID <> 0 Then
                frmCharge.mlng医嘱ID = mvNurseFilter.医嘱ID
            End If
            
            frmCharge.mbytUseType = 0
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
            varTemp = Array(lng记帐ID, 0, 3, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK And Visible Then '护士站调用
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
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
    
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '记帐查阅
    Else
        gbytBilling = 1 '划价查阅
    End If
    
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("多病人单"))) = 1 Then '批量记帐
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnNOMoved = mblnNOMoved
        frmBillings.mstrTime = strTime
        frmBillings.mblnDelete = blnDel
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 0, 1, strNO, 0, 0, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
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
    If mvNurseFilter.Nurse Then
        Call ShowBillsByNurse
    Else
        Call ShowBills(mstrFilter)
    End If
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
            mnuEditBillingBilling_Click
        Case "Price"
            mnuEditPriceBilling_Click
        Case "Auditing"
            mnuEditAuditingBilling_Click
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
        Case "Give"
            mnuEditGive_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim lngCount As Integer
    Dim str记帐单ID As String
    
    Select Case ButtonMenu.Key
        Case "BillingBilling"
            mnuEditBillingBilling_Click
        Case "BillingTable"
            mnuEditBillingTable_Click
        Case "BillingSimple"
            mnuEditBillingSimple_Click
        Case "PriceBilling"
            mnuEditPriceBilling_Click
        Case "PriceTable"
            mnuEditPriceTable_Click
        Case "PriceSimple"
            mnuEditPriceSimple_Click
        Case "AuditingBilling"
            mnuEditAuditingBilling_Click
        Case "AuditingTable"
            mnuEditAuditingTable_Click
        Case "AuditingSimple"
            mnuEditAuditingSimple_Click
        Case "AuditingPati"
            mnuEditAuditingPati_Click
        Case "AuditingBatch"
            mnuEditAuditingBatch_Click
        Case Else
            '自定义记帐
            str记帐单ID = Mid(ButtonMenu.Key, 2)
            For lngCount = mnuEditBillingCust.LBound To mnuEditBillingCust.UBound
                If str记帐单ID = mnuEditBillingCust(lngCount).Tag Then
                    Call mnuEditBillingCust_Click(lngCount)
                    Exit Sub
                End If
            Next
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
    With frmBillingFilter
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
    
    mnuEditAdjust.Enabled = blnUsed
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEditGive.Enabled = blnUsed And tbs.SelectedItem.Key = "Auditing"
    tbr.Buttons("Give").Enabled = mnuEditGive.Enabled
    
    mnuEditDel.Enabled = blnUsed
    mnuEditDelBat.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
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
        mnuEditBillingCust(0).Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '如果创建成功，再读出对应的菜单
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,名称 From 收费记帐单 Where substr(适用范围,2,1)='1' Order by 编号"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '增加到主菜单中
            Load mnuEditBillingCust(lngCount)
            mnuEditBillingCust(lngCount).Caption = rsTmp("名称") & "(&" & lngCount & ")"
            mnuEditBillingCust(lngCount).Tag = rsTmp("ID")
            '增到工具栏菜单中
            If lngCount = 1 Then
                tbr.Buttons("Billing").ButtonMenus.Add , , "-"
            End If
            tbr.Buttons("Billing").ButtonMenus.Add , "C" & rsTmp("ID"), rsTmp("名称")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditBillingCust(0).Visible = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
'说明：因为本窗体被护士站非模态调用，可能强行重复执行Form_Load进行初始化,因此有些语句前用Visible作了判断
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    mlngModul = glngModul
    
    If Not Visible Then
        Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
        Call SetCustBill  '设置自定义记帐单
        Call RestoreWinState(Me, App.ProductName)
        Set stbThis.Panels(5).Picture = Me.Picture
    
        '刷新方式
        For i = 0 To mnuViewRefeshOptionItem.UBound
            If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
                mnuViewRefeshOptionItem(i).Checked = True
            Else
                mnuViewRefeshOptionItem(i).Checked = False
            End If
        Next
    End If
    
    If mvNurseFilter.Nurse Then
        tbs.Tabs(IIf(mvNurseFilter.划价, "Price", "Auditing")).Selected = True
    ElseIf Not Visible Then
        i = IIf(zlDatabase.GetPara("页面", glngSys, mlngModul, "1") = "1", 1, 2)
        tbs.Tabs(i).Selected = True
    End If
            
    mlngCurRow = 1: mlngTopRow = 1
    
    '权限设置
    If InStr(mstrPrivsOpt, ";住院记帐;") = 0 Then
        mnuEditBilling.Visible = False
        tbr.Buttons("Billing").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";住院划价;") = 0 Then
        mnuEditPrice.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记帐审核;") = 0 Then
        mnuEditAuditing.Visible = False
        tbr.Buttons("Auditing").Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";药品发药;") = 0 Then
        mnuEditGive.Visible = False
        mnuEditGive_.Visible = False
        tbr.Buttons("Give").Visible = False
        tbr.Buttons("Give_").Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录调整;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 _
        And InStr(mstrPrivsOpt, ";记录调整;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0 Then
        mnuEditDel.Visible = False
        mnuEditDelBat.Visible = False
        '55380
        If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";销帐审核;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
        
        '护士工作站调用时,权限不足提示
        If mvNurseFilter.Nurse Then
            MsgBox "你不具有住院记帐管理模块对应的销帐权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
        Or mvNurseFilter.Nurse _
        Or InStr(1, mstrPrivsOpt, "部分销帐") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";销帐审核;") = 0 Or mvNurseFilter.Nurse Then
        mnuEditDelAudit.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";重打单据;") = 0 Then
        mnuEditPrint.Visible = False
    End If
    
    
    '病区或科室初始
    If Not InitUnits Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "没有发现你所属部门,且你不具有所有病区权限,不能使用住院记帐管理！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If

    mstrPage = tbs.SelectedItem.Key
    
    mbln记帐 = True
    mbln销帐 = False
    mstr操作员 = UserInfo.姓名
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    If mvNurseFilter.Nurse Then
        Call mnuViewReFlash_Click
        If mvNurseFilter.ReLoad Then
            If Me.WindowState = 1 Then Me.WindowState = 0
            mvNurseFilter.ReLoad = False
        End If
        
        '冲销指定单据时，自动调出进行冲销
        If mvNurseFilter.单据号 <> "" Then
            Call mnuEditDel_Click
            If Not Visible Then
                mvNurseFilter.Mode = True
                Unload Me: Exit Sub
            End If
        End If
    Else
        stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
    End If
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
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = Me.ScaleTop + cbrH + 15
    
    mshList.Left = 0
    mshList.Top = tbs.Top + tbs.TabFixedHeight + 30
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = Me.ScaleLeft
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height - mshList.Height
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    mvNurseFilter.Nurse = False
    mstrPrivs = ""
    mstrPrivsOpt = ""
    mstrFilter = ""
    mlngUnitID = 0
    mlngDeptID = 0
    
    mstr医嘱期效 = ""
    Unload frmBillingFilter
    Unload frmBillingGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "页面", tbs.SelectedItem.Index, glngSys, mlngModul, blnHavePrivs
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
End Sub

Private Sub mnuViewGo_Click()
    frmBillingGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBillingGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, j As Long, blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents

        '比较条件
        blnFill = True
        With frmBillingGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt床号.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("床号")) = .txt床号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
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
        If mshList.MouseCol = GetColNum("审核") Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        If mvNurseFilter.Nurse Then
            Call ShowBillsByNurse(True)
        Else
            Call ShowBills(, True)
        End If
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If tbs.SelectedItem.Key = "Auditing" Then
        strHead = "险类,1,0|单据类型,1,900|单据号,1,850|住院号,1,750|床号,1,500|姓名,1,700|费别,1,900|医疗付款方式,1,1400|应收金额,7,850|实收金额,7,850" & _
                "|开单科室,1,1000|开单人,1,800|划价人,1,800|记帐人,1,800|登记时间,1,1850|说明,1,850|符号,1,0|记录性质,1,0|多病人单,1,0|记帐单ID,1,0|病人ID,1,0|主页ID,1,0|开单部门ID,1,0"
    Else
        strHead = "审核,1,450|险类,1,0|单据类型,1,900|单据号,1,850|住院号,1,750|床号,1,500|姓名,1,700|费别,1,900|医疗付款方式,1,1400|应收金额,7,850|实收金额,7,850" & _
                "|开单科室,1,1000|开单人,1,800|划价人,1,800|记帐人,1,800|登记时间,1,1850|说明,1,850|符号,1,0|记录性质,1,0|多病人单,1,0|记帐单ID,1,0|病人ID,1,0|主页ID,1,0|开单部门ID,1,0"
    End If
    
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or (tbs.SelectedItem.Key <> mstrPage) Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or (tbs.SelectedItem.Key <> mstrPage) Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("符号"): mshList.ColWidth(i) = 0
        i = GetColNum("记录性质"): mshList.ColWidth(i) = 0
        i = GetColNum("多病人单"): mshList.ColWidth(i) = 0
        i = GetColNum("记帐单ID"): mshList.ColWidth(i) = 0
        
        If tbs.SelectedItem.Key = "Auditing" Then
            mshList.ColWidth(GetColNum("记帐人")) = 800
        Else
            mshList.ColWidth(GetColNum("记帐人")) = 0
        End If
        
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
    Dim strSql As String, str医嘱期效 As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call ZLCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        '缺省过滤条件(一天内)
        If strIF = "" Then
            strIF = " And 登记时间 Between trunc(sysdate) And trunc(sysdate+1)-1/24/60/60"
            If tbs.SelectedItem.Key = "Auditing" Then
                strIF = strIF & "  And 记录性质=2 And 记录状态 IN(1,3)"
            Else
                strIF = strIF & " And 记录性质=2 And 记录状态=0"
            End If
            mstr医嘱期效 = ""   '缺省为普通记帐+长嘱+临嘱
        End If
        
        '操作员单独控制
        If mstr操作员 <> "" Then
            If tbs.SelectedItem.Key = "Auditing" Then
                strIF = strIF & " And 操作员姓名||''=[7]"
            Else
                strIF = strIF & " And 划价人||''=[7]"
            End If
        End If
        
        '主界面条件,所有病区时病人病区ID=0
        If mlngUnitID > 0 Then strIF = strIF & " And 病人病区ID+0=[8]"
        
        '记录性质(自动记帐单)
        strIF = "Where 门诊标志=2 " & strIF

        
        '单据号,住院号,床号,姓名,费别,应收金额,实收金额,开单科室,开单人,划价人,记帐人,登记时间,说明,符号,记录性质,多病人单,记帐单ID
        'Sign(执行状态):当记帐与销帐时间相同时有必要,如自动记帐
        If tbs.SelectedItem.Key = "Auditing" Then
            strIF = strIF & " And 操作员姓名 IS NOT NULL"
            
            '筛选时的时间在最后一次转出之前,且当前列表不是划价单
            If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
                strIF = zlGetFullFieldsTable("住院费用记录", 2, strIF, False)
            Else
                strIF = zlGetFullFieldsTable("住院费用记录", 0, strIF, False)
            End If
            
            strSql = _
                "Select Decode(Nvl(A.多病人单,0),1,NULL,C.险类) 险类,Decode(A.记录性质,3,'自动记帐',Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐')) as 单据类型, A.NO as 单据号," & _
                " To_Number(Decode(Nvl(A.多病人单,0),1,NULL,A.标识号)) as 住院号," & _
                " To_Char(Decode(Nvl(A.多病人单,0),1,NULL,C.出院病床)) as 床号," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.姓名) as 姓名," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.费别) as 费别,Decode(Nvl(A.多病人单,0),1,NULL,C.医疗付款方式) as 医疗付款方式," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,B.名称) as 开单科室," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.开单人) as 开单人," & _
                " A.划价人,A.操作员姓名 as 记帐人,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
                " Decode(A.记录性质,3,Decode(Max(A.记录状态),2,'自动销帐','自动记帐'),Decode(Max(A.记录状态),2,'销帐记录','记帐记录')) as 说明," & _
                " Max(A.记录状态) as 符号,A.记录性质,A.多病人单,A.记帐单ID,Decode(Nvl(A.多病人单,0),1,0,A.病人ID) 病人ID,Decode(Nvl(A.多病人单,0),1,0,A.主页ID) 主页ID,A.开单部门ID" & _
                " From (" & strIF & ") A,部门表 B,病案主页 C,病人医嘱记录 D" & _
                " Where A.开单部门ID=B.ID And A.病人ID=C.病人ID And A.主页ID=C.主页ID " & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                " And A.医嘱序号=D.id(+) " & mstr医嘱期效 & _
                " Group by Sign(Decode(Nvl(A.执行状态,0),0,1,Nvl(A.执行状态,0))),Decode(Nvl(A.多病人单,0),1,NULL,C.险类),Decode(A.记录性质,3,'自动记帐',Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐')),A.NO," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,B.名称),Decode(Nvl(A.多病人单,0),1,NULL,A.开单人)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.标识号),Decode(Nvl(A.多病人单,0),1,NULL,C.出院病床)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.姓名),Decode(Nvl(A.多病人单,0),1,NULL,A.费别)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,C.医疗付款方式)," & _
                " A.划价人,A.操作员姓名,A.登记时间,A.记录性质,A.多病人单,A.记帐单ID," & _
                " Decode(Nvl(A.多病人单,0),1,0,A.病人ID),Decode(Nvl(A.多病人单,0),1,0,A.主页ID),A.开单部门ID" & _
                " Order by A.登记时间 Desc,A.NO Desc"
        Else
            strIF = strIF & " And 操作员姓名 IS NULL And 划价人 is Not NULL"
            
                    '筛选时的时间在最后一次转出之前,且当前列表不是划价单
            If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
                strIF = zlGetFullFieldsTable("住院费用记录", 2, strIF, False)
            Else
                strIF = zlGetFullFieldsTable("住院费用记录", 0, strIF, False)
            End If
        
            strSql = _
                "Select '√' 审核,Decode(Nvl(A.多病人单,0),1,NULL,C.险类) 险类,Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐') as 单据类型,A.NO as 单据号," & _
                " To_Number(Decode(Nvl(A.多病人单,0),1,NULL,A.标识号)) as 住院号," & _
                " To_Char(Decode(Nvl(A.多病人单,0),1,NULL,C.出院病床)) as 床号," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.姓名) as 姓名," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.费别) as 费别,Decode(Nvl(A.多病人单,0),1,NULL,C.医疗付款方式) as 医疗付款方式," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,B.名称) as 开单科室," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.开单人) as 开单人," & _
                " A.划价人,A.操作员姓名 as 记帐人,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
                " Decode(Max(A.记录状态),2,'销帐记录','记帐记录') as 说明,Max(A.记录状态) as 符号,A.记录性质,A.多病人单,A.记帐单ID,Decode(Nvl(A.多病人单,0),1,0,A.病人ID) 病人ID,Decode(Nvl(A.多病人单,0),1,0,A.主页ID) 主页ID,A.开单部门ID" & _
                " From (" & strIF & ") A,部门表 B,病案主页 C,病人医嘱记录 D" & _
                " Where A.开单部门ID=B.ID And A.病人ID=C.病人ID And A.主页ID=C.主页ID" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                " And A.医嘱序号=D.id(+) " & mstr医嘱期效 & _
                " Group by Decode(Nvl(A.多病人单,0),1,NULL,C.险类),Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐'),A.NO," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,B.名称),Decode(Nvl(A.多病人单,0),1,NULL,A.开单人)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.标识号),Decode(Nvl(A.多病人单,0),1,NULL,C.出院病床)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,A.姓名),Decode(Nvl(A.多病人单,0),1,NULL,A.费别)," & _
                " Decode(Nvl(A.多病人单,0),1,NULL,C.医疗付款方式)," & _
                " A.划价人,A.操作员姓名,A.登记时间,A.记录性质,A.多病人单,A.记帐单ID,Decode(Nvl(A.多病人单,0),1,0,A.病人ID),Decode(Nvl(A.多病人单,0),1,0,A.主页ID),A.开单部门ID" & _
                " Order by A.登记时间 Desc,A.NO Desc"
        End If
        With SQLCondition
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, mstr操作员, mlngUnitID, .FeeItems, .IncomeItems)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        '求实收合计金额
        If Not blnSort Then
            strSql = "Select Sum(实收金额) as 金额,Count(Distinct NO) as 单据 From (" & _
                Replace(strIF, "记录状态 IN(1,3)", "记录状态 IN(1,2,3)") & ") A,部门表 B Where A.开单部门ID = B.ID" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)"
            With SQLCondition
            Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, mstr操作员, mlngUnitID, .FeeItems, .IncomeItems)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '设置颜色
    If mbln销帐 And Not mbln记帐 And tbs.SelectedItem.Key = "Auditing" Then
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
    
    If Not blnSort Then Call ZLCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowBillsByNurse(Optional blnSort As Boolean)
'功能:按护士工作站调用条件读取要冲销的单据列表
'参数:blnSort=不重新读取数据,仅重新显示已排序的内容
'说明:因为要进行销帐操作,只读取在线数据;护士工作站调用之前应已作相关判断
    Dim strSql As String, strIF As String
    Dim i As Long, j As Long, k As Long
    
    On Error GoTo errH
    
    If Not blnSort Then
        '护士工作站主条件
        If mvNurseFilter.单据号 <> "" Then
            strIF = " Where 门诊标志=2 And 记录性质=2 And NO=[1]"
        ElseIf mvNurseFilter.医嘱ID <> 0 Then
            strIF = "Select ID From 病人医嘱记录 Where ID=[2] Or 相关ID=[2]"
            strIF = "Select NO From 病人医嘱发送 Where 记录性质=2 And 发送号=[3] And 医嘱ID IN(" & strIF & ")"
            strIF = "Where 门诊标志=2 And 记录性质=2 And NO IN(" & strIF & ")"
        ElseIf mvNurseFilter.发送号 <> 0 Then
            strIF = "Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 发送号=[3]"
            strIF = "Where 门诊标志=2 And 记录性质=2 And NO IN(" & strIF & ")"
        End If
                
        '记帐或划价
        If tbs.SelectedItem.Key = "Auditing" Then
            strIF = strIF & " And 记录状态 IN(1,3)"
        Else
            strIF = strIF & " And 记录状态=0"
        End If
        
        '病区或科室
        If mlngDeptID > 0 Then strIF = strIF & " And 病人科室ID+0=[4]"
        strIF = zlGetFullFieldsTable("住院费用记录", 0, strIF, False)
        
        strSql = _
            "Select " & IIf(tbs.SelectedItem.Key = "Price", " NULL as 审核,", "") & _
            " C.险类,Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐') as 单据类型,A.NO as 单据号," & _
            " A.标识号 as 住院号,A.床号,A.姓名,A.费别,C.医疗付款方式," & _
            " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
            " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
            " B.名称 as 开单科室,A.开单人,A.划价人,A.操作员姓名 as 记帐人,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
            " Decode(Max(A.记录状态),2,'销帐记录','记帐记录') as 说明," & _
            " Max(A.记录状态) as 符号,A.记录性质,A.多病人单,A.记帐单ID,A.病人ID,A.主页ID" & _
            " From (" & strIF & ") A,部门表 B,病案主页 C,病人医嘱记录 D" & _
            " Where A.开单部门ID=B.ID(+) And A.病人ID=C.病人ID And A.主页ID=C.主页ID And A.医嘱序号=D.ID" & _
            " Group by C.险类,Decode(D.医嘱期效,1,'临嘱记帐',0,'长嘱记帐','普通记帐')," & _
            " A.NO,B.名称,A.开单人,A.标识号,A.床号,A.姓名,A.费别,C.医疗付款方式," & _
            " A.划价人,A.操作员姓名,A.登记时间,A.记录性质,A.多病人单,A.记帐单ID,A.病人ID,A.主页ID" & _
            " Order by A.登记时间 Desc,A.NO Desc"
        With mvNurseFilter
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .单据号, .医嘱ID, .发送号, mlngDeptID)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        '求实收合计金额
        If Not blnSort Then
            strSql = "Select Sum(实收金额) as 金额,Count(Distinct NO) as 单据 From (" & Replace(strIF, "记录状态 IN(1,3)", "记录状态 IN(1,2,3)") & ")"
            With mvNurseFilter
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .单据号, .医嘱ID, .发送号)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    
    '设置颜色
    mshList.ForeColor = ForeColor
    k = GetColNum("符号")
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, k)) = 3 Then
            '包含销帐的用蓝色
            mshList.Row = i
            For j = 0 To mshList.Cols - 1
                mshList.Col = j
                mshList.CellForeColor = &HC00000
            Next
        End If
    Next
        
    Call SetHeader
    If mrsList.EOF Then Call SetDetail
    
    mshList.Redraw = True

    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    Dim blnVisible As Boolean

    If (tbs.SelectedItem.Key = mstrPage) And Visible Then Exit Sub      '启动时要进入后面判断权限
    
    
    If mvNurseFilter.Nurse Then
        blnVisible = True
    Else
        If tbs.SelectedItem.Key = "Auditing" Then
            blnVisible = InStr(mstrPrivs, ";查看记帐单;") > 0
        Else
            blnVisible = InStr(mstrPrivs, ";查看划价单;") > 0
        End If
    End If
    
    mnuEditView.Visible = blnVisible
    tbr.Buttons("View").Visible = blnVisible
    mnuViewFilter.Visible = blnVisible
    tbr.Buttons("Filter").Visible = blnVisible
    mnuViewreFlash.Visible = blnVisible
    
    If Not blnVisible Then
        mnuViewRefeshOptionItem(0).Checked = True '不刷新
        mnuViewRefeshOptionItem(1).Checked = False
        mnuViewRefeshOptionItem(2).Checked = False
        mnuViewRefeshOptionItem(1).Enabled = False
        mnuViewRefeshOptionItem(2).Enabled = False
        
        mshList.Clear
        mshList.Rows = 2
        mshDetail.Clear
        mshDetail.Rows = 2
        
        Call SetHeader
        Call SetDetail
        Call SetMenu(False)
        
        mstrPage = tbs.SelectedItem.Key
        Exit Sub
    End If
    
    mstrFilter = ""   ' 记录性质变了，要清除条件
    
    If Visible Then
        If mvNurseFilter.Nurse Then
            If Not mvNurseFilter.ReLoad Then Call ShowBillsByNurse
        Else
            Call ShowBills(mstrFilter) '进入窗体时缺省不显示任何单据
        End If
    End If
    
    If mshList.Visible And mshList.Enabled Then mshList.SetFocus
    
    mstrPage = tbs.SelectedItem.Key
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院科室
    Dim i As Long
    Dim strServiceRange As String
    
    On Error GoTo errH
    cboDept.Clear
    If InStr(";" & mstrPrivs, ";所有病区") > 0 Then cboDept.AddItem "所有病区"
        
    '有权则显示门诊观察室对应的临床科室,住院留观与住院相同
    If InStr(mstrPrivsOpt, ";门诊留观记帐;") And gbln门诊留观 Then
        strServiceRange = "1,2,3"
    Else
        strServiceRange = "2,3"
    End If
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";所有病区;") = 0, strServiceRange, "护理", True)
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboDept.AddItem mrsDept!编码 & "-" & mrsDept!名称
            cboDept.ItemData(cboDept.NewIndex) = mrsDept!ID
            
            '确定缺省的病区
            If mvNurseFilter.Nurse Then
                If mrsDept!ID = mvNurseFilter.病区ID Then cboDept.ListIndex = cboDept.NewIndex
            Else
                If UserInfo.部门ID = mrsDept!ID Then cboDept.ListIndex = cboDept.NewIndex
            End If
            
            mrsDept.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(";" & mstrPrivs, ";所有病区;") > 0 Then
        MsgBox "没有发现住院科室信息,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
    
    strHead = "住院号,1,750|床号,1,500|姓名,1,700|费别,1,750|开单科室,1,1000|开单人,1,700|类别,1,650|名称,1,1600" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|规格,1,1000|单位,4,500|数量,7,850|单价,7,850|应收金额,7,850|实收金额,7,850|统筹金额,7,850|执行科室,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    
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
        '刘兴洪:27990 2010-02-22 17:29:47
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(.Cols - 1) = 0
        
        '住院号,床号,姓名,费别,开单科室,开单人
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        
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
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '记帐划价单不检查是否在后备表中,因为不会转出到后备表
            mblnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, , CStr(bytFlag), Me.Caption)
        Else
            mblnNOMoved = False   '必须要有这一句
        End If
        
        strSql = _
        " Select A.标识号 as 住院号,A.床号,A.姓名,A.费别,F.名称 as 开单科室,A.开单人," & _
        "       C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
                IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 单位," & _
        "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ",'9999990.00000') as 数量, " & _
        "       To_Char(Sum(A.标准单价)" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.统筹金额),'9999999" & gstrDec & "') as 统筹金额, " & _
        "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
        "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行','第'||ABS(A.执行状态)||'次退费') as 说明, A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ," & _
        "       收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,部门表 F,药品规格 X" & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.收费细目ID=B.ID And A.收费类别=C.编码" & _
        "       And A.开单部门ID=F.ID(+) And A.执行部门ID=D.ID(+)" & _
        "       And A.NO=[1] And A.记录性质=[2] And A.门诊标志=2" & _
        "       And A.收费细目ID=X.药品ID(+) And A.记录状态" & IIf(blnDel, "=2", " IN(0,1,3)") & IIf(strTime <> "", " And A.登记时间=[3]", "") & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Group by Nvl(A.价格父号,A.序号),A.标识号,A.床号,A.姓名,A.费别,F.名称,A.开单人," & _
        "       C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格,A.计算单位,D.名称,Nvl(A.费用类型,B.费用类型)," & _
        "       A.执行状态,A.记录状态,X.药品ID,X.住院单位,Nvl(X.住院包装,1)" & _
        " Order by " & IIf(blnBat, "LPAD(A.床号,10,' '),A.标识号,Nvl(A.价格父号,A.序号)", "Nvl(A.价格父号,A.序号),A.标识号,LPAD(A.床号,10,' ')")
        
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag)
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
        '住院号,床号,姓名,费别,开单科室,开单人
        mshDetail.ColWidth(0) = 850
        mshDetail.ColWidth(1) = 800
        mshDetail.ColWidth(2) = 700
        mshDetail.ColWidth(3) = 500
        mshDetail.ColWidth(4) = 1000
        mshDetail.ColWidth(5) = 700
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then mshDetail.ColWidth(4) = 0
    End If
    
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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

