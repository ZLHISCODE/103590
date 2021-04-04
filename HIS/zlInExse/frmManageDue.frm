VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageDue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "应收款管理"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10815
   Icon            =   "frmManageDue.frx":0000
   KeyPreview      =   -1  'True
   Picture         =   "frmManageDue.frx":058A
   ScaleHeight     =   6195
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5080
      Left            =   2670
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5085
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   750
      Width           =   45
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7245
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
            Picture         =   "frmManageDue.frx":0718
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":0932
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":0B4C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":0D66
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":14E0
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":16FA
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":1914
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":1B2E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":1D48
            Key             =   "Add1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":1F62
            Key             =   "Add2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":265C
            Key             =   "Add3"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":2D56
            Key             =   "Add4"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":3450
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":366A
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7860
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
            Picture         =   "frmManageDue.frx":3884
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":3A9E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":3CB8
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":3ED2
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":464C
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":4866
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":4A80
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":4C9A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":4EB4
            Key             =   "Add1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":50CE
            Key             =   "Add2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":57C8
            Key             =   "Add3"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":5EC2
            Key             =   "Add4"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":65BC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDue.frx":67D6
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2730
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3930
      Width           =   8080
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10815
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6195
      MinHeight1      =   720
      Width1          =   4500
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   10695
         _ExtentX        =   18865
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
            NumButtons      =   13
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
               Caption         =   "收款"
               Key             =   "Add"
               Description     =   "收款"
               Object.ToolTipText     =   "收款"
               Object.Tag             =   "收款"
               ImageKey        =   "Add2"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退款"
               Key             =   "Del"
               Description     =   "退款"
               Object.ToolTipText     =   "对当前选中单据退款"
               Object.Tag             =   "退款"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin MSComctlLib.ImageList img16 
            Left            =   5880
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
                  Picture         =   "frmManageDue.frx":69F0
                  Key             =   "KM"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmManageDue.frx":72CA
                  Key             =   "KF"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img32 
            Left            =   6600
            Top             =   0
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
                  Picture         =   "frmManageDue.frx":7BA4
                  Key             =   "KM"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmManageDue.frx":847E
                  Key             =   "KF"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4875
      Left            =   -15
      TabIndex        =   0
      Top             =   960
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8599
      SortKey         =   1
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "姓名"
         Text            =   "姓名"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "住院号"
         Text            =   "住院号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "性别"
         Text            =   "性别"
         Object.Width           =   971
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "年龄"
         Text            =   "年龄"
         Object.Width           =   971
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "家庭地址"
         Text            =   "家庭地址"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "合约单位"
         Text            =   "合约单位"
         Object.Width           =   4304
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2835
      Left            =   2700
      TabIndex        =   2
      Top             =   1080
      Width           =   8110
      _ExtentX        =   14314
      _ExtentY        =   5001
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
      MouseIcon       =   "frmManageDue.frx":8D58
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBalance 
      Height          =   1875
      Left            =   2700
      TabIndex        =   4
      Top             =   3960
      Width           =   5300
      _ExtentX        =   9340
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
      MouseIcon       =   "frmManageDue.frx":9072
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshInvoice 
      Height          =   1875
      Left            =   8010
      TabIndex        =   3
      Top             =   3960
      Width           =   2805
      _ExtentX        =   4948
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
      MouseIcon       =   "frmManageDue.frx":938C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   2750
      TabIndex        =   1
      Top             =   720
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   635
      TabFixedWidth   =   2290
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "应收明细(&1)"
            Key             =   "Due"
            Object.ToolTipText     =   "当前病人的应收款明细"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "收款明细(&2)"
            Key             =   "Gathering"
            Object.ToolTipText     =   "当前病人的缴款明细"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   5835
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageDue.frx":96A6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
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
   Begin VB.Label lblDate 
      BackColor       =   &H00808080&
      Caption         =   " 日期:2006-12-27至2007-01-01"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   30
      TabIndex        =   8
      ToolTipText     =   "在该时间范围内的住院病人"
      Top             =   750
      Width           =   2550
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
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_DayReport 
         Caption         =   "打印应收款日报(&D)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFile_DayReport_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "收款(&A)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "退款(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "打印收款单(&P)"
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
Attribute VB_Name = "frmManageDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    PatientName As String
    PatientINID As String
    NO As String
    Invoice As String
    strUnitName As String '合约单位名称
    lngUnitID As Long '合约单位ID
    bln仅欠款病人 As Boolean
End Type
Private SQLCondition As Type_SQLCondition

Private Enum LVWCOL
    C0姓名 = 0
    C1住院号 = 1
    C2性别 = 2
    C3年龄 = 3
    C4家庭地址 = 4
    C5合约单位 = 5
End Enum

Private mrsList As ADODB.Recordset
Private mlngCurRow As Long, mblnMax As Boolean, mDatSys As Date
Private mblnNOMoved As Boolean, mblnDo As Boolean


Private Sub Form_Activate()
    On Error Resume Next
    Call mshList.SetFocus
    mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuEditAdd_Click()
    On Error Resume Next
    frmDue.mlng病人ID = Val(Mid(lvw.SelectedItem.Key, 2))
    frmDue.Show 1, Me
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

Private Sub mnuEditDel_Click()
    Dim strNO As String, lngFlag As Long, strSQL As String
    
    If tbs.SelectedItem.Key = "Gathering" Then
                
        strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        lngFlag = mshList.TextMatrix(mshList.Row, GetColNum("记录状态"))
        If strNO <> "" And lngFlag = 1 Then
            If MsgBox("你确定要对单据[" & strNO & "]进行退款吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            
            If CheckNOMoved(strNO) Then
                If MsgBox("该缴款单对应的结帐单已转出到后备数据表,确实要进行退款吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            End If
            
            On Error GoTo errH
            strSQL = "Zl_病人缴款记录_Delete('" & strNO & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            stbThis.Panels(2).Text = "单据[" & strNO & "]删除成功!"
            mnuViewReFlash_Click
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, lngFlag As Long, strSQL As String
    
    If tbs.SelectedItem.Key = "Gathering" Then
        strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        lngFlag = mshList.TextMatrix(mshList.Row, GetColNum("记录状态"))
        If strNO <> "" And lngFlag = 1 Then
            If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_1", Me) Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_1", Me, "NO=" & strNO, 2)
            End If
        End If
    End If
End Sub

Private Sub mnuFile_DayReport_Click()
    Dim lng病人ID As Long
    
    lng病人ID = Val(lvw.Tag)
    If lng病人ID <> 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_2", Me, "病人ID=" & lng病人ID, 2)
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    Dim strInNo As String
    Dim lngPatiID As Long, lngBalanceID As Long
    If Not lvw.SelectedItem Is Nothing Then
        strInNo = lvw.SelectedItem.ListSubItems(1).Text
        lngPatiID = Mid(lvw.SelectedItem.Key, 2)
    End If
    strNO = mshList.TextMatrix(mshList.Row, 0)
    If strNO <> "" Then
        If tbs.SelectedItem.Key = "Due" Then
            lngBalanceID = Val(mshList.TextMatrix(mshList.Row, 6))
        End If
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "NO=" & strNO, _
                "住院号=" & strInNo, _
                "病人ID=" & lngPatiID, _
                "结帐ID=" & lngBalanceID)
End Sub

Private Sub mnuViewFilter_Click()
    frmDueFilter.Show 1, Me
    If gblnOK Then
        With SQLCondition
            .DateB = frmDueFilter.dtpBegin.Value
            .DateE = frmDueFilter.dtpEnd.Value
            .PatientName = Trim(frmDueFilter.txt姓名.Text)
            .PatientINID = Trim(frmDueFilter.txt住院号.Text)
            .NO = Trim(frmDueFilter.txtNO.Text)
            .Invoice = Trim(frmDueFilter.txtInvoice.Text)
            '问题:40275
            .strUnitName = Trim(frmDueFilter.txtUnit.Text) '  As String '合约单位名称
            .lngUnitID = Val(frmDueFilter.txtUnit.Tag) ' As Long '合约单位ID
            .bln仅欠款病人 = frmDueFilter.chk仅显示欠费.Value = 1 ' As Boolean
            lblDate.Caption = "日期:" & Format(.DateB, "YYYY-MM-DD") & "至" & Format(.DateE, "YYYY-MM-DD")
            mblnNOMoved = zlDatabase.DateMoved(Format(.DateB, "yyyy-MM-dd HH:mm:ss"), , , Me.Caption)
        End With
        lvw.Tag = ""
        Call LoadPatients
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmDueGo.Show 1, Me
    If gblnOK Then Call SeekPatient
End Sub

Private Sub SeekPatient()
    Dim i As Long, lng病人ID As Long, blnFill As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    stbThis.Panels(2).Text = "正在定位满足条件的单据..."
    Screen.MousePointer = 11
    DoEvents
    On Error GoTo errH
    
    With frmDueGo
        If .txtNO.Text <> "" Or .txtInvoice.Text <> "" Then
            If .txtNO.Text <> "" Then
                strSQL = "NO = [1]"
                If .txtInvoice.Text <> "" Then strSQL = strSQL & " And 实际票号 = [2]"
            Else
                If .txtInvoice.Text <> "" Then strSQL = "实际票号 = [2] And 收费时间 Between [3] And [4]"
            End If
            
            strSQL = "Select 病人id From 病人结帐记录 Where " & strSQL
            If mblnNOMoved Then
                strSQL = strSQL & " Union All " & Replace(strSQL, "病人结帐记录", "H病人结帐记录")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .txtNO.Text, .txtInvoice.Text, SQLCondition.DateB, SQLCondition.DateE)
            If rsTmp.RecordCount > 0 Then
                lng病人ID = rsTmp!病人ID
            Else
                lng病人ID = -1
            End If
        End If
    
        For i = 1 To lvw.ListItems.Count
            If .txt住院号.Text <> "" Then
                blnFill = lvw.ListItems(i).SubItems(LVWCOL.C1住院号) = .txt住院号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = UCase(lvw.ListItems(i).Text) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If lng病人ID <> 0 Then
                blnFill = Val(Mid(lvw.ListItems(i).Key, 2)) = lng病人ID
            End If
            
            '满足则退出
            If blnFill Then
                lvw.ListItems(i).Selected = True
                lvw.ListItems(i).EnsureVisible
                lvw.Tag = ""
                Call lvw_ItemClick(lvw.ListItems(i))
                
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    End With
    
    stbThis.Panels(2).Text = "没有找到满足条件的病人!"
    Screen.MousePointer = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mshBalance_EnterCell()
    mshBalance.ForeColorSel = mshBalance.CellForeColor
End Sub

Private Sub mshBalance_GotFocus()
    Call SetActiveList(mshBalance)
End Sub

Private Sub mshInvoice_EnterCell()
    mshInvoice.ForeColorSel = mshInvoice.CellForeColor
End Sub

Private Sub mshInvoice_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_EnterCell()
    Dim bytFlag As Byte
    
    mshList.ForeColorSel = mshList.CellForeColor
    If tbs.SelectedItem.Key = "Gathering" Then
        mshList.Tag = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        bytFlag = Val(mshList.TextMatrix(mshList.Row, GetColNum("记录状态")))
        
        If bytFlag = 0 Or bytFlag = 1 Then
            mshBalance.Visible = True
            Call ShowDetail(mshList.Tag, mshBalance)
        ElseIf bytFlag = 2 Or bytFlag = 3 Then
            mshBalance.Visible = False
        End If
        
        Call ShowDetail(mshList.Tag, mshInvoice, bytFlag)
        Call Form_Resize
    Else
        mshList.Tag = ""
    End If
    Call SetMenu
End Sub

Private Sub mshList_GotFocus()
    
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If tbs.SelectedItem.Key = "Gathering" And mnuEditDel.Enabled And mnuEditDel.Visible Then Call mnuEditDel_Click
    End If
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub
Private Sub mnuFile_quit_Click()
    Unload Me
End Sub
Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub mnuViewReFlash_Click()
    If Not lvw.SelectedItem Is Nothing Then
        lvw.Tag = ""
        Call lvw_ItemClick(lvw.SelectedItem)
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
        If lvw.Width + X < 2550 Or mshList.Width - X < 2500 Then Exit Sub   'lblDate的最小宽度2550,tbs的最小宽度2500
        pic.Left = pic.Left + X
        lblDate.Width = lblDate.Width + X
        lvw.Width = lvw.Width + X
        
        tbs.Left = pic.Left + pic.Width
        tbs.Width = tbs.Width - X
        
        mshList.Left = tbs.Left
        mshList.Width = mshList.Width - X
        
        If mshBalance.Visible Then
            mshBalance.Left = mshList.Left
            mshBalance.Width = mshList.Width * 0.6
            mshInvoice.Left = mshBalance.Left + mshBalance.Width + 15
            mshInvoice.Width = mshList.Width - mshBalance.Width - 15
        ElseIf mshInvoice.Visible Then
            mshInvoice.Left = mshList.Left
            mshInvoice.Width = mshList.Width
        End If
        
        picHsc.Left = tbs.Left
        picHsc.Width = picHsc.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lvw.SetFocus
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshBalance.Visible And mshBalance.Height - Y < 1000 Or mshInvoice.Visible And mshInvoice.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        
        If mshBalance.Visible Then
            mshBalance.Top = mshBalance.Top + Y
            mshBalance.Height = mshBalance.Height - Y
            mshInvoice.Top = mshBalance.Top
            mshInvoice.Height = mshBalance.Height
        ElseIf mshInvoice.Visible Then
            mshInvoice.Top = mshInvoice.Top + Y
            mshInvoice.Height = mshInvoice.Height - Y
        End If
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
        Case "Add"
            mnuEditAdd_Click
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
    objOut.Title.Text = IIf(tbs.SelectedItem.Key = "Gathering", "病人应收款缴款清单", "病人应收款清单")
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(mDatSys, "yyyy年MM月dd日")
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

Private Sub SetMenu()
'功能：设置菜单可用状态
    Dim blnUsed As Boolean
    
    '以下菜单和按钮在有病人时允许
    blnUsed = lvw.ListItems.Count > 0
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    mnuViewGo.Enabled = blnUsed
    mnuViewStyle(0).Enabled = blnUsed
    mnuViewStyle(1).Enabled = blnUsed
    mnuViewStyle(2).Enabled = blnUsed
    mnuViewStyle(3).Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
    tbr.Buttons("Style").Enabled = blnUsed
    
    '登记缴款,定位到当前病人时
    blnUsed = Not lvw.SelectedItem Is Nothing
    mnuEditAdd.Enabled = blnUsed
    tbr.Buttons("Add").Enabled = blnUsed
    
    '删除和打印,当前单据为收款登记单,且记录状态为1时允许
    blnUsed = False
    If tbs.SelectedItem.Key = "Gathering" Then
        If mshList.Tag <> "" Then blnUsed = Val(mshList.TextMatrix(mshList.Row, GetColNum("记录状态"))) = 1
    End If
    mnuEditDel.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
End Sub

Private Function LoadPatients() As Boolean
'功能：读取指定范围内的病人列表
    Dim objItem As ListItem, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, strWhere As String, strWherePati As String
    
    On Error GoTo errH
    
    Call zlCommFun.ShowFlash("正在读取发生过应收款的病人清单,请稍候 ...", Me)
    Me.Refresh
    DoEvents
    strWhere = ""
    strWherePati = ""
    With SQLCondition
        strWhere = IIf(.NO <> "", " And A.NO=[4]", "")
        strWhere = strWhere & IIf(.Invoice <> "", " And A.实际票号=[5]", "")
       ' strWhere = strWhere & IIf(.strUnitName <> "", " And A.原因=[7]", "")
        strWherePati = IIf(.PatientINID <> "", " And A.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[3]) ", "")
        strWherePati = strWherePati & IIf(.lngUnitID <> 0, " And A.合同单位ID=[8]", "")
        strWherePati = strWherePati & IIf(.PatientName <> "", " And A.姓名=[6]", "")
        
        '40275
        If .bln仅欠款病人 Then
             strSQL = "" & _
                     "   Select Distinct A.病人id, A.住院号, A.姓名, A.性别, A.年龄, A.家庭地址,Q.名称 as 合约单位 " & _
                     "   From 病人信息 A,合约单位 Q," & _
                     "        (Select A.病人id, A.ID, Nvl(Sum(B.冲预交), 0) - Nvl(Max(J.金额), 0) As 欠款 " & _
                     "          From 病人结帐记录 A, 病人预交记录 B, 结算方式 C, 病人缴款对照 J " & _
                     "          Where A.收费时间 Between [1] And [2] And  A.记录状态 = 1 " & _
                     "                  And A.ID = B.结帐id And B.结算方式 = C.名称 " & _
                     "                  And  C.应收款 = 1 And A.ID = J.结帐id(+) " & strWhere & _
                     "          Group By A.病人id, A.ID " & _
                     "          Having Nvl(Sum(B.冲预交), 0) - Nvl(Max(J.金额), 0) > 0) B " & _
                     "   Where A.病人id = B.病人id And A.合同单位ID=Q.ID(+)" & strWherePati
        
        Else
                strSQL = "" & _
                "Select distinct A.病人id, A.住院号, A.姓名, A.性别, A.年龄,A.家庭地址,Q.名称 as 合约单位" & vbNewLine & _
                "From 病人信息 A,合约单位 Q," & vbNewLine & _
                "     ( Select Distinct A.病人id" & vbNewLine & _
                "       From 病人结帐记录 A, 病人预交记录 B, 结算方式 C" & vbNewLine & _
                "       Where A.收费时间 Between [1] And [2] And A.记录状态 = 1   " & _
                "                   And A.ID = B.结帐id And B.结算方式 = C.名称  " & _
                "                   And C.应收款 = 1" & vbNewLine & _
                                strWhere & _
                "       ) B" & vbNewLine & _
                "Where A.病人id = B.病人id  And A.合同单位ID=Q.ID(+) " & strWherePati
        End If
        If mblnNOMoved Then
            strSQL = strSQL & " Union All " & Replace(Replace(strSQL, "病人结帐记录", "H病人结帐记录"), "病人预交记录", "H病人预交记录")
        End If
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, Val(.PatientINID), .NO, .Invoice, .PatientName, .strUnitName, .lngUnitID)
    End With
  
    lvw.ListItems.Clear
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            Set objItem = lvw.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!姓名, IIf("" & rsTmp!性别 = "女", 2, 1), IIf("" & rsTmp!性别 = "女", 2, 1))
            objItem.SubItems(LVWCOL.C1住院号) = "" & rsTmp!住院号
            objItem.SubItems(LVWCOL.C2性别) = "" & rsTmp!性别
            objItem.SubItems(LVWCOL.C3年龄) = "" & NVL(rsTmp!年龄)
            objItem.SubItems(LVWCOL.C4家庭地址) = "" & NVL(rsTmp!家庭地址)
            objItem.SubItems(LVWCOL.C5合约单位) = "" & NVL(rsTmp!合约单位)
            rsTmp.MoveNext
        Next
        stbThis.Panels(2).Text = "共" & rsTmp.RecordCount & "名病人!"
        lvw.ListItems(1).Selected = True
        If Visible Then Call lvw_ItemClick(lvw.ListItems(1))
    Else
        stbThis.Panels(2).Text = "当前日期结帐期间内没有找到有应收款的病人!"
        Call ShowList
    End If
    
    Call zlCommFun.StopFlash
    Me.Refresh
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

Private Sub Form_Load()
    Dim i As Long
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Call RestoreWinState(Me, App.ProductName)
        
    '会调用click事件,但事件中已控制,此时不读数据
    mblnDo = False
    i = IIf(zlDatabase.GetPara("病人应收款页面", glngSys, mlngModul, "1") = "1", 1, 2)
    tbs.Tabs(i).Selected = True
    mblnDo = True
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '权限设置
    '暂无
    
    '缺省显示当天内结帐产生了应收款的病人
    mDatSys = zlDatabase.Currentdate
    With SQLCondition
        .DateE = mDatSys
        .DateB = CDate(Format(.DateE, "YYYY-MM-DD 00:00:01"))
        .Invoice = ""
        .NO = ""
        .PatientINID = ""
        frmDueFilter.dtpBegin.Value = .DateB
        frmDueFilter.dtpEnd.Value = .DateE
        lblDate.Caption = "日期:" & Format(.DateB, "YYYY-MM-DD") & "至" & Format(.DateE, "YYYY-MM-DD")
    End With
    mblnNOMoved = False
    
    Call LoadPatients
    Call SetHeader
    Call mshList_EnterCell
    
    mshBalance.Visible = (tbs.SelectedItem.Key = "Gathering"): mshInvoice.Visible = (tbs.SelectedItem.Key = "Gathering")
    
    
    '根据列表显示方式设置菜单
    Call SetView(lvw.View)
End Sub


Private Sub SetView(bytStyle As Byte)
'功能：调整列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub


Private Sub tbs_Click()
    Dim bln As Boolean
    
    If Not mblnDo Then Exit Sub
    If tbs.SelectedItem.Key = tbs.Tag Then Exit Sub
                
    bln = (tbs.SelectedItem.Key = "Gathering")
    mshBalance.Visible = bln: mshInvoice.Visible = bln
    Call Form_Resize
    
    Call ShowList
    
    tbs.Tag = tbs.SelectedItem.Key  '记录上一次的
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Mid(Item.Key, 2) = lvw.Tag Then Exit Sub
    
    tbs.Tag = ""
    Call tbs_Click
    
    lvw.Tag = Mid(Item.Key, 2)
    
    If Val(lvw.Tag) <> 0 Then stbThis.Panels(2).Text = "当前病人应收款余额:" & Format(GetPatientDue(Val(lvw.Tag)), "0.00")
    
    Call lvw.SetFocus
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


Private Sub SetActiveList(obj As Object)
    Select Case obj
        Case mshList
            mshList.BackColorSel = &HC0C0C0
            mshBalance.BackColorSel = &HE0E0E0
            mshInvoice.BackColorSel = &HE0E0E0
        Case mshBalance
            mshList.BackColorSel = &HE0E0E0
            mshBalance.BackColorSel = &HC0C0C0
            mshInvoice.BackColorSel = &HE0E0E0
        Case mshInvoice
            mshList.BackColorSel = &HE0E0E0
            mshBalance.BackColorSel = &HE0E0E0
            mshInvoice.BackColorSel = &HC0C0C0
    End Select
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, sngVsc As Single

    On Error Resume Next
    If WindowState = 1 Then Exit Sub
        
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    If mshInvoice.Visible Then sngVsc = 0.4
    
    lblDate.Left = Me.ScaleLeft
    lblDate.Top = Me.ScaleTop + cbrH + 50
    
    lvw.Left = Me.ScaleLeft
    lvw.Top = lblDate.Top + lblDate.Height
    lvw.Width = lblDate.Width
    lvw.Height = Me.ScaleHeight - staH - cbrH - lblDate.Height - 15
    
    pic.Left = lvw.Left + lvw.Width
    pic.Top = Me.ScaleTop + cbrH
    pic.Height = Me.ScaleHeight - cbrH - staH
    
    tbs.Left = pic.Left + pic.Width
    tbs.Top = Me.ScaleTop + cbrH
    tbs.Width = Me.ScaleWidth - lblDate.Width - pic.Width
        
    mshList.Left = pic.Left + pic.Width
    mshList.Top = tbs.Top + tbs.Height
    mshList.Width = Me.ScaleWidth - lblDate.Width - pic.Width
    mshList.Height = (Me.ScaleHeight - cbrH - staH - tbs.Height - IIf(mshInvoice.Visible, picHsc.Height, 15)) * (1 - sngVsc)
    
    picHsc.Left = mshList.Left
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Width = mshList.Width
    
    If mshBalance.Visible Then
        mshBalance.Top = picHsc.Top + picHsc.Height
        mshBalance.Left = mshList.Left
        mshBalance.Width = mshList.Width * 0.6
        mshBalance.Height = Me.ScaleHeight - cbrH - staH - tbs.Height - picHsc.Height - mshList.Height
        
        mshInvoice.Top = mshBalance.Top
        mshInvoice.Left = mshBalance.Left + mshBalance.Width + 15
        mshInvoice.Width = mshList.Width - mshBalance.Width - 15
        mshInvoice.Height = mshBalance.Height
    ElseIf mshInvoice.Visible Then
        mshInvoice.Top = picHsc.Top + picHsc.Height
        mshInvoice.Left = mshList.Left
        mshInvoice.Width = mshList.Width
        mshInvoice.Height = Me.ScaleHeight - cbrH - staH - tbs.Height - picHsc.Height - mshList.Height
    End If
    
    Me.Refresh
    mshList.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    Unload frmDueFilter
    Unload frmDueGo
    
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "病人应收款页面", tbs.SelectedItem.Index, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name & tbs.SelectedItem.Key)
        
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
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
        
        mlngCurRow = mshList.Row
        Set mshList.DataSource = Nothing
        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowList(True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If tbs.SelectedItem.Key = "Due" Then
        strHead = "单据号,4,850|票据号,4,1050|结帐时间,4,1850|结帐人,4,850|应收款额,7,1250|应收余额,7,1250|结帐ID,1,0"
    Else
        strHead = "单据号,4,850|收款人,4,800|收款时间,4,1850|记录状态,1,0|收款金额,7,1250"
    End If
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or tbs.SelectedItem.Key <> tbs.Tag Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or tbs.SelectedItem.Key <> tbs.Tag Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name & tbs.SelectedItem.Key)
        
        .RowHeight(0) = 320
    End With
End Sub

Private Sub ShowList(Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, bytFlag As Byte, lngCol As Long, lng病人ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
        
    If Not blnSort Then
        If Not lvw.SelectedItem Is Nothing Then lng病人ID = Val(Mid(lvw.SelectedItem.Key, 2))
        If lng病人ID <> 0 Then
            If tbs.SelectedItem.Key = "Due" Then
                '结帐作废之前必须先作废缴款登记单,清除[病人缴款对照]
                strSQL = "Select A.单据号, A.票据号, To_Char(A.结帐时间,'YYYY-MM-DD HH24:MI:SS') 结帐时间, A.结帐人, Ltrim(To_Char(Sum(B.冲预交),'999999999" & gstrDec & "')) 应收款额, Ltrim(To_Char(Sum(B.冲预交) - 冲应收,'999999999" & gstrDec & "')) 应收余额, A.ID As 结帐ID" & vbNewLine & _
                        "From (Select A.ID, A.单据号, A.票据号, A.结帐时间, A.结帐人, Nvl(Sum(B.金额), 0) 冲应收" & vbNewLine & _
                        "       From (Select ID, NO 单据号, 实际票号 票据号, 收费时间 结帐时间, 操作员姓名 结帐人" & vbNewLine & _
                        "              From 病人结帐记录" & vbNewLine & _
                        "              Where 病人id = [1] And 记录状态=1) A, 病人缴款对照 B" & vbNewLine & _
                        "       Where A.ID = B.结帐id(+)" & vbNewLine & _
                        "       Group By A.ID, A.单据号, A.票据号, A.结帐时间, A.结帐人) A, 病人预交记录 B, 结算方式 C" & vbNewLine & _
                        "Where A.ID = B.结帐id And B.结算方式 = C.名称 And C.应收款 = 1" & vbNewLine & _
                        "Group By A.单据号, A.票据号, A.结帐时间, A.结帐人, 冲应收, A.ID"
                If mblnNOMoved Then
                    strSQL = strSQL & " Union All " & Replace(Replace(strSQL, "病人结帐记录", "H病人结帐记录"), "病人预交记录", "H病人预交记录")
                End If
                strSQL = strSQL & " Order by 单据号 Desc"
            Else
                strSQL = "Select NO 单据号, 登记人 收款人, To_Char(登记时间,'YYYY-MM-DD HH24:MI:SS') 收款时间, 记录状态, Ltrim(To_Char(Sum(金额),'999999999" & gstrDec & "')) 收款金额" & vbNewLine & _
                        "From 病人缴款记录" & vbNewLine & _
                        "Where 病人id = [1]" & vbNewLine & _
                        "Group By NO, 登记人, 登记时间, 记录状态 Order By 单据号 Desc"
            End If
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
        Else
            Set mrsList = Nothing
        End If
    End If
    
    With mshList
        .Redraw = False
        .Clear: .Rows = 2
        If mrsList Is Nothing Then
            stbThis.Panels(2).Text = ""
        Else
            If mrsList.RecordCount = 0 Then
                If tbs.SelectedItem.Key = "Due" Then
                    stbThis.Panels(2).Text = "当前病人没有应收款结帐单据!"
                Else
                    stbThis.Panels(2).Text = "当前病人没有应收款缴款单据!"
                End If
            Else
                Set .DataSource = mrsList
                stbThis.Panels(2).Text = "当前病人共有" & mrsList.RecordCount & "张单据!"
            End If
        End If
        Call SetHeader
        
        If tbs.SelectedItem.Key = "Gathering" Then
            lngCol = GetColNum("记录状态")
            .ForeColor = ForeColor
            For i = 1 To .Rows - 1
                bytFlag = Val(.TextMatrix(i, lngCol))
                If bytFlag = 2 Or bytFlag = 3 Then
                    .Row = i
                    For j = 0 To .Cols - 1
                        .Col = j
                        .CellForeColor = IIf(bytFlag = 2, &HC0, &HC00000) '退款记录用红色,退过款的用蓝色
                    Next
                End If
            Next
        End If
        
        .Redraw = True
        .Row = IIf(blnSort, mlngCurRow, 1): .Col = 0: .ColSel = .Cols - 1
        '调用EnterCell事件,设置菜单和按钮,以及mshBalance,mshInvoice的可见性
        Call mshList_EnterCell
        If blnSort Then .TopRow = mlngCurRow
        
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetDetail(mshDetail As MSHFlexGrid)
    Dim strHead As String
    Dim i As Long
    
    If mshDetail Is mshBalance Then
        strHead = "结帐单,4,850|结帐票据,4,1050|结帐日期,4,1050|应收金额,7,850|冲应收,7,850"
    Else
        strHead = "结算方式,4,850|金额,7,850|结算号码,4,1200|备注,1,1600"
    End If
    
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
    End With
End Sub

Private Sub ShowDetail(strNO As String, mshDetail As MSHFlexGrid, Optional bytFlag As Byte)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If strNO <> "" Then
        If mshDetail Is mshBalance Then
            strSQL = "Select 结帐单, 结帐票据,To_Char(结帐时间,'YYYY-MM-DD') 结帐日期, Ltrim(To_Char(应收金额,'999999999" & gstrDec & "')) 应收金额, Ltrim(To_Char(冲应收,'999999999" & gstrDec & "')) 冲应收" & vbNewLine & _
                    "From (Select D.NO 结帐单, D.实际票号 结帐票据,D.收费时间 结帐时间, A.金额 冲应收, Sum(冲预交) 应收金额" & vbNewLine & _
                    "       From 病人缴款对照 A, 病人预交记录 B, 结算方式 C, 病人结帐记录 D" & vbNewLine & _
                    "       Where A.缴款单 = [1] And A.结帐id = B.结帐id And B.结算方式 = C.名称 And C.应收款 = 1 And B.结帐id = D.ID" & vbNewLine & _
                    "       Group By D.NO, D.实际票号,D.收费时间, A.金额)"
            If mblnNOMoved Then
                strSQL = strSQL & " Union All " & Replace(Replace(strSQL, "病人结帐记录", "H病人结帐记录"), "病人预交记录", "H病人预交记录")
            End If
            strSQL = strSQL & " Order by 结帐单"
        Else
            strSQL = "Select 结算方式, Ltrim(To_Char(金额,'999999999" & gstrDec & "')) 金额, 结算号 结算号码, 摘要 备注 From 病人缴款记录 Where NO = [1] And 记录状态=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytFlag)
    End If
            
    With mshDetail
        .Redraw = False
        .Clear: .Rows = 2
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then Set .DataSource = rsTmp
        End If
        Call SetDetail(mshDetail)
        .ForeColor = Me.ForeColor
        .Redraw = True
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
    End With
    
    If mshDetail Is mshBalance Then
        Call mshBalance_EnterCell
    Else
        Call mshInvoice_EnterCell
    End If
    
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

Private Function CheckNOMoved(strNO As String) As Boolean
'功能:根据缴款单据号检查对应的结帐单是否已转入到后备数据表

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select C.NO" & vbNewLine & _
            "From 病人缴款记录 A, 病人缴款对照 B, 病人结帐记录 C" & vbNewLine & _
            "Where A.NO = B.缴款单 And B.结帐id = C.ID And A.NO = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)

    If rsTmp.RecordCount > 0 Then
        CheckNOMoved = zlDatabase.NOMoved("病人结帐记录", rsTmp!NO, , , Me.Caption)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function






Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub tbs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then tbs.Tag = ""
End Sub
