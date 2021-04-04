VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmClientsParasSet 
   Caption         =   "站点参数配置"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10290
   Icon            =   "frmClientsParasSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10290
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   2655
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":030A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":052A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":0744
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":089E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":0ABE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   2820
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":0CDE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":0EFE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":1118
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":1272
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsParasSet.frx":148E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2040
      Left            =   4815
      TabIndex        =   30
      Top             =   5730
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   3598
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilt32"
      SmallIcons      =   "ilt16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "站点"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "用户"
         Object.Width           =   3705
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "恢复类型"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "说明"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4950
      Index           =   1
      Left            =   3705
      TabIndex        =   19
      Top             =   1095
      Width           =   7455
      Begin VB.CheckBox chk分站点 
         BackColor       =   &H8000000C&
         Caption         =   "用户分站点"
         Height          =   225
         Left            =   5985
         TabIndex        =   24
         Top             =   135
         Width           =   1395
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   ">"
         Height          =   300
         Index           =   0
         Left            =   3390
         TabIndex        =   23
         Top             =   510
         Width           =   615
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   ">>"
         Height          =   300
         Index           =   1
         Left            =   3390
         TabIndex        =   22
         Top             =   900
         Width           =   615
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "<"
         Height          =   300
         Index           =   2
         Left            =   3390
         TabIndex        =   21
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "<<"
         Height          =   300
         Index           =   3
         Left            =   3390
         TabIndex        =   20
         Top             =   1605
         Width           =   615
      End
      Begin MSComctlLib.ImageList ilt32 
         Left            =   100
         Top             =   1725
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":16AE
               Key             =   "User"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":19C8
               Key             =   "Client"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":2492
               Key             =   "Scheame"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilt16 
         Left            =   3000
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":27AC
               Key             =   "User"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":2AC6
               Key             =   "Client"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientsParasSet.frx":3590
               Key             =   "Scheame"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwUsered 
         Height          =   4500
         Left            =   4050
         TabIndex        =   25
         Top             =   405
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   7938
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilt32"
         SmallIcons      =   "ilt16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "用户名"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "姓名"
            Object.Tag             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "人员性质"
            Object.Tag             =   "人员性质"
            Text            =   "人员性质"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "所属部门"
            Object.Tag             =   "所属部门"
            Text            =   "所属部门"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwUser 
         Height          =   4500
         Left            =   75
         TabIndex        =   26
         Top             =   405
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   7938
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilt32"
         SmallIcons      =   "ilt16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "用户名"
            Object.Tag             =   "用户名"
            Text            =   "用户名"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "姓名"
            Object.Tag             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "人员性质"
            Object.Tag             =   "人员性质"
            Text            =   "人员性质"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "所属部门"
            Object.Tag             =   "所属部门"
            Text            =   "所属部门"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblForColor 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "未选用户"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   28
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblForColor 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "已选用户"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   3
         Left            =   4110
         TabIndex        =   27
         Top             =   165
         Width           =   780
      End
      Begin VB.Label lblBack 
         BackColor       =   &H8000000C&
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   29
         Top             =   120
         Width           =   7380
      End
   End
   Begin VB.Frame fra 
      Height          =   4965
      Index           =   0
      Left            =   3540
      TabIndex        =   9
      Top             =   1215
      Width           =   7455
      Begin VB.CommandButton cmdClientSel 
         Caption         =   ">"
         Height          =   300
         Index           =   0
         Left            =   3390
         TabIndex        =   13
         Top             =   690
         Width           =   615
      End
      Begin VB.CommandButton cmdClientSel 
         Caption         =   ">>"
         Height          =   300
         Index           =   1
         Left            =   3390
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdClientSel 
         Caption         =   "<"
         Height          =   300
         Index           =   2
         Left            =   3390
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdClientSel 
         Caption         =   "<<"
         Height          =   300
         Index           =   3
         Left            =   3390
         TabIndex        =   10
         Top             =   1785
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwCliented 
         Height          =   4500
         Left            =   4095
         TabIndex        =   14
         Top             =   405
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   7938
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilt32"
         SmallIcons      =   "ilt16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "站点"
            Object.Width           =   4304
         EndProperty
      End
      Begin MSComctlLib.ListView lvwClient 
         Height          =   4500
         Left            =   45
         TabIndex        =   15
         Top             =   405
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   7938
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilt32"
         SmallIcons      =   "ilt16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "站点"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblForColor 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "未选站点"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   165
         Width           =   780
      End
      Begin VB.Label lblForColor 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "已选站点"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   1
         Left            =   4140
         TabIndex        =   16
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblBack 
         BackColor       =   &H8000000C&
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   7365
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6180
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClientsParasSet.frx":38AA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13097
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1005
      Top             =   6450
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
            Picture         =   "frmClientsParasSet.frx":413E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw方案 
      Height          =   5190
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   9155
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilt32"
      SmallIcons      =   "ilt16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "方案"
         Text            =   "方案"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "上传站点"
         Text            =   "上传站点"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "上传用户"
         Text            =   "上传用户"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "方案描述"
         Text            =   "方案描述"
         Object.Width           =   4304
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10290
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   930
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除方案"
               Object.Tag             =   "删除方案"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存设置"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageKey        =   "Refresh"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab stb 
      Height          =   5370
      Left            =   2835
      TabIndex        =   3
      Top             =   765
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   9472
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "公共参数(&0)"
      TabPicture(0)   =   "frmClientsParasSet.frx":4458
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "私有参数(&1)"
      TabPicture(1)   =   "frmClientsParasSet.frx":4474
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "已恢复(&2)"
      TabPicture(2)   =   "frmClientsParasSet.frx":4490
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblClientEd"
      Tab(2).ControlCount=   1
      Begin VB.Label lblClientEd 
         AutoSize        =   -1  'True
         Caption         =   "方案信息:[02]测试方案"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   -71430
         TabIndex        =   8
         Top             =   75
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "分站点将以“公共参数”所涉及的站点为准。"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3540
         TabIndex        =   7
         Top             =   60
         Width           =   4020
      End
   End
   Begin VB.Image ImgLine_S 
      Height          =   4980
      Left            =   2640
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   45
   End
   Begin VB.Label lblForColor 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "方案选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   180
      Index           =   4
      Left            =   75
      TabIndex        =   0
      Top             =   765
      Width           =   780
   End
   Begin VB.Label lblBack 
      BackColor       =   &H8000000C&
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   1
      Top             =   720
      Width           =   2565
   End
End
Attribute VB_Name = "frmClientsParasSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mBlnFirst As Boolean
Dim mblnChange As Boolean
Dim mrs人员性质 As ADODB.Recordset

Private Sub SetCtlEnable()
    '---------------------------------------------------------------------------------------------------------
    '功能:设置控制的Enable属性
    '---------------------------------------------------------------------------------------------------------
    Dim intValue As Integer
    Me.cmdClientSel(0).Enabled = Not lvwClient.SelectedItem Is Nothing
    Me.cmdClientSel(1).Enabled = lvwClient.ListItems.Count <> 0
    
    Me.cmdClientSel(2).Enabled = Not lvwCliented.SelectedItem Is Nothing
    Me.cmdClientSel(3).Enabled = lvwCliented.ListItems.Count <> 0
    Me.cmdSel(0).Enabled = Not lvwUser.SelectedItem Is Nothing
    Me.cmdSel(1).Enabled = lvwUser.ListItems.Count <> 0
    
    Me.cmdSel(2).Enabled = Not lvwUsered.SelectedItem Is Nothing
    Me.cmdSel(3).Enabled = lvwUsered.ListItems.Count <> 0
    Me.chk分站点.Enabled = lvwUsered.ListItems.Count <> 0
    tlbThis.Buttons("保存").Enabled = (lvwCliented.ListItems.Count <> 0 Or Me.lvwUsered.ListItems.Count <> 0) And mblnChange
    tlbThis.Buttons("删除").Enabled = Not lvw方案.SelectedItem Is Nothing
End Sub

  

Private Sub chk分站点_Click()
    Dim lvwItem As ListItem
    Dim bln分站点 As Boolean
    If chk分站点.Value = 2 Then Exit Sub
    bln分站点 = Me.chk分站点.Value = 1
    For Each lvwItem In Me.lvwUsered.ListItems
        lvwItem.Checked = bln分站点
    Next
    mblnChange = True
    Call SetCtlEnable
End Sub

Private Sub cmdClientSel_Click(Index As Integer)
    Dim intItem As Integer
    Dim lvwItem As ListItem
    Dim lvwNewItem As ListItem
    Select Case Index
        Case 0 '单选:增加指定的置换人员
            If lvwClient.SelectedItem Is Nothing Then Exit Sub
            Set lvwItem = lvwClient.SelectedItem
            If lvwAddItem(lvwCliented, lvwItem) = False Then Exit Sub
            Call LvwReomveItem(lvwClient)
            mblnChange = True
        Case 1 '全选
            For intItem = 1 To Me.lvwClient.ListItems.Count
                Set lvwItem = lvwClient.ListItems(intItem)
                If lvwAddItem(lvwCliented, lvwItem) = False Then Exit Sub
            Next
            Me.lvwClient.ListItems.Clear
            mblnChange = True
        Case 2 '移除指定人员
            If lvwCliented.SelectedItem Is Nothing Then Exit Sub
            Set lvwItem = lvwCliented.SelectedItem
            
            If lvwAddItem(lvwClient, lvwItem) = False Then Exit Sub
            Call LvwReomveItem(lvwCliented)
            mblnChange = True
        Case 3 '移除所有已选人员
            For intItem = 1 To Me.lvwCliented.ListItems.Count
                Set lvwItem = lvwCliented.ListItems(intItem)
                If lvwAddItem(lvwClient, lvwItem) = False Then Exit Sub
            Next
            Me.lvwCliented.ListItems.Clear
            mblnChange = True
    End Select
    SetCtlEnable
End Sub

 

Private Sub cmdSel_Click(Index As Integer)
    Dim intItem As Integer
    Dim lvwItem As ListItem
    Dim lvwNewItem As ListItem
    Select Case Index
        Case 0 '单选:增加指定的置换人员
            If lvwUser.SelectedItem Is Nothing Then Exit Sub
            Set lvwItem = lvwUser.SelectedItem
            If lvwAddItem(lvwUsered, lvwItem) = False Then Exit Sub
            Call LvwReomveItem(lvwUser)
        Case 1 '全选
            For intItem = 1 To Me.lvwUser.ListItems.Count
                Set lvwItem = lvwUser.ListItems(intItem)
                If lvwAddItem(lvwUsered, lvwItem) = False Then Exit Sub
            Next
            Me.lvwUser.ListItems.Clear
        Case 2 '移除指定人员
            If lvwUsered.SelectedItem Is Nothing Then Exit Sub
            Set lvwItem = lvwUsered.SelectedItem
            
            If lvwAddItem(lvwUser, lvwItem) = False Then Exit Sub
            Call LvwReomveItem(lvwUsered)
        Case 3 '移除所有已选人员
            For intItem = 1 To Me.lvwUsered.ListItems.Count
                Set lvwItem = lvwUsered.ListItems(intItem)
                If lvwAddItem(lvwUser, lvwItem) = False Then Exit Sub
            Next
            Me.lvwUsered.ListItems.Clear
    End Select
    SetCtlEnable
End Sub

Private Sub RemoveLvwedToLvw()
    '功能:将所有已选的移到未选中去.
    Dim lvwItem As ListItem
    Err = 0: On Error Resume Next
    For Each lvwItem In Me.lvwCliented.ListItems
        If lvwAddItem(lvwClient, lvwItem) = False Then Exit Sub
    Next
    Me.lvwCliented.ListItems.Clear
    
    For Each lvwItem In Me.lvwUsered.ListItems
        If lvwAddItem(lvwUser, lvwItem) = False Then Exit Sub
    Next
    Me.lvwUsered.ListItems.Clear
End Sub
Private Function lvwAddItem(ByVal lvw As ListView, ByVal lvwItem As ListItem, Optional blnRed As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加lvw控件的listItem值
    '参数:lvw-需要增加的Lvw
    '     lvwItem-需要增加的lvwitem对象
    '返回:增加成功返回true,否则返回false
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim lvwNewItem As ListItem
    Dim intCount As Integer, i As Integer
    intCount = lvw.ColumnHeaders.Count - 1
    
    Err = 0: On Error GoTo errHand:
    Set lvwNewItem = lvw.ListItems.Add(, lvwItem.Key, lvwItem.Text, lvwItem.Icon, lvwItem.SmallIcon)
    
    lvwNewItem.Checked = lvwItem.Checked
    If blnRed Then
        lvwNewItem.ForeColor = vbRed
    End If
    For i = 1 To intCount
        lvwNewItem.SubItems(i) = lvwItem.SubItems(i)
    Next
lvwAddItem = True
    Exit Function
errHand:
    MsgBox "错误号:" & Err.Number & vbCrLf & "错误描述:" & Err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Function
Private Sub LvwReomveItem(ByVal lvw As ListView)
        '功能:选择一下Select
        '参数:lvw-移出的lvw控件
        Dim lngIndex As Long
        Dim strKey As String
        
        Err = 0: On Error Resume Next
        lngIndex = lvw.SelectedItem.Index
       ' strKey = lvw.SelectedItem.Key
        lvw.ListItems.Remove lngIndex
        
        If lvw.ListItems.Count <= 0 Then
             Exit Sub
        Else
            lvw.ListItems(lngIndex).Selected = True
            If Err <> 0 Then
                If lngIndex - 1 >= 0 Then
                    lvw.ListItems(lngIndex - 1).Selected = True
                Else
                    lvw.ListItems(1).Selected = True
                End If
            End If
        End If
End Sub
Private Sub Form_Activate()
    If mBlnFirst = False Then Exit Sub
    mBlnFirst = False
    
    If InitCard() = False Then Unload Me: Exit Sub
    Me.lvw方案.SetFocus
    mblnChange = False
    SetShowMode
    Call SetCtlEnable
End Sub

Private Sub Form_Load()
    Set mrs人员性质 = New ADODB.Recordset

    mBlnFirst = True
End Sub
  
Public Sub ShowEdit()
    '-------------------------------------------------------------------------------
    '--功能：显示设置信息
    '-------------------------------------------------------------------------------
    Me.Show 1, frmMDIMain
End Sub
Private Function Get人员性质(ByVal lng人员id As Long) As String
    '-----------------------------------------------------------------------------------------------
    '功能:获取人员性质
    '返回:人员性质说明，以逗号分离,比如:医生,护士
    '编制:刘兴宏
    '日期:2007/09/10
    '-----------------------------------------------------------------------------------------------
    Dim strTemp As String
    If lng人员id = 0 Then Exit Function
    strTemp = ""
    mrs人员性质.Filter = "人员id=" & lng人员id
    With mrs人员性质
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strTemp = strTemp & "," & Nvl(!人员性质)
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    End With
    Get人员性质 = strTemp
End Function
Private Function InitCard() As Boolean
    '功能:初始化卡片
    '返回:初始成功,返回true,否则返回False
    Dim lng方案号 As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    
    InitCard = False
    Err = 0: On Error Resume Next
        
    Set mrs人员性质 = New ADODB.Recordset
    
    gstrSQL = "Select 人员id,人员性质 From 人员性质说明"
    Call OpenRecordset(mrs人员性质, gstrSQL, Me.Caption)
    If Err <> 0 Then
        MsgBox "你不具备操作此功能的权限,请找系统管理员将“人员性质说明”授权给你!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
            
    '---加载用户
    '    strSQL = "" & _
    '        "   Select distinct 用户名" & _
    '        "   From 上机人员表 "
    
    strSQL = "" & _
    "   Select Distinct A.用户名, B.姓名, C.部门id, c.名称 as 所属部门,A.人员ID " & _
    "   From 上机人员表 A, 人员表 B, " & _
    "     (Select C.人员id, C.部门id, M.名称 From 部门人员 C, 部门表 M Where C.部门id = M.ID And 缺省 = 1) c " & _
    "   Where A.人员id = B.ID And A.人员id = C.人员id(+)"
    
    lvwUser.ListItems.Clear
    lvwUsered.ListItems.Clear
    Call OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With rsTmp
        If Err <> 0 Then
            MsgBox "你不具备操作此功能的权限,请找系统管理员将上机人员表授权给你!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        
        Do While Not .EOF
            Set objItem = lvwUser.ListItems.Add(, "K" & Nvl(!用户名), Nvl(!用户名), "User", "User")
            objItem.SubItems(1) = Nvl(rsTmp!姓名)
            objItem.SubItems(2) = Get人员性质(Val(rsTmp!人员id))
            objItem.SubItems(3) = Nvl(!所属部门)
            .MoveNext
        Loop
    End With
    
    '---加载站点

    Me.lvwClient.ListItems.Clear
    Me.lvwCliented.ListItems.Clear
    Dim lvwItem As ListItem
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client_station")
    
    With rsTmp
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Set lvwItem = Me.lvwClient.ListItems.Add(, "K" & Nvl(!工作站), Nvl(!站点), "Client", "Client")
            lvwItem.Tag = Nvl(!工作站)
            .MoveNext
        Loop
    End With
    If Me.lvwClient.ListItems.Count = 0 Then
        MsgBox "没有可选择的站点,请设置站点!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    '--加载方案
    '确定最好一次设置的方案
    
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Project_no")
    With rsTmp
        If .EOF Then
            lng方案号 = -99
        Else
            lng方案号 = Val(Nvl(!方案号))
        End If
    End With
    
   Call Load恢复方案(lng方案号)
    
    
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client_scheme")
    With rsTmp
        If .RecordCount <= 0 Then
            MsgBox "没有可选择的方案,请在客户端程序设置方案并上传参数!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
     
        Do While Not .EOF
            Set lvwItem = Me.lvw方案.ListItems.Add(, "K" & !方案号, Nvl(!方案名称), "Scheame", "Scheame")
            lvwItem.SubItems(1) = Nvl(!工作站)
            lvwItem.SubItems(2) = Nvl(!用户名)
            lvwItem.SubItems(3) = Nvl(!方案描述)
            If Val(Nvl(!方案号)) = lng方案号 Then
                lvwItem.ForeColor = vbRed
                lvwItem.Selected = True
                lvwItem.ListSubItems(1).ForeColor = vbRed
                lvwItem.ListSubItems(2).ForeColor = vbRed
                lvwItem.ListSubItems(3).ForeColor = vbRed
                lvwItem.EnsureVisible
                
            Else
                lvwItem.ForeColor = &H80000008
                lvwItem.ListSubItems(1).ForeColor = &H80000008
                lvwItem.ListSubItems(2).ForeColor = &H80000008
                lvwItem.ListSubItems(3).ForeColor = &H80000008
            End If
            .MoveNext
        Loop
    End With
    If Not Me.lvw方案.SelectedItem Is Nothing Then
        lvw方案_ItemClick Me.lvw方案.SelectedItem
    End If
    SetCtlEnable
    InitCard = True
End Function
Private Sub Load恢复方案(ByVal lng方案号 As Long)
    '功能:加载指定的方案号
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lvwItem As ListItem
    Err = 0: On Error Resume Next
    '加载已经恢复的信息
    
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Resile", lng方案号, 0)
    With rsTmp
        Dim i As Long
        If .RecordCount > 0 Then
            lblClientEd.Caption = Nvl(!方案名称)
            lblClientEd.Tag = lng方案号
        Else
            lblClientEd.Caption = "不存在恢复方案"
        End If
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            i = i + 1
            Set lvwItem = Me.lvwList.ListItems.Add(, "K" & i, Nvl(!工作站), "Client", "Client")
            
            If Nvl(!工作站) = "" And Nvl(!用户名) <> "" Then
                '表示不分站升级
                lvwItem.SubItems(3) = "不分站点恢复私有参数"
            ElseIf Nvl(!工作站) <> "" And Nvl(!用户名) = "" Then
                lvwItem.SubItems(3) = "恢复公共参数"
            Else
                lvwItem.SubItems(3) = "恢复私有参数"
            End If
            lvwItem.SubItems(1) = Nvl(!用户名, " ")
            
            If Val(Nvl(!恢复标志)) = 1 Then
                lvwItem.SubItems(2) = "未恢复"
                lvwItem.ForeColor = &H80000008
                lvwItem.ListSubItems(1).ForeColor = &H80000008
                lvwItem.ListSubItems(2).ForeColor = &H80000008
                lvwItem.ListSubItems(3).ForeColor = &H80000008
            Else
                lvwItem.SubItems(2) = "已恢复"
                lvwItem.ForeColor = vbRed
                lvwItem.ListSubItems(1).ForeColor = vbRed
                lvwItem.ListSubItems(2).ForeColor = vbRed
                lvwItem.ListSubItems(3).ForeColor = vbRed
            End If
            .MoveNext
        Loop
    End With
    For Each lvwItem In lvw方案.ListItems
        If lvwItem.Key = "K" & lng方案号 Then
            lvwItem.ForeColor = vbRed
            lvwItem.ListSubItems(1).ForeColor = vbRed
            lvwItem.ListSubItems(2).ForeColor = vbRed
            lvwItem.ListSubItems(3).ForeColor = vbRed
        Else
            lvwItem.ForeColor = &H80000008
            lvwItem.ListSubItems(1).ForeColor = &H80000008
            lvwItem.ListSubItems(2).ForeColor = &H80000008
            lvwItem.ListSubItems(3).ForeColor = &H80000008
        End If
    Next
End Sub

Private Sub Form_Resize()
    Dim sngcbrHeight As Single, sngSbrHeight As Single
    sngSbrHeight = IIf(stbThis.Visible, stbThis.Height, 0)
    sngcbrHeight = IIf(cbrThis.Visible, cbrThis.Height, 0)
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 6915 Then
        Me.Height = 6915
    End If
    If Me.Width < 10380 Then
        Me.Width = 10380
    End If
    With ImgLine_S
        .Top = ScaleTop + sngcbrHeight
        .Height = IIf(ScaleHeight - sngcbrHeight - sngSbrHeight - ScaleTop < 0, 0, ScaleHeight - sngcbrHeight - sngSbrHeight - ScaleTop)
    End With
    
    With stb
        .Left = ImgLine_S.Left + ImgLine_S.Width
        .Top = ImgLine_S.Top
        .Height = ImgLine_S.Height
        .Width = IIf(ScaleWidth - .Left < 0, 0, ScaleWidth - .Left)
    End With

'
'    With stb
'        .Height = ScaleHeight - .Top - stbThis.Height
'        .Width = ScaleWidth - .Left - 50
'    End With
    With Me.fra(0)
        .Left = stb.Left + 50
        .Top = stb.Top + stb.TabHeight + 25
        .Height = stb.Height - stb.TabHeight - 100
        .Width = stb.Width - 100
    End With
    With fra(1)
        .Left = Me.fra(0).Left
        .Top = Me.fra(0).Top
        .Height = stb.Height - stb.TabHeight - 100
        .Width = stb.Width - 100
        Me.lvwList.Top = .Top
        Me.lvwList.Left = .Left
        Me.lvwList.Width = .Width
        Me.lvwList.Height = .Height
    End With
    Dim sngWidth As Single
    Dim sngHeight As Single
    sngWidth = (fra(0).Width - cmdSel(0).Width - 200) \ 2
    sngHeight = fra(0).Height - lvwClient.Top - 50
    
    lblBack(0).Width = fra(0).Width - lblBack(0).Left - 50
    lblBack(1).Width = lblBack(0).Width
    With lvwClient
        .Width = sngWidth
        .Height = sngHeight
        lvwUser.Left = .Left
        lvwUser.Width = .Width
        lvwUser.Height = .Height
        lvwUser.Top = .Top
    End With
    With cmdSel(0)
        .Left = lvwClient.Left + sngWidth + 50
        cmdSel(1).Left = .Left
        cmdSel(2).Left = .Left
        cmdSel(3).Left = .Left
        cmdClientSel(0).Left = .Left
        cmdClientSel(1).Left = .Left
        cmdClientSel(2).Left = .Left
        cmdClientSel(3).Left = .Left
    End With
    lblForColor(1).Left = cmdSel(0).Left + cmdSel(0).Width + 50
    lblForColor(3).Left = lblForColor(1).Left
    
    With lvw方案
        .Left = ScaleLeft
        .Height = ScaleHeight - .Top - stbThis.Height
        .Width = ImgLine_S.Left - 15
    End With
    With lvwCliented
        .Width = sngWidth
        .Height = sngHeight
        .Left = lblForColor(1).Left
        lvwUsered.Left = .Left
        lvwUsered.Width = .Width
        lvwUsered.Height = .Height
    End With
    chk分站点.Left = lblBack(1).Width - chk分站点.Width
    chk分站点.Top = lblBack(1).Top + 10
    lblBack(2).Width = lvw方案.Width - 15
End Sub

Private Sub lvwClient_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Static intPreCol As Integer
     lvwClientSort ColumnHeader, lvwClient, intPreCol
End Sub
Private Sub lvwClientSort(ByVal ColumnHeader As MSComctlLib.ColumnHeader, ByVal lvw As ListView, ByRef intPreCol As Integer)
    '对lvw排序
    '排序
    If intPreCol = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        intPreCol = ColumnHeader.Index - 1
        lvw.SortKey = intPreCol
        lvw.SortOrder = lvwAscending
    End If
    
End Sub
Private Sub lvwClient_DblClick()
    If lvwClient.SelectedItem Is Nothing Then Exit Sub
    mblnChange = True
    cmdClientSel_Click 0
End Sub

Private Sub lvwClient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetCtlEnable
End Sub

 
Private Sub lvwClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub lvwCliented_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Static intPreCol As Integer
     lvwClientSort ColumnHeader, lvwCliented, intPreCol

End Sub

Private Sub lvwCliented_DblClick()
    If lvwCliented.SelectedItem Is Nothing Then Exit Sub
    mblnChange = True
    cmdClientSel_Click 2

End Sub

Private Sub lvwCliented_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetCtlEnable
End Sub

 

Private Sub lvwCliented_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 1
        stb.SetFocus
    End If
End Sub

Private Sub LvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Static intPreCol As Integer
     lvwClientSort ColumnHeader, lvwList, intPreCol
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Static intPreCol As Integer
     lvwClientSort ColumnHeader, lvwUser, intPreCol

End Sub

Private Sub lvwUser_DblClick()
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    mblnChange = True
    cmdSel_Click 0

End Sub

Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetCtlEnable
End Sub

Private Sub lvwUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub

Private Sub lvwUsered_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     Static intPreCol As Integer
     lvwClientSort ColumnHeader, lvwUsered, intPreCol

End Sub

Private Sub lvwUsered_DblClick()
    If lvwUsered.SelectedItem Is Nothing Then Exit Sub
    mblnChange = True
    cmdSel_Click 2
End Sub

Private Sub lvwUsered_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Me.chk分站点.Value = 2
        mblnChange = True
        SetCtlEnable
End Sub

Private Sub lvwUsered_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Call SetCtlEnable
End Sub

Private Sub lvwUsered_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 2
        stb.SetFocus
    End If
End Sub

Private Sub lvw方案_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intPreCol As Integer
    lvwClientSort ColumnHeader, lvw方案, intPreCol
End Sub

Private Sub lvw方案_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Call LoadScremeSet(Val(Mid(Item.Key, 2)))
        Call SetCtlEnable
End Sub
Private Function LoadScremeSet(ByVal lng方案号 As Long) As Boolean
    '功能:加载方案设置情况
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lvwItem As ListItem
    Call cmdClientSel_Click(3)
    Call cmdSel_Click(3)
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Resile", lng方案号, 1)
    With rsTemp
        Do While Not .EOF
            Err = 0: On Error Resume Next
            Set lvwItem = lvwClient.ListItems("K" & Nvl(!工作站))
            lvwItem.Selected = True
            
            If Err = 0 Then
                If lvwAddItem(lvwCliented, lvwItem, Val(Nvl(!恢复标志)) = 0) = False Then Exit Function
                Call LvwReomveItem(lvwClient)
            End If
            .MoveNext
        Loop
    End With
            
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Resile", lng方案号, 2)
    Dim str用户名 As String
    With rsTemp
        Do While Not .EOF
            Err = 0: On Error Resume Next
            If str用户名 <> Nvl(!用户名) Then
                Set lvwItem = lvwUser.ListItems("K" & Nvl(!用户名))
                If Err = 0 Then
                    '相同:判断是否为站点为NULL
                    If Nvl(!工作站) = "" Then
                        lvwItem.Checked = False
                    Else
                        lvwItem.Checked = True
                    End If
                    lvwItem.Selected = True
                    '分不出颜色
                    If lvwAddItem(lvwUsered, lvwItem) = False Then Exit Function
                    Call LvwReomveItem(lvwUser)
                End If
            Else
                '相同:判断是否为站点为NULL
                If Nvl(!工作站) = "" Then
                    Set lvwItem = lvwUsered.ListItems("K" & Nvl(!用户名))
                    lvwItem.Checked = False
                End If
            End If
            .MoveNext
        Loop
    End With
    mblnChange = False
End Function

Private Sub lvw方案_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub stb_Click(PreviousTab As Integer)
    SetShowMode
End Sub

Private Sub stb_GotFocus()
    SetShowMode
    Select Case stb.Tab
    Case 0
        If lvwClient.Enabled Then lvwClient.SetFocus
    Case 1
       If lvwUser.Enabled Then lvwUser.SetFocus
    Case 2
      If lvwList.Enabled Then lvwList.SetFocus
    End Select
End Sub
 

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strSQL As String
    Dim lng方案号 As Long
    Select Case Button.Key
    Case "保存"
        If SaveDataSet = False Then Exit Sub
        MsgBox "保存成功!", vbInformation + vbDefaultButton1, gstrSysName
        Call Load恢复方案(Val(Mid(lvw方案.SelectedItem.Key, 2)))
        mblnChange = False
        Call SetCtlEnable
    Case "删除"
        If lvw方案.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("你真的要删除“" & Me.lvw方案.SelectedItem.Text & "”的方案吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        lng方案号 = Val(Mid(lvw方案.SelectedItem.Key, 2))
        strSQL = "Delete zlClientScheme where 方案号=" & lng方案号
        gcnOracle.Execute strSQL
        LvwReomveItem lvw方案
        
        If lvw方案.ListItems.Count = 0 Or lng方案号 = Val(lblClientEd.Tag) Then
                Me.lvwList.ListItems.Clear
                Me.lblClientEd.Caption = "不存在恢复方案"
        End If
        Call SetCtlEnable
    Case "刷新"
        If InitCard() = False Then Exit Sub
        If lvw方案.Enabled Then Me.lvw方案.SetFocus
        mblnChange = False
        SetShowMode
        Call SetCtlEnable
    Case "帮助"
        ShowHelp Me.hwnd, "zl9svrtools\" & Me.name    '
    Case "Exit"
        Unload Me
    End Select
End Sub

Private Function SaveDataSet() As Boolean
    '功能:保存方案设置
    Dim lng方案号 As Long
    Dim clldata As New Collection
    Dim strSQL As String
    SaveDataSet = False
    If lvw方案.SelectedItem Is Nothing Then Exit Function
    lng方案号 = Val(Mid(lvw方案.SelectedItem.Key, 2))
    '保存用户
    
    
    '删除此方案的设置
    strSQL = "Delete Zlclientparaset "
    clldata.Add strSQL
'    方案号   number(18),
'    工作站 varchar2(50),
'    用户名 varchar2(20),
'    恢复标志 number(2))
    Dim lvwItem As ListItem
    Dim lvwItem1 As ListItem
    For Each lvwItem In Me.lvwUsered.ListItems
        If lvwItem.Checked Then
            '分站点:
            If lvwCliented.ListItems.Count = 0 Then
                MsgBox "你选择的用户限制了站点的,但未选择站点，请选择站点!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            
            For Each lvwItem1 In Me.lvwCliented.ListItems
                strSQL = "Insert into Zlclientparaset(方案号,工作站,用户名,恢复标志) values ("
                strSQL = strSQL & "" & lng方案号 & ","
                strSQL = strSQL & "'" & Mid(lvwItem1.Key, 2) & "',"
                strSQL = strSQL & "'" & Mid(lvwItem.Key, 2) & "',1)"
                clldata.Add strSQL
            Next
        Else
            '不分站点
            strSQL = "Insert into Zlclientparaset(方案号,工作站,用户名,恢复标志) values ("
            strSQL = strSQL & "" & lng方案号 & ","
            strSQL = strSQL & "NULL,"
            strSQL = strSQL & "'" & Mid(lvwItem.Key, 2) & "',1)"
            clldata.Add strSQL
        End If
    Next
    For Each lvwItem1 In Me.lvwCliented.ListItems
        strSQL = "Insert into Zlclientparaset(方案号,工作站,用户名,恢复标志) values ("
        strSQL = strSQL & "" & lng方案号 & ","
        strSQL = strSQL & "'" & Mid(lvwItem1.Key, 2) & "',"
        strSQL = strSQL & "NULL,1)"
        clldata.Add strSQL
    Next
    Err = 0: On Error GoTo errHand:
    Call InsertToDatabase(clldata)
    SaveDataSet = True
    Exit Function
errHand:
    gcnOracle.RollbackTrans
    If MsgBox("数据保存发生错误," & vbCrLf & "错误号:" & Err.Number & vbCrLf & "错误描述:" & Err.Description, vbRetryCancel + vbDefaultButton1, gstrSysName) = vbRetry Then Resume
End Function
Private Sub InsertToDatabase(ByVal clldata As Collection)
    '功能:保存数据
    Dim i As Long
    gcnOracle.BeginTrans
    For i = 1 To clldata.Count
        gcnOracle.Execute clldata(i)
    Next
    gcnOracle.CommitTrans
End Sub
Private Sub SetShowMode()
    '功能:设置当前的显示模式
    Select Case Me.stb.Tab
    Case 0
        fra(0).Visible = True
        fra(1).Visible = False
        lvwList.Visible = False
    Case 1
        fra(0).Visible = False
        fra(1).Visible = True
        lvwList.Visible = False
    Case Else
        fra(0).Visible = False
        fra(1).Visible = False
        lvwList.Visible = True
    End Select
End Sub
Private Sub ImgLine_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If ImgLine_S.Left + X < 1500 Then Exit Sub
    If Me.ScaleWidth - (ImgLine_S.Left + X) < 7380 Then Exit Sub
    ImgLine_S.Left = ImgLine_S.Left + X
    Form_Resize
End Sub

