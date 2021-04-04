VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDeptManage 
   Caption         =   "部门管理"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmDeptManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Tag             =   "可变化的"
   Begin VB.PictureBox picList 
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   840
      Width           =   2775
      Begin VB.PictureBox picSplit2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   255
         ScaleWidth      =   3720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Width           =   3720
         Begin VB.Label lbl病区科室 
            Caption         =   "  病区科室对应关系"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   60
            Width           =   1575
         End
      End
      Begin MSComctlLib.TreeView tvwMain_S 
         Height          =   2865
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5054
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvwDept 
         Height          =   1140
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2011
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin XtremeSuiteControls.TabControl tbcDetails 
      Height          =   855
      Left            =   4560
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
      _Version        =   589884
      _ExtentX        =   2143
      _ExtentY        =   1508
      _StockProps     =   64
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3720
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   5280
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3000
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   635
      SimpleText      =   $"frmDeptManage.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeptManage.frx":0351
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
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
   Begin MSComctlLib.ListView lvw部门性质_S 
      Height          =   1095
      Left            =   4470
      TabIndex        =   3
      Top             =   3420
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "工作性质"
         Object.Tag             =   "工作性质"
         Text            =   "工作性质"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "服务对象"
         Object.Tag             =   "服务对象"
         Text            =   "服务对象"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "说明"
         Object.Tag             =   "说明"
         Text            =   "说明"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3900
      Top             =   1290
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
            Picture         =   "frmDeptManage.frx":0BE5
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1231
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":154D
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4050
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":186D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1A8D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1CAD
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1ECD
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":20ED
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":230D
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":252D
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":274D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2969
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2B89
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3343
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3990
      Top             =   2160
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
            Picture         =   "frmDeptManage.frx":355D
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3BA9
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3EC5
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   4500
      TabIndex        =   4
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   4770
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":41E5
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4405
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4625
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4845
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4A65
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4C85
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4EA5
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":50C5
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":52E1
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":5501
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":5721
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":593B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10095
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
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Object.ToolTipText     =   "停用"
               Object.Tag             =   "停用"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "支持关键字模糊查找"
               Object.ToolTipText     =   "查找部门,支持关键字模糊查找"
               Object.Tag             =   "查找"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "人员查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   7560
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   15
            Top             =   240
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   74
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   8520
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "简码"
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label lblFind 
            Caption         =   "查找"
            Height          =   255
            Left            =   8400
            TabIndex        =   14
            Top             =   2520
            Width           =   615
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcDept 
      Height          =   870
      Left            =   6555
      TabIndex        =   17
      Top             =   5010
      Width           =   1230
      _Version        =   589884
      _ExtentX        =   2170
      _ExtentY        =   1535
      _StockProps     =   64
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "参数设置(&R)"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&T)"
      End
      Begin VB.Menu mnuEditRecovery 
         Caption         =   "恢复(&R)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExtend 
         Caption         =   "扩展信息维护(&E)"
      End
      Begin VB.Menu mnuEditExpand 
         Caption         =   "加长下级编码(&X)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuviewsplit1 
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "查找(&F)"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "显示所有下级(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "显示停用部门(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowDel 
         Caption         =   "显示已删除部门(&Y)"
      End
      Begin VB.Menu mnuViewReflash 
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
         Caption         =   "Web上的中联"
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
   Begin VB.Menu mnuShort1 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "增加(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "恢复(&R)"
         Index           =   4
      End
   End
   Begin VB.Menu mnuShort2 
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
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "恢复(&R)"
         Index           =   4
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
End
Attribute VB_Name = "frmDeptManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngStartX As Single, sngStartY As Single    '移动前鼠标的位置
Dim mblnLoad As Boolean  '窗口还未打开时为真
Dim mblnItem As Boolean  '为真表示单击到ListView某一项上
Dim mintColumn As Integer '
Dim mstrKey As String
Private Const mstrLvw As String = "名称,2000,0,1;编码,800,0,2;简码,1440,0,0;位置,2000,0,0;建档时间,1440,0,0;撤档时间,1440,0,0;上级部门,2000,0,0"
Dim mbln药店  As Boolean
Private mlngMode As Long
Private mstrPrivs As String                              '权限串
Private mint假删除 As Integer        '当有"已删除部门"分类时，删除操作都为假删除：0-真删除;1-假删除
Private mint性质 As Integer         '1-临床科室;2-病区;3-临床且病区
Private mlng配置中心 As Long
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String     '记录查询文本框的值
Private mint焦点 As Integer         '记录焦点位置 1在listview控件上 2在树表上
Private Const mint按层次 As Integer = 0   '页面显示方式 按层次关系显示
Private Const mint按性质 As Integer = 1   '页面显示方式   按性质关系显示
Private Const mCON部门性质 As Integer = 0
Private Const mCON扩展信息 As Integer = 1
Private mobjForm As frmDeptExtend
Private mblnPACSInterface As Boolean        '启用影像信息系统接口

Private Function CheckExistDepPres(ByVal lngDepID As Long) As Boolean
    '检查该部门下是否存在人员
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 人员id From 部门人员 " & _
        " Where 部门id In (Select ID From 部门表 " & _
        " Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null " & _
        " Start With ID = [1] Connect By Prior ID = 上级id) And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查部门人员", lngDepID)
    
    If rsTemp.RecordCount > 0 Then
        CheckExistDepPres = True
        Exit Function
    End If
End Function

Private Sub InitTabControl()
    Dim i As Integer
    '初始化Tabcontrol控件
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        .InsertItem(mint按层次, "按层次显示", picList.hwnd, 0).Tag = "按层次显示"
        .InsertItem(mint按性质, "按性质显示", picList.hwnd, 0).Tag = "按性质显示"
        
        .Item(mint按性质).Selected = True
        .Item(mint按层次).Selected = True
    End With
    
    With Me.tbcDept
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        
        Set mobjForm = New frmDeptExtend
        Call SetFormVisible(mobjForm.hwnd) '将窗体最大最小化隐藏

        .InsertItem(mCON部门性质, "部门性质", lvw部门性质_S.hwnd, 0).Tag = "部门性质"
        .InsertItem(mCON扩展信息, "扩展信息", mobjForm.hwnd, 0).Tag = "扩展信息"
        
        .Item(mCON扩展信息).Selected = True
        .Item(mCON部门性质).Selected = True
    End With
End Sub

Private Sub CheckHaveDelDept()
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select ID From 部门表 Where 编码 = '-'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否有已删除部门分类")
    
    If Not rsTemp.EOF Then
        mint假删除 = 1
    Else
        mint假删除 = 0
    End If
End Sub

Private Function Check配置中心(ByVal lngDeptID As Long) As Boolean
    '检查配置中心
    Dim rsData As ADODB.Recordset
    
    If mlng配置中心 = 0 Then
        Check配置中心 = True
        Exit Function
    End If
    
    '如果当前部门是已启用的输液配置中心
    If mlng配置中心 = lngDeptID Then
        MsgBox "该部门已被启用为医院的输液配置中心，不能删除或停用，请在基础参数设置中处理。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '如果当前部门的下级是已启用的输液配置中心
    gstrSQL = "Select Id,编码 || '-' || 名称 As 名称 From 部门表 " & _
        " Where ID In (Select ID From 部门表 Start With 上级id = [1] Connect By Prior ID = 上级id) And ID = [2] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "检查部门", lngDeptID, mlng配置中心)
    
    If rsData.RecordCount > 0 Then
        MsgBox "该部门的下级部门(" & rsData!名称 & ")已被启用为医院的输液配置中心，不能删除或停用，请在基础参数设置中处理。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Check配置中心 = True
End Function

Private Sub Show病区科室对应(ByVal str部门ID As String)
    '取病区科室对应关系
    Dim strCon As String
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    Dim sngTop, sngBottom As Single
    Dim nod As Node
        
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvwDept.Visible = False
    tvwMain_S.Top = 0
    tvwMain_S.Height = IIF(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top - picSplit2.Height, 0)
    tvwMain_S.Left = 0
    
    If str部门ID = "" Or str部门ID = "Root" Then Exit Sub
     
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "部门性质", Val(Mid(str部门ID, 2)))
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    mint性质 = 0
    Do While Not rsTmp.EOF
        If rsTmp("工作性质") = "临床" Then
            If mint性质 = 2 Then
                mint性质 = 3
            Else
                mint性质 = 1
            End If
        End If
        If rsTmp("工作性质") = "护理" Then
            If mint性质 = 1 Then
                mint性质 = 3
            Else
                mint性质 = 2
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If mint性质 = 0 Then Exit Sub
    
    If mint性质 = 1 Then
        strCon = " Select Distinct 病区id From 病区科室对应 Where 科室id = [1] "
    ElseIf mint性质 = 2 Then
        strCon = " Select Distinct 科室id From 病区科室对应 Where 病区id = [1] "
    ElseIf mint性质 = 3 Then
        strCon = " Select 病区id As ID From 病区科室对应 Where 科室id = [1] " & _
                " Union " & _
                " Select 科室id As ID From 病区科室对应 Where 病区id = [1] "
    End If

    If mnuViewShowStop.Checked = False Then
        strTemp = " And (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null ) "
    End If

    gstrSQL = "Select Id, 名称, 编码, 撤档时间 From 部门表 Where ID In (" & strCon & ") " & strTemp & " Order by 编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病区科室对应", Val(Mid(str部门ID, 2)))
    
    If rsTmp.RecordCount = 0 Then
        tvwDept.Visible = False
        tvwMain_S.Top = 0
        tvwMain_S.Height = IIF(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top - picSplit2.Height, 0)
        tvwMain_S.Left = 0
        Exit Sub
    End If
    
    tvwDept.Visible = True
    tvwDept.Height = 2000
    
    tvwMain_S.Height = tvwMain_S.Height - tvwDept.Height - picSplit2.Height - 50
    tvwDept.Left = tvwMain_S.Left
    tvwDept.Width = picSplit2.Width - 50
    tvwDept.Top = tvwMain_S.Top + tvwMain_S.Height + 50 + picSplit2.Height

    tvwDept.Nodes.Clear
    tvwDept.Nodes.Add , , "Root", tvwMain_S.SelectedItem.Text, "Root", "Root"
    tvwDept.Nodes("Root").Expanded = True
    Do Until rsTmp.EOF
        If CDate(IIF(IsNull(rsTmp("撤档时间")), CDate("3000/1/1"), rsTmp("撤档时间"))) = CDate("3000/1/1") Then
            strTemp = "Dept"
        Else
            strTemp = "Dept_No"
        End If

        tvwDept.Nodes.Add "Root", tvwChild, "C" & rsTmp("id"), "【" & rsTmp("编码") & "】" & rsTmp("名称"), strTemp, strTemp
        
        rsTmp.MoveNext
    Loop
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        FillTree
    End If
    mblnLoad = False
End Sub
Private Sub Form_Load()
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    SetParent txtFind.hwnd, Toolbar1.hwnd
    SetParent picFind.hwnd, Toolbar1.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    Call InitTabControl
    
    Call 权限控制
    Call CheckHaveDelDept
    '允许进行列删除的ListView须做标记
    lvwMain.Tag = "可变化的"
    '-----------
    RestoreWinState Me, App.ProductName
    lvw部门性质_S.Visible = True
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", 0)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", 0)) = 1)
    mnuViewShowDel.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示删除", 0)) = 1)
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
    End If
    '根据LvwMain显示设置对应菜单
     mnuViewIcon_Click lvwMain.View
     lvw部门性质_S.View = lvwReport
     lbl病区科室.BackStyle = 0
     
     mlng配置中心 = Val(zlDatabase.GetPara("配置中心", glngSys, 0))
    Call InitSystemPara
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    '初始化新网RIS接口
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    With Me.tbcDetails
        .Move 0, sngTop, Me.Width / 4, sngBottom - sngTop
        picList.Height = .Height - 400
    End With
    With tvwMain_S
        tvwMain_S.Top = 0
        tvwMain_S.Height = picList.Height
        tvwMain_S.Width = tbcDetails.Width - 60
        tvwMain_S.Left = 0
    End With
    
    If glngSys = 100 Then
        If tvwDept.Visible = True Then
            tvwDept.Height = 2000
            tvwMain_S.Height = tvwMain_S.Height - tvwDept.Height - 50 - picSplit2.Height
            tvwDept.Left = tvwMain_S.Left
            tvwDept.Width = picSplit2.Width - 50
            tvwDept.Top = tvwMain_S.Top + tvwMain_S.Height + 50 + picSplit2.Height
        End If
    End If
    
    picSplit.Top = sngTop
    picSplit.Height = IIF(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tbcDetails.Left + tbcDetails.Width
    
    tbcDept.Height = sngBottom / 3
    tbcDept.Left = picSplit.Left + picSplit.Width
    tbcDept.Top = sngBottom - tbcDept.Height
    
    If tbcDept.Top < tvwMain_S.Top + 2000 Then tbcDept.Top = tvwMain_S.Top + 2000
    
    picSplitH.Left = tbcDept.Left
    picSplitH.Top = tbcDept.Top - picSplitH.Height
    
    picSplit2.Left = tvwMain_S.Left
    picSplit2.Top = tvwMain_S.Top + tvwMain_S.Height
    picSplit2.Width = picList.Width
    lbl病区科室.Width = picList.Width
    tvwDept.Width = picList.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = picSplitH.Top - lvwMain.Top
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    tbcDept.Width = lvwMain.Width
    picSplitH.Width = lvwMain.Width
    
    lvw部门性质_S.Move 0, 400, tbcDept.Width, tbcDept.Height - 400
    lvw部门性质_S.ColumnHeaders.Item(3).Width = lvw部门性质_S.Width - lvw部门性质_S.ColumnHeaders.Item(1).Width - lvw部门性质_S.ColumnHeaders.Item(2).Width
    lvwMain.ColumnHeaders.Item(2).Width = 1000
    picSplit2.Visible = tvwDept.Visible
    Me.Refresh
    lvwMain.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    If Not mobjForm Is Nothing Then Set mobjForm = Nothing
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", IIF(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", IIF(mnuViewShowStop.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示删除", IIF(mnuViewShowDel.Checked, 1, 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lbl病区科室_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And picSplit2.Top + Y > tvwMain_S.Top + 200 And picSplit2.Top + Y < stbThis.Top - 1500 Then
        picSplit2.Top = picSplit2.Top + Y
        tvwMain_S.Height = tvwMain_S.Height + Y
        tvwDept.Move 0, tvwDept.Top + Y, picSplit2.Width - 50, tvwDept.Height - Y
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    mint焦点 = 1
    With lvwMain
        If .ListItems.Count = 0 Or .SelectedItem Is Nothing Then
            lvw部门性质_S.ListItems.Clear
            Call mobjForm.initVSf(0)
            Call SetMenu
        Else
            Call SetMenu
        End If
        stbThis.Panels(2).Text = "部门列表中共显示有" & .ListItems.Count & "个部门。"
    End With
End Sub

Public Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowAttribe Mid(Item.Key, 2)
    Call mobjForm.initVSf(Val(Mid(Item.Key, 2)))
    
    Item.Tag = GetClerk(Item.Key)
    
    mblnItem = True
    Call SetMenu
    stbThis.Panels(2).Text = "部门列表中共显示有" & lvwMain.ListItems.Count & "个部门" & IIF(Item.Tag = "0", "。", "，该部门有人员" & Item.Tag & "名。")
End Sub

Private Sub ShowAttribe(ByVal strKey As String)
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim str服务对象 As String
    
    gstrSQL = "select A.工作性质,A.服务对象,B.说明 from 部门性质说明 A,部门性质分类 B where A.工作性质=B.名称 and A.部门ID= [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strKey))
        
    lvw部门性质_S.ListItems.Clear
    Do Until rsTemp.EOF
        Select Case rsTemp("服务对象")
             Case 1
                str服务对象 = "门诊病人"
             Case 2
                str服务对象 = "住院病人"
             Case 3
                str服务对象 = "门诊和住院病人"
             Case Else
                str服务对象 = "不服务于病人"
        End Select
        Set lst = lvw部门性质_S.ListItems.Add(, rsTemp("工作性质"), rsTemp("工作性质"))
        If mbln药店 = True Then
            lst.SubItems(1) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        Else
            lst.SubItems(1) = str服务对象
            lst.SubItems(2) = IIF(IsNull(rsTemp("说明")), "", rsTemp("说明"))
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub
 
Private Sub lvwMain_LostFocus()
    mint焦点 = 0
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        If tbcDetails.Selected.Index = mint按层次 Then
            mnuShortMenu2(4).Visible = mnuViewShowDel.Checked And InStr(mstrPrivs, ";增删改;") > 0
        
            '已删除部门分类不允许操作
            If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 Then
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = lvwMain.ListItems.Count > 0
            Else
                mnuShortMenu2(1).Enabled = mnuEditNew.Enabled
                mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
                mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
                mnuShortMenu2(4).Enabled = False
            End If
        Else    '按性质
            If lvwMain.ListItems.Count = 0 Then
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = False
            End If
            If lvwMain.ListItems.Count <> 0 Then
                If lvwMain.SelectedItem.Icon = "Dept_No" And InStr(1, lvwMain.SelectedItem.ListSubItems(1).Text, "-") = 0 Then '停用
                    mnuShortMenu2(1).Enabled = False
                    mnuShortMenu2(2).Enabled = False
                    mnuShortMenu2(3).Enabled = False
                    mnuShortMenu2(4).Enabled = False
                End If
                If lvwMain.SelectedItem.Icon = "Dept_No" And InStr(1, lvwMain.SelectedItem.ListSubItems(1).Text, "-") > 0 Then '删除
                    mnuShortMenu2(1).Enabled = False
                    mnuShortMenu2(2).Enabled = False
                    mnuShortMenu2(3).Enabled = False
                    mnuShortMenu2(4).Enabled = True
                End If
                If lvwMain.SelectedItem.Icon = "Dept" Then  '正常
                    mnuShortMenu2(1).Enabled = True
                    mnuShortMenu2(2).Enabled = True
                    mnuShortMenu2(3).Enabled = True
                    mnuShortMenu2(4).Enabled = False
                End If
            Else
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = False
            End If
        End If
                        
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Long
    Dim strTemp As String
        
    If ActiveControl Is tvwMain_S Then
        If tbcDetails.Selected.Index = mint按层次 Then
            strTemp = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
        Else
            strTemp = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        End If
    
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "该部门或下级部门还有未删除的部门人员，不能删除该部门。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确认要删除名称为“" & tvwMain_S.SelectedItem.Text & "”的部门吗？" & vbCrLf & "如果有下级部门，也会一起被删除。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            If Check配置中心(strTemp) = False Then Exit Sub
            If Int(glngSys / 100) = 1 And mblnPACSInterface Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(10, RISBaseItemOper.Delete, strTemp) <> 1 Then
                        MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISBasicDictTable)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISBasicDictTable)接口，请与系统管理员联系。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            MousePointer = 11
            gstrSQL = "zl_部门表_DELETE(" & strTemp & "," & mint假删除 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            MousePointer = 0
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            Else
                tvwMain_S.SelectedItem.Parent.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            End If
            tvwMain_S.Nodes.Remove strKey
            '除非只有根结点一个，否则都可修改
            Call SetMenu
        End If
    Else
        If tbcDetails.Selected.Index = mint按层次 Then
            strTemp = Val(Mid(lvwMain.SelectedItem.Key, 2))
        Else
            strTemp = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "该部门或下级部门还有未删除的部门人员，不能删除该部门。", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("你确认要删除名称为“" & lvwMain.SelectedItem.Text & "”的部门吗？" & vbCrLf & "如果有下级部门，也会一齐被删除。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            If Check配置中心(strTemp) = False Then Exit Sub
            If Int(glngSys / 100) = 1 And mblnPACSInterface Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(10, RISBaseItemOper.Delete, strTemp) <> 1 Then
                        MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISBasicDictTable)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISBasicDictTable)接口，请与系统管理员联系。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            Me.MousePointer = 11
            gstrSQL = "zl_部门表_DELETE(" & strTemp & "," & mint假删除 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.MousePointer = 0
            With lvwMain
                '先删除TreeView中对应节点
'                tvwMain_S.Nodes.Remove .SelectedItem.Key
                '再删除ListView中对应节点
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                    lvwMain_ItemClick .SelectedItem
                Else
                    Call lvwMain_GotFocus
                End If
            End With
        End If
    End If
    FillTree
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = 0
End Sub

Private Sub mnuEditExtend_Click()
    Dim strKey As String
    Dim strName As String
    
    On Error Resume Next
    
    If ActiveControl Is lvwMain Then
        If tbcDetails.Selected.Index = mint按性质 Then
            strKey = Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1)
        Else
            strKey = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        strName = lvwMain.SelectedItem.Text
    Else
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint按性质 Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
            Else
                strKey = Mid(.Key, 2)
            End If
            strName = Mid(.Text, InStr(1, .Text, "】") + 1)
        End With
    End If
    
    Call frmDeptExtend.ShowMe(Me, strKey, strName, 0, 1)
    Call mobjForm.initVSf(Val(strKey), 0)
End Sub

Private Sub mnuEditModify_Click()
    Dim str编码 As String
    Dim str名称 As String
    Dim i As Integer
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim str上级编码 As String
    Dim strTemp As String
    
    On Error Resume Next
    If ActiveControl Is lvwMain Then
'        If tvwMain_S.SelectedItem.Key = "Root" Then
'            Exit Sub
'        End If
        
        If mnuViewShowAll.Checked = True Then
            '有上级部门列
            If tbcDetails.Selected.Index = mint按性质 Then
                strKey = Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1)
            Else
                strKey = Mid(lvwMain.SelectedItem.Key, 2)
            End If
            gstrSQL = "select a.上级id,b.编码,b.名称  from 部门表 a,部门表 b where a.上级id=b.id(+)  and a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询上级id", strKey)
            If Not rsTemp Is Nothing Then
                strTemp = rsTemp!上级id
            Else
                Exit Sub
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str上级编码 = ""
            Else
                str上级编码 = tvwMain_S.SelectedItem.Key
            End If
            
            frmDeptSet.编辑部门 mstrPrivs, strKey, 2, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), str上级编码, strTemp
        Else
            If tvwMain_S.SelectedItem.Key = "Root" Then
                Call frmDeptSet.编辑部门(mstrPrivs, Mid(lvwMain.SelectedItem.Key, 2), 2, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), "")
            Else
                frmDeptSet.编辑部门 mstrPrivs, Mid(lvwMain.SelectedItem.Key, 2), 2, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), tvwMain_S.SelectedItem.Parent.Key
            End If
        End If
    Else
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint按性质 Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
                gstrSQL = "select 上级id from 部门表  where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询上级id", strKey)
                If Not rsTemp Is Nothing Then
                    strTemp = rsTemp!上级id
                Else
                    Exit Sub
                End If
            Else
                strKey = Mid(.Key, 2)
                strTemp = Mid(.Parent.Key, 2)
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str上级编码 = ""
            Else
                str上级编码 = Mid(.Key, 1, InStr(1, .Key, "|") - 1)
            End If
            
            Call frmDeptSet.编辑部门(mstrPrivs, strKey, 2, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), str上级编码, strTemp)
        End With
    End If
End Sub

Private Sub mnuEditNew_Click()
    Dim str编码 As String
    Dim str名称 As String
    Dim i As Integer
    Dim strKey As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim str上级编码 As String
        
    If tbcDetails.Selected.Index = mint按性质 Then
        If ActiveControl Is tvwMain_S Then
            strKey = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        Else
            strKey = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        
        gstrSQL = "Select c.编码, a.Id, a.上级id" & _
                   " From 部门表 A, 部门性质说明 B, 部门性质分类 C" & _
                   " Where b.工作性质 = c.名称 And a.Id = b.部门id and a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询上级id", strKey)
        If Not rsTemp Is Nothing Then
            strTemp = rsTemp!上级id
            str上级编码 = rsTemp!编码 & "|" & rsTemp!ID
        End If
    Else
        strTemp = Mid(tvwMain_S.SelectedItem.Key, 2)
    End If

    Call frmDeptSet.编辑部门(mstrPrivs, "", 1, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), str上级编码, strTemp)
    Call FillTree
End Sub

Private Sub mnuEditRecovery_Click()
    Dim strKey As String
    Dim str上级编码 As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    '恢复已删除部门
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If tbcDetails.Selected.Index = mint按性质 Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
                gstrSQL = "select 上级id from 部门表  where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询上级id", strKey)
                If Not rsTemp Is Nothing Then
                    strTemp = rsTemp!上级id
                Else
                    Exit Sub
                End If
            Else
                strKey = Mid(.Key, 2)
                strTemp = Mid(.Parent.Key, 2)
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str上级编码 = ""
            Else
                If tbcDetails.Selected.Index = mint按性质 Then
                    str上级编码 = Mid(.Key, 1, InStr(1, .Key, "|") - 1)
                Else
                    str上级编码 = ""
                End If
            End If
        End With
        Call frmDeptSet.编辑部门(mstrPrivs, strKey, 1, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), str上级编码, strTemp)
    ElseIf ActiveControl Is lvwMain Then
        If tbcDetails.Selected.Index = mint按性质 Then
            strKey = Val(Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1))
        Else
            strKey = Val(Mid(lvwMain.SelectedItem.Key, 2))
        End If
        gstrSQL = "select a.上级id,b.编码,b.名称  from 部门表 a,部门表 b where a.上级id=b.id(+)  and a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询上级id", strKey)
        If Not rsTemp Is Nothing Then
            strTemp = rsTemp!上级id
        Else
            Exit Sub
        End If
        
        If tvwMain_S.SelectedItem.Key = "Root" Then
            str上级编码 = ""
        Else
            str上级编码 = tvwMain_S.SelectedItem.Key
        End If
        
        Call frmDeptSet.编辑部门(mstrPrivs, strKey, 1, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2), str上级编码, strTemp)
    End If
End Sub

Private Sub mnuEditStart_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim j As Integer
    Dim strTemp As String
    Dim str性质 As String
            
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint按层次 Then
                strKey = .Key
            Else
                strKey = "C" & Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
            End If
            
            gstrSQL = "zl_部门表_reuse(" & Mid(strKey, 2) & ")"
        End With
    Else
        If tbcDetails.Selected.Index = mint按层次 Then
            strKey = lvwMain.SelectedItem.Key
        Else
            strKey = "C" & Mid(lvwMain.SelectedItem.Key, 2)
        End If
        gstrSQL = "zl_部门表_reuse(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    
    '执行启用过程
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '改变图标和颜色
    If ActiveControl Is tvwMain_S Then
        tvwMain_S.SelectedItem.Image = "Dept"
        tvwMain_S.SelectedItem.SelectedImage = "Dept"
    Else
        If tbcDetails.Selected.Index = mint按层次 Then
            tvwMain_S.Nodes(strKey).Image = "Dept"
            tvwMain_S.Nodes(strKey).SelectedImage = "Dept"
        Else
            str性质 = "C" & Mid(tvwMain_S.SelectedItem.Key, 2, 1) & "|" & Mid(strKey, 2)
            tvwMain_S.Nodes(str性质).Image = "Dept"
            tvwMain_S.Nodes(str性质).SelectedImage = "Dept"
        End If
        With lvwMain.SelectedItem
            .Icon = "Dept"
            .SmallIcon = "Dept"
            .ForeColor = RGB(0, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain.ColumnHeaders.Count
                If i < lvwMain.ColumnHeaders.Count Then
                    .ListSubItems(i).ForeColor = RGB(0, 0, 0)
                End If
                '更新撤档时间
                If lvwMain.ColumnHeaders(i).Text = "撤档时间" Then
                    .SubItems(i - 1) = "3000-01-01"
                End If
            Next
        End With
    End If
    
    '处理上级目录
    If tbcDetails.Selected.Index = mint按层次 Then  '只有在按层次显示的页面中才处理上级图标
        If ActiveControl Is tvwMain_S Then
            j = Me.tvwMain_S.SelectedItem.Index
        Else
            j = Me.tvwMain_S.Nodes(lvwMain.SelectedItem.Key).Index
        End If
        
        While Me.tvwMain_S.Nodes(j).Parent.Image = "Dept_No"
            With tvwMain_S.Nodes(j)
                strKey = .Parent.Key
                gstrSQL = "zl_部门表_reuse(" & Mid(.Parent.Key, 2) & ")"
                '执行启用过程
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                '处理图标
                .Parent.Image = "Dept"
                .Parent.SelectedImage = "Dept"
                j = .Parent.Index
            End With
        Wend
    End If
    
    '改变状态栏和菜单
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim strTemp As String

    If ActiveControl Is tvwMain_S Then
        
        If tvwMain_S.SelectedItem.Key = "Root" Then Exit Sub
        If CheckStop = False Or tvwMain_S.SelectedItem.Tag <> "0" Then
            MsgBox "一个部门只有没有下级部门或所属人员时才能被停用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If tbcDetails.Selected.Index = mint按层次 Then
            strKey = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
        Else
            strKey = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        End If
        
        '检查业务相关性
        If CheckBusiness(Val(strKey)) = False Then Exit Sub
        
        If Check配置中心(Val(strKey)) = False Then Exit Sub
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            strKey = strKey
            gstrSQL = "zl_部门表_stop(" & strKey & ")"
        End With
    Else
        If CheckStop = False Or lvwMain.SelectedItem.Tag <> "0" Then
            MsgBox "一个部门只有没有下级部门或所属人员时才能被停用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查业务相关性
        If CheckBusiness(Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
        
        If Check配置中心(Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
        strKey = lvwMain.SelectedItem.Key
        gstrSQL = "zl_部门表_stop(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    '执行启用过程
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '改变图标和颜色
    If mnuViewShowStop.Checked = True Then '要显示停用部门
        If ActiveControl Is tvwMain_S Then
            tvwMain_S.SelectedItem.Image = "Dept_No"
            tvwMain_S.SelectedItem.SelectedImage = "Dept_No"
        Else
            If tbcDetails.Selected.Index = mint按层次 Then
                tvwMain_S.Nodes(strKey).Image = "Dept_No"
                tvwMain_S.Nodes(strKey).SelectedImage = "Dept_No"
            Else
                tvwMain_S.Nodes(tvwMain_S.SelectedItem.Key).Image = "Dept_No"
                tvwMain_S.Nodes(tvwMain_S.SelectedItem.Key).SelectedImage = "Dept_No"
            End If
            With lvwMain.SelectedItem
                .Icon = "Dept_No"
                .SmallIcon = "Dept_No"
                .ForeColor = RGB(255, 0, 0)
                
                Dim i As Integer
                For i = 1 To lvwMain.ColumnHeaders.Count
                    If i < lvwMain.ColumnHeaders.Count Then
                        .ListSubItems(i).ForeColor = RGB(255, 0, 0)
                    End If
                    '更新撤档时间
                    If lvwMain.ColumnHeaders(i).Text = "撤档时间" Then
                        .SubItems(i - 1) = Format(Date, "yyyy-MM-dd")
                    End If
                Next
            End With
        End If
        Call SetMenu
    Else '不显示停用部门
        If ActiveControl Is tvwMain_S Then
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            Else
                tvwMain_S.SelectedItem.Parent.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            End If
            tvwMain_S.Nodes.Remove strKey
            '除非只有根结点一个，否则都可修改
            Call SetMenu
        Else
            With lvwMain
                tvwMain_S.Nodes.Remove .SelectedItem.Key
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    .ListItems(1).Selected = True
                    .ListItems(1).EnsureVisible
                    lvwMain_ItemClick .SelectedItem
                Else
                    Call lvwMain_GotFocus
                End If
            End With
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditExpand_Click()
    Dim strTemp As String
    Dim str父编码 As String
    Dim str编码 As String
    Dim intNew As Integer '目前最长的

    On Error GoTo ErrHandle
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
            str父编码 = ""
            intNew = GetDownCodeLength("", "部门表")
        Else
            str父编码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "部门表")
        End If
        If intNew = 10 Then
            MsgBox "不能再加长编码，某一个下级已经用足了长度。", vbExclamation, gstrSysName
            Exit Sub
        End If
        str编码 = Mid(.Child.Text, 2, InStr(.Child.Text, "】") - 2)
        intNew = frmCodingL.GetLength(Len(str编码), 10 - (intNew - Len(str编码)), .Text)
        If intNew = 0 Then Exit Sub
        strTemp = str父编码 & String(intNew - Len(str编码), "0")
        If .Key = "Root" Then
            gstrSQL = "zl_部门表_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & ",0)"
        Else
            gstrSQL = "zl_部门表_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & "," & Mid(.Key, 2) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileParameter_Click()
    frmSetParameter.ShowMe Me
End Sub

Private Sub mnuFind_Click()
    frmPresFind.ShowOfType Me, 1, mnuViewShowStop.Checked, mnuViewShowDel.Checked, IIF(tbcDetails.Selected.Index = mint按层次, 1, 2)
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：部门=部门id，分类=分类id
    Dim lng分类id As Long
    Dim lng部门ID As Long
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        If tvwMain_S.SelectedItem.Key <> "Root" Then
            lng分类id = Mid(tvwMain_S.SelectedItem.Key, 2)
        End If
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng部门ID = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "分类=" & IIF(lng分类id = 0, "", lng分类id), _
        "部门=" & IIF(lng部门ID = 0, "", lng部门ID))
End Sub

'Private Sub mnuViewFind_Click()
'    frmDeptCharacter.显示部门
'End Sub

Private Sub mnuViewReflash_Click()
    FillTree
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '列有变化就要重新刷新
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewShowAll_Click()
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowDel_Click()
    mnuViewShowDel.Checked = Not mnuViewShowDel.Checked
    FillTree
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillTree
End Sub

'Private Sub opt层次_Click()
'    Call FillTree
'End Sub

'Private Sub opt性质_Click()
'    Call FillTree
'End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - sngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 500 Then
            picSplit.Left = sngTemp
            tbcDetails.Width = picSplit.Left - tvwMain_S.Left
            tvwMain_S.Width = tbcDetails.Width
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
            picSplit2.Width = tvwMain_S.Width
            
            If glngSys = 100 Then
                If tvwDept.Visible = True Then
                    tvwDept.Width = tvwMain_S.Width
                End If
            End If
            
            picSplitH.Left = lvwMain.Left
            tbcDept.Left = lvwMain.Left
            picSplitH.Width = lvwMain.Width
            tbcDept.Width = lvwMain.Width
            lvw部门性质_S.Left = 0
            lvw部门性质_S.Width = tbcDept.Width
        End If
        tvwMain_S.SetFocus
    End If
End Sub
'
Private Sub picSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And picSplit2.Top + Y > tvwMain_S.Top + 200 And picSplit2.Top + Y < stbThis.Top - 1500 Then
        picSplit2.Top = picSplit2.Top + Y
        tvwMain_S.Height = tvwMain_S.Height + Y
        tvwDept.Move 0, tvwDept.Top + Y, picSplit2.Width - 50, tvwDept.Height - Y
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - lvwMain.Top > 2500 And IIF(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - (sngTemp + picSplitH.Height) > 1200 Then
            picSplitH.Top = sngTemp
            lvwMain.Height = picSplitH.Top - tvwMain_S.Top - 800
            tbcDept.Top = picSplitH.Top + picSplitH.Height
            tbcDept.Height = IIF(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - tbcDept.Top
            lvw部门性质_S.Top = 400
            lvw部门性质_S.Height = tbcDept.Height - 400
        End If
        lvwMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub


Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Integer
    
    Call FillTree
    If InStr(tvwMain_S.SelectedItem.Text, "所有性质") > 0 And tbcDetails.Selected.Index = mint按性质 Then
        mnuEdit.Enabled = False
        mnuEditNew.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = False
        mnuEditRecovery.Enabled = False
        Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
        Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
        Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
        Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
        Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        Exit Sub
    End If
    If mblnLoad = False Then
        lvwMain.SetFocus
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "Find"
            mnuFind_Click
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "√", "  ")
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text, "  ", "√")
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
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

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "√", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "√")
    lvwMain.View = Index
End Sub


Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
        Case 4
            mnuEditRecovery_Click
    End Select

End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
        Case 4
            mnuEditRecovery_Click
    End Select

End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuhelptopic_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwDept_NodeClick(ByVal Node As MSComctlLib.Node)
    '''''
End Sub


Private Sub tvwMain_S_GotFocus()
    mint焦点 = 2
    If tbcDetails.Selected.Index = mint按层次 Then
        stbThis.Panels(2).Text = "本部门有" & tvwMain_S.SelectedItem.Children & "个直属部门" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "。", "，人员" & tvwMain_S.SelectedItem.Tag & "名。")
    Else
        If tvwMain_S.SelectedItem.Text = "所有性质" Then
            stbThis.Panels(2).Text = "本性质有" & tvwMain_S.SelectedItem.Children & "个直属性质"
        ElseIf InStr(1, tvwMain_S.SelectedItem.Text, "【") = 0 Then
            stbThis.Panels(2).Text = "本性质有" & tvwMain_S.SelectedItem.Children & "个直属部门" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "。", "，人员" & tvwMain_S.SelectedItem.Tag & "名。")
        ElseIf InStr(1, tvwMain_S.SelectedItem.Text, "【") > 0 Then
            stbThis.Panels(2).Text = "本部门有" & tvwMain_S.SelectedItem.Children & "个直属部门" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "。", "，人员" & tvwMain_S.SelectedItem.Tag & "名。")
        End If
    End If
    Call SetMenu
End Sub

Private Sub tvwMain_S_LostFocus()
    mint焦点 = 0
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If tbcDetails.Selected.Index = mint按层次 Then
            If mnuShortMenu1(1).Visible = False Then Exit Sub
            
            mnuShortMenu1(4).Visible = mnuViewShowDel.Checked
            
            '已删除部门分类不允许操作
            If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 Then
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = Trim(tvwMain_S.SelectedItem.Text) <> "【-】已删除部门"
            Else
                mnuShortMenu1(1).Enabled = mnuEditNew.Enabled
                mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
                mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
                mnuShortMenu1(4).Enabled = False
            End If
        Else    '按性质
            If InStr(1, tvwMain_S.SelectedItem.Text, "【") = 0 Then '性质
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = False
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "【-") Then    '删除
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = True
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "【") > 0 And InStr(1, tvwMain_S.SelectedItem.Text, "【-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then '停用
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = False
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "【") > 0 And InStr(1, tvwMain_S.SelectedItem.Text, "【-") = 0 And tvwMain_S.SelectedItem.Image = "Dept" Then  '正常
                mnuShortMenu1(1).Enabled = True
                mnuShortMenu1(2).Enabled = True
                mnuShortMenu1(3).Enabled = True
                mnuShortMenu1(4).Enabled = False
            End If
        End If
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strTemp As String
        
    If tbcDetails.Selected.Index = mint按性质 Then
        If InStr(1, tvwMain_S.SelectedItem.Text, "所有性质") > 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        
        If InStr(tvwMain_S.SelectedItem.Text, "【") = 0 Then    '顶级
            mnuEdit.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then '删除
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = True
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then '停用
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = True
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then '正常
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditDelete.Enabled = True
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = True
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        mnuEditExpand.Enabled = False
        strTemp = Mid(Node.Key, InStr(1, Node.Key, "|") + 1, Len(Node.Key) - InStr(1, Node.Key, "|"))
        strTemp = "C" & strTemp
    Else
        strTemp = Node.Key
        mnuEdit.Enabled = True
    End If
    
'    If mstrKey = strTemp Then Exit Sub
    
    mstrKey = strTemp
    
    Node.Tag = GetClerk(strTemp)
    
    FillList strTemp
    
    If InStr(tvwMain_S.SelectedItem.Text, "【") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then
        If tbcDetails.Selected.Index = mint按性质 Then
            Call ShowAttribe(Mid(Node.Key, 4))
            Call mobjForm.initVSf(Val(Mid(Node.Key, 4)))
        Else
            Call ShowAttribe(Mid(Node.Key, 2))
            Call mobjForm.initVSf(Val(Mid(Node.Key, 2)))
        End If
    End If
    
    If glngSys = 100 Then
        Show病区科室对应 (strTemp)
    End If
    picSplit2.Visible = tvwDept.Visible
    
    tvwMain_S_GotFocus
    
    If picSplit2.Visible = True Then
        picSplit2.Left = tvwMain_S.Left
        picSplit2.Top = tvwMain_S.Top + tvwMain_S.Height
        picSplit2.Width = tvwMain_S.Width
    End If
    If tvwDept.Visible = False Then
        tvwMain_S.Height = picList.Height
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "部门表"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub FillTree()
'功能:装入所有部门到tvwMain_S
'参数:
    Dim strTemp As String
    Dim strKey As String
    Dim rs部门 As New ADODB.Recordset
    Dim rs性质部门 As ADODB.Recordset
    Dim str性质 As String
    Dim i As Integer
    Dim rs部门性质 As ADODB.Recordset
    Dim str编码 As String
    Dim nod As Node
    Dim str图标 As String
    Dim str删除 As String
    
    mstrKey = ""
    On Error GoTo ErrHandle
    rs部门.CursorLocation = adUseClient
    rs部门.CursorType = adOpenKeyset
    rs部门.LockType = adLockReadOnly
    
    If tbcDetails.ItemCount = 0 Then Exit Sub
    If tbcDetails.Selected.Index = mint按层次 Then        '按层次显示
        If Not tvwMain_S.SelectedItem Is Nothing Then
            strKey = tvwMain_S.SelectedItem.Key
        End If
    
        If mnuViewShowStop.Checked = False Then
            strTemp = " where (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null ) "
        End If
       
        gstrSQL = "select id,上级id,编码 ,名称,to_char(撤档时间,'YYYY-MM-DD') as 撤档时间  from 部门表 " & strTemp & " " & _
                " start with 上级id is null And 编码 <> '-' " & _
                " connect by prior id =上级id"
        Set rs部门 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        tvwMain_S.Nodes.Clear
        tvwMain_S.Nodes.Add , , "Root", "所有部门", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        
        Do Until rs部门.EOF
            If CDate(IIF(IsNull(rs部门("撤档时间")), CDate("3000/1/1"), rs部门("撤档时间"))) = CDate("3000/1/1") Then
                strTemp = "Dept"
            Else
                strTemp = "Dept_No"
            End If
            
            If IsNull(rs部门("上级id")) Then
                tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
            Else
                tvwMain_S.Nodes.Add "C" & rs部门("上级id"), tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
            End If
            tvwMain_S.Nodes("C" & rs部门("id")).Sorted = True
                    
            rs部门.MoveNext
        Loop
        
        '"已删除部门"分类
        If mnuViewShowDel.Checked = True Then
            gstrSQL = " Select ID, 上级id, 编码, 名称, To_Char(撤档时间, 'YYYY-MM-DD') As 撤档时间 " & _
                    " From 部门表 Start With 编码 = '-' and 上级id is null Connect By Prior ID = 上级id"
                    
            Set rs部门 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    '        mint假删除 = 0
            strTemp = "Dept_No"
            
            If Not rs部门.EOF Then
                tvwMain_S.Nodes("Root").Sorted = False
    '            mint假删除 = 1
            End If
            
            Do Until rs部门.EOF
                If IsNull(rs部门("上级id")) Then
                    tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
                Else
                    tvwMain_S.Nodes.Add "C" & rs部门("上级id"), tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
                End If
                tvwMain_S.Nodes("C" & rs部门("id")).Sorted = True
                
                '已删除的部门用红色标记
                tvwMain_S.Nodes("C" & rs部门("id")).ForeColor = &HFF&
                
                rs部门.MoveNext
            Loop
        End If
    Else    '按性质显示
        If Not tvwMain_S.SelectedItem Is Nothing Then
            strKey = tvwMain_S.SelectedItem.Key
        End If
        
        gstrSQL = "select distinct a.编码,a.名称 from 部门性质分类 a,部门性质说明 c where a.名称=c.工作性质"
        Set rs性质部门 = zlDatabase.OpenSQLRecord(gstrSQL, "按性质查询部门")
        
        tvwMain_S.Nodes.Clear
        tvwMain_S.Nodes.Add , , "Root", "所有性质", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        
        str性质 = ""
        Do While Not rs性质部门.EOF
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs性质部门!编码, rs性质部门!名称, "Dept"
            str性质 = str性质 & rs性质部门!名称 & "|"
            rs性质部门.MoveNext
        Loop
        
        If mnuViewShowStop.Checked = False Then
            strTemp = " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
        End If
        
        str删除 = " and a.id not in(select id from 部门表 where 编码 like '-%')"
        
        For i = 0 To UBound(Split(str性质, "|"))
            gstrSQL = "Select a.id,a.上级id,a.名称,a.编码 as 部门编码,c.编码,b.工作性质,a.撤档时间 From 部门表 A, 部门性质说明 B,部门性质分类 c Where b.工作性质=c.名称 " & strTemp _
                & " and A.ID=B.部门ID and B.工作性质=[1]" & str删除
            str编码 = Split(str性质, "|")(i)
            Set rs部门性质 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str编码)
            Do While Not rs部门性质.EOF
                If CDate(IIF(IsNull(rs部门性质("撤档时间")), CDate("3000/1/1"), rs部门性质("撤档时间"))) = CDate("3000/1/1") Then
                    str图标 = "Dept"
                Else
                    str图标 = "Dept_No"
                End If
                If rs部门性质!编码 Like "-*" Then
                    str图标 = "Dept_No"
                End If
            
                tvwMain_S.Nodes.Add "C" & rs部门性质!编码, tvwChild, "C" & rs部门性质!编码 & "|" & rs部门性质!ID, "【" & rs部门性质!部门编码 & "】" & rs部门性质!名称, str图标
                rs部门性质.MoveNext
            Loop
        Next
        
        If mnuViewShowDel.Checked = True Then
            str删除 = "  and a.编码 like '-%'"
            For i = 0 To UBound(Split(str性质, "|"))
                gstrSQL = "Select a.id,a.上级id,a.名称,a.编码 as 部门编码,c.编码,b.工作性质,a.撤档时间 From 部门表 A, 部门性质说明 B,部门性质分类 c Where b.工作性质=c.名称 " & _
                    " and A.ID=B.部门ID and B.工作性质=[1]" & str删除
                str编码 = Split(str性质, "|")(i)
                Set rs部门性质 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str编码)
                Do While Not rs部门性质.EOF
                    If CDate(IIF(IsNull(rs部门性质("撤档时间")), CDate("3000/1/1"), rs部门性质("撤档时间"))) = CDate("3000/1/1") Then
                        str图标 = "Dept"
                    Else
                        str图标 = "Dept_No"
                    End If
                    If rs部门性质!编码 Like "-*" Then
                        str图标 = "Dept_No"
                    End If
                
                    tvwMain_S.Nodes.Add "C" & rs部门性质!编码, tvwChild, "C" & rs部门性质!编码 & "|" & rs部门性质!ID, "【" & rs部门性质!部门编码 & "】" & rs部门性质!名称, str图标
                    rs部门性质.MoveNext
                Loop
            Next
        End If
    End If
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillList(ByVal str部门ID As String)
'功能:装入对应部门的部门到lvwMain
'参数:str部门ID 部门的标识

    Dim rs部门 As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    Dim str停用 As String
    Dim str删除 As String
    
    If tbcDetails.Selected.Index = mint按层次 Then
        If Not lvwMain.SelectedItem Is Nothing Then
            '保留原有键值
            strKey = lvwMain.SelectedItem.Key
        End If
        
        rs部门.CursorLocation = adUseClient
        
        If mnuViewShowStop.Checked = False And InStr(1, tvwMain_S.SelectedItem.Text, "【-") = 0 Then
            str停用 = " (A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01','YYYY-MM-DD')"
            If mnuViewShowDel.Checked = True Then
                str停用 = str停用 & " Or A.编码 Like '-%'"
            End If
            str停用 = str停用 & ")"
        End If
        If mnuViewShowAll.Checked = True Then
            gstrSQL = "select A.*,B.名称 as 上级部门 from " & _
                "(select A.ID,A.上级ID,A.名称,A.编码,A.简码,A.位置,to_char(A.建档时间,'YYYY-MM-DD') as 建档时间,to_char(A.撤档时间,'YYYY-MM-DD') as 撤档时间 " & _
                " from 部门表 A " & IIF(str停用 = "", "", "where " & str停用) & " connect by prior A.id=A.上级id start with " & IIF(mnuViewShowDel.Checked = False, "编码 <> '-' And ", "") & " " & IIF(str部门ID = "Root", "A.上级ID is null ", "A.上级ID = [1]") & ") A,部门表 B where A.上级ID=B.ID(+)"
        Else
            gstrSQL = "select A.ID,A.上级ID,A.名称,A.编码,A.简码,A.位置,to_char(A.建档时间,'YYYY-MM-DD') as 建档时间,to_char(A.撤档时间,'YYYY-MM-DD') as 撤档时间,B.名称 as 上级部门 from 部门表 A,部门表 B where A.上级ID=B.ID(+) and " & IIF(str停用 = "", "", str停用 & " and ") & IIF(str部门ID = "Root", "A.上级ID is null ", "A.上级ID = [1]")
            If mnuViewShowDel.Checked = False Then
                gstrSQL = gstrSQL & " And A.编码 Not Like '-%'"
            End If
        End If
            
        Set rs部门 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(str部门ID, 2)))
    Else
        If mnuViewShowDel.Checked = False Then
            gstrSQL = "Select a.Id, a.上级id, a.名称, a.编码, a.简码,a.位置, To_Char(a.建档时间, 'YYYY-MM-DD') As 建档时间, To_Char(a.撤档时间, 'YYYY-MM-DD') As 撤档时间," & _
                              " c.名称 as 上级部门" & _
                       " From 部门表 A, 部门性质说明 B, 部门表 C" & _
                       " Where a.Id = b.部门id And a.上级id = c.Id(+) And b.工作性质 = [1] and a.id not in(select id from 部门表 where 编码 like '-%')"
        Else
            gstrSQL = "Select a.Id, a.上级id, a.名称, a.编码, a.简码,a.位置, To_Char(a.建档时间, 'YYYY-MM-DD') As 建档时间, To_Char(a.撤档时间, 'YYYY-MM-DD') As 撤档时间," & _
                              " c.名称 as 上级部门" & _
                       " From 部门表 A, 部门性质说明 B, 部门表 C" & _
                       " Where a.Id = b.部门id And a.上级id = c.Id(+) And b.工作性质 = [1] "
        End If
        If mnuViewShowStop.Checked = False Then
            str停用 = " and ((A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01','YYYY-MM-DD'))"
            If mnuViewShowDel.Checked = True Then
                str停用 = str停用 & " or a.编码 like '-%'" & ")"
            Else
                str停用 = str停用 & ")"
            End If
            gstrSQL = gstrSQL & str停用
        End If
        Set rs部门 = zlDatabase.OpenSQLRecord(gstrSQL, "查询下级", tvwMain_S.SelectedItem.Text)
    End If
    
    lvwMain.ListItems.Clear
    lvw部门性质_S.ListItems.Clear
    Call mobjForm.initVSf(0)
    
    Do Until rs部门.EOF
        If CDate(IIF(IsNull(rs部门("撤档时间")), CDate("3000/1/1"), rs部门("撤档时间"))) = CDate("3000/1/1") And rs部门("编码") <> "-" Then
            str停用 = "Dept"
        Else
            str停用 = "Dept_No"
        End If
        Set lst = lvwMain.ListItems.Add(, "C" & rs部门("ID"), rs部门("名称"), str停用, str停用)
        If str停用 = "Dept_No" Then lst.ForeColor = RGB(255, 0, 0)
        
        Dim lngCol  As Long
        Dim varValue As Variant
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rs部门(lvwMain.ColumnHeaders(lngCol).Text).value
            lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            If str停用 = "Dept_No" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rs部门.MoveNext
    Loop
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
        Else
            Item.Selected = True
            Item.EnsureVisible
        End If
        EnablePrint True
    Else
        EnablePrint False
    End If
End Sub

Private Sub SetMenu()
'功能:设置修改和删除按钮的有效值
'参数:blnEnabled 有效值
'功能:设置增加按钮的有效值
    Dim blnEnabled As Boolean
    
    If tvwMain_S.SelectedItem.Image = "Dept_No" Then
        Toolbar1.Buttons("New").Enabled = False
        mnuEditNew.Enabled = False
    Else
        Toolbar1.Buttons("New").Enabled = True
        mnuEditNew.Enabled = True
    End If
    
    mnuEditRecovery.Enabled = False
    mnuEditRecovery.Visible = mnuViewShowDel.Checked
    
    '是否根节点
    If tbcDetails.Selected.Index = mint按层次 And tvwMain_S Is ActiveControl And tvwMain_S.SelectedItem.Key = "Root" Or _
        Not (tvwMain_S Is ActiveControl) And lvwMain.ListItems.Count = 0 Then
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        Toolbar1.Buttons("Start").Enabled = False
        Toolbar1.Buttons("Stop").Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditExtend.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = False
        
        '已删除部门分类不允许操作
        If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 Then
            If mnuViewShowDel.Checked Then
                mnuEditRecovery.Enabled = True
                mnuEditNew.Enabled = False
                mnuEditModify.Enabled = False
                mnuEditExtend.Enabled = False
                mnuEditDelete.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
            Else
                mnuEdit.Enabled = False
            End If
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        mnuEditExpand.Enabled = False
        Exit Sub
    ElseIf tbcDetails.Selected.Index = mint按性质 And tvwMain_S Is ActiveControl And tvwMain_S.SelectedItem.Key = "Root" Or _
        Not (tvwMain_S Is ActiveControl) And lvwMain.ListItems.Count = 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
    End If
    If tvwMain_S.SelectedItem Is Nothing Then
        mnuEditExpand.Enabled = False
    Else
        mnuEditExpand.Enabled = tvwMain_S.SelectedItem.Children <> 0
    End If
    
    If tvwMain_S Is ActiveControl Then
        blnEnabled = (tvwMain_S.SelectedItem.Image = "Dept")
    Else
        blnEnabled = (lvwMain.SelectedItem.Icon = "Dept")
    End If
    
    Toolbar1.Buttons("Modify").Enabled = blnEnabled
    Toolbar1.Buttons("Delete").Enabled = blnEnabled
    mnuEditDelete.Enabled = blnEnabled
    mnuEditModify.Enabled = blnEnabled
    mnuEditExtend.Enabled = blnEnabled
    Toolbar1.Buttons("Start").Enabled = Not blnEnabled
    Toolbar1.Buttons("Stop").Enabled = blnEnabled
    mnuEditStart.Enabled = Not blnEnabled
    mnuEditStop.Enabled = blnEnabled
                    
    '已删除部门分类不允许操作
    If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 Then
        If mnuViewShowDel.Checked Then
            If UCase(Me.ActiveControl.Name) = "TVWMAIN_S" Then
                mnuEditRecovery.Enabled = InStr(tvwMain_S.SelectedItem.Text, "【-】已删除部门") = 0
            Else
                mnuEditRecovery.Enabled = True
            End If
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
        Else
            mnuEdit.Enabled = False
        End If
        Toolbar1.Buttons("New").Enabled = False
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        Toolbar1.Buttons("Start").Enabled = False
        Toolbar1.Buttons("Stop").Enabled = False
    End If
                    
    If tbcDetails.Selected.Index = mint按性质 And ActiveControl Is lvwMain Then
        With lvwMain
            If .ListItems.Count = 0 Then
                mnuEdit.Enabled = False
            Else
                If .SelectedItem.Icon = "Dept_No" And InStr(1, .SelectedItem.ListSubItems(1).Text, "-") > 0 Then '删除
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = False
                    mnuEditModify.Enabled = False
                    mnuEditExtend.Enabled = False
                    mnuEditDelete.Enabled = False
                    mnuEditStart.Enabled = False
                    mnuEditStop.Enabled = False
                    mnuEditRecovery.Enabled = True
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
                If .SelectedItem.Icon = "Dept_No" And InStr(1, .SelectedItem.ListSubItems(1).Text, "-") = 0 Then '停用
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = False
                    mnuEditModify.Enabled = False
                    mnuEditExtend.Enabled = False
                    mnuEditDelete.Enabled = False
                    mnuEditStart.Enabled = True
                    mnuEditStop.Enabled = False
                    mnuEditRecovery.Enabled = False
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
                If .SelectedItem.Icon = "Dept" Then  '正常
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = True
                    mnuEditModify.Enabled = True
                    mnuEditExtend.Enabled = True
                    mnuEditDelete.Enabled = True
                    mnuEditStart.Enabled = False
                    mnuEditStop.Enabled = True
                    mnuEditRecovery.Enabled = False
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
            End If
        End With
    ElseIf tbcDetails.Selected.Index = mint按性质 And ActiveControl Is tvwMain_S Then
        
        If InStr(tvwMain_S.SelectedItem.Text, "【") = 0 Then    '顶级
            mnuEdit.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【-") > 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then '删除
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = True
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then '停用
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = True
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "【") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then '正常
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditExtend.Enabled = True
            mnuEditDelete.Enabled = True
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = True
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(1, tvwMain_S.SelectedItem.Text, "所有性质") > 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        mnuEditExpand.Enabled = False
    End If
    EnablePrint lvwMain.ListItems.Count > 0
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'功能:设置打印和预鉴按钮的有效值
'参数:blnEnabled 有效值
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见

    If InStr(mstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu1(1).Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortMenu2(4).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
    End If
    
    If InStr(mstrPrivs, ";扩展信息维护;") = 0 Then
        mnuEditExtend.Visible = False
    End If
    
    mbln药店 = (glngSys \ 100 = 8)
    If mbln药店 = True Then
        '不显示服务对象
        lvw部门性质_S.ColumnHeaders.Remove 2
    End If
End Sub

Private Function GetClerk(ByVal strKey As String) As Long
    On Error GoTo errClerk
    Dim rsTemp As New ADODB.Recordset
    If strKey = "Root" Then Exit Function
    
    gstrSQL = "Select Count(b.Id) As 人员数" & vbNewLine & _
            "From 部门人员 A, 人员表 B," & vbNewLine & _
            "     (Select ID" & vbNewLine & _
            "       From 部门表" & vbNewLine & _
            "       Where ID = [1] Or ID In (Select ID From 部门表 Start With 上级id = [1] Connect By Prior ID = 上级id)) C" & vbNewLine & _
            "Where a.部门id = c.Id And a.人员id = b.Id And (b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') or 撤档时间 is null)"


    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strKey, 2)))
        
    GetClerk = rsTemp("人员数")
    Exit Function
errClerk:
    GetClerk = 0
End Function

Private Function CheckStop() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '修改人         曾超
    '修改时间       2004-10-18
    '功能           检查下级部门是否全部为停用部门
    '返回           =True全部为停用部门=False全部不为停用部门
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    
    CheckStop = True
    
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If .Children > 0 Then
                i = .Child.FirstSibling.Index
                If .Child.FirstSibling.Image = "Dept" Then
                    CheckStop = False
                    Exit Function
                End If
                While i <> .Child.LastSibling.Index
                    If Me.tvwMain_S.Nodes(i).Next.Image = "Dept" Then
                        CheckStop = False
                        Exit Function
                    End If
                    i = Me.tvwMain_S.Nodes(i).Next.Index
                Wend
            End If
        End With
    Else
        With tvwMain_S.Nodes(lvwMain.SelectedItem.Key)
            If .Children > 0 Then
                i = .Child.FirstSibling.Index
                If .Child.FirstSibling.Image = "Dept" Then
                    CheckStop = False
                    Exit Function
                End If
                While i <> .Child.LastSibling.Index
                    If Me.tvwMain_S.Nodes(i).Next.Image = "Dept" Then
                        CheckStop = False
                        Exit Function
                    End If
                    i = Me.tvwMain_S.Nodes(i).Next.Index
                Wend
            End If
        End With
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

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " 没有找到符合条件的信息！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " 已经定位完所有找到的信息，请重新输入条件！", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    With frmDeptManage.tvwMain_S
        If IsNull(mrsFind("上级ID")) Then
            If tbcDetails.Selected.Index = mint按层次 Then
                .Nodes("C" & mrsFind("ID")).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            Else
                strTemp = mrsFind!性质 & "|" & mrsFind!ID
                .Nodes("C" & strTemp).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            End If
        Else
            If tbcDetails.Selected.Index = mint按性质 Then
                strTemp = mrsFind!性质 & "|" & mrsFind!ID
                .Nodes("C" & strTemp).Selected = True
                .Nodes("C" & strTemp).Expanded = True
            Else
                .Nodes("C" & mrsFind("上级ID")).Selected = True
                .Nodes("C" & mrsFind("上级ID")).Expanded = True
            End If
            
            .SelectedItem.EnsureVisible
            frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            
            If tbcDetails.Selected.Index = mint按层次 Then
                frmDeptManage.lvwMain.ListItems("C" & mrsFind("ID")).Selected = True
                frmDeptManage.lvwMain.SelectedItem.EnsureVisible
                frmDeptManage.lvwMain_ItemClick frmDeptManage.lvwMain.SelectedItem
            End If
        End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
            gstrSQL = "Select a.id,a.上级id,a.名称,a.编码 ,c.编码 as 性质 From 部门表 A, 部门性质说明 B,部门性质分类 c Where b.工作性质=c.名称 " & _
                " and A.ID=B.部门ID  and (a.编码 like [1] or a.名称 like [2] or a.简码 like [3]) "
            
            If mnuViewShowStop.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "查询部门", UCase(txtFind.Text) & "%", UCase(txtFind.Text) & "%", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
End Sub

Private Function CheckBusiness(ByVal lngDeptID As Long) As Boolean
'功能：检查业务相关性
'参数：
'  lngDeptID：部门ID
'返回：True-通过；False非通过

    Dim strSQL As String, strMess As String
    Dim rsTmp As ADODB.Recordset, rsBusiness As ADODB.Recordset
    Dim lngSys As Long
    Dim dblStock As Double
    
    On Error GoTo hErr
    strSQL = "Select Count(1) Rec, 400 编号 From zlSystems Where 编号 = 400 And Nvl(共享号, 0) = 100 Union all " & vbCr & _
             "Select Count(1) Rec, 600 编号 From zlSystems Where 编号 = 600 And Nvl(共享号, 0) = 100 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取非标准系统信息")
    
    If rsTmp.EOF Then
        '没有共享安装其他系统，就不进行相关的业务检查
        rsTmp.Close
        CheckBusiness = True
        Exit Function
    End If
    
    Do While rsTmp.EOF = False
        dblStock = 0
        lngSys = NVL(rsTmp!编号, 0)
        Select Case lngSys
            Case 400    '物资系统
                If NVL(rsTmp!Rec, 0) = 1 Then
                    strMess = "该部门物资库存存在，请检查！"
                    strSQL = "Select Sum(实际数量) 实际数量 From 物资库存 Where 库房id = [1] "
                    Set rsBusiness = zlDatabase.OpenSQLRecord(strSQL, "检查物资库存", lngDeptID)
                    If rsBusiness.EOF = False Then
                        dblStock = NVL(rsBusiness!实际数量, 0)
                    End If
                    rsBusiness.Close
                End If
            Case 600    '设备系统
                If NVL(rsTmp!Rec, 0) = 1 Then
                    strMess = "该部门设备库存存在，请检查！"
                    strSQL = "Select Sum(实际数量) 实际数量 From 设备库存 Where 库房id = [1] "
                    Set rsBusiness = zlDatabase.OpenSQLRecord(strSQL, "检查设备库存", lngDeptID)
                    If rsBusiness.EOF = False Then
                        dblStock = NVL(rsBusiness!实际数量, 0)
                    End If
                    rsBusiness.Close
                End If
        End Select
        
        If dblStock <> 0 Then
            '存在库存数量
            rsTmp.Close
            MsgBox strMess, vbInformation, gstrSysName
            Exit Function
        End If
        
        rsTmp.MoveNext
    Loop
    
    CheckBusiness = True
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function
    
