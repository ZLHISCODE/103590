VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm发药窗口 
   Caption         =   "发药窗口管理"
   ClientHeight    =   4980
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   6648
   Icon            =   "frm发药窗口.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "93"
   Begin MSComctlLib.ImageList LvwBlack 
      Left            =   2730
      Top             =   2280
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList TvwImg 
      Left            =   2580
      Top             =   1020
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":0D90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList LvwColor 
      Left            =   2700
      Top             =   1680
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":13C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw 
      Height          =   2235
      Left            =   3120
      TabIndex        =   1
      Top             =   1380
      Width           =   2595
      _ExtentX        =   4572
      _ExtentY        =   3937
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "LvwBlack"
      SmallIcons      =   "LvwColor"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "名称"
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "编码"
         Text            =   "编码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "状态"
         Text            =   "状态"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "药房"
         Text            =   "药房"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "专家"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "叫号窗口"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   3345
      Left            =   240
      TabIndex        =   0
      Top             =   990
      Width           =   2445
      _ExtentX        =   4318
      _ExtentY        =   5906
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "TvwImg"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4290
      Top             =   660
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":16DE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":18FE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":1B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":1D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":1F58
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":2178
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":2398
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":25B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":2ACA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":2CEA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5040
      Top             =   630
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":2F0A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":312A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":334A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":3564
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":3784
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":39A4
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":3BC4
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":3DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":42F6
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm发药窗口.frx":4516
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6648
      _ExtentX        =   11726
      _ExtentY        =   1164
      BandCount       =   1
      _CBWidth        =   6648
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "Tbar"
      MinHeight1      =   612
      Width1          =   8376
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   612
         Left            =   24
         TabIndex        =   4
         Top             =   24
         Width           =   6552
         _ExtentX        =   11557
         _ExtentY        =   1080
         ButtonWidth     =   783
         ButtonHeight    =   1080
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新增"
               Key             =   "Add"
               Object.ToolTipText     =   "增加发药窗口"
               Object.Tag             =   "新增"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上班"
               Key             =   "Start"
               Object.ToolTipText     =   "上班"
               Object.Tag             =   "上班"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下班"
               Key             =   "Stop"
               Object.ToolTipText     =   "下班"
               Object.Tag             =   "下班"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看"
               Object.Tag             =   "查看"
               ImageIndex      =   8
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3228
      ScaleWidth      =   48
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   4620
      Width           =   6645
      _ExtentX        =   11726
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frm发药窗口.frx":4736
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6689
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileset 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
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
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "上班(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "下班(&T)"
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
   Begin VB.Menu mnuview 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuviewspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewText 
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
      Begin VB.Menu mnuViewShow 
         Caption         =   "包含下班窗口(&H)"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuviewr 
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
         Caption         =   "WEB上的中联(&W)"
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
End
Attribute VB_Name = "frm发药窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RecData As New ADODB.Recordset  '发药窗口记录集
Private BlnStartUp As Boolean
Private blnFirst As Boolean
Private LngLastRoot As Long
Private mlngMode As Long
Private mstrPrivs As String             '当前用户具有的当前模块的功能
Private mstr药房 As String
Dim mbln所有部门 As Boolean             '是否具有所有部门权限
Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    blnFirst = True
    LngLastRoot = 0
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mbln所有部门 = IsHavePrivs(mstrPrivs, "所有部门")
   
    mnuViewIcon_Click Me.Lvw.View
    If LoadInTree = False Then Exit Sub
    权限控制
    
    BlnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 8000 Then
        Me.Width = 8000
        Exit Sub
    End If
    
    If blnFirst Then picSplit.Left = 3100
    With picSplit
        .Top = IIf(Cbar.Visible = False, 0, Cbar.Height)
        .Height = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - .Top
    End With
    
    With Tree
        .Top = picSplit.Top
        .Width = picSplit.Left
        .Height = picSplit.Height
        .Left = 0
    End With
    
    With Lvw
        .Top = picSplit.Top
        .Left = picSplit.Left + picSplit.Width
        .Width = Me.ScaleWidth - .Left
        .Height = picSplit.Height
    End With
    blnFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw
        .Sorted = False
        .SortKey = ColumnHeader.index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw_DblClick()
    If Lvw.ListItems.count = 0 Then Exit Sub
    If Lvw.SelectedItem Is Nothing Then Exit Sub
    If Tbar.Buttons("Add").Visible = False Then Exit Sub
    
    mnuEditModify_Click
End Sub

Private Sub Lvw_ItemClick(ByVal Item As MSComctlLib.listItem)
     mnuEditStop.Enabled = IIf(Item.SubItems(2) = "上班", True, False)
     mnuEditStart.Enabled = mnuEditStop.Enabled Xor True
     Tbar.Buttons("Start").Enabled = mnuEditStart.Enabled
     Tbar.Buttons("Stop").Enabled = mnuEditStop.Enabled
End Sub

Private Sub Lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Lvw.ListItems.count = 0 Then Exit Sub
    If Lvw.SelectedItem Is Nothing Then Exit Sub
    If Tbar.Buttons("Add").Visible = False Then Exit Sub
    
    mnuEditModify_Click
End Sub

Private Sub Lvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If Tbar.Buttons("Add").Visible Or Tbar.Buttons("Start").Visible Then
        PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub mnuEditAdd_Click()
    With Frm发药窗口编辑
        .EditState = 1
        .Show 1, Me
    End With
    mnuviewr_Click
End Sub

Private Sub mnuEditDelete_Click()
    Dim dteBegin As Date, dteEnd As Date
    Dim lng药房ID As Long
    Dim strWinName As String
    Dim rsData As ADODB.Recordset
    Dim strDeptNode As String
    Dim blnAllStop As Boolean
    
    On Error GoTo ErrHand
    
    dteEnd = zldatabase.Currentdate
    dteBegin = DateAdd("D", -3, dteEnd)
    lng药房ID = Val(Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3))
    strWinName = Lvw.SelectedItem.Text
    strDeptNode = GetDeptStationNode(lng药房ID)
    
    gstrSQL = "Select 1 From 未发药品记录 Where 库房id = [1] And 发药窗口 = [2] And 填制日期 Between [3] And [4] And Rownum < 2 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lng药房ID, strWinName, dteBegin, dteEnd)
    
    '如果存在未发药的就打开调整发药窗口窗口
    If rsData.EOF Then
        If MsgBox("你确定要删除该发药窗口吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        MsgBox "该窗口还有未发药处方，请将其调整至其他窗口。", vbInformation, gstrSysName
        If frm调整发药窗口.ShowMe(lng药房ID, Me, dteBegin, dteEnd, strDeptNode, strWinName) = False Then
            If MsgBox("未将未发处方调整到其他发药窗口，是否坚持删除该发药窗口？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    gstrSQL = "zl_发药窗口_delete("
    '编码
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '库房ID
    gstrSQL = gstrSQL & "," & Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3)
    gstrSQL = gstrSQL & ")"
    
    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditModify_Click()
    With Frm发药窗口编辑
        .EditState = 2
        .Show 1, Me
    End With
    mnuviewr_Click
End Sub

Private Sub mnuEditStart_Click()
    gstrSQL = "zl_发药窗口_setwork("
    '编码
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '库房ID
    gstrSQL = gstrSQL & "," & Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3)
    '是否上班
    gstrSQL = gstrSQL & ",1"
    gstrSQL = gstrSQL & ")"

    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim dteBegin As Date, dteEnd As Date
    Dim lng药房ID As Long
    Dim strWinName As String
    Dim rsData As ADODB.Recordset
    Dim strDeptNode As String
    Dim blnAllStop As Boolean
    
    On Error GoTo ErrHand
    
    dteEnd = zldatabase.Currentdate
    dteBegin = DateAdd("D", -3, dteEnd)
    lng药房ID = Val(Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3))
    strWinName = Lvw.SelectedItem.Text
    strDeptNode = GetDeptStationNode(lng药房ID)
    
    gstrSQL = "select 1 from 发药窗口 where 药房id=[1] and 上班否=1 And 名称<>[2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lng药房ID, strWinName)
    If rsData.RecordCount = 0 Then blnAllStop = True
    
    gstrSQL = "Select 1 From 未发药品记录 Where 库房id = [1] And 发药窗口 = [2] And 填制日期 Between [3] And [4] And Rownum < 2 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "mnuEditStop_Click", lng药房ID, strWinName, dteBegin, dteEnd)
    
    '如果存在未发药的就打开调整发药窗口窗口
    If rsData.RecordCount > 0 Then
        If blnAllStop Then
            If MsgBox("该窗口还有未发药处方，其他窗口都已下班，是否坚持将当前窗口设置为下班？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("该窗口还有未发药处方，是否坚持将当前窗口设置为下班？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_发药窗口_setwork("
            '编码
            gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
            '库房ID
            gstrSQL = gstrSQL & "," & Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3)
            '是否上班
            gstrSQL = gstrSQL & ",0"
            gstrSQL = gstrSQL & ")"
        
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            MsgBox "请将未发药处方调整至其他窗口。", vbInformation, gstrSysName
            Call frm调整发药窗口.ShowMe(lng药房ID, Me, dteBegin, dteEnd, strDeptNode, strWinName)
            mnuviewr_Click
            Exit Sub
        End If
    End If
    
    gstrSQL = "zl_发药窗口_setwork("
    '编码
    gstrSQL = gstrSQL & "'" & Lvw.SelectedItem.SubItems(1) & "'"
    '库房ID
    gstrSQL = gstrSQL & "," & Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3)
    '是否上班
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"

    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mnuviewr_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    '默认参数：药房=药房id，发药窗口=发药窗口名称
    Dim Str窗口 As String
    If Not Me.Lvw.SelectedItem Is Nothing Then
        Str窗口 = Me.Lvw.SelectedItem.Text
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "药房=" & IIf(LngLastRoot = 0, "", LngLastRoot), _
        "发药窗口=" & Str窗口)
End Sub

Private Sub mnuViewButton_Click()
    mnuViewButton.Checked = Not mnuViewButton.Checked
    Cbar.Visible = mnuViewButton.Checked
    mnuViewText.Enabled = mnuViewButton.Checked
    Cbar.Bands("only").MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(index).Checked = True
    
    Select Case index
        Case 0
            Lvw.View = lvwIcon
        Case 1
            Lvw.View = lvwSmallIcon
        Case 2
            Lvw.View = lvwList
        Case 3
            Lvw.View = lvwReport
    End Select
End Sub

Private Sub mnuviewr_Click()
    If LoadInTree = False Then
        BlnStartUp = False
        Form_Activate
    End If
End Sub

Private Sub mnuViewShow_Click()
    mnuViewShow.Checked = mnuViewShow.Checked Xor True
    mnuviewr_Click
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewText_Click()
    Dim buttTemp As Button
    
    mnuViewText.Checked = Not mnuViewText.Checked
    For Each buttTemp In Tbar.Buttons
        If mnuViewText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    Cbar.Bands("only").MinHeight = Tbar.Height
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "发药窗口"
    Set objPrint.Body.objData = Lvw
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")

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

Private Function LoadInLvw(Optional ByVal Bln药房ID As Long = 0)
    Dim strCon As String
    
    If Not mbln所有部门 Then
        strCon = " And Id In (Select 部门id From 部门人员 Where 人员id = [1]) "
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select A.编码,A.名称,A.上班否,A.药房ID,A.专家,B.名称 as 药房,A.叫号窗口 From 发药窗口 A,部门表 B" & _
    " Where A.药房ID=B.ID " & strCon
    If Bln药房ID <> 0 Then gstrSQL = gstrSQL & " And 药房ID=[2]"
    If mnuViewShow.Checked = False Then gstrSQL = gstrSQL & " And 上班否=1"
    gstrSQL = gstrSQL & " Order by A.编码"
   Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, Bln药房ID)
   
   With RecData
        Lvw.ListItems.Clear
        Do While Not .EOF
            If !上班否 = 1 Then
                Lvw.ListItems.Add , "K_" & !药房ID & "," & !编码, !名称, 1, 1
            Else
                Lvw.ListItems.Add , "K_" & !药房ID & "," & !编码, !名称, 2, 2
            End If
            Lvw.ListItems("K_" & !药房ID & "," & !编码).SubItems(1) = !编码
            Lvw.ListItems("K_" & !药房ID & "," & !编码).SubItems(2) = IIf(!上班否 = 1, "上班", "下班")
            Lvw.ListItems("K_" & !药房ID & "," & !编码).SubItems(3) = !药房
            Lvw.ListItems("K_" & !药房ID & "," & !编码).SubItems(4) = IIf(IsNull(!专家), "", IIf(!专家 = 1, "√", ""))
            Lvw.ListItems("K_" & !药房ID & "," & !编码).SubItems(5) = zlStr.Nvl(!叫号窗口)
            .MoveNext
        Loop
        If Bln药房ID <> 0 Then
            Lvw.ColumnHeaders(4).Width = 0
        Else
            Lvw.ColumnHeaders(4).Width = 1500
        End If
        
        If .RecordCount = 0 Then
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            Tbar.Buttons("Modify").Enabled = False
            Tbar.Buttons("Delete").Enabled = False
            Tbar.Buttons("Start").Enabled = False
            Tbar.Buttons("Stop").Enabled = False
        Else
            mnuEditModify.Enabled = True
            mnuEditDelete.Enabled = True
            Tbar.Buttons("Modify").Enabled = True
            Tbar.Buttons("Delete").Enabled = True
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadInTree() As Boolean
    Dim strCon As String
    
    LoadInTree = False
    On Error GoTo errHandle
    If Not mbln所有部门 Then
        strCon = " And Id In (Select 部门id From 部门人员 Where 人员id = [1]) "
    End If
    
    gstrSQL = " Select ID,编码,名称 From 部门表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (" & _
          " Select distinct 部门ID From 部门性质说明" & _
          " Where 工作性质 Like '%药房')" & strCon & _
          " And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' Order by 编码"
       
    Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
        
    With RecData
        If .EOF Then
            MsgBox "药库药房信息不全（部门管理）或者你不是药房人员。", vbInformation, gstrSysName
            Exit Function
        End If
        
        Tree.Nodes.Clear
        Tree.Nodes.Add , , "R", "所有药房", 1, 1
        
        Do While Not .EOF
            Tree.Nodes.Add "R", 4, "K_" & !Id, "【" & !编码 & "】" & !名称, 2, 2
            .MoveNext
        Loop
        If LngLastRoot <> 0 Then
            Tree.Nodes("K_" & LngLastRoot).Selected = True
        Else
            Tree.Nodes("R").Selected = True
        End If
        Tree.SelectedItem.Selected = True
        Tree.SelectedItem.Expanded = True
        Tree_NodeClick Tree.SelectedItem
    End With
    
    LoadInTree = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 权限控制()
    If Not IsHavePrivs(mstrPrivs, "增删改") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit1.Visible = False

        Tbar.Buttons("Add").Visible = False
        Tbar.Buttons("Modify").Visible = False
        Tbar.Buttons("Delete").Visible = False
        Tbar.Buttons("split1").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "上下班") Then
        If Not IsHavePrivs(mstrPrivs, "增删改") Then
            mnuEdit.Visible = False
        Else
            mnuEditStart.Visible = False
            mnuEditStop.Visible = False
            mnuEditSplit1.Visible = False
        End If
        Tbar.Buttons("Start").Visible = False
        Tbar.Buttons("Stop").Visible = False
        Tbar.Buttons("split2").Visible = False
    End If
End Function

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    If picSplit.Left + x < 3000 Then Exit Sub
    If picSplit.Left + x > Me.ScaleWidth - 3000 Then Exit Sub
    
    picSplit.Left = picSplit.Left + x
    Form_Resize
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreView_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "View"
            If Lvw.View < lvwReport Then
                mnuViewIcon_Click Lvw.View + 1
            Else
                mnuViewIcon_Click 0
            End If
        Case "Help"
            mnuHelpTitle_Click
        Case "Quit"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    mnuViewIcon_Click ButtonMenu.index - 1
End Sub

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuview, 2
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "R" Then
        mnuEditAdd.Enabled = False
        Tbar.Buttons("Add").Enabled = False
        LngLastRoot = 0
        mstr药房 = ""
    Else
        mnuEditAdd.Enabled = True
        Tbar.Buttons("Add").Enabled = True
        LngLastRoot = Mid(Node.Key, 3)
        mstr药房 = Mid(Node.Text, InStr(1, Node.Text, "】") + 1)
    End If
    
    LoadInLvw IIf(Node.Key = "R", 0, Mid(Node.Key, 3))
    If Lvw.ListItems.count > 0 Then
        Lvw.ListItems(1).Selected = True
        Lvw.SelectedItem.Selected = True
        Lvw_ItemClick Lvw.SelectedItem
        
        mnuFilePrint.Enabled = True
        mnuFilePreview.Enabled = True
        mnuFileExcel.Enabled = True
        Tbar.Buttons("Preview").Enabled = True
        Tbar.Buttons("Print").Enabled = True
    Else
        mnuFilePrint.Enabled = False
        mnuFilePreview.Enabled = False
        mnuFileExcel.Enabled = False
        Tbar.Buttons("Preview").Enabled = False
        Tbar.Buttons("Print").Enabled = False
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

