VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPatientHistoryQuery 
   Caption         =   "病人历史记录查询"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   7890
   Icon            =   "frmPatientHistoryQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleMode       =   0  'User
      ScaleWidth      =   1608.75
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2145
   End
   Begin MSComctlLib.ListView LivItem 
      Height          =   1815
      Left            =   90
      TabIndex        =   11
      Top             =   3120
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LivInfo 
      Height          =   1725
      Left            =   3780
      TabIndex        =   8
      Top             =   1890
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin C1Chart2D8.Chart2D ChartMain 
      Height          =   1695
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1635
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   2884
      _ExtentY        =   2990
      _StockProps     =   0
      ControlProperties=   "frmPatientHistoryQuery.frx":020A
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   2730
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3195
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   45
   End
   Begin MSComctlLib.ListView LivPatient 
      Height          =   1635
      Left            =   30
      TabIndex        =   3
      Top             =   960
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   2940
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":078D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":09AD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":0BCD
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":0DED
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":100D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":122D
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":144D
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":166D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1889
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1AA9
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1CC9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   2910
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1EE3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2103
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2323
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2543
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2763
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2983
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2BA3
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2DC3
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2FDF
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":31FF
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":341F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   7890
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "科室"
      Child2          =   "CmbDepartment"
      MinWidth2       =   2505
      MinHeight2      =   300
      Width2          =   2685
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
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
               Caption         =   "过滤"
               Key             =   "Filter"
               Object.Tag             =   "过滤"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
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
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CmbDepartment 
         Height          =   300
         Left            =   5295
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5055
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   635
      SimpleText      =   $"frmPatientHistoryQuery.frx":3639
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatientHistoryQuery.frx":3680
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8837
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
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   3645
      Left            =   3570
      TabIndex        =   6
      Top             =   900
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   6429
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数据"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "图形"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label LabItem 
      AutoSize        =   -1  'True
      Caption         =   "检验项目"
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label LabPatient 
      AutoSize        =   -1  'True
      Caption         =   "病人信息"
      Height          =   180
      Left            =   30
      TabIndex        =   9
      Top             =   750
      Width           =   720
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
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "参数设置(&M)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
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
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
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
Attribute VB_Name = "frmPatientHistoryQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseStartX As Single                       '移动前鼠标的位置X
Dim MouseStartY As Single                       '移动前鼠标的位置Y
Dim OutPatient As Boolean                       '=True门诊病人=False住院病人
Dim StartDate As Date, EndDate As Date          '用于过滤的开始结束时间
Dim PatientInfo As String                       '病人信息(通过门诊、住院、姓名)
Dim NowFocus As Integer                         '=1选中了LivPatient;=2选中了LivItem;=3选中了LivInfo

Private Sub CmbDepartment_Click()
    '读入部门下的病人
    If Len(Me.CmbDepartment.Text) > 0 Then
        LoadPatientInfo Me.CmbDepartment.ItemData(Me.CmbDepartment.ListIndex)
    Else
        Me.LivPatient.ListItems.Clear
    End If
    '读入病人的检验项目
    If Me.LivPatient.ListItems.Count > 0 Then
        Me.LivPatient.ListItems(1).Selected = True
        LoadItem Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
        If Me.LivItem.ListItems.Count > 0 Then
            '读出数据并画线
            LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
        End If
    Else
        Me.LivItem.ListItems.Clear
        Me.LivInfo.ListItems.Clear
    End If
End Sub

Private Sub CoolBar1_Resize()
    Form_Resize
End Sub

Private Sub Form_Load()
    '初使化
    Initialization
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'LabPatient
    Me.LabPatient.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    Me.LabPatient.Left = 0
    
    'LivPatient
    Me.LivPatient.Left = 0
    Me.LivPatient.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + Me.LabPatient.Height
    Me.LivPatient.Width = Me.picSplit.Left
    Me.LivPatient.Height = Me.picSplit1.Top - Me.LabPatient.Height - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    
    'picSplit1
    Me.picSplit1.Left = 0
    Me.picSplit1.Width = Me.picSplit.Left
    
    'LabItem
    Me.LabItem.Top = Me.picSplit1.Top + Me.picSplit1.Height
    Me.LabItem.Left = 0
    
    'LivItem
    Me.LivItem.Top = Me.LabItem.Top + Me.LabItem.Height
    Me.LivItem.Left = 0
    Me.LivItem.Width = Me.LivPatient.Width
    Me.LivItem.Height = Me.ScaleHeight - Me.LabItem.Top - Me.LabItem.Height - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    'picSplit
    Me.picSplit.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    Me.picSplit.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    'Tabstrip
    Me.TabStrip.Left = Me.picSplit.Left + Me.picSplit.Width
    Me.TabStrip.Top = Me.LabPatient.Top
    Me.TabStrip.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    Me.TabStrip.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    
    'LivInfo
    Me.LivInfo.Top = Me.TabStrip.Top + 300
    Me.LivInfo.Left = Me.TabStrip.Left + 30
    Me.LivInfo.Height = Me.TabStrip.Height - 60 - 300
    Me.LivInfo.Width = Me.TabStrip.Width - 60
    
    'ChartMain
    Me.ChartMain.Top = Me.LivInfo.Top + 30
    Me.ChartMain.Left = Me.LivInfo.Left + 30
    Me.ChartMain.Width = Me.LivInfo.Width - 60
    Me.ChartMain.Height = Me.LivInfo.Height - 60
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '退出时保存私有设置
    SaveWinState Me, App.ProductName
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "门诊", OutPatient
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "开始日期", StartDate
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "结束日期", EndDate
End Sub

Private Sub LivInfo_Click()
    NowFocus = 3
End Sub

Private Sub LivItem_Click()
    NowFocus = 2
End Sub

Private Sub LivItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '读出数据
    LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
End Sub

Private Sub LivItem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.stbThis.Panels(2).Text = "提示:按住<Ctrl>后可用鼠标选中多个检验项目!"
End Sub

Private Sub LivPatient_Click()
    NowFocus = 1
End Sub

Private Sub LivPatient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '读入项目
    LoadItem Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
    If Me.LivItem.ListItems.Count > 0 Then
        '读出数据并画线
        LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
    End If
End Sub
Private Sub mnuFileExcel_Click()
    '输出Excel
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    '退出
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    '预览
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    subPrint 1
End Sub

Private Sub mnuFileSet_Click()
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuFileSetup_Click()
    '设数设置
    '过滤
    With frmPatientHistorySetup
        If OutPatient = True Then
            .OptInPatient.Value = 0
            .OptOutPatient.Value = 1
        Else
            .OptInPatient.Value = 1
            .OptOutPatient.Value = 0
        End If
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuHelp_Click()
    '显示帮助
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpAbout_Click()
     '显示关于
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub mnuHelpWebHome_Click()
    '显示主页
    Call zlHomePage(Me.Hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送Email
    Call zlMailTo(Me.Hwnd)
End Sub

Private Sub mnuViewFilter_Click()
    '过滤
    With frmPatientHistoryFilter
        .TxtPatient = PatientInfo
        .DTPBegin = StartDate
        .DTPEND = EndDate
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    '显示或隐藏标准按钮
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    '显示或隐藏文字
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

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseStartX = x
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveTmp As Single
    '暂时屏蔽方便查问题
    On Error Resume Next
    If Button = 1 Then
        
        '得到移动后的位置
        MoveTmp = Me.picSplit.Left + x - MouseStartX
        
        '超过最大或最小宽度时退出
        If MoveTmp <= 2000 Or Me.ScaleWidth - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit
        Me.picSplit.Left = MoveTmp
        
        'picSplit1
        Me.picSplit1.Width = Me.picSplit.Left
        
        'Livpatient
        Me.LivPatient.Width = Me.picSplit.Left
        
        'LivItem
        Me.LivItem.Width = Me.picSplit.Left
        
        'TabStrip
        Me.TabStrip.Left = Me.picSplit.Left + Me.picSplit.Width
        Me.TabStrip.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
        Me.TabStrip.Refresh
        
        'LivInfo
        Me.LivInfo.Left = Me.TabStrip.Left + 30
        Me.LivInfo.Width = Me.TabStrip.Width - 60
        Me.LivInfo.Refresh
        
        'Chartmain
        Me.ChartMain.Left = Me.LivInfo.Left + 30
        Me.ChartMain.Width = Me.LivInfo.Width - 60
        Me.ChartMain.Refresh
    End If
End Sub



Private Sub picSplit1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseStartY = y
    End If
End Sub

Private Sub picSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveTmp As Single
    '暂时屏蔽方便查问题
    On Error Resume Next
    If Button = 1 Then
        
        '得到移动后的位置
        MoveTmp = Me.picSplit1.Top + y - MouseStartY
        
        '超过最大或最小宽度时退出
        If MoveTmp <= 2000 Or Me.ScaleHeight - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit1
        Me.picSplit1.Top = MoveTmp
        
        'LivPatient
        Me.LivPatient.Height = Me.picSplit1.Top - Me.LabPatient.Height - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
        
        'LabItem
        Me.LabItem.Top = Me.picSplit1.Top + Me.picSplit1.Height
        
        'LivItem
        Me.LivItem.Top = Me.LabItem.Top + Me.LabItem.Height
        Me.LivItem.Height = Me.ScaleHeight - Me.LabItem.Top - Me.LabItem.Height - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
        
    End If
End Sub

Private Sub TabStrip_Click()
    '显示数据或图形
    If Me.TabStrip.SelectedItem.Index = 1 Then
        Me.ChartMain.Visible = False
        Me.LivInfo.Visible = True
    Else
        Me.LivInfo.Visible = False
        Me.ChartMain.Visible = True
    End If
End Sub

Sub LoadDepartmental(OutorIn As Boolean)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               读入门诊或住院科室
    '参数
    '    OutorIN        =True 表示门诊 = False 表示住院
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    If OutorIn = True Then
        gstrSql = "select DISTINCT a.科室id as id ,b.名称 as 部门名称 " & _
                  " from 病人挂号汇总 a , 部门表 b " & _
                  " Where a.科室id = b.ID "
    Else
        gstrSql = "select DISTINCT a.当前科室id as id , b.名称 as 部门名称 " & _
                  " from 病人信息 a , 部门表 b " & _
                  " Where a.当前科室id = b.ID " & _
                  " and a.当前科室id is not null"
    End If
    
    Me.CmbDepartment.Clear
    
    Me.MousePointer = 11
    
    OpenRecord rsTmp, gstrSql, Me.Caption
   
    Do Until rsTmp.EOF
        Me.CmbDepartment.AddItem rsTmp("部门名称")
        Me.CmbDepartment.ItemData(i) = rsTmp("ID")
        i = i + 1
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Me.MousePointer = 1
End Sub
Sub Initialization()
    ''''''''''''''''''''''''''
    '功能           初使化
    ''''''''''''''''''''''''''
    
    LoadColHead         '写入表头
    
    '恢复私有设置
    RestoreWinState Me, App.ProductName
    
    StartDate = date
    StartDate = DateAdd("d", -DatePart("d", date) + 1, date)
    EndDate = DateAdd("d", -1, DateAdd("m", 1, StartDate))
    
    OutPatient = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "门诊", "True")
    StartDate = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "开始日期", StartDate)
    EndDate = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\过滤", "结束日期", EndDate)
    
    '读入部门
    LoadDepartmental OutPatient
    
    '设置标注
    Me.ChartMain.Header.Text = "病人检验历史图形"
    Me.ChartMain.Header.Font.Size = 12
    Me.ChartMain.Header.Interior.ForegroundColor = vbBlue
    
    'X/Y轴标注
    Me.ChartMain.ChartArea.Axes("X").Title.Text = "时间"
    Me.ChartMain.ChartArea.Axes("Y").Title.Text = "结果"
    
    NowFocus = 1
End Sub
Sub LoadPatientInfo(DepartmentID As Long)
    ''''''''''''''''''''''''''''''''''''
    '功能              读出科室下的病人
    '参数
    '    Department    部门ID
    ''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    
    Me.LivPatient.ListItems.Clear
    
    gstrSql = " Select DISTINCT d.病人id, d.姓名, d.性别, d.年龄 from 检验标本记录 a , 检验普通结果 b , 病人医嘱记录 c , 病人信息 d " & _
              " Where a.ID = b.检验标本id " & _
              " and a.报告结果 = b.记录类型 " & _
              " and a.医嘱id = c.id " & _
              " and c.病人id = d.病人id " & _
              " and d.当前科室ID = " & DepartmentID & _
              " and a.检验时间 Between [2] and [3]"
    
    If Len(Trim(PatientInfo)) > 0 Then
        gstrSql = gstrSql & " and (d.住院号 like [1] " & _
                                   " or d.门诊号 like [1] " & _
                                   " or upper(d.姓名) like  upper([1]) )"
    End If
    
    Me.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientInfo, CDate(StartDate), CDate(EndDate))
    
    
    Do Until rsTmp.EOF
        With Me.LivPatient
            Set ItmX = .ListItems.Add(, "A" & rsTmp("病人ID"), rsTmp("姓名"))
            ItmX.SubItems(1) = rsTmp("性别")
            ItmX.SubItems(2) = rsTmp("年龄")
        End With
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
    
    Me.MousePointer = 1
End Sub
Sub LoadItem(PatientID As Long)
    '''''''''''''''''''''''''''''''''''
    '功能               读入检验项目
    '''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    
    Me.LivItem.ListItems.Clear
    
    gstrSql = " select distinct d.id , d.编码 , d.中文名 , d.英文名 " & _
              " from 检验标本记录 a , 检验普通结果 b , 病人医嘱记录 c , 诊治所见项目 d " & _
              " Where a.ID = b.检验标本id " & _
              " and a.报告结果 = b.记录类型 " & _
              " and a.医嘱id = c.id " & _
              " and b.检验项目id = d.id " & _
              " and c.病人id = " & PatientID & _
              " and a.检验时间 Between [2] and [3]"
              
    Me.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientID, CDate(StartDate), CDate(EndDate))
    
    Do Until rsTmp.EOF
        With Me.LivItem
            Set ItmX = .ListItems.Add(, "A" & rsTmp("ID"), zlCommFun.Nvl(rsTmp("编码")))
            ItmX.SubItems(1) = zlCommFun.Nvl(rsTmp("中文名"))
            ItmX.SubItems(2) = zlCommFun.Nvl(rsTmp("英文名"))
        End With
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
    Me.MousePointer = 1
End Sub
Sub LoadColHead()
    '''''''''''''''''''''''''''''''
    '功能               初使化表头
    '''''''''''''''''''''''''''''''
    
    With Me.LivPatient
        .ColumnHeaders.Add , "A", "姓名", 1100
        .ColumnHeaders.Add , "B", "性别", 700
        .ColumnHeaders.Add , "C", "年龄", 700
    End With
    
    With Me.LivItem
        .ColumnHeaders.Add , "A", "编码", 1000
        .ColumnHeaders.Add , "B", "中文名", 800
        .ColumnHeaders.Add , "C", "英文名", 800
        .ColumnHeaders.Add , "D", "缩写", 800
    End With
    
    With Me.LivInfo
        .ColumnHeaders.Add , "A", "检验项目", 1100
        .ColumnHeaders.Add , "B", "检验时间", 2000
        .ColumnHeaders.Add , "C", "检验人", 1000
        .ColumnHeaders.Add , "D", "审核人", 1000
        .ColumnHeaders.Add , "E", "结果", 1100
    End With
End Sub
Sub LoadInfo(PatientID As Long)
    '''''''''''''''''''''''''''''''''''''''''
    '功能               读入病人的检验结果
    '参数
    '    PatientID      病人ID
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    Dim DateTmp As Date
    Dim XX As Variant, YY As Variant
    Dim i As Integer, j As Integer, N As Integer
    Dim NextID As Long
    Dim ItemID As Long
    
    Me.LivInfo.ListItems.Clear
    
    With Me.ChartMain.ChartGroups(1)
        '清除线
        .Data.IsBatched = True
        .SeriesLabels.RemoveAll
        .PointLabels.RemoveAll
        .Data.NumSeries = 0
        .Data.IsBatched = False
    End With
    
    Me.ChartMain.ChartGroups(1).Data.IsBatched = True
    
    For i = 1 To Me.LivItem.ListItems.Count
        
        If Me.LivItem.ListItems(i).Selected = True Then
            
            ItemID = Mid(Me.LivItem.ListItems(i).Key, 2)
    
            gstrSql = " select a.id ,a.标本序号, a.检验时间 , a.检验人 , a.审核人 , b.检验结果 " & _
                      " from 检验标本记录 a , 检验普通结果 b , 病人医嘱记录 c " & _
                      " Where a.ID = b.检验标本id " & _
                      " and a.报告结果 = b.记录类型 " & _
                      " and a.医嘱id = c.id " & _
                      " and a.审核人 is not null " & _
                      " and c.病人ID = [1] " & _
                      " and b.检验项目ID = [2] " & _
                      " order by 检验时间"
              
            On Error GoTo Herr
                          
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientID, ItemID)
    
            '写入数据
            Do Until rsTmp.EOF
                With Me.LivInfo
                    NextID = NextID + 1
                    Set ItmX = .ListItems.Add(, "A" & NextID, Me.LivItem.ListItems(i).SubItems(1))
                    ItmX.SubItems(1) = zlCommFun.Nvl(rsTmp("检验时间"))
                    ItmX.SubItems(2) = zlCommFun.Nvl(rsTmp("检验人"))
                    ItmX.SubItems(3) = zlCommFun.Nvl(rsTmp("审核人"))
                    ItmX.SubItems(4) = zlCommFun.Nvl(rsTmp("检验结果"))
                End With
                rsTmp.MoveNext
            Loop
            
            If rsTmp.RecordCount > 0 Then
                '移动到开始位置
                rsTmp.MoveFirst
            End If
            
            '没有记录时不画线退出
            If rsTmp.EOF Then Exit Sub
            
            
            '画线
            With Me.ChartMain.ChartGroups(1)
                
                .Data.Layout = oc2dDataGeneral  '数据设置方式为每个Series拥有各自的X Points
                
                j = j + 1
                
                .Data.NumSeries = j             '设置有几条线
                
                .Data.NumPoints(j) = rsTmp.RecordCount
                Me.ChartMain.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
                
                ReDim XX(rsTmp.RecordCount - 1) As Date
                ReDim YY(rsTmp.RecordCount - 1) As Double
                
                N = 0
                
                Do Until rsTmp.EOF
                    DateTmp = Format(rsTmp("检验时间"), "yyyy-MM-dd HH:mm:ss")
                    XX(N) = DateTmp
                    YY(N) = Val(zlCommFun.Nvl(rsTmp("检验结果")))
                    N = N + 1
                    rsTmp.MoveNext
                Loop
                
                .Data.CopyXVectorIn j, XX
                .Data.CopyYVectorIn j, YY
                
                '图表旁标
                .SeriesLabels.Add Me.LivItem.ListItems(i).SubItems(1)
                Me.ChartMain.Legend.Anchor = oc2dAnchorNorth            '旁标位置
                Me.ChartMain.Legend.Orientation = oc2dOrientHorizontal  '旁标方向
                
                Select Case j
                    Case 1
                        .Styles(j).Symbol.Shape = oc2dShapeBox
                    Case 2
                        .Styles(j).Symbol.Shape = oc2dShapeCircle
                    Case 3
                        .Styles(j).Symbol.Shape = oc2dShapeCross
                    Case 4
                        .Styles(j).Symbol.Shape = oc2dShapeDiagonalCross
                    Case 5
                        .Styles(j).Symbol.Shape = oc2dShapeDiamond
                End Select

            End With
            rsTmp.Close
        End If
    Next
    Me.ChartMain.ChartGroups(1).Data.IsBatched = False
    Me.ChartMain.Refresh
    Exit Sub
Herr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Sub GetFilterStr(Patient As String, BegingDate As Date, OverDate As Date)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   读取过滤字串
    '参数
    '    OutPatient         =True门诊病人;=False住院病人
    '    Patient            病人
    '    StartDate          开始日期
    '    EndDate            结束日期
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PatientInfo = Patient
    StartDate = BegingDate
    EndDate = OverDate
    
    '刷新
    CmbDepartment_Click
End Sub

Public Sub GetFilterDate(InPatient As Boolean)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   读取过滤字串
    '参数
    '    OutPatient         =True门诊病人;=False住院病人
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If OutPatient <> InPatient Then
        '读入住院或门诊
        LoadDepartmental InPatient
    End If
    
    OutPatient = InPatient
    
    '刷新
    CmbDepartment_Click

End Sub

Private Sub subPrint(bytMode As Byte)
    '''''''''''''''''''''''''''''''''''''''''''
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '''''''''''''''''''''''''''''''''''''''''''
    Dim objPrint As New zlPrintLvw
    
    If gstrUserName = "" Then Call GetUserInfo
    
    Select Case NowFocus
        Case 1
            If LivPatient.SelectedItem Is Nothing Then Exit Sub
    
            If LivPatient.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivPatient
        Case 2
            If LivItem.SelectedItem Is Nothing Then Exit Sub
    
            If LivItem.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivItem
        Case 3
            If LivInfo.SelectedItem Is Nothing Then Exit Sub
    
            If LivInfo.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivInfo
    End Select
    
    objPrint.Title.Text = "质控查询"
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    '处理按下按钮
    Select Case Button.Key
        Case "Quit"
            '退出
            mnuFileExit_Click
        Case "Print"
            '打印
            mnuFilePrint_Click
        Case "Preview"
            '预览
            mnuFilePreview_Click
        Case "Help"
            '帮助
            mnuHelp_Click
        Case "Filter"
            '过滤
            mnuViewFilter_Click
    End Select
End Sub
