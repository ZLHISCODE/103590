VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "PACS备份管理"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timAutoPolicy 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7260
      Top             =   720
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   4845
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   7710
      Top             =   570
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
            Picture         =   "frmMain.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":074A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":096A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DAA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FCA
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1626
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1846
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B60
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D7A
            Key             =   "Auto"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F94
            Key             =   "Filtrate"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   8310
      Top             =   570
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
            Picture         =   "frmMain.frx":21AE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23CE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25EE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":280E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A2E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C4E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E6E
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":308E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32AA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34CA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36EA
            Key             =   "Hand"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A04
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C1E
            Key             =   "Auto"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E38
            Key             =   "Filtrate"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9210
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
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
               Caption         =   "手动"
               Key             =   "ManualArchive"
               Object.ToolTipText     =   "手动归档"
               Object.Tag             =   "手动"
               ImageKey        =   "Hand"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "反归"
               Key             =   "ManualDeArchive"
               Object.ToolTipText     =   "手动反归档"
               Object.Tag             =   "反归"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "自动"
               Key             =   "AutoArchiveSetup"
               Object.ToolTipText     =   "自动归档"
               Object.Tag             =   "自动"
               ImageKey        =   "Auto"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filtrate"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filtrate"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "人员查看方式"
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
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   2
      Top             =   6015
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   635
      SimpleText      =   $"frmMain.frx":4052
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":4099
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6750
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditManualArchive 
         Caption         =   "手动归档(&M)"
      End
      Begin VB.Menu mnuEditManualDeArchive 
         Caption         =   "手动反归档(&D)"
      End
      Begin VB.Menu mnuEditAutoArchiveSetup 
         Caption         =   "自动归档(&A)"
      End
      Begin VB.Menu mnuEditCollate 
         Caption         =   "校对(&L)"
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
      Begin VB.Menu mnuEditFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&


Dim rsData As New ADODB.Recordset

Private WithEvents mobjIcon As clsTaskIcon  '托盘类
Attribute mobjIcon.VB_VarHelpID = -1


Sub InitLiv()
    With LivMain
        .ColumnHeaders.Add , "A", "姓名"
        .ColumnHeaders.Add , "B", "影像类别"
        .ColumnHeaders.Add , "C", "检查号"
        .ColumnHeaders.Add , "D", "接收日期"
        .ColumnHeaders.Add , "E", "英文名"
        .ColumnHeaders.Add , "F", "姓别"
        .ColumnHeaders.Add , "G", "序列数"
        .ColumnHeaders.Add , "H", "图像数"
        .ColumnHeaders.Add , "I", "位置一"
        .ColumnHeaders.Add , "J", "位置二"
        .ColumnHeaders.Add , "K", "检查UID"
'        .ColumnHeaders.Add , "L", "处理状态"   '归档时，并没有根据列表中的数据归档，而是查询的数据库
    End With
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With LivMain
        .Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    End With
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
End Sub
Private Sub Form_Load()
    Dim strSQL As String
    Dim tmpset As ADODB.Recordset

    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    Call RestoreWinState(Me, App.EXEName)

    '----------加载托盘图标
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "PACS备份管理程序"
    mobjIcon.AddIcon
    '----------加载托盘图标
    
'    AddToTray Me

    ''''gcnOracle
    '初使化列表控件
    Call InitLiv
    '从注册表获取自动归档策略
    subReadPolicy
    '设置本程序启动的日期
    beginDay = Date
    timAutoPolicy.Enabled = True
    Call ShowChkRecord

'    SetTrayTip "PACS备份管理程序"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timAutoPolicy.Enabled = False
    
'    RemoveFromTray
    '清除托盘图标
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    
    Call SaveWinState(Me, App.EXEName)
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
On Error Resume Next
    '如果更新数据库和显示日志的模式窗口已经被打开，则退出，避免出现错误
'    If mfrmUpdateDB Is Nothing And mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            Me.Hide
        Else
            WindowState = vbNormal
            Me.Show
        End If
'    End If
    Err.Clear
End Sub

Public Sub ShowChkRecord()
    Dim strSQL As String
    Dim tmpset As New ADODB.Recordset
    Dim strStorePlace As String         '使用“位置一”，或“位置二”来筛选
    Dim objItem As ListItem             '列表对像
    '从注册表中取出过滤条件的临时变量
    Dim mDevice As String
    Dim mStorageDevice As String
    Dim mFStudy As String
    Dim mEStudy As String
    Dim mFTime As String
    Dim mETime As String
    Dim mArchiveState As String
    
    mDevice = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "影像类型", "所有类型")
    mStorageDevice = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "储存设备", cAllStorageDevice)
    mFStudy = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "开始检查号", "")
    mEStudy = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "结束检查号", "")
    mFTime = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "开始时间", zlDatabase.Currentdate - 90)
    mETime = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "结束时间", zlDatabase.Currentdate - 30)
    mArchiveState = GetSetting("ZLSOFT", "公共模块\归档管理\过滤", "归档状态", "未归档")
    
    With frmFilter
        strSQL = "Select 影像类别,检查号,接收日期,姓名,英文名,性别,Sum(1) As 序列数,Sum(图像数) As 图像数,位置一,位置二,b.检查UID From" & _
            " (Select a.检查UID,b.序列UID,Sum(1) As 图像数 from 影像检查记录 a,影像检查序列 b,影像检查图象 c" & _
            " Where a.检查UID=b.检查UID And b.序列UID=c.序列UID " & _
            IIf(mStorageDevice = cAllStorageDevice, "", IIf(mArchiveState = "未归档", " And a.位置一=" & mStorageDevice, " And a.位置二=" & mStorageDevice)) & _
            IIf(mArchiveState = "未归档", " And a.位置二 is null ", IIf(mArchiveState = "已归档已删除", " And a.位置一 is null ", " And a.位置一 is not null And a.位置二 is not null")) & _
            IIf(mDevice = "所有类型", "", " And a.影像类别='" & mDevice & "'") & _
            IIf(mFTime = "3000-01-01", "", " And a.接收日期>=to_Date('" & Format(mFTime, "yyyy-MM-dd HH:mm:SS") & "','YYYY-MM-DD HH24:Mi:SS')") & _
            IIf(mETime = "3000-01-01", "", " And a.接收日期<=to_Date('" & Format(mETime, "yyyy-MM-dd HH:mm:SS") & "','YYYY-MM-DD HH24:Mi:SS')") & _
            IIf(Len(Trim(mFStudy)) = 0 Or Not IsNumeric(mFStudy), "", " And a.检查号>=" & mFStudy) & _
            IIf(Len(Trim(mEStudy)) = 0 Or Not IsNumeric(mEStudy), "", " And a.检查号<=" & mEStudy) & _
            " Group By a.检查UID,b.序列UID) a, 影像检查记录 b Where a.检查UID=b.检查UID Group By 影像类别,检查号,接收日期,姓名,英文名,性别,位置一,位置二,b.检查UID"
    End With
    On Error GoTo errH
    zlDatabase.OpenRecordset rsData, strSQL, Me.Caption
    LivMain.ListItems.Clear
    Do Until rsData.EOF
        With LivMain
            Set objItem = .ListItems.Add(, "A" & rsData("检查号") & "UID:" & rsData("检查UID"), ZlCommFun.NVL(rsData("姓名")))
            objItem.SubItems(1) = IIf(IsNull(rsData("检查号")), "", rsData("检查号"))
            objItem.SubItems(2) = rsData("接收日期")
            objItem.SubItems(3) = rsData("影像类别")
            objItem.SubItems(4) = ZlCommFun.NVL(rsData("英文名"), "UnKnow")
            objItem.SubItems(5) = ZlCommFun.NVL(rsData("性别"), "未知")
            objItem.SubItems(6) = rsData("序列数")
            objItem.SubItems(7) = rsData("图像数")
            objItem.SubItems(8) = IIf(IsNull(rsData("位置一")), "", rsData("位置一"))
            objItem.SubItems(9) = IIf(IsNull(rsData("位置二")), "", rsData("位置二"))
            objItem.SubItems(10) = rsData("检查UID")
        End With
        rsData.MoveNext
    Loop
    
    '提示
    Me.stbThis.Panels(2).Text = "当前<" & mArchiveState & ">状态：共有" & rsData.RecordCount & "条记录。"
    
    '屏蔽工具条和菜单按钮
    If Me.LivMain.ListItems.Count > 0 Then
        Select Case mArchiveState
            Case "未归档"
                Me.mnuEditManualArchive.Enabled = True
                Me.mnuEditManualDeArchive.Enabled = False
                Me.tbrMain.Buttons("ManualArchive").Enabled = True
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = False
            Case "已归档已删除"
                Me.mnuEditManualArchive.Enabled = False
                Me.mnuEditManualDeArchive.Enabled = True
                Me.tbrMain.Buttons("ManualArchive").Enabled = False
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = True
            Case "已归档未删除"
                Me.mnuEditManualArchive.Enabled = True
                Me.mnuEditManualDeArchive.Enabled = True
                Me.tbrMain.Buttons("ManualArchive").Enabled = True
                Me.tbrMain.Buttons("ManualDeArchive").Enabled = True
        End Select
    Else
        Me.mnuEditManualArchive.Enabled = False
        Me.mnuEditManualDeArchive.Enabled = False
        Me.tbrMain.Buttons("ManualArchive").Enabled = False
        Me.tbrMain.Buttons("ManualDeArchive").Enabled = False
    End If
        
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function funcdoArchiveJob(Optional lngJobID As Long = 0) As Boolean
    '执行归档作业
    ''''''''''''''''''''''''''''''''''''''
    ''' 参数:lngJobID----需要执行的作业ID,如果等于0，则从作业数据库中检索出来一个作业并执行
    ''''''''''''''''''''''''''''''''''''''
    
    Dim tmpset As ADODB.Recordset
    Dim dsArchive As ADODB.Recordset
    Dim strSQL As String
    Dim strSourceDev As String, strDestinationDev As String, strAppointDev As String
    Dim bMove As Boolean, bDelete As Boolean, bAutoBackup As Boolean
    Dim strFilter As String
    
    funcdoArchiveJob = False
    '针对自动执行作业的情况，读取数据库，获取归档任务
    If lngJobID = 0 Then        '检索影像归档作业表，获取第一个可执行的作业ID
        strSQL = "Select 编码 From 影像归档作业 where 执行过程 = 0 "
        Set tmpset = gcnOracle.Execute(strSQL)
        If Not tmpset.EOF Then
            lngJobID = tmpset!编码
        Else
            Exit Function       '没有可执行的作业，退出函数
        End If
    End If
    '回写数据库，填写开始归档标识和开始时间
    strSQL = "update 影像归档作业 set 开始时间 = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
             "执行过程 = 1 where 编码=" & lngJobID
    gcnOracle.Execute (strSQL)
    
    '获取归档条件
    strSQL = "select 源设备,目的设备,指定设备,是否迁移,是否删除,自动备份,检索条件 from 影像归档作业 where 编码=" & lngJobID
    Set tmpset = gcnOracle.Execute(strSQL)
    strSourceDev = tmpset!源设备
    strDestinationDev = tmpset!目的设备
    If IsNull(tmpset!指定设备) Then
        strAppointDev = vbNullString
    Else
        strAppointDev = tmpset!指定设备
    End If
    bMove = tmpset!是否迁移
    bDelete = tmpset!是否删除
    bAutoBackup = tmpset!自动备份
    If IsNull(tmpset!检索条件) Then
        strFilter = vbNullString
    Else
        strFilter = tmpset!检索条件
    End If
    
    '归档操作
    '获取归档数据记录
    If bAutoBackup Then     '自动归档，需要根据条件进行检索
        strSQL = "Select 接收日期,位置一,位置二,检查UID From 影像检查记录 where not 位置一 is null and 位置二 is null"
        Set dsArchive = gcnOracle.Execute(strSQL)
        '需要对检索条件进行解析
    Else                    '手动归档，直接获取原来检索的记录
        Set dsArchive = rsData
    End If
    
    '调用函数执行归档操作
    funcdoArchiveJob = funcArchiveExec(strSourceDev, strDestinationDev, strAppointDev, bMove, bDelete, dsArchive)
    
    '归档完成，回写数据库，填写归档完成标识和完成时间
    
    If funcdoArchiveJob = True Then
        strSQL = "update 影像归档作业 set 结束时间 = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
                 "执行过程 = 2 where 编码=" & lngJobID
    
    Else            '填写标识，执行任务失败
        strSQL = "update 影像归档作业 set 结束时间 = to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss')," & _
                 "执行过程 = 3 where 编码=" & lngJobID
    End If
    gcnOracle.Execute (strSQL)
End Function

Private Function funcArchiveExec(strSourceDev As String, strDestinationDev As String, strAppointDevID As String, _
                                 bMove As Boolean, bDelete As Boolean, dsArchive As ADODB.Recordset) As Boolean
    Dim cDevice As New Collection       '保存全部归档设备信息的集合
    Dim clsOneDevice As clsBakDevice    '暂存一个设备信息的类
    Dim strSQL As String                '保存需要执行的临时SQL语句
    Dim tmpset As ADODB.Recordset       '保存SQL语句的执行结果数据集
    Dim strTempDir As String            '本机临时目录
    Dim strLocalIP As String            '本机IP地址
    Dim strLocalDirDest As String       '目的设备本机目录
    Dim lngResult As Long               '存储函数执行的返回信息
    Dim strDestDevID As String          '目的设备的ID
    Dim i As Integer                    '通用循环计数器
    
    If dsArchive.RecordCount <= 0 Then      '数据集为空，直接退出函数，标识完成任务
        funcArchiveExec = True
        Exit Function
    End If
    funcArchiveExec = False
    
    '使用一个集合来缓存被使用到的源设备IP，用户名，密码
    strSQL = "select 设备号,设备名,类型,IP地址,FTP目录,ftp用户名,ftp密码,状态,本机目录 from 影像设备目录 WHERE " & _
             "类型 = 1"
    Set tmpset = gcnOracle.Execute(strSQL)
    While Not tmpset.EOF
        With tmpset
            Set clsOneDevice = New clsBakDevice
            clsOneDevice.strDevID = IIf(IsNull(!设备号), "", !设备号)
            clsOneDevice.strDevName = IIf(IsNull(!设备名), "", !设备名)
            clsOneDevice.lngType = IIf(IsNull(!类型), "", !类型)
            clsOneDevice.strIP = IIf(IsNull(!ip地址), "", !ip地址)
            clsOneDevice.strPasswd = IIf(IsNull(!ftp密码), "", !ftp密码)
            clsOneDevice.strUser = IIf(IsNull(!ftp用户名), "", !ftp用户名)
            clsOneDevice.strVirtualPath = IIf(IsNull(!FTP目录), "", !FTP目录)
            clsOneDevice.strLocalPath = IIf(IsNull(!本机目录), "", !本机目录)
            clsOneDevice.lngStatus = IIf(IsNull(!状态), "", !状态)
            cDevice.Add clsOneDevice, clsOneDevice.strDevID
            .MoveNext
        End With
    Wend
    
    '获取本机的临时路径
    strTempDir = Environ("TEMP") & "\zlPacs"
    If Dir(strTempDir, vbDirectory) = vbNullString Then
        MkDir strTempDir
    End If
    '获取本机IP地址
    strLocalIP = Winsock1.LocalIP
    '设定目的设备号
    If strAppointDevID = vbNullString Then '自动选择
        For i = 1 To cDevice.Count
            If cDevice(i).lngStatus = 1 Then
                strDestDevID = cDevice(i).strDevID
                Exit For
            End If
        Next
    Else
        strDestDevID = strAppointDevID
    End If
    '如果目的设备是本机，设定本机目的设备目录
    strLocalDirDest = vbNullString
    If UCase(cDevice(strDestDevID).strIP) = "LOCALHOST" Or cDevice(strDestDevID).strIP = strLocalIP Then
        strLocalDirDest = cDevice(strDestDevID).strLocalPath
    End If
    
    '开始进行归档操作
    dsArchive.MoveFirst
    While Not dsArchive.EOF
        
        ''''调用对一条记录进行归档的实际操作'''''
        lngResult = funcArchiveOneRecord(dsArchive, cDevice, strLocalIP, strTempDir, strLocalDirDest, _
                             strSourceDev, strDestinationDev, strDestDevID, bMove, bDelete)
        Select Case lngResult
        Case 0            '成功完成，向后移动一条记录
            dsArchive.MoveNext
        Case 1, 2            '1--FTP连接失败，转移到下一个设备;2--'FTP迁移失败，标记当前设备满，转到下一个设备
            If lngResult = 2 Then           '标识设备满
                strSQL = "update 影像设备目录 set 目录满 = 1 where 设备号='" & strDestDevID & "'"
                gcnOracle.Execute strSQL
            End If
            '转移到下一个设备
            strDestDevID = vbNullString
            For i = 1 To cDevice.Count
                If cDevice(i).lngStatus = 1 Then '备份设备
                    strDestDevID = cDevice(i).strDevID
                    '如果目的设备是本机，设定本机目的设备目录
                    If UCase(cDevice(strDestDevID).strIP) = "LOCALHOST" Or cDevice(strDestDevID).strIP = strLocalIP Then
                        strLocalDirDest = cDevice(strDestDevID).strLocalPath
                    Else
                        strLocalDirDest = vbNullString
                    End If
                    Exit For
                End If
            Next
            If strDestDevID = vbNullString Then         '无设备可用，提示并退出
                If lngResult = 1 Then
                    MsgBox "没有可以连接的设备，可能是网络连接故障，请检查设备情况。"
                Else
                    MsgBox "没有可以连接的设备，可能是设备存储容量满，请检查设备情况。"
                End If
                Exit Function
            End If
        Case 3             'FTP删除失败，记录到日志中，向后移动一条记录
            '''''''''记录到日志中'''''''''''''''
            dsArchive.MoveNext
        Case 4             '未知错误，直接提示用户，中断操作
            MsgBox "出现异常错误，归档操作中止,可能是本机磁盘容量不足，请检查配置后重试。"
            Exit Function
        Case 5
            dsArchive.MoveNext
        End Select
    Wend
    funcArchiveExec = True
End Function
              
Private Function funcArchiveOneRecord(dsArchive As ADODB.Recordset, cDevice As Collection, _
                 strLocalIP As String, strTempDir As String, strLocalDirDest As String, _
                 strSourceDev As String, strDestinationDev As String, strDestDevID As String, _
                 bMove As Boolean, bDelete As Boolean) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''
'''功能：对一条记录执行归档（迁移和删除）操作
'''返回：成功返回0，连接失败返回1，迁移失败返回2，删除失败返回3，其他返回4
''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer                    '计数器
    Dim iImgCount As Integer            '图像计数器
    Dim objFileSystem As Object         '利用其中的文件复制函数
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Dim aImageFiles() As String         '图像文件名数组
    Dim strImgPath As String            '图像中间路径，由 “采集日期\检查UID”组成
    Dim strFTPImgPath As String         'FTP使用的图像中间路径，由 “采集日期/检查UID”组成
    Dim clsFTPsubs As New clsFTP        '为了调用clsFTP类里面的函数
    Dim strSQL As String                '保存需要执行的临时SQL语句
    Dim tmpset As ADODB.Recordset       '保存SQL语句的执行结果数据集
    Dim strLocalDirSource As String     '源设备本机目录
    Dim lngResult As Long               '保存返回值
    Dim strSourceDevID As String        '源设备的ID
    Dim lngFilesCount As Long           '所有需要归档或移动的文件的数目
    Dim aRptImgFiles() As String        '报告图像文件名数组
    Dim aOtherFiles() As String         '其他零碎文件名数组，包括报告图像，录音报告等
    
    funcArchiveOneRecord = 1
    '判断源设备中是否有本地设备，通过比较IP地址实现
    strSourceDevID = IIf(strSourceDev = "1", IIf(IsNull(dsArchive!位置一), "", dsArchive!位置一), IIf(IsNull(dsArchive!位置二), "", dsArchive!位置二))
    If UCase(cDevice(strSourceDevID).strIP) = "LOCALHOST" Or cDevice(strSourceDevID).strIP = strLocalIP Then
        '如果有本地设备，则直接使用本地目录
        strLocalDirSource = cDevice(strSourceDevID).strLocalPath
    Else
        strLocalDirSource = vbNullString
    End If
    
    '组合完成图像路径
    strFTPImgPath = Format(dsArchive!接收日期, "yyyymmdd") & "/" & dsArchive!检查uid
    strImgPath = Format(dsArchive!接收日期, "yyyymmdd") & "\" & dsArchive!检查uid
    
    '查询数据库，获取图像文件目录和文件名
    strSQL = "select 图像UID from 影像检查图象 a ,影像检查序列 b where b.序列UID = a.序列UID and b.检查UID = '" & _
             dsArchive!检查uid & "'"
    Set tmpset = gcnOracle.Execute(strSQL)
    ReDim aImageFiles(tmpset.RecordCount) As String
    i = 1
    While Not tmpset.EOF
        aImageFiles(i) = tmpset!图像uid
        i = i + 1
        tmpset.MoveNext
    Wend
    
    '查询数据库，获取报告图像文件名,录音报告文件名等零碎文件名
    strSQL = "select 报告图象 from 影像检查记录 where 检查UID = '" & dsArchive!检查uid & "'"
    Set tmpset = gcnOracle.Execute(strSQL)
    If Not IsNull(tmpset!报告图象) Then
        aOtherFiles = Split(tmpset!报告图象, "|")   '区分出报告图像和录音报告
        If UBound(aOtherFiles) > -1 Then            '有报告图像或录音报告，对其文件名添加到图像文件名数组中
            For i = 0 To UBound(aOtherFiles)
                aRptImgFiles = Split(aOtherFiles(i), ";")
                lngFilesCount = UBound(aImageFiles)
                If UBound(aRptImgFiles) > -1 Then   '有报告图像或录音报告
                    ReDim Preserve aImageFiles(lngFilesCount + UBound(aRptImgFiles) + 1) As String
                    For iImgCount = 0 To UBound(aRptImgFiles)
                        aImageFiles(iImgCount + 1 + lngFilesCount) = Trim(aRptImgFiles(iImgCount))
                    Next
                End If
            Next
        End If
    End If
    
    If bMove Then       '进行归档操作
        If strLocalDirSource <> vbNullString And strLocalDirDest <> vbNullString Then   '本机的文件复制
            '在目的设备创建目录
            If Dir(strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1), vbDirectory) = vbNullString Then
                MkDir strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            End If
            If Dir(strLocalDirDest & "\" & strImgPath, vbDirectory) = vbNullString Then
                MkDir strLocalDirDest & "\" & strImgPath
            End If
            For i = 1 To UBound(aImageFiles)
                objFileSystem.CopyFile strLocalDirSource & "\" & strImgPath & "\" & aImageFiles(i), _
                                       strLocalDirDest & "\" & strImgPath & "\"
            Next
            '复制文件
        ElseIf strLocalDirSource <> vbNullString Then       '从本机目录直接上传
            clsFTPsubs.strIPAddress = cDevice(strDestDevID).strIP
            clsFTPsubs.strPsw = cDevice(strDestDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strDestDevID).strUser
            lngResult = clsFTPsubs.FuncFtpMkDir(cDevice(strDestDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncUpLoadFiles(cDevice(strDestDevID).strVirtualPath & "/" & strFTPImgPath, strLocalDirSource & "\" & strImgPath, aImageFiles)
            If lngResult <> 0 Then
                funcArchiveOneRecord = lngResult
                Exit Function
            End If
        ElseIf strLocalDirDest <> vbNullString Then         '直接下载到本机
            '在目的设备创建目录
            MkDir strLocalDirDest & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            MkDir strLocalDirDest & "\" & strImgPath
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            lngResult = clsFTPsubs.FuncDownLoadFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, strLocalDirDest & "\" & strImgPath, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 4        '返回其他错误，可能是本机磁盘容量不足
                End If
                Exit Function
            End If
        Else        '使用临时目录做中转
            '如果没有本地设备，则使用系统临时目录作为中转
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            lngResult = clsFTPsubs.FuncDownLoadFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, strTempDir, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 5        '返回其他错误，可能未能读取源文件
                End If
                Exit Function
            End If
            clsFTPsubs.strIPAddress = cDevice(strDestDevID).strIP
            clsFTPsubs.strPsw = cDevice(strDestDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strDestDevID).strUser
            lngResult = clsFTPsubs.FuncFtpMkDir(cDevice(strDestDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncUpLoadFiles(cDevice(strDestDevID).strVirtualPath & "/" & strFTPImgPath, strTempDir, aImageFiles)
            If lngResult <> 0 Then
                funcArchiveOneRecord = lngResult
                Exit Function
            End If
            For i = 1 To UBound(aImageFiles) '删除中转文件
                objFileSystem.DeleteFile strTempDir & "\" & aImageFiles(i)
            Next
        End If
    End If
    
    '进行删除操作
    If bDelete Then
        If strLocalDirSource <> vbNullString Then       '删除本机文件和目录
            For i = 1 To UBound(aImageFiles)        '删除文件
                objFileSystem.DeleteFile strLocalDirSource & "\" & strImgPath & "\" & aImageFiles(i)
            Next
            '删除目录
            '检查目录是否为空,后删除
            If (Dir(strLocalDirSource & "\" & strImgPath & "\*.*") = vbNullString) Then
                objFileSystem.DeleteFolder strLocalDirSource & "\" & strImgPath
            End If
            If (Dir(strLocalDirSource & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1) & "\*.*") = vbNullString) Then
                objFileSystem.DeleteFolder strLocalDirSource & "\" & Left(strImgPath, InStr(strImgPath, "\") - 1)
            End If
        Else            '删除FTP文件和目录
            clsFTPsubs.strIPAddress = cDevice(strSourceDevID).strIP
            clsFTPsubs.strPsw = cDevice(strSourceDevID).strPasswd
            clsFTPsubs.strUser = cDevice(strSourceDevID).strUser
            '删除文件
            lngResult = clsFTPsubs.FuncDelFiles(cDevice(strSourceDevID).strVirtualPath & "/" & strFTPImgPath, aImageFiles)
            If lngResult <> 0 Then
                If lngResult = 1 Then
                    funcArchiveOneRecord = lngResult
                Else
                    funcArchiveOneRecord = 3        '返回删除失败
                End If
                Exit Function
            End If
            '删除目录
            lngResult = clsFTPsubs.FuncFtpDelDir(cDevice(strSourceDevID).strVirtualPath, strFTPImgPath)
            lngResult = clsFTPsubs.FuncFtpDelDir(cDevice(strSourceDevID).strVirtualPath, Left(strFTPImgPath, InStr(strFTPImgPath, "/") - 1))
        End If
    End If
    '向影像检查记录表中填写归档结果
    If bMove = True And bDelete = True Then
        strSQL = IIf(strDestinationDev = "2", "位置二 = '" & strDestDevID & "'", "位置一 = '" & strDestDevID & "'") & _
                 IIf(strSourceDev = "2", " , 位置二 =null", " , 位置一 = null")
    ElseIf bMove = True Then
        strSQL = IIf(strDestinationDev = "2", "位置二 = '" & strDestDevID & "'", "位置一 = '" & strDestDevID & "'")
    ElseIf bDelete = True Then
        strSQL = IIf(strSourceDev = "2", " 位置二 =null", " 位置一 = null")
    End If
    strSQL = "update 影像检查记录 set  " & strSQL & " where 检查UID = '" & dsArchive!检查uid & "'"
    gcnOracle.Execute (strSQL)
    funcArchiveOneRecord = 0        '操作正常结束，返回
End Function

Private Sub subManualArchive()
    '手动归档操作
    Dim tmpset As ADODB.Recordset
    Dim strSQL As String
    frmManualArchive.Caption = "手动归档"
    '填充归档设备
    frmManualArchive.cobDevice.Clear
    strSQL = "Select 设备名,设备号 From 影像设备目录 where 类型 = 1 and 状态 = 1 and NVL(目录满,0) = 0"
    Set tmpset = gcnOracle.Execute(strSQL)
    With tmpset
        While Not .EOF
            frmManualArchive.cobDevice.AddItem !设备号 & "-" & !设备名
            .MoveNext
        Wend
    End With
    frmManualArchive.cobDevice.ListIndex = IIf(frmManualArchive.cobDevice.ListCount > 0, 0, -1)
    frmManualArchive.cobManualMoveDelete.ListIndex = 0
    frmManualArchive.sstabManualArchive.Tab = 0
    frmManualArchive.cmdStep3.Caption = "开始归档"
    frmManualArchive.bArchive = True
    frmManualArchive.Show 1, Me
End Sub

Private Sub subManualDeArchive()
    
    '手反动归档操作
    Dim tmpset As ADODB.Recordset
    Dim strSQL As String
    frmManualArchive.Caption = "手动反归档"
    '填反充归档设备
    frmManualArchive.cobDevice.Clear
    strSQL = "Select 设备名,设备号 From 影像设备目录 where 类型 = 1 and NVL(目录满,0) = 0"
    Set tmpset = gcnOracle.Execute(strSQL)
    With tmpset
        While Not .EOF
            frmManualArchive.cobDevice.AddItem !设备号 & "-" & !设备名
            'frmManualArchive.cobDevice.ItemData(frmManualArchive.cobDevice.NewIndex) = !设备号
            .MoveNext
        Wend
    End With
    frmManualArchive.cobDevice.ListIndex = IIf(frmManualArchive.cobDevice.ListCount > 0, 0, -1)
    frmManualArchive.cobManualMoveDelete.ListIndex = 0
    frmManualArchive.sstabManualArchive.Tab = 0
    frmManualArchive.cmdStep3.Caption = "开始反归档"
    frmManualArchive.bArchive = False   '标识反归档
    frmManualArchive.Show 1, Me
End Sub

Private Sub mnuEditAutoArchiveSetup_Click()
    Call subAutoArchive
End Sub

Private Sub mnuEditCollate_Click()
    frmCollate.Show vbModal, Me
End Sub

Private Sub mnuEditFilter_Click()
    If frmFilter.ShowMe(Me) = True Then
        Call ShowChkRecord
    End If
End Sub

Private Sub mnuEditManualArchive_Click()
    Call subManualArchive
End Sub

Private Sub mnuEditManualDeArchive_Click()
    Call subManualDeArchive
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileSet_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
'    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    LivMain.View = Index
    
End Sub

Private Sub mnuViewReflash_Click()
    ShowChkRecord
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = tbrMain.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrMain.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = tbrMain.Height
    Form_Resize
End Sub

Private Sub timAutoPolicy_Timer()
    Dim aTempPolicy() As String         '保存解析后策略的数组
    Dim strSQL As String                '暂存SQL语句
    Dim tmpset As ADODB.Recordset       '暂存查询数据集
    Dim lngJobNum As Long               '新的归档作业号
    
    '检测时间
    If strTimePolicy = vbNullString Then Exit Sub
    
    aTempPolicy = Split(strTimePolicy, ",")
    If aTempPolicy(1) <> "N/A" Then         '有时间策略，处理时间策略
        '判断是否满足归档条件
        If UCase(aTempPolicy(1)) = "DAY" Then      '这是每天策略，检测间隔天数和时间
            If (Date - beginDay) >= aTempPolicy(2) Then     '检测当前时间
                If Time = CDate(aTempPolicy(3)) Then
                    '添加一条备份记录
                    strSQL = "select 影像归档作业_ID.nextval as JobID from dual"
                    Set tmpset = gcnOracle.Execute(strSQL)
                    lngJobNum = tmpset!JobID
                    strSQL = "Insert into 影像归档作业 (编码,名称,执行时间,源设备,目的设备,指定设备,是否迁移,是否删除,自动备份,执行过程) values (" & _
                             lngJobNum & ",'自动" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                             ",'1','2',''," & aTempPolicy(5) & "," & aTempPolicy(4) & ",1,0)"
                    gcnOracle.Execute (strSQL)
        
                    '通知备份程序    '执行归档作业
                    frmMain.funcdoArchiveJob lngJobNum
                    '修改beginDay为今天
                    beginDay = Date
                End If
            End If
        ElseIf UCase(aTempPolicy(1)) = "MONTH" Then     '这是每月策略，检测今天是否归档日期
            If Day(Date) = aTempPolicy(2) Then     '这是归档日期
                If Time = CDate(aTempPolicy(3)) Then    '这是归档时间
                    '添加一条备份记录
                    strSQL = "select 影像归档作业_ID.nextval as JobID from dual"
                    Set tmpset = gcnOracle.Execute(strSQL)
                    lngJobNum = tmpset!JobID
                    strSQL = "Insert into 影像归档作业 (编码,名称,执行时间,源设备,目的设备,指定设备,是否迁移,是否删除,自动备份,执行过程) values (" & _
                             lngJobNum & ",'自动" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                             ",'1','2',''," & aTempPolicy(5) & "," & aTempPolicy(4) & ",1,0)"
                    gcnOracle.Execute (strSQL)
        
                    '通知备份程序    '执行归档作业
                    frmMain.funcdoArchiveJob lngJobNum
                End If
            End If
        End If
    End If
End Sub

Public Sub subReadPolicy()
    strTimePolicy = GetSetting("ZLSOFT", "公共模块\归档管理", "时间归档策略")
    bAutoArchive = IIf(GetSetting("ZLSOFT", "公共模块\归档管理", "使用自动归档", "False") = "True", True, False)
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ManualArchive"
            '手动归档
            mnuEditManualArchive_Click
        Case "ManualDeArchive"
            '手动反归档
            mnuEditManualDeArchive_Click
        Case "AutoArchiveSetup"
            '自动归档
            mnuEditAutoArchiveSetup_Click
        Case "Filtrate"
            '过滤
            mnuEditFilter_Click
        Case "Preview"
            '预览
            subPrint 2
        Case "Print"
            '打印
            subPrint 1
        Case "View"
            '查看
            mnuViewIcon(LivMain.View).Checked = False
            If LivMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                LivMain.View = 0
            Else
                mnuViewIcon(LivMain.View + 1).Checked = True
                LivMain.View = LivMain.View + 1
            End If
        Case "Help"
            '帮助
            mnuHelpTopic_Click
        Case "Quit"
            '退出
            mnufileexit_Click
    End Select
End Sub
Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "PACS备份管理"
    Set objPrint.Body.objData = LivMain
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
Private Sub subAutoArchive()
    '显示自动策略设置界面
    frmAutoArchive.Show 1, Me
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    LivMain.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mobjIcon.MouseState X
End Sub
