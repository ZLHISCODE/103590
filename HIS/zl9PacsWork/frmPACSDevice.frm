VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPACSDevice 
   AutoRedraw      =   -1  'True
   Caption         =   "影像设备目录"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   Icon            =   "frmPACSDevice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   7305
      Left            =   75
      ScaleHeight     =   7275
      ScaleWidth      =   11175
      TabIndex        =   26
      Top             =   570
      Width           =   11205
      Begin MSComctlLib.ImageList imgKind 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":058A
               Key             =   "Server"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":06E4
               Key             =   "Gate"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":0C7E
               Key             =   "Printer"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":1218
               Key             =   "影像设备"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":1532
               Key             =   "远程目录"
            EndProperty
         EndProperty
      End
      Begin VB.Frame FraInfor 
         Height          =   825
         Left            =   7455
         TabIndex        =   31
         Top             =   6195
         Width           =   3705
      End
      Begin VB.Frame FraDevice 
         Height          =   6165
         Left            =   7440
         TabIndex        =   27
         Top             =   0
         Width           =   3705
         Begin VB.TextBox txtSDPassword 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   37
            Top             =   3660
            Width           =   2060
         End
         Begin VB.TextBox txtSDUser 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   35
            Top             =   3315
            Width           =   2060
         End
         Begin VB.TextBox txtShareDir 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   33
            Top             =   2970
            Width           =   2060
         End
         Begin VB.ComboBox Cbosort 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmPACSDevice.frx":3CE4
            Left            =   1470
            List            =   "frmPACSDevice.frx":3CE6
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   2060
         End
         Begin VB.TextBox TxtName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1140
            Width           =   2060
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "取消(&C)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2475
            TabIndex        =   32
            Top             =   5655
            Width           =   1000
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "连接测试(&T)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   225
            TabIndex        =   24
            Top             =   5655
            Width           =   1140
         End
         Begin VB.ComboBox cboType 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmPACSDevice.frx":3CE8
            Left            =   1470
            List            =   "frmPACSDevice.frx":3CFB
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   810
            Width           =   2060
         End
         Begin VB.TextBox txtDevAdress 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1470
            Width           =   2060
         End
         Begin VB.TextBox txtDevPort 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   5
            TabIndex        =   11
            Top             =   4830
            Width           =   2060
         End
         Begin VB.TextBox txtDevNO 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   3
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   150
            Width           =   2060
         End
         Begin VB.TextBox txtDevLocalAE 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   17
            Top             =   4125
            Width           =   2060
         End
         Begin VB.TextBox txtDevAE 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   19
            ToolTipText     =   "等幅度幅"
            Top             =   4470
            Width           =   2060
         End
         Begin VB.TextBox txtFtpPath 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   9
            Top             =   1935
            Width           =   2060
         End
         Begin VB.TextBox txtUser 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2280
            Width           =   2060
         End
         Begin VB.TextBox txtPassWord 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   2625
            Width           =   2060
         End
         Begin VB.CommandButton cmdPath 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3225
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   5205
            Width           =   300
         End
         Begin VB.CommandButton CmdDevSave 
            Caption         =   "保存(&S)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   1425
            TabIndex        =   23
            Top             =   5655
            Width           =   1000
         End
         Begin VB.TextBox txtDirPath 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   21
            ToolTipText     =   "Ftp目录在服务器上的本地路径"
            Top             =   5175
            Width           =   2060
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   200
            X2              =   3500
            Y1              =   4040
            Y2              =   4040
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   200
            X2              =   3500
            Y1              =   1850
            Y2              =   1850
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "共享目录密码"
            Height          =   180
            Left            =   300
            TabIndex        =   38
            ToolTipText     =   "【存储设备参数】连接共享FTP目录的密码。"
            Top             =   3720
            Width           =   1080
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "共享目录用户名"
            Height          =   180
            Left            =   60
            TabIndex        =   36
            ToolTipText     =   "【存储设备参数】连接共享FTP目录的用户名。"
            Top             =   3375
            Width           =   1260
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "共享目录"
            Height          =   180
            Left            =   600
            TabIndex        =   34
            ToolTipText     =   "【存储设备参数】""FTP目录""的只读共享目录名称，优先使用共享目录方式下载图像。"
            Top             =   3015
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "影像类别(&I)"
            Height          =   180
            Left            =   300
            TabIndex        =   0
            ToolTipText     =   "设置设备对应的影像类别。"
            Top             =   540
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "名称(&M)"
            Height          =   180
            Left            =   675
            TabIndex        =   4
            ToolTipText     =   "设置设备的名称。"
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblRoom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型(&T)"
            Height          =   180
            Left            =   675
            TabIndex        =   2
            ToolTipText     =   "设置设备的职能类型。"
            Top             =   870
            Width           =   630
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "IP地址(&A)"
            Height          =   180
            Left            =   495
            TabIndex        =   6
            ToolTipText     =   "设置设备的网络IP地址。"
            Top             =   1530
            Width           =   810
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "端口(&P)"
            Height          =   180
            Left            =   675
            TabIndex        =   10
            ToolTipText     =   $"frmPACSDevice.frx":3D31
            Top             =   4860
            Width           =   630
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "设备号  "
            Height          =   255
            Left            =   330
            TabIndex        =   29
            ToolTipText     =   "设备的唯一标识，只读。"
            Top             =   173
            Width           =   975
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "本地AE(&L)"
            Height          =   180
            Left            =   495
            TabIndex        =   16
            ToolTipText     =   $"frmPACSDevice.frx":3DB2
            Top             =   4185
            Width           =   810
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "设备AE(&D)"
            Height          =   180
            Left            =   495
            TabIndex        =   18
            ToolTipText     =   $"frmPACSDevice.frx":3E77
            Top             =   4530
            Width           =   810
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ftp目录(&F)"
            Height          =   180
            Left            =   405
            TabIndex        =   8
            ToolTipText     =   "【存储设备参数】存放影像的FTP目录名称。"
            Top             =   1995
            Width           =   900
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FTP用户名(&N)"
            Height          =   180
            Left            =   225
            TabIndex        =   12
            ToolTipText     =   "【存储设备参数】连接FTP目录的用户名。"
            Top             =   2310
            Width           =   1080
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FTP密码(&W)"
            Height          =   180
            Left            =   405
            TabIndex        =   14
            ToolTipText     =   "【存储设备参数】连接FTP目录的密码。"
            Top             =   2655
            Width           =   900
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "本地路径(&L)"
            Height          =   180
            Left            =   315
            TabIndex        =   20
            ToolTipText     =   "设置远程目录本地路径。"
            Top             =   5235
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   6855
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgKind"
         SmallIcons      =   "imgKind"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   7875
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPACSDevice.frx":3F2D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12965
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   255
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPACSDevice.frx":47C1
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPACSDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '**********************************API调用*****************************************

Private mstrPrivs As String

'***********************************************************************************
Private blnBeginchange As Boolean   '开始变更，含新增和修改

Private Sub InitSubWindow()
Dim Pane1 As Pane
    With dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(0, 0, 0, DockTopOf, Nothing)
    Pane1.Title = "设备列表"
    Pane1.Handle = PicList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable

End Sub
Private Sub InitCbosort()
Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "影像检查类别")
    Cbosort.Clear
        Cbosort.AddItem ""
    Do Until rsTemp.EOF
        Cbosort.AddItem rsTemp!编码 & "-" & rsTemp!名称
        rsTemp.MoveNext
    Loop
End Sub
Private Sub InitlvwItem()
    lvwItem.ColumnHeaders.Clear
    With lvwItem.ColumnHeaders
        .Clear
        .Add , "_设备号", "设备号", 800
        .Add , "_影像类别", "影像类别", 900
        .Add , "_名称", "名称", 1500
        .Add , "_类型", "类型", 900
        .Add , "_IP地址", "IP地址", 1500
        .Add , "_端口号", "端口号", 800
        .Add , "_Ftp目录", "Ftp目录", 900
        .Add , "_FTP用户名", "FTP用户名", 1200
        .Add , "_共享目录", "共享目录", 1200
        .Add , "_共享目录用户名", "共享目录用户名", 1600
        .Add , "_本地路径", "本地路径", 900
        .Add , "_本地AE", "本地AE", 1200
        .Add , "_设备AE", "设备AE", 1200
        .Add , "_状态", "状态", 800
        .Add , "_共享目录密码", "共享目录密码", 0
    End With
    With lvwItem
        .SortKey = .ColumnHeaders("_设备号").Index - 1
        .SortOrder = lvwAscending
    End With
    lvwItem.ListItems.Add , , , , 1
    lvwItem.ListItems.Clear
    Call FillData '写数据
End Sub
Private Sub FillData()
Dim strCurrKey As String, objItem As ListItem, rsTemp As New ADODB.Recordset
    If Not lvwItem.SelectedItem Is Nothing Then strCurrKey = lvwItem.SelectedItem.Key
    gstrSQL = "Select A.设备号,B.名称 影像类别,A.设备名,Decode(Nvl(A.类型,1),1,'存储设备',2,'接收主机',3,'胶片打印',4,'影像设备',5,'远程目录') As 设备类型," & _
        "Nvl(A.类型,1) As 类型,A.IP地址,A.端口号,A.Ftp目录,A.FTP用户名,A.FTP密码,A.本地AE,A.设备AE,A.本机目录,A.状态, " & _
        "A.共享目录用户名,A.共享目录密码,A.共享目录" & _
        " From 影像设备目录 A,影像检查类别 B WHERE A.影像类别=B.编码(+) order by 设备号"

    err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取数据")
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !设备号, !设备号, Val(!类型), Val(!类型))
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_影像类别").Index - 1) = Nvl(!影像类别)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = !设备名
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = !设备类型
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_IP地址").Index - 1) = Nvl(!IP地址)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_端口号").Index - 1) = Nvl(!端口号)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ftp目录").Index - 1) = Nvl(!ftp目录)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_FTP用户名").Index - 1) = Nvl(!FTP用户名)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_共享目录").Index - 1) = Nvl(!共享目录)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_共享目录用户名").Index - 1) = Nvl(!共享目录用户名)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_本地路径").Index - 1) = Nvl(!本机目录)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_本地AE").Index - 1) = Nvl(!本地AE)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_设备AE").Index - 1) = Nvl(!设备AE)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_状态").Index - 1) = Decode(Nvl(!状态), 1, "在用", "已停用")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_共享目录密码").Index - 1) = Nvl(!共享目录密码)
            objItem.tag = Nvl(!FTP密码)
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.Count > 0 Then
        err = 0: On Error Resume Next
        lvwItem.ListItems(strCurrKey).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
        '默认选中第一行记录
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False


'菜单定义
'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)") '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)") '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…") '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With


'Begin----------------------编辑菜单--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "服务配置(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用设备(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用设备(&R)")
    End With
    
'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '二级菜单
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With


'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题", -1, False)
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "WEB上的中联(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With


'----------------------快键绑定------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, VK_F1, conMenu_Help_Help              '帮助-------------F1
        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem   '新增-------------CTRL+N
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify    '修改-------------CTRL+M
        .Add FCONTROL, vbKeyD, conMenu_Edit_Delete    '删除-------------CTRL+D
        .Add 0, VK_F5, conMenu_View_Refresh           '刷新-------------F5
        .Add FCONTROL, vbKeyP, conMenu_File_Parameter '服务配置
        .Add 0, VK_F9, conMenu_Edit_Stop              '停用-------------F9
        .Add 0, VK_F10, conMenu_Edit_Reuse            '启用-------------F10
    End With


'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.Style = xtpButtonIconAndCaption '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.Style = xtpButtonIconAndCaption '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "服务配置"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用"): cbrControl.BeginGroup = True:  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.Style = xtpButtonIconAndCaption '固有
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.Style = xtpButtonIconAndCaption  '固有
    End With
End Sub

Private Sub cboType_Click()
    Call ControlEnabled(cboType.list(cboType.ListIndex))
End Sub
Private Sub ControlEnabled(ByVal Typekey As String)
'控件可用状态
        Select Case Typekey
        Case "存储设备"
            Cbosort.Enabled = False: Cbosort.ListIndex = -1
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = False: txtDevAE.BackColor = &H80000000
            txtDevLocalAE.Enabled = False: txtDevLocalAE.BackColor = &H80000000
            txtPassWord.Enabled = True: txtPassWord.BackColor = &H80000005
            txtUser.Enabled = True: txtUser.BackColor = &H80000005
            txtDevPort.Enabled = False: txtDevPort.BackColor = &H80000000
            txtFtpPath.Enabled = True: txtFtpPath.BackColor = &H80000005
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = True: txtShareDir.BackColor = &H80000005
            txtSDUser.Enabled = True: txtSDUser.BackColor = &H80000005
            txtSDPassword.Enabled = True: txtSDPassword.BackColor = &H80000005
        Case "接收主机", "胶片打印"
            Cbosort.Enabled = True
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = True: txtDevAE.BackColor = &H80000005
            txtDevLocalAE.Enabled = True: txtDevLocalAE.BackColor = &H80000005
            txtDevLocalAE.ToolTipText = ""
            txtPassWord.Enabled = False: txtPassWord.BackColor = &H80000000
            txtUser.Enabled = False: txtUser.BackColor = &H80000000
            txtDevPort.Enabled = True: txtDevPort.BackColor = &H80000005
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
        Case "影像设备"
            Cbosort.Enabled = True
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = True: txtDevAE.BackColor = &H80000005
            txtDevLocalAE.Enabled = True: txtDevLocalAE.BackColor = &H80000005
            txtDevLocalAE.ToolTipText = "用于Q/R查询的本地服务AE"
            txtPassWord.Enabled = False: txtPassWord.BackColor = &H80000000
            txtUser.Enabled = False: txtUser.BackColor = &H80000000
            txtDevPort.Enabled = True: txtDevPort.BackColor = &H80000005
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
        Case "远程目录"
            Cbosort.Enabled = True: Cbosort.ListIndex = -1
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = True
            txtDirPath.Enabled = True: txtDirPath.BackColor = &H80000005
            txtDevAE.Enabled = False: txtDevAE.BackColor = &H80000000
            txtDevLocalAE.Enabled = False: txtDevLocalAE.BackColor = &H80000000
            txtPassWord.Enabled = True: txtPassWord.BackColor = &H80000005
            txtUser.Enabled = True: txtUser.BackColor = &H80000005
            txtDevPort.Enabled = False: txtDevPort.BackColor = &H80000000
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = False: txtDevAdress.BackColor = &H80000000
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
    End Select
End Sub
Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
'---------------------------文件----------------
        Case conMenu_File_Exit      '退出
            Unload Me
        Case conMenu_File_PrintSet, conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Call Menu_File_PrintMore(control.ID)
        Case conMenu_Edit_NewItem '新增
            Call Menu_Edit_NewItem
        Case conMenu_Edit_Modify '修改
            Call Menu_Edit_Modify
        Case conMenu_Edit_Delete '删除
            Call Menu_Edit_Delete
        Case conMenu_Edit_Stop   '停用
            Call Menu_Edit_Stop
        Case conMenu_Edit_Reuse  '启用
            Call Menu_Edit_Reuse
        Case conMenu_File_Parameter '服务配置
            Call Menu_File_Parameter
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Refresh '刷新
            Call InitlvwItem
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
End Sub
Private Sub Menu_Edit_Delete()
On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBoxD(Me, "真的将“" & Me.lvwItem.SelectedItem.SubItems(2) & "”从影像设备目录中删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "zl_影像设备目录_Delete('" & Mid(Me.lvwItem.SelectedItem.Key, 2) & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Call Me.lvwItem.ListItems.Remove(Me.lvwItem.SelectedItem.Key)
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog

End Sub
Private Sub Menu_Edit_Stop()
'停用设备
    On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwItem.SelectedItem
        If MsgBoxD(Me, "真的将“" & .SubItems(2) & "“ 停用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_影像设备目录_Update('" & .Text & "','" & .SubItems(lvwItem.ColumnHeaders("_名称").Index - 1) & _
                    "'," & Decode(.SubItems(lvwItem.ColumnHeaders("_类型").Index - 1), "存储设备", 1, "接收主机", 2, "胶片打印", 3, "影像设备", 4, "远程目录", 5, 6) & ",'" & .SubItems(lvwItem.ColumnHeaders("_IP地址").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_Ftp目录").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_端口号").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FTP用户名").Index - 1) & "','" & Trim(.tag) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录用户名").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录密码").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_本地AE").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_设备AE").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_本地路径").Index - 1) & "', 0)"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    Call InitlvwItem
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub
Private Sub Menu_Edit_Reuse()
'启用设备
On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwItem.SelectedItem
        If MsgBoxD(Me, "真的将“" & .SubItems(2) & "”启用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_影像设备目录_Update('" & .Text & "','" & .SubItems(lvwItem.ColumnHeaders("_名称").Index - 1) & _
                    "'," & Decode(.SubItems(lvwItem.ColumnHeaders("_类型").Index - 1), "存储设备", 1, "接收主机", 2, "胶片打印", 3, "影像设备", 4, "远程目录", 5, 6) & ",'" & .SubItems(lvwItem.ColumnHeaders("_IP地址").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_Ftp目录").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_端口号").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FTP用户名").Index - 1) & "','" & Trim(.tag) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录用户名").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_共享目录密码").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_本地AE").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_设备AE").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_本地路径").Index - 1) & "', 1)"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    
    Call InitlvwItem
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub
Private Sub Menu_File_Parameter()
    Call frmPacsSrvSet.ShowMe(Mid(lvwItem.SelectedItem.Key, 2), _
                lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_名称").Index - 1), _
                lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_IP地址").Index - 1), Me)
End Sub
Private Sub Menu_Edit_NewItem()
    blnBeginchange = True
    txtDevNO = ""
    cboType.ListIndex = -1
    Cbosort.ListIndex = -1
    txtName = ""
    txtDirPath = ""
    txtDevAE = ""
    txtDevLocalAE = ""
    txtPassWord = ""
    txtUser = ""
    txtDevPort = ""
    txtFtpPath = ""
    txtDevAdress = ""

    CmdDevSave.Enabled = True
    cmdCancel.Enabled = True
    cboType.Enabled = True
    Cbosort.Enabled = True
    txtName.Enabled = True
    txtDevNO = GetNewNo
    If Cbosort.ListCount <= 0 Then Call InitCbosort
    txtName.SetFocus
End Sub
Private Sub Menu_Edit_Modify()
    blnBeginchange = True
    CmdDevSave.Enabled = True
    cmdCancel.Enabled = True
    Call ControlEnabled(lvwItem.SelectedItem.SubItems(2))
End Sub
Private Sub Menu_File_PrintMore(ByVal lngType As Long)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    
Dim objPrint As New zlPrintLvw, bytType As Byte
    
    On Error Resume Next
    If lvwItem.ListItems.Count <= 0 Then Exit Sub
    
    objPrint.Title.Text = "设备列表"
    Set objPrint.Body.objData = lvwItem
    objPrint.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")

    Select Case lngType
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Preview
            zlPrintOrViewLvw objPrint, 2
        Case conMenu_File_Print
            bytType = zlPrintAsk(objPrint)
            If bytType <> 0 Then zlPrintOrViewLvw objPrint, bytType
        Case conMenu_File_Excel
            zlPrintOrViewLvw objPrint, 3
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_Edit_NewItem '新增 没权限和增删改、配置任一功能开始后其它按键都不可用
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Modify '修改
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Delete '删除
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_File_Parameter '服务配置
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_类型").Index - 1) <> "影像设备" Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Stop      '停用
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_状态").Index - 1) <> "在用" Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_Edit_Reuse     '启用
            If Not CheckPopedom(mstrPrivs, "增删改") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_状态").Index - 1) <> "已停用" Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_File_Excel '导出EXCEL
            If Not CheckPopedom(mstrPrivs, "增删改") Then control.Enabled = False
    End Select
End Sub

Private Sub cmdCancel_Click()
    If MsgBoxD(Me, "当前变更尚未保存,确实要取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then txtName.SetFocus: Exit Sub
    blnBeginchange = False
    Call ClearControl
    If Not lvwItem.SelectedItem Is Nothing Then
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    End If
End Sub
Private Sub ClearControl()
    cmdCancel.Enabled = False
    CmdDevSave.Enabled = False
    cmdTest.Enabled = False
    cboType.Enabled = False
    cmdPath.Enabled = False
    txtName.Enabled = False
    Cbosort.Enabled = False
    txtDirPath.Enabled = False
    txtDevAE.Enabled = False
    txtDevLocalAE.Enabled = False
    txtPassWord.Enabled = False
    txtUser.Enabled = False
    txtDevPort.Enabled = False
    txtFtpPath.Enabled = False
    txtDevAdress.Enabled = False
    txtDevNO = ""
    cboType.ListIndex = -1
    txtName = ""
    txtDirPath = ""
    txtDevAE = ""
    txtDevLocalAE = ""
    txtPassWord = ""
    txtUser = ""
    txtDevPort = ""
    txtFtpPath = ""
    txtDevAdress = ""
End Sub
Private Sub CmdDevSave_Click()
    '对于“影像设备”和“胶片打印”，先检查数量是否超过授权限制
    If cboType.ListIndex = 2 Or cboType.ListIndex = 3 Then
        If funCanAddModality = False Then Exit Sub
    End If
    '保存
    If Not DevSave Then Exit Sub
    '所有控件改为不可用
    Call ClearControl
    blnBeginchange = False
    '刷新数据
    Call InitlvwItem
End Sub
Private Function ValidData() As Boolean
    Dim j As Integer
    
    On Error GoTo err
    If cboType.list(cboType.ListIndex) = "存储设备" Then
        If Trim(txtUser) = "" Or Trim(txtPassWord) = "" Then
            MsgBoxD Me, "必须指定设备的用户名和密码,请检查.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Len(Trim(txtDevNO)) = 0 Then
        MsgBoxD Me, "请输入设备号！", vbInformation, gstrSysName
        txtDevNO.SetFocus
        Exit Function
    End If
    
    If Not (cboType.list(cboType.ListIndex) = "存储设备" Or cboType.list(cboType.ListIndex) = "远程目录") Then
        If Cbosort.Text = "" Then
            MsgBoxD Me, "请选择影像类别！", vbInformation, gstrSysName
            Cbosort.SetFocus: Exit Function
        End If
    End If
    
    If Len(Trim(txtName)) = 0 Then
        MsgBoxD Me, "请输入设备名！", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    
    If Me.cboType.ListIndex <> cboType.ListCount - 1 Then
        If UBound(Split(Trim(txtDevAdress), ".")) <> 3 Then
            MsgBoxD Me, "IP格式不正确，请检查！", vbInformation, gstrSysName
            txtDevAdress.SetFocus: Exit Function
        Else
            For j = 0 To 3
                If Not IsNumeric(Split(Trim(txtDevAdress), ".")(j)) Then
                    MsgBoxD Me, "IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                Else
                    If Split(Trim(txtDevAdress), ".")(j) < 0 Or Split(Trim(txtDevAdress), ".")(j) >= 256 Then
                        MsgBoxD Me, "IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                    End If
                End If
            Next
        End If
    Else
        If Trim(txtDirPath) = "" Then
            MsgBoxD Me, "请输入本地路径！", vbInformation, gstrSysName: txtDirPath.SetFocus: Exit Function
        End If
    End If
        
    If InStr(Trim(txtFtpPath.Text), ":") > 0 Then
        MsgBoxD Me, "FTP目录格式不正确，请检查！", vbInformation, gstrSysName
        txtFtpPath.SetFocus: Exit Function
    End If
    
    If cboType.ListIndex = 1 And (Len(Trim(txtDevPort)) = 0 Or Not IsNumeric(txtDevPort)) Then
        MsgBoxD Me, "请输入正确的端口号！", vbInformation, gstrSysName
        txtDevPort.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(txtName), vbFromUnicode)) > txtName.MaxLength Then
        MsgBoxD Me, "设备名超长（最多" & txtName.MaxLength & "个字符或" & CInt(txtName.MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    ValidData = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function DevSave() As Boolean
Dim DevType As String
Dim objFtp As New clsFtp
Dim strEncryptionPassW As String

    On Error GoTo DBError
    If ValidData = False Then Exit Function
    If zlStr.NeedName(Cbosort.list(Cbosort.ListIndex)) <> "" Then
        DevType = Split(Cbosort.list(Cbosort.ListIndex), "-")(0)
    End If
    
    '加密ftp密码
    If Trim(txtPassWord.Text) <> "" Then
        strEncryptionPassW = objFtp.GetEncryptionPassW(Trim(txtPassWord.Text))
        strEncryptionPassW = Mid(strEncryptionPassW, 1, 1) & "※" & Mid(strEncryptionPassW, 2)
        strEncryptionPassW = "★" & strEncryptionPassW & "★"
        strEncryptionPassW = Replace(strEncryptionPassW, "'", "''")
    End If
    
    gstrSQL = "zl_影像设备目录_Update('" & txtDevNO & "','" & Trim(txtName) & "'," & cboType.ListIndex + 1 & _
        ",'" & Trim(txtDevAdress) & "','" & Trim(txtFtpPath) & "','" & Trim(txtDevPort) & "','" & Trim(txtUser) & "','" & _
        strEncryptionPassW & "','" & Trim(txtShareDir) & "','" & Trim(txtSDUser) & "','" & Trim(txtSDPassword) & "','" & Trim(txtDevLocalAE) & "','" & Trim(txtDevAE) & "','" & Trim(txtDirPath) & "', 1,'" & DevType & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    DevSave = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdPath_Click()
    Dim strTmp As String
    '得到路径
    strTmp = BrowPath(Me.hWnd, "请选定保存的文件目录：")
    '当用新的路径时才保存
    If strTmp <> "" And strTmp <> txtDirPath.Text Then
        txtDirPath.Text = strTmp
    End If
End Sub

Private Sub cmdTest_Click()
    If ValidData = False Then Exit Sub
    Me.MousePointer = vbHourglass: cmdTest.Enabled = False
    Select Case cboType.Text
        Case "存储设备"
            If Len(Dir(Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsFtpTools.exe")) > 0 Then
                Shell Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsFtpTools.exe   " & txtUser.Text & "||" & txtPassWord & "||" & txtDevAdress.Text & "||" & txtFtpPath.Text, 1
            Else
                Call TestFTPDev
            End If
        Case "接收主机", "胶片打印", "影像设备"
            Call TestDev
        Case "远程目录"
            Call TestPath
    End Select
    Me.MousePointer = vbDefault: cmdTest.Enabled = True
End Sub

Private Sub Form_Load()
    blnBeginchange = False
    mstrPrivs = gstrPrivs
    Me.Icon = imgKind.ListImages(4).Picture
    Call InitCommandBars '初始化菜单
    Call InitSubWindow  '初始化子窗体
    Call InitlvwItem '初始化表格
    
    Call RestoreWinState(Me, App.ProductName)
    
    gintDICOM设备数量 = getLicenseCount(LOGIN_TYPE_DICOM设备)
    gint胶片打印机数量 = getLicenseCount(LOGIN_TYPE_胶片打印机)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    Call SaveWinState(Me, App.ProductName)
    Unload Me
End Sub
Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub
Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub
Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub




Private Sub lvwItem_DblClick()
Dim cbrControl As CommandBarControl
    If blnBeginchange Then txtName.SetFocus: Exit Sub
    Set cbrControl = cbrMain.FindControl(xtpControlButton, conMenu_Edit_Modify)
    If Not cbrControl Is Nothing Then Call cbrMain_Execute(cbrControl)
End Sub
Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objFtp As New clsFtp
    Dim strDecryptionPassW As String
    Dim i As Integer
'显示数据
    If blnBeginchange Then '开始修改或新增后
        If MsgBoxD(Me, "当前变更尚未保存，确实要更换查看吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtName.SetFocus
            Exit Sub
        End If
    End If
    blnBeginchange = False
    cmdCancel.Enabled = False
    CmdDevSave.Enabled = False
    cmdTest.Enabled = True
    If Cbosort.ListCount <= 0 Then Call InitCbosort
    txtDevNO = lvwItem.SelectedItem.Text
    txtName = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_名称").Index - 1)
    
    cboType.ListIndex = Decode(lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_类型").Index - 1), "存储设备", 0, "接收主机", 1, "胶片打印", 2, "影像设备", 3, "远程目录", 4)
    
    If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_影像类别").Index - 1) = "" Then '存储和远程目录不指定影像类别
        Cbosort.ListIndex = -1
    Else
        Call SeekIndex(Cbosort, lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_影像类别").Index - 1))
    End If
    txtDirPath = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_本地路径").Index - 1)
    txtDevAE = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_设备AE").Index - 1)
    txtDevLocalAE = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_本地AE").Index - 1)
    
    '解密ftp密码
    If Len(lvwItem.SelectedItem.tag) >= 3 Then
        If Mid(lvwItem.SelectedItem.tag, 1, 1) & Mid(lvwItem.SelectedItem.tag, 3, 1) & Mid(lvwItem.SelectedItem.tag, Len(lvwItem.SelectedItem.tag), 1) = "★※★" Then
            strDecryptionPassW = Mid(lvwItem.SelectedItem.tag, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = objFtp.GetDecryptionPassW(strDecryptionPassW)
            
            txtPassWord = strDecryptionPassW
        Else
            txtPassWord = lvwItem.SelectedItem.tag
        End If
    Else
        txtPassWord = lvwItem.SelectedItem.tag
    End If
    
    txtUser = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_FTP用户名").Index - 1)
    txtDevPort = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_端口号").Index - 1)
    txtFtpPath = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_Ftp目录").Index - 1)
    txtDevAdress = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_IP地址").Index - 1)
    txtShareDir = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_共享目录").Index - 1)
    txtSDUser = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_共享目录用户名").Index - 1)
    txtSDPassword = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_共享目录密码").Index - 1)
'相当于修改
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    With lvwItem
        .Top = 0
        .Left = 0
        .Width = PicList.Width - FraDevice.Width
        .Height = PicList.Height - 370
    End With
    With FraDevice
        .Top = 0
        .Left = lvwItem.Width
    End With
    With FraInfor
        .Top = FraDevice.Height
        .Left = FraDevice.Left
        .Height = PicList.Height - FraDevice.Height - 390
    End With
End Sub
Private Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    On Error GoTo OpenFileError
    With udtBI
        '设置浏览窗口
        .lngHwnd = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
    Exit Function
OpenFileError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetNewNo() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo DBError
    strSql = "Select Nvl(Max(To_Char(设备号,'000')),1) From 影像设备目录"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.EOF Then
        GetNewNo = "001"
    Else
        GetNewNo = Format(Val(rsTmp(0)) + 1, "000")
    End If
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub TestDev()
Dim objGlobal As New DicomGlobal
    On Error GoTo TestError
    If Trim(txtDevAdress.Text) = "" Then
        MsgBoxD Me, "请输入IP地址！", vbInformation, gstrSysName
        txtDevAdress.SetFocus: Exit Sub
    End If
    If Trim(txtDevPort.Text) = "" Or Not IsNumeric(txtDevPort.Text) Then
        MsgBoxD Me, "请输入正确的端口号！", vbInformation, gstrSysName
        txtDevPort.SetFocus: Exit Sub
    End If
    
    If Trim(txtDevAE.Text) = "" Then
        MsgBoxD Me, "请输入正确的设备AE！", vbInformation, gstrSysName
        txtDevAE.SetFocus: Exit Sub
    End If
    
    If Trim(txtDevLocalAE.Text) = "" Then
        MsgBoxD Me, "请输入正确的设备本地AE！", vbInformation, gstrSysName
        txtDevLocalAE.SetFocus: Exit Sub
    End If
    
    With objGlobal
        If .Echo(txtDevAdress, CLng(txtDevPort), txtDevLocalAE, txtDevAE) <> 0 Then
            MsgBoxD Me, "无法连接到指定的接收主机！", vbInformation, gstrSysName
            txtDevAdress.SetFocus
        Else
            MsgBoxD Me, "连接测试成功！", vbInformation, gstrSysName
        End If
    End With
    Exit Sub
TestError:
    Me.MousePointer = vbDefault
    MsgBoxD Me, "无法连接到指定的设备！", vbInformation, gstrSysName
End Sub
Private Sub TestFTPDev()
Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP类
    strPath = Format(zlDatabase.Currentdate, "yyyymmddHHMMSS")
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Print #1, "测试文件"
    Close #1
    If FtpNet.FuncFtpConnect(txtDevAdress, txtUser, txtPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir("/", "FTP测试" & strPath) > 0 Then
            MsgBoxD Me, "当前设备写目录访问失败", vbInformation, gstrSysName
        Else
            FtpNet.FuncFtpDelDir "/", "FTP测试" & strPath
            If CheckFtpDir(FtpNet, txtFtpPath) Then
                If FtpNet.FuncFtpMkDir(txtFtpPath, "FTP测试" & strPath) > 0 Then
                    MsgBoxD Me, "当前设备创建目录失败", vbInformation, gstrSysName
                ElseIf FtpNet.FuncUploadFile(txtFtpPath, strTmpPath, "temp.txt") > 0 Then
                    MsgBoxD Me, "当前设备上传文件失败", vbInformation, gstrSysName
                ElseIf FtpNet.FuncFtpGetFileSize(txtFtpPath, "temp.txt") <= 0 Then
                    MsgBoxD Me, "当前设备读取文件大小失败，" & IIf(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 1) <> 0, "建议取消注册表参数“启用FTP文件大小对比”。", ""), vbInformation, gstrSysName
                ElseIf FtpNet.FuncDelFile(txtFtpPath, "temp.txt") > 0 Then
                    MsgBoxD Me, "当前设备删除文件失败", vbInformation, gstrSysName
                Else
                    FtpNet.FuncFtpDisConnect '先断开，再删除，不然删不掉
                    If FtpNet.FuncFtpConnect(txtDevAdress, txtUser, txtPassWord) <= 0 Then
                        MsgBoxD Me, "当前设备不能连接！", vbInformation, gstrSysName
                    ElseIf FtpNet.FuncFtpDelDir(txtFtpPath, "FTP测试" & strPath) > 0 Then
                        MsgBoxD Me, "当前设备删除目录访问失败", vbInformation, gstrSysName
                    Else
                        MsgBoxD Me, "测试连接成功！", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    Else
        MsgBoxD Me, "当前设备不能连接！", vbInformation, gstrSysName
    End If
    FtpNet.FuncFtpDisConnect
    Kill strTmpPath
End Sub

Private Function CheckFtpDir(objFtp As clsFtp, strFtpDir As String) As Boolean
    
    CheckFtpDir = True
    If Len(Trim(Replace(Replace(strFtpDir, "/", ""), "\", ""))) <> 0 Then
        objFtp.FuncChangeDir ""
        If objFtp.FuncChangeDir(strFtpDir) <> 0 Then
            If MsgBox("当前FTP目录不存在，是否创建后测试？", vbYesNo, gstrSysName) = vbYes Then
                If objFtp.FuncFtpMkDir("", strFtpDir) <> 0 Then
                    CheckFtpDir = False
                    MsgBoxD Me, "当前FTP目录不合法！", vbInformation, gstrSysName
                    txtFtpPath.SetFocus
                End If
            Else
                CheckFtpDir = False
                MsgBoxD Me, "当前FTP目录不存在！", vbInformation, gstrSysName
                txtFtpPath.SetFocus
            End If
        End If
    End If
End Function

Private Sub TestPath()
Dim duTime As Double
    On Error GoTo TestError
    If Trim(txtDirPath.Text) = "" Then
        MsgBoxD Me, "请选择或输入要访问的远程主机", vbInformation, gstrSysName
        txtDirPath.SetFocus
        Exit Sub
    End If
    
    duTime = Timer
    Do Until CLng(Timer - duTime) >= 20
        Shell "net use " & txtDirPath & " " & txtPassWord & " /user:" & txtUser, vbHide
        If WriteTest(False) = True Then
            MsgBoxD Me, "连接测试成功！", vbInformation, gstrSysName
        Else
            MsgBoxD Me, "无法连接到指定的接收主机！", vbInformation, gstrSysName
        End If
        Exit Do
        DoEvents
    Loop
    Shell "net use " & txtDirPath & " /delete "
    Exit Sub
TestError:
    Me.MousePointer = vbDefault
    MsgBoxD Me, "无法连接到指定的设备！", vbInformation, gstrSysName
End Sub
Private Function WriteTest(ShowErrMsg As Boolean) As Boolean
    Dim strTmpPath As String
    On Error GoTo CopyError
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Close #1
    FileCopy strTmpPath, IIf(Len(txtDirPath) > 3, txtDirPath & "\", txtDirPath) & "temp.txt"
    Kill IIf(Len(txtDirPath) > 3, txtDirPath & "\", txtDirPath) & "temp.txt"
    Kill strTmpPath
    WriteTest = True
    Exit Function
CopyError:
    If ShowErrMsg = False Then Exit Function
    If err.Number = 75 Then
        MsgBoxD Me, "写入测试失败!请查看[" & txtDirPath & "]是否有写入权限!", vbInformation, App.EXEName
    Else
        MsgBoxD Me, "发生其他错误！", vbQuestion, App.EXEName
    End If
End Function

Private Function funCanAddModality() As Boolean
'检查DICOM设备和胶片打印机的数量，判断是否可以新增
'参数：
'返回值：   True--可以添加；False--不允许添加
    Dim i As Integer
    Dim str类型 As String   '"影像设备"或者“胶片打印”
    Dim intSum As Integer
    
    On Error GoTo err
    str类型 = cboType.list(cboType.ListIndex)
    intSum = 0
    For i = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(i).SubItems(lvwItem.ColumnHeaders("_类型").Index - 1) = str类型 And _
            lvwItem.ListItems(i).SubItems(lvwItem.ColumnHeaders("_状态").Index - 1) = "在用" Then
            intSum = intSum + 1
        End If
    Next i
    '新增的设备
    If txtDevNO = GetNewNo Then
        intSum = intSum + 1
    End If
    
    If str类型 = "影像设备" Then
        If intSum <= gintDICOM设备数量 Or gintDICOM设备数量 = -1 Then
            funCanAddModality = True
            Exit Function
        End If
    ElseIf str类型 = "胶片打印" Then
        If intSum <= gint胶片打印机数量 Or gint胶片打印机数量 = -1 Then
            funCanAddModality = True
            Exit Function
        End If
    End If
    funCanAddModality = False
    MsgBoxD Me, str类型 & "超过您购买的总数量（" & _
        IIf(str类型 = "影像设备", gintDICOM设备数量, gint胶片打印机数量) & _
        "），无法添加。请向软件供应商联系。", vbOKOnly, gstrSysName
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

