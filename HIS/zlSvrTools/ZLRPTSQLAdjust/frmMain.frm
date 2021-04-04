VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "中联软件辅助开发管理工具"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgThis 
      Left            =   4995
      Top             =   3090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BB2
            Key             =   "_程序组"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F4C
            Key             =   "_程序项"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22E6
            Key             =   "_工具集"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2680
            Key             =   "_字典管理"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A1A
            Key             =   "_消息收发"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DB4
            Key             =   "_提醒查阅"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":314E
            Key             =   "_系统选项"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E8
            Key             =   "_"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3882
            Key             =   "__参数上传"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C1C
            Key             =   "_参数恢复"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   5655
      Top             =   1935
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
            Picture         =   "frmMain.frx":4E08
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5260
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56B2
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4380
      Top             =   1350
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
            Picture         =   "frmMain.frx":59CC
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E24
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6276
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog DlgMain 
      Left            =   4140
      Top             =   4515
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   4590
      Top             =   2340
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
            Picture         =   "frmMain.frx":6590
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B2A
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70C4
            Key             =   "child"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   9585
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   450
      Width1          =   930
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   450
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgToolsStard"
         HotImageList    =   "imgToolsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "help"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "exit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Height          =   5400
      Left            =   0
      ScaleHeight     =   5340
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   510
      Width           =   2955
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1485
         Index           =   0
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1500
         _cx             =   2646
         _cy             =   2619
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   13811126
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483634
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
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
      Begin VB.PictureBox picVbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4260
         Left            =   2730
         MousePointer    =   9  'Size W E
         ScaleHeight     =   4260
         ScaleWidth      =   45
         TabIndex        =   4
         Top             =   165
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcThis 
         Height          =   4875
         Left            =   525
         TabIndex        =   6
         Top             =   210
         Width           =   2025
         _Version        =   589884
         _ExtentX        =   3572
         _ExtentY        =   8599
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5910
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   661
      SimpleText      =   $"frmMain.frx":765E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":76A5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11853
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
   Begin MSComctlLib.ImageList imgToolsHot 
      Left            =   4770
      Top             =   705
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
            Picture         =   "frmMain.frx":7F37
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8151
            Key             =   "print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":836B
            Key             =   "view"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8585
            Key             =   "help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":879F
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   3900
      Top             =   690
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
            Picture         =   "frmMain.frx":89B9
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BD3
            Key             =   "print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DED
            Key             =   "view"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9007
            Key             =   "help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9221
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogout 
         Caption         =   "注销(&L)"
      End
      Begin VB.Menu mnuFileExit 
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
         Begin VB.Menu mnuViewToolSpilt1 
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
      Begin VB.Menu mnuViewSpilt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolList 
         Caption         =   "工具列表(&L)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "自定义"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuItems 
         Caption         =   "自定义项目0"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim intCount As Integer
Dim strPreKey As String     '上次选择的Key值
Private Enum mCol
    空列 = 0
    标题 = 1
    说明 = 3
    模块 = 4
    系统 = 5
    部件 = 6
End Enum
Private Sub MDIForm_Load()
    Dim objNode As Node
    
    Me.Caption = Me.Caption & " [" & gstrUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    
    gstrSysName = gstrProductName & "软件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    mnuHelpWeb.Caption = "&Web上的" & gstrWebSustainer
    mnuHelpWebHome.Caption = gstrWebSustainer & "主页"
    With Me.tbcThis
        .SetImageList Me.imgThis
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .Color = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
    End With
    
    LoadFunctionMenu
 End Sub
Private Function LoadFunctionMenu() As Boolean
    '功能:加载功能菜单
    '返回:成功返回true,否则返回False
    Dim i As Long
    Dim strProgram(0 To 4) As String
    Dim lngCount As Long
    
    lngCount = Me.tbcThis.ItemCount
    
    If lngCount > 0 Then Load Me.vfgThis(lngCount): Me.vfgThis(lngCount).Visible = True
    
    Call Me.tbcThis.InsertItem(lngCount, "安装脚本管理", Me.vfgThis(lngCount).hwnd, 2)
    
    With Me.vfgThis(lngCount)
        .MergeCells = flexMergeFree
        .ForeColor = vbBlue
        .Rows = 0: .Cols = 7
        .ColWidth(mCol.空列) = 280: .ColWidth(mCol.标题) = .ColWidth(0)
        .ColWidth(mCol.标题 + 1) = .Width - .ColWidth(0) * 2 - Screen.TwipsPerPixelX
        .ColWidth(mCol.说明) = 0: .ColWidth(mCol.模块) = 0: .ColWidth(mCol.系统) = 0: .ColWidth(mCol.部件) = 0
        
        For i = 0 To 2
            .Rows = .Rows + 1: .MergeRow(.Rows - 1) = True
            Select Case i
            Case 0
                strProgram(0) = "01"
                strProgram(1) = "生成安装脚本"
                strProgram(2) = "根据数据库中的相关结构数据，自动生成相关的安装脚本"
                strProgram(3) = ""
            Case 1
                strProgram(0) = "02"
                strProgram(1) = "报表脚本工具"
                strProgram(2) = "生成报表数据安装脚本和执行指定文件的报表安装脚本"
                strProgram(3) = ""
            Case 2
                strProgram(0) = "03"
                strProgram(1) = "字典核对工具"
                strProgram(2) = "根据数据结构或脚本文件，核对数据字典是否相符"
                strProgram(3) = ""
            End Select
            .TextMatrix(.Rows - 1, mCol.模块) = strProgram(0)
            .TextMatrix(.Rows - 1, mCol.标题) = strProgram(1):
            .TextMatrix(.Rows - 1, mCol.说明) = .TextMatrix(.Rows - 1, mCol.标题)
            .TextMatrix(.Rows - 1, mCol.说明) = strProgram(2)
            Set .Cell(flexcpPicture, .Rows - 1, mCol.标题) = Me.imgThis.ListImages(i + 4).Picture
        Next
    End With
    
End Function
Private Sub MDIForm_Resize()
    On Error Resume Next
    If picVbar.Left < 2000 Then picVbar.Left = 2000
    If picVbar.Left > Width - 3000 Then picVbar.Left = Width - 3000
    picVbar.Top = 0
    picVbar.Height = picFunc.Height
    picFunc.Width = picVbar.Left + picVbar.Width + 45
    
       
    
    If stbThis.Panels(2) = "" Then
        '特殊处理，不然状态栏的宽度不正确
        stbThis.Panels(2) = " "
        stbThis.Panels(2) = ""
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmChild As Form
    Err = 0: On Error Resume Next
    For Each frmChild In Forms
        Unload frmChild
    Next
   '----------------------------------------
    '关闭公共部件的窗体
    CloseWindows
    
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    gfrmActive.subPrint 3
End Sub

Private Sub mnuFileLogout_Click()
    Unload Me
    Call Main
End Sub

Private Sub mnuFilePreView_Click()
    gfrmActive.subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    gfrmActive.subPrint 1
End Sub

Private Sub mnuFileSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.ShowAbout
End Sub

Private Sub mnuHelpWebHome_Click()
    ShellExecute hwnd, "open", "http://" & gstrWebURL, "", "", 1
End Sub

Private Sub mnuHelpWebMail_Click()
    ShellExecute hwnd, "open", "mailto:" & gstrWebEmail, "", "", 1
End Sub

'Private Sub mnuToolSysManage_Click(Index As Integer)
'    On Error Resume Next
'    tvwFunc.Nodes("_01").Expanded = True
'    tvwFunc.Nodes("_" & mnuToolSysManage(Index).Tag).EnsureVisible
'    tvwFunc.Nodes("_" & mnuToolSysManage(Index).Tag).Selected = True
'    Call RunByModule(mnuToolSysManage(Index).Tag)
'End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call MDIForm_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call MDIForm_Resize
End Sub

Private Sub mnuViewToolList_Click()
    '显示或隐藏工具列表
    mnuViewToolList.Checked = Not mnuViewToolList.Checked
    picFunc.Visible = mnuViewToolList.Checked
End Sub

Private Sub mnuViewToolText_Click()
    If mnuViewToolText.Checked = False Then
        tbrThis.TextAlignment = tbrTextAlignRight
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Tag = tbrThis.Buttons(intCount).Caption
            tbrThis.Buttons(intCount).Caption = ""
        Next
        tbrThis.TextAlignment = tbrTextAlignBottom
    End If
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    tbrThis.Refresh
End Sub

Private Sub MyOutLook_ItemClick(item As OutItem)
    If strPreKey = item.Key Then Exit Sub
    strPreKey = item.Key
    Call RunByModule(Mid(item.Key, 2))
End Sub

Private Sub picFunc_Resize()
    Dim lngCount  As Long
    Call MDIForm_Resize
    
    With Me.tbcThis
        .Left = 10: .Width = picFunc.ScaleWidth
        .Top = 10: .Height = picFunc.ScaleHeight
    End With
    For lngCount = 0 To Me.tbcThis.ItemCount - 1
        With Me.vfgThis(lngCount)
            .ColWidth(2) = picFunc.ScaleWidth - .ColWidth(0) * 2 - Screen.TwipsPerPixelX
        End With
    Next
End Sub

Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIf(picVbar.Left + x < 2000, 2000, picVbar.Left + x)
        Call MDIForm_Resize
    End If
End Sub

Private Sub tabCol_SelectedChanged(ByVal item As XtremeSuiteControls.ITabControlItem)

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "preview"
        mnuFilePreView_Click
    Case "print"
        mnuFilePrint_Click
    Case "help"
         '
    Case "exit"
        Call mnuFileExit_Click
    End Select
End Sub

'Private Sub tvwFunc_DblClick()
'    If tvwFunc.SelectedItem Is Nothing Then Exit Sub
'    If tvwFunc.SelectedItem.Selected = False Then Exit Sub
'    Call tvwFunc_NodeClick(tvwFunc.SelectedItem)
'End Sub

'Private Sub tvwFunc_NodeClick(ByVal Node As MSComctlLib.Node)
'    If tvwFunc.Tag = Mid(Node.Key, 2) Then Exit Sub
'Private Sub tvwFunc_NodeClick(ByVal Node As MSComctlLib.Node)
'    If tvwFunc.Tag = Mid(Node.Key, 2) Then Exit Sub
'    Call RunByModule(Mid(Node.Key, 2))
'End Sub

Private Sub RunByModule(ByVal strModule As String)
    Dim frmChild As Form
    
    For Each frmChild In Forms
        If frmChild Is Me Then
        ElseIf frmChild.MDIChild = True And frmChild.Enabled = True Then
            Unload frmChild
        End If
    Next
    
    Set gfrmActive = Nothing
    
    Select Case strModule
        Case "0101"   '安装脚本生成器
            Set gfrmActive = frmAppSteupSQLBuild
        Case "0102" '报表脚本管理
            Set gfrmActive = frmRptSQLMgr
        Case "0103" '数据字典核对工具"
            Set gfrmActive = frmCheckScrip
    End Select
    If Not gfrmActive Is Nothing Then
        Call FindWindowAndSetActive(gfrmActive)
        gfrmActive.Show
        gfrmActive.ZOrder 0
    End If
    Call SetEnable
End Sub

Private Sub SetEnable()
    Dim blnEnable As Boolean
    
    If gfrmActive Is Nothing Then
        blnEnable = False
    Else
        blnEnable = gfrmActive.SupportPrint()
    End If
    mnuFilePrint.Enabled = blnEnable
    mnuFilePreView.Enabled = blnEnable
    mnuFileExcel.Enabled = blnEnable
    tbrThis.Buttons("preview").Enabled = blnEnable
    tbrThis.Buttons("print").Enabled = blnEnable
End Sub

Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    
    '--如果该窗体已经打开,则激活它(这样,窗体的大小不会发生变化)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub
Private Sub ExecuteFunc(ByVal lngSys As Long, ByVal strDLLName As String, ByVal lngModul As Long)
    '-------------------------------------------------------------
    '功能：调用执行指定部件的功能程序
    '参数： lngSys-系统
    '       strDLLName-部件名
    '       lngModul-模块号
    '-------------------------------------------------------------
    If lngModul = 0 Then Exit Sub
        
End Sub

Private Sub vfgThis_AfterSelChange(Index As Integer, ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Err = 0: On Error Resume Next
    With Me.vfgThis(Index)
        .CellBorderRange OldRowSel, 0, OldRowSel, 2, RGB(255, 255, 255), 0, 0, 0, 0, 0, 0
        .CellBorderRange NewRowSel, 0, NewRowSel, 2, RGB(0, 64, 128), 1, 1, 1, 1, 0, 0
    End With
End Sub

Private Sub vfgThis_DblClick(Index As Integer)
   With vfgThis(Index)
    ExecuteFunc Val(.TextMatrix(.Row, mCol.系统)), .TextMatrix(.Row, mCol.部件), Val(.TextMatrix(.Row, mCol.模块))
   End With
End Sub
Private Sub vfgThis_GotFocus(Index As Integer)
    Me.Height = Me.Height + Screen.TwipsPerPixelY
    Me.Height = Me.Height - Screen.TwipsPerPixelY
End Sub

