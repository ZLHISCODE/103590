VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLISSrv 
   AutoRedraw      =   -1  'True
   Caption         =   "仪器数据接收"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmLISSrv.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   StartUpPosition =   2  '屏幕中心
   WindowState     =   1  'Minimized
   Begin VB.Timer timConn 
      Interval        =   30000
      Left            =   9930
      Top             =   3030
   End
   Begin Zl9LISComm.ctrlComm DevComm 
      Height          =   495
      Index           =   0
      Left            =   8280
      TabIndex        =   10
      Top             =   1785
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   873
   End
   Begin VB.PictureBox picTmp 
      Height          =   1080
      Left            =   8535
      ScaleHeight     =   1020
      ScaleWidth      =   1170
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picIcon 
      Height          =   285
      Left            =   8790
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   4170
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame fraUD_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   3480
      Width           =   8745
   End
   Begin MSComctlLib.ListView lvwResult 
      Height          =   2730
      Left            =   150
      TabIndex        =   6
      Top             =   3720
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwLISRec 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   735
      Top             =   2190
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
            Picture         =   "frmLISSrv.frx":08CA
            Key             =   "_0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":0E64
            Key             =   "_1"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10995
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinWidth1       =   4995
      MinHeight1      =   720
      NewRow1         =   0   'False
      Caption2        =   "仪器"
      Child2          =   "cboDev"
      MinWidth2       =   3795
      MinHeight2      =   300
      Width2          =   5940
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboDev 
         Height          =   300
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   5100
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   4995
         _ExtentX        =   8811
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "预览"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "打印"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "连接"
               Key             =   "连接"
               Object.ToolTipText     =   "连接当前仪器准备接收数据"
               Object.Tag             =   "连接"
               ImageKey        =   "连接"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "断开"
               Key             =   "断开"
               Object.ToolTipText     =   "与当前仪器断开连接"
               Object.Tag             =   "断开"
               ImageKey        =   "断开"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLISSrv.frx":13FE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3330
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   1335
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":1C92
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":1EAC
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":20C6
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":22E0
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":24FA
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":2BF4
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":32EE
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":39E8
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":40E2
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":47DC
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":4ED6
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":55D0
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":5CCA
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":63C4
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":6ABE
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":71B8
            Key             =   "连接"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":7932
            Key             =   "断开"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1935
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":80AC
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":82C6
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":84E0
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":86FA
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":8914
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":900E
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":9708
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":9E02
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":A4FC
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":ABF6
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":B2F0
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":B9EA
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":C0E4
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":C7DE
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":CED8
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":D5D2
            Key             =   "连接"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISSrv.frx":DD4C
            Key             =   "断开"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock WinsockS 
      Left            =   9840
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "参数设置(&S)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFtpSet 
         Caption         =   "FTP设置(&F)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "连接(&C)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "断开(&D)"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevExp 
         Caption         =   "导出基础数据(&P)"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "科室选择(&D)"
            Checked         =   -1  'True
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuViewTool_1 
            Caption         =   "-"
            Visible         =   0   'False
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCharge 
         Caption         =   "只显示已经收费的病人(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "数据过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewComm 
         Caption         =   "通讯监控(&C)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReLoad 
         Caption         =   "重启通讯(&L)"
      End
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
Attribute VB_Name = "frmLISSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '要求变量声明
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private lngPort As Long
Private lngErrCounts As Long, strWhere As String, strBeginDate As String
Private strDevIDs As String '连接的设备ID

'********************返回给技师站的信息*****************************
Private Const strSend_Refresh = "Refresh"      '已保存数据可以刷新
Private Const strSend_True = "True"            '已操作成功
Private Const strSend_False = "False"          '操作失败
Private Const strSend_AutoCompute = "AutoCompute" '计算完成
'*******************************************************************

Private WithEvents frmView As frmViewComm
Attribute frmView.VB_VarHelpID = -1
Private mblnOwner As Boolean '是否所有者登录
Private mfsoTmp As New FileSystemObject  '文件对象

Private Sub cboDev_Click()
    ListSeq strWhere
    ShowMenu
End Sub

Private Sub DevComm_DevDecode(Index As Integer, ByVal commport As String, ByVal str结果 As String)
    Dim strCOM As String
    If frmView Is Nothing Then Exit Sub
    
    If InStr(commport, ".") <= 0 Then strCOM = "COM" & Val(commport)
    
    If InStr(cboDev.List(cboDev.ListIndex), strCOM) > 0 Then
        If str结果 <> "" Then
         '  显示收到的解析结果
            Call frmView.ShowDecode(0, str结果)
        End If
    End If
End Sub

Private Sub DevComm_DevRefresh(Index As Integer, ByVal lngID As Long)
    '发送刷新消息到LISWORK
    If lngID <> 0 Then
        Me.WinsockS.SendData Me.WinsockS.LocalIP & ";" & strSend_Refresh & ";" & lngID
    End If
End Sub

Private Sub DevComm_ItemUnknown(Index As Integer, ByVal commport As String, ByVal strItems As String)
    Dim strCOM As String
    If frmView Is Nothing Then Exit Sub
    
    If InStr(commport, ".") <= 0 Then strCOM = "COM" & Val(commport)
    
    If InStr(cboDev.List(cboDev.ListIndex), strCOM) > 0 Then
        If strItems <> "" Then
         '  显示收到的未知项
            Call frmView.ShowDecode(1, strItems)
        End If
    End If
End Sub

Private Sub DevComm_ReturnCompute(Index As Integer, ByVal strReturn As String)
    
    If DevComm(Index) Is Nothing Then Exit Sub
    If blnDataReceived Then Exit Sub
    blnDataReceived = True
    With Me.WinsockS
        .SendData .LocalIP & ";" & "AutoQCCompute|" & strReturn
    End With
    blnDataReceived = False
End Sub

Private Sub Form_Activate()
    ListSeq strWhere
    ShowMenu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub fraUD_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    fraUD_s.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraUD_s.Top + y < 2000 Then
        fraUD_s.Top = 2000
    ElseIf Me.ScaleHeight - fraUD_s.Top - y < 4000 Then
        fraUD_s.Top = Me.ScaleHeight - 4000
    Else
        fraUD_s.Top = fraUD_s.Top + y
    End If
End Sub

Private Sub fraUD_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    fraUD_s.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub frmView_CloseWindow()
    mnuViewComm.Checked = False
End Sub

Private Sub lvwLISRec_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call gobjControl.LvwSortColumn(lvwLISRec, ColumnHeader.Index)
End Sub

Private Sub lvwLISRec_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ListResult
End Sub

Private Sub mnuDevExp_Click()
    Dim strFile As String
    strFile = ExpLisDevData
    If strFile <> "" Then
        MsgBox "已导出到 " & strFile & "！"
    End If
    
End Sub

Private Sub mnuFileClose_Click()
    If Me.cboDev.ListIndex = -1 Then Exit Sub
    
'    Me.DevComm(Me.cboDev.ListIndex + 1).ClosePort
    ShowMenu
End Sub

Private Sub mnuFileOpen_Click()
    If Me.cboDev.ListIndex = -1 Then Exit Sub
    If gblnFromDB Then
        mMakeNoRule = gobjDatabase.GetPara("标本序号生成规则", glngSys, 1208, "今  天")
    Else
        mMakeNoRule = GetSetting("ZLSOFT", "公共模块\zl9LisWork\frmLabMain", "标本序号生成规则", "今  天")
    End If
'    Me.DevComm(Me.cboDev.ListIndex + 1).OpenPort
    ShowMenu
End Sub

Private Sub mnuFileSetup_Click()

    If frmParaSet.ShowMe(Me) Then
        If Not ReadPara("ResetExe") Then Unload Me
    End If
End Sub

Private Sub mnuFtpSet_Click()
    frmFtpSet.Show vbModal
End Sub

Private Sub mnuReLoad_Click()
    Dim intTime As Integer
    Dim tsmTmp As TextStream
    Dim objWait As New clsLISComm
    On Error GoTo errH

'    For inttime = LBound(g仪器) To UBound(g仪器)
'        If mfsoTmp.FileExists(g仪器(inttime).通讯目录 & "\Lock.txt") Then
'            Set tsmTmp = mfsoTmp.CreateTextFile(g仪器(inttime).通讯目录 & "\Send\CloseExe.txt")
'            tsmTmp.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss")
'            tsmTmp.Close
'            Set tsmTmp = Nothing
'        End If
'    Next
    
    Call KillProc("zlLisReceiveSend.exe")
    For intTime = LBound(g仪器) To UBound(g仪器)
        If Dir(g仪器(intTime).通讯目录 & "\Lock.txt") <> "" Then Kill g仪器(intTime).通讯目录 & "\Lock.txt"
    Next
    '延时1.5秒再启动接口
    objWait.Wait 1500
    Set objWait = Nothing
    Call ReadPara("")
    Exit Sub
errH:
    WriteLog "mnuReload", LOG_错误日志, Err.Number, Err.Description
End Sub

Private Sub mnuViewComm_Click()
    
    If mnuViewComm.Checked Then
        If Not frmView Is Nothing Then Unload frmView
    Else
        mnuViewComm.Checked = True
        If frmView Is Nothing Then Set frmView = New frmViewComm
        Call frmView.ShowMe(cboDev.List(cboDev.ListIndex), Me.DevComm(Me.cboDev.ListIndex + 1).CommSetting, Me.DevComm(Me.cboDev.ListIndex + 1).DevProgName)
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    ListSeq strWhere
End Sub

Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '--------------------------------------------------------------------------------------------------
    '功能:  处理图标的各种处理事件
    '--------------------------------------------------------------------------------------------------
    On Error Resume Next
    Select Case Button '
        Case vbLeftButton
            Me.Show
            Me.WindowState = vbNormal
        Case vbRightButton
            ModifyIcon picIcon.hwnd, Me.Icon, , False
            Me.PopupMenu Me.mnuFile
            ModifyIcon picIcon.hwnd, Me.Icon
    End Select '
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "连接"
            mnuFileOpen_Click
        Case "断开"
            mnuFileClose_Click
        Case "退出"
            mnuFileQuit_Click
        Case "打印"
            mnuFilePrint_Click
        Case "预览"
            mnuFilePreview_Click
        Case "帮助"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuFilePrintSet_Click()
'功能：打印设置
    Call gobjPrintMode.zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
'功能：输出到Excel
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
'功能：打印预览
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
'功能：打印
    Call OutputList(1)
End Sub

Private Sub mnuHelpTitle_Click()
'功能：调用帮助主题
    gobjComLib.ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    If MsgBox("退出后将不能接收仪器数据！是否确定要退出？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTmp As ADODB.Recordset
    
    '和技师站的通讯接口
    With Me.WinsockS
        .Protocol = sckUDPProtocol
        .RemoteHost = "Localhost"
        .RemotePort = 1001
        .Bind 1000
    End With
    
    strBeginDate = Format(date & " " & Time, "yyyy-MM-dd hh:mm:ss")
    lngErrCounts = 0
    strWhere = ""
    
    With lvwLISRec
        With .ColumnHeaders
            .Clear
            .Add , , "病人姓名", 1000
            .Add , , "标本号", 800, 1
            .Add , , "标本", 1000
            .Add , , "申请项目", 2500
            .Add , , "申请科室", 1200
            .Add , , "医生", 1000
            .Add , , "申请时间", 2000
            .Add , , "检验人", 1000
            .Add , , "检验时间", 2000
            .Add , , "质控品", 800
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    With lvwResult
        With .ColumnHeaders
            .Clear
            .Add , , "检验项目", 2000
            .Add , , "检验结果", 1200, 1
            .Add , , "标志", 1000
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    
    
    '获取连接的设备
    If Dir(App.Path & "\zlLisReceiveSend.exe") = "" Then
        MsgBox "缺少通讯程序，zlLisReceiveSend.Exe，程序不能运行！", vbQuestion, "zl9LisComm"
        End
    Else
        If Not ReadPara("") Then End
    End If
    

    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    
    If gblnFromDB Then
        mMakeNoRule = gobjDatabase.GetPara("标本序号生成规则", glngSys, 1208, "今  天")
    Else
        mMakeNoRule = GetSetting("ZLSOFT", "公共模块\zl9LisWork\frmLabMain", "标本序号生成规则", "今  天")
    End If
    gstrSQL = "Select 所有者 from zlsystems where 编号=[1] And 所有者=[2]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, glngSys, UCase(gstrDbUser))
    mblnOwner = False
    If rsTmp.RecordCount > 0 Then
        mblnOwner = True
    End If

    Call mnuReLoad_Click
End Sub

Private Sub cbr_Resize()
    Call Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    gobjComLib.ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click(Index As Integer)
    Dim blnEnabled As Boolean, blnVisible As Boolean, i As Integer
    
    mnuViewToolItem(Index).Checked = Not mnuViewToolItem(Index).Checked
    cbr.Bands(Index + 1).Visible = Not cbr.Bands(Index + 1).Visible

    blnEnabled = False: blnVisible = False
    For i = 1 To cbr.Bands.Count
        '只有有一个ToolBar可见,则"显示文本"菜单可见
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            If cbr.Bands(i).Visible Then
                blnEnabled = True
            End If
        End If
        '只要有一个Band可见,则CoolBar可见
        If cbr.Bands(i).Visible Then
            blnVisible = True
        End If
    Next
    mnuViewToolText.Enabled = blnEnabled
    cbr.Visible = blnVisible
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To cbr.Bands.Count
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            For j = 1 To cbr.Bands(i).Child.Buttons.Count
                cbr.Bands(i).Child.Buttons(j).Caption = IIf(mnuViewToolText.Checked, cbr.Bands(i).Child.Buttons(j).Tag, "")
            Next
            If Not mnuViewToolText.Checked Then
                cbr.Bands(i).Child.TextAlignment = tbrTextAlignBottom
            End If
            cbr.Bands(i).MinHeight = cbr.Bands(i).Child.ButtonHeight
            cbr.Bands(i).Child.Refresh
        End If
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    gobjComLib.zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    gobjComLib.zlMailTo hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long

    On Error Resume Next
    
    Select Case WindowState
        Case vbMinimized
            Me.Hide
            AddIcon picIcon.hwnd, Me.Icon
        Case Else
            RemoveIcon picIcon.hwnd
    End Select

    If WindowState = 1 Then Exit Sub
    
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With Me.fraUD_s
        If .Top > Me.ScaleHeight Then .Top = cbrH + (Me.ScaleHeight - cbrH) / 2
        .Left = 0: .Width = Me.ScaleWidth
    End With
    
    With lvwResult
        .Left = 0: .Top = fraUD_s.Top + fraUD_s.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - staH - .Top
    End With
    
    With lvwLISRec
        .Left = 0
        .Top = cbrH
        .Height = fraUD_s.Top - .Top
        .Width = Me.ScaleWidth
    End With
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To DevComm.UBound
        Unload DevComm(i)
    Next
    RemoveIcon picIcon.hwnd
    Call gobjComLib.SaveWinState(Me, App.ProductName)
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能: 输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As Object
    Set objOut = CreateObject("zl9PrintMode.zlPrintLvw")
    On Error Resume Next
    
    
    If lvwLISRec.SelectedItem Is Nothing Then Exit Sub
    
    Set objOut.Body.objData = Me.lvwLISRec
    objOut.Title.Text = "检验记录"
    objOut.UnderAppItems.Add ""
    objOut.UnderAppItems.Add "时间：" & strBeginDate & " - " & Format(date & " " & Time, "yyyy-MM-dd HH:mm:SS")
    If bytStyle = 1 Then
        bytStyle = gobjPrintMode.zlPrintAsk(objOut)
        If bytStyle <> 0 Then gobjPrintMode.zlPrintOrViewLvw objOut, bytStyle
    Else
        gobjPrintMode.zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub ListSeq(ByVal strWhere As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strCurKey As String
    Dim tmpItem As MSComctlLib.ListItem
    Dim strIDs As String '标本记录ID枚举
    Dim aDevIDs() As String
    
    If cboDev.ListIndex = -1 Then Me.lvwLISRec.ListItems.Clear: Exit Sub
    On Error GoTo DBError
    If Not lvwLISRec.SelectedItem Is Nothing Then strCurKey = lvwLISRec.SelectedItem.Key

    Me.lvwLISRec.ListItems.Clear
    If Len(strDevIDs) > 0 Then
        aDevIDs = Split(strDevIDs, ",")
        gstrSQL = "Select Distinct A.ID,A.标本序号,A.标本类型,A.申请时间,A.检验人,A.检验时间,A.是否质控品," & _
            "C.姓名,B.医嘱内容,D.名称,B.开嘱医生 " & _
            "From 检验标本记录 A,病人医嘱记录 B,病人信息 C,部门表 D " & _
            "Where A.医嘱ID=B.ID(+) And B.病人ID=C.病人ID(+) And B.开嘱科室ID=D.ID(+) " & _
            " And A.检验时间 Between [1] And Sysdate" & _
            " And A.仪器ID =[2]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, CDate(strBeginDate), CLng(aDevIDs(Me.cboDev.ListIndex)))
        Do While Not rsTmp.EOF
            Set tmpItem = lvwLISRec.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("姓名")))
            With tmpItem
                .SubItems(1) = Nvl(rsTmp("标本序号"))
                .SubItems(2) = Nvl(rsTmp("标本类型"))
                .SubItems(3) = Nvl(rsTmp("医嘱内容"))
                .SubItems(4) = Nvl(rsTmp("名称"))
                .SubItems(5) = Nvl(rsTmp("开嘱医生"))
                .SubItems(6) = Nvl(rsTmp("申请时间"))
                .SubItems(7) = Nvl(rsTmp("检验人"))
                .SubItems(8) = Nvl(rsTmp("检验时间"))
                .SubItems(9) = IIf(Nvl(rsTmp("是否质控品"), 0) = 0, "  ", "√")

                If .Key = strCurKey Then .Selected = True
            End With

            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
DBError:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog("frmLISSrv.ListSeq", LOG_错误日志, Err.Number, Err.Description)
End Sub

Private Sub ListResult()
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    lvwResult.ListItems.Clear
    If lvwLISRec.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo DBError
    gstrSQL = "Select A.ID,B.中文名,A.检验结果,Decode(A.结果标志,Null,'正常',1,'正常',2,'偏低',3,'偏高') As 结果标志 " & _
        "From 检验普通结果 A,诊治所见项目 B,检验标本记录 C " & _
        "Where A.检验项目ID=B.ID And A.检验标本ID=C.ID And A.记录类型=C.报告结果 And A.检验标本ID=[1]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, Me.Caption, CLng(Mid(lvwLISRec.SelectedItem.Key, 2)))
    Do While Not rsTmp.EOF
        Set tmpItem = lvwResult.ListItems.Add(, "_" & rsTmp("ID"), Nvl(rsTmp("中文名")))
        With tmpItem
            .SubItems(1) = Nvl(rsTmp("检验结果"))
            .SubItems(2) = Nvl(rsTmp("结果标志"))
        End With

        rsTmp.MoveNext
    Loop
    Exit Sub
DBError:
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog("frmLISSrv.ListResult", LOG_错误日志, Err.Number, Err.Description)
End Sub

Private Function ReadPara(ByVal strCmd As String) As Boolean
    '
    '获取连接的设备并打开串口
    'strCmd= ResetExe -重启接口 CloseExe-关闭接口
    
    Dim aDevices As Variant, i As Integer, y As Integer
    Dim iData As Integer, blnNextIP As Boolean
    Dim strSQL As String, strSet As String, strTmp As String, varSet As Variant, lngID As Long, lngSaveAsID As Long
    Dim strCOM As String, rsTmp As ADODB.Recordset
    Dim aPorts As Variant, str仪器列表 As String, blnAdd As Boolean
    '清除控件
    On Error GoTo errH
    
    For i = 1 To Me.DevComm.Count - 1
        Unload Me.DevComm(i)
    Next
    Me.cboDev.Clear
    strDevIDs = ""

    ReDim g仪器(1)

    
    If gblnFromDB Then
        gblnClearData = gobjDatabase.GetPara("清空接收日志", glngSys, 1208, 1)
        '从数据库读参数
        '设置格式:  仪器id,类型,COM口,波特率,数据位,校验位,停止位,握手,TCPIP端口,IP地址,字符模式,另存为的仪器ID,主机,自动应答,可发已核标本,通讯目录,自动审核人,自动计算质控,另存为通道码
       
        strSet = Trim(gobjDatabase.GetPara("本机连接仪器", glngSys, 1208, ""))
        If strSet = "" Then
            ShowMenu
            '没有连接任何仪器
        Else
            varSet = Split(strSet, ";")
            
            ReDim g仪器(UBound(varSet))
            For i = LBound(g仪器) To UBound(g仪器)
                g仪器(i).ID = 0
                g仪器(i).IP = ""
                g仪器(i).IP端口 = 6666
                g仪器(i).SaveAsID = 0
                g仪器(i).波特率 = 9600
                g仪器(i).类型 = 0
                g仪器(i).COM口 = 0
                g仪器(i).数据位 = 0
                g仪器(i).停止位 = 0
                g仪器(i).握手 = 0
                g仪器(i).校验位 = "N"
                g仪器(i).字符模式 = 0
                g仪器(i).主机 = 0
                g仪器(i).编码名称 = ""
                g仪器(i).自动应答 = "0"
                g仪器(i).可发已核标本 = 1
                g仪器(i).通讯目录 = ""
                g仪器(i).通讯程序 = ""
                g仪器(i).自动审核人 = ""
                g仪器(i).自动计算质控 = 0
                g仪器(i).另存为通道码 = 0
            Next
            
            str仪器列表 = ""
            If gstr仪器数量 <> "" Then
                If Val(gstr仪器数量) <> 0 Then
                    Set rsTmp = GetDevices
                    Do Until rsTmp.EOF
                        str仪器列表 = str仪器列表 & "," & rsTmp!ID
                        rsTmp.MoveNext
                    Loop
                End If
            End If
            
            For i = LBound(varSet) To UBound(varSet)
                
                If varSet(i) <> "" Then
                    lngID = Val(Split(varSet(i), ",")(0))
                    If lngID > 0 Then
                        blnAdd = True

                        If gstr仪器数量 <> "" Then
                            If Val(gstr仪器数量) <> 0 Then
                                If InStr("," & str仪器列表 & ",", "," & lngID & ",") <= 0 Then
                                    blnAdd = False
                                End If
                            Else
                                blnAdd = False
                            End If
                        End If
                        
                        If blnAdd Then
                            strCOM = Split(varSet(i), ",")(1)
                            If Val(strCOM) = 0 Then
                                strTmp = "COM" & Split(varSet(i), ",")(2)
                            Else
                                strTmp = Split(varSet(i), ",")(9) & ":" & Trim(Split(varSet(i), ",")(8))
                            End If
                            
                            Set rsTmp = gobjDatabase.OpenSqlRecord("Select 编码,名称, 通讯程序名 From 检验仪器 where ID=[1]", "取检验仪器名", lngID)
                            Do Until rsTmp.EOF
                                strTmp = strTmp & " " & rsTmp!名称
                                g仪器(i).编码名称 = "(" & rsTmp!编码 & ")" & rsTmp!名称
                                g仪器(i).通讯程序 = Trim("" & rsTmp!通讯程序名)
                                rsTmp.MoveNext
                            Loop
                            If g仪器(i).编码名称 <> "" Then
                                strDevIDs = strDevIDs & "," & lngID
                                lngSaveAsID = Split(varSet(i), ",")(11)
                                Set rsTmp = gobjDatabase.OpenSqlRecord("Select 名称 From 检验仪器 where ID=[1]", "取另存检验仪器名", lngSaveAsID)
                                Do Until rsTmp.EOF
                                    strTmp = strTmp & " -> " & rsTmp!名称
                                    rsTmp.MoveNext
                                Loop
                                 If strTmp <> "" Then Me.cboDev.AddItem strTmp
                            
                                 With g仪器(i)
                                     .ID = lngID
                                     .类型 = Trim(Split(varSet(i), ",")(1))
                                     .COM口 = Trim(Split(varSet(i), ",")(2))
                                     .波特率 = Trim(Split(varSet(i), ",")(3))
                                     .数据位 = Trim(Split(varSet(i), ",")(4))
                                     .校验位 = Trim(Split(varSet(i), ",")(5))
                                     .停止位 = Trim(Split(varSet(i), ",")(6))
                                     .握手 = Trim(Split(varSet(i), ",")(7))
                                     .IP端口 = Trim(Split(varSet(i), ",")(8))
                                     .IP = Trim(Split(varSet(i), ",")(9))
                                     .字符模式 = Trim(Split(varSet(i), ",")(10))
                                     .SaveAsID = lngSaveAsID
                                     .主机 = Trim(Split(varSet(i), ",")(12))
                                     .自动应答 = Trim(Split(varSet(i), ",")(13))
                                     If UBound(Split(varSet(i), ",")) >= 14 Then
                                        .可发已核标本 = Val(Split(varSet(i), ",")(14))
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 15 Then
                                        .通讯目录 = Split(varSet(i), ",")(15)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 16 Then
                                        .自动审核人 = Split(varSet(i), ",")(16)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 17 Then
                                        .自动计算质控 = Split(varSet(i), ",")(17)
                                     End If
                                     If UBound(Split(varSet(i), ",")) >= 18 Then
                                        .另存为通道码 = Split(varSet(i), ",")(18)
                                     End If
                                 End With
                                
                                 Load Me.DevComm(Me.DevComm.Count)
                                 Me.DevComm(Me.DevComm.Count - 1).InitContrl i, strCmd
    
                            End If
                        End If '可以加
                    End If
                End If
            Next
            If Len(strDevIDs) > 0 Then strDevIDs = Mid(strDevIDs, 2)
            If Me.cboDev.ListCount > 0 Then Me.cboDev.ListIndex = 0
        End If
    Else
        gblnClearData = GetSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", 1)
        '从注册表读参数
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = LBound(aPorts) To UBound(aPorts)
                aPorts(i, 0) = "COM" & (i + 1)
            Next
        End If
        
        If IsEmpty(aPorts) Then
            ShowMenu
    '        MsgBox "没有连接任何仪器，系统将不能接收检验数据！请进入参数设置进行处理。", vbInformation, gstrSysName
        Else
            ReDim g仪器(UBound(aPorts))
            For i = LBound(g仪器) To UBound(g仪器)
                g仪器(i).ID = 0
                g仪器(i).IP = ""
                g仪器(i).IP端口 = 6666
                g仪器(i).SaveAsID = 0
                g仪器(i).波特率 = 9600
                g仪器(i).类型 = 0
                g仪器(i).COM口 = 0
                g仪器(i).数据位 = 0
                g仪器(i).停止位 = 0
                g仪器(i).握手 = 0
                g仪器(i).校验位 = "N"
                g仪器(i).字符模式 = 0
                g仪器(i).主机 = 0
                g仪器(i).编码名称 = ""
                g仪器(i).自动应答 = "0"
                g仪器(i).可发已核标本 = "1"
                g仪器(i).通讯目录 = ""
                g仪器(i).通讯程序 = ""
                g仪器(i).自动审核人 = ""
                g仪器(i).自动计算质控 = 0
                g仪器(i).另存为通道码 = 0
            Next
            
            For i = LBound(aPorts) To UBound(aPorts)
                lngID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                If lngID > 0 Then
                    
                    strCOM = IIf(aPorts(i, 0) Like "COM*", 0, 1)
                    If strCOM = 0 Then
                    
                        strTmp = aPorts(i, 0)
                    Else
                        strTmp = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1") & ":" & _
                                 GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Port", "6666")
                        
                    End If
                
                    Set rsTmp = gobjDatabase.OpenSqlRecord("Select 编码,名称,通讯程序名 From 检验仪器 where ID=[1]", "取检验仪器名", lngID)
                    Do Until rsTmp.EOF
                        strTmp = strTmp & " " & rsTmp!名称
                        g仪器(i).编码名称 = "(" & rsTmp!编码 & ")" & rsTmp!名称
                        g仪器(i).通讯程序 = Trim("" & rsTmp!通讯程序名)
                        rsTmp.MoveNext
                    Loop
                    
                    If g仪器(i).编码名称 <> "" Then
                        strDevIDs = strDevIDs & "," & lngID
                        lngSaveAsID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "SaveAs", "0"))
                        
                        Set rsTmp = gobjDatabase.OpenSqlRecord("Select 名称 From 检验仪器 where ID=[1]", "取另存检验仪器名", lngSaveAsID)
                        Do Until rsTmp.EOF
                            strTmp = strTmp & " -> " & rsTmp!名称
                            rsTmp.MoveNext
                        Loop
                        If strTmp <> "" Then Me.cboDev.AddItem strTmp
                    
                        With g仪器(i)
                            .ID = lngID
                            .类型 = strCOM
                            .COM口 = IIf(strCOM = 0, Replace(aPorts(i, 0), "COM", ""), "0")
                            .波特率 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                            .数据位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                            .校验位 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Parity", "N")
                            .停止位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                            .握手 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "HandShaking", "0"))
                            .IP端口 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Port", "6666"))
                            .IP = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            .字符模式 = IIf(strCOM = 0, Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")), _
                                        Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InMode", "0")))
                            .SaveAsID = lngSaveAsID
                            .主机 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Host", "0"))
                            .自动应答 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            .可发已核标本 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            .通讯目录 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "ReceiveDir", "")
                            .自动审核人 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "AutoCheckMan", "")
                            .自动计算质控 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "AutoQCCalc", 0))
                            .另存为通道码 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "SaveAsTonDao", 0))
                        End With
                        Load Me.DevComm(Me.DevComm.Count)
                        Me.DevComm(Me.DevComm.Count - 1).InitContrl i, strCmd
                    End If
                End If
            Next
            If Len(strDevIDs) > 0 Then strDevIDs = Mid(strDevIDs, 2)
            If Me.cboDev.ListCount > 0 Then Me.cboDev.ListIndex = 0
        End If
        
    End If ''是否从数据库读取参数
    ReadPara = True
    Exit Function
errH:
    MsgBox Err.Description
    
End Function

Private Sub ShowMenu()
    Dim blnEnabled As Boolean
    If Me.cboDev.ListIndex = -1 Then
        Me.mnuFileOpen.Enabled = False
        Me.mnuFileClose.Enabled = False
        Me.tbrMain.Buttons("连接").Enabled = False
        Me.tbrMain.Buttons("断开").Enabled = False
    Else
        blnEnabled = Me.DevComm(cboDev.ListIndex + 1).PortOpened
        
        Me.mnuFileOpen.Enabled = Not blnEnabled
        Me.mnuFileClose.Enabled = blnEnabled
        Me.tbrMain.Buttons("连接").Enabled = Not blnEnabled
        Me.tbrMain.Buttons("断开").Enabled = blnEnabled
    End If
    Me.mnuDevExp.Enabled = mblnOwner
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal lngSampleNO As Long, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
'发送标本记录到仪器
    Dim i As Integer
    SendSample = True
    
    For i = 1 To DevComm.UBound
        If DevComm(i).DeviceID = lngDeviceID Then
            SendSample = DevComm(i).SendSample(lngDeviceID, strSampleDate, lngSampleNO, strAdviceIDs, blnUndo, iType)
            Exit For
        End If
    Next
End Function

Private Sub timConn_Timer()
    Dim dateNow As Date, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select 1 From dual"
    
    Set rsTmp = gcnOracle.Execute(strSQL)
    
    Exit Sub
errH:
    WriteLog "数据库连接断开", LOG_错误日志, Err.Number, Err.Description
    i = 0
    Do While i <= 30
        If Err.Number <> 0 Then
            Err.Clear
            If gcnOracle.State = 1 Then gcnOracle.Close
            gcnOracle.Open mstrConn
        Else
            WriteLog "数据库连接已恢复", LOG_错误日志, 0, "本次重试次数=" & i
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

Private Sub timLsn_Timer()
    Call mnuReLoad_Click
End Sub

Private Sub WinsockS_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim aItem() As String
    On Error Resume Next
    With Me.WinsockS
        .GetData strData
    End With
    If Len(Trim(strData)) = 0 Then Exit Sub
    aItem = Split(strData, ",")
    If UBound(aItem) <= 0 Then Exit Sub
    If aItem(1) <> Me.WinsockS.LocalIP Then Exit Sub            '不是同一IP时退出
    Select Case aItem(0)
        Case "SendSample"
            SendSample aItem(2), aItem(3), aItem(4), IIf(aItem(5) = "", "", _
            Replace(aItem(5), ";", ",")), IIf(aItem(6) = "", "", aItem(6)), IIf(aItem(7) = "", "", aItem(7))
        Case "ResultFromFile"
            ResultFromFile aItem(2), aItem(3), aItem(4), aItem(5), aItem(6)
    End Select
    '返回操作完成时
    With Me.WinsockS
        .SendData .LocalIP & ";" & strSend_True
    End With
End Sub

Private Function ExpLisDevData() As String
    '导出检验仪器的数据
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    Dim strPath As String, strFileName As String
    Dim strLog As String '记录导出步骤
    
    Dim strTable As String, strFiled As String
    
    On Error GoTo errH
    
    
    
    strTable = "检验项目"
    strFiled = "临床意义,VARCHAR2(4000)|隐私项目,NUMBER(1)"
    If Not CheckFiled(strTable, strFiled) Then Exit Function
    
    strTable = "检验仪器项目"
    strFiled = "糖耐量项目,NUMBER(1)"
    If Not CheckFiled(strTable, strFiled) Then Exit Function
    
    
    '------------------------------------------------------------------------------
    strFileName = App.Path & "\zlLis基础数据_" & Format(date, "yyyyMMdd") & ".txt"
    
    If objFileSystem.FileExists(strFileName) Then Kill strFileName
    Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    objStream.WriteLine "[用户名称]"
    objStream.WriteLine gstr单位名称
    
    strSQL = "Select ID||Chr(9)||编码||Chr(9)||名称||Chr(9)||通讯程序名||Chr(9)||仪器类型 As Line From 检验仪器 Where 微生物 <> 1"
    Call WritData("[仪器]", strSQL, objStream)
    
    strSQL = "Select Distinct B.项目id||Chr(9)||D.名称||Chr(9)||A.缩写||Chr(9)||A.项目类别||Chr(9)||A.结果类型||Chr(9)||D.适用性别||Chr(9)||D.操作类型||Chr(9)||D.标本部位||Chr(9)||A.单位||Chr(9)||A.默认值||Chr(9)||" & vbNewLine & _
            "                A.取值序列||Chr(9)||A.隐私项目||Chr(9)||A.阳性公式||Chr(9)||A.弱阳性公式||Chr(9)||A.Cutoff公式||Chr(9)||A.计算公式||Chr(9)||A.检验方法||Chr(9)||A.临床意义 As Line" & vbNewLine & _
            "From 诊疗项目目录 D, 检验报告项目 C, 检验项目 A, 检验仪器项目 B, 检验仪器 E" & vbNewLine & _
            "Where D.组合项目 <> 1 And C.诊疗项目id = D.ID And C.报告项目id = B.项目id And A.诊治项目id = B.项目id And E.ID = B.仪器id And E.微生物 <> 1"
    Call WritData("[项目]", strSQL, objStream)

    strSQL = "Select Distinct A.项目id||Chr(9)||A.标本类型||Chr(9)||A.性别域||Chr(9)||A.年龄下限||Chr(9)||A.年龄上限||Chr(9)||A.年龄单位||Chr(9)||A.参考低值||Chr(9)||A.参考高值||Chr(9)||A.临床特征 As Line" & vbNewLine & _
            "From 检验项目参考 A, 检验仪器项目 B, 检验仪器 E" & vbNewLine & _
            "Where A.项目id = B.项目id And (Nvl(A.年龄下限, 0) <> 0 Or Nvl(A.年龄上限, 0) <> 0 Or Nvl(A.参考低值, 0) <> 0 Or Nvl(A.参考高值, 0) <> 0) And" & vbNewLine & _
            "      E.ID = B.仪器id And E.微生物 <> 1"
    Call WritData("[项目参考]", strSQL, objStream)
    
    strSQL = "Select 仪器id||Chr(9)||项目id||Chr(9)||通道编码||Chr(9)||小数位数||Chr(9)||糖耐量项目 as Line" & vbNewLine & _
            "From 检验仪器项目 A, 检验仪器 E" & vbNewLine & _
            "Where E.ID = A.仪器id And E.微生物 <> 1"
    Call WritData("[仪器项目]", strSQL, objStream)
    
    '微生物--字典管理数据
    
    strSQL = "select 编码||chr(9)||名称||chr(9)||简码 as Line from 检验细菌菌属"
    Call WritData("[检验细菌菌属]", strSQL, objStream)
    
    strSQL = "select 编码||chr(9)||名称||chr(9)||简码||chr(9)||缺省标志 as Line from 检验细菌类别"
    Call WritData("[检验细菌类别]", strSQL, objStream)
    
    strSQL = "select 编码||chr(9)||名称||chr(9)||简码||chr(9)||缺省标志 as Line from 革兰染色分类"
    Call WritData("[革兰染色分类]", strSQL, objStream)
    
    strSQL = "select 编码||chr(9)||名称||chr(9)||简码 as Line from 细菌检测方法"
    Call WritData("[细菌检测方法]", strSQL, objStream)
    
    '微生物--数据
    strSQL = "select id||chr(9)||编码||chr(9)||名称||chr(9)||英文||chr(9)||简码 as Line from 检验抗生素组"
    Call WritData("[检验抗生素组]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||编码||chr(9)||中文名||chr(9)||英文名||chr(9)||简码||chr(9)||说明||chr(9)||药敏方法||chr(9)||whonet码||chr(9)||用法用量1||chr(9)||血药浓度1||chr(9)||尿药浓度1||chr(9)||用法用量2||chr(9)||血药浓度2||chr(9)||尿药浓度2 as Line from 检验用抗生素"
    Call WritData("[检验用抗生素]", strSQL, objStream)
    
    strSQL = "select 抗生素id||chr(9)||抗生素分组id as Line From 检验抗生素用药"
    Call WritData("[检验抗生素用药]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||编码||chr(9)||中文名称||chr(9)||英文名称||chr(9)||简码 as Line from 检验细菌类型"
    Call WritData("[检验细菌类型]", strSQL, objStream)
    
    strSQL = "select id||chr(9)||编码||chr(9)||中文名||chr(9)||英文名||chr(9)||类型id||chr(9)||简码||chr(9)||默认药敏||chr(9)||默认方法||chr(9)||whonet码||chr(9)||默认结果||chr(9)||细菌类别||chr(9)||细菌菌属||chr(9)||革兰氏分类 as Line from 检验细菌"
    Call WritData("[检验细菌]", strSQL, objStream)
    
    strSQL = "select 细菌id||chr(9)||抗生素分组id||chr(9)||缺省标志 as Line From 检验细菌抗生素"
    Call WritData("[检验细菌抗生素]", strSQL, objStream)
    
    strSQL = "select 细菌id||chr(9)||抗生素分组id||chr(9)||抗生素id||chr(9)||药敏方法||chr(9)||参考低值||chr(9)||参考高值||chr(9)||判断方式||chr(9)||备注 as Line From 检验细菌抗生素参考"
    Call WritData("[检验细菌抗生素参考]", strSQL, objStream)
    
    '微生物--仪器及仪器项目对照
    strSQL = "Select ID||Chr(9)||编码||Chr(9)||名称||Chr(9)||通讯程序名||Chr(9)||仪器类型 As Line  From 检验仪器 Where 微生物=1"
    Call WritData("[微生物仪器]", strSQL, objStream)
    
    strSQL = "Select 仪器id||Chr(9)||通道编码||Chr(9)||细菌id||Chr(9)||抗生素id as Line From 仪器细菌对照"
    Call WritData("[仪器细菌对照]", strSQL, objStream)
    
    objStream.Close
    Set objStream = Nothing
    ExpLisDevData = strFileName
    Exit Function
errH:
    MsgBox "导出数据时出现错误：" & Err.Description & vbCrLf & strLog & vbCrLf & strSQL

End Function

Private Sub WritData(ByVal strHead As String, ByVal str_Sql As String, objStream As TextStream)
    Dim rsTmp As ADODB.Recordset
    
    If str_Sql = "" Then Exit Sub
    If objStream Is Nothing Then Exit Sub
    If InStr(UCase(str_Sql), UCase(" as Line")) <= 0 Then Exit Sub
    
    objStream.WriteLine strHead
    Set rsTmp = gobjDatabase.OpenSqlRecord(str_Sql, Me.Caption)
    With rsTmp
        Do Until .EOF
            objStream.WriteLine "" & !Line
            .MoveNext
        Loop
    End With
    objStream.WriteLine ""
    
End Sub

Private Function CheckFiled(ByVal strTable As String, ByVal strFileds As String) As Boolean
    '检查数据结构，差的话就加。
    
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Integer
    Dim strName As String, strTypeLen As String
    Dim varFiled As Variant
    strSQL = "Select Data_Type As 类型, Data_Precision As 整数, Data_Scale As 小数, Data_Length As 长度" & vbNewLine & _
            "From User_Tab_Columns" & vbNewLine & _
            "Where Table_Name = [1] And Column_Name = [2]"
    
    varFiled = Split(strFileds, "|")
    
    strSQL = "Select upper(Column_Name) as 字段名 " & vbNewLine & _
    "From User_Tab_Columns" & vbNewLine & _
    "Where Table_Name = [1]"
    Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, Me.Caption, strTable)
    If rsTmp.RecordCount <= 0 Then
        MsgBox "缺少表“" & strTable & "”,不能导出数据!"
        Exit Function
    End If
    
    For i = LBound(varFiled) To UBound(varFiled)
        strName = UCase(Split(varFiled(i), ",")(0))
        strTypeLen = UCase(Split(varFiled(i), ",")(1))
        rsTmp.Filter = "字段名= '" & strName & "'"
        If rsTmp.EOF Then
            strSQL = "Alter Table " & strTable & " Add " & strName & " " & strTypeLen
            gcnOracle.Execute strSQL
        End If
    Next

    CheckFiled = True
    
End Function
