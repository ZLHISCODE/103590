VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSGate 
   AutoRedraw      =   -1  'True
   Caption         =   "影像接收服务"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmPACSGate.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraUD_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   270
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   3480
      Width           =   7635
   End
   Begin VB.Timer Timer1 
      Left            =   5010
      Top             =   3270
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "fgfg"
         Object.Width           =   2540
      EndProperty
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
            Picture         =   "frmPACSGate.frx":058A
            Key             =   "_0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":0B24
            Key             =   "_1"
         EndProperty
      EndProperty
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
            Picture         =   "frmPACSGate.frx":10BE
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
      TabIndex        =   1
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
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1952
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1B6C
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1D86
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1FA0
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":21BA
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":28B4
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":2FAE
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":36A8
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":3DA2
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":449C
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":4B96
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":5290
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":598A
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":6084
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":677E
            Key             =   "作废"
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
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":6E78
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":7092
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":72AC
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":74C6
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":76E0
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":7DDA
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":84D4
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":8BCE
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":92C8
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":99C2
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":A0BC
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":A7B6
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":AEB0
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":B5AA
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":BCA4
            Key             =   "作废"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picView 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   8235
      TabIndex        =   4
      Top             =   3720
      Width           =   8295
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2055
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   3735
         _Version        =   262147
         _ExtentX        =   6588
         _ExtentY        =   3625
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1191
      ButtonWidth     =   820
      ButtonHeight    =   1138
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgGray"
      HotImageList    =   "imgColor"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
            Caption         =   "帮助"
            Key             =   "帮助"
            Object.ToolTipText     =   "当前帮助主题"
            Object.Tag             =   "帮助"
            ImageKey        =   "帮助"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "退出"
            Object.ToolTipText     =   "退出"
            Object.Tag             =   "退出"
            ImageKey        =   "退出"
         EndProperty
      EndProperty
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
      Begin VB.Menu mmuProcessLog 
         Caption         =   "记录处理日志"
      End
      Begin VB.Menu mmuCommLog 
         Caption         =   "记录通讯日志"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示通讯日志"
         Index           =   1
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示错误日志"
         Index           =   2
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "显示当前服务"
         Index           =   3
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
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
      Begin VB.Menu mmuUpdateDB 
         Caption         =   "升级数据库(&U)"
      End
      Begin VB.Menu mnuHelp_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmPACSGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public LastState As Integer
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mstrPrivs As String         '权限字串
Private mBufferDir As String
Private lngErrCounts As Long
Private mstrMWLModality As String            'worklist通讯中使用的影像类别
Private mintWLCount As Integer               'worklist通讯中用于自动递增的记数器

Private strWhere As String
Private blnNewImg As Boolean '有新的影像，需要刷新列表
Private strDirURL As String, strHost As String

Private mdtLastAssociation As Date           '最近接收到Association的时间

'打印路由
Dim DGlobal As DicomGlobal
Dim PrintRouterDss As DicomDataSets
Dim printerobject As DicomDataSet

Private WithEvents mobjIcon As clsTaskIcon  '托盘类
Attribute mobjIcon.VB_VarHelpID = -1
Private mfrmUpdateDB As frmUpdateDB
Private mfrmShowLog As frmShowLog

Private Sub DViewer_AssociationClosed(ByVal connection As DicomObjects.DicomConnection)
    Dim Session As DicomDataSet
    
    For Each Session In connection.Tag
        subRemove Session.instanceUID
    Next
End Sub

Private Sub DViewer_AssociationRequest2(ByVal connection As DicomObjects.DicomConnection, isOK As Boolean)
    Dim context As DicomContext
    Dim strLog As String, strTmp As String
    Dim i As Integer
    Dim blnMatch As Boolean     '是否跟允许的服务对匹配
    Dim blnNext As Boolean
    
    On Error GoTo ProcError
    
    '记录接收到连接请求的时间
    mdtLastAssociation = Time
    
    '处理打印路由，预先给connection的tag创建一个空数据集
    Set connection.Tag = New DicomDataSets
    
    '当数据断开时重新连接
    CheckDBConnect
    
    strLog = "请求的AE是：" & connection.CallingAET & _
        ",请求的IP是：" & connection.RemoteIP & ",被呼叫的AE是：" & connection.CalledAET
    WriteCommLog "AssociationRequest", "接收到通讯请求", strLog
    
    '判断该请求是否是服务对里面允许的
    '对于“打印路由”和“胶片接收”请求，只判断CalledAE
    '对于图像接收、worklist、Q/R请求，判断设备IP,设备AE,服务AE。
    
    For i = 1 To UBound(Services)
        If UCase(Services(i).ServiceAE) = UCase(connection.CalledAET) _
            And (UCase(Services(i).SOP) = "PRINT" Or Services(i).SOP = "胶片接收") Then
            blnMatch = True
            Exit For
        ElseIf Services(i).DeviceIP = connection.RemoteIP And UCase(Services(i).DeviceAE) = UCase(connection.CallingAET) _
            And UCase(Services(i).ServiceAE) = UCase(connection.CalledAET) Then
            blnMatch = True
            Exit For
        End If
    Next i
    
    If blnMatch = False Then
        isOK = False
        strLog = "请求的服务跟本机支持的服务不匹配，请检查AE名称，请求被拒绝。请求的AE是：" & connection.CallingAET & _
        ",请求的IP是：" & connection.RemoteIP & ",被呼叫的AE是：" & connection.CalledAET
        WriteCommLog "AssociationRequest", "拒绝不被允许的服务", strLog
        Exit Sub
    End If
        
    '拒绝所有非DICOM SOP 服务类的请求
    strLog = ""
    For Each context In connection.Contexts
        If Left(context.AbstractSyntax, 14) <> "1.2.840.10008." Then
'            context.Reject 3
'            WriteCommLog "AssociationRequest", "请求被拒绝", "请求的语法为：" & _
'                context.AbstractSyntax & ",不属于1.2.840.10008类"
        Else
            strLog = strLog & ": " & context.AbstractSyntax
            '保存图像存储请求的Association
            '只判断第一个上下文
            '保存除了Q/R，Worklist，Verify之外的连接，这种连接应该就是图像存储连接
            'Q/R的抽象语法为 “1.2.840.10008.5.1.4.1.2.x.x”
            'Worklist的抽象语法为 “1.2.840.10008.5.1.4.31”
            'Verify的抽象语法为 “1.2.840.10008.1.1”
            'Print 中 BasicGrayScalePrint 的抽象语法为 “1.2.840.10008.5.1.1.9”
            If blnNext = False Then
                If context.AbstractSyntax <> "1.2.840.10008.5.1.4.31" And context.AbstractSyntax <> "1.2.840.10008.1.1" _
                    And Left(context.AbstractSyntax, 24) <> "1.2.840.10008.5.1.4.1.2." Then
                    '保存连接参数
                    subSaveAssociation connection
                    blnNext = True
                End If
            End If
        End If
    Next context
    If strLog <> "" Then
        WriteCommLog "AssociationRequest", "请求被允许", "允许请求的语法为：" & strLog
    End If
    
    Exit Sub
ProcError:
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Sub DViewer_ImageReceived(ByVal ReceivedImage As DicomObjects.DicomImage, ByVal Association As Long, isAdded As Boolean, Status As Long)
    Dim blnReceived As Boolean
    
    On Error GoTo ProcError
    
    '记录接收到图像的时间
    mdtLastAssociation = Time
    
    ReceivedImage.Tag = Association
    '当数据断开时重新连接
    CheckDBConnect
    If ImageExist(DViewer.Images, ReceivedImage) Then
        isAdded = False: Status = 0
        Exit Sub
    End If
    
    blnReceived = True
    '有一类图像需要特殊处理。飞利浦Intera MR的一种图像，因为无法解析所以不接收
    If Not IsNull(ReceivedImage.Attributes(&H8, &H60).value) Then
        If UCase(ReceivedImage.Attributes(&H8, &H60).value) = "MR" And Not IsNull(ReceivedImage.Attributes(&H8, &H16).value) Then
            If Left(ReceivedImage.Attributes(&H8, &H16).value, Len(ReceivedImage.Attributes(&H8, &H16).value) - 1) = "1.3.46.670589.11.0.0.12." Or _
                ReceivedImage.Attributes(&H8, &H16).value = "1.2.840.10008.5.1.4.1.1.66" Then
                  '类型为MR的，经测试得知，如果Sop Class UID ="1.3.46.670589.11.0.0.12.2"或"1.3.46.670589.11.0.0.12.4" ，则也不做任何处理
                  '还可能有其他的SOP ClassUID,因此判断前缀“1.3.46.670589.11.0.0.12.xxx”
                  blnReceived = False
            End If
        End If
    End If
    
    If blnReceived = False Then
        isAdded = False: Status = 0
        Exit Sub
    End If
    
    DViewer.Images.Add ReceivedImage
    DoEvents
    isAdded = False: Status = 0: blnNewImg = True
    ProcSave

    Me.stbThis.Panels(2).Text = "最近接收到的图像信息：病人－" & NVL(ReceivedImage.Name) & " 检查时间－" & Time
    WriteCommLog "DViewer_ImageReceived", "接收到图像", Me.stbThis.Panels(2).Text
    Exit Sub
ProcError:
    On Error Resume Next
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Sub DViewer_NormalisedReceived(ByVal connection As DicomObjects.DicomConnection)
    Dim command As DicomDataSet, ds As DicomDataSet, a As DicomAttribute
    Dim operation As Integer, rclass As String, ruid As String, aclass As String, auid As String
    Dim dss As DicomDataSets, ds1 As DicomDataSet, ds2 As DicomDataSet, i As Integer
    Dim sessionUID As String
    Dim lngPrintStatus As Long
    
    
    Set command = connection.command
    
    '重点参考DICOM标准第七章
    operation = command.Attributes(0, &H100)    'Command Field 针对每一个消息都有一个特定值
    rclass = command.Attributes(0, 3) & ""      'Requested SOP Class UID
    ruid = command.Attributes(0, &H1001) & ""   'Requested SOP Instance UID
    aclass = command.Attributes(0, 2) & ""      'Affected SOP Class UID
    auid = command.Attributes(0, &H1000) & ""   'Affected SOP Instance UID
    
    On Error Resume Next
    Select Case operation
    Case &H110  'N-GET      响应一个N-GET请求，返回所请求的数据集
        Set ds = funGetDataset(ruid)
        connection.SendData ds, 0
        connection.SendStatus 0
    
    Case &H140  'N-CREATE   响应一个N-CREATE请求，Affected SOP Class UID是需要创建的SOP Instance 的Class UID
                'Affected Instance UID 是需要创建的SOP Instance 的Instance UID
                'Connection.Request.Attributes是N-CREATE请求里面附带的数据集
                '当对方Open 的时候，先接收到FilmSession和它的参数
                '当对方PrintImage的时候，接收到FilmBox和它的参数
                
        Set ds = NewDataSet         'ds就是本次N-CREAT创建的主数据集
        '设置默认值，使用请求中的数据集复制一份
        For Each a In connection.request.Attributes
            ds.Attributes.Add a.group, a.element, a.value
        Next
    
        ds.Attributes.Add 8, &H16, aclass       'SOP Class UID ,就是请求的Class UID
        If auid = "" Then auid = DGlobal.NewUID
        ds.Attributes.Add 8, &H18, auid         'SOP Instance UID,就是请求的Instance UID
        
        '处理Film Box，处理doSOP_BasicFilmBox类型的N-CREATE请求，创建FilmBox
        If aclass = doSOP_BasicFilmBox Then
            
            '检查Image boxs  的数量，然后创建他们
            '从Image Display Format中解析图像的布局,计算出图像数量
            Dim intImgNum As Integer
            DecodeFormat ds.Attributes(&H2010, &H10), intImgNum
            
            
            '根据图像数字量，创建基本灰度的ImageBox
            'dss为ds中附加的数据集，保存在它的Tag中，在前面创建ds的时候，已经给其TAG赋值成了DicomDataSets
            Set dss = ds.Tag
            For i = 1 To intImgNum
                '创建ImageBox的数据集，通过NewDataSet保存到PrintRouterDss集合里面
                Set ds1 = NewDataSet
                ds1.instanceUID = DGlobal.NewUID                            'instanceUID以后作为索引调用
                ds1.Attributes.Add 8, &H1155, ds1.instanceUID               'Referenced SOP Instance UID
                ds1.Attributes.Add 8, &H1150, doSOP_BasicGrayscaleImageBox  'Referenced SOP Class UID
                dss.Add ds1
            Next i
            '把ImageBox的内容添加到ds中
            ds.Attributes.Add &H2010, &H510, dss        'Referenced Image Box Sequence 添加Image Box序列
            
            '关联到session
            Dim SessionSeq As DicomDataSets
            Set SessionSeq = ds.Attributes(&H2010, &H500).value     'Referenced Film Session Sequence,指向Session 序列
            sessionUID = SessionSeq(1).Attributes(8, &H1155)        'Referenced SOP Instance UID
            
            PrintRouterDss(sessionUID).Tag.Add ds                   '这个TAG里面指向的是一个DicomDataSets
        End If
        
        '处理session ,跟connecion相连接
        If aclass = doSOP_BasicFilmSession Then             '处理doSOP_BasicFilmSession类型的N-CREATE请求
            connection.Tag.Add ds
        End If
        
        connection.SendData ds, 0
    
    Case &H130  '响应N-ACTION请求，开始打印Instance UID对应的内容
        
        '执行打印操作-------------------------------------
        '判断当前AE是打印路由，还是胶片接收
        
        lngPrintStatus = funPrintOut(ruid, connection)
        connection.SendStatus lngPrintStatus
    
    Case &H150  '响应N-DELETE请求，删除请求的Instance UID
        subRemove ruid
        connection.SendStatus 0
    Case &H120  '响应一个N-SET请求，设置对应的SOP Class和Instance UID指定的数据集，实际上是接收图像
        
        Set ds = funGetDataset(ruid)
        For Each a In connection.request.Attributes
            ds.Attributes.Add a.group, a.element, a.value
        Next
        connection.SendStatus 0
    End Select
End Sub

Private Sub DViewer_VerifyReceived(Status As Long)
    On Error GoTo ProcError
    Status = 0
    Exit Sub
ProcError:
    On Error Resume Next
'    Status = err.Number
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(1, err.Number, err.Description)
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

Private Sub fraUD_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    fraUD_s.BackColor = IIf(Y > 0, vbWhite, RGB(0, 0, 0))
    On Error Resume Next
    If fraUD_s.Top + Y < 2000 Then
        fraUD_s.Top = 2000
    ElseIf Me.ScaleHeight - fraUD_s.Top - Y < 4000 Then
        fraUD_s.Top = Me.ScaleHeight - 4000
    Else
        fraUD_s.Top = fraUD_s.Top + Y
    End If
End Sub

Private Sub fraUD_s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub

    fraUD_s.BackColor = Me.BackColor
    Form_Resize
End Sub


Private Sub lvwSeq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwSeq, ColumnHeader.Index)
End Sub

Private Sub mmuCommLog_Click()
    mmuCommLog.Checked = Not mmuCommLog.Checked
End Sub



Private Sub mmuProcessLog_Click()
    mmuProcessLog.Checked = Not mmuProcessLog.Checked
    gblnProcessLog = mmuProcessLog.Checked
End Sub

Private Sub mmuShowLog_Click(Index As Integer)
    Set mfrmShowLog = New frmShowLog
    mfrmShowLog.intLogType = Index
    mfrmShowLog.Show 1, Me
    Set mfrmShowLog = Nothing
End Sub

Private Sub mmuShowService_Click()
        
End Sub

Private Sub mmuUpdateDB_Click()
    Set mfrmUpdateDB = New frmUpdateDB
    Set mfrmUpdateDB.m_cnAccess = gcnAccess
    mfrmUpdateDB.Show 1, Me
    Set mfrmUpdateDB = Nothing
End Sub

Private Sub subListenPorts(iType As Integer)
'-----------------------------------------------------------------------------
'功能:启动或停止对服务端口的侦听
'参数: iType = 0 停止侦听；iType = 1 启动侦听。
'修改人:黄捷
'修改日期:2007-11-30
'-----------------------------------------------------------------------------
    Dim strPort As String
    Dim i As Integer
    
    '启动本机服务端口
    For i = 1 To UBound(Services)
        If InStr(strPort, Services(i).ServicePort) = 0 Then
            strPort = strPort & "," & Services(i).ServicePort
            If iType = 0 Then   '停止侦听
                DViewer.Unlisten Val(Services(i).ServicePort)
                Services(i).Started = False
            ElseIf iType = 1 Then   '启动侦听
                If Not DViewer.Listen(Val(Services(i).ServicePort)) Then
                    Services(i).Started = False
                    MsgBox "端口：" & Services(i).ServicePort & "已被使用，" & _
                    "系统无法监听！请重新设置监听端口。", vbExclamation, gstrSysName
                Else
                    Services(i).Started = True
                End If
            End If
        Else
            Services(i).Started = IIf(iType = 0, False, True)
        End If
    Next i
End Sub


Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
    
        ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "退出"
            Me.WindowState = vbMinimized ' mnuFileQuit_Click
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
    Call zlPrintSet
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
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset

    On Error Resume Next
    
    If WindowState = vbMinimized Then
        LastState = vbNormal
        Me.Hide
    Else
        LastState = WindowState
    End If

    
    '----------加载托盘图标
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "ZLPACS服务网关"
    mobjIcon.AddIcon
    '----------加载托盘图标
    
    mstrPrivs = gstrPrivs
    
    With lvwSeq
        With .ColumnHeaders
            .Clear
        
            .Add , , "影像类别", 1000
            .Add , , "检查号", 800, 1
            .Add , , "检查设备", 1500
            .Add , , "姓名", 1000
            .Add , , "英文名", 1000
            .Add , , "性别", 600
            .Add , , "年龄", 600, 1
            .Add , , "影像数", 800, 1
            .Add , , "接收时间", 2000
            .Add , , "检查UID", 2700
            .Add , , "序列UID", 3000
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    
    '初始化本地参数
    ReDim AEconnections(0) As AEconnection
    
    gstrAccessPath = App.Path & "\ZlPacsLog"
    gstrAccessName = gstrAccessPath & ".mdb"
    '建立到本地Access（日志记录）的连接
    With gcnAccess
        .ConnectionString = "DBQ=" & gstrAccessName & ";DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then MsgBox "不能打开本地日志文件，系统将无法记录接收过程！", vbInformation, gstrSysName
    End With
    
    strBeginDate = Format(Date & " " & Time, "yyyy-MM-dd hh:mm:ss")
    
    '给DGlobal赋初值
    Set DGlobal = New DicomGlobal
    '给PrintRouterDss集合赋初值
    Set PrintRouterDss = New DicomDataSets
    
    '创建打印数据集
    MakePrinterdataset
    
    If Not ReadPara Then Unload Me: Exit Sub
    
    
    '启动侦听服务端口
    Call subListenPorts(1)
    
    lngErrCounts = 0
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    
    strWhere = ""
    ListSeq strWhere
    
    Me.WindowState = vbMinimized
    blnNewImg = False
    
    If funCanStartServer = False Then
        Unload Me
    End If
    
    gblnProcessLog = False
    
    '记录网关启动
    Call WriteLog(5802, 5802, "网关启动，网关版本为：" & App.Major & "." & App.Minor & "." & App.Revision)
    
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click()
    
    mnuViewToolItem.Checked = Not mnuViewToolItem.Checked
    tbrMain.Visible = mnuViewToolItem.Checked
    tbrMain.Enabled = tbrMain.Visible
    mnuViewToolText.Enabled = tbrMain.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    
    For i = 1 To tbrMain.Buttons.Count
        tbrMain.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrMain.Buttons(i).Tag, "")
    Next i
    If mnuViewToolText.Checked Then
        tbrMain.TextAlignment = tbrTextAlignBottom
    End If
    tbrMain.Refresh
    Form_Resize
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long

    On Error Resume Next
    
    If WindowState = 1 Then
        Me.Hide
        Exit Sub
    End If
    cbrH = IIf(tbrMain.Visible, tbrMain.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With Me.fraUD_s
        If .Top > Me.ScaleHeight Then .Top = cbrH + (Me.ScaleHeight - cbrH) / 2
        .Left = 0: .Width = Me.ScaleWidth
    End With
    
    With picView
        .Left = 0: .Top = fraUD_s.Top + fraUD_s.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - staH - .Top
    End With
    
    With lvwSeq
        .Left = 0
        .Top = cbrH
        .Height = fraUD_s.Top - .Top
        .Width = Me.ScaleWidth
    End With
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call WriteLog(5803, 5803, "网关关闭。")
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    Call SaveWinState(Me, App.ProductName)
    '停止侦听服务端口
    Call subListenPorts(0)
    '清除托盘图标
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
    '如果更新数据库和显示日志的模式窗口已经被打开，则退出，避免出现错误
    If mfrmUpdateDB Is Nothing And mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            Me.Hide
        Else
            WindowState = vbNormal
            Me.Show
        End If
    End If
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mobjIcon.MouseState X
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能: 输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrintLvw

    On Error Resume Next
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    
    Set objOut.Body.objData = Me.lvwSeq
    objOut.Title.Text = "影像序列"
    objOut.UnderAppItems.Add ""
    objOut.UnderAppItems.Add "时间：" & strBeginDate & " - " & Format(Date & " " & Time, "yyyy-MM-dd HH:mm:SS")
    If bytStyle = 1 Then
        bytStyle = zlPrintAsk(objOut)
        If bytStyle <> 0 Then zlPrintOrViewLvw objOut, bytStyle
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub ListSeq(ByVal strWhere As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strCurKey As String
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer

    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    If gcnAccess.State = adStateOpen Then
        gstrSQL = "Select 影像类别,检查号,检查设备,姓名,英文名,性别,年龄," & _
            " 影像数,接收时间,对应检查,检查UID,序列UID,ID" & _
            " From 影像接收序列 Where " & _
            IIf(strWhere = "", "接收时间>cDate('" & _
            strBeginDate & "')", strWhere) & _
            " Order By 接收时间 Desc"
        Set rsTmp = gcnAccess.Execute(gstrSQL)
        
        Me.lvwSeq.ListItems.Clear
        Do While Not rsTmp.EOF
            i = i + 1
            If i > 500 Then Exit Do
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("ID"), NVL(rsTmp("影像类别")))
            With tmpItem
                .SubItems(1) = NVL(rsTmp("检查号"))
                .SubItems(2) = NVL(rsTmp("检查设备"))
                .SubItems(3) = NVL(rsTmp("姓名"))
                .SubItems(4) = NVL(rsTmp("英文名"))
                .SubItems(5) = NVL(rsTmp("性别"))
                .SubItems(6) = NVL(rsTmp("年龄"))
                .SubItems(7) = NVL(rsTmp("影像数"))
                .SubItems(8) = NVL(rsTmp("接收时间"), Date)
                .SubItems(9) = NVL(rsTmp("检查UID"))
                .SubItems(10) = NVL(rsTmp("序列UID"))
                
                .SmallIcon = "_" & IIf(NVL(rsTmp("对应检查"), 1), 0, 1)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
DBError:
'    If ErrCenter() = 1 Then Resume
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(2, err.Number, err.Description)
End Sub

Private Sub ProcSave()
    On Error GoTo ProcError
    If DViewer.Images.Count > 0 Then
        SaveImages DViewer.Images, mBufferDir
    End If
    Exit Sub
ProcError:
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(0, err.Number, err.Description)
End Sub

Private Function ReadPara() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objFile As New Scripting.FileSystemObject
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo DBError
    ReadPara = True
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(1))
    
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        ReadPara = False: Exit Function
    End If
    
    '设置和创建临时目录
    mBufferDir = App.Path & "\TempImage\"
    If Not objFile.FolderExists(mBufferDir) Then objFile.CreateFolder mBufferDir
    
    Timer1.Interval = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "存储间隔", "10")) * 1000

    
    '获取本机IP地址
    gstrLocalIP = funcGetLocalIP & ",127.0.0.1"
    
    '从数据库获取服务对设置
    strSQL = "Select 设备名,影像类别,PACSAE名称,PACS端口,设备IP地址,设备AE名称,设备端口,服务功能 From 影像dicom服务对 a ,影像设备目录 b " & _
             " Where a.设备号=b.设备号 And (a.PACS角色='SCP' or a.PACS角色='SCU' ) and NVL(b.状态,0)=1  And instr([1],PACSIP地址)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取DICOM服务对", gstrLocalIP)
    If rsTmp.RecordCount > 0 Then
        ReDim Services(rsTmp.RecordCount) As Service
        i = 1
    Else
        ReDim Services(0) As Service
    End If
    
    While Not rsTmp.EOF
        Services(i).DeviceAE = NVL(rsTmp!设备AE名称)
        Services(i).DeviceIP = NVL(rsTmp!设备IP地址)
        Services(i).DevicePort = NVL(rsTmp!设备端口)
        Services(i).DeviceName = NVL(rsTmp!设备名)
        Services(i).ServiceAE = NVL(rsTmp!PACSAE名称)
        Services(i).ServicePort = NVL(rsTmp!PACS端口)
        Services(i).SOP = NVL(rsTmp!服务功能)
        Services(i).Modality = NVL(rsTmp!影像类别)
        Services(i).Started = False
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    '从数据库获取参数设置
    strSQL = "Select distinct a.设备IP地址,a.PACSAE名称,b.参数名称,b.参数值 From 影像DICOM服务对 a,影像DICOM服务参数 b,影像设备目录 c " & _
             "Where a.服务ID = b.服务ID and c.设备号=a.设备号 And a.PACS角色='SCP' And NVL(c.状态,0)=1 and  instr([1],a.PACSIP地址)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取DICOM服务参数", gstrLocalIP)
    If rsTmp.RecordCount > 0 Then
        ReDim AEParas(rsTmp.RecordCount) As AEPara
        i = 1
    End If
    While Not rsTmp.EOF
        AEParas(i).AE = rsTmp!PACSAE名称
        AEParas(i).IP = rsTmp!设备IP地址
        AEParas(i).ParaName = rsTmp!参数名称
        AEParas(i).ParaValue = NVL(rsTmp!参数值)
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    '从数据库获取FTP存储设备
    strSQL = "Select 设备号,IP地址,FTP目录,FTP用户名,FTP密码 From 影像设备目录 Where 类型 =1 And NVL(状态,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取FTP存储设备")
    If rsTmp.RecordCount > 0 Then
        ReDim FTPDevices(rsTmp.RecordCount) As FTPDevice
        i = 1
    End If
    While Not rsTmp.EOF
        FTPDevices(i).No = rsTmp!设备号
        FTPDevices(i).IP = rsTmp!IP地址
        FTPDevices(i).FTPDir = NVL(rsTmp!FTP目录)
        FTPDevices(i).User = NVL(rsTmp!FTP用户名)
        FTPDevices(i).Password = NVL(rsTmp!FTP密码)
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function funGetServiceIndex(strServiceAE As String, Optional strDeviceIP As String = "") As Integer
'根据服务AE和设备IP查找对应的服务ID
'参数：     strServiceAE --- 服务的AE名称
'           strDeviceIP --- 设备的IP地址，对应打印路由，不需要输入设备的IP地址，因为设备的IP地址没有登记

    Dim i As Integer
    
    funGetServiceIndex = -1
    If strDeviceIP = "" Then
        For i = 1 To UBound(Services)
            If UCase(Services(i).ServiceAE) = UCase(strServiceAE) Then
                funGetServiceIndex = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To UBound(Services)
            If UCase(Services(i).ServiceAE) = UCase(strServiceAE) And Services(i).DeviceIP = strDeviceIP Then
                funGetServiceIndex = i
                Exit For
            End If
        Next i
    End If
End Function

Private Sub DViewer_QueryRequest(ByVal connection As DicomObjects.DicomConnection)
    Dim result As Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rq As DicomDataSet
    Dim rqs As DicomDataSets
    Dim rq1 As DicomDataSet
    Dim sql As String
    Dim D1 As DicomAttribute
    Dim NullSequence As New DicomDataSets
    Dim root As String
    Dim resultDS As DicomDataSets
    Dim Level As String
    Dim ds As DicomDataSet
    Dim RemoteAET As String
    Dim ResultImages As DicomImages
    Dim iService As Integer             '当前服务对应的服务对编号
    Dim blnAddModality As Boolean       '记录是否添加了影像类别的参数
    '服务基本参数
    Dim intFilterModality As Integer    '是否MWL过滤方法 0--按影像类别过滤，1--按IP地址过滤
    Dim intDayInterval As Integer       '查询延续日期
    Dim blnUseForceResult As Boolean    '是否使用强制结果
    Dim blnCGet As Boolean              '是否允许C-GET
    Dim intPatientIDMatch As Integer    'QR查询时，PatientID的匹配方式0--检查号，1--住院号/门诊号，2--医嘱ID
    Dim intBodypartType As Integer      '多部位方式，0-无；1-分隔符；2-多记录；3-多序列
    Dim strBodypartSplitter As String   '多部位分隔符
    Dim strMultiBodypartsName As String     '多部位名称，用“,”连接
    Dim strMultiBodypartsCode As String     '多部位代码，用“,”连接
    Dim intResultFilter As Integer          '查询结束条件，0-图像采集；1-检查完成
    Dim i As Integer
    Dim dtCurDate As Date                   '在外部程序获取数据库时间，避免在里面的多级循环中重复查询数据库
    
    On Error GoTo ProcError
    iService = -1
    '当数据断开时重新连接
    CheckDBConnect
    '获取连接中传过来的数据集
    Set rq = connection.request
    root = connection.root    '获取本次请求的根，有四种：PATIENT;STUDY;PATIENT/STUDY;WORKLIST
    If root = "WORKLIST" Then    '处理Worklist的请求

        '记录处理日志
        Call WriteProcessLog("DViewer_QueryRequest", "接收到Worklist请求", "请求的IP地址是： " & connection.RemoteIP & "，被呼叫的AE是：" & connection.CalledAET)
        
        '获取服务的基本参数设置
        funGetAEMWLParas connection.CalledAET, connection.RemoteIP, intFilterModality, intDayInterval, blnUseForceResult, _
                    intBodypartType, strBodypartSplitter, intResultFilter
        
        '记录处理日志
        Call WriteProcessLog("DViewer_QueryRequest", "提取Worklist服务参数完成", "intFilterModality = " & intFilterModality & ",intDayInterval= " & intDayInterval _
                        & ", blnUseForceResult = " & blnUseForceResult & ", intBodypartType= " & intBodypartType & ", strBodypartSplitter=" & strBodypartSplitter _
                        & ",intResultFilter = " & intResultFilter)
                        
        subLogDataset rq, "QueryRequest", "WORKLIST接收数据集"
        
        sql = "select /*+RULE*/ e.影像类别,a.医嘱ID,c.医嘱内容,a.发送号,a.英文名,b.执行间,b.发送人,a.出生日期,a.性别,b.首次时间,b.执行过程,a.检查UID,a.检查号 " & _
              ",Decode(C.病人来源,2,D.住院号,D.门诊号) As 标识号,a.年龄,a.姓名,a.检查设备,a.体重,a.附加主述 " & _
              "from 影像检查记录 a, 病人医嘱发送 b,病人医嘱记录 C,病人信息 D, 影像设备目录 E where a.医嘱ID = b.医嘱ID and a.发送号 = b.发送号 And A.检查设备 =E.设备号 " & _
              "And B.医嘱ID=C.ID And C.病人ID=D.病人ID And A.是否安排=1 And C.诊疗类别 in('D','E') And B.执行状态=3 AND C.相关ID IS NULL "
        '根据查询结束条件配置SQL查询条件
        If intResultFilter = 1 Then
            sql = sql & "And B.执行过程>=2 And B.执行过程<6 "
        Else
            sql = sql & "And B.执行过程=2 And A.检查UID IS Null "
        End If
              
        '根据参数过滤类型，增加搜索条件， 影像类被或者IP地址
        If intFilterModality = 0 Then
            '增加搜索条件 影像类别
            If rq.Attributes(&H40, &H100).Exists And Not IsNull(rq.Attributes(&H40, &H100).value) Then
                Set rqs = rq.Attributes(&H40, &H100).value   '一个嵌套的数据集
                Set rq1 = rqs(1)
                If rq1.Attributes(&H8, &H60).Exists And Not IsNull(rq1.Attributes(&H8, &H60).value) Then
                    If rq1.Attributes(&H8, &H60) <> "*" Then
                        sql = sql & " And UPPER(e.影像类别)='" & UCase(rq1.Attributes(&H8, &H60).value) & "'"
                        blnAddModality = True
                    End If
                End If
            End If
            If blnAddModality = False Then
                If iService = -1 Then iService = funGetServiceIndex(connection.CalledAET, connection.RemoteIP)
                sql = sql & " And UPPER(e.影像类别)='" & UCase(NVL(Services(iService).Modality)) & "'"
            End If
        ElseIf intFilterModality = 1 Then   '按照IP地址过滤
            sql = sql & " And E.IP地址='" & connection.RemoteIP & "'"
        End If
        
        If rq.Attributes(&H10, &H20).Exists And Not IsNull(rq.Attributes(&H10, &H20).value) Then
            If Trim(rq.Attributes(&H10, &H20).value) <> "*" Then
                sql = sql & " AND (1=2 "
                AddCondition sql, rq.Attributes(&H10, &H20), "a.检查号", False
                AddIDCondition sql, rq.Attributes(&H10, &H20), "D.门诊号", "", False
                AddIDCondition sql, rq.Attributes(&H10, &H20), "D.住院号", "", False
                sql = sql & ")"
            End If
        End If
        
        sql = sql & " AND B.首次时间>=SysDate-" & intDayInterval
        
        AddCondition sql, rq.Attributes(&H10, &H10), "Upper(a.英文名)"
        '检查号
        If rq.Attributes(&H8, &H50).Exists And Not IsNull(rq.Attributes(&H8, &H50).value) Then
            AddCondition sql, rq.Attributes(&H8, &H50), "a.检查号", True
        End If
        
'        AddLinkedDateTimeCondition sql, rq1.Attributes(&H40, 2), rq1.Attributes(&H40, 3), "b.首次时间"
        sql = sql & " order by a.医嘱ID"
        WriteCommLog "QueryRequest", "接收到WORKLIST请求", Replace(sql, "'", "‘")
        
        '因为设置好之后，每次WORKLIST的请求，参数都是一样的，因此就不需要绑定参数
        Set result = zlDatabase.OpenSQLRecord(sql, "查询WORKLIST请求")
        
        '获取数据库时间
        dtCurDate = zlDatabase.Currentdate
        
        '传输DICOM连接信息
        '添加强制信息
        If blnUseForceResult Then subReturnDataSet connection, 2, dtCurDate
        
        If strBodypartSplitter = "" Then strBodypartSplitter = "|"
        
        While Not result.EOF
            strMultiBodypartsName = ""
            strMultiBodypartsCode = ""
            '如果选择对码部位，不论是多记录还是分隔符，查询该医嘱ID对应的对码部位名称和对码部位代码
            If intBodypartType = 1 Or intBodypartType = 2 Or intBodypartType = 3 Then
                funGetBodypartValue result!医嘱ID, connection, strBodypartSplitter, strMultiBodypartsName, strMultiBodypartsCode
            End If
            
            If intBodypartType = 2 And strMultiBodypartsName <> "" Then '多记录方式返回多个对码部位，则循环返回多个部位
                For i = 0 To UBound(Split(strMultiBodypartsName, strBodypartSplitter))
                    subReturnDataSet connection, 1, dtCurDate, result, Split(strMultiBodypartsName, strBodypartSplitter)(i), _
                            Split(strMultiBodypartsCode, strBodypartSplitter)(i)
                Next i
            ElseIf intBodypartType = 3 And strMultiBodypartsName <> "" Then '多序列方式返回多个对码部位
                subReturnDataSet connection, 1, dtCurDate, result, strMultiBodypartsName, strMultiBodypartsCode, strBodypartSplitter
            Else
                subReturnDataSet connection, 1, dtCurDate, result, strMultiBodypartsName, strMultiBodypartsCode
            End If
            result.MoveNext
        Wend
        connection.SendStatus 0
        Exit Sub
    ElseIf root = "PATIENT" Or root = "STUDY" Or root = "PATIENT/STUDY" Then '处理查询检索Query/Retrieve请求
        '记录进入Query/Retrieve处理
        subLogDataset rq, "QueryRequest", "接收到Query/Retrieve请求"
        
        '读取QR的服务参数
        funGetQRParas connection.CalledAET, connection.RemoteIP, blnCGet, intPatientIDMatch
        
        '获取本次请求的层次,有四种：PATIENT;STUDY;SERIES;IMAGE
        Level = rq.Attributes(&H8, &H52)
        '处理C-FIND类型的连接请求，查询数据库，只返回查询结果，不返回图像
        If connection.operation = "C-FIND" Then
            Set resultDS = New DicomDataSets
            
            If Level = "PATIENT" Then   '处理病人级别的查询,支持病人姓名，病人ID作为查询条件
                '根据病人ID的匹配方式组织SQL查询语句
                If intPatientIDMatch = 0 Then       '检查号
                    sql = "Select a.医嘱ID,a.检查号 as PatientID,a.姓名,a.性别,a.英文名,a.出生日期 From 影像检查记录 a Where 检查uid Is Not Null "
                ElseIf intPatientIDMatch = 1 Then   '住院号/门诊号
                    sql = "Select a.医嘱ID,Decode(b.病人来源, 2, c.住院号, c.门诊号) as PatientID,a.姓名,a.性别,a.英文名,a.出生日期 " _
                          & " From 影像检查记录 a,病人医嘱记录 b,病人信息 c Where a.检查uid Is Not Null and a.医嘱ID=b.Id And b.病人ID=c.病人ID "
                Else                                '或者医嘱ID
                    sql = "Select a.医嘱ID,a.医嘱ID as PatientID,a.姓名,a.性别,a.英文名,a.出生日期 From 影像检查记录 a Where 检查uid Is Not Null "
                End If
               
                AddStringCondition sql, rq.Name, "a.英文名"
                AddDateCondition sql, rq.Attributes(&H10, &H30), "a.出生日期"
                
                '添加PatientID条件,不为* 则按照给定的PatientID查询
                If rq.PatientID <> "*" And rq.PatientID <> "" Then
                    If intPatientIDMatch = 0 Then   '检查号,去掉多余的*号
                        sql = sql & " and a.检查号= '" & Replace(rq.PatientID, "*", "") & "'"
                    ElseIf intPatientIDMatch = 1 Then   '住院号，门诊号
                        sql = sql & " and ((c.住院号=" & Val(rq.PatientID) & " And b.病人来源=2) Or (c.门诊号=" _
                            & Val(rq.PatientID) & " And b.病人来源<>2))"
                    Else    '医嘱ID
                        sql = sql & " and a.医嘱ID = " & Val(rq.PatientID)
                    End If
                End If
                
                WriteCommLog "QueryRequest", "病人级别的Query", Replace(sql, "'", "‘")
                Set result = zlDatabase.OpenSQLRecord(sql, "病人级别的Query")
                
                '返回查询结果数据集，英文名，PatientID，出生日期，性别，被检索的AE
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H10, &H10, NVL(result!英文名)
                    AddResultItem ds, rq, &H10, &H20, result!PatientID
                    AddResultItem ds, rq, &H10, &H30, NVL(result!出生日期)
                    AddResultItem ds, rq, &H10, &H40, IIf(NVL(result!性别) = "女", "F", IIf(NVL(result!性别) = "男", "M", "O"))
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '处理检查级别的查询
            If Level = "STUDY" Then
                '根据病人ID的匹配方式组织SQL查询语句,支持英文名，出生日期，检查UID，检查日期（接收日期），PatientID 作为查询条件
                If intPatientIDMatch = 0 Then       '检查号
                    sql = "Select  a.医嘱ID,a.检查号 as PatientID,a.影像类别,a.检查uid,a.性别,a.英文名,a.出生日期,a.接收日期 " _
                        & " From 影像检查记录 a Where 检查uid Is Not Null "
                ElseIf intPatientIDMatch = 1 Then   '住院号，门诊号
                    sql = "Select a.医嘱ID,Decode(b.病人来源, 2, c.住院号, c.门诊号) as PatientID,a.影像类别,a.检查uid,a.性别,a.英文名,a.出生日期,a.接收日期 " _
                          & " From 影像检查记录 a,病人医嘱记录 b,病人信息 c Where a.检查uid Is Not Null and a.医嘱ID=b.Id And b.病人ID=c.病人ID "
                Else                                '医嘱ID
                    sql = "Select a.医嘱ID,a.医嘱ID as PatientID,a.影像类别,a.检查uid,a.性别,a.英文名,a.出生日期,a.接收日期 From 影像检查记录 a Where 检查uid Is Not Null "
                End If
                
                If root = "STUDY" Then
                    '对于以检查为根的查询，在检查层次上返回的是病人名字和检查的组合,因此允许使用姓名和出生日期做条件
                    AddStringCondition sql, rq.Name, "a.英文名"
                    AddDateCondition sql, rq.Attributes(&H10, &H30), "a.出生日期"
                End If
                AddStringCondition sql, rq.StudyUID, "a.检查uid"
                AddDateTimeCondition sql, rq.Attributes(&H8, &H20), rq.Attributes(&H8, &H30), "a.接收日期"
                
                '添加PatientID条件,不为* 则按照给定的PatientID查询
                If rq.PatientID <> "*" And rq.PatientID <> "" Then
                    If intPatientIDMatch = 0 Then   '检查号
                        sql = sql & " and a.检查号= '" & Replace(rq.PatientID, "*", "") & "'"
                    ElseIf intPatientIDMatch = 1 Then   '住院号，门诊号
                        sql = sql & " and ((c.住院号=" & Val(rq.PatientID) & " And b.病人来源=2) Or (c.门诊号=" _
                            & Val(rq.PatientID) & " And b.病人来源<>2))"
                    Else    '医嘱ID
                        sql = sql & " and a.医嘱ID = " & Val(rq.PatientID)
                    End If
                End If
                
                WriteCommLog "QueryRequest", "检查级别的Query", Replace(sql, "'", "‘")
                Set result = zlDatabase.OpenSQLRecord(sql, "检查级别的Query")
                
                '组织返回的数据集，包括检查UID，检查日期（接收日期），英文名，PatientID，出生日期，被检索的AE，检查的影像类别
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H20, &HD, NVL(result!检查UID, 1)
                    '检查描述是可选项，我们的数据库里面 没有这个数据，因此暂不支持此项
                    'AddResultItem ds, rq, &H8, &H1030, result!StudyDescription
                    AddResultItem ds, rq, &H8, &H20, Format(NVL(result!接收日期, "19000101"), "YYYYMMDD")
                    AddResultItem ds, rq, &H8, &H30, Format(NVL(result!接收日期, "12:01:01"), "hhmmss")
                    If root = "STUDY" Then
                        AddResultItem ds, rq, &H10, &H10, result!英文名
                        AddResultItem ds, rq, &H10, &H20, result!PatientID
                        AddResultItem ds, rq, &H10, &H30, result!出生日期
                    End If
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    AddResultItem ds, rq, &H8, &H60, result!影像类别
                    AddResultItem ds, rq, &H8, &H61, result!影像类别
                    AddResultItem ds, rq, &H10, &H40, IIf(NVL(result!性别) = "女", "F", IIf(NVL(result!性别) = "男", "M", "O"))
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '处理序列级别的查询,支持的条件包括：检查UID，序列UID
            If Level = "SERIES" Then
                sql = "select /*+RULE*/ b.序列uid,b.序列描述,b.序列号,a.影像类别 from 影像检查记录 a ,影像检查序列 b " _
                            & "where  a.检查uid = b.检查uid"
                AddStringCondition sql, rq.StudyUID, "a.检查uid"
                AddStringCondition sql, rq.SeriesUID, "b.序列uid"
                
                WriteCommLog "QueryRequest", "序列级别的Query", Replace(sql, "'", "‘")
                Set result = zlDatabase.OpenSQLRecord(sql, "序列级别的Query")
                
                '组织返回的数据集，包括：序列UID，序列描述，序列号，影像类别，被检索的AE,图像总数
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H20, &HE, result!序列uid
                    AddResultItem ds, rq, &H8, &H103E, result!序列描述
                    AddResultItem ds, rq, &H20, &H11, result!序列号
                    AddResultItem ds, rq, &H8, &H60, result!影像类别
                    AddCountItem ds, rq, &H20, &H1209, "SeriesUID", result!序列uid, "InstanceUID"
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '处理图像级别的查询，支持的条件包括：序列UID，图像UID
            If Level = "IMAGE" Then
                sql = "  select /*+RULE*/ t.图像uid,t.序列uid,t.图像号 from 影像检查图象 t where 1=1"
                AddStringCondition sql, rq.SeriesUID, "t.序列uid"
                AddStringCondition sql, rq.instanceUID, "t.图像uid"
                
                WriteCommLog "QueryRequest", "图像级别的Query", Replace(sql, "'", "‘")
                Set result = zlDatabase.OpenSQLRecord(sql, "图像级别的Query")
                
                '组织返回的数据集，包括：图像UID，图像号，被检索的AE
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H8, &H18, result!图像uid
                    AddResultItem ds, rq, &H20, &H13, result!图像号
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            
            For Each ds In resultDS
                subLogDataset ds, "QueryRequest", "Query/Retrieve查询结果"
            Next ds
            
            '发送查询结果
            connection.SendData resultDS, &HFF00
        ElseIf connection.operation = "C-GET" Or connection.operation = "C-MOVE" Then
            '处理C-GET和C-MOVE，根据请求的方法，将图像返回给SCU
            If connection.operation = "C-MOVE" Then
                '检查AE名称是否在被许可的AE集中，如果不在，则拒绝传送图像
                '因为C-MOVE是必须使用一个新的连接来传送图像的，不可以使用现有的连接，因此需要根据
                '传过来的AE名称，查找到该AE对应的IP地址和端口号。
                RemoteAET = connection.Destination
                sql = "Select decode(PACS角色,'SCP',PACSIP地址,设备IP地址) As IP地址,decode(PACS角色,'SCP',PACS端口,设备端口) As 端口号 " & _
                      "From 影像DICOM服务对 Where 服务功能='图像接收' And (PACS角色='SCP' And upper(PACSAE名称) =[1]) Or (PACS角色='SCU' And 设备AE名称=[1])"
                WriteCommLog "QueryRequest", "C-MOVE查找图像移动目的主机", Replace(sql, "'", "‘")
                WriteCommLog "QueryRequest", "C-MOVE查找图像移动目的主机--参数", "[1] = " & UCase(RemoteAET)
                
                Set result = zlDatabase.OpenSQLRecord(sql, "查找C-MOVE的目的地", UCase(RemoteAET))
                If result.EOF Then
                    WriteCommLog "QueryRequest", "C-MOVE主机查找错误", "图像移动的目的地不可知"
                    connection.Errors.Attributes.Add 0, &H902, "图像移动的目的地不可知。"
                    connection.SendStatus (&HA801)
                    Exit Sub
                End If
                '以下这个设置目的地是必须的，如果不设置，将不能成功的发送图像。
                '因为C-MOVE操作要求一定要使用一个新的连接来发送图像。
                WriteCommLog "QueryRequest", "C-MOVE找到图像移动目的主机", "IP地址：" & result!IP地址 _
                             & "，端口号：" & result!端口号 & ",本机AE:" & connection.CalledAET & ",远程AE:" & RemoteAET
                          
                connection.SetDestination result!IP地址, result!端口号, connection.CalledAET, RemoteAET
            ElseIf connection.operation = "C-GET" Then
                '在这里处理不允许C-GET的情况,不允许C-GET,则拒绝它
                '判断是否支持C-GET
                If blnCGet = False Then Exit Sub
            End If
            Set ResultImages = New DicomImages
            If Level = "PATIENT" Then
            '根据病人ID的匹配方式组织SQL查询语句
                If intPatientIDMatch = 0 Then       '检查号
                    sql = "select 病人ID from 影像检查记录 a ,病人医嘱记录 b where a.医嘱id=b.id and a.检查号=[1]"
                ElseIf intPatientIDMatch = 1 Then   '住院号/门诊号
                    sql = "select 病人id from  病人信息  where 门诊号=[1] or 住院号=[1]"
                Else                                '或者医嘱ID
                    sql = "select 病人id from  病人医嘱记录  where ID=[1]"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(sql, "查询病人ID", IIf(intPatientIDMatch = 0, rq.PatientID, Val(rq.PatientID)))
                
                WriteCommLog "QueryRequest", "C-MOVE查找病人ID", "sql = " & sql & " ,[1] = " & Val(rq.PatientID)
                
                If Not rsTmp.EOF Then
                    Set ResultImages = GetAllImageFiles(Level, rsTmp!病人id)
                End If
            End If
            
            If Level = "STUDY" Then Set ResultImages = GetAllImageFiles(Level, rq.StudyUID)
            If Level = "SERIES" Then Set ResultImages = GetAllImageFiles(Level, rq.SeriesUID)
            If Level = "IMAGE" Then Set ResultImages = GetAllImageFiles(Level, rq.instanceUID)
            If ResultImages Is Nothing Then
                WriteCommLog "QueryRequest", "发送图像", "未返回任何图像"
                connection.SendStatus 0
                WriteCommLog "QueryRequest", "QR通讯完成", "未返回任何图像"
            Else
                WriteCommLog "QueryRequest", "准备发送图像", "图像数量为：" & ResultImages.Count & _
                         "图像检查UID为：" & IIf(ResultImages.Count = 0, "无", ResultImages(1).StudyUID)
                connection.SendImages ResultImages
                WriteCommLog "QueryRequest", "发送图像完成", "图像数量为：" & ResultImages.Count
            End If
        End If
    End If
    Exit Sub
ProcError:
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Function GetAllImageFiles(Level As String, SearchValue As String) As DicomImages
'------------------------------------------------
'功能：在Q/R查询中使用，用来返回符合条件的图像集合
'参数： Level －－查询级别
'       SearchValue－－查询条件
'返回：DicomImages查询到的图像集合
'-----------------------------------------------
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double
    Dim FrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim Item As MSComctlLib.ListItem
    Dim clsUseFTP1 As New clsFtp
    Dim clsUseFTP2 As New clsFtp
    
    
    Dim aSeriesUIDs() As String     '保存用于获取图像的序列UID集合
    Dim i As Integer                '循环记数器
    Dim OneSeriesUID As String      '保存单个序列UID
    Dim lngResult As Long           '保存返回值
    Dim AllImages As New DicomImages
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    
    Dim curImage As DicomImage, GetAllImages As New DicomImages
    
    Dim bln1stDev As Boolean
    bln1stDev = True
    
    On Error GoTo DBError
    Screen.MousePointer = vbHourglass
    
    '要分病人，检查，序列，图像 四种层次来获取并返回图像,根据层次的不同，获取序列UID集合的方法不同
    If Level = "PATIENT" Then
        strSQL = "select /*+RULE*/ e.序列uid from 影像检查记录 c , 影像检查序列 e , " _
                    & "(select a.病人id,b.医嘱id,b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " _
                    & "where a.病人id=" & SearchValue & "　AND a.相关ID IS NULL  and a.id=b.医嘱id) d " _
                    & "Where c.医嘱id = d.医嘱id And c.发送号 = d.发送号 and c.检查uid = e.检查uid"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "STUDY" Then
        strSQL = "select /*+RULE*/ b.序列uid from 影像检查记录 a, 影像检查序列 b where a.检查uid = b.检查uid " _
                    & "and a.检查uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "SERIES" Then
        strSQL = "select /*+RULE*/ t.序列uid from 影像检查序列 t where t.序列uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "IMAGE" Then
        strSQL = "select /*+RULE*/ t.序列uid from 影像检查序列 t ,影像检查图象 q where t.序列uid = q.序列uid " _
                    & "and q.图像uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    End If
  
    WriteCommLog "GetAllImageFiles", "根据ID查找图像", "sql = " & strSQL & " ,[1] = " & SearchValue
  
    If rsTmp.RecordCount <= 0 Then Exit Function    '没有结果则返回
    '处理查询结果，将查询出来的序列UID放入aSeriesUIDs集合中
    ReDim aSeriesUIDs(rsTmp.RecordCount) As String
    i = 1
    While Not rsTmp.EOF
        aSeriesUIDs(i) = rsTmp!序列uid
        i = i + 1
        rsTmp.MoveNext
    Wend
    For i = 1 To UBound(aSeriesUIDs)
        OneSeriesUID = aSeriesUIDs(i)
        strSQL = "Select /*+RULE*/ A.图像号,D.FTP用户名 as User1 ,D.FTP密码 as Psw1 , D.IP地址 as IP1 , " & _
            "'/'||D.Ftp目录||'/' As FtpPath1,D.设备号 as 设备号1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
            "||C.检查UID||'/' As Path,A.图像UID as ImgName , " & _
            "E.FTP用户名 as User2, E.FTP密码 as Psw2 , E.IP地址 as IP2 , " & _
            "'/'||E.Ftp目录||'/' As FtpPath2,E.设备号 as 设备号2 " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
            "And A.序列UID= [1] Order By A.图像号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, OneSeriesUID)
        If rsTmp.RecordCount > 0 Then
            If strDeviceNO1 <> rsTmp("设备号1") Then
                clsUseFTP1.FuncFtpDisConnect
                clsUseFTP1.FuncFtpConnect rsTmp("IP1"), rsTmp("User1"), rsTmp("Psw1")
                strDeviceNO1 = rsTmp("设备号1")
            End If
            If strDeviceNO2 <> rsTmp("设备号2") Then
                clsUseFTP2.FuncFtpDisConnect
                clsUseFTP2.FuncFtpConnect rsTmp("IP2"), rsTmp("User2"), rsTmp("Psw2")
                strDeviceNO2 = rsTmp("设备号2")
            End If
            
            Do While Not rsTmp.EOF
                '判断检索的级别是否IMAGE，如果是，则检查图像的UID
                If Level <> "IMAGE" Or (Level = "IMAGE" And SearchValue = rsTmp("ImgName")) Then
                    If Dir(mBufferDir & rsTmp("ImgName")) = vbNullString Then
                        '使用FTP从服务器获取图像
                        lngResult = clsUseFTP1.FuncDownloadFile(rsTmp("FtpPath1") & rsTmp("Path"), mBufferDir & rsTmp("ImgName"), rsTmp("ImgName"))
                        If lngResult <> 0 Then  '从设备1中没有图像
                            lngResult = clsUseFTP2.FuncDownloadFile(rsTmp("FtpPath2") & rsTmp("Path"), mBufferDir & rsTmp("ImgName"), rsTmp("ImgName"))
                        End If
                    End If
                    AllImages.ReadFile mBufferDir & rsTmp("ImgName")
                End If
                rsTmp.MoveNext
            Loop
        End If
    Next i
    clsUseFTP1.FuncFtpDisConnect
    clsUseFTP2.FuncFtpDisConnect
    Screen.MousePointer = vbDefault
    Set GetAllImageFiles = AllImages
    Exit Function

DBError:
    clsUseFTP1.FuncFtpDisConnect
    clsUseFTP2.FuncFtpDisConnect
    Screen.MousePointer = vbDefault
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
    Call WriteLog(2, err.Number, err.Description)
End Function

Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String)
'功能： 记录通讯日志
'参数： logSubName -- 通讯所在的过程名称
'       logTitle  --  通讯标题
'       logDesc --    日志描述

    Dim strSQL As String
    
    On Error Resume Next
    
    If mmuCommLog.Checked Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        strSQL = "Insert into DICOM通讯日志 (通讯时间,通讯函数,记录标题,记录内容) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
End Sub

Private Sub subLogDataset(ds As DicomDataSet, logSubName As String, logTitle As String)
    Dim strLog As String
    If mmuCommLog.Checked Then
        AppendAttributes strLog, "", ds.Attributes
        WriteCommLog logSubName, logTitle, Replace(strLog, "'", "‘")
    End If
End Sub

Private Sub AppendAttributes(ByRef list As String, prefix As String, ByRef ob As Object)
    Dim at As DicomAttribute
    Dim s As DicomDataSets
    Dim i As Integer
    Dim v As Variant
    For Each at In ob
        list = list & prefix & "(" & hex4(at.group) & "," & hex4(at.element) & ") : "
        list = list & Left(at.Description & Space(30), 30) & ": "
        If (at.group = &H7FE0) Then ' pixel data
            list = list & "Pixel data" & vbCrLf
        ElseIf (VarType(at.value) = 9) Then ' i.e. a sequence
            Set s = at.value
            list = list & "Sequence of " & s.Count & " items:" & vbCrLf
            For i = 1 To s.Count
                list = list & prefix & ">---------------" & vbCrLf
                AppendAttributes list, prefix & ">", s(i).Attributes
            Next
            list = list & prefix & ">---------------" & vbCrLf
        Else
            v = at.value ' could be variant or array
            If (VarType(v) > 8192) Then ' i.e. an array
                list = list & "Multiple values :" & vbCrLf & "              "
                If UBound(v, 1) > 32 Then
                    list = list & "Array of " & UBound(v, 1) & " elements"
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        list = list & v(i)
                        If i <> UBound(v, 1) Then list = list & " : "
                    Next
                End If
                list = list & vbCrLf
            Else
                list = list & v & vbCrLf
            End If
        End If
    Next
End Sub

Private Function hex4(ByVal v As Integer) As String
    hex4 = Right("000" & Hex(v), 4)
End Function

Private Sub subReturnDataSet(connection As DicomConnection, intType As Integer, dtCurDate As Date, Optional rsOracle As Recordset, _
    Optional ByVal strBodypartName As String = "", Optional ByVal strBodyparCode As String = "", _
    Optional ByVal strBodypartSplitter As String = "")
'组织并返回Worklist的数据集
'参数： connection ---Worklist所在的Dicom连接
'       intType --- 类型，1-返回正常查询结果；2-返回强制结果
'       dtCurDate --- 当前数据库时间
'       rsOracle --- 要返回的Oracle数据集
'       strBodypartName --- 对码部位名称
'       strBodyparCode --- 对码部位代码
'       strBodypartSplitter --- 对码部位的分隔符，如果有分隔符，则说明是使用序列来传对码部位，因此需要处理序列

    Dim dsSeqItem(5) As New DicomDataSet
    Dim dsSeqItemTemp As DicomDataSet
    Dim dsResult As New DicomDataSet
    Dim dssSeq(5) As New DicomDataSets
    Dim rq As DicomDataSet
    Dim rq1(5) As DicomDataSet
    Dim rqs(5) As DicomDataSets
    Dim D1 As DicomAttribute
    Dim NullSequence As New DicomDataSets
    Dim NullSeqItem As New DicomDataSet
    Dim strValue As String      '保存数据集返回的一个结果
    Dim intValueType As Integer     '结果的类型，0-普通结果；1-对码部位代码；2-对码部位名称
    Dim strField As String
    Dim i As Integer
    Dim strFieldValue As String
    Dim strSQL As String
    Dim UpID() As Integer
    Dim dbResult As Recordset
    Dim intBodypartNum As Integer   '对码部位的部位数量
    Dim arrBodypartName() As String '对码部位名称串
    Dim arrBodypartCode() As String '对码部位代码串
    Dim int上级ID As Integer        '数据集的上级ID
    
    On Error GoTo ProcError
    '初始化返回数据集
    Set rq = connection.request
    
    '初始化非序列的属性
    Set dsResult = rq
    
    '先判断是否使用序列来传对码部位，如果使用，则解析对码部位的数量
    If strBodypartSplitter <> "" Then
        arrBodypartName = Split(strBodypartName, strBodypartSplitter)
        arrBodypartCode = Split(strBodyparCode, strBodypartSplitter)
        intBodypartNum = UBound(arrBodypartName) + 1
    End If
    
    '初始化序列属性
    If rq.Attributes(&H8, &H1110).Exists And Not IsNull(rq.Attributes(&H8, &H1110).value) Then
        Set rqs(1) = rq.Attributes(&H8, &H1110).value
        If rqs(1).Count > 0 Then
            Set rq1(1) = rqs(1)(1)
            For Each D1 In rq1(1).Attributes
                dsSeqItem(1).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(1).Add dsSeqItem(1)
            dsResult.Attributes.Add &H8, &H1110, dssSeq(1)
        End If
    Else
        Set rqs(1) = NullSequence
        Set rq1(1) = NullSeqItem
    End If
    If rq.Attributes(&H8, &H1120).Exists And Not IsNull(rq.Attributes(&H8, &H1120).value) Then
        Set rqs(2) = rq.Attributes(&H8, &H1120).value
        If rqs(2).Count > 0 Then
            Set rq1(2) = rqs(2)(1)
            For Each D1 In rq1(2).Attributes
                dsSeqItem(2).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(2).Add dsSeqItem(2)
            dsResult.Attributes.Add &H8, &H1120, dssSeq(2)
        End If
    Else
        Set rqs(2) = NullSequence
        Set rq1(2) = NullSeqItem
    End If
    
    '这个序列需要处理对码部位，（8,100）可能是部位对码
    If rq.Attributes(&H32, &H1064).Exists And Not IsNull(rq.Attributes(&H32, &H1064).value) Then
        Set rqs(3) = rq.Attributes(&H32, &H1064).value
        If rqs(3).Count > 0 Then
            Set rq1(3) = rqs(3)(1)
            If intBodypartNum > 1 Then
                For i = 1 To intBodypartNum
                    Set dsSeqItemTemp = New DicomDataSet
                    For Each D1 In rq1(3).Attributes
                        dsSeqItemTemp.Attributes.Add D1.group, D1.element, D1.value
                    Next
                    dssSeq(3).Add dsSeqItemTemp
                    If i = 1 Then
                        Set dsSeqItem(3) = dsSeqItemTemp
                    End If
                Next i
            Else
                For Each D1 In rq1(3).Attributes
                    dsSeqItem(3).Attributes.Add D1.group, D1.element, D1.value
                Next
                dssSeq(3).Add dsSeqItem(3)
            End If
            dsResult.Attributes.Add &H32, &H1064, dssSeq(3)
        End If
    Else
        Set rqs(3) = NullSequence
        Set rq1(3) = NullSeqItem
    End If
    
    If rq.Attributes(&H40, &H100).Exists And Not IsNull(rq.Attributes(&H40, &H100).value) Then
        Set rqs(5) = rq.Attributes(&H40, &H100).value
        If rqs(5).Count > 0 Then
            Set rq1(5) = rqs(5)(1)
            For Each D1 In rq1(5).Attributes
                dsSeqItem(5).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(5).Add dsSeqItem(5)
            dsResult.Attributes.Add &H40, &H100, dssSeq(5)
        End If
        
        '这个序列需要处理对码部位，（8,100）可能是部位对码
        If rq1(5).Attributes(&H40, &H8).Exists And Not IsNull(rq1(5).Attributes(&H40, &H8).value) Then
            Set rqs(4) = rq1(5).Attributes(&H40, &H8).value
            If rqs(4).Count > 0 Then
                Set rq1(4) = rqs(4)(1)
                If intBodypartNum > 1 Then
                    For i = 1 To intBodypartNum
                        Set dsSeqItemTemp = New DicomDataSet
                        For Each D1 In rq1(4).Attributes
                            dsSeqItemTemp.Attributes.Add D1.group, D1.element, D1.value
                        Next
                        dssSeq(4).Add dsSeqItemTemp
                        If i = 1 Then
                            Set dsSeqItem(4) = dsSeqItemTemp
                        End If
                    Next i
                Else
                    For Each D1 In rq1(4).Attributes
                        dsSeqItem(4).Attributes.Add D1.group, D1.element, D1.value
                    Next
                    dssSeq(4).Add dsSeqItem(4)
                End If
                dsSeqItem(5).Attributes.Add &H40, &H8, dssSeq(4)
            End If
        Else
            Set rqs(4) = NullSequence
            Set rq1(4) = NullSeqItem
        End If
    
    Else
        Set rqs(5) = NullSequence
        Set rq1(5) = NullSeqItem
    End If
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '首先创建数据集数组
    strSQL = "Select a.Id, a.组号, a.元素号 From 影像MWL结果集 a ,影像DICOM服务对 b " & _
             " Where a.值类型 = 'SQ' and a.服务ID =b.服务ID and  a.选中 = 1 and upper(b.PACSAE名称)=[1] and b.设备IP地址 =[2]  Order by id "
    Set dbResult = zlDatabase.OpenSQLRecord(strSQL, "读取MWL数据集序列", CStr(UCase(connection.CalledAET)), CStr(connection.RemoteIP))
    Dim lngMin As Long
    Dim lngMax As Long
     
    If dbResult.RecordCount = 0 Then
        '查询到的数据集为0，是错误的，最少需要选择上 （40,100），因此退出本过程，记录错误日志
        err.Raise 10, , Replace(strSQL, "'", "‘") & vbNewLine & "查询不到数据集，最少应该选择返回（40,100）这个数据集"
    End If
     
    lngMin = dbResult!id
    dbResult.MoveLast
    lngMax = dbResult!id
    ReDim UpID(lngMin To lngMax) As Integer
    
    dbResult.MoveFirst
    While Not dbResult.EOF
        If dbResult!组号 = "0008" And dbResult!元素号 = "1110" Then
            UpID(dbResult!id) = 1
        ElseIf dbResult!组号 = "0008" And dbResult!元素号 = "1120" Then
            UpID(dbResult!id) = 2
        ElseIf dbResult!组号 = "0032" And dbResult!元素号 = "1064" Then
            UpID(dbResult!id) = 3
        ElseIf dbResult!组号 = "0040" And dbResult!元素号 = "0008" Then
            UpID(dbResult!id) = 4
        ElseIf dbResult!组号 = "0040" And dbResult!元素号 = "0100" Then
            UpID(dbResult!id) = 5
        End If
        dbResult.MoveNext
    Wend
    
    '循环结果集，填写结果值
    strSQL = "Select a.组号, a.元素号, a.上级id, a.数据值, a.是否递增,a.值类型, a.元素类型, a.强制结果值 From 影像MWL结果集 a , " & _
             " 影像DICOM服务对 b Where a.服务ID =b.服务ID and  a.选中 = 1 and upper(b.PACSAE名称)=[1] and b.设备IP地址 =[2]"
    Set dbResult = zlDatabase.OpenSQLRecord(strSQL, "读取MWL数据集", CStr(UCase(connection.CalledAET)), CStr(connection.RemoteIP))
    
    While Not dbResult.EOF
        intValueType = 0
        If dbResult!值类型 <> "SQ" Then  '数据类型不是SQ的，才解码返回值
            '从数据库中读取返回值字符串
            If intType = 1 Then         '返回正常查询结果
                strValue = NVL(dbResult!数据值)
                '解码返回字符串
                Do While InStr(strValue, "[") <> 0
                    If InStr(strValue, "]") = 0 Or InStr(strValue, "]") < InStr(strValue, "[") Then
                        strValue = ""
                        Exit Do
                    End If
                    strField = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
                    
                    strFieldValue = ""
                    On Error Resume Next
                    If strField = "CallingAET" Then
                        strFieldValue = connection.CallingAET
                    ElseIf strField = "对码部位名称" Then
                        strFieldValue = strBodypartName
                        intValueType = 2
                    ElseIf strField = "对码部位代码" Then
                        strFieldValue = strBodyparCode
                        intValueType = 1
                    Else
                        strFieldValue = funGetFieldValue(strField, rsOracle, dtCurDate)
                    End If
                    
                    strValue = Replace(strValue, "[" & strField & "]", strFieldValue)
                Loop
                
            ElseIf intType = 2 Then         '返回强制结果
                strValue = NVL(dbResult!强制结果值)
            End If
            '处理递增的结果
            If dbResult!是否递增 = True Then
                strValue = strValue & mintWLCount
                mintWLCount = mintWLCount + 1
            End If
            '处理元素类型为1或1C的，结果不允许返回空值
            If dbResult!元素类型 = "1" Or UCase(dbResult!元素类型) = "1C" Then
                If strValue = "" Then strValue = "1"
            End If
        End If
        
        '处理非序列类型的数据
        If dbResult!值类型 <> "SQ" Then
            If IsNull(dbResult!上级ID) Then      '上级ID为空，直接填写数据集
                AddResultItem dsResult, rq, Int("&H" & dbResult!组号), Int("&H" & dbResult!元素号), strValue
            Else    '有上级ID，说明是嵌套在序列中的数据
                '知道上级ID，需要查找到使用那个数据集
                int上级ID = UpID(dbResult!上级ID)
                If intValueType = 1 And strBodypartSplitter <> "" Then    '对码部位代码
                    For i = 1 To intBodypartNum
                        AddResultItem dssSeq(int上级ID)(i), rq1(int上级ID), Int("&H" & dbResult!组号), Int("&H" & dbResult!元素号), arrBodypartCode(i - 1)
                    Next i
                ElseIf intValueType = 2 And strBodypartSplitter <> "" Then    '对码部位名称
                    For i = 1 To intBodypartNum
                        AddResultItem dssSeq(int上级ID)(i), rq1(int上级ID), Int("&H" & dbResult!组号), Int("&H" & dbResult!元素号), arrBodypartName(i - 1)
                    Next i
                Else
                    AddResultItem dsSeqItem(int上级ID), rq1(int上级ID), Int("&H" & dbResult!组号), Int("&H" & dbResult!元素号), strValue
                End If
            End If
        End If
        dbResult.MoveNext
    Wend
    
    connection.SendData dsResult, &HFF00  '如果有不匹配的字段，可以使用&HFF01
    subLogDataset dsResult, "subReturnDataset", IIf(intType = 2, "WORKLIST强制返回数据集", "WORKLIST返回数据集")
    Exit Sub
ProcError:
    Call WriteLog(10, err.Number, err.Description)
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "错误：" & Format(lngErrCounts, "@@@@@@") & "条"
End Sub

Private Function funGetFieldValue(strField As String, rsDataSet As Recordset, dtCurDate As Date) As String
    Dim lngAge As Long
    Dim strAge As String
        
    Select Case strField
        Case "首次日期"
            funGetFieldValue = Format(NVL(rsDataSet!首次时间, "30000101"), "YYYY-MM-DD")
        Case "首次时间"
            funGetFieldValue = Format(NVL(rsDataSet!首次时间, "000001"), "HH:MM:SS")
        Case "影像类别"
            funGetFieldValue = rsDataSet!影像类别
        Case "执行间"
            funGetFieldValue = NVL(rsDataSet!执行间, "XX")
        Case "执行过程"
            funGetFieldValue = NVL(rsDataSet!执行过程, "2")
        Case "医嘱ID"
            funGetFieldValue = rsDataSet!医嘱ID
        Case "检查部位"
            If InStr(NVL(rsDataSet!医嘱内容), ":") > 0 Then
                funGetFieldValue = Split(rsDataSet!医嘱内容, ":")(1)
            Else
                funGetFieldValue = rsDataSet!医嘱内容
            End If
        Case "发送号"
            funGetFieldValue = rsDataSet!发送号
        Case "检查号"
            funGetFieldValue = rsDataSet!检查号
        Case "标识号"
            funGetFieldValue = rsDataSet!标识号
        Case "英文名"
            funGetFieldValue = rsDataSet!英文名
        Case "性别"
            funGetFieldValue = IIf(NVL(rsDataSet!性别) = "男", "M", IIf(NVL(rsDataSet!性别) = "女", "F", "O"))
        Case "年龄"
            If NVL(rsDataSet!出生日期) <> "" Then
                '根据出生日期转换位dicom格式的年龄
                
                '按岁计算
                lngAge = DateDiff("yyyy", CDate(rsDataSet!出生日期), dtCurDate)
                If lngAge >= 3 Then
                    funGetFieldValue = Format(lngAge, "000") & "Y"
                    Exit Function
                End If
                
                '按月计算
                lngAge = DateDiff("m", CDate(rsDataSet!出生日期), dtCurDate)
                If lngAge >= 3 Then
                    funGetFieldValue = Format(lngAge, "000") & "M"
                    Exit Function
                End If
                
                
                '按周计算
                lngAge = DateDiff("w", CDate(rsDataSet!出生日期), dtCurDate)
                If lngAge >= 4 Then
                    funGetFieldValue = Format(lngAge, "000") & "W"
                    Exit Function
                End If
                
                '按天计算
                lngAge = DateDiff("d", CDate(rsDataSet!出生日期), dtCurDate)
                funGetFieldValue = Format(lngAge, "000") & "D"
                
                Exit Function
            Else
                '根据录入的年龄转换为dicom格式的年龄形式
                strAge = NVL(rsDataSet!年龄, "0")
                
                lngAge = Val(strAge)
                
                Select Case True
                    Case (InStr(strAge, "岁") > 0), (InStr(UCase(strAge), "Y") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "Y"
                    Case (InStr(strAge, "月") > 0), (InStr(UCase(strAge), "M") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "M"
                    Case (InStr(strAge, "周") > 0), (InStr(UCase(strAge), "W") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "W"
                    Case Else
                        funGetFieldValue = Format(lngAge, "000") & "D"
                End Select
                    
            End If
        Case "出生日期"
            funGetFieldValue = Format(NVL(rsDataSet!出生日期), "YYYYMMDD")
        Case "中文名"
            funGetFieldValue = NVL(rsDataSet!姓名)
        Case "检查设备"
            funGetFieldValue = NVL(rsDataSet!检查设备)
        Case "体重"
            funGetFieldValue = Val(NVL(rsDataSet!体重))
        Case "附加主述"
            funGetFieldValue = NVL(rsDataSet!附加主述)
    End Select
End Function

Private Sub Timer1_Timer()
    
    On Error GoTo err
    
    '当数据断开时重新连接
    CheckDBConnect
    
    '保存剩下的图像到FTP中
    Call ProcSave
    
    '刷新显示列表
    If blnNewImg Then
        ListSeq strWhere
        blnNewImg = False
    End If
    
    '判断当前是否有图像，如果没有图像，而且300秒(5分钟)之内没有Association，则清空Association数组
    If DateDiff("S", mdtLastAssociation, Time) > 300 And DViewer.Images.Count = 0 Then
        ReDim AEconnections(0) As AEconnection
    End If
    
    '判断日志文件是否超过600M，超过则创建新的日志文件
    If FileLen(gstrAccessName) > 600000000 Then
        Call subNewLogFile
    End If
    Exit Sub
err:
    Call WriteLog(5801, err.Number, "Timer 出错，错误描述是：" & err.Description)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub subRemove(instanceUID As String)
    Dim children As DicomDataSets, child As DicomDataSet
    Dim thisDataSet As DicomDataSet
    
    ' Object may already have been removed by delete so trap and ignore errors
    On Error GoTo er1
    Set thisDataSet = PrintRouterDss(instanceUID)
    
    Set children = thisDataSet.Tag
    For Each child In children
        subRemove child.instanceUID
    Next
    
    PrintRouterDss.Remove instanceUID
    
cont:
    Exit Sub
    
er1:
    Resume cont
End Sub

Private Function funGetDataset(uID As String)
    '这个函数可以在不同的数据集合中保存不同的数据集类
    ' this function would allow youto keep different classes of dataset in different collections if you wished
    Set funGetDataset = PrintRouterDss(uID)
End Function

Private Function NewDataSet() As DicomDataSet
    Set NewDataSet = PrintRouterDss.AddNew
    Set NewDataSet.Tag = New DicomDataSets
End Function

Private Sub DecodeFormat(ByVal strFormat As String, ImgCount As Integer)
'根据图像格式字符串,计算出图像的数量
'需要考虑STANDARD，ROW和COL的格式
'常用格式表示方法为：“STANDARD\2,2”，“ROW\1,2,2”，“COL\1,2,2”等

    Dim strNum() As String       '存放格式内图像数量的数组
    Dim strPrintFormat As String
    Dim strImageCount As String
    Dim i As Integer
    
    strPrintFormat = Left(strFormat, InStr(strFormat, "\") - 1)
    strImageCount = Right(strFormat, Len(strFormat) - InStr(strFormat, "\"))
    
    strNum = Split(strImageCount, ",")
    
    On Error Resume Next
    
    If strPrintFormat = "STANDARD" Then
        ImgCount = Val(strNum(0)) * Val(strNum(1))
    ElseIf strPrintFormat = "ROW" Or strPrintFormat = "COL" Then
        For i = 0 To UBound(strNum)
            ImgCount = ImgCount + Val(strNum(i))
        Next i
    Else
        ImgCount = 0
    End If
End Sub


Private Function funSaveFilmImages(ByVal ruid As String, connection As DicomConnection, ByVal iService As Integer)
    '接收胶片，把胶片中的图像全部保存下来
    Dim DFilmBox As DicomDataSet
    Dim DImageBoxs As DicomDataSets
    Dim DSessions As DicomDataSets
    Dim DSession As DicomDataSet
    Dim intImageCount As Integer
    Dim strFormat As String
    Dim DImageds As DicomDataSet
    Dim DImageAtt As DicomAttribute
    Dim DImages As DicomImages
    Dim DImage As DicomImage
    Dim DTempImage As New DicomImage
    Dim strStudyUID As String
    Dim strSeriesUID As String
    Dim curDate As Date
    
    Dim i As Integer
    
    On Error GoTo err1
    
    '根据ruid从公共数据集PrintRouterDss中获得一个指向FilmBox的数据集
    Set DFilmBox = PrintRouterDss(ruid)
    '读取ImageBox序列 Referenced Image Box Sequence
    Set DImageBoxs = DFilmBox.Attributes(&H2010, &H510).value
    '读取session  Referenced Film Session Sequence
    Set DSessions = DFilmBox.Attributes(&H2010, &H500).value
    Set DSession = DSessions(1)
    Set DSession = PrintRouterDss(DSession.Attributes(8, &H1155).value)
    
    '读取图像打印格式
    If DFilmBox.Attributes(&H2010, &H10).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H10)) Then
        strFormat = DFilmBox.Attributes(&H2010, &H10)
    Else
        strFormat = "STANDARD\1,1"
    End If
    '根据格式，计算图像数量
    DecodeFormat strFormat, intImageCount
    
    '读取新创建图像的 检查UID和序列UID
    strStudyUID = DTempImage.StudyUID
    strSeriesUID = DTempImage.SeriesUID
    
    '提前读取数据库时间，避免在下面的循环中多次查询数据库
    curDate = zlDatabase.Currentdate
    
    '循环ImageBoxs 打印每一个图像
    For i = 1 To intImageCount
        '读取图像数据集
        '(8,&h1155) = Referenced SOP Instance UID
        Set DImageds = PrintRouterDss(DImageBoxs(i).Attributes(8, &H1155).value)
        
        '先尝试读取灰度图象 (&H2020, &H110) = Basic Grayscale Image Sequence
        '每一个DImageAtt里面保存一个DicomImage，实际上DImageAtt是从DImageds里面读出来的
        Set DImageAtt = DImageds.Attributes(&H2020, &H110)
        
        '如果灰度图象没有读取成功，则读取彩色图像
        '(&H2020, &H111) = Basic Colour Image Sequence
        If Not DImageAtt.Exists Then
            Set DImageAtt = DImageds.Attributes(&H2020, &H111)
        End If
        
        '找到图象，则开始保存图像
        If DImageAtt.Exists Then
            Set DImages = New DicomImages
            DImages.Add DImageAtt.value.Item(1)
            Set DImage = DImages(1)
            '把图像放到DViewer中，等待保存
            subWriteDicomPara DImage, iService, Str(i), curDate
            DImage.StudyUID = strStudyUID
            DImage.SeriesUID = strSeriesUID
            DImage.Tag = connection.Association
'            DImage.Tag = "胶片接收-" & iService
            
            '保存连接参数
            subSaveAssociation connection
            
            DViewer.Images.Add DImage
        End If
    Next i
    
    blnNewImg = True
    funSaveFilmImages = 0
    Exit Function
err1:
    funSaveFilmImages = 1
End Function

Private Sub subWriteDicomPara(img As DicomImage, ByVal iService As Integer, strImageNum As String, curDate As Date)
'------------------------------------------------
'功能：给输入的图像填写DICOM文件头信息
'参数：img－－输入的DICOM文件
'       iService--服务信息数组的索引
'       strImageNum--图像号
'返回：无，直接文件头信息写入img的文件头
'------------------------------------------------
    Dim g As New DicomGlobal
    
    img.instanceUID = g.NewUID
    img.Attributes.Add &H8, &H8, ""                             'ImageType  空
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID，二次捕捉
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date 检查日期
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date 序列日期
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date 采集日期
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   图像日期
    img.Attributes.Add &H8, &H30, Format(curDate, "HH:MM:SS")     'Study Time   检查时间
    img.Attributes.Add &H8, &H31, Format(curDate, "HH:MM:SS")     'Series Time  序列时间
    img.Attributes.Add &H8, &H32, Format(curDate, "HH:MM:SS")     'Acquisition Time  采集时间
    img.Attributes.Add &H8, &H33, Format(curDate, "HH:MM:SS")     'Image Time  图像时间
    img.Attributes.Add &H8, &H50, ""                            'Accession Number 空
    img.Attributes.Add &H8, &H60, Services(iService).Modality   'Modality 影像类别
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer 厂商
    img.Attributes.Add &H8, &H80, gstr单位名称                  'Institution Name 单位名称
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name 空
    img.Attributes.Add &H8, &H1030, ""                          'Study Description 检查描述 空
    img.Attributes.Add &H10, &H10, ""                           'Name 姓名
    img.Attributes.Add &H10, &H20, ""                           'Patient ID 病人ID
    img.Attributes.Add &H10, &H30, ""                           'BirthDate 生日
    img.Attributes.Add &H10, &H40, ""                           'Sex 性别
    img.Attributes.Add &H10, &H1010, ""                         'Age 年龄
    img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
    img.Attributes.Add &H20, &H10, ""                           'Study ID 检查ID
    img.Attributes.Add &H20, &H11, "1"                          'Series Number 序列号
    img.Attributes.Add &H20, &H13, strImageNum                         'ImageNumber 图像号
    img.Attributes.Add &H20, &H20, ""                           'Orientation 空
End Sub


Private Function funPrintOut(ByVal ruid As String, connection As DicomConnection) As Long
    '把打印的请求转给真实的打印机,或者转发给图像接收
    
    Dim DPrinter As New DicomPrint
    Dim DFilmBox As DicomDataSet
    Dim DImageBoxs As DicomDataSets
    Dim DSessions As DicomDataSets
    Dim DSession As DicomDataSet
    Dim strOrientation As String
    Dim strFilmSize As String
    Dim intCopies As Integer
    Dim intImageCount As Integer
    Dim DImageds As DicomDataSet
    Dim DImageAtt As DicomAttribute
    Dim DImages As DicomImages
    Dim DImage As DicomImage
    Dim iService As Integer
    
    Dim i As Integer
    
    iService = -1
    
    '获取主服务器的IP地址，端口，AE名称等
    If iService = -1 Then iService = funGetServiceIndex(connection.CalledAET)
    
    If iService = -1 Then
        '找不到打印路由的主服务设置，退出打印
        funPrintOut = 1
        Exit Function
    Else
        '如果是“胶片接收”则转到子程序执行
        If Services(iService).SOP = "胶片接收" Then
            funPrintOut = funSaveFilmImages(ruid, connection, iService)
            Exit Function
        End If
        
        DPrinter.Node = Services(iService).DeviceIP
        DPrinter.Port = Services(iService).DevicePort
        DPrinter.CalledAE = Services(iService).DeviceAE
        DPrinter.CallingAE = Services(iService).ServiceAE
    End If
    
    On Error GoTo err1
    
    '根据ruid从公共数据集PrintRouterDss中获得一个指向FilmBox的数据集
    Set DFilmBox = PrintRouterDss(ruid)
    '读取ImageBox序列 Referenced Image Box Sequence
    Set DImageBoxs = DFilmBox.Attributes(&H2010, &H510).value
    '读取session  Referenced Film Session Sequence
    Set DSessions = DFilmBox.Attributes(&H2010, &H500).value
    Set DSession = DSessions(1)
    Set DSession = PrintRouterDss(DSession.Attributes(8, &H1155).value)
    
    
    
    '打印图像的位数，必须
    '在这里无法知道图像的位数，而且作为打印路由，接收到的图像是已经处理好可以直接打印的图像了，
    '因此不再处理打印图像的位数，而是在打印每一个图像的时候，PrintImage中使用Raw=True的参数。
    '使用 ImageBox 中的来读取 (0028,0101) : Bits Stored
    '''''''''''''''''''Printer的直接参数''''''''''''''''''''''''''''''''''
    

    '''''''''''''''''''Session的参数''''''''''''''''''''''''''''''''''
    '打印份数，必须
    If Not IsNull(DSession.Attributes(&H2000, &H10)) Then
        intCopies = DSession.Attributes(&H2000, &H10)           '读取Number of Copies
        DPrinter.Copies = intCopies
    Else
        DPrinter.Copies = 1
    End If
    
    'Print Priority 优先级，可选
    If DSession.Attributes(&H2000, &H20).Exists And Not IsNull(DSession.Attributes(&H2000, &H20)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H20)
    End If
    
    'Medium Type 介质类型，可选
    If DSession.Attributes(&H2000, &H30).Exists And Not IsNull(DSession.Attributes(&H2000, &H30)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H30)
    End If
    
    
    'Film Destination 介质目标，可选
    If DSession.Attributes(&H2000, &H40).Exists And Not IsNull(DSession.Attributes(&H2000, &H40)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H40)
    End If
    
    
    '''''''''''''''''''''''''''''''''''''    '打开打印机'''''''''''''''''''''''''''
    DPrinter.Open
    
    '''''''''''''''''''FilmBox的参数''''''''''''''''''''''''''''''''''
    '胶片方向 ，必须
    If Not IsNull(DFilmBox.Attributes(&H2010, &H40)) Then
        strOrientation = DFilmBox.Attributes(&H2010, &H40)      '读取Film Orientation
        DPrinter.Orientation = strOrientation
    Else
        DPrinter.Orientation = "PORTRAIT"
    End If
    
    '打印格式，必须
    If DFilmBox.Attributes(&H2010, &H10).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H10)) Then
        DPrinter.Format = DFilmBox.Attributes(&H2010, &H10)
    Else
        DPrinter.Format = "STANDARD\1,1"
    End If
    
    '胶片大小,默认值为空，使用打印机的默认值
    If Not IsNull(DFilmBox.Attributes(&H2010, &H50)) Then
        strFilmSize = DFilmBox.Attributes(&H2010, &H50)
        DPrinter.FilmSize = strFilmSize
    End If
    
    '放大方式,必须
    If DFilmBox.Attributes(&H2010, &H60).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H60)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H60, DFilmBox.Attributes(&H2010, &H60)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H60, "CUBIC"
    End If
    
    'Smoothing Type '平滑,可选
    If DFilmBox.Attributes(&H2010, &H80).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H80)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H80, DFilmBox.Attributes(&H2010, &H80)
    End If
    
    'border density 边缘密度，必须
    If DFilmBox.Attributes(&H2010, &H100).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H100)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H100, DFilmBox.Attributes(&H2010, &H100)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H100, "BLACK"
    End If
    
    'empty image density 空白密度，必须
    If DFilmBox.Attributes(&H2010, &H110).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H110)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H110, DFilmBox.Attributes(&H2010, &H110)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H110, "BLACK"
    End If

    '剪切方式
    If DFilmBox.Attributes(&H2010, &H140).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H140)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H140, DFilmBox.Attributes(&H2010, &H140)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H140, "NO"
    End If
        
    'Polarity 极性,可选
    If DFilmBox.Attributes(&H2020, &H20).Exists And Not IsNull(DFilmBox.Attributes(&H2020, &H20)) Then
        DPrinter.FilmBox.Attributes.Add &H2020, &H20, DFilmBox.Attributes(&H2020, &H20)
    End If
        
    'Requested Resolution ID 分辨率，可选
    If DFilmBox.Attributes(&H2020, &H50).Exists And Not IsNull(DFilmBox.Attributes(&H2020, &H50)) Then
        DPrinter.FilmBox.Attributes.Add &H2020, &H50, DFilmBox.Attributes(&H2020, &H50)
    End If
        
    '循环ImageBoxs 打印每一个图像
    DecodeFormat DPrinter.Format, intImageCount
    For i = 1 To intImageCount
        '读取图像数据集
        '(8,&h1155) = Referenced SOP Instance UID
        Set DImageds = PrintRouterDss(DImageBoxs(i).Attributes(8, &H1155).value)
        
        '先尝试读取灰度图象 (&H2020, &H110) = Basic Grayscale Image Sequence
        '每一个DImageAtt里面保存一个DicomImage，实际上DImageAtt是从DImageds里面读出来的
        Set DImageAtt = DImageds.Attributes(&H2020, &H110)
        
        '如果灰度图象没有读取成功，则读取彩色图像
        '(&H2020, &H111) = Basic Colour Image Sequence
        If Not DImageAtt.Exists Then
            Set DImageAtt = DImageds.Attributes(&H2020, &H111)
        End If
        
        '找到图象，则开始打印
        If DImageAtt.Exists Then
            Set DImages = New DicomImages
            DImages.Add DImageAtt.value.Item(1)
            Set DImage = DImages(1)
            DPrinter.PrintImage DImage, True, False
        End If
    Next i
    DPrinter.PrintFilm
    DPrinter.Close
    
    funPrintOut = 0
    Exit Function
err1:
    funPrintOut = 1
    
End Function

Private Sub MakePrinterdataset()
    Dim p As DicomDataSet
    Set p = NewDataSet
    
    p.instanceUID = doInstance_Printer
    p.Attributes.Add 8, &H16, doSOP_Printer             'SOP Class UID
    p.Attributes.Add 8, &H70, "ZLSOFT"                  'Manufacturer
    p.Attributes.Add 8, &H1090, "Demo Printer SCP"      'Manufacturer's Model Name
    p.Attributes.Add &H18, &H1000, "serial no 1234"     'Device Serial Number
    p.Attributes.Add &H18, &H1020, DGlobal.Version            'Software Version(s)
    Set printerobject = p
End Sub

Private Function funCanStartServer() As Boolean
'检查设备数量是否超过限制,超过限制则不允许启动服务
'参数：
'返回值：
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    gint胶片打印机数量 = getLicenseCount(LOGIN_TYPE_胶片打印机)
    gintDICOM设备数量 = getLicenseCount(LOGIN_TYPE_DICOM设备)
    strSQL = "select 设备号,设备名,类型 from 影像设备目录 Where  NVL(状态,0)=1 and (类型=3 Or 类型=4 )"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询设备数量")
    
    rsTemp.Filter = "类型=3"
    If (rsTemp.RecordCount > gint胶片打印机数量 And gint胶片打印机数量 <> -1) Or gint胶片打印机数量 = 0 Then
        MsgBox LOGIN_TYPE_胶片打印机 & "超过您购买的总数量（" & gint胶片打印机数量 & "），服务无法启动。请向软件供应商联系", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rsTemp.Filter = "类型=4"
    If (rsTemp.RecordCount > gintDICOM设备数量 And gintDICOM设备数量 <> -1) Or gintDICOM设备数量 = 0 Then
        MsgBox LOGIN_TYPE_DICOM设备 & "超过您购买的总数量（" & gintDICOM设备数量 & "），服务无法启动。请向软件供应商联系", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    funCanStartServer = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetBodypartValue(lngOrderID As Long, connection As DicomConnection, strBodypartSplitter As String, strBodypartName As String, _
    strBodypartCode As String) As Boolean
'根据医嘱ID，查询对码部位名称和对码部位代码
'参数： lngOrderID【IN】 --- 医嘱ID
'       connection【IN】 --- DICOM连接
'       strBodyPartSplitter【IN】 --- 多部位的分隔符
'       strBodypartName 【OUT】--- 对码部位名称串，用strBodyPartSplitter分隔
'       strBodypartCode 【OUT】--- 对码部位代码串，用strBodyPartSplitter分隔
'返回值：   True-成功；False-失败
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    'PACS部位是“标本部位+检查方法”共同组成部位对码中的“PACS部位名称”
    strSQL = "Select x.部位方法 ,b.设备部位名称, b.设备部位代码 " & _
             " From  (Select a.标本部位||a.检查方法 As 部位方法 From 病人医嘱记录 a Where 相关id = [1]) x, " & _
             " 影像mwl部位对码 b, 影像dicom服务对 c " & _
             " Where x.部位方法 =b.Pacs部位名称 And b.服务id = c.服务id And Upper(c.Pacsae名称) = [2] and c.设备ip地址 = [3] " & _
             " order by 设备部位代码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngOrderID, UCase(CStr(connection.CalledAET)), CStr(connection.RemoteIP))
    
    While rsTemp.EOF = False
        If InStr(strBodypartCode, NVL(rsTemp!设备部位代码)) = 0 Then
            strBodypartName = strBodypartName & strBodypartSplitter & NVL(rsTemp!设备部位名称)
            strBodypartCode = strBodypartCode & strBodypartSplitter & NVL(rsTemp!设备部位代码)
        End If
        rsTemp.MoveNext
    Wend
    
    If strBodypartName <> "" Then
        strBodypartName = Mid(strBodypartName, 2)
    End If
    
    If strBodypartCode <> "" Then
        strBodypartCode = Mid(strBodypartCode, 2)
    End If
    funGetBodypartValue = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    funGetBodypartValue = False
End Function
