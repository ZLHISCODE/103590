VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotify 
   BorderStyle     =   0  'None
   Caption         =   "医嘱提醒"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   Icon            =   "frmNotify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmNotify"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Time_Flash 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1590
      Top             =   60
   End
   Begin VB.Timer TimNotify 
      Interval        =   500
      Left            =   2010
      Top             =   60
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   3630
      Top             =   45
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
            Picture         =   "frmNotify.frx":000C
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":05A6
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":0B40
            Key             =   "warn"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":0EDA
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":1274
            Key             =   "Change"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      ScaleHeight     =   345
      ScaleWidth      =   4305
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   4305
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱提醒"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   90
         Width           =   3735
      End
      Begin VB.Image imgShow 
         Height          =   360
         Left            =   0
         Picture         =   "frmNotify.frx":160E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8040
      Left            =   0
      ScaleHeight     =   8040
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   0
      Width           =   4490
      Begin VB.PictureBox picNotify 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7560
         Left            =   15
         ScaleHeight     =   7560
         ScaleWidth      =   4485
         TabIndex        =   4
         Top             =   480
         Width           =   4490
         Begin XtremeReportControl.ReportControl rptNotify 
            Height          =   7290
            Left            =   0
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   255
            Width           =   4470
            _Version        =   589884
            _ExtentX        =   7885
            _ExtentY        =   12859
            _StockProps     =   0
            BorderStyle     =   1
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.OptionButton optNotify 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            Caption         =   "全病区"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2250
            TabIndex        =   7
            Top             =   15
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optNotify 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            Caption         =   "本人负责"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   1005
            TabIndex        =   6
            Top             =   15
            Width           =   1155
         End
         Begin VB.Label lblNotify 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00EAFFFF&
            Caption         =   "提醒范围："
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   135
            TabIndex        =   8
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱提醒"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgHide 
         Height          =   360
         Left            =   30
         Picture         =   "frmNotify.frx":1D10
         Stretch         =   -1  'True
         Top             =   30
         Width           =   360
      End
   End
   Begin XtremeCommandBars.ImageManager imgIcons 
      Left            =   1350
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNotify.frx":2412
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1350
      Top             =   3480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeft_Form As Long       '窗体的坐标
Private mlngTop_Form As Long        '窗体的坐标
Private mlngLeft_Mouse As Long      '鼠标点击时的坐标
Private mlngTop_Mouse As Long       '鼠标点击时的坐标

Private mintCount As Integer
Private mstrPreNotify As String
Private mlngPreID As Long
Private mblnFirstIn As Boolean      '是否第一次进入医嘱提醒界面区域

Public mblnExecCollapse As Boolean  '执行任务时提醒窗口是否自动折叠
Public mblnNormal As Boolean        'TRUE-最大化;FALSE-最小化
Public mblnOrientation As Boolean   'TRUE-横向;False-纵向
Public mblnFirst As Boolean         '第一次启动立即刷新,或切换病区时强制刷新
Public mlng病区ID As Long
Public mstrScope As String
Public mdtOutBegin As Date, mdtOutEnd As Date
Public mintNotify As Integer '医嘱提醒自动刷新间隔(分钟)
Public mintNotifyDay As Integer '提醒多少天内的医嘱
Public mstrNotifyAdvice As String '提醒的医嘱类型
Public mstrRelatedUnitID As String '整体护理病区ID
Public mbln整体护理消息 As Boolean '控制是否显示整体护理消息

Private mstrBlankTime As String     '作废医嘱时间
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean


Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private mclsPublicAdvice As zlPublicAdvice.clsPublicAdvice

Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人ID = 1
    C_主页ID = 2
    c_姓名 = 3
    c_住院号 = 4
    c_床号 = 5
    C_状态 = 6
    
    '隐藏列
    C_消息 = 7
    C_序号 = 8
    C_日期 = 9
    C_业务 = 10
    C_就诊病区 = 11
    C_唯一标识 = 12 '用于区分消息的唯一性
End Enum

Private Enum EFun_医嘱提醒
    E校对 = 0
    E停止 = 1
End Enum

Private Enum Msg_Type '消息提醒类别
    m新开 = 1
    m新停 = 2
    m新废 = 3
    m安排 = 4
    m危机值 = 5
    m输液拒绝 = 6
    m销帐申请 = 7
    m取血通知 = 8
    m备血完成 = 9
End Enum

Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll " (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'横向
Private Const lngWidth_Normal As Long = 4530
Private Const lngHeight_Normal As Long = 8070
Private Const lngWidthH_Collapse As Long = 4020
Private Const lngHeightH_Collapse As Long = 410
'纵向
Private Const lngWidthV_Collapse As Long = 410
Private Const lngHeightV_Collapse As Long = 4020

Private Const conMenu_最小化 As Long = 15
Private Const conMenu_纵向 As Long = 14
Private Const conMenu_横向 As Long = 13
Private Const conMenu_折叠 As Long = 12
Private Const conMenu_展开 As Long = 11
Private mobjMenu As CommandBarPopup

Private mbytFontSize As Byte

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小 在Form_Load之后调用
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "宋体"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            If UCase(objCtrl.Name) = UCase("lblInfo") Then
                If mblnOrientation = True Then
                    objCtrl.Height = TextHeight("刘") + 20
                Else
                    objCtrl.Width = TextWidth("刘") + 20
                End If
            Else
                objCtrl.Height = TextHeight("刘") + 20
            End If
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth(objCtrl.Caption & "刘刘")
            objCtrl.Height = TextHeight("刘") + 20
            
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        End Select
    Next
    optNotify(0).Left = lblNotify.Left + lblNotify.Width
    optNotify(1).Left = optNotify(0).Left + optNotify(0).Width + 100
    Call Form_Resize
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    mlngLeft_Form = Me.Left
    mlngTop_Form = Me.Top
    
    Select Case Control.ID
    Case conMenu_最小化
        mblnExecCollapse = mblnExecCollapse Xor True
    Case conMenu_纵向
        Call AdjustInfo
        mblnOrientation = False
    Case conMenu_横向
        Call AdjustInfo
        mblnOrientation = True
    Case conMenu_折叠
        mblnNormal = False
    Case conMenu_展开
        mblnNormal = True
    Case conMenu_View_Refresh
        mblnFirst = True
        Exit Sub
    End Select
    Call SetMode
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_最小化
        Control.Checked = mblnExecCollapse
    Case conMenu_纵向
        Control.Checked = Not mblnOrientation
    Case conMenu_横向
        Control.Checked = mblnOrientation
    Case conMenu_折叠
        Control.Checked = Not mblnNormal
    Case conMenu_展开
        Control.Checked = mblnNormal
    End Select
End Sub

'示例:
'其中dwFlags有LWA_ALPHA和LWA_COLORKEY
'LWA_ALPHA被设置的话,通过bAlpha决定透明度.
'LWA_COLORKEY被设置的话 , 则指定被透明掉的颜色为crKey, 其他颜色则正常显示
'因此只要设置LWA_COLORKEY和crKey，并且将窗体背景色和控件颜色设为   不同的颜色，就可以做到楼主的要求，经实际测试可行
'Dim rtn As Long
'rtn = GetWindowLong(hwnd, GWL_EXSTYLE) '取的窗口原先的样式
'rtn = rtn Or WS_EX_LAYERED '使窗体添加上新的样式WS_EX_LAYERED
'SetWindowLong hwnd, GWL_EXSTYLE, rtn '把新的样式赋给窗体
'SetLayeredWindowAttributes hwnd, 要透明的颜色, 0, LWA_COLORKEY
'SetLayeredWindowAttributes hwnd, 0, 透明度, LWA_ALPHA


'如果在左右侧小于窗体宽度时,折叠时以纵向显示

Private Sub Form_Load()
    Dim strCoord As String
    Dim objCol As ReportColumn
    Dim RectState As RECT
    Dim blnStateLR As Boolean  '状态栏是否左右显示
    
    mintCount = 0
    mblnFirst = True
    mblnFirstIn = True
    mstrPreNotify = ""
    mlngPreID = 0
    mstrBlankTime = Format(Now, "MM-dd")
    
    '获取状态栏的位置
    On Error Resume Next
    GetWindowRect FindWindow("Shell_TrayWnd", vbNullString), RectState
    blnStateLR = ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY = Screen.Height)
    err.Clear
    On Error GoTo ErrHand
    '创建象
    Set mclsPublicAdvice = New zlPublicAdvice.clsPublicAdvice
    Call mclsPublicAdvice.InitCommon(gcnOracle, glngSys)
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1265, gstrPrivs)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    mblnExecCollapse = (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AutoCollapse", "1") = 1)
    '取窗口方向
    mblnOrientation = (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "Orient", "1") = 1)
    '如果是横向则不处理,纵向处理方法:横为纵,纵为横
    If Not mblnOrientation Then
        Call AdjustInfo
    End If
    
    '在设置窗体大小前获取上次退出时窗体的位置
    strCoord = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "Coord", "-8000|300")
    mlngLeft_Form = Split(strCoord, "|")(0)
    mlngTop_Form = Split(strCoord, "|")(1)
    
    '79247
    If mlngLeft_Form + Me.Width < 0 Then  '缺省显示在右上角(主要是为了新使用的用户)
        mlngLeft_Form = Screen.Width - Me.Width - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0)
    ElseIf mlngLeft_Form < IIf(blnStateLR = True, IIf(RectState.Left = 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX), 0), 0) Then
        mlngLeft_Form = IIf(blnStateLR = True, IIf(RectState.Left = 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX), 0), 0)
    ElseIf mlngLeft_Form + IIf(mblnOrientation = True, lngWidthH_Collapse, lngWidthV_Collapse + 100) > Screen.Width - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0) Then
        mlngLeft_Form = Screen.Width - IIf(mblnOrientation = True, lngWidthH_Collapse, lngWidthV_Collapse + 100) - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0)
    End If
    
    If mlngTop_Form < IIf(blnStateLR = False, IIf(RectState.Top = 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY), 0), 0) Then
        mlngTop_Form = IIf(blnStateLR = False, IIf(RectState.Top = 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY), 300), 300)
    ElseIf mlngTop_Form + IIf(mblnOrientation = True, lngHeightH_Collapse, lngHeightV_Collapse) > Screen.Height - IIf(blnStateLR = False, IIf(RectState.Top = 0, 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY)), 0) Then
        mlngTop_Form = Screen.Height - IIf(mblnOrientation = True, lngHeightH_Collapse, lngHeightV_Collapse) - IIf(blnStateLR = False, IIf(RectState.Top = 0, 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY)), 0)
    End If
    '设置窗体的大小与显示状态
    Call SetMode
    
    If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "NotifyRang", "0")) = 0 Then
        optNotify(1).Value = True
    Else
        optNotify(0).Value = True
    End If
    
    '90278:按照床位排列
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", 70, True)
        Set objCol = .Columns.Add(c_床号, "床号", 60, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
        
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_就诊病区, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_唯一标识, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有提醒内容..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '排序
        '95547:紧急医嘱优先显示
        .SortOrder.Add .Columns.Find(C_序号)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns.Find(c_床号)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns.Find(C_日期)
        .SortOrder(2).SortAscending = False
    End With

    Call MainDefCommandBar
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AdjustInfo()
    Dim lngTmp As Long
    
    lngTmp = picInfo.Width
    picInfo.Width = picInfo.Height
    picInfo.Height = lngTmp
    
    lngTmp = lblInfo.Width
    lblInfo.Width = lblInfo.Height
    lblInfo.Height = lngTmp
    lngTmp = lblInfo.Top
    lblInfo.Top = lblInfo.Left
    lblInfo.Left = lngTmp
End Sub

Private Sub MainDefCommandBar()
    Dim objControl As CommandBarControl

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgIcons.Icons

    '菜单定义
    '-----------------------------------------------------
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "操作(&O)", -1, False) '固有
    mobjMenu.ID = 1 '对xtpControlPopup类型的命令ID需重新赋值
    mobjMenu.Visible = False
    With mobjMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_横向, "横向")   '打印床头卡
        Set objControl = .Add(xtpControlButton, conMenu_纵向, "纵向")
        Set objControl = .Add(xtpControlButton, conMenu_展开, "展开"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_折叠, "折叠")
        Set objControl = .Add(xtpControlButton, conMenu_最小化, "执行任务时窗口折叠"): objControl.BeginGroup = True
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picNotify.ZOrder 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '收折时保存上次的位置
    If mblnNormal Then
        mlngLeft_Form = Me.Left
        mlngTop_Form = Me.Top
    End If
    
    If Not (mclsPublicAdvice Is Nothing) Then
        Set mclsPublicAdvice = Nothing
    End If
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If

    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AutoCollapse", IIf(mblnExecCollapse, "1", "0"))
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "Orient", IIf(mblnOrientation, "1", "0"))
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "Coord", mlngLeft_Form & "|" & mlngTop_Form)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "NotifyRang", IIf(optNotify(0).Value = True, 1, 0))
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
End Sub

Private Sub imgHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    '收折时保存上次的位置
    mlngLeft_Form = Me.Left
    mlngTop_Form = Me.Top
    
    '设置窗体显示模式
    mblnNormal = False
    Call SetMode
End Sub

Private Sub imgShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picForm_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picForm_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim blnRecToLis As Boolean '是否加载到提醒列表中
    Dim rsMsg As ADODB.Recordset
    On Error GoTo ErrHand
    
    If mlng病区ID = 0 Then Exit Sub
    
    If strMsgItemIdentity = "ZLHIS_TRANSFUSION_001" And Mid(mstrNotifyAdvice, 6, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_001" And Mid(mstrNotifyAdvice, 1, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_002" And Mid(mstrNotifyAdvice, 2, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_003" And Mid(mstrNotifyAdvice, 3, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CHARGE_001" And Mid(mstrNotifyAdvice, 7, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 4, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_LIS_003,ZLHIS_PACS_005,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 5, 1) = "1" Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optNotify_Click(Index As Integer)
    '窗体加载后才点击才刷新, 避免初始化启用该事件
    If picNotify.Visible = True Then mblnFirst = True
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picForm.Tag = 1
        mlngLeft_Mouse = X
        mlngTop_Mouse = Y
    Else
        Call AddMenus
    End If
End Sub

Private Sub AddMenus()
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    '组装右键菜单
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_横向, "横向"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_纵向, "纵向"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_展开, "展开"): cbrPopupItem.Visible = True: cbrPopupItem.BeginGroup = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_折叠, "折叠"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_最小化, "执行任务时窗口折叠"): cbrPopupItem.BeginGroup = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新医嘱"): cbrPopupItem.BeginGroup = True

    cbrPopupBar.ShowPopup
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '移动窗体的位置
    If Button <> 1 Then Exit Sub
    If Val(picForm.Tag) <> 1 Then Exit Sub
    
    Call MoveWindow(Me.hwnd, (Me.Left + X - mlngLeft_Mouse) / Screen.TwipsPerPixelX, (Me.Top + Y - mlngTop_Mouse) / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, True)
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngLeft_Mouse = 0
    mlngTop_Mouse = 0
    picForm.Tag = 0
End Sub

Private Sub SetMode()
    If mblnNormal Then
        Me.Width = lngWidth_Normal
        Me.Height = lngHeight_Normal
        picInfo.Visible = False
        
        picNotify.Visible = True
        picForm.Height = lngHeight_Normal
        picForm.Width = lngWidth_Normal
        picForm.ZOrder 0
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
    Else
        Me.picNotify.Visible = False
        If mblnOrientation Then
            '横向
            Me.Width = lngWidthH_Collapse
            Me.Height = lngHeightH_Collapse
            
            picInfo.Visible = True
            picInfo.Width = Me.Width - 60
            picInfo.ZOrder 0
            picInfo.Refresh
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
        Else
            '纵向
            Me.Width = lngWidthV_Collapse
            Me.Height = lngHeightV_Collapse
            
            picInfo.Visible = True
            picInfo.Height = Me.Height - 60
            picInfo.ZOrder 0
            picInfo.Refresh
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
        End If
        picForm.Height = Me.Height
        picForm.Width = Me.Width
        
    End If
    
    '设置窗体的透明效果
    Call zlControl.PicShowFlat(picForm, 2)
    Call SetTransparence(Not mblnNormal)
    Me.Refresh
End Sub

Private Sub SetTransparence(Optional ByVal blnTransp As Boolean = True)
    Dim rtn As Long
    
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE) '取的窗口原先的样式
    rtn = rtn Or WS_EX_LAYERED '使窗体添加上新的样式WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn '把新的样式赋给窗体
    
    '38595,刘鹏飞,2012-09-10,修改医嘱提醒看不见可点击的情况(被设置为了透明)
    '取消下面背景颜色为&HEAFFFF都设置为透明这段代码，统一改成展开不透明，折叠透明
    'SetLayeredWindowAttributes hwnd, &HEAFFFF, 0, LWA_COLORKEY
    SetLayeredWindowAttributes hwnd, 0, IIf(blnTransp, 180, 255), LWA_ALPHA
End Sub

Private Sub picinfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '设置窗体显示模式
    If Button = 1 Then
        mblnNormal = True
        Call SetMode
    Else
        Call AddMenus
    End If
End Sub

Private Sub picinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    If mblnNormal Then Exit Sub
    
    '--判断当前鼠标位置是否在菜单上--
    MouseOver = (0 <= X) And (X <= picInfo.Width) And (0 <= Y) And (Y <= picInfo.Height)
    If MouseOver Then
        Call SetCapture(picInfo.hwnd)
        If mblnFirstIn Then
            Call SetTransparence(False)
            mblnFirstIn = False
        End If
    Else
        Call ReleaseCapture
        Call SetTransparence(True)
        mblnFirstIn = True
    End If
End Sub

Private Sub Time_Flash_Timer()
    '有新消息，闪烁三次停止
    Time_Flash.Enabled = False
    mintCount = mintCount + 1
    
    If mintCount Mod 2 = 0 Then
        lblInfo.ForeColor = 0
        lblTitle.ForeColor = 0
        If Not mblnNormal Then Call SetTransparence(True)
    Else
        lblInfo.ForeColor = 255
        lblTitle.ForeColor = 255
        If Not mblnNormal Then Call SetTransparence(False)
    End If

    Time_Flash.Enabled = True
    If mintCount = 10 Then
        mintCount = 0
        '49547,刘鹏飞,2012-09-05,如果存在医嘱需要一直提醒
        'Time_Flash.Enabled = False
    End If
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Static strPreTime2 As String
    Dim curTime As Date
    
    curTime = Now
    If gbln启用影像信息系统预约 Or mbln整体护理消息 Then
        If strPreTime2 = "" Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime2), curTime) > 300 Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            If gbln启用影像信息系统预约 = True Then
                If mclsPublicAdvice.GetMsgRISReady(mlng病区ID) Then
                    Call LoadNotify
                Else
                    If mbln整体护理消息 = True Then GoTo NurseMsg
                End If
            Else
NurseMsg:
                Call LoadNurseIntegrateMsg
                Call SetNotifyState
            End If
        End If
    End If
    
    If mbln消息语音 Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                TimNotify.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                TimNotify.Enabled = True
            End If
        End If
    End If
    
    '刷新病历审查提醒
    If mintNotify > 0 Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Or mblnFirst Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            '启用消息平台则不按照固定刷新时限为准
            If mclsMipModule.IsConnect = False Or mblnFirst Then
                strPreTime2 = "" '消息刷新重新记时
                Call LoadNotify
            End If
            mblnFirst = False
        End If
     Else
        If mblnFirst = True Then
            strPreTime2 = "" '消息刷新重新记时
            Call LoadNotify
            mblnFirst = False
        End If
    End If
    
    '测试是否同步刷新时使用
'    If Right(Format(Now, "mm:ss"), 2) Mod 5 = 0 Then
'        Time_Flash.Enabled = True
'    End If
End Sub

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset, rsOut As New ADODB.Recordset
    Dim objOut As Collection, intType As Integer
    Dim strTmp As String, strSQL As String, strTmpRIS As String
    Dim i As Long, blnOk As Boolean
    
    lblTitle.Caption = IIf(mbln整体护理消息 = True, "消息提醒：", "医嘱提醒：")
    lblInfo.Caption = lblTitle.Caption
    
    Screen.MousePointer = 11
    On Error GoTo errH
    blnOk = mclsPublicAdvice.GetAdviceRemind(rsTmp, mlng病区ID, IIf(optNotify(0).Value = True, UserInfo.姓名, ""))
    Screen.MousePointer = 0
    If blnOk = False Then Exit Function
    rptNotify.Records.DeleteAll
    If rsTmp Is Nothing Then GoTo GOEND
    If rsTmp.State = adStateClosed Then GoTo GOEND
    
    '90256:出院和转科医嘱要求显示图标(只针对新开的医嘱)
    strTmp = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For i = 1 To rsTmp.RecordCount
        If rsTmp!类型编码 = "ZLHIS_CIS_001" And InStr("," & strTmp & ",", "," & rsTmp!病人ID & ":" & rsTmp!主页ID & ",") = 0 Then
            strTmp = strTmp & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
        End If
        rsTmp.MoveNext
    Next
    If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
    Set objOut = New Collection
    If strTmp <> "" Then
        '92088:当病人和婴儿同时存在出院医嘱且时间相同，此SQL将过滤出两条数据,故加Distinct(objOut.Add 中也加以判断)
        strSQL = "Select /*+ RULE*/ " & vbNewLine & _
            " Distinct a.病人id, a.主页id, First_Value(b.操作类型) Over(Partition By a.病人id, a.主页id Order By a.开嘱时间 Desc) As 医嘱类型" & vbNewLine & _
            " From 病人医嘱记录 a, 诊疗项目目录 b, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) c" & vbNewLine & _
            " Where a.诊疗项目id + 0 = b.Id And a.诊疗类别 = 'Z' And Instr(',3,5,11,', ',' || b.操作类型 || ',') > 0 And a.医嘱状态 = 1 And" & vbNewLine & _
            "      a.病人id = c.C1 And a.主页id = c.C2"
        Set rsOut = zlDatabase.OpenSQLRecord(strSQL, "提取医嘱信息", strTmp)
        For i = 1 To rsOut.RecordCount
            '不存在才添加
            If GetOutType(objOut, rsOut!病人ID & "_" & rsOut!主页ID) = 0 Then
                objOut.Add Decode(Val(NVL(rsOut!医嘱类型, 0)), 3, 4, 5, 3, 11, 3, 0), rsOut!病人ID & "_" & rsOut!主页ID
            End If
            rsOut.MoveNext
        Next i
    End If
    strTmp = ","
    strTmpRIS = ","
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!类型编码 & ""
        Case "ZLHIS_PACS_006", "ZLHIS_PACS_007"
            'ZLHIS_PACS_006 ZLHIS_PACS_007 消息为一条消息（按医嘱项目为单位）显示一行
            If InStr(strTmpRIS, "," & rsTmp!类型编码 & "," & rsTmp!业务标识 & ",") = 0 Then
                strTmpRIS = strTmpRIS & rsTmp!类型编码 & "," & rsTmp!业务标识 & ","
                intType = 0
                Call AddReportRow(intType, rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, NVL(rsTmp!姓名), NVL(rsTmp!住院号), NVL(rsTmp!床号), NVL(rsTmp!消息内容), _
                    rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "", rsTmp!病人来源 & "", NVL(rsTmp!险类, 0), NVL(rsTmp!就诊病区id, 0), rsTmp!类型编码 & "," & rsTmp!业务标识)
            End If
        Case Else
            If InStr(strTmp, "," & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ",") = 0 Then
                strTmp = strTmp & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码 & ","
                intType = 0
                If rsTmp!类型编码 = "ZLHIS_CIS_001" Then intType = GetOutType(objOut, rsTmp!病人ID & "_" & rsTmp!主页ID)
                Call AddReportRow(intType, rsTmp!病人ID & "," & rsTmp!主页ID, rsTmp!病人ID, rsTmp!主页ID, NVL(rsTmp!姓名), NVL(rsTmp!住院号), NVL(rsTmp!床号), NVL(rsTmp!消息内容), _
                    rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!业务标识 & "", rsTmp!病人来源 & "", NVL(rsTmp!险类, 0), NVL(rsTmp!就诊病区id, 0), rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!类型编码)
            End If
        End Select
        rsTmp.MoveNext
    Next
    
GOEND:
    Call LoadNurseIntegrateMsg '刷新移动护理消息
    
    Call SetNotifyState
    
    LoadNotify = True
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p住院护士站)) = 1
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(2)
        End If
        If Not rsTmp Is Nothing Then
            If Not rsTmp.State = adStateClosed Then
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    Set mrsMsg = rsTmp
                End If
            End If
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOutType(ByVal objOut As Collection, ByVal strKey As String) As Integer
    Dim intType As Integer
    On Error Resume Next
    intType = Val(objOut(strKey))
    If err <> 0 Then err.Clear
    GetOutType = intType
End Function

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!提醒场合, 3, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!部门IDs & ",", "," & mlng病区ID & ",") > 0 Or _
        InStr("," & rsMsg!提醒人员 & ",", "," & UserInfo.姓名 & ",") > 0 Then
        
        '判断列表是否已经有这类消息了，不放 AddReportRow 中判断，这样可能会减少一次SQL查询
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊id) Then
                    Exit Sub
                End If
            End If
        Next
        
        strSQL = "Select a.住院号, a.姓名, a.性别, a.年龄, a.当前床号 As 床号, a.险类,B.当前病区ID 就诊病区id From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID  And  B.病人id =[1] and B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!病人ID), Val(rsMsg!就诊id))
        
        If Not rsTmp.EOF Then
            Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊id, rsMsg!病人ID, rsMsg!就诊id, NVL(rsTmp!姓名), NVL(rsTmp!住院号), NVL(rsTmp!床号), NVL(rsMsg!消息内容), _
                 rsMsg!类型编码 & "", rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "", rsMsg!病人来源 & "", NVL(rsTmp!险类, 0), NVL(rsTmp!就诊病区id, 0), rsMsg!病人ID & "," & rsMsg!就诊id & "," & rsMsg!类型编码)
            Call SetNotifyState
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ByVal intType As Integer, ParamArray arrInput() As Variant)
'功能：向消息提配列表中增加一行
'intType:新开医嘱的类型(3:出院,4-转科)

    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strRowID As String '提醒列表行的唯一标识，"病人id,主页id,消息编码"
    Dim strNO As String
    Dim str业务 As String
    Dim str病人来源 As String
    Dim int优先级 As Integer
    Dim int险类 As Integer
    Dim Index As Integer
    Dim objItemIcon As ReportRecordItem
    
    On Error GoTo errH
    
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值
    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '就诊id
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))): Index = Index + 1 '姓名
    If intType = 3 Or intType = 4 Then objItem.Icon = intType '图标序号:3-出院,4-转科
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '住院号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(zlStr.Lpad(CStr(arrInput(Index)), 10, " ")) '床号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '消息编号
    objRecord.AddItem strNO: Index = Index + 1
    
    int优先级 = Val(arrInput(Index))                     '序号
    objRecord.AddItem int优先级: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '日期
    
    str业务 = arrInput(Index): Index = Index + 1              '业务标识
    str病人来源 = arrInput(Index): Index = Index + 1          '病人来源
    int险类 = arrInput(Index)
    
    If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_003,", "," & strNO & ",") > 0 Then '危机值消息特殊处理，阅读时触发消息
        objRecord.AddItem str业务 & "," & Val(str病人来源)
    Else
        objRecord.AddItem str业务
    End If
    
    Index = Index + 1
    objRecord.AddItem Val(arrInput(Index))    '病区ID
    Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index))    '消息唯一标识
    
    If int优先级 > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int优先级 = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
        If (strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002") And int优先级 = 2 Then objItemIcon.Icon = 2
    End If
    '保险病人用红色显示
    If int险类 > 0 And int优先级 <> 3 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            objRecord.Item(Index).ForeColor = &HC0&
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetNotifyState()
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    mstrPreNotify = IIf(rptNotify.Rows.Count = 0, "", rptNotify.Rows.Count)
    If rptNotify.Rows.Count > 0 Then
        lblTitle.Caption = IIf(mbln整体护理消息 = True, "消息提醒：", "医嘱提醒：") & "共有" & rptNotify.Rows.Count & "条消息需要处理"
        lblInfo.Caption = lblTitle.Caption
        Time_Flash.Enabled = ((mstrPreNotify <> "") Or (mlngPreID <> mlng病区ID))
    Else '没有医嘱信息将停止闪烁
        lblTitle.Caption = IIf(mbln整体护理消息 = True, "消息提醒：", "医嘱提醒：")
        lblInfo.Caption = lblTitle.Caption
        Time_Flash.Enabled = False
        lblInfo.ForeColor = 0
        lblTitle.ForeColor = 0
    End If
    mlngPreID = mlng病区ID
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：自动进入医嘱校对、确认停止的执行界面
    Dim blnExecute As Boolean
    Dim intFunc As Integer
    Dim str业务 As String
    Dim strSQL As String
    Dim objControl As CommandBarControl
    Dim strPrivs As String, strPatis As String
    Dim blnOnePati As Boolean
    Dim lng病人ID As Long, lng主页ID As Long
    Dim blnCollateAutoFind As Boolean
    Dim blnTmp As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strNoteKey As String '消息唯一标识
    Dim blnNurseIntegrate As Boolean '是否是整体护理消息
    
    On Error GoTo ErrHand
    blnCollateAutoFind = (Val(zlDatabase.GetPara("医嘱处理后自动定位到医嘱页面", glngSys, 1265, 0)) = 1)
    strNoteKey = ""
    intFunc = -1
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                str业务 = .Item(C_业务).Value
                lng病人ID = Val(.Item(C_病人ID).Value)
                lng主页ID = Val(.Item(C_主页ID).Value)
                strNoteKey = .Item(C_唯一标识).Value
                If InStr(",ZLHIS_CIS_001,ZLHIS_CIS_002,", .Item(C_消息).Value) > 0 Then
                    strPrivs = GetInsidePrivs(p住院医嘱发送)
                    If .Item(C_消息).Value = "ZLHIS_CIS_001" Then
                        If Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, 0)) = 1 Then
                            intFunc = 0
                        Else
                            intFunc = 1
                        End If
                    ElseIf .Item(C_消息).Value = "ZLHIS_CIS_002" Then
                        intFunc = 2
                    End If
                Else
                    strTmp = ""
                    '55430:刘鹏飞,2013-02-27,双击作废医嘱定位到病人事物的医嘱页面,护士站不能处理危急值消息
                    Select Case .Item(C_消息).Value
                        Case "ZLHIS_BLOOD_003", "ZLHIS_BLOOD_001", "ZLHIS_BLOOD_007" '取血提醒,备血完成提醒,血袋回收提醒
                            intFunc = 3
                        Case "ZLHIS_CIS_003" '作废医嘱
                            intFunc = 3
                        Case "ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015" '安排提醒
                            intFunc = -1
                        Case "ZLHIS_TRANSFUSION_001" '输液审核未通过
                            intFunc = 11
                        Case "ZLHIS_CHARGE_001" '费用销帐申请
                            intFunc = 12
                        Case "ZLHIS_LIS_003" '检验危急值
                            'strTmp = "ZLHIS_CIS_014"
                            Exit Sub
                        Case "ZLHIS_PACS_005" '检查危急值
                            'strTmp = "ZLHIS_CIS_025"
                            Exit Sub
                        Case "ZLHIS_NURSE_INTEGRATE" '整体护理消息
                            blnNurseIntegrate = True
                    End Select
                    If strTmp <> "" And blnNurseIntegrate = False Then
                        If Not (mclsMipModule Is Nothing) Then
                            If mclsMipModule.IsConnect Then
                                strSQL = "select 出院科室ID,当前病区ID from 病案主页 where 病人ID=[1] and 主页ID=[2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Item(C_病人ID).Value), Val(.Item(C_主页ID).Value))
                                Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, .Item(C_病人ID).Value, .Item(c_姓名).Value, .Item(c_住院号).Value, , Val(Split(str业务, ",")(1)), _
                                        .Item(C_主页ID).Value, Val(rsTmp!当前病区ID & ""), Val(rsTmp!出院科室ID & ""), .Item(c_床号).Value, Val(Split(str业务, ",")(0)))
                            End If
                        End If
                    End If
                End If
                strSQL = ""
                If blnNurseIntegrate = False Then
                    '更新消息阅读状态(业务消息相关表，过程为公共部分，不用单独授权)
                    If .Item(C_消息).Value = "ZLHIS_PACS_006" Or .Item(C_消息).Value = "ZLHIS_PACS_007" Then
                        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & .Item(C_消息).Value & "',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ",null,null,'" & .Item(C_业务).Value & "')"
                    ElseIf .Item(C_消息).Value = "ZLHIS_BLOOD_007" And gbln血库系统 Then     '未回收前不允许设为已读
                        If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                        If gobjPublicBlood.zlIsBloodMessageDone(1, lng病人ID, lng主页ID, 3, mlng病区ID) Then
                            If strNoteKey <> "" Then
                                Call ReMoveItemByKey(strNoteKey)
                                Call SetNotifyState
                            End If
                            If intFunc > -1 Then
                                Call frmSublimeInNurseStation.ExecFuncs(intFunc)
                            End If
                        End If
                        Exit Sub
                    Else
                        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & .Item(C_消息).Value & "',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ")"
                    End If
                Else
                    Call ReadNurseIntegrateMsg(strNoteKey)
                    Exit Sub
                End If
            End With
        End If
    End If
    
    If intFunc > -1 Then
        If mblnExecCollapse Then Call imgHide_MouseDown(1, 0, 0, 0)
    End If
    
    Select Case intFunc
    Case 0, 1
        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
            If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
            Call ReMoveItemByKey(strNoteKey)
            Call SetNotifyState
        Else
            If intFunc = 0 Then
                If InStr(strPrivs, ";发送药疗临嘱;") > 0 Or InStr(strPrivs, ";发送药疗长嘱;") > 0 Or InStr(strPrivs, ";发送其他临嘱;") > 0 Or InStr(strPrivs, ";发送其他长嘱;") > 0 Then
                    Call mclsPublicAdvice.AdviceSend(Me, mlng病区ID, lng病人ID, lng主页ID, gstrPrivs, mclsMipModule)
                    
                    If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
                        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
                        Call ReMoveItemByKey(strNoteKey)
                        Call SetNotifyState
                    End If
                    If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
                End If
            ElseIf intFunc = 1 Then
                If InStr(strPrivs, ";医嘱校对处理;") > 0 Then
                    blnOnePati = Val(zlDatabase.GetPara("批量医嘱校对", glngSys, p住院医嘱发送)) = 0
                    blnTmp = mclsPublicAdvice.AdviceOperate(Me, gstrPrivs, 3, lng病人ID, lng主页ID, mlng病区ID, Val(str业务), mclsMipModule, strPatis, blnOnePati)
                    If strPatis <> "" And blnTmp Then Call BatchRemove(strPatis)
                    If Not blnTmp Then
                        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
                            If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
                            Call ReMoveItemByKey(strNoteKey)
                            Call SetNotifyState
                        End If
                    End If
                    If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
                End If
            End If
        End If
    Case 2
        If InStr(strPrivs, ";医嘱确认停止;") > 0 Then
            If Not HaveOperateAdvice(lng病人ID, lng主页ID, 1) Then
                If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
                Call ReMoveItemByKey(strNoteKey)
                Call SetNotifyState
            Else
                blnTmp = mclsPublicAdvice.AdviceOperate(Me, gstrPrivs, 2, lng病人ID, lng主页ID, mlng病区ID, Val(str业务), mclsMipModule, strPatis, True)
                If strPatis <> "" And blnTmp Then
                    Call ReMoveItemByKey(strNoteKey)
                    Call SetNotifyState
                End If
                If Not blnTmp Then
                    If Not HaveOperateAdvice(lng病人ID, lng主页ID, 1) Then
                        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
                        Call ReMoveItemByKey(strNoteKey)
                        Call SetNotifyState
                    End If
                End If
                If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
            End If
        End If
    Case Else
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_业务消息清单_Read")
        If strNoteKey <> "" Then
            Call ReMoveItemByKey(strNoteKey)
            Call SetNotifyState
        End If
        If intFunc > -1 Then
            Call frmSublimeInNurseStation.ExecFuncs(intFunc)
        End If
    End Select
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim strBed As String, strKey As String, strNoteKey As String
    Dim strNO As String
    Dim lng就诊病区ID As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    strNO = CStr(Trim(rptNotify.SelectedRows(0).Record.Item(C_消息).Value))
    lng病人ID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_病人ID).Value))
    lng主页ID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_主页ID).Value))
    lng就诊病区ID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_就诊病区).Value))
    
    strBed = Trim(rptNotify.SelectedRows(0).Record.Item(c_床号).Value)
    strKey = Trim(rptNotify.SelectedRows(0).Record.Item(C_病人ID).Value) & "|" & Trim(rptNotify.SelectedRows(0).Record.Item(C_主页ID).Value)
    strNoteKey = Trim(rptNotify.SelectedRows(0).Record.Item(C_唯一标识).Value)
    
    If ReadAndSendMsg(strNO, lng病人ID, lng主页ID, lng就诊病区ID) Then
        Call ReMoveItemByKey(strNoteKey)
        If rptNotify.Records.Count = 0 Then
            Call imgHide_MouseDown(1, 0, 0, 0)
        End If
        Call SetNotifyState
        Exit Sub
    End If
    
    Call frmSublimeInNurseStation.SelPatiCard(strBed, strKey)
    
    rptNotify.SetFocus
End Sub


Private Function ReadAndSendMsg(ByVal strNO As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng就诊病区ID As Long) As Boolean
    '功能：新开消息时，该消息的病人已经不再当前病区，则先将消息设为已读，再重新发送消息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL() As String
    Dim lng当前科室ID As Long
    Dim lng当前病区ID As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    
    strSQL = "select nvl(A.当前科室ID,0) as 当前科室ID, nvl(A.当前病区ID,0) as 当前病区ID from 病人信息 A where A.病人ID = [1] and 主页ID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)

    If rsTmp.EOF Then Exit Function
    
    lng当前科室ID = Val(rsTmp!当前科室id)
    lng当前病区ID = Val(rsTmp!当前病区ID)
    
    If lng就诊病区ID <> lng当前病区ID And lng当前病区ID <> 0 Then
        If strNO <> "ZLHIS_CIS_001" Then Exit Function
        If Not HaveOperateAdvice(lng病人ID, lng主页ID, 0) Then
            '设置消息为已读
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "','0010','" & _
            UserInfo.姓名 & "'," & lng就诊病区ID & ")"
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans: blnTrans = False
        Else
            strSQL = "select A.消息内容, A.提醒场合,A.类型编码,A.业务标识,A.优先程度 From 业务消息清单 A Where a.病人id=[1] And a.就诊id=[2] And a.类型编码 =[3] and a.就诊病区ID =[4]  And a.是否已阅=0 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, strNO, lng就诊病区ID)
            If rsTmp.RecordCount > 0 Then
                For i = 0 To rsTmp.RecordCount - 1
                    ReDim Preserve arrSQL(i)
                    arrSQL(UBound(arrSQL)) = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng主页ID & "," & lng当前科室ID & "," & lng当前病区ID & ",2,'" & rsTmp!消息内容 & "','" & rsTmp!提醒场合 & "','" & rsTmp!类型编码 & "','" & rsTmp!业务标识 & "'," & rsTmp!优先程度 & ",0,null," & lng当前病区ID & ",null)"
                Next
            End If
            
            '设置消息为已读
            strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng主页ID & ",'" & strNO & "','0010','" & _
            UserInfo.姓名 & "'," & lng就诊病区ID & ")"
            
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '重新发送消息
            If UBound(arrSQL) <> -1 Then
                For i = 0 To UBound(arrSQL)
                    zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
                Next
            End If
            gcnOracle.CommitTrans: blnTrans = False
        End If
        ReadAndSendMsg = True
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BatchRemove(ByVal strPatis As String, Optional ByVal blnNurseIntegrateMsg As Boolean = False)
    '功能：批量移除医嘱提示信息(以病人为单位)。或整体护理消息。当blnNurseIntegrateMsg=True 时，则医嘱整体护理消息
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If blnNurseIntegrateMsg = False Then
                If InStr(";" & strPatis & ";", ";" & objRow.Record.Tag & ";") > 0 And objRow.Record(C_消息).Value = "ZLHIS_CIS_001" Then
                    strIndexs = strIndexs & "," & objRow.Record.Index
                End If
            Else
                If objRow.Record(C_消息).Value = "ZLHIS_NURSE_INTEGRATE" Then
                    strIndexs = strIndexs & "," & objRow.Record.Index
                End If
            End If
        End If
    Next
    If strIndexs <> "" Then
        strIndexs = Mid(strIndexs, 2)
        arrTmp = Split(strIndexs, ",")
        For i = UBound(arrTmp) To 0 Step -1
            Call rptNotify.Records.RemoveAt(Val(arrTmp(i)))
        Next
        Call SetNotifyState
    End If
End Sub

Private Sub ReMoveItemByKey(ByVal strNoteKey As String)
'功能：根据消息列表唯一标识列内容，移除对应消息
    Dim objRow As ReportRow
    If strNoteKey = "" Then Exit Sub
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            '唯一标识列肯定只有一列
            If objRow.Record(C_唯一标识).Value = strNoteKey Then
                Call rptNotify.Records.RemoveAt(objRow.Record.Index)
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub LoadNurseIntegrateMsg()
'功能：读取整体护理消息列表
    Dim strMsg As String, strErrMsg As String
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    Dim i As Integer
    
    '消息节点数据属性
    Dim strID As String, strPatientID As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strName As String, strPatiNo As String, strBedNo As String, strContent As String, lng就诊病区ID As Long, int险类 As Integer
    Dim strCreateTime As String, strToUser As String, strRetrun As String
    Dim objPati As Collection
    Dim blnAdd As Boolean
    
    If mbln整体护理消息 = True Then
        If InitNurseIntegrate = True Then
            If gobjNurseIntegrate.GetMsg(mstrRelatedUnitID, strMsg, strErrMsg) = True Then
                '添加前先删除整体护理的消息信息
                Call BatchRemove("", True)
                If objXML.loadXML(strMsg) = False Then Exit Sub
                Set objNodeList = objXML.selectNodes(".//List//Msg")
                'XML返回格式
                Set objPati = New Collection
                For i = 0 To objNodeList.length - 1
                    strID = objNodeList.Item(i).childNodes(0).Text
                    strName = objNodeList.Item(i).childNodes(1).Text
                    strBedNo = objNodeList.Item(i).childNodes(4).Text
                    strContent = objNodeList.Item(i).childNodes(5).Text
                    strCreateTime = objNodeList.Item(i).childNodes(7).Text
                    strPatientID = objNodeList.Item(i).childNodes(9).Text
                    lng病人ID = Val(objNodeList.Item(i).childNodes(10).Text)
                    lng主页ID = Val(objNodeList.Item(i).childNodes(11).Text)
                    strToUser = objNodeList.Item(i).childNodes(13).Text
                    
                    blnAdd = False
                    If optNotify(0).Value = True And UCase(strToUser) = UCase(UserInfo.用户名) Then
                        blnAdd = True
                    Else
                        blnAdd = True
                    End If
                    If blnAdd = True Then
                        '根据病人ID获取移动相关数据
                        strRetrun = GetPatiData(lng病人ID, lng主页ID, objPati)
                        If strRetrun <> "" Then
                            strPatiNo = Split(strRetrun, "'")(3)
                            lng就诊病区ID = Val(Split(strRetrun, "'")(4))
                            int险类 = Val(Split(strRetrun, "'")(6))
                            Call AddReportRow(0, lng病人ID & "," & lng主页ID, lng病人ID, lng主页ID, strName, strPatiNo, strBedNo, strContent, _
                                "ZLHIS_NURSE_INTEGRATE" & "", 1 & "", Format(strCreateTime & "", "yyyy-MM-dd HH:mm:ss"), strID & "", 2 & "", int险类, lng就诊病区ID, strID)
                        Else
                            Call AddReportRow(0, lng病人ID & "," & lng主页ID, lng病人ID, lng主页ID, strName, strPatiNo, strBedNo, strContent, _
                                "ZLHIS_NURSE_INTEGRATE" & "", 1 & "", Format(strCreateTime & "", "yyyy-MM-dd HH:mm:ss"), strID & "", 2 & "", 0, mlng病区ID, strID)
                        End If
                    End If
                Next i
            Else
                MsgBox "获取整体护理病区ID接口调用失败！" & vbCrLf & "详细信息：" & strErrMsg, vbInformation, gstrSysName
            End If
        End If
    Else
        '清除整体护理消息
        Call BatchRemove("", True)
    End If
End Sub

Private Function GetPatiData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, objPati As Collection) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strRetrun As String, strKey As String
    Dim strSQL As String
    
    On Error Resume Next
    strKey = lng病人ID & "_" & lng主页ID
    strRetrun = objPati(strKey)
    If err <> 0 Then
        err.Clear
        On Error GoTo ErrHand
        strSQL = "Select 姓名,性别,年龄,住院号,当前病区ID,出院病床,险类 From 病案主页 where 病人ID=[1] And 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病人信息", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            strRetrun = rsTemp!姓名 & "'" & rsTemp!性别 & "'" & rsTemp!年龄 & "'" & rsTemp!住院号 & "'" & NVL(rsTemp!当前病区ID, 0) & "'" & rsTemp!出院病床 & "'" & NVL(rsTemp!险类, 0)
            objPati.Add strRetrun, strKey
        End If
    End If
    GetPatiData = strRetrun
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ReadNurseIntegrateMsg(ByVal strID As String)
    Dim strErrMsg As String
    If InitNurseIntegrate = True Then
        If gobjNurseIntegrate.ReplyMsg(strID, strErrMsg) = True Then
            Call ReMoveItemByKey(strID)
            Call SetNotifyState
        Else
            MsgBox strErrMsg
        End If
    End If
End Sub
