VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "药品自动化设备管理工具"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   10080
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picTray 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   460
      Left            =   1680
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picLog 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1440
      ScaleHeight     =   1815
      ScaleWidth      =   2415
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfLog 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _cx             =   2355
         _cy             =   1720
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":1994
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":2226
      Left            =   1200
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":223A
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_BILL As String = "时间,,3,2000,dt|日志信息,,3,7000"

Private mobjPopupMenu  As CommandBar
Private mblnShow As Boolean
Private mblnMenuClose As Boolean
Private mblnAction As Boolean
Private mtypParams As TYPE_PARAMS
Private WithEvents mclsVSF As clsVSFlexGridEx
Attribute mclsVSF.VB_VarHelpID = -1
Private WithEvents mclsTransmit As zlDrugMachineTimer.clsDataTransmit
Attribute mclsTransmit.VB_VarHelpID = -1
Private mstrPrivs As String
Private mstrSupportData As String

'Public mobjSOAP As Object
Public mobjHTTP As Object

Public Property Get SupportData() As String
    SupportData = mstrSupportData
End Property

Private Sub InitCommandbars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003 'xtpthemeoffice2000有凹凸感
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        .EnableCustomization False
        Set .Icons = Me.imgMain.Icons
        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
'    picLine01_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
'    picLine02.BackColor = picLine01_S.BackColor
    
    '文件
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.文件, "文件(&F)", -1, False)
    With cbpTmp
        .ID = enuMenus.文件
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印设置, "打印设置(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印预览, "打印预览(&V)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印, "打印")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.输出Excel, "输出到&Excel...")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.参数设置, "参数设置")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.退出, "退出")
        cbcTmp.BeginGroup = True
    End With
    
    '操作
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.操作, "操作(&O)", -1, False)
    With cbpTmp
        .ID = enuMenus.操作
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.设备接口管理, "设备接口管理(&M)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.基础数据传送, "基础数据传送(&D)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.启用, "定时传送开启(&S)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.停用, "定时传送停止(&P)")
    End With
    
    '查看
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.查看, "查看(&V)", -1, False)
    With cbpTmp
        .ID = enuMenus.查看
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.工具栏, "工具栏(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.标准按钮, "标准按钮(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.文本标签, "文本标签(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大图标, "大图标(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.状态栏, "状态栏(&S)")
        cbcTmp.BeginGroup = True
    End With
    
    '帮助
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.帮助, "帮助(&H)", -1, False)
    With cbpTmp
        .ID = enuMenus.帮助
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB上的中联, "&WEB上的中联")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联主页, "中联主页(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联论坛, "中联论坛(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.发送反馈, "发送反馈(&K)")
'        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.关于, "关于(&A)")
'        cbcTmp.BeginGroup = True
    End With
    
    '弹出菜单
    Set mobjPopupMenu = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopupMenu
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.显示, "显示")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.隐藏, "隐藏")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.退出, "退出")
        cbcTmp.BeginGroup = True
    End With
    
    '菜单项的快键绑定
    With cbsMain.KeyBindings
        .Add 8, vbKeyP, enuMenus.打印
        .Add 8, vbKeyX, enuMenus.退出
        .Add 0, vbKeyF12, enuMenus.参数设置
        .Add 0, vbKeyF1, enuMenus.帮助主题
    End With
    
    '定义工具栏
    Set cbrTmp = cbsMain.Add("工具栏", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.设备接口管理, "设备接口管理")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.基础数据传送, "基础数据传送")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.启用, "定时传送开启")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.停用, "定时传送停用")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.退出, "退出")
        cbcTmp.BeginGroup = True
    End With
    
    '有图标，无文本的按钮风格
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.Type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIconAndCaption
        End If
    Next
    
End Sub

Private Sub InitDockPane()
    Dim panLeft As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
        
        Set panLeft = .CreatePane(1, 250, 0, DockLeftOf)
        With panLeft
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = "日志信息"
        End With
    End With
End Sub

Private Sub InitVSF()
    With mclsVSF
        .Bunding = vsfLog
        .Init
        .Head = MSTR_BILL
        .ColsReadonly = ""
        .Editable = EM_Display
        .Repaint RT_Columns
    End With
    With vsfLog
        .RowHeight(0) = 350
        .ExplorerBar = flexExNone       '不支持鼠标点列头排序
    End With
End Sub

Private Sub InitTray()
    With picTray
        .Top = -.Height
        .Left = -.Width
    End With
    mdlTray.AddIcon picTray, App.ProductName
    App.TaskVisible = True
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim objPrint As Object
    Dim colData As Collection
    
    Select Case Control.ID
    Case enuMenus.参数设置
        If frmParameters.ShowMe(Me) Then
            '更新部件的参数值
            Call mclsTransmit.ReadParams
            Call ReadParams
        End If
    
    Case enuMenus.设备接口管理
        frmMachineManage.ShowMe Me
    
    Case enuMenus.基础数据传送
        If frmTransmitBD.ShowMe(Me, colData) Then
            If Not mclsTransmit Is Nothing Then
                For i = 1 To colData.Count
                    mclsTransmit.Transmit colData(i)
                Next
            End If
        End If
        Set colData = Nothing
        
    Case enuMenus.启用
        mclsTransmit.TimerAction = True
        mblnAction = mclsTransmit.TimerAction
        
    Case enuMenus.停用
        mclsTransmit.TimerAction = False
        mblnAction = mclsTransmit.TimerAction
        
    Case enuMenus.退出
        If mclsTransmit.Transmitting Then
            MsgBox "数据正在传送，请稍后退出！", vbInformation, GSTR_MSG
        Else
            mblnMenuClose = True
            Unload Me
        End If
    
    Case enuMenus.打印设置
        If gobjZLPrint Is Nothing Then Exit Sub
        Call gobjZLPrint.zlPrintSet
    
    Case enuMenus.打印预览, enuMenus.打印, enuMenus.输出Excel
        If TypeName(Me.ActiveControl) = "VSFlexGrid" Then
            If gobjZLPrint Is Nothing Then Exit Sub
            
            Set objPrint = CreateObject("zl9PrintMode.zlPrint1Grd")
            Set objPrint.Body = vsfLog
            
            On Error GoTo hErr
            If Control.ID = enuMenus.打印预览 Then
                gobjZLPrint.zlPrintOrView1Grd objPrint, 0
            ElseIf Control.ID = enuMenus.打印 Then
                gobjZLPrint.zlPrintOrView1Grd objPrint, 1
            Else
                gobjZLPrint.zlPrintOrView1Grd objPrint, 3
            End If
            On Error GoTo 0
        End If
    
    Case enuMenus.显示
        mdlTray.TrayStatus True, Me
        
    Case enuMenus.隐藏
        mdlTray.TrayStatus False, Me
        
    Case enuMenus.标准按钮
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
        
    Case enuMenus.文本标签
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
        
    Case enuMenus.大图标
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
        
    Case enuMenus.状态栏
        stbMain.Visible = Not Control.Checked
        cbsMain.RecalcLayout
    
    Case enuMenus.帮助主题
        Call gobjComLib.ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        
    Case enuMenus.中联主页
        Call gobjComLib.zlHomePage(Me.hwnd)
        
    Case enuMenus.中联论坛
        Call gobjComLib.zlWebForum(Me.hwnd)
        
    Case enuMenus.发送反馈
        Call gobjComLib.zlMailTo(Me.hwnd)
        
    End Select
    
    Exit Sub
    
hErr:
    Call gobjComLib.ErrCenter
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbMain.Visible Then
        Bottom = stbMain.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case enuMenus.参数设置
         Control.Enabled = mstrPrivs Like "*;参数设置;*"
    Case enuMenus.基础数据传送
        Control.Enabled = mblnAction
    Case enuMenus.启用
        Control.Enabled = mblnAction = False
    Case enuMenus.停用
        Control.Enabled = mblnAction
    Case enuMenus.显示
        Control.Enabled = mdlTray.mblnVisible = False
    Case enuMenus.隐藏
        Control.Enabled = mdlTray.mblnVisible
    Case enuMenus.标准按钮
        Control.Checked = Me.cbsMain(2).Visible
    Case enuMenus.文本标签
        Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
    Case enuMenus.大图标
        Control.Checked = cbsMain.Options.LargeIcons
    Case enuMenus.状态栏
        Control.Checked = Me.stbMain.Visible
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID Then
        Item.Handle = picLog.hwnd
    End If
End Sub

Private Sub Form_Activate()
    If mblnShow = True Then
        Screen.MousePointer = vbHourglass
        Me.Visible = False
        
        Me.Caption = App.ProductName
        
        Call InitTray
        Call InitDockPane
        Call InitCommandbars
        Call InitVSF
        Call ReadParams
        
        gobjComLib.RestoreWinState Me, App.EXEName
        
        If mclsTransmit Is Nothing Then
            MsgBox "请正确注册zlDrugMachineTimer.EXE（ActiveX EXE）！", vbInformation, GSTR_MSG
            mblnMenuClose = True
            Unload Me
            Exit Sub
        ElseIf Not mstrPrivs Like "*;基本;*" Then
            MsgBox "你无权使用，请检查用户权限或联系管理员！", vbInformation, GSTR_MSG
            mblnMenuClose = True
            Unload Me
            Exit Sub
        End If
        
        cbsMain.RecalcLayout
        mblnShow = False
        
        Me.Visible = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    Dim strFile As String
    
    '检查参数文件是否存在
    strFile = App.Path & "\" & GSTR_CONFIG_FILE
    Call VerifyConfigFile(strFile)
    
    '初始化对象
    Set mclsVSF = New clsVSFlexGridEx
    
    On Error Resume Next
    mstrSupportData = ""
    Set mclsTransmit = New zlDrugMachineTimer.clsDataTransmit
    If Err.Number <> 0 Then
        mblnShow = True
        MsgBox Err.Description, vbCritical, GSTR_MSG
        Exit Sub
    End If
    mstrSupportData = mclsTransmit.SupportData
    
    On Error GoTo hErr
    mclsTransmit.Init UserInfo.用户名, gobjComLib
    
'    Call CreateSOAP(mobjSOAP)
    Call CreateHTTP(mobjHTTP)
    mstrPrivs = ";" & gobjComLib.GetPrivFunc(GLNG_SYSTEM, GLNG_MODULE) & ";"
    
    mblnShow = True
    
    Exit Sub
    
hErr:
    MsgBox Err.Description, vbInformation, GSTR_MSG
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnMenuClose = False Then
        '隐藏
        mdlTray.TrayStatus False, Me
        Cancel = True
    Else
        If mblnShow = False Then
            If MsgBox("确认退出“" & App.ProductName & "”？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
            Else
                Cancel = False
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = Val("1-最小化") Then
        mdlTray.TrayStatus False, Me
    End If
    
    If Width < 6000 Then Width = 6000
    If Height < 4000 Then Height = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gobjComLib.SaveWinState Me, App.EXEName
    mdlTray.DeleteIcon picTray
    
    '窗体
    Set mclsVSF = Nothing
    Set mclsTransmit = Nothing
'    Set mobjSOAP = Nothing
    Set mobjHTTP = Nothing
    
    '全局
    Set gobjZLPrint = Nothing
    Set gobjXML = Nothing
    Set gobjFile = Nothing
    Set gcnOracle = Nothing
    Set gobjComLib = Nothing
    Set gobjRegister = Nothing
End Sub

Private Sub mclsTransmit_AfterTransmit(ByVal strLog As String)
    Dim l As Long

    '日志信息
    With vsfLog
        .Redraw = False
        .Rows = .Rows + 1
        
        l = .Rows - 1
        .TextMatrix(l, .ColIndex("时间")) = Now
        .TextMatrix(l, .ColIndex("日志信息")) = strLog
        .TopRow = l
        .Row = l
        
        If .Rows > mtypParams.显示最大行数 + 1 Then .RemoveItem 1
        
        .Redraw = True
    End With
End Sub

Private Sub picLog_Resize()
    On Error Resume Next
    With vsfLog
        .Top = 0
        .Left = 0
        .Width = picLog.ScaleWidth
        .Height = picLog.ScaleHeight
    End With
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'托盘消息接收与响应
    Dim lngMSG As Long
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    lngMSG = X / Screen.TwipsPerPixelX
    
    Select Case lngMSG
    Case WM_RBUTTONUP
        mobjPopupMenu.ShowPopup
    Case WM_LBUTTONDBLCLK
        Set objControl = mobjPopupMenu.Controls.Find(, enuMenus.显示, , True)
        If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
    End Select
End Sub

Private Sub ReadParams()
    Dim strFile As String

    strFile = App.Path & "\" & GSTR_CONFIG_FILE

    '读取参数文件的信息
    If gobjXML.OpenXMLFile(strFile) = False Then
        MsgBox "管理工具的参数文件不正确！", vbInformation, GSTR_MSG
        Exit Sub
    End If

    With mtypParams
        .显示最大行数 = Val(GetParameter(gobjXML, "viewlines"))
    End With
    
    gobjXML.CloseXMLDocument
End Sub
