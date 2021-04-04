VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPass 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "合理用药监测"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPass.frx":1A62
   ScaleHeight     =   795
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraLight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3500
      TabIndex        =   0
      Top             =   240
      Width           =   270
      Begin VB.Image imgFlag 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   0
         Picture         =   "frmPass.frx":231F
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   240
      Top             =   840
   End
   Begin VB.Label lblFont 
      Caption         =   "Label1"
      Height          =   15
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblBtn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblDrug 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   205
      Left            =   795
      TabIndex        =   1
      ToolTipText     =   "药品名称"
      Top             =   240
      Width           =   2180
   End
   Begin VB.Image imgBack 
      Height          =   630
      Left            =   0
      Picture         =   "frmPass.frx":8B71
      Top             =   0
      Width           =   4200
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1080
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'-----系统托盘相关声明
'----------------------------------------------------------------------------------------------------
Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const SW_RESTORE = 9
Private Const SW_HIDE = 0

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type
Private mnfIconData As NOTIFYICONDATA

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'------------------------------------------------------------------------------------------
Private mudtRect As RECT
Private mudtPoint As POINTAPI
Private mblnMove As Boolean '判断指针是否位于移动栏
Private mblnMoveStart As Boolean '判断移动是否开始
Private mMoveX As Long, mMoveY As Long  '记录窗体移动前，窗体左上角与鼠标指针位置间的纵横距离
Private mfrmDrug As frmPassDrug
Attribute mfrmDrug.VB_VarHelpID = -1
Private mfrmDrugTip As New frmPassDrug
'
Public mstrDrugCode As String          '药品本位码
Private mstrDrugCurr As String          '已经加载显示药品

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    
    Case conMenu_Drug_View
        Call lblDrug_Click
    Case conMenu_PAR_SET        '参数设置
        frmPassPara.Show 1, Me
    Case conMenu_FRM_VISIBLE    '隐藏到托盘
        If Control.Caption = "隐藏界面" Then
            Me.Hide
        Else
           ShowWindow Me.hWnd, SW_RESTORE
        End If
    End Select
End Sub

Public Function GetTipForm() As Object
'功能:将提示信息界面与主界面绑定
    Set GetTipForm = mfrmDrugTip
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
     
    Case conMenu_PAR_SET
        
        If glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
            Control.Visible = PassCheckPrivs(glngModel)
        Else
            Control.Visible = False
        End If
        
    Case conMenu_FRM_VISIBLE
        If Me.Visible Then
            Control.Caption = "隐藏界面"
        Else
            Control.Caption = "显示界面"
        End If
    
    End Select
    If Control.Visible And Control.Enabled Then
        If gblnBreak Then Control.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim lngRet As Long
    Dim intIdx As Integer
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    Set gobjAir = CreateObject("zl9ComLib.clsAirBubble")
    
    Me.Top = 0       '窗体始终放在屏幕顶部
    Me.Left = Screen.Width \ 2 - Me.Width \ 2
    Me.Height = imgBack.Height
    Me.Width = imgBack.Width
    lblBtn.ForeColor = conCOLOR_TITLE_BAR
    lblFont.Visible = False
    lblDrug.ToolTipText = ""
    lngRet = GetWindowRect(Me.hWnd, mudtRect)
    Call InitCommandBar
    Call gobjFrm.SetLight("蓝")
    Me.BackColor = vbBlue
    SetFormTranslucency Me.hWnd, Me.BackColor, 0, LWA_COLORKEY
    Call SetNotifyIcon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMsg As Long
    Dim objPopup As CommandBarPopup
    
    If gblnBreak Then Exit Sub
    
    lngMsg = X / Screen.TwipsPerPixelX
    If Not mfrmDrug Is Nothing Then If mfrmDrug.IsOpen Then Exit Sub
    If frmPassResultZL.IsOpen Then Exit Sub
    If frmPassPara.IsOpen Then Exit Sub
    
    If lngMsg = WM_LBUTTONDBLCLK Then
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            If objPopup.CommandBar.Visible Then
                Exit Sub
            End If
        End If
        If Me.Visible Then
            Me.Hide
        Else
           ShowWindow Me.hWnd, SW_RESTORE
        End If
    ElseIf lngMsg = WM_RBUTTONUP Then '鼠标右键
        ShowPopup 1
    End If
End Sub

Private Sub Form_Paint()
    Dim blnDo As Boolean
    blnDo = True
    If Not mfrmDrug Is Nothing Then
        '药品说明书打开的情况下,不允许窗体SetWindowPos,否则药品说明书中查询下拉框会被影藏到药品说明书后面
        If mfrmDrug.IsOpen Then blnDo = False
    End If
    '使窗体始终置于最前面
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) = 0 And blnDo Then
         SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left \ Screen.TwipsPerPixelX, _
              Me.Top \ Screen.TwipsPerPixelY, Me.Width \ Screen.TwipsPerPixelX, _
              Me.Height \ Screen.TwipsPerPixelY, SWP_NOACTIVATE
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    imgBack.Move 0, 0
    lblBtn.Move imgBack.Width - 360, imgBack.Height / 2 - 100, 240, 240
    lblDrug.Move 800, imgBack.Height / 2 - 120, 2200, 240
    fraLight.BackColor = vbWhite
    fraLight.Move lblBtn.Left - (fraLight.Width + 30), imgBack.Height / 2 - fraLight.Height / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmDrug = Nothing
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    If Not gobjAir Is Nothing Then
        gobjAir.CloseAirBubble
        Set gobjAir = Nothing
    End If
End Sub

Private Sub fraLight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBack_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub fraLight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBack_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub fraLight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgback_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMove Then
         mMoveX = mudtPoint.X - mudtRect.Left
         mMoveY = mudtPoint.Y - mudtRect.Top
         mblnMoveStart = True
    End If
End Sub

Private Sub lblBtn_Click()
    ShowPopup 0
End Sub

Private Sub ShowPopup(ByVal bytFunc As Byte)
    Dim objPopup As CommandBarPopup
    
    If cbsMain Is Nothing Then Exit Sub
    If Not mfrmDrug Is Nothing Then
       If mfrmDrug.IsOpen Then Exit Sub
    End If
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
    If bytFunc = 0 Then
        Call GetWindowRect(Me.hWnd, mudtRect)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup , mudtRect.Right * Screen.TwipsPerPixelX - 1965, mudtRect.Bottom * Screen.TwipsPerPixelY
        End If
    Else
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup , mudtPoint.X * Screen.TwipsPerPixelX - 1965, mudtPoint.Y * Screen.TwipsPerPixelY - 120
        End If
    End If
End Sub

Private Sub imgFlag_Click()
    If grsRet Is Nothing Then Exit Sub
    If imgFlag.Picture <> frmIcons.imgPass.ListImages("蓝_4").Picture Then
        grsRet.Filter = "(Category=0 And Tag = 0)"
        frmPassResultZL.ShowMe gfrmMain, grsRet, 2
    End If
End Sub

Private Sub lblDrug_Click()
'功能:药品说明书
    Call GetDrugInstructions(Me, mfrmDrug, 0, mstrDrugCode, lblDrug.Caption)
End Sub

Private Sub lblDrug_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgBack_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblDrug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lblDrug.Font.Underline Then
        lblDrug.Font.Underline = True
        lblFont.AutoSize = True
        lblFont.Caption = lblDrug.Tag
        If lblFont.Width + 300 > lblDrug.Width Then
            lblDrug.ToolTipText = lblDrug.Tag
        Else
            lblDrug.ToolTipText = ""
        End If
    End If
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRet As Long
    Dim objPoint As RECT
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If mblnMoveStart Then
        lngLeft = (mudtPoint.X - mMoveX) * Screen.TwipsPerPixelX
        lngTop = (mudtPoint.Y - mMoveY) * Screen.TwipsPerPixelY
        If lngTop < 0 Then lngTop = 0      '窗体始终放在屏幕顶部
        If lngTop + Me.Height > Screen.Height Then lngTop = Screen.Height - Me.Height
        If lngLeft < 0 Then lngLeft = 0
        If lngLeft + Me.Width > Screen.Width Then lngLeft = Screen.Width - Me.Width
        
        If Not mfrmDrug Is Nothing Then
            If mfrmDrug.IsOpen Then
                Call mfrmDrug.ShowStyle(2, lngLeft, lngTop)
                Exit Sub
            End If
        End If
        If Not mfrmDrugTip Is Nothing Then
             If mfrmDrugTip.IsOpen Then
                Call mfrmDrugTip.ShowStyle(2, lngLeft, lngTop)
                Exit Sub
            End If
        End If
        Me.Left = lngLeft
        Me.Top = lngTop
    End If
    If lblDrug.Font.Underline Then
        lblDrug.Font.Underline = False
    End If
End Sub

Private Sub imgback_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRet As Long
    lngRet = GetWindowRect(Me.hWnd, mudtRect)
    mblnMoveStart = False
End Sub

Private Sub lblDrug_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgback_MouseUp Button, Shift, X, Y
End Sub

Private Sub tmrTime_Timer()
    Dim lngRet As Long
    Dim lngTime As Long
    Dim strURL As String
    Dim blnRet As Boolean
    
    lngRet = GetCursorPos(mudtPoint)
    '判断鼠标指针是否位于窗体拖动区
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If Not gobjPati Is Nothing Then
        If gobjPati.lng病人ID <> glngPatiID Then
            glngPatiID = gobjPati.lng病人ID
            Set grsRet = Nothing
        End If
    End If
    If gblnBreak Then
        lngTime = Timer
        If lngTime - gsngCheckLinkTime > gsngAutoLinkTime * 60 Then
            blnRet = HttpGet("http://" & gstrDrugIP & ":" & gstrDrugPort & "/DrugCorrect/Debug", responseText, 0.3, gblnBreak)
            If blnRet Or Not gblnBreak Then
                SetNotifyIcon
                ShowWindow Me.hWnd, SW_RESTORE
            Else
                gblnBreak = True
            End If
            gsngCheckLinkTime = lngTime
        End If
    End If
End Sub

Private Sub InitCommandBar()
'功能：初始化工具栏
    Dim objControl As CommandBarControl
    Dim objMenu As CommandBarPopup

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "药嘱审查", -1, False)
    objMenu.id = conMenu_EditPopup

    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Drug_View, "药品说明书")
        objControl.IconId = 801
        Set objControl = .Add(xtpControlButton, conMenu_PAR_SET, "个性化")
        objControl.IconId = 3202
        
        Set objControl = .Add(xtpControlButton, conMenu_FRM_VISIBLE, "隐藏")
        objControl.IconId = 181
        objControl.BeginGroup = True
    End With

End Sub

Public Sub SetDrug(ByVal strDrugName As String, ByVal strDrugCode As String)
'功能:将HIS本位码传入悬浮窗体
    If Not (GetWindowLong(Me.hWnd, GWL_EXSTYLE) & WS_EX_TOPMOST) And Me.Tag <> "置顶" Then
        '首次加载窗体未置顶强制置顶
        Call Form_Paint
        Me.Tag = "置顶"
    End If
     
    lblDrug.ToolTipText = ""
    If strDrugName = "" Then
        lblDrug.Caption = "药品名称"
    Else
        lblDrug.Tag = strDrugName
        Call GetSubString(lblDrug.Width, strDrugName)
        lblDrug.Caption = strDrugName
    End If
     
    mstrDrugCode = strDrugCode

    If Not mfrmDrug Is Nothing Then
        If mfrmDrug.IsOpen Then
            Call lblDrug_Click
        End If
    End If
End Sub

Public Sub SetLight(ByVal strLight As String)
'功能:设置灯的状态
    Dim arrLight(0 To 4) As String
    
    strLight = strLight & "_4"
    Set imgFlag.Picture = frmIcons.imgPass.ListImages(strLight).Picture
End Sub

Public Sub CloseGetDrugInstructions()
'功能:审查界面查看药品说明书前先检查药品说明书是否已经
    If Not mfrmDrug Is Nothing Then
        If mfrmDrug.IsOpen Then
            Unload mfrmDrug
        End If
    End If
End Sub

Public Sub SetNotifyIcon()
    Me.Hide  '隐藏窗体
    '下面的代码可以将图标添加到系统图标
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    mnfIconData.hWnd = Me.hWnd
    mnfIconData.uID = Me.Icon '这里确定使用哪个图标
    mnfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mnfIconData.uCallbackMessage = WM_MOUSEMOVE
    mnfIconData.hIcon = Me.Icon.Handle
    If gblnBreak Then
        mnfIconData.szTip = "中联合理用药" & vbCrLf & "连接状态:断开" & vbNullChar  '这里是将鼠标移到图标上时，将显示的文字
    Else
        mnfIconData.szTip = "中联合理用药" & vbCrLf & "连接状态:正常" & vbNullChar '这里是将鼠标移到图标上时，将显示的文字
    End If
    mnfIconData.cbSize = Len(mnfIconData)
    Call Shell_NotifyIcon(NIM_ADD, mnfIconData)
End Sub

Private Sub GetSubString(ByVal lngLen As Long, ByRef strSource As String)
'功能:按指定长度截取字符
    Dim lngMid As Long
    Dim lngMin As Long, lngMax As Long
    Me.FontSize = lblDrug.FontSize
    If TextWidth(strSource) < lngLen Then Exit Sub
    Do While strSource <> ""
        lngMin = 1: lngMax = Len(strSource)
        Do While lngMin <= lngMax
            lngMid = (lngMin + lngMax) \ 2
            If TextWidth(Mid(strSource, 1, lngMid)) > lngLen Then
                lngMax = lngMid - 1
            ElseIf TextWidth(Mid(strSource, 1, lngMid)) < lngLen Then
                lngMin = lngMid + 1
            Else
                Exit Do
            End If
        Loop
        strSource = Mid(strSource, 1, lngMid)
        Exit Do
    Loop
    
    Do While TextWidth(Mid(strSource, 1, lngMid)) >= lngLen
        lngMid = lngMid - 2
        strSource = Mid(strSource, 1, lngMid)
    Loop
End Sub
