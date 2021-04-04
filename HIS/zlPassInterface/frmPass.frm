VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPass 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "������ҩ���"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ҩƷ����"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "ҩƷ����"
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
'-----ϵͳ�����������
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
Private mblnMove As Boolean '�ж�ָ���Ƿ�λ���ƶ���
Private mblnMoveStart As Boolean '�ж��ƶ��Ƿ�ʼ
Private mMoveX As Long, mMoveY As Long  '��¼�����ƶ�ǰ���������Ͻ������ָ��λ�ü���ݺ����
Private mfrmDrug As frmPassDrug
Attribute mfrmDrug.VB_VarHelpID = -1
Private mfrmDrugTip As New frmPassDrug
'
Public mstrDrugCode As String          'ҩƷ��λ��
Private mstrDrugCurr As String          '�Ѿ�������ʾҩƷ

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    
    Case conMenu_Drug_View
        Call lblDrug_Click
    Case conMenu_PAR_SET        '��������
        frmPassPara.Show 1, Me
    Case conMenu_FRM_VISIBLE    '���ص�����
        If Control.Caption = "���ؽ���" Then
            Me.Hide
        Else
           ShowWindow Me.hWnd, SW_RESTORE
        End If
    End Select
End Sub

Public Function GetTipForm() As Object
'����:����ʾ��Ϣ�������������
    Set GetTipForm = mfrmDrugTip
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
     
    Case conMenu_PAR_SET
        
        If glngModel = PM_����༭ Or glngModel = PM_סԺ�༭ Or glngModel = PM_סԺҽ���嵥 Or glngModel = PM_����ҽ���嵥 Then
            Control.Visible = PassCheckPrivs(glngModel)
        Else
            Control.Visible = False
        End If
        
    Case conMenu_FRM_VISIBLE
        If Me.Visible Then
            Control.Caption = "���ؽ���"
        Else
            Control.Caption = "��ʾ����"
        End If
    
    End Select
    If Control.Visible And Control.Enabled Then
        If gblnBreak Then Control.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim lngRet As Long
    Dim intIdx As Integer
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    Set gobjAir = CreateObject("zl9ComLib.clsAirBubble")
    
    Me.Top = 0       '����ʼ�շ�����Ļ����
    Me.Left = Screen.Width \ 2 - Me.Width \ 2
    Me.Height = imgBack.Height
    Me.Width = imgBack.Width
    lblBtn.ForeColor = conCOLOR_TITLE_BAR
    lblFont.Visible = False
    lblDrug.ToolTipText = ""
    lngRet = GetWindowRect(Me.hWnd, mudtRect)
    Call InitCommandBar
    Call gobjFrm.SetLight("��")
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
    ElseIf lngMsg = WM_RBUTTONUP Then '����Ҽ�
        ShowPopup 1
    End If
End Sub

Private Sub Form_Paint()
    Dim blnDo As Boolean
    blnDo = True
    If Not mfrmDrug Is Nothing Then
        'ҩƷ˵����򿪵������,��������SetWindowPos,����ҩƷ˵�����в�ѯ������ᱻӰ�ص�ҩƷ˵�������
        If mfrmDrug.IsOpen Then blnDo = False
    End If
    'ʹ����ʼ��������ǰ��
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
    If imgFlag.Picture <> frmIcons.imgPass.ListImages("��_4").Picture Then
        grsRet.Filter = "(Category=0 And Tag = 0)"
        frmPassResultZL.ShowMe gfrmMain, grsRet, 2
    End If
End Sub

Private Sub lblDrug_Click()
'����:ҩƷ˵����
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
        If lngTop < 0 Then lngTop = 0      '����ʼ�շ�����Ļ����
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
    '�ж����ָ���Ƿ�λ�ڴ����϶���
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If Not gobjPati Is Nothing Then
        If gobjPati.lng����ID <> glngPatiID Then
            glngPatiID = gobjPati.lng����ID
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
'���ܣ���ʼ��������
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҩ�����", -1, False)
    objMenu.id = conMenu_EditPopup

    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Drug_View, "ҩƷ˵����")
        objControl.IconId = 801
        Set objControl = .Add(xtpControlButton, conMenu_PAR_SET, "���Ի�")
        objControl.IconId = 3202
        
        Set objControl = .Add(xtpControlButton, conMenu_FRM_VISIBLE, "����")
        objControl.IconId = 181
        objControl.BeginGroup = True
    End With

End Sub

Public Sub SetDrug(ByVal strDrugName As String, ByVal strDrugCode As String)
'����:��HIS��λ�봫����������
    If Not (GetWindowLong(Me.hWnd, GWL_EXSTYLE) & WS_EX_TOPMOST) And Me.Tag <> "�ö�" Then
        '�״μ��ش���δ�ö�ǿ���ö�
        Call Form_Paint
        Me.Tag = "�ö�"
    End If
     
    lblDrug.ToolTipText = ""
    If strDrugName = "" Then
        lblDrug.Caption = "ҩƷ����"
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
'����:���õƵ�״̬
    Dim arrLight(0 To 4) As String
    
    strLight = strLight & "_4"
    Set imgFlag.Picture = frmIcons.imgPass.ListImages(strLight).Picture
End Sub

Public Sub CloseGetDrugInstructions()
'����:������鿴ҩƷ˵����ǰ�ȼ��ҩƷ˵�����Ƿ��Ѿ�
    If Not mfrmDrug Is Nothing Then
        If mfrmDrug.IsOpen Then
            Unload mfrmDrug
        End If
    End If
End Sub

Public Sub SetNotifyIcon()
    Me.Hide  '���ش���
    '����Ĵ�����Խ�ͼ����ӵ�ϵͳͼ��
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    mnfIconData.hWnd = Me.hWnd
    mnfIconData.uID = Me.Icon '����ȷ��ʹ���ĸ�ͼ��
    mnfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mnfIconData.uCallbackMessage = WM_MOUSEMOVE
    mnfIconData.hIcon = Me.Icon.Handle
    If gblnBreak Then
        mnfIconData.szTip = "����������ҩ" & vbCrLf & "����״̬:�Ͽ�" & vbNullChar  '�����ǽ�����Ƶ�ͼ����ʱ������ʾ������
    Else
        mnfIconData.szTip = "����������ҩ" & vbCrLf & "����״̬:����" & vbNullChar '�����ǽ�����Ƶ�ͼ����ʱ������ʾ������
    End If
    mnfIconData.cbSize = Len(mnfIconData)
    Call Shell_NotifyIcon(NIM_ADD, mnfIconData)
End Sub

Private Sub GetSubString(ByVal lngLen As Long, ByRef strSource As String)
'����:��ָ�����Ƚ�ȡ�ַ�
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
