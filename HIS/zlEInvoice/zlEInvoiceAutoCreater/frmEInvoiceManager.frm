VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEInvoiceManager 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ʊ���Զ����߹���"
   ClientHeight    =   4860
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8025
   Icon            =   "frmEInvoiceManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraCons 
      Height          =   1845
      Left            =   90
      TabIndex        =   1
      Top             =   1080
      Width           =   7605
      Begin VB.TextBox txtSplit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1290
         TabIndex        =   3
         Text            =   "60"
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lbl��ѯ��� 
         AutoSize        =   -1  'True
         Caption         =   "ִ�м����           �����ӣ�"
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2610
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4500
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEInvoiceManager.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "14:53:01"
            TextSave        =   "17:41"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEInvoiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngTimerID As Long, mblnStart As Boolean

Public Function ShowMe(ByVal strPrivs As String) As Boolean
    '�������
    mstrPrivs = strPrivs
    Me.Show
End Function

Private Sub Form_Load()
    Call InitCommandBar
    Call RestoreWinState(Me, App.ProductName)
    
    '��������ʾͼ��
    With nfIconData
        .hwnd = Me.hwnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle
        '��������ƶ���������ʱ��ʾ��Tip
        .szTip = Me.Caption + "(�汾 " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '�˵�����
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_ViewLog, "������־(&L)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Disuse, "ͣ��(&P)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Disuse, "ͣ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
    With cbsMain.Options
        '.AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        '.AddHiddenCommand conMenu_File_Excel '�����Excel
    End With
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    '����Left,Top,Right,Bottom�����صĿͻ�������λ�öԴ����е������ؼ������ŷ�
    fraCons.Move lngLeft + 10, lngTop, lngRight - lngLeft - 10, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean

    Select Case Control.ID
    Case conMenu_Edit_Reuse '����
        Control.Enabled = Not mblnStart
    Case conMenu_Edit_Disuse 'ͣ��
        Control.Enabled = mblnStart And Not gblnExecuting
    
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub StartTimer()
    '������ʱ��
    If Len(txtSplit.Text) > 4 Then
        MsgBox "ִ�м��ʱ����Ч�����������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtSplit.Text) <= 0 Then
        MsgBox "ִ�м��ʱ����Ч�����������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    zlWritLog glngModul, "�����Զ����ߵ���Ʊ������", "StartTimer", "ִ�м����" & glngSplitTime
    
    gblnExecuting = False
    Set gfrmMain = Me
    glngSplitTime = Val(txtSplit.Text) * 60
     
    mlngTimerID = SetTimer(0, 0, glngSplitTime * 1000, AddressOf TimerProc)
    If mlngTimerID = 0 Then
        MsgBox "��ʱ������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnStart = True
    txtSplit.Locked = True
End Sub

Private Sub StopTimer()
    'ֹͣ��ʱ��
    If gblnExecuting Then
        MsgBox "��ǰ����ִ�е���Ʊ�ݵĿ��ߣ����Ժ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If mlngTimerID = 0 Then Exit Sub
     
    If KillTimer(0, mlngTimerID) = 0 Then
        MsgBox "��ʱ��ͣ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnStart = False
    txtSplit.Locked = False
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
    Case conMenu_File_ViewLog '��־�鿴
    
    Case conMenu_Edit_Reuse '����
        Call StartTimer
        
    Case conMenu_Edit_Disuse 'ͣ��
        Call StopTimer
    
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_LBUTTONUP
        '�������/�Ҽ�����ʾ����
        ShowWindow Me.hwnd, SW_RESTORE
        '���������Ŀ���ǰѴ�����ʾ�ڴ������
        Me.Show
        Me.SetFocus
    Case WM_RBUTTONUP
        'PopupMenu MenuTray '�������ϵͳTrayͼ���ϵ��Ҽ����򵯳��˵�MenuTray
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONDBLCLK
    Case Else
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gblnExecuting Then
        MsgBox "��ǰ����ִ�е���Ʊ���Զ������������Ժ�", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    
    If mblnStart Then
        If MsgBox("����Ʊ���Զ��������������ã��˳����Զ�ֹͣ����ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
        Call StopTimer
        Exit Sub
    End If
    
    If MsgBox("��ǰδ���õ���Ʊ���Զ�����������ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
    
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide '��С��ʱ���ش���
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mlngTimerID = 0
    mblnStart = False
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtSplit_GotFocus()
    zlControl.TxtSelAll txtSplit
End Sub

Private Sub txtSplit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
