VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiseaseReportSetting 
   Caption         =   "������������"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmDiseaseReportSetting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picParameter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1965
      ScaleHeight     =   675
      ScaleWidth      =   3930
      TabIndex        =   1
      Top             =   1485
      Width           =   3930
      Begin VB.OptionButton optParameter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʾ�༭���濨"
         Height          =   255
         Index           =   0
         Left            =   465
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton optParameter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�����༭���濨"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   315
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ҳ�����:"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   45
         Width           =   1110
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiseaseReportSetting.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12330
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2070
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   300
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDiseaseReportSetting.frx":0E1C
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDiseaseReportSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�������
'-----------------------------------------------------
Const conPane_Parameter = 1
Const conPane_Request = 2
Const conPane_Compend = 3

Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mlngFileID As Long
Private WithEvents mfrmRequest As frmEPRFileRequest     'Ӧ��Ҫ�󴰸�
Attribute mfrmRequest.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmEPRFileContent     '������ٴ���
Attribute mfrmContent.VB_VarHelpID = -1
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFileID As Long, lngCopyId As Long
    Dim cbrControl As CommandBarControl
    Dim str��� As String, str���� As String
    
    Select Case Control.ID
    Case conMenu_File_Exit
        Unload Me
    Case conMenu_Edit_ApplyTo
        If mlngFileID = 0 Then Exit Sub
        If frmEPRFileApplyTo.ShowMe(Me, mlngFileID) Then Call mfrmRequest.zlRefresh(mlngFileID)
    Case conMenu_Edit_Request
        If frmEPRFileDisease.ShowMe(Me, mlngFileID) Then Call mfrmRequest.zlRefresh(mlngFileID)
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
    End Select

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnZWave As Boolean
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Parameter
        Item.Handle = picParameter.hWnd
    Case conPane_Request
        If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
        Item.Handle = mfrmRequest.hWnd
    Case conPane_Compend
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "���ÿ���(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����Ҫ��(&R)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("T"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("R"), conMenu_Edit_Request
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "ʹ�ÿ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����Ҫ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panParameter As Pane, panRequest As Pane, panCompend As Pane, rsTemp As New ADODB.Recordset
    If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
    If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
    
    gstrSQL = "Select ID From �����ļ��б� Where ���� = 5 And ���� = 4"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    mlngFileID = NVL(rsTemp!ID, 0)
    Call mfrmRequest.zlRefresh(mlngFileID)
    Call mfrmContent.zlRefresh(mlngFileID)
    
    Set panParameter = dkpMan.CreatePane(conPane_Parameter, 400, 50, DockTopOf, Nothing)
    panParameter.Title = "��������"
    panParameter.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    panParameter.MaxTrackSize.Height = 50: panParameter.MinTrackSize.Height = 50
    
    Set panRequest = dkpMan.CreatePane(conPane_Request, 400, 90, DockBottomOf, Nothing)
    panRequest.Title = "Ӧ��Ҫ��"
    panRequest.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption

    Set panCompend = dkpMan.CreatePane(conPane_Compend, Me.ScaleX(Screen.Width, vbTwips, vbPixels) - 400, 100, DockRightOf, Nothing)
    panCompend.Title = "�ļ���ʽ"
    panCompend.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    
    If zlDatabase.GetPara("��ҳ�����༭���ر��濨", glngSys, 1277, "0") = 0 Then
        optParameter(0).Value = True
    Else
        optParameter(1).Value = True
    End If
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    mstrPrivs = gstrPrivs
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmRequest
    Unload mfrmContent
    Set mfrmRequest = Nothing
    Set mfrmContent = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub mfrmRequest_DblClick(lngWhere As zlEnumDClick)
Dim cbrControl As CommandBarControl

    Select Case lngWhere
    Case cprEmDClickApplyTo: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_ApplyTo)
    Case cprEmDClickRequest: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Request)
    Case Else: Set cbrControl = Nothing
    End Select
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub optParameter_Click(Index As Integer)
    Call zlDatabase.SetPara("��ҳ�����༭���ر��濨", CStr(Index), glngSys, 1277)
End Sub
