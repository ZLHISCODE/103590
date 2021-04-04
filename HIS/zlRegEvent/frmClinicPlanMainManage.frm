VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmClinicPlanMainManage 
   Caption         =   "���ﰲ�Ź���"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "frmClinicPlanMainManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   10260
      Visible         =   0   'False
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10590
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicPlanMainManage.frx":1082
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18124
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   89
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   953
            MinWidth        =   882
            Text            =   "ְ��"
            TextSave        =   "ְ��"
            Key             =   "DoctorsTitle"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "PlanColor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   720
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmClinicPlanMainManage.frx":1916
      Left            =   1260
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmClinicPlanMainManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean

Private mWorkPan As Pane '��ǰ����
Private mfrmCurForm As Form '��ǰ���ܴ���
Public mFunListActived As Boolean

Private mrsְ�� As ADODB.Recordset  '����ҽ��רҵ����ְ�ƺͶ�Ӧ�ı�ʶ��������

Private mfrmClinicPlanMainFun As frmClinicPlanMainFun

Private mfrmClinicWorkTimeManage As frmClinicWorkTimeManage
Private mfrmClinicHolidayManage As frmClinicHolidayManage
Private mfrmClinicOfficeManage As frmClinicOfficeManage
Private mfrmClinicSignalSourceManage As frmClinicSignalSourceManage
    
Private mfrmClinicFixedPlanManage As frmClinicFixedPlanManage
Private mfrmClinicPlanDaysManage As frmClinicPlanDaysManage
Private mfrmClinicPlanTempletManage As frmClinicPlanTempletManage
Private mfrmClinicPlanStopVisitManage As frmClinicPlanStopVisitManage
Private mfrmClinicPlanTempletByDayManage As frmClinicPlanTempletByDayManage

Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    Select Case CommandBar.Parent.id
    Case conMenu_View_FindType
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.InitCommandsPopup(CommandBar)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    mblnUnload = False
    If mblnFirst Then mblnFirst = False: Exit Sub
    
    Err = 0: On Error Resume Next
    '���mFunListActived������ActiveFormChange�¼���Ϊ�˿��ƽ���
    If Not mfrmCurForm Is Nothing Then
        If mFunListActived = False And mfrmCurForm.Visible Then mfrmCurForm.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    '�̶������Զ����ɳ����¼
    On Error Resume Next
    mblnUnload = False
    zlCommFun.ShowFlash "���ڼ������ݣ����Ե�...", Me
    zlDatabase.ExecuteProcedure "zl1_auto_buildingregisterplan(Null)", Me.Caption
    
    Err = 0: On Error GoTo errHandler
    '�����¼վ���޹̶�������¼�����Զ������ٴ�������¼
    'Zl_�ٴ������_Add(
    strSQL = "Zl_�ٴ������_Add("
    '  ��������_In         Number,
    strSQL = strSQL & "" & "2" & ","
    '  ����id_In           �ٴ������.Id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �������_In         �ٴ������.�������%Type,
    strSQL = strSQL & "'" & "�̶������" & "',"
    '  վ��_In             ���ű�.վ��%Type,
    strSQL = strSQL & "'" & gstrNodeNo & "',"
    '  ȫԺ��Դ����վ��_In ���ű�.վ��%Type,
    strSQL = strSQL & "'" & gVisitPlan_ModulePara.str��Դά��վ�� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    mblnFirst = True
    Set gobjRegist = New clsRegist
    gobjRegist.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    mstrPrivs = gstrPrivs
    mlngModule = glngModul
    
    Set mfrmClinicPlanMainFun = New frmClinicPlanMainFun
    Call mfrmClinicPlanMainFun.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)

    Call DefMainCommandBars
    Call InitPanel '��ʼ��dkpMain
    Call RestoreWinState(Me, App.ProductName)
    Call Loadְ��  '����ҽ��ְ�Ƽ���ʶ��
    
    zlCommFun.StopFlash
    Exit Sub
errHandler:
    zlCommFun.StopFlash
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnUnload = True
End Sub

Private Sub InitPanel()
    Dim objPane As Pane

    Err = 0: On Error GoTo errHandler
    Set objPane = dkpMain.CreatePane(Pane_FunFace, 150, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable
    objPane.MinTrackSize.Width = 130
    objPane.MaxTrackSize.Width = 240
    objPane.Tag = Pane_FunFace

    Set mWorkPan = dkpMain.CreatePane(Pane_Face, 700, 400, DockRightOf, objPane)
    mWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mWorkPan.Tag = Pane_Face

    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    Err = 0: On Error GoTo errHandler

    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False

    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Sign, "ְ�Ʊ�ʶ����(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportPlan, "���롰�ҺŰ��š�(&I)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrSubControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    '��ʾ�Զ��屨��˵�
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs, _
        "ZL" & glngSys \ 100 & "_INSIDE_1114_1", "ZL" & glngSys \ 100 & "_INSIDE_1114_2", _
        "ZL" & glngSys \ 100 & "_INSIDE_1114_3", "ZL" & glngSys \ 100 & "_INSIDE_1114_4")

    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")

        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With

    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DefSubCommandBars(ByVal ObjItem As Pane)
    '���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String

    Err = 0: On Error GoTo errHandler
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsThis.Count >= 2 Then
        idx = GetFirstCommandBar(cbsThis(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsThis(2).Visible
            bytStyle = cbsThis(2).Controls(idx).Style
        End If
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.Hwnd)
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next

    '���������¼���
    Call DefMainCommandBars

    '�Ӵ������¼���
    If Not mfrmCurForm Is Nothing Then
        Call mfrmCurForm.zlDefCommandBars
    End If

    '�ָ����̶���һЩ�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop + xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsThis.Count
        cbsThis(lngCount).ContextMenuPresent = False
        cbsThis(lngCount).ShowTextBelowIcons = False
        cbsThis(lngCount).EnableDocking xtpFlagStretched + xtpFlagHideWrap
        For Each objControl In cbsThis(lngCount).Controls
            If objControl.Type <> xtpControlLabel _
                And objControl.Type <> xtpControlEdit Then
                objControl.Style = bytStyle
            End If
        Next
        cbsThis(lngCount).Visible = blnShowBar
    Next

    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub SetProgressBarPostion(ByVal lngValue As Long, _
    Optional ByVal blnInit As Boolean, Optional ByVal lngMax As Long)
    
    With ProgressBar
        If blnInit Then
            .Max = lngMax
            .Left = stbThis.Panels(2).Left + 50
            .Top = stbThis.Top + (stbThis.Height - .Height) / 2 + 20
            .Width = stbThis.Panels(2).Width - 100
            .ZOrder
        Else
            .Value = lngValue
        End If
    End With
End Sub

Public Sub ActiveFormChange(objForm As Form)
    Err = 0: On Error Resume Next
    mFunListActived = objForm Is mfrmClinicPlanMainFun
End Sub

Public Sub NodeChanged(ByVal strKey As String)
    Call mfrmClinicPlanMainFun.RefreshVisitTable(strKey)
End Sub

Public Sub StatusShowInfoChanged(ByVal PanelIndex As Integer, ByVal strInfo As String)
    If PanelIndex = 2 Then
        stbThis.Panels(2).Text = strInfo
    Else
        stbThis.Panels(3).Text = strInfo
    End If
End Sub

Public Sub SelectedChange(ByVal bytMode As RegistPlanFun, _
    Optional ByVal lng����ID As Long, _
    Optional ByVal intYear As Integer, Optional ByVal intMonth As Integer, _
    Optional ByVal strTitle As String, Optional ByVal bytģ������ As Byte)
    
    Err = 0: On Error Resume Next
    stbThis.Panels(2).Text = ""
    stbThis.Panels(3).Visible = False
    stbThis.Panels("PlanColor").Visible = False
    stbThis.Panels("DoctorsTitle").Visible = False
    
    Select Case bytMode
    Case Pane_WorkTime
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicWorkTimeManage Is Nothing Then
                Set mfrmClinicWorkTimeManage = New frmClinicWorkTimeManage
                Call mfrmClinicWorkTimeManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicWorkTimeManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            'ˢ���Ӵ���˵���������
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
    Case Pane_Holiday
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicHolidayManage Is Nothing Then
                Set mfrmClinicHolidayManage = New frmClinicHolidayManage
                Call mfrmClinicHolidayManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicHolidayManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.RefrashData(Year(zlDatabase.Currentdate))
    Case Pane_DoctorOffice
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicOfficeManage Is Nothing Then
                Set mfrmClinicOfficeManage = New frmClinicOfficeManage
                Call mfrmClinicOfficeManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicOfficeManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
    Case Pane_SignalSource
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicSignalSourceManage Is Nothing Then
                Set mfrmClinicSignalSourceManage = New frmClinicSignalSourceManage
                Call mfrmClinicSignalSourceManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicSignalSourceManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
        If mrsְ��.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_StopPlan 'ͣ�����
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanStopVisitManage Is Nothing Then
                Set mfrmClinicPlanStopVisitManage = New frmClinicPlanStopVisitManage
                Call zlControl.FormSetCaption(mfrmClinicPlanStopVisitManage, False, False)
                Call mfrmClinicPlanStopVisitManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanStopVisitManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData
    Case Pane_PlanTemplet
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanTempletManage Is Nothing Then
                Set mfrmClinicPlanTempletManage = New frmClinicPlanTempletManage
                Call mfrmClinicPlanTempletManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanTempletManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData IIf(bytģ������ = 0, 2, 1), lng����ID, True
        stbThis.Panels(3).Visible = True
        If mrsְ��.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_MonthTemplet
        '2-�����Ű�����Ű�ģ��
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanTempletByDayManage Is Nothing Then
                Set mfrmClinicPlanTempletByDayManage = New frmClinicPlanTempletByDayManage
                Call mfrmClinicPlanTempletByDayManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanTempletByDayManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData 1, lng����ID, True, intYear, intMonth, strTitle
        stbThis.Panels(3).Visible = True
        If mrsְ��.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_FixedPlan
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicFixedPlanManage Is Nothing Then
                Set mfrmClinicFixedPlanManage = New frmClinicFixedPlanManage
                Call mfrmClinicFixedPlanManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicFixedPlanManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData lng����ID, True
        stbThis.Panels("PlanColor").Visible = True
        stbThis.Panels(3).Visible = True
        If mrsְ��.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_MonthPlan, Pane_WeekPlan
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanDaysManage Is Nothing Then
                Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
                Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanDaysManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData IIf(bytMode = Pane_MonthPlan, 1, 2), lng����ID, True, intYear, intMonth, strTitle
        stbThis.Panels("PlanColor").Visible = True
        stbThis.Panels(3).Visible = True
        If mrsְ��.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
    Set mWorkPan = Nothing
    Set mfrmCurForm = Nothing
    'ж�����д���
    If Not mfrmClinicWorkTimeManage Is Nothing Then Unload mfrmClinicWorkTimeManage: Set mfrmClinicWorkTimeManage = Nothing
    If Not mfrmClinicHolidayManage Is Nothing Then Unload mfrmClinicHolidayManage: Set mfrmClinicHolidayManage = Nothing
    If Not mfrmClinicOfficeManage Is Nothing Then Unload mfrmClinicOfficeManage: Set mfrmClinicOfficeManage = Nothing
    If Not mfrmClinicSignalSourceManage Is Nothing Then Unload mfrmClinicSignalSourceManage: Set mfrmClinicSignalSourceManage = Nothing

    If Not mfrmClinicPlanDaysManage Is Nothing Then Unload mfrmClinicPlanDaysManage: Set mfrmClinicPlanDaysManage = Nothing
    If Not mfrmClinicFixedPlanManage Is Nothing Then Unload mfrmClinicFixedPlanManage: Set mfrmClinicFixedPlanManage = Nothing
    If Not mfrmClinicPlanTempletManage Is Nothing Then Unload mfrmClinicPlanTempletManage: Set mfrmClinicPlanTempletManage = Nothing
    If Not mfrmClinicPlanStopVisitManage Is Nothing Then Unload mfrmClinicPlanStopVisitManage: Set mfrmClinicPlanStopVisitManage = Nothing
    If Not mfrmClinicPlanTempletByDayManage Is Nothing Then Unload mfrmClinicPlanTempletByDayManage: Set mfrmClinicPlanTempletByDayManage = Nothing
    Unload mfrmClinicPlanMainFun: Set mfrmClinicPlanMainFun = Nothing
    
    On Error Resume Next
    Unload frmClinicPlanTemp '�ر���ʱ����
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_ImportPlan
        Control.Visible = HavePrivs(mstrPrivs, "���ﰲ��;���п���", True)
        Control.Enabled = Control.Visible
    Case conMenu_File_Sign
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ְ�Ʊ�ʶ����")
        Control.Enabled = Control.Visible
    Case Else
        If Not mfrmCurForm Is Nothing Then
            Call mfrmCurForm.zlUpdateCommandBars(Control)
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl

    Err = 0: On Error GoTo errHandler
    Select Case Control.id
    Case conMenu_File_Sign
        If frmClinicDoctorTitleSet.ShowMe(Me) = False Then Exit Sub
        Call Loadְ��
        'ˢ�³��������
        If mWorkPan.Tag = Pane_FixedPlan Or mWorkPan.Tag = Pane_MonthPlan _
            Or mWorkPan.Tag = Pane_WeekPlan _
            Or mWorkPan.Tag = Pane_PlanTemplet _
            Or mWorkPan.Tag = Pane_MonthTemplet Then
            Call mfrmClinicPlanMainFun.RefreshVisitTable
        ElseIf mWorkPan.Tag = Pane_SignalSource Then
            Call SelectedChange(Pane_SignalSource)
        End If
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_View_StatusBar
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        cbsThis(2).Visible = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each objControl In cbsThis(2).Controls
            objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Control.Checked = Not Control.Checked
        cbsThis.Options.LargeIcons = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))

        ElseIf Control.id = conMenu_File_Parameter Then
            '��������
            Dim frmPara As New frmClinicPlanParaSet
            If frmPara.ShowMe(Me, mlngModule, mstrPrivs) Then
                Call InitLocVisitPlanPar(mlngModule)
            End If
        ElseIf Control.id = conMenu_File_ImportPlan Then
            '���밲��
            If ImportPlan() Then
                MsgBox "���ҺŰ��š�������ɣ�", vbInformation, gstrSysName
                mfrmClinicPlanMainFun.RefreshVisitTable  '���¶�ȡ
            End If
        Else
            If Control.id = conMenu_View_Refresh And mFunListActived And Val(mWorkPan.Tag) > 5 Then
                'ˢ�³�����б�
                mfrmClinicPlanMainFun.RefreshVisitTable
                Exit Sub
            End If

            If Val(mWorkPan.Tag) = Pane_PlanTemplet _
                Or Val(mWorkPan.Tag) = Pane_FixedPlan _
                Or Val(mWorkPan.Tag) = Pane_MonthPlan _
                Or Val(mWorkPan.Tag) = Pane_WeekPlan _
                Or Val(mWorkPan.Tag) = Pane_StopPlan _
                Or Val(mWorkPan.Tag) = Pane_MonthTemplet Then
                
                If ExecuteAddNewPlan(Control) Then Exit Sub
            End If
            If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.zlExecuteCommandBars(Control)
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExecuteAddNewPlan(ByVal Control As CommandBarControl) As Boolean
    '��������
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strKey As String

    Err = 0: On Error GoTo errHandler
    Select Case Control.id
    Case conMenu_Edit_AddTemplet 'ģ��
        If mfrmClinicPlanTempletManage Is Nothing Then
            Set mfrmClinicPlanTempletManage = New frmClinicPlanTempletManage
            Call mfrmClinicPlanTempletManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanTempletManage.AddNewPlanTemplet
    Case conMenu_Edit_AddMonthPlan '�°���
        If mfrmClinicPlanDaysManage Is Nothing Then
            Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
            Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanDaysManage.AddNewPlan(True)
    Case conMenu_Edit_AddWeekPlan '�ܰ���
        If mfrmClinicPlanDaysManage Is Nothing Then
            Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
            Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanDaysManage.AddNewPlan(False)
    Case Else
        Exit Function
    End Select
    If strKey = "" Then Exit Function
    Call mfrmClinicPlanMainFun.RefreshVisitTable(strKey)
    ExecuteAddNewPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo errHandler
    If Item.Tag = Pane_FunFace Then
        Item.Handle = mfrmClinicPlanMainFun.Hwnd
    ElseIf Not mfrmCurForm Is Nothing Then
        Item.Handle = mfrmCurForm.Hwnd
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ImportPlan() As Boolean
    '������ʷ����
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim strTemp As String, cllSQL As Collection, blnDo As Boolean
    Dim rsPlanAll As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '���:
    '1.������ڰ��ţ��������ٵ���
    '2.�����Դ�д������ݣ������ڰ��ţ������ѡ�����ʱ��Щ��Դ���ᱻ���ǣ��Ƿ�������룿��
    '3.ԭ�ҺŰ����е��ϰ�ʱ�������в����ڵģ��������룬Ҫ���������ӣ��簲����ʹ����"����"�������ϰ�ʱ�������û��"����"��
    strSQL = "Select ���� From �ҺŰ��� Order By ID Desc"
    Set rsPlanAll = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsPlanAll.EOF Then
        MsgBox "�����ڹҺŰ��ţ�����Ҫ���룡", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select 1 From �ٴ������ A, �ٴ����ﰲ�� B Where a.Id = b.����id And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������Դ B, ���ű� C" & vbNewLine & _
                " Where a.��Դid = b.Id And b.����id = c.Id And (c.վ�� Is Null Or c.վ�� = [1]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
        If rsTemp.EOF Then
            '���Ǳ�վ���
            MsgBox "��ǰ����Ժ���Ѿ������ٴ����ﰲ���ˣ�����ɾ�������������룡", vbInformation, gstrSysName
        Else
            MsgBox "��ǰ�Ѿ������ٴ����ﰲ���ˣ�����ɾ�������������룡", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    strSQL = "Select 1 From �ٴ������Դ Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        If MsgBox("��ǰ�����ٴ������Դ���ڵ���ʱ��Щ��Դ���ᱻ���ǣ��Ƿ�������룿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    strSQL = "Select f_List2str(Cast(Collect(s.ʱ���) As t_Strlist)) As ʱ���" & vbNewLine & _
            " From (Select ʱ���, Row_Number() Over(Partition By ʱ��� Order By ʱ���) As ���" & vbNewLine & _
            "        From (Select Decode(b.�к�, 1, a.��һ, 2, a.�ܶ�, 3, a.����, 4, a.����, 5, a.����, 6, a.����, a.����) As ʱ���" & vbNewLine & _
            "               From (Select ��һ, �ܶ�, ����, ����, ����, ����, ����" & vbNewLine & _
            "                      From �ҺŰ���" & vbNewLine & _
            "                      Union All" & vbNewLine & _
            "                      Select ��һ, �ܶ�, ����, ����, ����, ����, ���� From �ҺŰ��żƻ�) A," & vbNewLine & _
            "                    (Select Level As �к� From Dual Connect By Level <= 7) B)" & vbNewLine & _
            "        Where ʱ��� Is Not Null) S, ʱ��� T" & vbNewLine & _
            " Where s.ʱ��� = t.ʱ���(+) And t.ʱ��� Is Null And s.��� = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!ʱ���) <> "" Then
            MsgBox "ԭ�ҺŰ����е��ϰ�ʱ��Ρ�" & Nvl(rsTemp!ʱ���) & "�������ڣ������ڡ���������>�ϰ�ʱ���������ӣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '���밲��
    zlCommFun.ShowFlash "���ڵ��밲�ţ����Ե�...", Me
    Set cllSQL = New Collection
    Do While Not rsPlanAll.EOF
        cllSQL.Add "Zl_�ٴ������_����('" & rsPlanAll!���� & "'," & IIf(cllSQL.Count = 0, 1, 0) & ")"
        rsPlanAll.MoveNext
    Loop
    
    'ִ��SQL���
    Call SetProgressBarPostion(0, True, cllSQL.Count)
    Me.ProgressBar.Visible = True
    blnDo = True: gcnOracle.BeginTrans
    For i = 1 To cllSQL.Count
        zlDatabase.ExecuteProcedure cllSQL(i), Me.Caption
        Call SetProgressBarPostion(i)
    Next
    gcnOracle.CommitTrans: blnDo = False
    Me.ProgressBar.Visible = False
    
    zlCommFun.StopFlash
    ImportPlan = True
    Exit Function
errHandler:
    If blnDo Then gcnOracle.RollbackTrans
    Me.ProgressBar.Visible = False
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '����:������ص��Զ��屨��

    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If frmClinicPlanTemp.Visible Then Exit Sub
    Select Case Panel.Key
    Case "PlanColor"
        Call frmClinicPlanTemp.ShowPlanColor(Me)
    Case "DoctorsTitle"
        Call frmClinicPlanTemp.ShowDoctorsTitle(Me, mrsְ��)
    End Select
End Sub

Public Function GetPopupCommandBarSub() As CommandBar
    '��ȡ�����˵�
    If mfrmCurForm Is Nothing Then Exit Function
    Set GetPopupCommandBarSub = GetPopupCommandBar(mfrmCurForm, cbsThis)
End Function

'����:��ȡҽ��רҵ����ְ������Ƽ���ʶ��
Public Sub Loadְ��()
    Dim strSQL As String
    strSQL = "Select ����, ��ʶ�� From רҵ����ְ��" & vbNewLine & _
             "Where ���� like '23%'and ����<>'23'" & vbNewLine & _
             "And ��ʶ�� Is Not Null"
    Set mrsְ�� = zlDatabase.OpenSQLRecord(strSQL, "��ȡְ�ƺͱ�ʶ��")
    
    If mrsְ��.RecordCount = 0 Then
        stbThis.Panels("DoctorsTitle").Visible = False
    Else
        If mWorkPan.Tag = Pane_FixedPlan Or mWorkPan.Tag = Pane_MonthPlan _
            Or mWorkPan.Tag = Pane_WeekPlan _
            Or mWorkPan.Tag = Pane_MonthTemplet _
            Or mWorkPan.Tag = Pane_PlanTemplet _
            Or mWorkPan.Tag = Pane_SignalSource Then
            stbThis.Panels("DoctorsTitle").Visible = True
        Else
            stbThis.Panels("DoctorsTitle").Visible = False
        End If
    End If
End Sub


