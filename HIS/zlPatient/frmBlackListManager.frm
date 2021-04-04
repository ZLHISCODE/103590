VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJOCK.DOCKINGPANE.UNICODE.9600.OCX"
Begin VB.Form frmBlackListManager 
   Caption         =   "���˲�����¼����"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   Icon            =   "frmBlackListManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   15240
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBlackListManager.frx":06EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21802
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBlackListManager.frx":0F7E
      Left            =   1260
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBlackListManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean
Private mobjWorkPan As Pane '��ǰ����
Private mfrmCurForm As Form '��ǰ���ܴ���

Public mblnFunListActived As Boolean

Private WithEvents mfrmBlackListMainFun As frmBlackListMainFun    '��Ҫ���ܲ˵�
Attribute mfrmBlackListMainFun.VB_VarHelpID = -1
Private WithEvents mfrmBlackTypeManage As frmBlackTypeManage   '������Ϊ�������
Private WithEvents mfrmBlackListReasonManage As frmBlackListReasonManage   '������Ϊ���õ�ԭ�����
Private WithEvents mfrmBlackListRecordManage As frmBlackListRecordManage   '������Ϊ��¼����
Attribute mfrmBlackListRecordManage.VB_VarHelpID = -1


Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If Me.Visible = False Then Exit Sub
    Err = 0: On Error Resume Next
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.InitCommandsPopup(CommandBar)
    End Select
End Sub

Private Sub Form_Activate()

    If mblnUnload Then Unload Me: Exit Sub
    mblnUnload = False
    If mblnFirst Then mblnFirst = False: Exit Sub
    
    Err = 0: On Error Resume Next
    '���mblnFunListActived������ActiveFormChange�¼���Ϊ�˿��ƽ���
    If Not mfrmCurForm Is Nothing Then
        If mblnFunListActived = False And mfrmCurForm.Visible Then mfrmCurForm.SetFocus
    End If
End Sub
Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ģ�鼶����
    '����:���˺�
    '����:2018-11-08 10:36:07
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    Set mfrmBlackListMainFun = New frmBlackListMainFun
    Call mfrmBlackListMainFun.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
     
    Err = 0: On Error GoTo errHandle
    mblnFirst = True: mstrPrivs = gstrPrivs: mlngModule = glngModul
    Call InitVar
    Call DefMainCommandBars
    Call InitPanel '��ʼ��dkpMain
    Call RestoreWinState(Me, App.ProductName)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     mblnUnload = True
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����������������
    '����:���˺�
    '����:2018-11-08 10:38:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    
    Set objPane = dkpMain.CreatePane(gEM_BlackListFun.Em_Pane_FunFace, 150, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable
    objPane.MinTrackSize.Width = 130
    objPane.MaxTrackSize.Width = 240
    objPane.Tag = Em_Pane_FunFace

    Set mobjWorkPan = dkpMain.CreatePane(gEM_BlackListFun.Em_Pane_Face, 700, 400, DockRightOf, objPane)
    mobjWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mobjWorkPan.Tag = Em_Pane_Face

    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 10:41:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    Err = 0: On Error GoTo ErrHandler

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
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
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
    cbrMenuBar.ID = conMenu_HelpPopup
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
    Set cbrToolBar = GetCommbarFromName(cbsThis, "������")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    End If
    
   ' Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub DefSubCommandBars(ByVal objItem As Pane)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ӵ���˵���������
    '���:ObjItem-��ǰ����ҳ����
    '����:���˺�
    '����:2018-11-08 10:42:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl, bytStyle As XTPButtonStyle, blnShowBar As Boolean
    Dim lngCount As Long, lngIndex As Long, objCustom As CommandBarControlCustom
    Dim strName As String, cbrToolBar As CommandBar

    Err = 0: On Error GoTo ErrHandler
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsThis.Count >= 2 Then
        lngIndex = zlGetFirstCommandBar(cbsThis(2).Controls)
        If lngIndex > 0 Then
            blnShowBar = cbsThis(2).Visible
            bytStyle = cbsThis(2).Controls(lngIndex).Style
        End If
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hWnd)
    
    'ɾ�����ڵĹ������������˵���
    cbsThis.ActiveMenuBar.Controls.DeleteAll
    If Not mfrmCurForm Is Nothing Then mfrmCurForm.zlCancelBands
    
    Set cbrToolBar = GetCommbarFromName(cbsThis, "������")
    cbrToolBar.Controls.DeleteAll

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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub mfrmBlackListMainFun_zlActivate(ByVal frmSubForm As Form)
    '�Ӵ��弯��ʱ�������¼�
    mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub
Private Sub mfrmBlackListReasonManage_zlActivate(ByVal frmSubForm As Form)
  mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub

Private Sub mfrmBlackTypeManage_zlActivate(ByVal frmSubForm As Form)
    mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub

Private Sub mfrmBlackListMainFun_SelectedChange(ByVal bytFunMode As gEM_BlackListFun, ByVal strBlackLitType As String)
    '����ѡ��ı�󴥷����¼�
    
    On Error GoTo errHandle
    stbThis.Panels(2).Text = ""
 
    
    Select Case bytFunMode
    Case Em_Pane_Type   '�����������
        If mobjWorkPan.Tag <> bytFunMode Then
        
            If Val(mobjWorkPan.Tag) = Em_Pane_Record And Not mfrmBlackListRecordManage Is Nothing Then
                Call mfrmBlackListRecordManage.zlCancelBands
            End If
            
            If mfrmBlackTypeManage Is Nothing Then
                Set mfrmBlackTypeManage = New frmBlackTypeManage
                Call mfrmBlackTypeManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackTypeManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            'ˢ���Ӵ���˵���������
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData
    Case Em_Pane_Reason  '����ԭ��
        If mobjWorkPan.Tag <> bytFunMode Then
            If Val(mobjWorkPan.Tag) = Em_Pane_Record And Not mfrmBlackListRecordManage Is Nothing Then
                Call mfrmBlackListRecordManage.zlCancelBands
            End If
            If mfrmBlackListReasonManage Is Nothing Then
                Set mfrmBlackListReasonManage = New frmBlackListReasonManage
                Call mfrmBlackListReasonManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackListReasonManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData
        
    Case Em_Pane_Record '������¼����
        
        If mobjWorkPan.Tag <> bytFunMode Then
            If mfrmBlackListRecordManage Is Nothing Then
                Set mfrmBlackListRecordManage = New frmBlackListRecordManage
                Call mfrmBlackListRecordManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackListRecordManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData(strBlackLitType)
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub mfrmBlackListRecordManage_zlShowStatusText(ByVal bytPancel As Byte, ByVal strText As String)
     stbThis.Panels(bytPancel).Text = strText
End Sub
 


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
    Set mobjWorkPan = Nothing
    Set mfrmCurForm = Nothing
    'ж�����д���
    If Not mfrmBlackTypeManage Is Nothing Then Unload mfrmBlackTypeManage: Set mfrmBlackTypeManage = Nothing
    If Not mfrmBlackListReasonManage Is Nothing Then Unload mfrmBlackListReasonManage: Set mfrmBlackListReasonManage = Nothing
    If Not mfrmBlackListRecordManage Is Nothing Then Unload mfrmBlackListRecordManage: Set mfrmBlackListRecordManage = Nothing
    If Not mfrmBlackListMainFun Is Nothing Then Unload mfrmBlackListMainFun: Set mfrmBlackListMainFun = Nothing
End Sub


Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If mfrmCurForm Is Nothing Then Exit Sub
    Call mfrmCurForm.zlUpdateCommandBars(Control)
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Parameter '��������
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
    Case conMenu_View_Refresh   'ˢ��
            
            
        If mfrmBlackListMainFun Is Nothing Then Exit Sub
        
        If Not (mblnFunListActived And Val(mobjWorkPan.Tag) = 13) Then
            If mfrmCurForm Is Nothing Then Exit Sub
            Call mfrmCurForm.zlExecuteCommandBars(Control)
            Exit Sub
        End If
        Call mfrmBlackListMainFun.zlRefresh
    
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlOpenCustomReport(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        Else
            If mfrmCurForm Is Nothing Then Exit Sub
            Call mfrmCurForm.zlExecuteCommandBars(Control)
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo ErrHandler
    If Item.Tag = Em_Pane_FunFace Then
        Item.Handle = mfrmBlackListMainFun.hWnd
    ElseIf Not mfrmCurForm Is Nothing Then
        Item.Handle = mfrmCurForm.hWnd
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub zlOpenCustomReport(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ص��Զ��屨��
    '���:frmMain-���õĸ�����
    '     lngSys-ϵͳ��
    '     strReprotName-��������
    '����:���˺�
    '����:2018-11-08 11:16:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
End Sub

Public Function GetPopupCommandBarSub() As CommandBar
    '��ȡ�����˵�
    If mfrmCurForm Is Nothing Then Exit Function
    Set GetPopupCommandBarSub = zlGetPopupCommandBar(mfrmCurForm, cbsThis)
End Function
Private Sub mfrmBlackTypeManage_zlChangeType()
    If mfrmBlackListMainFun Is Nothing Then Exit Sub
    Call mfrmBlackListMainFun.zlRefresh(True)
End Sub
