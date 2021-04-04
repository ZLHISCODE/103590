VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDockInTends 
   BorderStyle     =   0  'None
   Caption         =   "�����¼����"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   5115
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   9022
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   150
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDockInTends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mblnSearch As Boolean                           '��ǰʹ�����Ƿ�߱���������(1273)Ȩ
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean
Private mbytFontSize As Byte                            '�����С0-9������,1-12������
Private WithEvents mfrmDockInTendFile As frmDockInTendFile
Attribute mfrmDockInTendFile.VB_VarHelpID = -1
Private WithEvents mfrmDockInTendData As frmDockInTendData
Attribute mfrmDockInTendData.VB_VarHelpID = -1
Private WithEvents mfrmDockInTendEPR As frmDockInTendEPR
Attribute mfrmDockInTendEPR.VB_VarHelpID = -1

Private mcbsThis As Object          'CommandBar�ؼ�
Private eMySignLevel As EPRSignLevelEnum
Public Event Activate()

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С(����ģ���Ѿ����ص���)
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    '�����ļ�
    Call mfrmDockInTendFile.SetFontSize(bytSize)
    '��������
    Call mfrmDockInTendData.SetFontSize(bytSize)
    '������
    Call mfrmDockInTendEPR.SetFontSize(bytSize)
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
            objCtrl.PaintManager.Layout = xtpTabLayoutAutoSize
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        End Select
    Next
End Sub


Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_Resize()
    Call Form_Resize
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

'######################################################################################################################
Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Load()

    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1255)

    mlngPatiId = -1
    mlngPageId = -1
    
    '------------------------------------------
    '����ѡ������
    With Me.tbcThis
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
            .Position = xtpTabPositionRight
            
        End With
        
        Set mfrmDockInTendFile = New frmDockInTendFile
        Call mfrmDockInTendFile.InitData(Me, mstrPrivs)
        
        Set mfrmDockInTendData = New frmDockInTendData
        Call mfrmDockInTendData.InitData(Me, mstrPrivs)
        
        Set mfrmDockInTendEPR = New frmDockInTendEPR
        Call mfrmDockInTendEPR.InitData(mstrPrivs)
        
        .InsertItem(0, "�����ļ�", mfrmDockInTendFile.hWnd, 0).Tag = "_�����ļ�"
        .InsertItem(1, "�����¼", mfrmDockInTendData.hWnd, 0).Tag = "_��������"
        .InsertItem(2, "������", mfrmDockInTendEPR.hWnd, 0).Tag = "_������"
        
        .Item(0).Selected = True
    End With
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsTools.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsTools.VisualTheme = xtpThemeOffice2003
    cbsTools.EnableCustomization False
    Set cbsTools.Icons = zlCommFun.GetPubIcons
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsTools.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If cbsTools(1).Controls.Count = 0 Then lngTop = 0
    tbcThis.Move lngLeft, lngTop, Me.ScaleWidth, Me.ScaleHeight - lngTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmDockInTendFile Is Nothing Then Unload mfrmDockInTendFile
    If Not mfrmDockInTendData Is Nothing Then Unload mfrmDockInTendData
    If Not mfrmDockInTendEPR Is Nothing Then Unload mfrmDockInTendEPR
    Set mfrmDockInTendFile = Nothing
    Set mfrmDockInTendData = Nothing
    Set mfrmDockInTendEPR = Nothing
    Set mcbsThis = Nothing
End Sub


'------------------------------------------------------------
'����Ϊ��������
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object, Optional ByVal blnChildToolBar As Boolean = False)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    eMySignLevel = GetUserSignLevel(glngUserId, , mlngPatiId, mlngPageId) '��ȡ��ǰ�û�ǩ������
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
        
        '���ڵ���ΪXML�ļ�֮��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "�б��ӡ(&T)", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "������¼(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "����¼��(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ļ�¼(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ����¼(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���Ĳ���(&U)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "����鵵(&R)"): cbrControl.Parameter = "����鵵": cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "������(&U)")
        cbrControl.IconId = conMenu_Edit_Archive
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�����鵵(&R)"): cbrControl.Parameter = "�����鵵"
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "������ͼ(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "��¼ǩ��(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "��ʷ�汾(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "��������(&S)"): cbrControl.BeginGroup = True
    End With


    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", cbrMenuBar.Index, False)
        cbrControl.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Option, "����ѡ��(&O)"): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_File_Parameter
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "�����������(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "���˲�������(&S)")
    End With
    
    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    If blnChildToolBar Then
        cbsTools.DeleteAll
        Set cbrToolBar = cbsTools.Add("��������", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
    End If
    With cbrToolBar.Controls
        'Set cbrControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "ȡ����ӡ(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "�鵵", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "����", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "�鵵", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        For Each cbrControl In cbrToolBar.Controls
            cbrControl.STYLE = xtpButtonIconAndCaption
        Next
    End With
    Call cbsTools.RecalcLayout
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_File_PrintDayDetail
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("G"), conMenu_Edit_MarkMap
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F11, conMenu_Tool_Option
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        Call zlPrintSet
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Option                            '����ѡ��
        If Not CreateBodyEditor Then Exit Sub
        
        If gobjBodyEditor.GetCaseTendBodyPara.ShowPara(Me, mstrPrivs) Then
            Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlExecuteCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlExecuteCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlExecuteCommandBars(Control)
        End Select
    End Select
    
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
            
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    Case conMenu_Edit_NoPrint
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Enabled = False
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Control.Enabled = False
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_ExportToXML
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_RowPrint
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintDayDetail    '����¼��

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
                
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
        
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
        
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
        
            Control.Visible = False
            Control.Enabled = False
            
        Case "_������"
            
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select
    Case conMenu_Tool_SignVerify
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Control.Visible = False
            Control.Enabled = False
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
    
        Select Case tbcThis.Selected.Tag
        Case "_�����ļ�"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_��������"
            Control.Visible = False
            Control.Enabled = False
        Case "_������"
            Control.Visible = False
            Control.Enabled = False
        End Select
    End Select
End Sub

'------------------------------------------------------------
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnForce As Boolean, Optional ByVal blnDoctorStation As Boolean, Optional ByVal blnSeekCase As Boolean) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    mlngDeptId = lngDeptId: mblnEdit = blnEdit

    mlngPatiId = lngPatiID: mlngPageId = lngPageId
    
    mblnDoctorStation = blnDoctorStation
    mblnMoved_HL = False
        
    If mlngPatiId <> 0 Then
        gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", mlngPatiId, mlngPageId)
        mblnMoved_HL = NVL(rs!����ת��, 0) <> 0
    End If
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendData.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendEPR.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, , mblnMoved_HL)
    
    
End Function
    
Public Sub zlLocateData(ByVal intType As Integer)
    tbcThis.Item(intType).Selected = True
End Sub

Private Sub mfrmDockInTendData_AfterArchiveChanged(ByVal blnArchived As Boolean)
    mfrmDockInTendFile.TendArchive = blnArchived
End Sub

Private Sub mfrmDockInTendData_AfterDataChanged()
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
End Sub

Private Sub mfrmDockInTendData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    Select Case Button
    Case 2

        If mcbsThis.ActiveMenuBar.Controls(3).Visible = False Then Exit Sub

        Set cbrMenuBar = mcbsThis.ActiveMenuBar.Controls(3)
        Set cbrPopupBar = mcbsThis.Add("�����˵�", xtpBarPopup)
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            cbrPopupItem.IconId = cbrControl.IconId
        Next
        cbrPopupBar.ShowPopup

    End Select
End Sub

Private Sub mfrmDockInTendData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objTmp As CommandBarControl

    Set objTmp = mcbsThis.FindControl(, conMenu_Edit_Modify)
    If Not (objTmp Is Nothing) Then
        If objTmp.Enabled And objTmp.Visible Then
            Call zlExecuteCommandBars(objTmp)
        End If
    End If

End Sub

Private Sub mfrmDockInTendEPR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    Select Case Button
    Case 2

        If mcbsThis.ActiveMenuBar.Controls(3).Visible = False Then Exit Sub

        Set cbrMenuBar = mcbsThis.ActiveMenuBar.Controls(3)
        Set cbrPopupBar = mcbsThis.Add("�����˵�", xtpBarPopup)
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            cbrPopupItem.IconId = cbrControl.IconId
        Next
        cbrPopupBar.ShowPopup

    End Select
End Sub

Private Sub mfrmDockInTendFile_AfterDataChanged()
    Call mfrmDockInTendData.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
End Sub
