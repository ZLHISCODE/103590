VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDockInTendMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmDockInTendMain.frx":0000
      Left            =   390
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInTendMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mblnFirst As Boolean
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mblnSearch As Boolean                           '��ǰʹ�����Ƿ�߱���������(1273)Ȩ
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mintBaby As Integer
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean

Private WithEvents mfrmDockInTend_TendList As frmDockInTend_TendList
Attribute mfrmDockInTend_TendList.VB_VarHelpID = -1
Private WithEvents mfrmDockInTend_Data As frmDockInTend_Data
Attribute mfrmDockInTend_Data.VB_VarHelpID = -1

Private mcbsThis As Object          'CommandBar�ؼ�
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar
Private rsTemp As New ADODB.Recordset
Private mintPageSel As Integer

Private Enum enmSEL
    ����
    ����
End Enum

Public Event Activate()
Public Event RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmDockInTend_TendList.hwnd
    Case 2
        Item.Handle = mfrmDockInTend_Data.hwnd
    End Select
End Sub

Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mfrmDockInTend_TendList.Show
        mfrmDockInTend_Data.Show
        mblnFirst = False
    End If
    
    RaiseEvent Activate
End Sub

Private Sub InitDOCK()
    Dim objPane As Pane
    With DkpMain
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.CloseGroupOnButtonClick = True
        .Options.HideClient = True
        .SetCommandBars cbsTools
        
        Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "�ļ��б�": objPane.Options = PaneNoCaption
        Set objPane = .CreatePane(2, 500, 500, DockRightOf, objPane): objPane.Title = "����ҳ��": objPane.Options = PaneNoCaption
    End With
End Sub

Private Sub InitCommandBar()
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
    cbsTools.Icons = frmPubIcons.imgPublic.Icons
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1255)
    
    '���ش���
    Set mfrmDockInTend_TendList = New frmDockInTend_TendList
    Call mfrmDockInTend_TendList.InitData(Me, mstrPrivs)
    Load mfrmDockInTend_TendList
    Set mfrmDockInTend_Data = New frmDockInTend_Data
    Call mfrmDockInTend_Data.InitData(Me, mstrPrivs)
    Load mfrmDockInTend_Data
    
    Call InitDOCK
    Call InitCommandBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmDockInTend_TendList Is Nothing Then Unload mfrmDockInTend_TendList
    If Not mfrmDockInTend_Data Is Nothing Then Unload mfrmDockInTend_Data
End Sub


'------------------------------------------------------------
'����Ϊ��������
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object, Optional ByVal blnChildToolBar As Boolean = False)
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "�ļ�����(&N)")
    
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "����¼��(&B)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "��¼ǩ��(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): cbrControl.IconId = 229
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "�ϼ���ǩ(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "ȡ����ǩ(&C)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "����鵵(&R)"): cbrControl.Parameter = "����鵵": cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "������(&U)")
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "�ļ�", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "��ǩ", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "�鵵", cbrControl.Index + 1): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        For Each cbrControl In cbrToolBar.Controls
            cbrControl.Style = xtpButtonIconAndCaption
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
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_Option                            '����ѡ��
        If Not CreateBodyEditor Then Exit Sub

        If gobjBodyEditor.GetCaseTendBodyPara.ShowPara(Me, mstrPrivs) Then
            Call mfrmDockInTend_TendList.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
        End If
    Case conMenu_Edit_FileMan
        If frmNurseFileMan.ShowEditor(mlngPatiId, mlngPageId, mintBaby) Then
            Call mfrmDockInTend_TendList.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
        End If
    Case Else
'        If mfrmDockInTend_Data.tbcData.Selected.Index = 2 Then   '����ҳ��
'            Call mfrmDockInTend_Data.zlExecuteCommandBars(Control)
'        Else
            Call mfrmDockInTend_TendList.zlExecuteCommandBars(Control)
'        End If
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Select Case Control.ID
    Case conMenu_Help_Help, conMenu_Tool_Option
    Case Else
        Call mfrmDockInTend_TendList.zlUpdateCommandBars(Control)
    End Select
End Sub

'------------------------------------------------------------
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnForce As Boolean, Optional ByVal blnDoctorStation As Boolean, Optional ByVal blnSeekCase As Boolean) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Rs As New ADODB.Recordset
    mlngDeptId = lngDeptID: mblnEdit = blnEdit
    mlngPatiId = lngPatiID: mlngPageId = lngPageId
    mblnDoctorStation = blnDoctorStation
    
    Call mfrmDockInTend_TendList.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
End Function
    
Public Sub zlLocateData(ByVal intType As Integer)
'    tbcData.Item(intType).Selected = True
End Sub

Private Sub mfrmDockInTend_Data_Activate()
    On Error Resume Next
    Me.SetFocus
End Sub

Private Sub mfrmDockInTend_Data_AfterDataChanged(ByVal blnChange As Boolean)
    Call mfrmDockInTend_TendList.SetChange(blnChange)
End Sub

Private Sub mfrmDockInTend_Data_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent RefreshPrompt(strInfo, blnImportant)
    Call mfrmDockInTend_TendList.SetState(blnSign, blnArchive)
End Sub

Private Sub mfrmDockInTend_TendList_Activate()
'    On Error Resume Next
'    Me.SetFocus
End Sub

Private Sub mfrmDockInTend_TendList_ArchiveDocument(blnOK As Boolean)
    Call mfrmDockInTend_Data.zlArchiveDocument(blnOK)
End Sub

Private Sub mfrmDockInTend_TendList_PrintDocument(ByVal bytKind As Byte, ByVal bytMode As Byte)
    Call mfrmDockInTend_Data.zlPrintDocument(bytKind, bytMode)
End Sub

Private Sub mfrmDockInTend_TendList_SaveDocument(blnSave As Boolean)
    Call mfrmDockInTend_Data.zlSaveDocument(blnSave)
End Sub

Private Sub mfrmDockInTend_TendList_ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte)
    mintBaby = intBaby
    Call mfrmDockInTend_Data.zlRefreshTend(mlngPatiId, mlngPageId, intBaby, lngDept, mblnEdit, mblnDoctorStation, lngFile, bytSel)
End Sub

Private Sub mfrmDockInTend_TendList_SignDocument(blnOK As Boolean, blnVerify As Boolean)
    Call mfrmDockInTend_Data.zlSignDocument(blnOK, blnVerify)
End Sub

Private Sub mfrmDockInTend_TendList_ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String)
    Call mfrmDockInTend_Data.zlViewAnimalHeat(strPara, bytMode, strPrivs)
End Sub

Private Sub mfrmDockInTend_TendList_ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean)
    Call mfrmDockInTend_Data.zlViewFile(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit)
End Sub
