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
Private mlngPatiID As Long                              '����id
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
Private mbytFontSize As Byte

Private Enum enmSEL
    ����
    ����
End Enum

Public Event Activate()
Public Event RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event StartTimer(ByVal blnStart As Boolean)

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize(False)
End Sub

Private Sub ReSetFontSize(Optional ByVal blnStart As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С(����ģ���Ѿ����ص���)
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    
    If mlngPatiID = 0 Then Exit Sub
    
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    Call mfrmDockInTend_Data.SetFontSize(bytSize)
    Call mfrmDockInTend_TendList.SetFontSize(bytSize)

    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
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


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmDockInTend_TendList.hWnd
    Case 2
        Item.Handle = mfrmDockInTend_Data.hWnd
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
        '53330,������,2012-09-04,ȡ���öδ��룬���ҽ��վ�������Exeִ�б��δ�����򱨴�����
'        mfrmDockInTend_TendList.Show
'        mfrmDockInTend_Data.Show
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
    Set cbsTools.Icons = ZLCommFun.GetPubIcons
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
    Set cbsThis.Icons = ZLCommFun.GetPubIcons
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '�������:���ڵ�һ��
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '���������Excel֮��
        '51588,2012-12-12,������,�����ļ����������ӡ
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print * 100# + 1, "������ӡ(&L)��", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_File_Print
        cbrControl.BeginGroup = True
        
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "���߱༭(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "���༭(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "������ʾ(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surgery_Edit, "����/��������(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "��¼ǩ��(&S)"): cbrControl.BeginGroup = True
        '51589:������,2013-03-01,��ӽ���ǩ��
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignShiftExchange, "����ǩ��(&K)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): cbrControl.IconId = 229
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "�ϼ���ǩ(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "ȡ����ǩ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����¼��(&E)"): cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "����", cbrControl.Index + 1): cbrControl.ToolTipText = "�������߱༭": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "���", cbrControl.Index + 1): cbrControl.ToolTipText = "���±��༭"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "��ʾ", cbrControl.Index + 1): cbrControl.ToolTipText = "����������ʾ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surgery_Edit, "����", cbrControl.Index + 1): cbrControl.ToolTipText = "��������/����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "ǩ��", cbrControl.Index + 1): cbrControl.BeginGroup = True
        '51589:������,2013-03-01,��ӽ���ǩ��
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignShiftExchange, "����ǩ��", cbrControl.Index + 1): cbrControl.ToolTipText = "���Ӱ�ǩ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "��ǩ", cbrControl.Index + 1)
       ' Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("E"), conMenu_Edit_Billing
        .Add FCONTROL, Asc("L"), conMenu_File_Print * 100# + 1
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F11, conMenu_Tool_Option
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFormat As Long, lng��� As Long
    
    Select Case Control.ID
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_Option '����ѡ��
        Call mfrmDockInTend_TendList.zlExecuteCommandBars(Control)
    Case conMenu_Edit_FileMan
        '�õ��±༭���ļ��ĸ�ʽID,������ݸ�ʽID��λ���һ���ļ�
        If frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, False, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), lngFormat, lng���) Then
            Call mfrmDockInTend_TendList.RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, lngFormat, lng���)
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
    Case conMenu_Help_Help
    
    Case conMenu_Tool_Option
        Call mfrmDockInTend_TendList.zlUpdateCommandBars(Control)
    Case Else
        Call mfrmDockInTend_TendList.zlUpdateCommandBars(Control)
    End Select
End Sub

'------------------------------------------------------------
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnForce As Boolean, Optional ByVal blnDoctorStation As Boolean, Optional ByVal blnSeekCase As Boolean, Optional ByVal intCurveReSize As Integer = 0) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    mlngDeptId = lngDeptID: mblnEdit = blnEdit
    mlngPatiID = lngPatiID: mlngPageId = lngPageId
    mblnDoctorStation = blnDoctorStation
    Call mfrmDockInTend_TendList.RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, , , intCurveReSize, False)
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

Private Sub mfrmDockInTend_Data_ISChartArchive(ByVal blnArchive As Boolean)
    Call mfrmDockInTend_TendList.SetState(True, blnArchive)
End Sub

Private Sub mfrmDockInTend_Data_StartTimer(ByVal blnStart As Boolean)
    Call mfrmDockInTend_TendList.StartTimer(blnStart)
End Sub

Private Sub mfrmDockInTend_Data_zlRefreshViewFile()
    Call mfrmDockInTend_TendList.zlRefreshViewFile
End Sub

Private Sub mfrmDockInTend_TendList_Activate()
'    On Error Resume Next
'    Me.SetFocus
End Sub

Private Sub mfrmDockInTend_TendList_ArchiveDocument(blnOK As Boolean)
    Call mfrmDockInTend_Data.zlArchiveDocument(blnOK)
End Sub


Private Sub mfrmDockInTend_TendList_BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)
    Call mfrmDockInTend_Data.BulkPrintDocument(lngPatiID, lngPageId, lngDeptID, intBaby)
End Sub

Private Sub mfrmDockInTend_TendList_PrintTendFile(ByVal bytKind As Byte, ByVal bytMode As Byte)
    Call mfrmDockInTend_Data.zlPrintTendFile(bytKind, bytMode)
End Sub

Private Sub mfrmDockInTend_TendList_SaveDocument(blnSave As Boolean)
    Call mfrmDockInTend_Data.zlSaveDocument(blnSave)
End Sub

Private Sub mfrmDockInTend_TendList_ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte, ByVal intCurveReSize As Integer)
    mintBaby = intBaby
    Call mfrmDockInTend_Data.zlRefreshTend(mlngPatiID, mlngPageId, intBaby, lngDept, mblnEdit, mblnDoctorStation, lngFile, bytSel, intCurveReSize)
End Sub

Private Sub mfrmDockInTend_TendList_SignDocument(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)
    Call mfrmDockInTend_Data.zlSignDocument(blnOK, blnVerify, blnExchange)
End Sub

Private Sub mfrmDockInTend_TendList_SignMarker()
    Call mfrmDockInTend_Data.SignMarker
End Sub

Private Sub mfrmDockInTend_TendList_ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewAnimalHeat(strPara, bytMode, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewCaveData(ByVal intDataEditor As Integer)
    Call mfrmDockInTend_Data.zlViewCaveData(intDataEditor)
End Sub

Private Sub mfrmDockInTend_TendList_ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewFile(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit, bytSize)
End Sub


Private Sub mfrmDockInTend_TendList_Viewpartogram(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewpartogram(strPara, bytMode, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewpartogramEditor(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
     Call mfrmDockInTend_Data.ViewReSetFontSize(intSEL, bytSize)
End Sub
