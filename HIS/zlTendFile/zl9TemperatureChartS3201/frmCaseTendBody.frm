VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCaseTendBody 
   Caption         =   "������ͼ"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin zl9TemperatureChartS3201.usrBodyEditor BodyEdit 
      Height          =   3495
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6165
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
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
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'���˻�����Ϣ
'***************************************************************
Private Type type_Patient
    lng����id As Long
    lng��ҳid As Long
    lng����ID As Long
    lng����ID As Long
    lng��Ժ As Long
    lngӤ�� As Long
    lng�༭ As Long
    lng����ȼ� As Long
    lng�ļ�ID As Long
    lngԭʼ��С As Long
    lngPage As Long
End Type

Private T_Info As type_Patient

Private mblnChildForm As Boolean
Private mcbrToolBar As CommandBar
Private mcbr�鿴 As CommandBarControl
Private mstrPrivs As String
Private mstrSQL As String
Private mblnShowing As Boolean
Private mblnChanged As Boolean
Private mfrmMain As Form
Private mIntDataEditor As Integer
Private mblnMove As Boolean
Private mfrmTendBody As Object

Public Event AfterPrint()
Public Event CmdClick(ByVal strParam As String)

'######################################################################################################################
'�Զ��庯������������

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mblnMove = False
    mblnChildForm = True
    mblnChanged = False
    mstrPrivs = strPrivs
    
    blnShowing = mblnShowing
    
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    If blnShowing Then
        If Val(varParam(0)) = T_Info.lng����id Or Val(varParam(1)) = T_Info.lng��ҳid And T_Info.lng����ID = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set BodyEdit.ParentForm = Me
    Set mfrmMain = frmMain

    '������ʽ������ID;��ҳID;����ID;�ļ�ID;��Ժ;�༭;Ӥ��;�Ƿ���ߴ����С�Զ�У�����µ���ʽ(1 �� 0 У��)ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    
    '��ʼ������
    
    T_Info.lngӤ�� = 0
    T_Info.lng��Ժ = 0
    T_Info.lng�༭ = 0
    T_Info.lngԭʼ��С = 0
    T_Info.lngPage = 0
    
    T_Info.lng����id = Val(varParam(0))
    T_Info.lng��ҳid = Val(varParam(1))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng�ļ�ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng��Ժ = Val(varParam(4))
    If UBound(varParam) > 4 Then T_Info.lng�༭ = Val(varParam(5))
    If InStr(1, ";" & mstrPrivs & ";", ";���µ���ͼ;") = 0 Then
        T_Info.lng�༭ = 0
    Else
        T_Info.lng�༭ = 1
    End If
    If UBound(varParam) > 5 Then T_Info.lngӤ�� = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lngԭʼ��С = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    If blnShowing = False Then Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID,nvl(����ת��,0) ת��  from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����id, T_Info.lng��ҳid)
    If RS.BOF = False Then
        T_Info.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
        If T_Info.lng��Ժ = 1 Then mblnMove = (Val(RS("ת��")) <> 0)
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    If blnShowing = False Then
        Hook Me
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        
        ShowEdit = mblnChanged
    End If
End Function

Public Function zlInit() As Boolean
    mblnChildForm = True
End Function

Public Function GetCurvePage() As Long
   GetCurvePage = BodyEdit.intPage
End Function

Public Sub zlDataEditor(ByVal intDataEditor As Integer)
    BodyEdit.DateEditor = intDataEditor
End Sub

Public Function zlRefresh(ByVal frmParent As Form, strParam As String, Optional strPrivs As String) As Boolean

   '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim intBaby As Integer
    
    mblnMove = False
    mstrPrivs = strPrivs
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    mblnChanged = False
    
    Set BodyEdit.ParentForm = frmParent
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    '������ʽ������ID;��ҳID;����ID;�ļ�ID;��Ժ;�༭;Ӥ��;�Ƿ���ߴ����С�Զ�У�����µ���ʽ(1 �� 0У��);ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    If Val(varParam(3)) <> T_Info.lng�ļ�ID Then
        glngCurPage = 0
    Else
        If UBound(varParam) > 5 Then
            intBaby = Val(varParam(6))
        Else
            intBaby = 0
        End If
        
        If T_Info.lngӤ�� <> intBaby Then
            glngCurPage = 0
        End If
    End If
    
    '��ʼ������
    T_Info.lngӤ�� = 0
    T_Info.lng��Ժ = 0
    T_Info.lng�༭ = 0
    T_Info.lngԭʼ��С = 0
    T_Info.lngPage = 0
    
    T_Info.lng����id = Val(varParam(0))
    T_Info.lng��ҳid = Val(varParam(1))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng�ļ�ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng��Ժ = Val(varParam(4))
    If UBound(varParam) > 4 Then T_Info.lng�༭ = Val(varParam(5))
    If InStr(1, ";" & mstrPrivs & ";", ";���µ���ͼ;") = 0 Then
        T_Info.lng�༭ = 0
    Else
        T_Info.lng�༭ = 1
    End If
    If UBound(varParam) > 5 Then T_Info.lngӤ�� = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lngԭʼ��С = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID,nvl(����ת��,0) ת�� from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����id, T_Info.lng��ҳid)
    If RS.BOF = False Then
        T_Info.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
        If T_Info.lng��Ժ = 1 Then mblnMove = (Val(RS("ת��")) <> 0)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    zlRefresh = True
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    
    T_Info.lng����ȼ� = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����id, T_Info.lng��ҳid)
    If RS.BOF = False Then T_Info.lng����ȼ� = zlCommFun.Nvl(RS("����ȼ�"), 3)
    
    '������ȡ�ļ�ID
    gstrSQL = "select A.ID from ���˻����ļ� A,�����ļ��б� B" & _
       "    where A.����ID=[1] and A.��ҳId=[2] and A.Ӥ��=[3] and A.����ID=[4] and A.��ʽID=B.ID and B.����=3 and B.����=-1"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����id, T_Info.lng��ҳid, T_Info.lngӤ��, T_Info.lng����ID)
    If mblnMove = True Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    
    If RS.BOF = False Then T_Info.lng�ļ�ID = Val(zlCommFun.Nvl(RS("ID")))
    '��ʼ�����߲˵�
    If InitBodyLine = False Then Exit Function
    
    '����������ID;��ҳID;����ID;�ļ�ID;��Ժ��־;�༭��־;Ӥ��;����ȼ�;ԭʼ��С;ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    strParam = T_Info.lng����id & ";" & T_Info.lng��ҳid & ";" & T_Info.lng����ID & ";" & T_Info.lng�ļ�ID & ";" & _
        T_Info.lng��Ժ & ";" & T_Info.lng�༭ & ";" & T_Info.lngӤ�� & ";" & T_Info.lng����ȼ� & ";" & T_Info.lngԭʼ��С & ";" & T_Info.lngPage
    Call BodyEdit.zlMenuClick("��ʼ��", strParam)
        
    OpenPatientMap = True
    
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo errHand
    
    '--�������ü��
    mstrSQL = "SELECT A.��¼��,A.��Ŀ��� FROM ���¼�¼��Ŀ A,�����¼��Ŀ B " & _
            "WHERE A.��¼�� =1 And A.��Ŀ���=B.��Ŀ��� AND B.����ȼ�>=[1]  And Nvl(b.Ӧ�÷�ʽ,0)=1 " & _
            "ORDER BY A.�������"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng����ȼ�)
    If rsTmp.BOF Then
        MsgBox "�����µ�����������Ŀ�����ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '--��¼Ƶ��ʱ������ü��
    mstrSQL = " Select Distinct nvl(��¼Ƶ��,2) Ƶ��  From ���¼�¼��Ŀ A,�����¼��Ŀ B" & _
            "   WHERE A.��¼�� =2 And A.��Ŀ���<>3 And  ��Ŀ��ʾ<>4 And A.��Ŀ���=B.��Ŀ��� AND B.����ȼ�>=[1] And Nvl(b.Ӧ�÷�ʽ,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng����ȼ�)
    
    Do While Not rsTmp.EOF
        strSQL = "select Count(*) ��¼�� From ������ĿƵ�� where Ƶ��=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!Ƶ��))
        If Val(rsData!��¼��) < Val(rsTmp!Ƶ��) Then
            MsgBox "������Ŀ��¼Ƶ��ʱ�����ò����������ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    rsTmp.MoveNext
    Loop
    '--������Ŀʱ������ü��
    mstrSQL = "select count(*) ��¼�� from �������ʱ�� Where ����=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If Val(rsTmp!��¼��) < 3 Then
        MsgBox "�������ʱ�����ò����������ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitBodyLine = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    Dim strPage  As String, strParam As String
    
    '�����˴�ӡ������,˵����������ӡ,�Զ��ӵ�1ҳ��ʼ��ӡ,�������κ�ѯ��
    '����:0-ȡ��,2-Ԥ��,1-��ӡ
    
    frmCaseTendBodyPrintSet.cmdPrint.Visible = (bytMode = 1)
    frmCaseTendBodyPrintSet.cmdPreview.Visible = (bytMode = 2)
    
    
    If strPrintDevice = "" Then
        'strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����Id & ";" & T_Info.lng����Id
        strParam = T_Info.lng�ļ�ID & ";" & Me.BodyEdit.AllPage
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, strParam, intPrintRange, lngBeginY, intBeginPage, strPage, mstrPrivs)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
    
    '��ӡ��ǰҳ���뵱ǰҳ��
    If intPrintRange = 0 Then
        strPage = Me.BodyEdit.intPage - 1
    End If
    
    Select Case bytMode
    Case 2  '��ӡ
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice, strPage)
    Case 1  'Ԥ��
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice, strPage)
    End Select

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '���:1-Ԥ��,2-��ӡ
    '����ֵ:0-ʧ��;1-�ɹ�;2-��ӡ
    gblnPrinted = False
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
       
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "��������(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�ָ�����(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "���ü�¼(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "������ʾ(&D)")
    End With

    Set mcbr�鿴 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With mcbr�鿴.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
'    Set mcbrToolBar = cbsThis.Add("��׼", xtpBarTop)
'    mcbrToolBar.ShowTextBelowIcons = False
'    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
'
'    With mcbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�ָ�")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
'        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
'    End With
'
'    '��λ������
'    '------------------------------------------------------------------------------------------------------------------
'
'    For Each cbrControl In mcbrToolBar.Controls
'        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
'            cbrControl.Style = xtpButtonIconAndCaption
'        End If
'    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Adjust
        .Add FCONTROL, Asc("D"), conMenu_View_Show
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub BodyEditCur(ByVal intDataEditor As Integer, Optional ByVal strParam As String = "")
    Call GetTendEidor
    If intDataEditor = 0 Then
        Call BodyEdit.zlMenuClick("�������ݱ༭", strParam)
    ElseIf intDataEditor = 1 Then
         Call BodyEdit.zlMenuClick("����������ʾ����", strParam)
    End If
End Sub

Private Sub BodyEdit_DbClickCur(ByVal intDataEditor As Integer)
    Call BodyEditCur(intDataEditor)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_File_PrintSet   '��ӡ����
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '��ӡԤ��
            
            Call PrintData(2)
            
        Case conMenu_File_Print  '��ӡ
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button

'            cbsThis(2).Visible = Not cbsThis(2).Visible
'            cbsThis.RecalcLayout

        Case conMenu_View_ToolBar_Text

'            For Each cbrControl In cbsThis(1).Controls
'                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
'            Next
'
'            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_Adjust '���ü�¼
            Call BodyEditCur(0)
            
        Case conMenu_View_Show '������ʾ
            Call BodyEditCur(1)
            
        Case conMenu_Edit_Save '��������
            
        Case conMenu_Edit_Reuse '���ݻָ�
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Adjust, conMenu_View_Show
        
        Control.Enabled = (T_Info.lng�༭ = 1)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .mblnResize = True
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .mblnResize = False
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub Form_Load()
    If Not mblnChildForm Then
         Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub GetTendEidor()
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    Set gobjTendEditor = Me
End Sub

Private Sub BodyEdit_CmdClick(ByVal strParam As String)
    Dim arrParam() As String
    If mfrmTendBody Is Nothing Then Set mfrmTendBody = New frmCaseTendBody
    
    If mfrmTendBody.ShowEdit(BodyEdit.ParentForm, strParam, 0, mstrPrivs) Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) > 6 Then arrParam(7) = 0
        If UBound(arrParam) > 7 Then
            strParam = arrParam(0) & ";" & arrParam(1) & ";" & arrParam(2) & ";" & arrParam(3) & ";" & arrParam(4) & ";" & arrParam(5) & ";" & arrParam(6) & ";" & arrParam(7)
        Else
            strParam = Join(arrParam, ";")
        End If
        
        'ˢ�����µ�ҳ��
        Call zlRefresh(BodyEdit.ParentForm, strParam, mstrPrivs)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook Me
    
    mblnShowing = False
    Set mfrmTendBody = Nothing
    
    If Not mblnChildForm Then
        Call SaveWinState(Me, App.ProductName)
    Else
        mblnChanged = True
    End If
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    
    'ж���û��ؼ����� ������ر�ʱ�û��ؼ��� UserControl_Terminate �¼��޷����� ���Է��ڸ�����ر�ִ�� ��
    Call BodyEdit.ReleaseObj
End Sub

