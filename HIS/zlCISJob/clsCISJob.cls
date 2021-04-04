VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Public Type PatiInfo
    ���� As String
    ״̬ As Integer '������ҳ.״̬
    ���� As Integer '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    סԺ�� As String
    ���� As String
    ��ҳID As Long
    ����ID As Long
    ����ID As Long
    ���� As Boolean
    ��Ժ���� As Date
    ��Ժ���� As Date
    ·��״̬ As Integer
    ����״̬ As Integer
    ����ת�� As Boolean
    ���� As Integer
    ���� As Boolean
End Type

Public Type TYPE_CATE
    strName As String
    lngBeginRow As Long
    lngEndRow As Long
End Type
Private mobjArchiveView As frmArchiveView
Private mobjOneArchiveView As frmArchiveView '��ʾĳһ�����˵ĵ��Ӳ���
Private mobjArchiveViewAndPrint As frmArchiveViewAndPrint

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���
    Dim objPlugIn As Object
    Dim strMsg As String, intMode As Integer
    On Error GoTo ErrHand

    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub
    On Error Resume Next
    Set gobjEmr = gfrmMain.mobjEMR  '�²�����Ҫ����
    If err.Number <> 0 Then
        Set gobjEmr = Nothing
        err.Clear
    End If
    
    On Error GoTo ErrHand
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    If frmMain Is Nothing Then  'ҽ���ӿڵ��Ե���
        gblnShowInTaskBar = blnShowInTaskBar
    Else
        gblnShowInTaskBar = blnShowInTaskBar Or frmMain.frmHide
    End If
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '��ʼ�����Ĳ���
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, blnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs, gobjEmr)
    
    '�ж��ٴ�����վ����ʹ�ò��
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara(208, 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "��ǰϵͳ����Ϊ����ʹ�ò��(zlPlugIn)������ȱ�ٸò��������ϵ����Ա��", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If
    '����°滤ʿ����վ�Ƿ��������廤������
    If lngModul = P�°滤ʿվ Or lngModul = pסԺҽ��վ Then
        '����Ƿ����������廤��(�����������ӿڲ������ڡ�������������ȷ�����û���¼�ɹ�)
        gbln�������廤���ӿ� = False
        If InitNurseIntegrate = True Then
            gbln�������廤���ӿ� = gobjNurseIntegrate.IsUseIntergrate(strMsg, intMode)
            If gbln�������廤���ӿ� = False Then
                If intMode = 1 Then
                    MsgBox "���廤������������ʧ�ܣ����޷��������廤����ز�����" & vbCrLf & "��ϸ��Ϣ��" & strMsg, vbInformation, gstrSysName
                ElseIf intMode = 2 Then
                    'IE�汾���Բ���ʾ���������ֵ���δװIE11��Ҳ��ʹ�����廤�����˴�������ʾ��
'                    MsgBox strMsg, vbInformation, gstrSysName
                End If
            Else
                If lngModul = P�°滤ʿվ Then
                    If gobjNurseIntegrate.UserLogin(strMsg) = False Then
                        MsgBox "���¼��������������ʧ�ܣ����޷��������廤����ز�����" & vbCrLf & "��ϸ��Ϣ��" & strMsg, vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    End If
    '���ò˵�Ӧ��ģ��
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '����ҽ������վ
            Call InitCommunity
            frmOutDoctorStation.Show , frmMain
        Case 1261 'סԺҽ������վ
            frmInDoctorStation.ShowMe frmMain
        Case 1262 'סԺ��ʿ����վ
            frmInNurseStation.Show , frmMain
        Case 1263 'ҽ������վ
            frmTechnicStation.Show , frmMain
        Case 1265 '�°�סԺ��ʿ����վ
            frmSublimeInNurseStation.Show , frmMain
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274, 1277
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1078, 1275, 1083, 1272 '�ٴ�·������,�ٴ�·������,����·������,����·������
            Call gobjPath.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1266
            frmExamineKSS.Show , frmMain
        Case 1267
            frmExamineOPS.Show , frmMain
        Case 1268
            frmExamineTransfuse.ShowMe frmMain, 2
        Case 1269
            frmKssStatistics.Show , frmMain
        Case 1284 'Σ��ֵ��ѯ
            Call gobjKernel.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1910  'ѪҺ���յǼ�
            If gblnѪ��ϵͳ = False Then
                MsgBox "Ѫ��ϵͳδ������װ��δ��ѡϵͳ��������Ѫ��ϵͳ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood = True Then
                Call gobjPublicBlood.zlShowBloodReceives(frmMain, 100, lngModul, 0)
            End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���
    Dim objPlugIn As Object
    Dim strMsg As String
    Dim intMode As Integer
    
    Set gfrmMain = Nothing
    gblnShowInTaskBar = True
    'If Not CheckValid Then Exit Sub
    
    If Not objLogin Is Nothing Then
        Set gobjEmr = objLogin.mobjEMR  '�²�����Ҫ����
        If err.Number <> 0 Then
            Set gobjEmr = Nothing
            err.Clear
        End If
    End If
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '��ʼ�����Ĳ���
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, gblnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    
    '�ж��ٴ�����վ����ʹ�ò��
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara(208, 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "��ǰϵͳ����Ϊ����ʹ�ò��(zlPlugIn)������ȱ�ٸò��������ϵ����Ա��", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If
    '����°滤ʿ����վ�Ƿ��������廤������
    If lngModul = P�°滤ʿվ Or lngModul = pסԺҽ��վ Then
        '����Ƿ����������廤��(�����������ӿڲ������ڡ�������������ȷ)
        gbln�������廤���ӿ� = False
        If InitNurseIntegrate = True Then
            gbln�������廤���ӿ� = gobjNurseIntegrate.IsUseIntergrate(strMsg, intMode)
            If gbln�������廤���ӿ� = False Then
                If intMode = 1 Then
                    MsgBox "���廤������������ʧ�ܣ����޷��������廤����ز�����" & vbCrLf & "��ϸ��Ϣ��" & strMsg, vbInformation, gstrSysName
                ElseIf intMode = 2 Then
                    'IE�汾���Բ���ʾ���������ֵ���δװIE11��Ҳ��ʹ�����廤�����˴�������ʾ��
'                    MsgBox strMsg, vbInformation, gstrSysName
                End If
            Else
                If lngModul = P�°滤ʿվ Then
                    If gobjNurseIntegrate.UserLogin(strMsg) = False Then
                        MsgBox "���¼��������������ʧ�ܣ����޷��������廤����ز�����" & vbCrLf & "��ϸ��Ϣ��" & strMsg, vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    End If
    '���ò˵�Ӧ��ģ��
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '����ҽ������վ
            Call InitCommunity
            Call OS.ShowChildWindow(frmOutDoctorStation.hwnd, lngMain, True)
        Case 1261 'סԺҽ������վ
            Call OS.ShowChildWindow(frmInDoctorStation.hwnd, lngMain, True)
        Case 1262 'סԺ��ʿ����վ
            Call OS.ShowChildWindow(frmInNurseStation.hwnd, lngMain, True)
        Case 1263 'ҽ������վ
            Call OS.ShowChildWindow(frmTechnicStation.hwnd, lngMain, True)
        Case 1265 '�°�סԺ��ʿ����վ
            Call OS.ShowChildWindow(frmSublimeInNurseStation.hwnd, lngMain, True)
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1078, 1275 '�ٴ�·������,�ٴ�·������
            Call gobjPath.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1266
            Call OS.ShowChildWindow(frmExamineKSS.hwnd, lngMain, True)
        Case 1267
            Call OS.ShowChildWindow(frmExamineOPS.hwnd, lngMain, True)
        Case 1268
            Call OS.ShowChildWindow(frmExamineTransfuse.hwnd, lngMain, True)
        Case 1269
            Call OS.ShowChildWindow(frmKssStatistics.hwnd, lngMain, True)
        Case 1910 'ѪҺ���յǼ�
            If gblnѪ��ϵͳ = False Then
                MsgBox "Ѫ��ϵͳδ������װ��δ��ѡϵͳ��������Ѫ��ϵͳ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood(True) = True Then
                Call gobjPublicBlood.zlShowBloodReceives(lngMain, 100, lngModul, 0)
            End If
    End Select
End Sub

Private Sub InitCommunity()
'���ܣ���ڳ�ʼ�������ӿڶ���
    If gobjCommunity Is Nothing Then
        On Error Resume Next
        Set gobjCommunity = CreateObject("zlCommunity.clsCommunity")
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjCommunity Is Nothing Then
        If Not gobjCommunity.Initialize(gcnOracle) Then
            Set gobjCommunity = Nothing
        End If
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    Call gobjEmr.CloseForms
    blnChildren = gobjRichEPR.CloseWindows()
    blnChildren = blnChildren And gobjKernel.CloseWindows()
    blnChildren = blnChildren And gobjPath.CloseWindows()
    '���ýӿڹر��ƶ��ؼ�����,�����޷��ͷŽ���
    If gbln�������廤���ӿ� = True Then
        If InitNurseIntegrate = True Then
            Call gobjNurseIntegrate.UnloadWebKixt
        End If
    End If
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    If Not gobjCommunity Is Nothing Then
        Call gobjCommunity.Terminate
        Set gobjCommunity = Nothing
    End If
    If Not gobjPass Is Nothing Then
        Set gobjPass = Nothing
    End If
    On Error Resume Next
    Unload mobjOneArchiveView
    Unload mobjArchiveView
    Unload mobjArchiveViewAndPrint
    Set mobjOneArchiveView = Nothing
    Set mobjArchiveView = Nothing
    Set mobjArchiveViewAndPrint = Nothing
End Sub

Public Sub ShowArchive(ByRef objForm As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal blnModal As Boolean = False)
'���ܣ����õ��Ӳ������Ĺ��ܣ��������˲���
    If mobjOneArchiveView Is Nothing Then Set mobjOneArchiveView = New frmArchiveView
    Call mobjOneArchiveView.ShowArchive(objForm, lng����ID, lng��ҳID, blnModal)
End Sub

Public Function GetArchiveOneFrom() As Object
'���ܣ����õ��Ӳ������Ĺ��ܣ��������˲��ĵĴ���
    Set GetArchiveOneFrom = mobjOneArchiveView
End Function

Public Function GetArchiveFrom() As Object
'���ܣ����õ��Ӳ������Ĺ��ܣ�Ƕ��ʽ��ȡ�������
    Set mobjArchiveView = New frmArchiveView
    mobjArchiveView.BorderStyle = FormBorderStyleConstants.vbBSNone '����Ϊ�ޱ߿�
    mobjArchiveView.Caption = mobjArchiveView.Caption       '�ص�����һ��
    Set GetArchiveFrom = mobjArchiveView
End Function

Public Sub ShowArchiveViewAndPrint(ByVal frmParent As Object, ByVal strPrivs As String)
'���ܣ����ò������Ĵ�ӡ����
    If mobjArchiveViewAndPrint Is Nothing Then Set mobjArchiveViewAndPrint = New frmArchiveViewAndPrint
    mobjArchiveViewAndPrint.ShowArchive frmParent, strPrivs
End Sub

Public Function GetArchiveViewAndPrintFrom() As Object
'���ܣ���ȡ�������Ĵ�ӡ�Ĵ���
    If mobjArchiveViewAndPrint Is Nothing Then Set mobjArchiveViewAndPrint = New frmArchiveViewAndPrint
    Set GetArchiveViewAndPrintFrom = mobjArchiveViewAndPrint
End Function

Public Sub zlArchiveRefresh(ByVal lng����ID As Long, ByVal lng����ID As Long)
'���ܣ�ˢ�µ��Ӳ������Ľ���
    Call mobjArchiveView.zlRefresh(lng����ID, lng����ID)
End Sub

Public Sub ViewEPRReport(frmParent As Object, ByVal lng����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(frmParent, lng����ID, CanPrint)
End Sub

Public Sub ViewPACSImage(frmParent As Object, ByVal lngҽ��ID As Long, ByVal blnMoved As Boolean)
'���ܣ�PACS��Ƭ����
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(lngҽ��ID, frmParent, blnMoved)
    End If
End Sub

Public Sub InitCISJob(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal strMainPrivs As String, Optional ByRef objEmr As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    
    glngSys = lngSys

    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    gstrPrivs = strMainPrivs
    If Not objEmr Is Nothing Then
        Set gobjEmr = objEmr
    End If
    '��ʼ������Ա��Ϣ
    Call GetUserInfo
    If gstrDBUser = "" Then gstrDBUser = UserInfo.�û���   '�������Ĵ�ӡ
    '��ʼ��ϵͳ����:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call InitSysPar
    '��ʼ�����Ĳ���
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs, gobjEmr)
End Sub

Public Function InitCISJobPar() As Boolean
    InitCISJobPar = InitSysPar
End Function