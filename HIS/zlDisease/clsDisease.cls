VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Public Event PatiTransfer(ByVal lng����ID As Long, ByVal str�Һ�No As String)  'ת��
Public Event Closed(ByVal lngFunID As Long, ByVal strTag As String)  'lngFunID������ID��0����ʾ ���Խ����������д����Ϻ�رմ��壻strTag ��չ������
Private WithEvents frmDisRegist As frmDiseaseRegist
Attribute frmDisRegist.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    If frmDisRegist Is Nothing Then Set frmDisRegist = New frmDiseaseRegist
End Sub

Private Sub Class_Terminate()
     If Not frmDisRegist Is Nothing Then Set frmDisRegist = Nothing
End Sub

Public Sub InitDisease(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal lngModul As Long, ByVal strMainPrivs As String, Optional ByRef objMip As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    Call zlGetComLib
    glngSys = lngSys
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = gobjComlib.GetUnitName
    gstrPrivs = strMainPrivs
    If strMainPrivs = "" Then
        gstrPrivs = gobjComlib.GetPrivFunc(lngSys, lngModul)
    End If
    '��ʼ������Ա��Ϣ
    Call GetUserInfo
    '��ʼ��ϵͳ����:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call InitSysPar
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    End If
    gstrDBOwer = GetDbOwner(lngSys)
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
End Sub

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���
    On Error GoTo errHand

    Set gfrmMain = frmMain
    
    Set gcnOracle = cnMain
    Call zlGetComLib
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDbUser
    gstrDBOwer = GetDbOwner(lngSys)
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrPrivs = gobjComlib.GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub

    gstrUnitName = gobjComlib.GetUnitName
    gblnShowInTaskBar = blnShowInTaskBar
    
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    End If
    Call InitSysPar
    On Error Resume Next
    Select Case lngModul
        Case 1277 '������������
            frmDiseaseReportSetting.Show , frmMain
        Case 1278 '��Ⱦ������վ
            frmDiseaseStation.Show , frmMain
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���
    
    Set gfrmMain = Nothing
    
    Set gcnOracle = cnMain
    Call zlGetComLib
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDbUser
    gstrDBOwer = GetDbOwner(lngSys)
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = gobjComlib.GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    
    gstrUnitName = gobjComlib.GetUnitName
    gblnShowInTaskBar = True
    
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
    End If
    Call InitSysPar
    On Error Resume Next
    Select Case lngModul
        Case 1277 '������������
            gobjComlib.ZLCommFun.ShowChildWindow frmDiseaseReportSetting.hwnd, lngMain, True
        Case 1278 '��Ⱦ������վ
            gobjComlib.ZLCommFun.ShowChildWindow frmDiseaseStation.hwnd, lngMain, True
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    blnChildren = gObjRichEPR.CloseWindows()
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function

Private Sub frmDisRegist_Closed(ByVal lngFunID As Long, ByVal strTag As String)
    RaiseEvent Closed(lngFunID, strTag)
End Sub

Private Sub frmDisRegist_PatiTransfer(ByVal lng����ID As Long, ByVal str�Һ�No As String)
     RaiseEvent PatiTransfer(lng����ID, str�Һ�No)
End Sub

Public Sub HideFrm(ByVal lngFrmID As Long)
'���ܣ����ش���
'������lngFrmID,�����ʶ��0����ʾ frmDisRegist
    On Error Resume Next
    frmDisRegist.Hide
    Err.Clear
End Sub

Public Function ShowDisRegist(ByRef frmParent As Object, ByVal intType As Integer, Optional ByVal lngID As Long, _
                Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һ�No As String, _
                Optional ByVal lngҽ��id As Long, Optional ByVal var�Ǽǿ��� As Variant, Optional ByVal dat�ͼ�ʱ�� As Date, Optional ByVal var�ͼ���� As Variant, _
                Optional ByVal str�ͼ�ҽ�� As String, Optional ByVal str�걾���� As String, Optional ByVal str������� As String, _
                Optional ByVal dat���ʱ�� As Date, Optional ByVal str���ƴ�Ⱦ�� As String, Optional ByRef objMip As Object, Optional ByVal str�Ǽ��� As String) As Integer
'���ܣ����ô�Ⱦ�����Խ��������
'������ intType 0��ʾ��д��ֻ��ʾ�ϰ벿�֣���1-��ʾҽ��������ֻ���°벿�ֿɱ༭����2-�鿴�����в��ɱ༭���ɲ鿴�°벿�֣�
'       lngID  = �������Լ�¼ ID
'       lng����ID = ����ID
'       lng��ҳID=סԺ:��ҳID
'       str�Һ�No =����Һŵ�NO
'       lngҽ��ID =Ҫ��д��������ҽ����ҽ��ID
'       var�Ǽǿ��� = ��ǰ��д�������Ŀ���
'       dat�ͼ�ʱ�� = �걾�����ʱ��
'       var�ͼ���� = �ò���ǰ�����Ŀ���
'       str�ͼ�ҽ�� = �ò���ǰ������ҽ��
'       str�걾���� = ���˼�����õı걾
'       str������� = ��ǰ��д�ķ������
'       dat���ʱ�� = ���˵ļ��ʱ��
'       str���ƴ�Ⱦ��  ���˵����ƴ�Ⱦ��
'       objMip  ��Ϣ����
'      ���intTypeΪ0�Ļ���lng����ID,lng��ҳID,str�Һ�No����
'      ���intTypeΪ1����2�Ļ���lngID����
'����ֵ��1��ֻ��intType Ϊ0��ʾ��дʱ�Ż���֣���ʾ�Ѿ���д�÷��������ͳɹ���
'        2��ֻ��intType Ϊ1��ʾҽ������ʱ�Ż���֣���ʾ�Ѿ���д����˵�����������Ѿ�������ɣ�
'        3��ֻ��intType Ϊ1��ʾҽ������ʱ�Ż���֣���ʾ��������������������ת�ƴ�����
'        �����������ֵΪ0

    ShowDisRegist = frmDisRegist.ShowDiseaseRegist(frmParent, intType, lngID, lng����ID, lng��ҳID, str�Һ�No, lngҽ��id, var�Ǽǿ���, dat�ͼ�ʱ��, var�ͼ����, str�ͼ�ҽ��, str�걾����, str�������, dat���ʱ��, str���ƴ�Ⱦ��, objMip, str�Ǽ���)
End Function

Public Function ShowDisQuery(ByVal lng�Ǽǿ���ID As Variant) As Long
'���ܣ����ô�Ⱦ�����Խ����ѯ����
'������lng�Ǽǿ���ID=�������Լ�¼ �� �Ǽǿ���ID
'����ֵ�����ز�ѯ���ļ�¼����
    ShowDisQuery = frmDiseaseQuery.ShowDiseaseQuery(lng�Ǽǿ���ID)
End Function

Public Function ShowPatiDis(ByVal rsTemp As ADODB.Recordset, ByRef frmParent As Object) As Long
'���ܣ����ô�Ⱦ�����Խ����ѯ����
'������rsTemp-���˵����Է�������¼��
'����ֵ������ѡ�еķ�����ID
    ShowPatiDis = frmDiseaseQuery.ShowPatiDis(rsTemp, frmParent)
End Function

Public Function ShowRegistByPati(ByRef frmParent As Object, ByVal intType As Integer, ByVal lng����ID As Long, _
                            Optional ByVal lng��ҳID As Long, Optional ByVal str�Һ�No As String, Optional ByVal var�Ǽǿ��� As Variant) As Boolean
'���ܣ����ݴ���Ĳ���ID����ҳID������Ļ��ǹҺŵ�NO������ѯ�ò��˵����Խ��������
'������intType��1=ҽ��������2-�鿴��3-�޸�
'      lng����ID = ����ID
'      lng��ҳID=סԺ:��ҳID
'      str�Һ�No =����Һŵ�NO
'      var�Ǽǿ��� = ��д�������Ŀ���
    ShowRegistByPati = QueryRegistByPati(frmParent, intType, lng����ID, lng��ҳID, str�Һ�No, var�Ǽǿ���)
End Function

Public Function ModifyDiseaseDoc(ByVal frmPar As Object, ByVal lngFileID As Long, _
            ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long) As Boolean
'���ܣ� �޸Ĵ�Ⱦ������
'������
'   lngFileID:����ID
'   lngPatiID:����ID
'   lngPageId:������ҳID���߹Һ�ID
'   lngPatFrom:������Դ��1-���2-סԺ
'   lngDeptId:���˿���ID
    Dim objDoc As cEPRDocument
    Set objDoc = New cEPRDocument
    Call objDoc.InitEPRDoc(cprEM_�޸�, cprET_�������༭, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
    Call objDoc.ShowEPREditor(frmPar, InStr(1, gstrPrivs, "������ӡ") > 0, 1)
    Set objDoc = Nothing
    ModifyDiseaseDoc = True
End Function

Public Function EditNotFillReason(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long) As Boolean
'���ܣ��༭���治��д���濨��ԭ��
    Dim rsDisease As ADODB.Recordset
    Dim strTmp As String
    Dim strID As String
    If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDisease) Then
        If rsDisease.RecordCount = 1 Then
            strTmp = frmReason.ShowMe(frmPar, NVL(rsDisease!�������˵��))
            Call SaveReason(strTmp, rsDisease!ID, rsDisease!��¼״̬)
        ElseIf rsDisease.RecordCount > 1 Then
            If frmFeedBackList.ShowMe(frmPar, rsDisease, strID, 2) Then
            rsDisease.Filter = "ID = " & strID
            If rsDisease.RecordCount > 0 Then
                strTmp = frmReason.ShowMe(frmPar, NVL(rsDisease!�������˵��))
                Call SaveReason(strTmp, rsDisease!ID, rsDisease!��¼״̬)
            End If
            End If
        End If
    End If
End Function

Public Function ShowDiseaseStation(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                                   ByVal intPatiFrom As Integer, ByVal lng����ID As Long, ByVal str����ID As String, ByVal str���ID As String, Optional ByRef blnNotView As Boolean) As Boolean
'���ܣ���ѯָ����Աһ�����Ƿ���д����Ⱦ�����濨
'������lng����ID    ����ID
'      lng��ҳID    סԺΪ ��ҳID������Ϊ �Һ�ID
'      intPatiFrom  ������Դ סԺΪ 2�� ����Ϊ 1
'      lng����ID    ���� ID
'      str����ID    ����ID
'      str���ID    ���ID
    ShowDiseaseStation = frmDiseaseStation.ShowDiseaseStation(frmParent, lng����ID, lng��ҳID, intPatiFrom, lng����ID, str����ID, str���ID, blnNotView)
End Function
