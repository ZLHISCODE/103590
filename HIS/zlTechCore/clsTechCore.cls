VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTechCore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub InitTechCore(cnMain As ADODB.Connection)
'------------------------------------------------
'���ܣ� �����ĳ�ʼ������������ı�Ҫ���ݴ�������
'������
'   cnMain:����������ݿ�����
'���أ�
'------------------------------------------------
    Set gcnOracle = cnMain
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    glngSys = 100
    Call GetUserInfo
    
    gstrUnitName = GetUnitName
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If Not gblnPass Then
        gblnPass = GetSysParVal(30) = "1"
        If gblnPass Then
            If Not PassInitialize() Then gblnPass = False
        End If
    End If
End Sub

Private Sub Class_Terminate()
    'PASS�ӿ���ֹ
    If gblnPass Then Call PassQuit
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub


'�ⲿ�ӿڣ�ҽ���༭
'�����������Χϵͳ
'By:��ͮ��
Public Function Editҽ��(frmMain As Object, _
                            ByVal PatientID As Long, _
                            ByVal PageID As Variant, _
                            ByVal lngҽ��ID As Long, _
                            cnMain As ADODB.Connection, _
                            ByVal strDBUser As String, _
                            ByVal lngSys As Long, _
                            Optional ByVal strPrivs As String = "") As Boolean
    
    '---------------------------------------------------------------------------------------------------------
    '����:
    '---------------------------------------------------------------------------------------------------------
    If TypeName(PageID) = "String" Then
        If Not CheckAdviceWindow("����ҽ���༭") Then Exit Function
    Else
        If Not CheckAdviceWindow("סԺҽ���༭") Then Exit Function
    End If
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If TypeName(PageID) = "String" Then
        Editҽ�� = frmOutAdviceEdit.ShowMe(frmMain, strPrivs, PatientID, PageID, lngҽ��ID)
    Else
        Editҽ�� = frmInAdviceEdit.ShowMe(frmMain, strPrivs, PatientID, PageID, lngҽ��ID)
    End If
End Function

'�ⲿ�ӿڣ���ָ����������ʾҽ���б�
'�����������Χϵͳ
'By:��ͮ��
Public Function ListAdviceInObject(frmMain As Object, objContainer As Object, _
                            ByVal PatientID As Long, _
                            ByVal PageID As Variant, _
                            ByVal lng����ID As Long, ByVal lng����ID As Long, _
                            ByVal lngҽ��ID As Long, _
                            cnMain As ADODB.Connection, _
                            ByVal strDBUser As String, _
                            ByVal lngSys As Long, _
                            Optional ByVal strPrivs As String = "") As Object

    '---------------------------------------------------------------------------------------------------------
    '����:
    '---------------------------------------------------------------------------------------------------------
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If TypeName(PageID) = "String" Then
        Set ListAdviceInObject = New frmOutDoctorAdvice
    Else
        Set ListAdviceInObject = New frmInDoctorAdvice
    End If
    ListAdviceInObject.mstrPrivs = strPrivs
    Set ListAdviceInObject.mfrmParent = frmMain
    
    Load ListAdviceInObject
    
    lngStyle = GetWindowLong(ListAdviceInObject.Hwnd, GWL_STYLE)
    SetWindowLong ListAdviceInObject.Hwnd, GWL_STYLE, lngStyle Or WS_CHILD
    
    SetParent ListAdviceInObject.Hwnd, objContainer.Hwnd
    MoveWindow ListAdviceInObject.Hwnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1
            
    '���ù��ܴ�������ˢ��
    If TypeName(PageID) = "String" Then
        Call ListAdviceInObject.zlRefresh(PatientID, CStr(PageID), 1, 0)
    Else
        Call ListAdviceInObject.zlRefresh(PatientID, CLng(PageID), lng����ID, lng����ID, False)
    End If

    ListAdviceInObject.Show
End Function

Public Function ListChargeInObject(ByVal frmMain As Object, _
                                    ByVal objContainer As Object, _
                                    ByVal lngҽ��ID As Long, _
                                    ByVal lng����id As Long, _
                                    ByVal cnMain As ADODB.Connection, _
                                    ByVal strDBUser As String, _
                                    ByVal lngSys As Long, _
                                    Optional ByVal strPrivs As String = "", _
                                    Optional ByVal strClass As String = "����", _
                                    Optional ByVal strSys As String = "") As Object
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
        
    Set ListChargeInObject = New frmChargeRequest
    
    Call InitSysPar '��ʼ��ϵͳ����
    
    Load ListChargeInObject
    
    lngStyle = GetWindowLong(ListChargeInObject.Hwnd, GWL_STYLE)
    
    Call SetWindowLong(ListChargeInObject.Hwnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(ListChargeInObject.Hwnd, objContainer.Hwnd)
    Call MoveWindow(ListChargeInObject.Hwnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
    '���ù��ܴ�������ˢ��
    
    Call ListChargeInObject.zlRefresh(frmMain, lngҽ��ID, lng����id, strPrivs, strClass, strSys)
    
    ListChargeInObject.Show
End Function

Public Function InDoctorAdvice() As Object
    Set InDoctorAdvice = New frmInDoctorAdvice
End Function

Public Function OutDoctorAdvice() As Object
    Set OutDoctorAdvice = New frmOutDoctorAdvice
End Function
'�༭����
Public Function EditExpense(objParent As Object, ByVal mbytInState As Integer, ByVal int��¼���� As Integer, ByVal mstrPrivs As String, ByVal strNO As String, _
    ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal int������Դ As Integer, ByVal lng��������ID As Long, ByVal lng���˿���ID As Long, Optional ByVal mbln���õǼ� As Boolean = False) As Boolean
    
    Dim bln��� As Boolean
    
    Call InitSysPar '��ʼ��ϵͳ����
    
    If int��¼���� = 2 And mbytInState = 0 And Len(strNO) > 0 Then
        bln��� = BillisZeroLog(strNO)
    Else
        bln��� = mbln���õǼ�
    End If

    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = mbytInState
    frmTechnicExpense.mbln���õǼ� = bln���
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mlngҽ��ID = lngҽ��ID
    frmTechnicExpense.mlng���ͺ� = lng���ͺ�
    frmTechnicExpense.mlng����ID = lng����ID
    frmTechnicExpense.mlng��ҳID = lng��ҳID
    frmTechnicExpense.mint������Դ = int������Դ
    frmTechnicExpense.mint��¼���� = int��¼����
    frmTechnicExpense.mlng��������ID = lng��������ID
    frmTechnicExpense.mlng���˿���id = lng���˿���ID
    On Error Resume Next
    frmTechnicExpense.Show 1, objParent
    On Error GoTo 0
    EditExpense = gblnOK
End Function