VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterFacePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mobjPubPatient As Object
'��������ֵ�ľֲ�����
Private mvarblnInit As Boolean '�ֲ�����
Private mvarblnRealName As Boolean '�ֲ�����

Public Property Get blnRealName() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnRealName
    blnRealName = mvarblnRealName
End Property

Public Property Get blnInit() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnInit
    blnInit = mvarblnInit
End Property


Public Function Init(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
                    ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
    '��ʼ����ģ�����
    On Error GoTo errHandler
    If CreatePublicPatient(lngSys, strDBUser, cnOracle) = False Then Exit Function
    mvarblnInit = True
    mvarblnRealName = mobjPubPatient.blnRealName
    Init = True
    Exit Function
errHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsRealNamed(ByVal lng����id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鲡���Ƿ���ʵ������֤
    '��Σ�lng����ID ����id
    '���أ�True/False  ��ʵ������֤����True�����򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    IsRealNamed = mobjPubPatient.IsRealNamed(lng����id)
End Function

Public Function CheckPatiIdCard(ByVal strIdCard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����֤����Ϸ���У��
    '��Σ�strIdCard ����֤����
    '���Σ�strBirthday  ��������TrueΪ��������
    '         strSex ��������TrueΪ�Ա�
    '         strErrInfo ��������FalseΪ������Ϣ
    '         datCalc �������� ȱʡ��ϵͳʱ�����
    '���أ�True/False  ����֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)�����򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckPatiIdCard = mobjPubPatient.CheckPatiIdCard(strIdCard, strBirthday, strAge, strSex, strErrInfo, datCalc)
End Function

Public Function CheckPatiAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", _
    Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '   strAge:��������
    '   strBirthDay:��������
    '   bytTag:����zl_Age_Check�������ص�ѯ�����͵���Ϣ���Ƿ�Ҫǿ����ֹ�����Ǳ���ѯ��.0-����ѯ��,1-��ֹ
    '   strCalcDate:��������,ȱʡ���յ�ǰʱ�����
    '���أ�TRUE��FALSE��TRUE:����,FALSE:��ֹ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckPatiAge = mobjPubPatient.CheckPatiAge(strAge, strBirthday, bytTag, strCalcDate)
End Function

Public Function RecalcBirthDay(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ��������ȡ���˳�������
    '���:strAge:��������,�磺23�ꡢ1��2��
    'strCalcDate-�����������
    '����:����Ĳ��������ʽ��ȷ����㷵�س�������,���򷵻ؿ�
    '  strMsg-���ؾ�����Ϣ
    '��ȷ�����ʽ:X��[X��]��X��[X��]��X�졢XСʱ[X����]
    '    X��:X���ܴ���200,X��:X���ܴ���12,X��:X���ܴ���31,XСʱ:X���ܴ���24,X����:X���ܴ���59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    RecalcBirthDay = mobjPubPatient.RecalcBirthDay(strAge, strDateOfBirth, strCalcDate, strMsg)
End Function

Public Function SavePatiBaseInfo(ByVal lng����id As Long, ByVal lng����ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal strģ�� As String, Optional ByVal int���� As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean = True, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������˻�����Ϣ(��ҵ�����ݵ�ͬ������)
    '��Σ�lng����ID-����ID (����Ϊ��/0)
    '         lng����ID-�Һ�ID����ҳID(��Ϊ0)
    '         strName-���� (����Ϊ��)
    '         strSex-�Ա� (����Ϊ��)
    '         strAge-���� (����Ϊ��)
    '         strBirthDay-�������� (����Ϊ��)
    '         strģ��-���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '         int���� 1-����;2-סԺ(lng����ID=0,��Ĭ��Ϊ1;lng����ID<>0,1-lng����IDΪ�Һ�ID,2-lng����IDΪ��ҳID)
    '         blnXWHIS-������Ϣ����ʱ�Ƿ����RIS�Ľӿ� ȱʡ=True���ã��ò������ڱ��ⲡ����Ϣ���ظ�����RIS�ӿڣ�
    '         blnEMPI-T EMPIƽ̨�Ѿ�������F-EMPIƽ̨δ����
    ' ���Σ�strInfo:���³ɹ�-��Ϣ�������µı仯��Ϣ(����True); ����ʧ��-��Ϣ����δ�ɹ���ԭ��
    ' ���أ�TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    SavePatiBaseInfo = mobjPubPatient.SavePatiBaseInfo(lng����id, lng����ID, strName, strSex, strAge, strBirthday, strģ��, int����, strInfo, blnXWHIS, blnEMPI)
End Function

Public Function PatiImageGatherer(frmMain As Object, ByRef strPictureFilePath As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͼƬ�ɼ�
    '���:frmMain-����������
    '����:strPictureFilePath-���زɼ�ͼƬ��·��(�����ļ�������׺)
    '����:TRUE��False
    '����:������
    '����:2014-07-16
    '����ʾ��:�˽ӿڷ���TRue����ֱ��ʹ��LoadPicture(strPictureFilePath)����ͼƬ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    PatiImageGatherer = mobjPubPatient.PatiImageGatherer(frmMain, strPictureFilePath)
End Function

Private Function CreatePublicPatient(ByVal lngSys As Long, ByVal strDBUser As String, _
                    ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(cnOracle, lngSys, strDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Sub Class_Terminate()
    mvarblnInit = False
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
End Sub

Public Function CheckIsRealNameColumn(ByVal lng����id As Long, ByVal strInputItem As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ŀ�жϣ��Ƿ�ʵ����֤��
    '���:strInputItem-��ǰ������Ŀ
    '����:��ʵ����֤����true,���򷵻�False
    '����:���˺�
    '����:2018-12-26 16:53:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckIsRealNameColumn = mobjPubPatient.CheckIsRealNameColumn(lng����id, strInputItem)
End Function


