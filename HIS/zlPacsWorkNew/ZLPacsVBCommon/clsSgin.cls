VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsESign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const mstrSysName = "ǩ��"

Private mobjESign As Object
Private mstrServerName As String

Public IsConnection As Boolean

'�û���Ϣ
Private mstrUserName As String
Private mstrUserPwd As String

'���ݿ�ʵ����
Public Property Get ServerName() As String
    ServerName = mstrServerName
End Property

Public Property Let ServerName(ByVal vNewValue As String)
    mstrServerName = vNewValue
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ָ���û����û���
Public Property Get UserName() As String
    UserName = mstrUserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    mstrUserName = vNewValue
End Property

'ָ���û�������
Public Property Get UserPwd() As String
    UserPwd = mstrUserPwd
End Property

Public Property Let UserPwd(ByVal vNewValue As String)
    mstrUserPwd = vNewValue
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'��ʼ������ǩ������
Public Function Initialize(ByVal lngSys As Long) As Boolean
On Error GoTo ErrorHand

    If mobjESign Is Nothing Then
        mobjESign = CreateObject("zl9ESign.clsESign")

        If mobjESign Is Nothing Then
            MsgBox "����zl9ESign.dll����ʧ�ܣ�����ϵ����Ա!", vbExclamation, mstrSysName
            Exit Function
        End If
    End If

    '��ʼ������ǩ������
    If Not mobjESign Is Nothing Then
        If mobjESign.Initialize(gcnOracle, lngSys) = False Then
            MsgBox "����֤���ʼ��ʧ�ܣ���ʹ����ȷ������֤��ǩ����", vbInformation + vbOKOnly, mstrSysName
            Exit Function
        End If
    End If

    Initialize = True
    Exit Function
ErrorHand:
    Initialize = False
    Err = 0
End Function

'�������֤�����½�û��Ƿ�һ��
Public Function CheckCertificate(ByVal strDbUser As String) As Boolean
    CheckCertificate = False

On Error GoTo ErrorHand

    If Not mobjESign Is Nothing Then
        CheckCertificate = mobjESign.CheckCertificate(strDbUser)
    End If
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'���֤���Ƿ�ͣ��
Public Function CertificateStoped(ByVal strDBName As String) As Boolean
    CertificateStoped = False

On Error GoTo ErrorHand

    If Not mobjESign Is Nothing Then
        CertificateStoped = mobjESign.CertificateStoped(strDBName)
    End If
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'ʹ������ǩ����Դ�Ľ���ǩ������
'���أ�ǩ����ϢstrSign-���ܺ��Դ�ģ�lngCertID-ǩ��ʹ�õ�֤���¼ID��strʱ��� --ǩ��֮���ʱ���
Public Function Signature(ByVal strSource As String, _
                          ByVal strDbUser As String, _
                          lngCertID As Long, _
                          Optional strTimeStamp As String, _
                          Optional objSignPic As StdPicture, _
                          Optional strTimeStampCode As String) As String
    Dim strSign As String
    
On Error GoTo ErrorHand
    
    If Not mobjESign Is Nothing Then
        strSign = mobjESign.Signature(strSource, strDbUser, lngCertID, strTimeStamp, objSignPic, strTimeStampCode)
        If strSign = "" Then MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, mstrSysName
    End If

    Signature = strSign
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'��֤ǩ��,����Ҫ����USB-Key
Public Function VerifySignature(ByVal strSource As String, ByVal lngSignID As Long, ByVal intSignType As Integer) As Boolean
    VerifySignature = False
    
On Error GoTo ErrorHand
    
    If Not mobjESign Is Nothing Then
        VerifySignature = mobjESign.VerifySignature(strSource, lngSignID, intSignType)
    End If
    
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

Private Sub Class_Initialize()
    mstrServerName = ""
    mstrUserName = ""
    mstrUserPwd = ""
    IsConnection = False
End Sub

Private Sub Class_Terminate()
    If Not mobjESign Is Nothing Then mobjESign = Nothing
End Sub

Public Function OraDataOpen() As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    Dim strUserPwd As String

    On Error GoTo ErrorHand
    
    IsConnection = False
    OraDataOpen = False
    
    'ת������
    If Not (UCase(mstrUserName) = "SYS" Or UCase(mstrUserName) = "SYSTEM") Then
        strUserPwd = TranPasswd(mstrUserPwd)
    Else
        strUserPwd = mstrUserPwd
    End If

    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close

        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & ServerName, UserName, strUserPwd

        If Err <> 0 Then
            MsgBox "�ʻ���֤ʧ�ܣ��������ʻ��������벻��ȷ", vbExclamation, mstrSysName
            Exit Function
        End If
    End With
    
    IsConnection = True
    OraDataOpen = True
    Exit Function

ErrorHand:
    MsgBox "�ʻ���֤ʧ�ܣ��������ʻ��������벻��ȷ", vbExclamation, mstrSysName
    Err = 0
End Function