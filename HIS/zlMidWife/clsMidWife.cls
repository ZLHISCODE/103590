VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMidWife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjFrm As frmBrowser

Public Function GetForm() As Object
'���ܣ�����������Ϣ�Ĵ������
    Set GetForm = mobjFrm
End Function

Public Sub RefreshForm(ByVal lngPatiID As Long, ByVal lngPageID As Long)
'���ܣ�ˢ��ָ�����˵�������Ϣ����
'������lngPatiID-����id,lngPageID-��ҳID
    Call mobjFrm.RefreshForm(lngPatiID, lngPageID)
End Sub

Public Function Initialize(ByRef cnOracle As ADODB.Connection, Optional ByVal lngPid As Long, Optional ByVal lngTid As Long) As Boolean
'���ܣ��ӿڳ�ʼ��
'������cnOracle=��¼ZLHIS����̨�Ļ����
    Dim rsTemp As ADODB.Recordset, strSql As String, strTemp As String
    Dim varPar As Variant, strIP As String, strPubUserName As String, strPwd As String, strTime As String
    
    On Error GoTo errHand
    
    If gstrURLLogin = "" Then
        Set gcnOracle = cnOracle
        
        strSql = "Select user From dual"
        Set rsTemp = gcnOracle.Execute(strSql)
                
        If rsTemp.RecordCount > 0 Then gstrUser = rsTemp!user
        'gstrUser = "ZLHIS"
            
        strTemp = GetSysPar(325, 100)
        'varPar = Split("192.168.0.60:8080/ords+127", "+")
        varPar = Split(strTemp, "+")
        If UBound(varPar) < 1 Or varPar(0) = "" Then
            MsgBox "û�ж�����ʿϵͳ���г�ʼ�����ã�����ϵͳ������������������ʿ����վ����ز�����", vbInformation, "������Ϣ"
            Exit Function
        End If
        strIP = "http://" & varPar(0) & "/f?p=" & varPar(1)
        
        strPubUserName = "zlsol_pub"
        strPwd = Replace(strTemp, varPar(0) & "+" & varPar(1) & "+", "")
        strPwd = zlStr.Sm4DecryptEcb(strPwd)
        strPwd = GetEncrypt(strPwd)
        strTime = GetEncrypt(Replace(Now, " ", "-"))
        
        gstrURL = strIP & ":14:[SESSION_TOKEN]::NO:RP:P14_PID,P14_TID:"
        gstrURLLogin = strIP & ":LOGIN_DESKTOP" & "&username=" & strPubUserName & "&pwd=" & strPwd & "&key=" & strTime & "&pid=" & lngPid & "&tid=" & lngTid & "&hisname=" & gstrUser
        glngPatiID = lngPid
        glngPageID = lngTid
    End If
    
    mobjFrm.InitLoad
    
    Initialize = True
    Exit Function
    
errHand:
    MsgBox Err.Description, vbExclamation, "������Ϣ��ʼ��"
    Initialize = False
End Function

Private Sub Class_Initialize()
    '��ʹ���Instancing��������Ϊ5-MultiUse(���������һ�����ʵ��),������3-SingleUse��ÿ�������Ӧһ�����ʵ������Ҳ�����
    '�൱���Ƕ���ĳ�ʼ��������һ�������ʵ����ʼ��
    glngInstance = glngInstance + 1
    Set mobjFrm = New frmBrowser    '֧��ҽ��վ����ʿվ�ֱ�ͬʱ��������
End Sub

Private Sub Class_Terminate()
    glngInstance = glngInstance - 1
    Set mobjFrm = Nothing
    
    If glngInstance = 0 Then
        Set gcnOracle = Nothing
    End If
End Sub