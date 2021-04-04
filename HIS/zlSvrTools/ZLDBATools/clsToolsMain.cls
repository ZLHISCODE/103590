VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function InitConnection(ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection) As Boolean
    '���ܣ� ��ʼ�����ݿ����Ӷ���,�ɹ��򷵻�True
    '������-IsDba���Ƿ���DBA�û�    -cnConnection�����ݿ����Ӷ���
    Dim blnResult As Boolean
    
    '��DBA�����Ӷ���Ϊ��
    If Not blnIsDba Or cnConnection Is Nothing Then
        blnResult = False
    Else
        blnResult = True
    End If
    
    InitConnection = blnResult
End Function


Private Function InitParameters(ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection, ByVal strUserName As String, _
                                                ByVal strPwd As String) As Boolean
    '����:��ʼ�����ȫ�ֱ�����
    '������-IsDba���Ƿ���DBA�û�    -cnConnection�����ݿ����Ӷ���   -strUserName���û���    -strPwd������
    Dim strFileName As String, strPName As String * 256
    
    On Error GoTo errH
    
    '���Ƚ���һ���򵥵��жϣ��������ZL��ͷ�ĳ�����ߵ��Գ�����ôֱ���˳�
    Call GetModuleFileName(0, strPName, 256)
    strFileName = Left(strPName, InStr(strPName, Chr(0)) - 1)
    strFileName = UCase(Mid(strFileName, InStrRev(strFileName, "\") + 1))

    If strFileName <> "VB6.EXE" Then
        If (Not UCase(strFileName) Like "ZL*") Or GetSetting("ZLSOFT", "����ȫ��", "ִ���ļ�") = "" Then
            InitParameters = False
            Exit Function
        End If
    End If

    '������ݿ����ӳ�ʼ��ʧ�ܻ��DBA�û�,����False
    If Not InitConnection(blnIsDba, cnConnection) Then
        InitParameters = False
        Exit Function
    End If
    
    If gcnOracle = cnConnection Then
        gcnOracle.CursorLocation = adUseClient
        InitParameters = True
        Exit Function
    Else
        gcnOracle.CursorLocation = adUseClient
        Set gcnOracle = cnConnection
    End If
    
    
    '��ʼ������
    gblnRAC = CheckRAC(gintInstId)
    gintCpuCount = GetCpuCount(gintCpuAdvise, gintCpuMax) '��ȡ���ж�
    gblnIsZlhis = CheckTblExist("ZLSYSTEMS")
    gstrBigVer = GetOracleVersion
    gstrVerNum = GetOracleVersion(True)
    gstrUserName = strUserName
    gstrPassword = strPwd
    gblnHasBigtables = CheckTblExist("ZLBIGTABLES")
    gblnHasZltables = CheckTblExist("ZLTABLES")
    Call CreateStr2list
    Call CreateList2str
    
    InitParameters = True
    Exit Function
errH:
    ErrCenter
    InitParameters = False
End Function

Public Function GetFrmByMdoudle(ByVal strMoudleNum As String, ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection, ByVal strUserName As String, _
                                                ByVal strPwd As String) As Object
    Dim objForm As Object
                                                
    If Not InitParameters(blnIsDba, cnConnection, strUserName, strPwd) Then
        Set objForm = Nothing
        Exit Function
    End If
    
    Select Case strMoudleNum
        Case "0601"   '���ܼ��
            Set objForm = New frmMonitorMain
        Case "0602"  'SQL�����Ż�
            Set objForm = New frmTunning
        Case "0604"     '�Ự����
            Set objForm = New frmKillBlockers
        Case "0605"     '�������
            Set objForm = New frmIdxInfo
        Case "0606"     '�ռ�����
            Set objForm = New frmReused
    End Select
    
    Set GetFrmByMdoudle = objForm
End Function