Attribute VB_Name = "mdlLCDShowPublic"
Option Explicit
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjPrintMode As Object
Public gobjReport As Object
'Public gcnOracle As New ADODB.Connection
Public gstrSysName As String                'ϵͳ����
Public gstrSQL As String

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Sub SaveErrLog(ByVal strInfo As String)
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    Dim rsTemp As New ADODB.Recordset
    
    strFile = "C:\zl9LCDShow_" & Format(gobjDatabase.Currentdate, "yyyyMMdd") & ".TXT"
    '����ļ��Ƿ���ڣ��������򴴽�
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strInfo & vbCrLf
    objText.Close
End Sub
'Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
'    '------------------------------------------------
'    '���ܣ� ��ָ�������ݿ�
'    '������
'    '   strServerName�������ַ���
'    '   strUserName���û���
'    '   strUserPwd������
'    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
'    '------------------------------------------------
'    Dim strError As String
'
'    On Error Resume Next
'    With cnOracle
'        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
'        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
'    End With
'    If err <> 0 Then
'        If blnMessage = True Then
'            '���������Ϣ
'            strError = err.Description
'            If InStr(strError, "�Զ�������") > 0 Then
'                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-12154") > 0 Then
'                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-12541") > 0 Then
'                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-01033") > 0 Then
'                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
'            Else
'                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
'            End If
'        End If
'
'        err.Clear
'        OraDataOpen = False
'        Exit Function
'    End If
'    OraDataOpen = True
'End Function
'Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnTemp As ADODB.Connection)
''���ܣ��򿪼�¼��
'    If rsTemp.State = adStateOpen Then rsTemp.Close
'
'    Call gobjComLib.SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), IIf(cnTemp Is Nothing, gcnOracle, cnTemp), adOpenStatic, adLockReadOnly
'    Call gobjComLib.SQLTest
'End Sub
Public Sub SaveDebug(ByVal strInfo As String)
    '�������=1����ʾ���Ե�����Ϣ,2-����ʽ��Ϣд���ı���������������������Ϣ
    '�ж��Ƿ��ǵ���״̬��������ʾ��ʾ��

    'д�ı��ļ�
    '��������Ϣд���ļ���
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String, strExchange As String
    Dim rsTemp As New ADODB.Recordset
'    If gint���� <> 0 Then
        If strExchange = "" Then strExchange = "C:\ҩ���Ŷӽк�"
        strFile = strExchange & "\������Ϣ_" & Format(Now, "yyyyMMdd") & ".TXT"
        '����ļ����Ƿ���ڣ��������򴴽�
        If Not objFile.FolderExists(strExchange) Then objFile.CreateFolder strExchange
        '����ļ��Ƿ���ڣ��������򴴽�
        If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
            
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfo
        objText.Close
'    End If
End Sub
