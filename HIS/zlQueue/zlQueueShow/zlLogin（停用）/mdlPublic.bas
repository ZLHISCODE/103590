Attribute VB_Name = "mdlPublic"
Option Explicit

Public gobjFile As New FileSystemObject
Public SplashObj As New frmSplash

Public gstrSysName As String
Public gstrUserName As String

Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'��ȡ��������
Public Function getEncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(20)
    strRandom = GetRandom(20)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

'��ȡ��������
Public Function getDecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    getDecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
End Function

Private Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, StrBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        StrBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(StrBit = "0", "W", StrBit = "1", "I", StrBit = "2", "N", StrBit = "3", "T", StrBit = "4", "E", StrBit = "5", "R", StrBit = "6", "P", StrBit = "7", "L", StrBit = "8", "U", StrBit = "9", "M", _
                   StrBit = "A", "H", StrBit = "B", "T", StrBit = "C", "I", StrBit = "D", "O", StrBit = "E", "K", StrBit = "F", "V", StrBit = "G", "A", StrBit = "H", "N", StrBit = "I", "F", StrBit = "J", "J", _
                   StrBit = "K", "B", StrBit = "L", "U", StrBit = "M", "Y", StrBit = "N", "G", StrBit = "O", "P", StrBit = "P", "W", StrBit = "Q", "R", StrBit = "R", "M", StrBit = "S", "E", StrBit = "T", "S", _
                   StrBit = "U", "T", StrBit = "V", "Q", StrBit = "W", "L", StrBit = "X", "Z", StrBit = "Y", "C", StrBit = "Z", "X", True, StrBit)
        Case 2
            strNew = strNew & _
                Switch(StrBit = "0", "7", StrBit = "1", "M", StrBit = "2", "3", StrBit = "3", "A", StrBit = "4", "N", StrBit = "5", "F", StrBit = "6", "O", StrBit = "7", "4", StrBit = "8", "K", StrBit = "9", "Y", _
                   StrBit = "A", "6", StrBit = "B", "J", StrBit = "C", "H", StrBit = "D", "9", StrBit = "E", "G", StrBit = "F", "E", StrBit = "G", "Q", StrBit = "H", "1", StrBit = "I", "T", StrBit = "J", "C", _
                   StrBit = "K", "U", StrBit = "L", "P", StrBit = "M", "B", StrBit = "N", "Z", StrBit = "O", "0", StrBit = "P", "V", StrBit = "Q", "I", StrBit = "R", "W", StrBit = "S", "X", StrBit = "T", "L", _
                   StrBit = "U", "5", StrBit = "V", "R", StrBit = "W", "D", StrBit = "X", "2", StrBit = "Y", "S", StrBit = "Z", "8", True, StrBit)
        Case 0
            strNew = strNew & _
                Switch(StrBit = "0", "6", StrBit = "1", "J", StrBit = "2", "H", StrBit = "3", "9", StrBit = "4", "G", StrBit = "5", "E", StrBit = "6", "Q", StrBit = "7", "1", StrBit = "8", "X", StrBit = "9", "L", _
                   StrBit = "A", "S", StrBit = "B", "8", StrBit = "C", "5", StrBit = "D", "R", StrBit = "E", "7", StrBit = "F", "M", StrBit = "G", "3", StrBit = "H", "A", StrBit = "I", "N", StrBit = "J", "F", _
                   StrBit = "K", "O", StrBit = "L", "4", StrBit = "M", "K", StrBit = "N", "Y", StrBit = "O", "D", StrBit = "P", "2", StrBit = "Q", "T", StrBit = "R", "C", StrBit = "S", "U", StrBit = "T", "P", _
                   StrBit = "U", "B", StrBit = "V", "Z", StrBit = "W", "0", StrBit = "X", "V", StrBit = "Y", "I", StrBit = "Z", "W", True, StrBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    
    On Error Resume Next
    
    strUserPwd = TranPasswd(strUserPwd)
    
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName & ";Persist Security Info=false;", strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If

            Set OraDataOpen = Nothing
            Exit Function
        End If
    End With

    Err = 0
    On Error GoTo errHand

    'gstrDbUser = UCase(strUserName)
    'gobjComLib.SetDbUser gstrDbUser
    
    Set OraDataOpen = cnOracle
    Exit Function

errHand:
    MsgBox strError, vbInformation, gstrSysName
    Set OraDataOpen = Nothing
    Err = 0
End Function

