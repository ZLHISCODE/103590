Attribute VB_Name = "mdlMain"
Option Explicit

'---------------------------------------------------------------
'˵�����������̡��߼�����ģ��
'���ƣ�������
'---------------------------------------------------------------

Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Sub Main()
    Dim blnNew As Boolean
    
    On Error Resume Next
    
    Set gobjFile = New FileSystemObject
    If Err.Number <> 0 Then
        MsgBox "������FileSystemObject������ʧ�ܣ�����������ֹ������ϵ����Ա��", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjRegister = Nothing
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err.Number <> 0 Then
        MsgBox "������zl9ComLib������ʧ�ܣ�����������ֹ������ϵ����Ա��", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If Err.Number <> 0 Then
'        MsgBox "������zlRegister������ʧ�ܣ�����������ֹ������ϵ����Ա��", vbInformation, GSTR_MSG
'        Exit Sub
        Set gobjRegister = Nothing
        Err.Clear
    End If
    
    Set gcnOracle = New ADODB.Connection
    If Err.Number <> 0 Then
        MsgBox "��Microsoft ADO�����δ��װ������ϵ����Ա��", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjXML = New clsXML
    If Err.Number <> 0 Then
        MsgBox "������clsXML����ʧ�ܣ� ����������ֹ������ϵ����Ա��", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjZLPrint = CreateObject("zl9PrintMode.zlPrintMethod")
    If Err.Number <> 0 Then
        Set gobjZLPrint = Nothing
        MsgBox "������zl9PrintMode������ʧ�ܣ���Ӱ���ӡ����ع��ܣ�", vbInformation, GSTR_MSG
    End If
    
    Set gobjEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        Set gobjEncrypt = Nothing
        MsgBox "������zlEncryptPub������ʧ�ܣ���Ӱ������Կ��ع��ܣ�", vbInformation, GSTR_MSG
    End If
    
    On Error GoTo 0
    
    frmLogin.Show vbModal
    
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gobjComLib.InitCommon gcnOracle
            If mdlMain.GetUserInfo(gstrUser) Then
                frmMain.Show
            End If
        End If
    End If
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'���ܣ���ע���
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע���ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע���ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Properties("Persist Security Info") = True
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_MSG
            Else
                MsgBox strError, vbInformation, GSTR_MSG
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Err.Clear
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

'Public Function GetUserInfo() As Boolean
''���ܣ���ȡ��½�û���Ϣ
''���أ�True�ɹ���Falseʧ��
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo hErr
'
'    UserInfo.���� = UserInfo.�û���
'    Set rsTmp = mdlMain.GetUserInfo
'    If Not rsTmp Is Nothing Then
'        If Not rsTmp.EOF Then
'            UserInfo.ID = rsTmp!ID
'            UserInfo.��� = rsTmp!���
'            UserInfo.����ID = gobjComLib.zlCommFun.NVL(rsTmp!����ID, 0)
'            UserInfo.���� = gobjComLib.zlCommFun.NVL(rsTmp!����)
'            UserInfo.���� = gobjComLib.zlCommFun.NVL(rsTmp!����)
'            UserInfo.�û��� = rsTmp!�û���
'            GetUserInfo = True
'        End If
'        rsTmp.Close
'    End If
'
'    Exit Function
'
'hErr:
'    If gobjComLib.ErrCenter = 1 Then Resume
'End Function

Public Function GetUserInfo(ByVal strDBUser As String) As Boolean
'���ܣ���ȡ��ǰ�û��Ļ�����Ϣ
'���أ�����Ado��¼��
    Dim strSQL As String, strDefault As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strDefault = " And C.ȱʡ = 1"
    strSQL = "Select User,A.Id, A.���, A.����, A.����, A.רҵ����ְ��,B.�û���, C.����id, D.���� As ������, D.���� As ������ " & vbNewLine & _
             "From ��Ա�� A, �ϻ���Ա�� B, ������Ա C, ���ű� D " & vbNewLine & _
             "Where A.Id = B.��Աid And A.Id = C.��Աid And C.����id = D.Id And B.�û��� = [1] "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    If rsTemp.RecordCount = 0 Then
        strDefault = " And Rownum < 2"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    End If
    If rsTemp.RecordCount > 0 Then
        UserInfo.ID = rsTemp!ID
        UserInfo.��� = rsTemp!���
        UserInfo.����ID = gobjComLib.zlCommFun.NVL(rsTemp!����ID, 0)
        UserInfo.���� = gobjComLib.zlCommFun.NVL(rsTemp!����)
        UserInfo.���� = gobjComLib.zlCommFun.NVL(rsTemp!����)
        UserInfo.�û��� = rsTemp!�û���
        GetUserInfo = True
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'����:���������뷨����ر����뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ��������뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '�������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���true�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function

Public Sub LoadServer(ByVal cbxVar As ComboBox, ByRef colVar As Collection)
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim blnFinish As Boolean
    
    cbxVar.Clear
    
'    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
'    If Not gobjFile.FolderExists(strPath) Then '10G
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '10Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '10Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then    '10G ��ҵ��
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then    '10G ��ҵ��
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home2", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '11Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb11g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '11Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb11g_home2", "ORACLE_HOME")
'    End If
'    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
'    If Not gobjFile.FileExists(strFile) Then
'        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
'        If Not gobjFile.FileExists(strFile) Then Exit Sub
'    End If

    '����ע�����ȡOracle��װ·��
    strPath = GetRegItemValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", blnFinish)
    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
    If Not gobjFile.FileExists(strFile) Then
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
        If Not gobjFile.FileExists(strFile) Then Exit Sub
    End If

    Set colVar = New Collection
    
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '��ע���л����
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '���е����ݾ��Ƿ��������ˣ����������ݶ���ʼ��
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '���е�������������
                If InStr(strLine, "PROTOCOL = TCP") > 0 And strLine Like "*PORT = 152[0-9]*" Then
                    '�������ǵĳ���Ҫ��
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '���е�������ʵ����
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '�Ѿ��õ�������Ҫ������
                        colVar.Add Array(strServer, strComputer, strSID)
                        cbxVar.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
End Sub

Public Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'���ܣ���zlDrugMachine.cfg�ļ��л�ȡָ��������ֵ
'������
'  objXML��cfg�ļ������ݼ��غ��XML����
'  strName���������ƣ�����XML�������
'���أ�����ֵ

    Dim strValue As String

    If objXML Is Nothing Then Exit Function
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function

Public Function VerifyConfigFile(ByVal strFile As String) As Boolean
'���ܣ���������ĵ��Ƿ���ڣ������ھ��Զ�����
'������
'���أ�True���ɹ���False���ʧ��

    Dim fsoFile As New FileSystemObject
    Dim tsmFile As TextStream
    
    On Error GoTo hErr
    
    If fsoFile.FileExists(strFile) = False Then
        '���������ĵ�
        Set tsmFile = fsoFile.CreateTextFile(strFile)
        
        'Ĭ�������ĵ�����
        With tsmFile
            .WriteLine "<root>"
            .WriteLine "    <log>"
            .WriteLine "        <output>0</output>"
            .WriteLine "        <detailed>0</detailed>"
            .WriteLine "        <savedays>7</savedays>"
            .WriteLine "    </log>"
            .WriteLine "    <timer>"
            .WriteLine "        <enabled>0</enabled>"
            .WriteLine "        <businessdata></businessdata>"
            .WriteLine "        <cycle>5</cycle>"
            .WriteLine "        <validdays>2</validdays>"
            .WriteLine "        <viewlines>200</viewlines>"
            .WriteLine "    </timer>"
            .WriteLine "</root>"
        End With
        tsmFile.Close
    End If
    
    VerifyConfigFile = True
    Exit Function
    
hErr:
    Call gobjComLib.ErrCenter
End Function

Public Sub SetTextMaxLen(ByRef txtVal As TextBox, ByVal strTableField As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
'    gstrSQL = zlStr.FormatString("Select [2] as �ֶ� From [1] Where Rownum < 1 ", _
'                        CStr(Split(strTableField, ".")(0)), _
'                        CStr(Split(strTableField, ".")(1)))
    strSQL = "Select " & Split(strTableField, ".")(1) & " as �ֶ� From " & Split(strTableField, ".")(0) & " Where Rownum < 1 "
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ֶ���Ϣ")
    txtVal.MaxLength = rsTmp.Fields(0).DefinedSize
    rsTmp.Close

    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub CreateSOAP(ByRef objSOAP As Object)
    On Error Resume Next
    Set objSOAP = Nothing
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
        If Err.Number <> 0 Then
            MsgBox "ʵ������SoapClient��ʧ�ܣ�����ϵ������Ա��" & vbCrLf & _
                   "ע�⣺SoapClient��WinXP�°�װ2.0�汾��", _
                   vbInformation, GSTR_MSG
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub CreateHTTP(ByRef objHTTP As Object)
    On Error Resume Next
    Set objHTTP = Nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "ʵ������WinHttp��ʧ�ܣ�����ϵ������Ա��", vbInformation, GSTR_MSG
    End If
    On Error GoTo 0
End Sub

Public Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip/Pixel)
'���أ�blnTwip=True-����Twip��λ��False-�������ص�λ

    Dim vRect As RECT
    
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function FormatString(ByVal strFormat As String, ParamArray arrParams() As Variant) As String
'���ܣ���ʽ���ַ���
'������
'  strFormat�����ʽ��[1-x]Ϊ�����Źؼ��֣����ӣ�"����ֵΪ��[1]"
'  arrParams�����ʽ�Ĳ�������ӦstrFormat�еĲ����Źؼ���
'���أ���ʽ������ַ���

    Dim i As Integer, intSN As Integer
    Dim strKey As String, strTmp As String
    Dim blnStart As Boolean

    FormatString = strFormat

    If Len(strFormat) > 60000 Then Exit Function
    If Not strFormat Like "*[[]*[]]*" Then Exit Function
    If UBound(arrParams) < 0 Then Exit Function

    On Error GoTo errHandle

    For i = 1 To Len(strFormat)
        If Mid(strFormat, i, 1) = "[" Then
            blnStart = True
        End If
        If blnStart Then
            If Mid(strFormat, i, 1) = "]" Then
                intSN = Val(Mid(strKey, 2))
                If intSN > 0 Then
                    If UBound(arrParams) >= intSN - 1 Then
                        strTmp = strTmp & arrParams(intSN - 1)
                    End If
                Else
                    strTmp = strTmp & Mid(strKey, 2)
                End If
                blnStart = False
                strKey = ""
            Else
                strKey = strKey & Mid(strFormat, i, 1)
            End If
        Else
            strTmp = strTmp & Mid(strFormat, i, 1)
        End If
    Next

    FormatString = strTmp
    Exit Function

errHandle:
End Function

Public Function VerifyString(ByVal strTarget As String, ByVal strStandard As String, _
    Optional ByVal blnStandard As Boolean = True) As Boolean
    
'���ܣ����Ŀ���ַ������޷Ǳ�׼�ַ�
'������
'  strStandard����׼���ַ���
'  strTarget��Ҫ����Ŀ���ַ���
'  blnStandard��True-strStandardΪ��׼�ַ���False-strStandardΪ�Ǳ�׼�ַ�
'���أ�Trueͨ����Falseδͨ��

    Dim i As Integer
    
    For i = 1 To Len(strTarget)
        If blnStandard Then
            If InStr(strStandard, Mid(strTarget, i, 1)) <= 0 Then
                VerifyString = False
                Exit Function
            End If
        Else
            If InStr(strStandard, Mid(strTarget, i, 1)) > 0 Then
                VerifyString = False
                Exit Function
            End If
        End If
    Next
    
    VerifyString = True

End Function

Public Function Encrypt(ByVal strSource As String) As String
'����
    Dim BLowData As Byte
    Dim BHigData As Byte
    Dim i As Long
    Dim k As Integer
    Dim strEncrypt As String
    Dim StrChar As String
    Dim KeyTemp As String
    Dim Key1 As Byte
    
    For k = 1 To 30
        KeyTemp = KeyTemp & CStr(Int(Rnd * (9) + 1))
    Next
    
    Key1 = CByte(Mid(KeyTemp, 11, 1) & Mid(KeyTemp, 27, 1))
    
    For i = 1 To Len(strSource)
        StrChar = Mid(strSource, i, 1)                                      '�Ӵ������ַ�����ȡ��һ���ַ�
        BLowData = AscB(MidB(StrChar, 1, 1)) Xor Key1                       'ȡ�ַ��ĵ��ֽں�Key1�����������
        BHigData = AscB(MidB(StrChar, 2, 1))                                'ȡ�ַ��ĸ��ֽ�
        strEncrypt = strEncrypt & ChrB(BLowData) & ChrB(BHigData)       '�����������ݺϳ��µ��ַ�
    Next i
    
    Encrypt = KeyTemp & strEncrypt
End Function

Public Function Decrypt(ByVal strSource As String) As String
'����
    Dim BLowData As Byte
    Dim BHigData As Byte
    Dim i As Long
    Dim k As Integer
    Dim StrDecrypt As String
    Dim StrChar As String
    Dim KeyTemp As String
    Dim Key1 As Byte
    
    KeyTemp = Mid(strSource, 1, 30)
    Key1 = CByte(Mid(KeyTemp, 11, 1) & Mid(KeyTemp, 27, 1))
    
    For i = 31 To Len(strSource)
        StrChar = Mid(strSource, i, 1)                                      '�Ӵ������ַ�����ȡ��һ���ַ�
        BLowData = AscB(MidB(StrChar, 1, 1)) Xor Key1                       'ȡ�ַ��ĵ��ֽں�Key1�����������
        BHigData = AscB(MidB(StrChar, 2, 1))                                'ȡ�ַ��ĸ��ֽ�
        StrDecrypt = StrDecrypt & ChrB(BLowData) & ChrB(BHigData)       '�����������ݺϳ��µ��ַ�
    Next i
    
    Decrypt = StrDecrypt
End Function

Public Function Base64Encode(strSource As String) As String
    Dim arrBase64() As String
    Dim arrB() As Byte, bTmp(2) As Byte, bT As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    If UBound(arrBase64) = -1 Then
        arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
    End If
    
    arrB = StrConv(strSource, vbFromUnicode)

    j = UBound(arrB)
    For i = 0 To j Step 3
        Erase bTmp
        bTmp(0) = arrB(i + 0)
        bTmp(1) = arrB(i + 1)
        bTmp(2) = arrB(i + 2)

        bT = (bTmp(0) And 252) / 4
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(0) And 3) * 16
        bT = bT + bTmp(1) \ 16
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(1) And 15) * 4
        bT = bT + bTmp(2) \ 64
        If i + 1 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If

        bT = bTmp(2) And 63
        If i + 2 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If
    Next
End Function

Public Function Base64Decode(strEncoded As String) As String '??
    Dim arrB() As Byte, bTmp(3) As Byte, bT As Long, bRet() As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

Public Function GetRegItemValue(ByVal lngKey As Long, ByVal strSubKey As String, _
    ByRef blnFinish As Boolean) As String
    
'���ܣ�������Ŀ¼���ض�����Ŀ���Ƶ�ֵ
'������
'  lngKey��ע�������
'  strSubKey��ע���Ŀ¼��
'���أ�ָ����Ŀ��ֵ
    
    Const STR_HOME_KEY_1 As String = "ORACLE_HOME"
    Const STR_HOME_KEY_2 As String = "ORA_CRS_HOME"

    Dim lngRet As Long, lngResult As Long, lngLen As Long, lngIndex As Long, lngReserved As Long, lngClass As Long
    Dim strName As String, strClass As String, strResult As String, strTmp As String
    Dim LWT As FILETIME
    Dim blnTemp As Boolean
    
    lngRet = RegOpenKey(lngKey, strSubKey, lngResult)
    
    Do While lngRet = ERROR_SUCCESS
        strName = String(255, Chr(0))
        lngLen = Len(strName)
        lngRet = RegEnumKeyEx(lngResult, lngIndex, strName, lngLen, lngReserved, strClass, lngClass, LWT)
        If lngRet = ERROR_SUCCESS Then
            strName = Left(strName, InStr(strName, Chr(0)) - 1)
            strTmp = strSubKey & "\" & strName
'Debug.Print strTmp
            strResult = GetRegItemValue(lngKey, strTmp, blnTemp)
            If strResult = "" Then
                '����Ŀ¼ʱ����ʼ����Ŀ����Ŀֵ
                strResult = GetKeyValue(lngKey, strTmp, STR_HOME_KEY_1)
                If strResult <> "" Then
                    GetRegItemValue = strResult
                    blnFinish = True
                    Exit Do
                Else
                    strResult = GetKeyValue(lngKey, strTmp, STR_HOME_KEY_2)
                    If strResult <> "" Then
                        GetRegItemValue = strResult
                        blnFinish = True
                        Exit Do
                    End If
                End If
            ElseIf blnFinish Then
                GetRegItemValue = strResult
                Exit Do
            End If
        End If
        lngIndex = lngIndex + 1
    Loop
    
    Call RegCloseKey(lngRet)

End Function
