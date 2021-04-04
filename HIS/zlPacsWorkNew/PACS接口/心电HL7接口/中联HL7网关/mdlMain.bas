Attribute VB_Name = "mdlMain"
Option Explicit

Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrDbUserPwd As String              '��ǰ���ݿ�����
Public gstrServerName As String             '��ǰ���ݿ������

Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstr��λ���� As String
Public glngSys As Long
'-----------------------------------------
'�����롢ע���롢�������������ע���������
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

Public gcnAccess As New ADODB.Connection
Public gstrAccessPath As String         'Access���ݿ���ļ�·�����ļ���ǰ׺���ޡ�.mdb��
Public gstrAccessName As String         'Access���ݿ���ļ�·�����ļ���

Public gstrLocalIP As String             '�洢����IP��ַ
Public gblnServiceStart As Boolean         '��������

'��¼��־
Public gblnProcessLog As Boolean        '�Ƿ��¼��־����������������صĴ���
Public glngProcessLogLevel As Long      '��¼������־�ļ��𣬷ֳ�3����1��ֻ��ͨѶϢ�������־��2����¼ͨѶ�ʹ������־��3����¼ͨѶ�ʹ������ϸ��־

'���յ���HL7��Ϣ�Ķ���
Public gstrMsgQueue() As String         '��Ϣ���ݶ���
Public gblnQueueBusy As Boolean         '��¼��ǰ�����Ƿ��ڴ�����У�һ��ֻ����һ�����������
Public gintQueueIndex As Integer        '��¼�����е�һ����Ϣ������
Public gblnMsgProcessing As Boolean     '���ڴ�����Ϣ
Public gintTimeOut() As Integer         '��¼ÿһ����Ϣ���ӵ�ʱ��
Public gintTimeOutMax As Integer        '��ʱ�����ʱ��

'��Ϣ���ղ���
Public gintInputDataType As Integer     '������Ϣ�ķ�ʽ��Ĭ����0��0-socket��ʽ��1-�ļ���ʽ��
Public gstrFileDir As String            '�ļ���ʽ������Ϣ��·��
Public gstrFileSuffix As String         '�ļ���ʽ������Ϣ�ĺ�׺
Public gstrFileBackupDir As String      '�ļ���ʽ������Ϣ�ı���·��

Public gstrRegPath As String            '�������ע���·��

Public gzlDatabase As Object           '�������ݿ�ģ�� zlComLib��zlDatabase
Public gzlComLib As Object             '�������ݿ�ģ�� zlComLib��zlDatabase

'-------------����10.35.10֮ǰ�İ汾����--------------
Public gblnBefore3510 As Boolean       '����10.35.10ǰ��汾��True=10.35.10֮ǰ�汾,��ʹ��zlRegister����ʼ��comlibʱ��ҪSetDbUser��RegCheck
Public SplashObj As New frmSplash
Public gstrStation As String           '������վ����
Public gstrParseRegCode As String
'-------------����10.35.10֮ǰ�İ汾����--------------


'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------
Public Sub Main()
    Dim objRegister As Object         '10.35.10֮���ע�����
    
    Set gzlDatabase = CreateObject("zl9ComLib.clsDatabase")
    Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
    
    If App.PrevInstance Then
        MsgBox "HL7���ط����Ѿ������������ٴ����С�", vbInformation, "����"
        Exit Sub
    End If
    
    On Error Resume Next
    '��ͨ��zlRegister�����ж��ǲ���10.35.10֮��İ汾������汾֮�󣬵�¼���ݿ�����벻һ����
    Set objRegister = GetObject("", "zlRegister.clsRegister")
    If objRegister Is Nothing Then
        gblnBefore3510 = True   '35.10֮ǰ�İ汾
    Else
        gblnBefore3510 = False
        Set objRegister = Nothing
    End If
    
    err.Clear
    On Error GoTo err

    If gblnBefore3510 Then
        If LoginBefore3510 = False Then Exit Sub
    Else
        If LoginAfter3510 = False Then Exit Sub
    End If
   
    CodeMan 2000
    Exit Sub
err:
    MsgBox "�������س��ִ��󣬴��������ǣ�" & err.Description, vbOKOnly
End Sub

Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '���أ�
    '------------------------------------------------
    Dim rsUser As ADODB.Recordset
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = gzlComLib.GetUnitName()
    
    '��ȡ�û�����Ϣ
    Set rsUser = gzlDatabase.GetUserInfo
    If rsUser.RecordCount <> 0 Then
        glngUserId = Nvl(rsUser!ID)
        gstrUserCode = Nvl(rsUser!���)
        gstrUserName = Nvl(rsUser!����)
        gstrUserAbbr = Nvl(rsUser!����)
        glngDeptId = Nvl(rsUser!����ID)
        gstrDeptCode = Nvl(rsUser!������)
        gstrDeptName = Nvl(rsUser!������)
    Else
        glngUserId = 0
        gstrUserCode = ""
        gstrUserName = ""
        gstrUserAbbr = ""
        glngDeptId = 0
        gstrDeptCode = ""
        gstrDeptName = ""
    End If
    
    gstrPrivs = gzlComLib.GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 2000
            frmHL7Main.Show
    End Select
End Sub

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
    
    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    gstrServerName = strServerName
    gstrDbUserPwd = strUserPwd
    gstrDbUser = UCase(strUserName)
    gzlComLib.SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If gzlComLib.ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
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

Public Sub CheckDBConnect()
    '������ݿ�Ͽ����������������ݿ�
    On Error GoTo ConnErr
    If gcnOracle.State <> 1 Then
        gcnOracle.Provider = "MSDataShape"
        gcnOracle.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServerName, gstrDbUser, gstrDbUserPwd
    End If
    gzlDatabase.OpenSQLRecord "select '����'  from dual", "���������Ƿ�ɹ�"
    Exit Sub
ConnErr:
    On Error Resume Next
    If gcnOracle.State = 1 Then
        gcnOracle.Close
    End If
End Sub

Private Function LoginBefore3510() As Boolean
'10.35.10֮ǰ�ĵ�¼����

    Dim BlnShowFlash As Boolean
    Dim StrUnitName As String
    Dim intCount As Integer
    Dim lngReturn As Long
    Dim strCode As String
    
    
    LoginBefore3510 = False
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    
    BlnShowFlash = False
    Load SplashObj
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If StrUnitName <> "" Then
        With SplashObj
            '��������Ҫ����
            Call gzlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call gzlComLib.ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .lbl������.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
        End With
        
        BlnShowFlash = True
        DoEvents
    End If
    
    gstrStation = Space(200)
    lngReturn = GetComputerName(gstrStation, 200)
    gstrStation = Trim(gstrStation)
    If Len(gstrStation) > 1 Then
        gstrStation = Left(gstrStation, Len(gstrStation) - 1)
    Else
        gstrStation = "..."
    End If
    
    '�û�ע��
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Function
    End If
    
    '��ʼ����������
    gzlComLib.InitCommon gcnOracle
    If gzlComLib.RegCheck = False Then
        Unload SplashObj
        Exit Function
    End If
    
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gstrParsePublish = gzlComLib.zlRegInfo("��Ʒ����")
    gstrParseRegCode = gzlComLib.zlRegInfo("��λ����", , -1)
    
    gstrSysName = gstrParsePublish & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\�����ļ�"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl����֧����.Caption = gzlComLib.zlRegInfo("����֧����", , -1)
            .LblProductName = gzlComLib.zlRegInfo("��Ʒ����")
            
            strCode = gzlComLib.zlRegInfo("��Ʒ������", , -1)
            .lbl������.Caption = ""
            For intCount = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(intCount) & vbCrLf
            Next
            Call gzlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gstrParseRegCode
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", gzlComLib.zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", gzlComLib.zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", gzlComLib.zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", gzlComLib.zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", gzlComLib.zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", gzlComLib.zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", gzlComLib.zlRegInfo("֧����URL")

    gstrSystems = " (ϵͳ =100 Or ϵͳ Is NULL)"
    glngSys = 100
    '-------------------------------------------------------------
    '����ͬ���
    '-------------------------------------------------------------
    gzlDatabase.ExecuteProcedure "Zl_Createsynonyms(" & glngSys & ")", "����ͬ���"
    
    '-------------------------------------------------------------
    'ѡ����ò�ͬ��񵼺�̨
    '-------------------------------------------------------------
    On Error Resume Next
    err = 0
    
    Unload SplashObj
    
    LoginBefore3510 = True
End Function

Private Function LoginAfter3510() As Boolean
'10.35.10֮��ĵ�¼����

    Dim objLogin As Object
    Dim strCommand As String
    
    LoginAfter3510 = False
    
    Set objLogin = DynamicCreate("zlLogin.clsLogin", "zlLogin.dll")
    If objLogin Is Nothing Then Exit Function
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    
    Set gcnOracle = objLogin.Login(0, CStr(Command()))
    
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.ConnectionString = "" Then Exit Function

    gstrUserName = objLogin.DBUser
    
    '��ʼ����������
    gzlComLib.InitCommon gcnOracle
    
    gstrParsePublish = gzlComLib.zlRegInfo("��Ʒ����")
    gstrSysName = gstrParsePublish & "���"
    
    gstrSystems = " (ϵͳ =100 Or ϵͳ Is NULL)"
    glngSys = 100
    
    LoginAfter3510 = True
End Function
