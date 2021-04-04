Attribute VB_Name = "mdlPublic"
Option Explicit

'���ò�������   1:ZLHIS:83:1:0:0
'��������ҽ��   2:ZLHIS:1:1:0:0
'����סԺҽ��   3:ZLHIS:83:1:0:0
'����PACS����   4:ZLHIS:83:1:1:1008

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'���̼䴫���ڴ�ռ䣬���Դ��ַ���
Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Public Const SW_RESTORE = 9
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const C_LOG = 0 '�Ƿ��¼��־,0��������־��1Ҫ��¼������־

'��ϢHook����
Public plngPreWndProc As Long       'ԭ������Ϣ�������
Public Const MSG_SPLIT = ":"

Private mobjRegister As Object                  '10.35.10֮���ע�����

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public gstrZLHIS�����ַ��� As String
Public gstr�û��� As String
Public gstr���� As String
Public gbln�Ƿ�ת������ As Boolean
Public glng����ID As Long
Public glng��ҳID As Long
Public glngFunID As Long
Public glng����ID As Long
Public glng����ID As Long
Public glng���ܺ� As Long '0-��������,1-LIS����,2-ҽ������;3-ִ�ж˸�������;4-ִ�ж˸���;5-���ݴ�ӡ;99-�Զ��屨��


Public glngPid As Long

Public gstrHwndOLD As String
Public gstrHwndNew As String



Private mclsReport As Object


'��������
Public gblnXWRISInterfaceLog As Boolean         '�����ݿ���д����־
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetRandom(ByVal lngBase As Long) As String
'------------------------------------------------
'���ܣ���ȡһ��1����lngBase-1��֮������������������ASCII��
'������ lngBase  --  ����������ֵ����������Ϊ��lngBase-1��
'���أ���������ASCII��
'------------------------------------------------
    Dim lngNum As Long
    
    Randomize
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function


Public Function getEncryptionWord(ByVal strPassW As String) As String
'------------------------------------------------
'���ܣ���ȡ�������ģ�ʹ��1-29֮���ASCII�룬��Ϊ����������ַ������ܣ�����㷨ֻ��������1-29��ASCII�룬������Χ�ᵼ�½���ʧ��
'������ strPassW  --  ��Ҫ���ܵ�Դ��
'���أ�������������
'------------------------------------------------
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    '���strBase=strRandom�����ܺ�����ģ������ԭ�ģ�����Ҫȷ��������ֵ����ͬ
    If strRandom = strBase Then
        If Asc(strBase) >= 29 Then
            strRandom = Chr(1)
        Else
            strRandom = Chr(Asc(strRandom) + 1)
        End If
    End If
    
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionWord = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

Public Function getDecryptionWord(ByVal strPassW As String) As String
'------------------------------------------------
'���ܣ���ȡ���ܵ�Դ��
'������ strPassW  --  ��Ҫ���ܵ�����
'���أ���������Դ��
'------------------------------------------------
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

    getDecryptionWord = Join(strTemp, "") '���ܺ���ִ�
End Function

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'���ܣ�������
'������ logSubName  --  ��������ĺ�����
'       logTitle   -- ��������
'       logDesc   --  ��������
'���أ�1-�������Resume��0-�����˳�
'------------------------------------------------
    
    errHandle = 0
    
    '��ʾ����
    MsgBox errTitle & errDesc, vbOKOnly, "�ӿ�zlSoftCISInterface���ִ���"
    
    '�������
    err.Clear
    
End Function


Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function ConnectDB(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'���ܣ��������ݿ⣬��ע����ж�ȡ���ܺ�����ݿ�������Ϣ���û��������룬������
'������
'���أ�True-�ɹ���False-ʧ��
'------------------------------------------------
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnTransPassword As Boolean
    
    ConnectDB = False
    
    On Error GoTo err
    
    If gcnOracle.State <> adStateOpen Then
        strDBServer = gstrZLHIS�����ַ���
        strDBUser = gstr�û���
        strDBPassword = gstr����
        blnTransPassword = gbln�Ƿ�ת������
                
        '�������ݿ�
        If OraDataOpen(strDBServer, strDBUser, strDBPassword, blnTransPassword) = False Then
           
            Exit Function
        End If
    End If
    
    ConnectDB = True
    Exit Function
err:
    If errHandle("zlSoftCISInterface.ConnectDB", "�������ݿ⺯�����ִ���", err.Description) = 1 Then Resume
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal blnTransPassword As Boolean) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '   blnTransPassword �� �Ƿ���Ҫת������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error GoTo ErrHand
    

    If gblnBefore3510 = True Then
        '�����10.35.10֮ǰ�İ汾��ֱ�����û����������¼���ݿ�
        OraDataOpen = OpenOracle(gcnOracle, strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, IIf(blnTransPassword = True, TranPasswd(strUserPwd), strUserPwd)))
    Else
        '�����10.35.10֮��İ汾��ʹ��zlRegister��ȡ���ݿ�����
        Set gcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, , strError, True)
        If gcnOracle.State = adStateOpen Then
            OraDataOpen = True
        Else
            OraDataOpen = False
        End If
    End If
    
    If OraDataOpen = True Then
        gstrDBUser = UCase(strUserName) '����ΪʲôҪǿ�ƴ�д���ǲ���comlib��Ҫ��
        If gblnBefore3510 = True Then
            '10.35.10֮ǰ�İ汾
            gzlComLib.SetDbUser gstrDBUser
        End If
    End If
    
    Exit Function
    
ErrHand:
    
    If errHandle("zlSoftCISInterface.OraDataOpen", "�������ݿ����", err.Description) = 1 Then Resume
    OraDataOpen = False
End Function

Private Function OpenOracle(ByRef cnOrcle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ����Oracle���ݿ�
    '������
    '   cnOrcle �����ݿ�����
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With cnOrcle
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
            
            OpenOracle = False
            Exit Function
        End If
    End With
    
    OpenOracle = True
    err = 0
    
    Exit Function
    
End Function

Private Function TranPasswd(strOld As String) As String
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

Public Sub ShowSubWindow(ByVal lngHwnd As Long, Optional ByVal lngMainHwnd As Long)
'���ܣ���ʾָ���Ĵ��壬���Ӵ��巽ʽ
'������lngHwnd=Ҫ��Ϊ�Ӵ�����ʾ�Ĵ���ľ��
'      lngMainHwnd=��������������ʱ�������������Ӵ�����ʾ
'˵�����������Ҫ������ZLBH���ںϵ���ZLHIS������ʾ
    Dim vRect1 As RECT, vRect2 As RECT
    Dim X As Long, Y As Long
    
    If lngHwnd <= 0 Then Exit Sub
    
    If lngMainHwnd <> 0 Then
        SetParent lngHwnd, lngMainHwnd
    Else
        SetParent lngHwnd, 0
    End If

    '��ʾ�ڸ����������
    If IsWindowVisible(lngHwnd) = 0 Then
        GetWindowRect lngHwnd, vRect1
        GetWindowRect lngMainHwnd, vRect2
        
        X = ((vRect2.Right - vRect2.Left) - (vRect1.Right - vRect1.Left)) / 2
        Y = ((vRect2.Bottom - vRect2.Top) - (vRect1.Bottom - vRect1.Top)) / 2
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        
        SetWindowPos lngHwnd, 0, X, Y, 0, 0, &H40 Or &H1 'HWND_TOP=0
    End If
    
    ShowWindow lngHwnd, SW_RESTORE
End Sub

Public Function UpdateEmrInterface() As Object
    Dim objEmr As Object
    Dim strDBPassword As String
    Dim strDBServer As String

    
    On Error Resume Next
    err.Clear
    If GetEMRLoginUser(strDBServer, strDBPassword) Then
        Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
        If err.Number = 0 Then
            Call objEmr.CheckUpdate1(gstrDBUser, strDBPassword, True)
            If err.Number <> 0 Then
                err.Clear
                If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                    Exit Function
                End If
            End If
        Else
            err.Clear
        End If
    Else
        Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
        If err.Number = 0 Then
            '��ע����ȡ���ݿ�������Ϣ
            strDBServer = gstrZLHIS�����ַ���
            strDBPassword = gstr����
    
            Call objEmr.CheckUpdate1(gstrDBUser, strDBPassword, True)
            If err.Number <> 0 Then
                err.Clear
                If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                    Exit Function
                End If
            End If
        Else
            err.Clear
        End If
    End If

    
    Set UpdateEmrInterface = objEmr
    On Error GoTo 0
End Function

Public Function InitInterface(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'���ܣ���ʼ���ӿڣ�����ComLib���������ݿ�
'��������
'���أ�True-�ɹ���False-ʧ��
'------------------------------------------------
    
    On Error GoTo err
    InitInterface = False
    
    '��ʼ��ϵͳ��Ϊ100��ģ���Ϊ1287
    glngSys = 100
    glngModule = 1287
    
    '������־Ŀ¼
    Call MkLocalDir(gstrLogPath + "\")
    Call MkLocalDir(gstrBackupPath + "\")
 
On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        If mobjRegister Is Nothing Then gblnBefore3510 = True '35.10֮ǰ�İ汾
    End If
    
    err.Clear
On Error GoTo err
    If gzlComLib Is Nothing Then
        If gblnBefore3510 Then
            '10.35.10֮ǰ�İ汾
            Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
        Else
            '10.35.10֮��İ汾
            Set gzlComLib = GetObject("", "zl9ComLib.clsComLib")
        End If
    End If
    
    '����Ǵ�RIS������DLL�����ݿ�����gzlComLib.CurrentConn�ǿյģ���Ҫ��ע����ȡ�û������룬�����������ݿ�
    If gzlComLib.CurrentConn Is Nothing Then
        '��ע����ȡ�û������룬�������ݿ�
        
        '���gcnOracle�����ڣ�Ҫ�½�һ��
        If gcnOracle Is Nothing Then Set gcnOracle = New ADODB.Connection
        Call ConnectDB(strDBUser)

        '��ʼ����������
        gzlComLib.InitCommon gcnOracle
        

        If gblnBefore3510 = True Then
            '10.35.10֮ǰ�İ汾
            If gzlComLib.RegCheck = False Then
                
                Exit Function
            End If
        End If
    Else
        '����Ǵ�HIS����̨������DLL���򴴽�zl9ComLib֮�󣬻��Զ�������gzlComLib.CurrentConn
        '������ʱû�д� CodeMan��ȡ�� gcnOracle��������Ҫ��zl9ComLibȡ��gcnOracle����
        'gstrDBUser��ע����ж�ȡ��������clsHISInner.SaveDBConnectInfo
        
        If gcnOracle Is Nothing Then Set gcnOracle = gzlComLib.CurrentConn
    End If
    
    InitInterface = True
    
  
    Exit Function
err:
    If errHandle("zlSoftCISInterface.InitInterface", "��ʼ���ӿڳ���", err.Description) = 1 Then Resume
End Function

Public Function InitSysParameter() As Boolean
'------------------------------------------------
'���ܣ���ʼ��ȫ�ֲ���
'��������
'���أ�True-�ɹ���False-ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    InitSysParameter = False
    
    ''��ȡ�Ƿ�����Ӱ����Ϣϵͳ�ӿ�
    gblnUseInterface = Val(gzlComLib.zlDatabase.GetPara(255, glngSys)) = 1
    
    InitSysParameter = True
    
On Error GoTo Error
    'gblnXWRISInterfaceLogĬ��Ϊfalse
    gblnXWRISInterfaceLog = False
    
    strSQL = "Select ���� From zlRegInfo Where ��Ŀ = '��¼רҵ��RIS��־'"
    Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface")
    
    If rsData.RecordCount > 0 Then
        gblnXWRISInterfaceLog = Nvl(rsData!����, "0") = "1"
    End If
    
    Exit Function
Error:
    If errHandle("zlSoftCISInterface.InitSysParameter", "�ж��Ƿ��¼��־�����ݿ�ʱ���ִ���", strSQL) = 1 Then Resume
End Function

'======================================================================================================================
'����           ByteToHexString         ��16�����ַ���ת��Ϊ�ֽ���
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'bstrInput      String                  16�����ַ���
'lngRetBytLen   Long(Optional)          ָ�����ص��ֽ���ĳ���,0-��ԭʼ���ȷ��أ�<>0����ָ���ĳ��ȣ����㲹�루��0�������˽�ȡ
'======================================================================================================================
Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    lngLen = Len(strInput)
    If lngRetBytLen <> 0 Then
        lngLen = lngLen \ 2
        If lngLen > lngRetBytLen Then
            lngLen = lngRetBytLen
        End If
        ReDim arrReturn(lngRetBytLen - 1)
    Else
        lngLen = lngLen \ 2
        ReDim arrReturn(lngLen - 1)
    End If
    
    For i = 0 To lngLen - 1
        arrReturn(i) = Val("&H" & Mid(strInput, 2 * i + 1, 2))
    Next
    
    HexStringToByte = arrReturn()
End Function

Private Function TruncZeroInside(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�,�������ù���,���Ե�������clsstring
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZeroInside = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZeroInside = strInput
    End If
End Function

'======================================================================================================================
'����           sm_version              ��ȡZLSM4�İ汾��
'����ֵ         Long                    ZLSM4�İ汾��
'����б�:
'======================================================================================================================
Public Function sm_version() As Long
    Dim lngVersion As Long
    On Error Resume Next
    lngVersion = get_sm_version
    If err.Number <> 0 Then
        err.Clear
        sm_version = 1
    Else
        sm_version = lngVersion
    End If
End Function


Private Function GetKey(ByVal strKey As String, ByVal intType As Integer) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    If strKey <> "" Then
        arrReturn = HexStringToByte(strKey, 16)
    Else
        ReDim arrReturn(15)
        If intType = 0 Then
            For i = 0 To 15
                arrReturn(i) = i * 15
            Next
        ElseIf intType = 1 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_IV)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        ElseIf intType = 2 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_KEY)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        End If
    End If
    GetKey = arrReturn
End Function

'======================================================================================================================
'����           Sm4DecryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ��������ַ�����Sm4EncryptEcb���ɵĽ����
'strKey         String(Optional)        ������ԿҲ���ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'======================================================================================================================
Public Function Sm4DecryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))
    
    Call sm4_crypt_ecb(CM_Decrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
    If lngVersion = 1 Then
        Sm4DecryptEcb = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptEcb = TruncZeroInside(StrConv(arrOutPut(), vbUnicode))
    End If
End Function

Private Function GetEMRLoginUser(strUser As String, strPwd As String) As Boolean
'���ܣ���ȡEMP��ʼ�����û�������
'���أ��Ƿ��ȡ�ɹ���������ֻ����2500ϵͳ����������ļ���ȡ����������100��2500ϵͳ���򷵻�FALSE

    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim rsTest      As ADODB.Recordset
    Dim strConn     As String
    Dim objFSO      As New FileSystemObject
    Dim arrInfo As Variant
    Dim strCode     As String

    On Error GoTo errH
    strSQL = "Select Floor(a.��� / 100) ��� From zlSystems A Where Floor(a.��� / 100) In (1, 25)"
    Set rsTmp = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "GetEMRLoginUser")
    If rsTmp.RecordCount <> 0 Then
        rsTmp.Filter = "���=1"
        If rsTmp.RecordCount = 0 Then
            rsTmp.Filter = "���=25"
            If rsTmp.RecordCount <> 0 Then
                strSQL = "Select ����ֵ From zlOptions Where  ������ =[1]"
                Set rsTest = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡLIS��������", "LISϵͳ��������")
                If rsTest.RecordCount > 0 Then
                    strConn = rsTest("����ֵ") & ""
                End If
                If strConn <> "" Then
                    strCode = Sm4DecryptEcb(strConn)
                    arrInfo = Split(strCode, "<SP 1>")
                    If UBound(arrInfo) >= 1 Then
                        If arrInfo(0) <> "" And arrInfo(1) <> "" Then
                            strUser = arrInfo(0)
                            strPwd = IIf(UCase(arrInfo(0)) = "SYS" Or UCase(arrInfo(0)) = "SYSTEM", "[DBPASSWORD]", "") & arrInfo(1)
                            GetEMRLoginUser = True
                        End If
                    End If
                End If
            End If
        Else
            strUser = gstrDBUser
            strPwd = IIf(gbln�Ƿ�ת������, "", "[DBPASSWORD]") & gstr����
            GetEMRLoginUser = True
        End If
    End If
    Exit Function
errH:
    err.Clear
End Function

Public Function AnalyseComputer() As String
'��ȡ�������
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function InStrW(wStr As String, wIn As String, wTimes As Long) As Long
    Dim sPos As Long
    Dim s As String
    
    s = Replace(wStr, wIn, "", 1, wTimes - 1, vbBinaryCompare)
    sPos = InStr(s, wIn)
    InStrW = sPos + wTimes - 1
End Function


Public Function ProcessMessage(strmsg As String) As Long
    '-----------------------------------------------
    '������յ�����Ϣ
    '��Ϣ��ʽ��Oracle�����ַ���:�û���:����:�Ƿ�����ת��(0��1):���ù��ܺ�(0-��������,1-�����鱨��,2-ҽ������;3-ִ�ж˸�������;4-ִ�ж˸���;5-���ݴ�ӡ;99-������;999-���ܳ�ʼ��):...
    '              ���ܺŲ�ͬ�����������ĸ�ʽ�뺬��Ҳ��ͬ
    '              ����=0,1,2ʱ:���ܺź����������ID,��ҳID
    '              ����=3��999ʱ�����ܺ��޲���
    '              ����=4ʱ,���ܺ�Ϊ:����ID:ҽ����Ϣ:NOs
    '                      ����ҽ����Ϣ��NOs���δ�һ������,ҽ����Ϣ��ִ�п���|ҽ��IDs(����ö��ŷָ�);NOs: ����ö��ŷָ�
'                  ����=5ʱ,���ܺ�Ϊ����ӡ���(0=����ӡ��Ԥ��,1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF,99-��ӡ����):(��ʽ��������,���ݺ�(par)������,���ݺ�)���ܺ�Ϊ:  ������,���ݺ�(par)������,���ݺ�(par)������,���ݺ�
    '              ����=99ʱ,���ܺ�Ϊ��ϵͳ��:������:��ӡ���(0=����ӡ��Ԥ��,1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF,99-��ӡ����):�������(��Ϊ�� ʾ����ʽ��"����id=1<par>PDF=C:\1.PDF<par>ExcelFile=C:\1.xls")
    
    
    '-----------------------------------------------
    Dim blnLis As Boolean
    Dim strErr As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strҽ����Ϣ As String
    Dim strNos     As String
    
    Dim lngϵͳ�� As Long
    Dim str������ As String
    Dim lng��ӡ��� As Long '��ӡ���(0=����ӡ��Ԥ��,1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF,99-��ӡ����)
    Dim str������� As String '
    Dim arrPar As Variant
    Dim arrPar�̶�(50) As String
    Dim arrNoPar As Variant



    Dim i As Long
                     
    
    Dim varData As Variant
    
    
    On Error GoTo err
    ProcessMessage = 1
    
    varData = Split(strmsg & String(7, MSG_SPLIT), MSG_SPLIT)
    gstrZLHIS�����ַ��� = varData(0)
    gstr�û��� = varData(1)
    gstr���� = varData(2)
    gbln�Ƿ�ת������ = Val(varData(3)) = 1
    glng���ܺ� = Val(varData(4)) '0-��������,1-LIS����,2-ҽ������;3-ִ�ж˸�������;4-ִ�ж˸���;5-���ݴ�ӡ
    If glng���ܺ� <> 3 And glng���ܺ� <> 99 And glng���ܺ� <> 5 Then
        glng����ID = Val(varData(5))
        If glng���ܺ� = 4 Then
            strҽ����Ϣ = varData(6)
            strNos = varData(7)
        Else
            glng��ҳID = Val(varData(6))
        End If
        glngFunID = IIf(glng���ܺ� = 2, 3001, 0)
    End If
    
    If glng���ܺ� = 99 Then
        lngϵͳ�� = Val(varData(5))
        str������ = varData(6)
        lng��ӡ��� = Val(varData(7))
        
        str������� = Mid(strmsg, InStrW(strmsg, MSG_SPLIT, 8) + 1) '����split��Ԥ�������������ð��
        If str������� <> "" Then
            str������� = Mid(str�������, 2)
            str������� = Mid(str�������, 1, Len(str�������) - 1)
        End If
    End If
    If glng���ܺ� = 5 Then
        lng��ӡ��� = Val(varData(5))
        str������� = varData(6)
    End If
    
    
    
    If InStr("0,1,2", glng���ܺ�) > 0 And (glng����ID = 0) Then
        Exit Function
    ElseIf glng���ܺ� = 4 And (strҽ����Ϣ = "" And strNos = "") Then
        Exit Function
    ElseIf glng���ܺ� = 5 And str������� = "" Then
        Exit Function
    ElseIf glng���ܺ� = 99 And (str������ = "") Then
        Exit Function
    End If
    
    blnLis = glng���ܺ� = 1
    
    If glng���ܺ� = 2 Then
        strSQL = "Select a.��ǰ����ID,a.��Ժ����ID From ������ҳ a Where a.����id=[1] and a.��ҳid=[2]"
        Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface", glng����ID, glng��ҳID)
        If Not rsData.EOF Then
            glng����ID = Val(rsData!��ǰ����ID & "")
            glng����ID = Val(rsData!��Ժ����ID & "")
        End If
    End If
    
    If Not mfrmShowHisForms Is Nothing Then
        Call GetWindowThreadProcessId(mfrmShowHisForms.hWnd, glngPid)
    End If
    
    gstrHwndOLD = "": EnumChildWindows GetDesktopWindow, AddressOf EnumChildProcOld, ByVal 0
    Select Case glng���ܺ�
        Case 0, 1 '��������
            '���鱨�����
            If Not mclsReport Is Nothing Then
                mclsReport.CloseWindows
            End If
            If blnLis Then
                If mobjLisInsideComm Is Nothing Then
                    Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
                    If mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                        If strErr <> "" Then
                            If errHandle("zlSoftCISInterface.ProcessMessage", strErr) = 1 Then Resume
                            Exit Function
                        End If
                    End If
                    If mobjLisInsideComm Is Nothing Then
                         MsgBox "LIS�ӿڳ�ʼ��ʧ��", vbInformation, "��ʾ"
                         Exit Function
                    End If
                End If
                mfrmShowHisForms.TimerShow.Enabled = True
                Call mobjLisInsideComm.PatientSampleBrowse(frmShowHisForms, glng����ID, "", 0, 0, IIf(glng��ҳID = 0, 1, 2), glng��ҳID)
            Else
                '���Ӳ�����ѯ
                If mclsArchive Is Nothing Then
                    '��һ�ε��õ��Ӳ������ģ���ֵ
                    Set mclsArchive = New clsArchive
                End If
                Call mclsArchive.zlOpenArchiveForm(glng����ID, glng��ҳID)
            End If
        Case 2
            If mclsOrder Is Nothing Then
                Set mclsOrder = New clsOrder
            End If
            
            mclsOrder.zlCloseOrderForm
            mclsOrder.zlOpenOrderForm
        Case 3 'ִ�ж˸�������
            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            Call mclsFee.zlDeviceSetup
        Case 4 'ִ�ж˸���
            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            If mclsFee.zlSquareAffirm(glng����ID, strҽ����Ϣ, strNos) = False Then Exit Function
        Case 99 '�Զ��屨��
            If mclsReport Is Nothing Then
                Set mclsReport = CreateObject("zl9Report.clsReport")
            End If
            
            If (Not mclsReport Is Nothing) And (Not gcnOracle Is Nothing) Then
                mclsReport.CloseWindows
                If lng��ӡ��� = 99 Then
                    Call mclsReport.ReportPrintSet(gcnOracle, lngϵͳ��, str������, mfrmShowHisForms)
                Else
                    If str������� <> "" Then
                        arrPar = Split(str�������, "<par>")
                        For i = LBound(arrPar) To UBound(arrPar)
                            arrPar�̶�(i) = arrPar(i)
                        Next
                    End If
                    
                    mfrmShowHisForms.TimerShow.Enabled = True
                    
                    
                    Call mclsReport.ReportOpen(gcnOracle, lngϵͳ��, str������, mfrmShowHisForms, arrPar�̶�(0), arrPar�̶�(1), arrPar�̶�(2), arrPar�̶�(3), arrPar�̶�(4), arrPar�̶�(5), _
                             arrPar�̶�(6), arrPar�̶�(7), arrPar�̶�(8), arrPar�̶�(9), arrPar�̶�(10), arrPar�̶�(11), _
                             arrPar�̶�(12), arrPar�̶�(13), arrPar�̶�(14), arrPar�̶�(15), arrPar�̶�(16), arrPar�̶�(17), _
                             arrPar�̶�(18), arrPar�̶�(19), lng��ӡ���)
                End If
            End If
        Case 5 '���ݴ�ӡ
            If mclsReport Is Nothing Then
                Set mclsReport = CreateObject("zl9Report.clsReport")
            End If
            
            If (Not mclsReport Is Nothing) And (Not gcnOracle Is Nothing) Then
                mclsReport.CloseWindows
            
            
            
                mfrmShowHisForms.TimerShow.Enabled = True
                If lng��ӡ��� = 99 Then
                    Call mclsReport.ReportPrintSet(gcnOracle, 100, str�������, mfrmShowHisForms)
                Else
                    If str������� <> "" Then
                        arrPar = Split(str�������, "(par)")
                        For i = LBound(arrPar) To UBound(arrPar)
                            arrNoPar = Split(arrPar(i), ",")
                            
                             Call mclsReport.ReportOpen(gcnOracle, 100, arrNoPar(0), mfrmShowHisForms, "NO=" & arrNoPar(1), "����=1", "ҽ��ID=0", "PrintEmpty=0", lng��ӡ���)
                        Next
                    End If
                End If
            End If
            
            If lng��ӡ��� = 2 Then Call CloseAllForms
        Case 999 '��ʼ��
            
            On Error Resume Next
            If mobjLisInsideComm Is Nothing Then
                Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
                If Not mobjLisInsideComm Is Nothing Then
                    Call mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr)
                End If
            End If
            
            Set mclsReport = CreateObject("zl9Report.clsReport")
            
            err.Clear
            '���Ӳ�����ѯ
            If mclsArchive Is Nothing Then
                '��һ�ε��õ��Ӳ������ģ���ֵ
                Set mclsArchive = New clsArchive
                
                If Not mclsArchive Is Nothing Then
                   Call mclsArchive.zlOpenArchiveForm(0, 0, True)
                End If
            End If
            
            If mclsOrder Is Nothing Then
                Set mclsOrder = New clsOrder
            End If

            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            
            err.Clear
            
    End Select
    ProcessMessage = 0
    Exit Function
err:
    MsgBox err.Description
End Function

Public Function CloseAllForms() As Boolean
    On Error GoTo err
    
    If Not mfrmShowHisForms Is Nothing Then
        Call GetWindowThreadProcessId(mfrmShowHisForms.hWnd, glngPid)
        KillPID glngPid
    End If
    
    
    '�رձ������
    If Not mclsReport Is Nothing Then
        Set mclsReport = Nothing
    End If
    
    
    '�ر��շѶ���
    If Not mclsFee Is Nothing Then
        Set mclsOrder = Nothing
    End If
    
    '�ر�ҽ��������
    If Not mclsOrder Is Nothing Then
        mclsOrder.zlCloseOrderForm
        Set mclsOrder = Nothing
    End If
    
    '�رյ��Ӳ������Ĵ���
    If Not mclsArchive Is Nothing Then
        mclsArchive.zlCloseArchiveForm
    End If
'
    '�ر�LIS���������
    If Not mobjLisInsideComm Is Nothing Then
        Set mobjLisInsideComm = Nothing
    End If
    
    '�ر���Ϣѭ��������
    If Not mfrmShowHisForms Is Nothing Then
        Unload mfrmShowHisForms
        Set mfrmShowHisForms = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
   
    Resume Next
End Function

Public Sub writeTestLog(ByVal strInfo As String)
'API����ʱ�Ӻ���
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'����  microsoft script runtime  ��(C:\Windows\System32\scrrun.dll)
    Dim objFile As FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    Dim strTmp As String
    
    If C_LOG = 0 Then Exit Sub
    
    On Error Resume Next
    
    Set objFile = New FileSystemObject
    
    strFile = App.Path & "\zlSoftShowArchiveView.Log"
    
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    
'    If FileLen(strFile) > 52428800 Then     '�ж��ļ��Ƿ����50M
'        Name strFile As App.Path & "\CISJOBTest" & Format(Now, "yyyymmddhhmm") & ".bak"  '�޸��ļ���
'        '���ļ���������֮����Ҫ���¼�鲢�����ļ�
'        Call writeTestLog(strInfo)
'        Exit Sub
'    End If

'    strTmp = "'" & strInfo & "' from dual Union All Select"
    strTmp = strInfo
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strTmp
    objText.Close
'4072 �� Union All Select ����
'insert into �������� (ID,����,�ı�)
'select ��������_ID.Nextval,sysdate,a.* from (
'select �ı� from  �������� Where 1 = 0 Union All Select

'���ļ�Ŀ¼' from dual Union All Select
'���ļ�Ŀ¼' from dual Union All Select
'���ļ�Ŀ¼' from dual Union All Select

'�ı� from  �������� Where 1 = 0) a;


'select id,�ı�,substr(�ı�,1,instr(�ı�,'|')-1) as ����,
'substr(replace(�ı�,substr(�ı�,1,instr(�ı�,'|')) ,''),1,instr(replace(�ı�,substr(�ı�,1,instr(�ı�,'|')) ,''),'|')-1) as ģ��,
'replace(�ı�,substr(�ı�,1,instr(�ı�,'|',-1)) ,'') as ����
'from �������� where id>=4848;
'������־
'    Set objText = objFile.OpenTextFile(strFile, ForReading)
'    Do While Not objText.AtEndOfStream
'        strTmp = objText.ReadLine
'    Loop
'    objText.Close
End Sub

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
 Dim cklm As String * 50 '��������
 Dim lngPid As Long
    On Error Resume Next
    GetClassName hWnd, cklm, 50
    lngPid = 0
    If InStr(LCase(Blank(cklm)), "form") > 0 And InStr(gstrHwndOLD, "," & hWnd & ",") = 0 Then
        Call GetWindowThreadProcessId(hWnd, lngPid)

        If lngPid = glngPid And lngPid <> 0 Then
            
            SetWindowPos hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
            SetWindowPos hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
            BringWindowToTop hWnd
            SetForegroundWindow hWnd
        End If
    End If
    EnumChildProc = 1
End Function


Public Function EnumChildProcOld(ByVal hWnd As Long, ByVal lParam As Long) As Long
 Dim cklm As String * 50 '��������
 Dim lngPid As Long
    On Error Resume Next
    GetClassName hWnd, cklm, 50
    lngPid = 0
    If InStr(LCase(Blank(cklm)), "form") > 0 Then
        Call GetWindowThreadProcessId(hWnd, lngPid)

        If lngPid = glngPid And lngPid <> 0 Then
             gstrHwndOLD = "," & gstrHwndOLD & "," & hWnd & ","
        End If
    End If
    EnumChildProcOld = 1
End Function

Public Function Blank(ByVal szString As String) As String
    Dim l As Integer
    l = InStr(szString, Chr(0))
    If l > 0 Then
        Blank = Left(szString, l - 1)
    Else
        Blank = szString
    End If
End Function


Public Function KillPID(ByVal lngPid As Long) As Boolean
    
    'ɱ������
    On Error Resume Next
    Shell ("taskkill /pid " & lngPid & " -t -f")
End Function

