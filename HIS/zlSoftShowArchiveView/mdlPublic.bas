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
Public C_LOG As Long  '�Ƿ��¼��־,0��������־��1Ҫ��¼������־

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
    MsgBox errTitle & errDesc, vbOKOnly, "�ӿ�zlSoftShowHisForms���ִ���"
    
    '�������
    err.Clear
    
End Function

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
    If errHandle("zlSoftShowHisForms.ConnectDB", "�������ݿ⺯�����ִ���", err.Description) = 1 Then Resume
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
    
    If errHandle("zlSoftShowHisForms.OraDataOpen", "�������ݿ����", err.Description) = 1 Then Resume
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
    Dim blnOut As Boolean
 
    On Error Resume Next

    writeTestLog "UpdateEmrInterface ** strDBServer=" & gstrZLHIS�����ַ��� & ",strDBPassword=" & gstr����
    Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Not objEmr Is Nothing Then
        '��ע����ȡ���ݿ�������Ϣ
        strDBServer = gstrZLHIS�����ַ���
        strDBPassword = gstr���� '3510�������ת�������仯ͳһ��δת��������
        
        If objEmr.CheckUpdate1(gstr�û���, strDBPassword, True) = False Then
           blnOut = False
        Else
            blnOut = True
        End If
        If err.Number <> 0 Then
            err.Clear
            If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                blnOut = False
            Else
                blnOut = True
            End If
        End If
     End If
  If blnOut Then
    Set UpdateEmrInterface = objEmr
   End If
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
    If errHandle("zlSoftShowHisForms.InitInterface", "��ʼ���ӿڳ���", err.Description) = 1 Then Resume
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
    Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "��¼רҵ��RIS��־")
    
    If rsData.RecordCount > 0 Then
        gblnXWRISInterfaceLog = Nvl(rsData!����, "0") = "1"
    End If
    
    Exit Function
Error:
    If errHandle("zlSoftShowHisForms.InitSysParameter", "�ж��Ƿ��¼��־�����ݿ�ʱ���ִ���", strSQL) = 1 Then Resume
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

Public Function ProcessMessage(strmsg As String) As Long
    '������յ�����Ϣ
    '���ݴ���Ĳ����жϴ����ĸ���������Ϣ��ʽ��HIS�����ַ���:�û���:����:�Ƿ�ת������(0/1):����ID:��ҳID��
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim blnLis As Boolean
    Dim strErr As String
    
    On Error GoTo err
    ProcessMessage = 1
    
    If UBound(Split(strmsg, MSG_SPLIT)) = 5 Then
        gstrZLHIS�����ַ��� = Split(strmsg, MSG_SPLIT)(0)
        gstr�û��� = Split(strmsg, MSG_SPLIT)(1)
        gstr���� = Split(strmsg, MSG_SPLIT)(2)
        gbln�Ƿ�ת������ = Val(Split(strmsg, MSG_SPLIT)(3)) = 1
        lng����ID = Val(Split(strmsg, MSG_SPLIT)(4))
        lng��ҳID = Val(Split(strmsg, MSG_SPLIT)(5))
    ElseIf UBound(Split(strmsg, MSG_SPLIT)) = 6 Then
        blnLis = Val(Split(strmsg, MSG_SPLIT)(0)) = 25
        gstrZLHIS�����ַ��� = Split(strmsg, MSG_SPLIT)(1)
        gstr�û��� = Split(strmsg, MSG_SPLIT)(2)
        gstr���� = Split(strmsg, MSG_SPLIT)(3)
        gbln�Ƿ�ת������ = Val(Split(strmsg, MSG_SPLIT)(4)) = 1
        lng����ID = Val(Split(strmsg, MSG_SPLIT)(5))
        lng��ҳID = Val(Split(strmsg, MSG_SPLIT)(6))
    Else
        Exit Function
    End If
    
    '���鱨�����
    If blnLis Then
        If mobjLisInsideComm Is Nothing Then
            Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
            If mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    If errHandle("zlSoftShowHisForms.ProcessMessage", strErr) = 1 Then Resume
                    Exit Function
                End If
            End If
            If mobjLisInsideComm Is Nothing Then
                 MsgBox "LIS�ӿڳ�ʼ��ʧ��", vbInformation, "��ʾ"
                 Exit Function
            End If
        End If
        Call mobjLisInsideComm.PatientSampleBrowse(frmShowHisForms, lng����ID, "", 0, 0, IIf(lng��ҳID = 0, 1, 2), lng��ҳID)
    Else
    
        '���Ӳ�����ѯ
        If mclsArchive Is Nothing Then
            '��һ�ε��õ��Ӳ������ģ���ֵ
            Set mclsArchive = New clsArchive
        End If
        Call mclsArchive.zlOpenArchiveForm(lng����ID, lng��ҳID)
    End If

    ProcessMessage = 0
    Exit Function
err:
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
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
    
    strFile = App.Path & "\zlSoftShowArchiveView" & Format(Now, "YYYY_MM_DD") & ".Log"
    
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    strTmp = strInfo
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strTmp
    objText.Close
End Sub
