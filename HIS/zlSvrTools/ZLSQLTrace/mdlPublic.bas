Attribute VB_Name = "mdlPublic"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public gobjFile As New FileSystemObject
Public gstrFilePath As String
Public gcnOracle As adodb.Connection    '�������ݿ�����
Public gstrDBUser As String
Public gblnOwner As Boolean
Public gcolSort As Collection
Public gfrmFind As New frmFind
Public gblnIsRac As Boolean
Public gintInstId As Integer
Public gblnZlhis As Boolean
Public gstrCompareExe As String
Public gstrLeft As String
Public gstrSysName As String                'ϵͳ����
Public gstrUserName As String               '�û���
Public gstrPassword As String               '�û�����
Public gstrToolsPwd As String               '�����ߵ�����
Public gstrServer As String                 '��������
Public gstrSQL    As String                 'ͨ�õ�SQL������
Public gblnDBA As Boolean                   '�Ƿ�DBA
Public gdtStart As Long
Public gblnOK As Boolean
Public glngSessionID As Long

Public gblnHasZltables As Boolean '��¼�Ƿ���zltable���ű�

'********************************************************************
'CommandBar����ID
Public Enum CommandBarIDCond
    conMenu_FilePopup = 1
    conMenu_EditPopup = 2
    conMenu_ViewPopup = 8
    conMenu_HelpPopup = 9
    
    '���һ���Աȹ�������
    conMenu_ComparePopup = 3
    '�ļ��˵�
    conMenu_File_Open = 101
    conMenu_File_CompareExe = 210
    conmenu_File_Logout = 108
    conMenu_File_Exit = 109
    
    '�༭�˵�
    conMenu_Edit_Trace = 201
    conMenu_Edit_Trace_1 = 2011
    conMenu_Edit_Trace_4 = 2012
    conMenu_Edit_Trace_8 = 2013
    conMenu_Edit_Trace_12 = 2014
    conMenu_Edit_ChangeReg = 2015
    conMenu_Edit_TraceOff = 202
    conMenu_Edit_CompareLeft = 211
    conMenu_Edit_Compare = 212
    
    '�鿴�˵�
    conMenu_View_Style = 801
    conMenu_View_Style_Report = 8011
    conMenu_View_Style_Table = 8012
    conMenu_View_Filter = 802
    conMenu_View_SQLPrev = 803
    conMenu_View_SQLNext = 804
    conMenu_View_Find = 805
    conMenu_View_FindNext = 806
    conMenu_View_Refresh = 809
    conMenu_View_Close = 810
    
    '�����˵�
    conMenu_Help_About = 901
End Enum

'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000 '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392 '�������б�
Public Const ID_INDICATOR_CAPS = 59137 '״̬������д��
Public Const ID_INDICATOR_NUM = 59138 '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139 '״̬����������

'CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************
Public Const CB_SETDROPPEDWIDTH As Long = &H160
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'-------------------------------------------------------------
Public Const Process_Query_Information = &H400
Public Const Still_Active = &H103
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'-------------------------------------------------------------
Public Const GWL_EXSTYLE = (-20)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'-------------------------------------------------------------
Public Const EM_LINESCROLL = &HB6 'lngW=��������,lngL=��������
Public Const EM_SCROLL = &HB5 '������������
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETLINECOUNT = &HBA 'lngR(>=1,�����Զ��۵���)
Public Const EM_LINELENGTH = &HC1 '��һ��δ����ǰ��Ч
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETSEL = &HB1

Public Const FR_DOWN = &H1
Public Const FR_WHOLEWORD = &H2
Public Const FR_MATCHCASE = &H4
Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Public Type FINDTEXT
    chrg As CHARRANGE
    lpstrText As String
End Type

Public Const WM_USER = &H400
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
'-------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode���ս��ַ���
Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Const REG_DWORD = 4                      ' 32-bit ����

' ע���ؼ��ְ�ȫѡ��...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

' ����ֵ...
Public Const ERROR_SUCCESS = 0
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long


Public Sub Main()
    Dim strTmp As String
    Dim strServerName As String, strUserName As String, strUserPwd As String
    Dim intUserPosition As Integer, intPwdPosition As Integer, intServerPosition As Integer
    
    Call InitCommonControls
    
    gblnOwner = False
    strTmp = Command
    
    If strTmp = "" Then
        '�û�ע��
        frmUserLogin.Show 1
        If gcnOracle Is Nothing Then
            Set gcnOracle = New adodb.Connection
        End If
    Else
        '�ڹ������У�ͨ��Command������е�¼
        intUserPosition = InStr(1, strTmp, "zlUserName=") + Len("zlUserName=")
        intPwdPosition = InStr(1, strTmp, "zlPassword=") + Len("zlPassword=")
        intServerPosition = InStr(1, strTmp, "zlServer=") + Len("zlServer=")
        
        strUserName = Mid(Left(strTmp, InStr(1, strTmp, "zlPassword=") - 1), intUserPosition)
        strUserPwd = Mid(Left(strTmp, InStr(1, strTmp, "zlServer=") - 1), intPwdPosition)
        strServerName = Mid(strTmp, intServerPosition)
        gstrDBUser = UCase(strUserName)
        
        If Not OraDataOpen(strServerName, strUserName, strUserPwd) Then
            Exit Sub
        End If
    End If

    frmMain.Show
End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnת�� As Boolean) As Boolean
    Dim rstmp As adodb.Recordset
    Dim strSql As String, i As Integer
    
    On Error Resume Next
    
    If gcnOracle Is Nothing Then
        Set gcnOracle = New adodb.Connection
    End If
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    With gcnOracle
        .CursorLocation = adUseClient
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        MsgBox "����ʧ�ܣ�����ȷ���û�����������������", vbInformation, App.Title
        Err.Clear: Exit Function
    End If

    With rstmp
        strSql = "Select 1 From User_Role_Privs Where Granted_Role = 'DBA'"
        If .State = adStateOpen Then .Close
        .Open strSql, gcnOracle, adOpenKeyset
        gblnDBA = Not (.EOF Or .BOF)
    End With

    '���ܣ�����Ƿ�ΪRAC����
    Err.Clear
    strSql = "Select 1 from gv$active_instances"
    Set rstmp = OpenSQLRecord(strSql, "CheckRAC")
    gblnIsRac = rstmp.RecordCount > 0
    
    If gblnIsRac Then
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rstmp = OpenSQLRecord(strSql, "CheckRAC")
        gintInstId = Val("" & rstmp!INST_ID)
    End If
    
    If Err.Number > 0 Then Exit Function
    
    gstrUserName = strUserName: gstrPassword = strUserPwd: gstrServer = strServerName
    OraDataOpen = True
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

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetShortName(ByVal strFile As String) As String
    Dim strShort As String, lngLen As Long
    
    GetShortName = strFile
    
    If InStr(strFile, " ") > 0 Then
        If gobjFile.FileExists(strFile) Then
            GetShortName = gobjFile.GetFile(strFile).ShortPath
        ElseIf gobjFile.FolderExists(strFile) Then
            GetShortName = gobjFile.GetFolder(strFile).ShortPath
        Else
            strShort = Space(255)
            lngLen = GetShortPathName(strFile, strShort, 255)
            GetShortName = Left(strShort, lngLen)
        End If
    End If
End Function

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'���ܣ���ComboBoxʵ������������Զ���ɵĹ���
'˵������Combox.KeyPress�¼��е���
    Dim strInput As String
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.Style <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput

        '���ݼ�������ݵõ����ܵ��б���
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            .ListIndex = lngIndex
            '.Text = .List(lngIndex)
            
            .SelStart = Len(strInput)
            .SelLength = Len(.Text) - Len(strInput)
        Else
            .Text = strInput
            .SelStart = Len(strInput)
        End If
    End With
End Sub


Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As adodb.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(gcnOracle, strSql, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal cnOracle As adodb.Connection, ByVal strSql As String, ByVal strTitle As String, arrInput() As Variant) As adodb.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New adodb.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    '�����Զ���[x]����
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
    
    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = cnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If

    cmdData.CommandText = strSql
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing

End Function



Public Function CheckZlhis() As Boolean
    Dim strSql As String, rstmp As adodb.Recordset
    
    On Error GoTo errh
    
    strSql = "Select 1 From dba_tables Where table_name = 'ZLSYSTEMS'"
    Set rstmp = OpenSQLRecord(strSql, "CheckZlhis")
    
    CheckZlhis = rstmp.RecordCount > 0
    Exit Function
errh:
    MsgBox "��ȡZLHIS����ʧ�ܡ�"
    CheckZlhis = False
End Function

Public Sub CheckSqlPlan(vsfPlanTbl As VSFlexGrid, ByVal intOptCol As Integer, ByVal intObjCol As Integer, _
                                            rsBigtbl As adodb.Recordset, rsBigIdx As adodb.Recordset, rsLowIdx As adodb.Recordset)
'����:���VSF����е�ִ�мƻ�
'         1.���ȫ��ɨ��zltables+zlbigtable+zlbaktables��
'         2.���ͱ�ȫ��ɨ��(�����ͳ����Ϣ��User_tab_statistics:num_rows>3000(ҩƷĿ¼һ�������ֵ����) AND num_rows<100 0000��������)
'         3.��������û�����(�Ǵ��)������ϵ�����
'         4.�������ͱ�����ȫɨ�裨inex full scan��INDEX FAST FULL SCAN��
'         5.�������ͱ���Ծʽ����ɨ�裨INDEX SKIP SCAN��
'����:
'vsfPlanTbl - ִ�мƻ����
'intOptCol - ������,��:Index full scan ,intObjCol - �����漰�Ķ�����,��: ����ҽ����¼_IX_ID
'rsBigtbl,rsBigIdx,rsLowIdx -�漰�ı�/����
    
    Dim strOperation As String, strObject As String
    Dim strTmp() As String, i As Integer, j As Integer
    Dim blnTmp As Boolean
    
    On Error GoTo errh
    With vsfPlanTbl
        If .Redraw = flexRDNone Then Exit Sub
        
        '�������,��ȡ����
        For i = .FixedRows To .Rows - .FixedRows
            If intOptCol <> intObjCol Then
                'ִ�мƻ��Ĳ����Ͷ�����һ����,ֱ�ӻ�ȡ
                strOperation = TrimEx(.TextMatrix(i, intOptCol))
                strObject = TrimEx(.TextMatrix(i, intObjCol))
            Else
                '�漰���:TABLE ACCESS FULL/INDEX FAST FULL SCAN/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN
                strTmp = Split("TABLE ACCESS FULL/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN/INDEX FAST FULL SCAN", "/")
                
                For j = 0 To UBound(strTmp)
                    If InStr(1, TrimEx(.TextMatrix(i, intOptCol)), strTmp(j)) > 0 Then
                        strOperation = strTmp(j)
                        strObject = Split(Trim(Replace(TrimEx(.TextMatrix(i, intOptCol)), strTmp(j), "")), " ")(0)
                        Exit For
                    End If
                Next
            End If
            
            If strOperation <> "" And strObject <> "" Then
                If strOperation = "TABLE ACCESS FULL" Then '��ȡȫ��ɨ��
                    blnTmp = CheckRs(rsBigtbl, "���� = '" & strObject & "'") Or gcnOracle = ""
                ElseIf InStr(1, "INDEX FULL SCAN/INDEX SKIP SCAN/INDEX FAST FULL SCAN", strOperation) > 0 Then '����ȫɨ��\������ɨ��
                    blnTmp = CheckRs(rsBigIdx, "������ = '" & strObject & "'") Or gcnOracle = ""
                ElseIf strOperation = "INDEX RANGE SCAN" And gcnOracle <> "" Then '������Χɨ��:��Ч����
                    blnTmp = CheckRs(rsLowIdx, "Լ����= '" & GetFkByIdx(strObject) & "'")
                End If
            End If
                
            If blnTmp Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HF0F0FF
            strOperation = "": strObject = ""
            blnTmp = False
        Next

    End With
    Exit Sub
errh:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub


Public Sub GetMidTabSize(ByRef lngMinSize As Long, ByRef lngMaxSize As Long)
    '����:��ȡ���ͱ��С
    
    Dim strSql As String, rstmp As adodb.Recordset
    
    lngMinSize = 3000: lngMaxSize = 1000000
    
    On Error GoTo errh
    strSql = "Select A.������,Nvl(A.����ֵ,A.ȱʡֵ) As ����ֵ " & _
                 "From zlParameters A " & _
                 "Where A.������ = '������ͱ�' And a.ϵͳ is null And a.ģ�� is null"
    Set rstmp = OpenSQLRecord(strSql, "GetMidTabSize")
    
    If rstmp.EOF Then Exit Sub
    lngMinSize = Split(rstmp!����ֵ, ",")(0)
    lngMaxSize = Split(rstmp!����ֵ, ",")(1)
    
    Exit Sub
errh:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCheckObj(ByVal intMod As Integer, Optional ByVal lngMinSize As Long, Optional ByVal lngMaxSize As Long) As adodb.Recordset
'����:��ȡ�漰��������ı�/��������,����һ����¼��
'����intMod: 1-��,2-����,3-��Ч����
'lngMinSize,lngMaxSize - �ж����ͱ������,������Ĭ��Ϊ3000-1000000

    Dim strSql As String
    
    On Error GoTo errh
    
    If gblnHasZltables Then
        strSql = "Union Select Distinct ���� From Zltables Where ���� In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')"
    Else
        strSql = "Union Select Distinct ���� From Zlbigtables" & vbNewLine & _
                        "Union" & vbNewLine & _
                        "Select Distinct ���� From zlBakTables"
    End If
    
    Select Case intMod
        Case 1
            strSql = "Select distinct  Table_Name ����" & vbNewLine & _
                            "From Dba_Tab_Statistics" & vbNewLine & _
                            "Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql
                            
        Case 2
            strSql = "Select distinct Index_Name ������" & vbNewLine & _
                            "From Dba_Indexes" & vbNewLine & _
                            "Where Table_Name In" & vbNewLine & _
                            " ( Select Table_Name ���� From Dba_Tab_Statistics Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql & ")"

        Case 3
            strSql = "Select distinct  a.Constraint_Name Լ����" & vbNewLine & _
                            "From Dba_Constraints A, Dba_Indexes B" & vbNewLine & _
                            "Where a.Constraint_Type = 'R' And b.uniqueness='UNIQUE' And a.r_Constraint_Name = b.Index_Name And a.r_Owner = b.Owner And" & vbNewLine & _
                            "      b.Table_Name Not In" & vbNewLine & _
                            "      (Select Distinct ���� From Zlbigtables" & vbNewLine & _
                            "       Union Select Distinct ���� From zlBakTables" & vbNewLine & _
                            IIf(gblnHasZltables, "Union Select Distinct ���� From Zltables Where ���� In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')", "") & vbNewLine & _
                            "       )"

    End Select
    
    Set GetCheckObj = OpenSQLRecord(strSql, "GetCheckObj")
    Exit Function
errh:
    Set GetCheckObj = Nothing
    MsgBox Err.Description
End Function


Public Function CheckRs(rsData As adodb.Recordset, ByVal strFilter As String) As Boolean
'����:�Դ���ļ�¼����ӹ���,�����ƥ�����򷵻�True
    
    If rsData Is Nothing Then Exit Function
    rsData.Filter = strFilter
    CheckRs = Not rsData.EOF
    rsData.Filter = 0
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function

Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '���ܣ����ݱ����жϱ��Ƿ����
    '������strTableName - Ҫ��ѯ�ı���
    Dim strSql As String, rsData As adodb.Recordset
    
    On Error GoTo errh
    strSql = "select 1 from dba_all_tables where table_name =[1] "
    Set rsData = OpenSQLRecord(strSql, "CheckTblExist", strTableName)
    CheckTblExist = (rsData.RecordCount > 0)
    
    Exit Function
errh:
    MsgBox Err.Description
End Function

Public Function GetFkByIdx(ByVal strIdxName As String) As String
'����:���ݴ�����������ض�Ӧ�����Լ������
    
    Dim strSql As String, rsData As adodb.Recordset
    
    On Error GoTo errh:
    
    strSql = "Select Distinct a.Constraint_Name" & vbNewLine & _
                    "From Dba_Cons_Columns A, Dba_Ind_Columns B" & vbNewLine & _
                    "Where a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name And a.Position = b.Column_Position And" & vbNewLine & _
                    "      b.Index_Name = [1]"

    Set rsData = OpenSQLRecord(strSql, "GetFkByIdx", strIdxName)
        
    If Not rsData.EOF Then
        GetFkByIdx = rsData!Constraint_Name & ""
    End If
    Exit Function
errh:
    GetFkByIdx = ""
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Function
