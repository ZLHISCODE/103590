Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '����̨
Public gcnOracle As ADODB.Connection     '�������ݿ�����
Public gobjRelogin As New clsRelogin  '������������Ķ���ʵ��
Public gobjWait As Object 'չʾ��ģ̬��������ʹ�����˳��Ķ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA(���ڷ���DBA_ROLE_PRIVS��ͼ������IO�ϸߣ���ʱû���ж�,����Ҫʹ��ʱ���ж�)����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�

Public gstrSystems As String

'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Public Enum Register
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

'---------------------------------------------------------------
'����ʱ�䣬�����ж�������Ļ�ĵȴ�ʱ��
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'����:������ؽ��̴����API����:2008-10-30 11:34:11:���˺�
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Const PROCESS_TERMINATE = &H1
Public gcll_His_PId As Collection        '�洢��صĽ�����Ϣ:array(��������,PID,���ڸ���),"K"+������

#Const SYS_TRYUSE = "��ʽ" '��ʽ/����
Private Sub SetAppBusyState()
'���������̶���δ�������ʱ���滻��ִ�������̹���ʱ�����ġ����������𡱶Ի���
On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "���������ڴ����������ĵȴ���"
    App.OleRequestPendingMsgText = "�������������������ĵȴ���"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Sub Main()
    Dim rsMenu As ADODB.Recordset
    Dim objRIS As Object
    Dim strStyle As String
    
    '����:ɱ�쳣����:2008-10-30(BUG:14365)
    Call zlKillHISPID
    Set gcnOracle = gobjRelogin.Login(0, CStr(Command()))
    If gcnOracle Is Nothing Then
        Set gobjRelogin = Nothing
        Exit Sub
    End If
    gstrDeptName = gobjRelogin.DeptName
    gstrDbUser = gobjRelogin.DBUser
    
    'д�뱾�����������EXE�ļ���
    Call SaveSetting("ZLSOFT", "����ȫ��", "ִ���ļ�", App.EXEName & ".exe")
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\�����ļ�"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
    
    
    On Error Resume Next
    Set objRIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    If Not objRIS Is Nothing Then
        Call objRIS.SaveDBConnectInfo(gobjRelogin.InputUser, gobjRelogin.InputPwd, gobjRelogin.ServerName, gobjRelogin.IsTransPwd)
    End If
    gstrSystems = gobjRelogin.Systems
    Call GetUserInfo(IIf(gobjRelogin.Systems = "REPORT", 0, Replace(gobjRelogin.Systems, "'", "")))
    
    '��ȡ��¼����
    gstrUserFlag = IIf(gobjRelogin.IsSysOwner, "01", "00")
    gstrStation = OS.ComputerName
    If gstrStation = "" Then
        gstrStation = "..."
    End If
    
    '-------------------------------------------------------------
    '�����˵�������
    '-------------------------------------------------------------
    Set rsMenu = MenuGranted(gobjRelogin.MenuGroup)
    If rsMenu.EOF Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Set gobjRelogin = Nothing
        Exit Sub
    End If
    '-------------------------------------------------------------
    '�����ٴ�������ͬ��ʣ��������ڰ�װ������ʱ������˽�е��ڽ���ģ��ʱ����
    '-------------------------------------------------------------
    '-------------------------------------------------------------
    'ѡ����ò�ͬ��񵼺�̨
    '-------------------------------------------------------------
    On Error Resume Next
    Err = 0
    
    strStyle = zlDatabase.GetPara("����̨", , , "zlBrw")
    Set ZlBrowerDll = CreateObject(strStyle & ".Cls" & Mid(strStyle, 3))
    If Err <> 0 Then
        If strStyle = "ZLBRW" Then
            MsgBox "����ʧ�ܣ������������ļ���ʧ�������°�װ��", vbInformation, gstrSysName
            Set gobjRelogin = Nothing
            Exit Sub
        Else
            Err = 0
            Set ZlBrowerDll = CreateObject("ZLBRW.ClsBrw")
            If Err <> 0 Then
                MsgBox "����ʧ�ܣ������������ļ���ʧ�������°�װ��", vbInformation, gstrSysName
                Set gobjRelogin = Nothing
                Exit Sub
            End If
        End If
    End If
    '��������ע������ֵ
    Call UpdateParameters
    '���������ֹ������ֹ
    Set gobjWait = frmSelClient
    Load gobjWait
    Call ZlBrowerDll.SetEnvironment(gstrSysName, gstrVersion, gstrAviPath, _
                          gstrUserFlag, gstrDbUser, glngUserId, _
                          gstrUserCode, gstrUserName, gstrUserAbbr, _
                          glngDeptId, gstrDeptCode, gstrDeptName, _
                          gstrStation, gstrMenuSys, CStr(Command()))
    Call ZlBrowerDll.InitBrower(gobjRelogin, gcnOracle, rsMenu)
End Sub

Public Function MenuGranted(ByVal strMenuGroup As String) As ADODB.Recordset
    '-------------------------------------------------------------
    '���ܣ�������Ȩʹ�ò���װ�Ĳ���������������Ȩʹ�õĲ˵�����
    '������ע����
    '-------------------------------------------------------------
    Dim ArrCommand
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim intCount As Integer
    Dim strSystems As String
    Dim BlnOnlySys As Boolean 'ֻ�б���ϵͳ
    Dim strSYS As String
    
    On Error GoTo errH
    BlnOnlySys = (gstrSystems = "REPORT")
    If BlnOnlySys Then
        strSystems = "'0'"
        strSYS = "0"
    Else
        strSystems = Replace(gstrSystems, "','", ",")
        strSYS = Replace(gstrSystems, "'", "")
    End If
    
    If strMenuGroup <> "" Then gstrMenuSys = strMenuGroup
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs = "" Then strObjs = "'Zl9Common'"
    strObjs = Replace(strObjs, "','", ",")
    If OS.IsDesinMode Then
        strSQL = "Select ���, ID As ���, Nvl(�ϼ�id, 0) As �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) as �̱���, ���, ˵��, Nvl(ģ��, 0) As ģ��, Nvl(ϵͳ, 0) As ϵͳ, " & _
                 "        Nvl(ͼ��, 0) As ͼ��, ����, Decode(Upper(RTrim(����)), 'ZL9REPORT', 1, 0) As ���� " & _
                 " From Table(Cast(ZLTOOLS.f_Reg_Menu([1], [2], [3]) As ZLTOOLS.t_Menu_Rowset)) " & _
                 " Union " & _
                 " Select A.���, A.ID, Nvl(�ϼ�id, 0) As �ϼ�, A.����, Decode(Nvl(A.�̱���,'��'),'��',A.����,A.�̱���) As �̱���, A.���, A.˵��, Nvl(A.ģ��, 0) As ģ��, " & _
                 "        Nvl(A.ϵͳ, 0) As ϵͳ, Nvl(ͼ��, 0) As ͼ��, C.����, Decode(C.����, 'ZL9REPORT', 1, 0) As ���� " & _
                 " From (Select Level As ���, ID, �ϼ�id, ����, �̱���, ���, ˵��, Nvl(ģ��,0) ģ��, ϵͳ, ͼ�� " & _
                 "        From zlMenus " & _
                 "        Where ��� = [1] And Nvl(ϵͳ, 0) IN(" & strSYS & ") " & _
                 "        Start With �ϼ�id Is Null " & _
                 "        Connect By Prior ID = �ϼ�id) A, " & _
                 "      (Select ϵͳ, Nvl(ģ��,0) ģ�� " & _
                 "        From zlMenus A " & _
                 "        Where ��� = [1] And Nvl(ϵͳ, 0) IN (" & strSYS & ") " & _
                 "        Minus " & _
                 "        Select ϵͳ * 100, ��� From Zlregfunc Where ϵͳ * 100 IN (" & strSYS & ")) B," & _
                 "      (select ϵͳ, Upper(RTrim(����)) as ����,��� From zlPrograms ) C " & _
                 " Where A.ϵͳ = B.ϵͳ And A.ģ�� = B.ģ�� And A.ģ�� = C.���(+) and A.ϵͳ = C.ϵͳ"

    Else
        strSQL = "SELECT ���, Id AS ���, Nvl(�ϼ�id, 0) AS �ϼ�, ����, Decode(Nvl(�̱���,'��'),'��',����,�̱���) As �̱���, ���, ˵��, Nvl(ģ��, 0) AS ģ��, Nvl(ϵͳ, 0) AS ϵͳ, " & _
                 "        Nvl(ͼ��, 0) AS ͼ��, ����, Decode(Upper(Rtrim(����)), 'ZL9REPORT', 1, 0) AS ���� " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu([1], [2], [3]) As " & _
                 " Zltools.t_Menu_Rowset)) "
    End If
    'ʵ�ֱ����������,ģ��ſ�����zlReports.����id,Ҳ������zlRPTGroups.����id,����zlReports
    'ֻ��ȡ��������ģ��ı���
    strSQL = "Select ���, ���, �ϼ�, ����, �̱���, ���, ˵��, ģ��, ϵͳ, ͼ��, ����, ����, ������, �Ƿ�ͣ��" & vbNewLine & _
                    "From (Select a.*, Decode(a.����, 0, Null, Nvl(b.���, c.���)) ������, Nvl(b.�Ƿ�ͣ��, 0) + Nvl(c.�Ƿ�ͣ��, 0) �Ƿ�ͣ��" & vbNewLine & _
                    "       From (" & strSQL & ")  a," & vbNewLine & _
                    "            (Select Nvl(b.ϵͳ, 0) ϵͳ, b.����id, b.���, b.�Ƿ�ͣ��" & vbNewLine & _
                    "              From Zlprograms a, Zlreports b" & vbNewLine & _
                    "              Where Nvl(a.ϵͳ, 0) = Nvl(b.ϵͳ, 0) And a.��� = Nvl(b.����id, 0) And Upper(a.����) = 'ZL9REPORT') b, (Select ���, Nvl(ϵͳ, 0) ϵͳ, ����id, �Ƿ�ͣ�� From Zlrptgroups) c" & vbNewLine & _
                    "       Where a.ϵͳ = b.ϵͳ(+) And a.ģ�� = b.����id(+) And a.ϵͳ = c.ϵͳ(+) And a.ģ�� = c.����id(+))" & vbNewLine & _
                    "Order By ���, ����, ϵͳ, ģ��, ���, ������"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, gstrMenuSys, Replace(strSystems, "'", ""), Replace(strObjs, "'", ""))

    Set MenuGranted = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As ADODB.Recordset, rsUser As ADODB.Recordset
    Dim strSQL As String, i As Integer
    On Error GoTo errH
    '���û���Ϣ���蹫����������������ʹ��
    strSQL = "Select S.������" & _
            " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
            " Where S.������=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������")
    
    With rsTmp
        If Not .EOF Then
            '��Ϊ���ܸ��û����ж��ϵͳ����ݣ�����ѭ��ȡ���
            glngUserId = 0 '��ǰ�û�id
            gstrUserCode = "" '��ǰ�û�����
            gstrUserName = "" '��ǰ�û�����
            gstrUserAbbr = "" '��ǰ�û�����
            glngDeptId = 0 '��ǰ�û�����id
            gstrDeptCode = "" '��ǰ�û�
            gstrDeptName = "" '��ǰ�û�
            
            For i = 1 To .RecordCount
                strSQL = "Select R.��ԱID,R.����ID,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
                        " From " & !������ & ".�ϻ���Ա�� U," & !������ & ".��Ա�� P," & !������ & ".���ű� D," & !������ & ".������Ա R" & _
                        " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=[1] And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
                Set rsUser = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա��Ϣ", gstrDbUser)
                                
                If Not rsUser.EOF Then
                    glngUserId = rsUser!��ԱID '��ǰ�û�id
                    gstrUserCode = rsUser!��� '��ǰ�û�����
                    gstrUserName = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    gstrUserAbbr = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    glngDeptId = rsUser!����ID '��ǰ�û�����id
                    gstrDeptCode = rsUser!���ű��� '��ǰ�û�
                    gstrDeptName = rsUser!�������� '��ǰ�û�
                    Exit For
                End If
                .MoveNext
            Next
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'**********************************************************************************************************************
'����:���´�����ؽ��̵ĺ���
'����:���˺�
'����:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ɱ������HIS����������Ƴ�����(ɱ��������:����ZLHIS+.exe�Ľ��������κδ���)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo errHand:
    '��һ��:��Ҫ������ص�ZLHIS����ؽ���
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '���������صĴ��󣬾�ֱ�ӷ�����
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '�ڶ���:��Ҫ�������ZLHIS����ؽ��̵���ش��ڸ���,�����ź��жϳ���صĽ����Ƿ�����쳣,�����쳣�ģ��͵�ɱ��
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '�϶�������С��1����,��ô�϶����쳣����Ҫɱ����
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '����δ�ɹ������޴���������
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
errHand:
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���д��ڷ���HIS�Ľ��̵Ĵ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hwnd) = 0 Then
        '��ȡ hWnd ���Ӵ�����
        strTittle = String(80, 0)
        Call GetWindowText(hwnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hwnd, lngPID)
            If IsWindowVisible(hwnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' ��ʾ�����о� hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���Ľ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo errHand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
errHand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡZLHIS�Ľ��̸���صļ���(gcll_HIS_Pid)
    '���:
    '����:cll_His_Pid-������HIS.exe�ĳ���װ�ظü�����
    '����:
    '����:���˺�
    '����:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strExeName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Dim StrSessionID As String '��ǰ�ỰID
    Dim StrHISSessionID As String '����ZLHIS���̻ỰID
    Const TH32CS_SNAPPROCESS = &H2
    
    
    Err = 0: On Error GoTo errHand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '��ȡ��ǰӦ�ó������
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    StrSessionID = GetCurSessionID(lngCurPid)
    
    
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '�����ڵ�ǰ���̵ĲŴ���
            If lngCurPid <> uProcess.lProcessId Then
                strExeName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strExeName Like strCurExeName Then '"ZLHIS+.EXE"
                    StrHISSessionID = GetCurSessionID(uProcess.lProcessId)
                    '�����ǰzlhis+�Ľ��̻ỰID�������ĻỰID��ͬ,�Ž��йرմ���
                    If StrSessionID = StrHISSessionID Then
                        cll_His_Pid.Add Array(strExeName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                    End If
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
errHand:
End Function

Private Function GetCurSessionID(ByVal lngCurPid As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ���̵ĻỰID
    '���:��ǰ����PID
    '����:
    '����:�ỰID
    '����:ף��
    '����:2012-06-06 10:15:00
    '-----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Dim WMI, objProcess, colProcessList As Object
    Set WMI = GetObject("WinMgmts:")
    Set colProcessList = WMI.InstancesOf("Win32_Process")
    For Each objProcess In colProcessList
        If objProcess.Handle = lngCurPid Then
            GetCurSessionID = objProcess.SessionId
            Exit Function
        End If
    Next
    GetCurSessionID = "-1"
End Function

