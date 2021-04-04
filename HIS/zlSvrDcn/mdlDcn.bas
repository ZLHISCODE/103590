Attribute VB_Name = "mdlDcn"
Option Explicit
Private Const GWL_WNDPROC = -4
Public Const GWL_USERDATA = (-21)
Public Const WM_SIZE = &H5
Public Const WM_USER = &H400
Public Const WM_BROADCAST = &H218
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function OCI_ConnCreate Lib "zlNoticeLib.dll" (ByVal strServer As String, ByVal strUser As String, ByVal strPwd As String) As Boolean
Public Declare Sub OCI_Register Lib "zlNoticeLib.dll" (ByVal lngHandler As Long, ByVal strTable As String)
Public Declare Sub OCI_UnRigister Lib "zlNoticeLib.dll" ()

Public lpPrevWndProc As Long    '������
Public gcolNotice As New Collection '֪ͨ���漯��
Public gstrBuild           As New clsStringBulider

Public gstrIp As String
Public glngPort As Long
Public glngSid As Long
Public gintState As Integer
Public gintLog As Integer
Public gintInterval As Integer

Public Function Hook(ByVal hwnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    lpPrevWndProc = GetWindowLong(hwnd, GWL_WNDPROC)
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf WindowProc
    
    Hook = lpPrevWndProc
End Function

Public Sub UnHook(ByVal hwnd As Long)
    Dim temp As Long
    'Cease subclassing.
    temp = SetWindowLong(hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim byteRowid(100) As Byte, strNotice As String
    '����ԭ���Ĵ��ڹ���
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    
    If uMsg = WM_USER + 1 Then
        On Error Resume Next
        
        CopyMemory byteRowid(0), ByVal wParam, 100
        strNotice = StrConv(byteRowid, vbUnicode)
        strNotice = zlCommFun.TruncZero(strNotice)
        gcolNotice.Add strNotice
    End If
    
End Function


'--------------------------------------------------------------------------------------------------------
'���ݲ���
Public Function GetNoticeList() As ADODB.Recordset
    '����:��ȡNoticeList,������һ����¼��
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Noticecode, Noticename, Tableowner, Tablename, Receivercols,  Changetype," & vbNewLine & _
                    "SplitChar,Noticekind, Comments, Status ,Filter,ReceiverTab ,ReceiverRelas,ReceiverIP ," & vbNewLine & _
                    "ReceiverStaffKind ,ReceiverDeptKind,Interval " & vbNewLine & _
                    "From Zltools.Zlnoticelists Order by 1"
                        
    Set GetNoticeList = zlDatabase.OpenSQLRecord(strSql, "��ȡNoticeList")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetClientPort() As ADODB.Recordset
    '����:��ȡ�ͻ���UDP�˿�����
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ����ֵ From zloptions Where ������=9"
                        
    Set GetClientPort = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ͻ���UDP�˿�")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetJobs() As ADODB.Recordset
    Dim strSql  As String
    
    On Error GoTo errH
    strSql = "Select ���ʱ��, Nvl(��ҵ��, 0) ��ҵ�� From zlAutoJobs Where ���� = 3 And ��� = 3 And ϵͳ Is Null"
                        
    Set GetJobs = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ͻ���UDP�˿�")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetUserNotices() As ADODB.Recordset
    '����:��ȡ�Զ���֪ͨ,������һ����¼��
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.���, a.��������, a.�������, a.��������, To_Char(a.��ʼʱ��, 'yyyy/mm/dd hh24:mi') ��ʼʱ��," & vbNewLine & _
                "       To_Char(a.��ֹʱ��, 'yyyy/mm/dd hh24:mi') ��ֹʱ��" & vbNewLine & _
                "From zlNotices A" & vbNewLine & _
                "Order By 1"

    Set GetUserNotices = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Զ���֪ͨ")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetSid() As Long
    '����:��ȡ��ǰ�Ự��Sid
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select Userenv('SessionID') AudSid From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡSID")
    GetSid = Val(rsTmp!AudSid & "")
    
    Exit Function
errH:
    ErrCenter
End Function

Public Function CheckSidState(ByVal lngSid As Long) As Boolean
    '���ָ����Sid�Ƿ�����
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1 From GV$session Where Audsid = [1] And STATUS <> 'KILLED'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡSID", lngSid)
    CheckSidState = rsTmp.RecordCount <> 0
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateDcnState2DB(ByVal intType As Integer) As Boolean
    '�����ݿ����޸���Ϣ�շ���״̬
    'intType=1:����     intType=0:����
    Dim strSql As String, rsClient As ADODB.Recordset
    Dim strValue As String
    On Error GoTo errH
    
    strValue = gstrIp & ";" & glngPort & ";" & intType & ";" & glngSid
    
    '���ڵ��ò���ʱ,�Ѿ���zlclientSession���в����˵�ǰ�Ự��Ϣ,����ȸ��ݻỰ��ɾ��
    
    If intType = 1 Then
        strSql = "Delete From zltools.zlclientsession Where �Ự�� = " & glngSid
        gcnOracle.Execute strSql
    End If
    
    '����շ���״̬
    strSql = "SELECT ����ֵ FROM zltools.zloptions WHERE ������=[1]"
    Set rsClient = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", 27)
    
    If rsClient.RecordCount > 0 Then
        If rsClient!����ֵ & "" <> "" Then
            If intType = 1 And Split(rsClient!����ֵ & "", ";")(2) = 1 And CheckSidState(Split(rsClient!����ֵ & "", ";")(3)) Then
                MsgBox "���ݱ䶯֪ͨ�����Ѿ���IP" & Split(rsClient!����ֵ & "", ";")(0) & "�ѿ������޷��ٴο�����"
                Exit Function
            End If
        End If
        strSql = "Update Zltools.zlOptions Set ����ֵ = '" & strValue & "' Where ������ =27"
        gcnOracle.Execute strSql
    Else
        MsgBox "��������ȱʧ������zlOption�е�27�Ų����Ƿ���ڡ�", vbExclamation, "ע��"
        Exit Function
    End If
    
    UpdateDcnState2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeServerSet2DB(ByVal lngPort As Integer, ByVal intLog As Integer, ByVal intInterval As Integer) As Boolean
    '����:�޸�DCN������������
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    
    On Error GoTo errH
    glngPort = lngPort: gintLog = intLog: gintInterval = intInterval
    strValue = gstrIp & ";" & glngPort & ";" & gintState & ";" & glngSid
        
    strSql = "Update Zltools.zlOptions Set ����ֵ = '" & strValue & "' Where ������ = 27"
    gcnOracle.Execute strSql
    
    strSql = "Update Zltools.zlOptions Set ����ֵ = " & intLog & " Where ������ = 28"
    gcnOracle.Execute strSql
    
    strSql = "Update Zltools.zlOptions Set ����ֵ = " & intInterval & " Where ������ = 29"
    gcnOracle.Execute strSql
    
    ChangeServerSet2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeClientSet2DB(ByVal lngPortS As Long, ByVal lngPortE As Long, ByVal lngCheckInterval As Long) As Boolean
    '����:�޸Ŀͻ�����Ϣ���ն˿ںͼ����״̬Ƶ��
    Dim strSql As String
    
    On Error GoTo errH
    
    '�޸Ķ˿�
    If lngPortS = 0 Or lngPortE = 0 Then
        strSql = "Update zloptions set ����ֵ = 0 Where ������ = 9"
    Else
        strSql = "Update zloptions set ����ֵ = '" & lngPortS & "-" & lngPortE & "' Where ������ = 9"
    End If
    gcnOracle.Execute strSql
        
    '�޸ļ����״̬��Ƶ��
    strSql = "Update zloptions set ����ֵ = " & lngCheckInterval & " Where ������ = 32"
    gcnOracle.Execute strSql
    
    ChangeClientSet2DB = True

    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeJobSet2DB(ByVal intType As Integer, Optional ByVal intInterval As Integer = 5) As Boolean
    '����:�޸��Զ���ҵ��Ϣ
    'intType =1: �ύ�Զ�����  intType =2:�޸��Զ����� intType =3:ɾ���Զ�����
    'intInterval-�Զ�����ִ��Ƶ��,Ĭ��ÿ5����ִ��һ��
    Dim strSql As String
    
    On Error GoTo errH
    Select Case intType
    Case 1
        strSql = "Begin" & vbNewLine & _
                    "  Execute Immediate 'Update  zlAutoJobs Set ���ʱ�� = " & intInterval & "  Where ���� = 3 And ��� = 3 And ϵͳ Is Null';" & vbNewLine & _
                    "  zltools.Zl_Jobsubmit(Null,3,3);" & vbNewLine & _
                    "End;"
    Case 2
        strSql = "Begin" & vbNewLine & _
                    "  Execute Immediate 'Update  zlAutoJobs Set ���ʱ�� = " & intInterval & "  Where ���� = 3 And ��� = 3 And ϵͳ Is Null';" & vbNewLine & _
                    "  zltools.Zl_Jobchange(Null,3,3);" & vbNewLine & _
                    "End;"
    Case 3
        strSql = "Begin" & vbNewLine & _
                    "  zltools.Zl_Jobremove(Null,3,3);" & vbNewLine & _
                    "End;"
    End Select
    
    gcnZltools.Execute strSql
    ChangeJobSet2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Private Function GetZltoolsConnection(ByVal strPwd As String) As Boolean
    '����: ��ȡzltools���Ӷ���
    
    On Error Resume Next
    
    If gcnZltools Is Nothing Then
        Set gcnZltools = New ADODB.Connection
    End If
    
    With gcnZltools
        .Provider = "OraOLEDB.Oracle"
        .Open "PLSQLRSet=1;Data Source=" & gstrServer, "ZLTOOLS", strPwd
      
        If .State = adStateOpen Then
            GetZltoolsConnection = True
        End If
    End With

End Function

Public Function GetZltools() As Boolean
    Dim blnResult As Boolean, strPwd As String
    
    blnResult = GetZltoolsConnection("ZLTOOLS")
    
    If Not blnResult Then
        blnResult = GetZltoolsConnection("ZLSOFT")
    End If
    
    If Not blnResult Then
        blnResult = frmUserCheckLogin.GetZltoolsByLogin
    End If
    
    GetZltools = blnResult
End Function

Public Function GetCheckInterval() As Long
    '��ȡ "DCN���ʱ����¼��"
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select ����ֵ From zltools.zlOptions Where ������=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "DCN���ʱ����¼��", 32)
    
    GetCheckInterval = Val(rsTmp!����ֵ)
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateDcnTime() As Boolean
    '����:ά�����������´��ʱ��
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1" & vbNewLine & _
                "From Dba_Change_Notification_Regs" & vbNewLine & _
                "Where Table_Name In (Select Tableowner || '.' || Tablename From Zlnoticelists)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���DCN״̬")
    
    If rsTmp.RecordCount > 0 Then
        strSql = "Update zloptions set ����ֵ = to_char(Sysdate,'YYYY-MM-DD hh24:mi:ss') where ������=31"
        gcnOracle.Execute strSql
    End If
    
    UpdateDcnTime = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateNoticeInterval(ByVal lngNoticeCode As Long, ByVal lngInterval As Long)
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Update zlNoticeLists set Interval= " & lngInterval & " where NoticeCode=" & lngNoticeCode
    gcnOracle.Execute strSql
        
    UpdateNoticeInterval = True
    Exit Function
errH:
    ErrCenter
End Function
