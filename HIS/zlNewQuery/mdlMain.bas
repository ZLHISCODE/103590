Attribute VB_Name = "mdlMain"
Option Explicit
'Public gobjDemand As Object                '����̨
Public SplashObj As New frmSplash
Public gcnOracle As ADODB.Connection     '�������ݿ�����

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�
Public gobjLogin As Object
Public gobjRegister As Object

'-----------------------------------------
'�����롢ע���롢�������������ע���������
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------

Public gstrSystems As String

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public glngOld As Long, glngFormW As Long, glngFormH As Long

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String
    Dim IntCount As Integer
    Dim StrStyle As String
    Dim rsMenu As ADODB.Recordset
    Dim StrHaveSys As String
    
    If gobjLogin Is Nothing Then
        Set gobjLogin = CreateObject("zlLogin.clsLogin")
    End If
    If gobjLogin Is Nothing Then
        Err = 0: On Error GoTo 0
        MsgBox "����zlLogin����ʧ��,����zlLogin�ļ���ʧ�����Ƿ���ȷע�ᣡ", vbInformation + vbOKOnly, "��½��֤"
        Exit Sub
    End If
    Set gcnOracle = gobjLogin.Login(0, CStr(Command()), , , App.Path & "\" & App.EXEName & ".exe", App.hInstance)
    If gcnOracle Is Nothing Then
        Set gobjLogin = Nothing
        Exit Sub
    End If
    gstrSystems = gobjLogin.Systems
    gstrServerName = gobjLogin.ServerName
    gstrDbUser = gobjLogin.DBUser

    '��ʼ����������
    InitCommon gcnOracle
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gstrParsePublish = zlRegInfo("��Ʒ����")
    gstrParseRegCode = zlRegInfo("��λ����", , -1)
    gstrSysName = gstrParsePublish & "���"
    
    StrHaveSys = gstrSystems
    If gstrSystems = "REPORT" Then
        gstrSystems = ""
    Else
        gstrSystems = " (ϵͳ in (" & gstrSystems & ") Or ϵͳ Is NULL)"
    End If
    If gstrSystems = "" Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '�����˵�������
    '-------------------------------------------------------------
    gstrSQL = "SELECT ϵͳ FROM zlPrograms WHERE ���=1536 AND ϵͳ IN (" & StrHaveSys & ")"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMain")
    
    If gRs.BOF Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngSys = gRs("ϵͳ").Value
    If InStr(1, GetPrivFunc(glngSys, 1536), "����") <= 0 Then
        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    '����ͬ���
    '-------------------------------------------------------------
    Call CreateSynonyms(glngSys, 1536)
    
    Call GetUserInfo
    Call CodeMan(glngSys, 1536)
    
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    zlDatabase.ExecuteProcedure strSQL, "����ͬ���"
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
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
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
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(strUserName)
    gstrServerName = strServerName
    SetDbUser gstrDbUser
    
    gstrConnect = strServerName & ";" & strUserName & ";" & strUserPwd
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
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

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDbUser
    UserInfo.���� = gstrDbUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.������ = IIf(IsNull(rsTmp!������), "", rsTmp!������)
        UserInfo.���� = IIf(IsNull(rsTmp!������), "", rsTmp!������)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CodeMan(lngSys As Long, ByVal lngModul As Long)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������

    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
       
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    
    glngSys = lngSys

    gstrUnitName = GetUnitName
    gblnInsure = True
    Call gclsInsure.InitOracle(gcnOracle)
    '-------------------------------------------------
        
    frmMainQuery.Show
    
End Sub

Public Sub InitData()

End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '����:�ر������Ӵ���
    
    Dim FrmThis As Form
    
    On Error Resume Next

    CloseChildWindows = True
    
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption And FrmThis.Caption <> FrmSon.Caption Then Unload FrmThis
    Next
    
    '�رչ��������Ĵ���
    If CloseChildWindows Then CloseChildWindows = CloseWindows

End Function

Public Sub RunMudal(ByVal lngNO As Long)
    Select Case lngNO
    Case 1
        frmDefTable.Show , gfrmMain
    Case 2
        frmPicture.Show , gfrmMain
    Case 3
        frmDoctor.Show , gfrmMain
    Case 4
        frmAdvice.Show , gfrmMain
    Case 5
        frmDefQuery.Show , gfrmMain
    Case 6
        frmDefTree.Show , gfrmMain
    Case 7
'        If gblnInsure Then
'            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
'        End If
        
        Call gclsInsure.InitOracle(gcnOracle)
        
        frmMainQuery.Show , gfrmMain
    Case 8
        Call InitLocPar
        Call InitSysPar
        
        On Error Resume Next
        
        frmselectinfo.Show , gfrmMain
    Case 9
        frmLisPrinterSetup.Show , gfrmMain
    End Select
End Sub


