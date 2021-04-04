Attribute VB_Name = "mdlRunAfter"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2018/12/25
'ģ��           mdlRunAfter
'˵��           �ӳٽű�ִ����
'==================================================================================================
Private Const mstrCurModule     As String = "mdlRunAfter"           '��ǰģ������
'˵����
'1��һ�����ݿ�����ж�����߿⣬��LIS�ͱ�׼����һ��ʵ���϶�����װ��
'2����ʷ����Ӻ������ű�û����Ҫ�����߿�ִ�еĽű�����ʷ���������ǵ����Ŀ⡣
'3��ͳ����Ϣ�ռ�ֻ��Ҫ��ǰʵ����DBA��
'4���Ӻ�ִ�����͵������ű���ʱ�������������ӿ���ṹ��
'��������ԭ��ֻ����DBA�û�����ʷ���û���֤��
Private Enum ScriptStruct
    ESS_�汾�� = 0
    ESS_���� = 1
    ESS_������ = 2
    ESS_������ = 3
    ESS_���������� = 4
    ESS_���������� = 5
    ESS_DBA�û��� = 6
    ESS_DBA���� = 7
    ESS_��ʷ�� = 8
    ESS_��ʷ��ű� = 9
    ESS_ͳ����Ϣ = 10
    ESS_�ӳ�ִ�нű� = 11
End Enum
Private mstrServer              As String                           '��ǰ����������
Private mrsHistoryDeferred      As ADODB.Recordset                  '��ǰ������������ʷ����ӳ�ִ�нű�����ʱ��֧�֣�����������
Private mrsAppDeferred          As ADODB.Recordset                  '��ǰ������������Ӧ��ϵͳ�ӳ�ִ�нű�����ʱ��֧�֣�����������
Private mrsToolDeferred         As ADODB.Recordset                  '��ǰ������������ZLTOOLS�ӳ�ִ�нű�����ʱ��֧�֣�����������
'"ID", adInteger, Empty, Empty, "ϵͳ", adInteger, Empty, Empty, "BAKDBName", adVarChar, 100, Empty, _
"BAKUser", adVarChar, 100, Empty, "������", adVarChar, 500, Empty, "DBLINK", adVarChar, 200, Empty, _
"SQL", adVarChar, 500, Empty, "ExecOrder", adInteger, Empty, Empty, "FixType", adInteger, Empty, Empty, _
"ExecDB", adInteger, Empty, Empty, "ExecLater", adInteger, Empty, Empty,"DB_ID", adInteger, Empty, Empty,ScriptNO", adInteger, Empty, Empty, "DDLParallel", adInteger, Empty, Empty))
Private mrsHisScript            As ADODB.Recordset                  '��ǰ��������������ʷ�������ű�����ʷ�⵱ǰû����Ҫ�����߿�ִ�е�SQL
'"ID", adInteger, Empty, Empty, "Owner", adInteger, Empty, Empty, "TableName", adVarChar, 100, Empty, _
"SQL", adVarChar, 500, Empty,"ScriptNO", adInteger, Empty, Empty))
Private mrsStatistics           As ADODB.Recordset                  '��ǰ������������ͳ����Ϣ�ռ��ű�
'"ϵͳ���", adInteger, Empty, Empty, "ϵͳ����", adVarChar, 50, Empty, "ϵͳ�汾", adVarChar, 20, Empty, "�����ļ�", adVarChar, 2000, Empty, _
"���", adInteger, Empty, Empty, "����", adVarChar, 30, Empty, "������", adVarChar, 50, Empty, _
"��ǰ", adInteger, Empty, Empty, "DB����", adVarChar, 200, Empty, "����", adVarChar, 100, Empty, _
"������", adVarChar, 500, Empty, "����", adInteger, Empty, Empty, "��ǰ�汾", adVarChar, 20, Empty, _
"Ŀ��汾", adVarChar, 20, Empty, "��ֹ��Ϣ", adVarChar, 2000, Empty, "������", adInteger, 1, 0, "�����", adVarChar, 2000, Empty, _
"��ǰĿ��汾", adVarChar, 20, Empty, "��ǰ��ֹ��Ϣ", adVarChar, 2000, Empty, "����ǰ����", adInteger, 1, 0, "��ǰ�����", adVarChar, 2000, Empty, _
"��֤", adInteger, Empty, Empty))
Private mrsHistory              As ADODB.Recordset                  '��ʷ����Ϣ
Private mlngCurFileLen          As Long                             '��ǰ׼��������ļ�����
Private mlngCurMaxScriptNo      As Long                             '��ǰ׼��������ļ������ű���

Private mstrDBAUser             As String                           'DBA�û���
Private mstrDBAPWD              As String                           'DBA�û�����
Private mblnDBAOK               As Boolean                          'DBA�û��Ƿ����ӳɹ�
Private mcnDBA                  As ADODB.Connection

Private mblnExecAgain           As Boolean                          '��ǰ�����������������ű�������ڴ�ִ�С���ʱ���ٽ����û���֤
Private mcllHistory             As New Collection                   '�Ѿ���֤����ʷ��
Private mlngHisID               As Long                             '��ʷ��ID��ǣ�����ѡ��
'���÷�����
Public Property Get Server() As String
    Server = mstrServer
End Property

Public Property Let Server(strServer As String)
    If mstrServer <> strServer Then
        Set mrsHistory = Nothing
        Set mrsHisScript = Nothing
        Set mrsStatistics = Nothing
        Set mcllHistory = Nothing
        mlngCurFileLen = 0
        mlngCurMaxScriptNo = -1
        mblnDBAOK = False
        mstrDBAUser = ""
        mstrDBAPWD = ""
        mblnExecAgain = False
        Set mcnDBA = Nothing
    Else
        mblnExecAgain = True
        Set mrsHisScript = Nothing
        Set mrsStatistics = Nothing
        Set mrsHistory = Nothing
    End If
    mstrServer = strServer
End Property
'����DAB
Public Property Get DBAUser() As String
    DBAUser = mstrDBAUser
End Property

Public Property Let DBAUser(strDBAUser As String)
    mstrDBAUser = strDBAUser
End Property

Public Property Get DBAPWD() As String
    DBAPWD = mstrDBAPWD
End Property

Public Property Let DBAPWD(strDBAPWD As String)
    mstrDBAPWD = strDBAPWD
End Property

Public Property Get IsDBAOK() As Boolean
    IsDBAOK = mblnDBAOK
End Property

Public Property Let IsDBAOK(blnDBAOK As Boolean)
    mblnDBAOK = blnDBAOK
End Property
'--------------------------------------------------------------------------------------------------
'�ӿ�               RunUpgradeAfter
'����               ִ�������Ƿ����
'����ֵ
'����б�:
'������         ����                        ˵��
'-------------------------------------------------------------------------------------------------
Public Function RunUpgradeAfter() As Boolean
    Dim lngScriptNo         As Long, lngOjbectNo    As Long, lngSQLID     As Long, intIniFileNo   As Integer, blnOk As Boolean
    Dim arrTmp              As Variant
    Dim conTmp              As ADODB.Connection
    Dim strLastCon          As String
    Dim cllHisCon           As New Collection
    Dim comTmp              As New ADODB.Command
    Dim i                   As Long, intLastDDLParallel As Integer
    Dim lngTotal            As Long, lngCurCount        As Long
    
    
    On Error GoTo errH
    '˵���н�������ִ�У����˳���������֤��Ԥ����������ֹ�������̴���
    If Not SaveOrReadExecuteRunAfterInfo(intIniFileNo, lngScriptNo, lngOjbectNo, lngSQLID) Then
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash("���ڶ�ȡ�Ӻ�ִ�нű���", , , Server)
    '��ȡ�ű�
    If Not ReadRunAfter(lngScriptNo, lngOjbectNo, lngSQLID) Then
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, lngScriptNo, lngOjbectNo, lngSQLID)
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash
    If Not RunAterIdentifyUsers Then
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, lngScriptNo, lngOjbectNo, lngSQLID)
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash
    If Not mrsHisScript Is Nothing Then
        mrsHisScript.Filter = ""
        lngTotal = mrsHisScript.RecordCount
    End If
    If Not mrsStatistics Is Nothing And IsDBAOK Then
        mrsStatistics.Filter = ""
        lngTotal = lngTotal + mrsStatistics.RecordCount
    End If
    If lngTotal = 0 Then lngTotal = 1
    
    For i = lngScriptNo To mlngCurMaxScriptNo
        If Not mrsHisScript Is Nothing Then
            mrsHisScript.Filter = "ScriptNO=" & i
            mrsHisScript.Sort = "ID"
            Do While Not mrsHisScript.EOF
                If strLastCon <> "K_" & mrsHisScript!������ & "|" & mrsHisScript!BAKUser Then
                    '�رղ���
                    If strLastCon <> "" Then
                        If Not conTmp Is Nothing Then
                            Call SetSessionParallel(conTmp, False, intLastDDLParallel)
                        End If
                    End If
                    strLastCon = "K_" & mrsHisScript!������ & "|" & mrsHisScript!BAKUser
                    intLastDDLParallel = Val(mrsHisScript!DDLParallel)
                    If Not InCollection(cllHisCon, strLastCon) Then
                        mrsHistory.Filter = "ϵͳ���=" & mrsHisScript!ϵͳ & " And ����='" & mrsHisScript!BAKDBName & "'"
                        Set conTmp = gobjRegister.GetConnection(mrsHistory!������, mrsHistory!������, mrsHistory!����, False, MSODBC, "", False)
                        If conTmp.State = adStateClosed Then
                            Set conTmp = Nothing
                        End If
                        cllHisCon.Add conTmp, strLastCon
                        '��������
                        If Not conTmp Is Nothing Then
                            Call SetSessionParallel(conTmp, True, intLastDDLParallel)
                        End If
                    Else
                        Set conTmp = cllHisCon(strLastCon)
                    End If
                End If
                lngCurCount = lngCurCount + 1
                Call ShowFlash("��  �ȣ�" & lngCurCount & "/" & lngTotal & "  ��ʷ������Լ������", lngCurCount / lngTotal, mrsHisScript!SQL, Server)
                If Not conTmp Is Nothing Then
                    On Error Resume Next
                    If mrsHisScript!ExecDB = 1 Then
                        '��ǰ�����߿�ִ�еĽű�δ�����Ӻ�ִ��
'                        Set comTmp.ActiveConnection = mcnOracle
                    Else
                        Set comTmp.ActiveConnection = conTmp
                    End If
                    comTmp.CommandText = mrsHisScript!SQL
                    DoEvents
                    comTmp.Execute
                    If Err.Number <> 0 Then
                        Debug.Print Err.Description & "-" & mrsHisScript!SQL
                        Err.Clear
                    End If
                    On Error GoTo errH
                End If
                Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 2, mrsHisScript!Id)
                mrsHisScript.MoveNext
            Loop
        End If
        '�رղ��С����������Ϊ�������нű�����һ�����ݿ⣬����ִ�д�����
        If strLastCon <> "" Then
            If Not conTmp Is Nothing Then
                Call SetSessionParallel(conTmp, False, intLastDDLParallel)
            End If
        End If
        '��ǵ�ǰ�ű����е���ʷ���Լ��������
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 3, 0)
        If IsDBAOK Then
            If mcnDBA Is Nothing Then
                Set mcnDBA = gobjRegister.GetConnection(mstrServer, DBAUser, DBAPWD, False, MSODBC, "", False)
            End If
            If Not mrsStatistics Is Nothing And mcnDBA.State = adStateOpen Then
                mrsStatistics.Filter = "ScriptNO=" & i
                mrsStatistics.Sort = "ID" '���DB_ID�ֶα�֤�����Ψһ��
                Do While Not mrsStatistics.EOF
                    '���ð�ʱָ������������ODBC���ӷ�ʽ֧��
                    '��connection����excute������Options����ֵΪ�⼸�������ԣ�adCmdUnknown 'adCmdStoredProc 'adExecuteNoRecords
                    '��Command���󣬱���ָ��CommandType = adCmdStoredProc
                    On Error Resume Next
                    lngCurCount = lngCurCount + 1
                    Call ShowFlash("��  �ȣ�" & lngCurCount & "/" & lngTotal & "  ͳ����Ϣ�ռ�", lngCurCount / lngTotal, mrsStatistics!SQL, Server)
                    DoEvents
                    mcnDBA.Execute mrsStatistics!SQL & "", , adCmdStoredProc
                    If Err.Number <> 0 Then
                        Debug.Print Err.Description & "-" & mrsStatistics!SQL
                        Err.Clear
                    End If
                    On Error GoTo errH
                    Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 3, mrsStatistics!Id)
                    mrsStatistics.MoveNext
                Loop
            End If
        End If
        '��ǣ���ǰ�ű����е�ͳ����Ϣ�Ѿ��ռ����
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i + 1, 0, 0)
    Next

    If mlngCurFileLen = FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL") Then '��ֹ�ڽű�ִ���нű������仯
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, mlngCurMaxScriptNo + 1, 0, 0)
        Kill IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini"
        Name IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL" As IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & "_" & Format(Now, "YYYYMMDDHHmmss") & ".SQL"
        RunUpgradeAfter = True
    Else
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, mlngCurMaxScriptNo + 1, 0, 0)
    End If
    Call ShowFlash
    Exit Function
errH:
    Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, , , , True)
    Call ShowFlash
    RunUpgradeAfter = True
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function


'--------------------------------------------------------------------------------------------------
'�ӿ�           ReadRunAfter
'����           -��ȡRunAfter.SQL�Ľű�����
'����ֵ         Boolean                    �Ƿ��ȡ�ɹ�
'����б�:
'������         ����                        ˵��
'lngScriptNo    Long                        ִ�е��Ľű���λ��
'lngOjbectNo    Long                        ִ�е��Ľű����ж�������
'lngSQLID       Long                        �Ѿ�ִ�е�SQLID
'˵����
'��ʷ��������SQL�����һ�θ���ʷ�������Ϊ׼��
'ͳ����Ϣ�ռ������һ��Ϊ��׼���𽥵���������һ�β����ڵı����
'-------------------------------------------------------------------------------------------------
Public Function ReadRunAfter(ByVal lngScriptNo As Long, ByVal lngOjbectNo As Long, ByVal lngSQLID As Long) As Boolean
    Dim objTxt          As TextStream, strLine              As String
    Dim lngCurScriptNo  As Long, arrScript()                As Variant, arrLine             As Variant, i               As Long
    Dim cllStatictics    As New Collection
    Dim rsTmpHis        As ADODB.Recordset, rsTmpHisAfter   As ADODB.Recordset, rsTmpSta As ADODB.Recordset
    Dim conTmp          As ADODB.Connection
    Dim strFileter      As String

    On Error GoTo errH
    If gobjFSO.FileExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL") Then
        gobjFSO.CopyFile IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL", IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL", True '����һ����ʱ�ļ���ִ�нű�����ֹ�ļ���ִ�й�����д��
        '--[SERVER]:Oracle
        '--[SCRIPT]:SerializeMulti(�汾��ʶ,���в���,������,Ӧ��ϵͳ������, Sm4EncryptEcb(Ӧ��ϵͳ����������), Sm4EncryptEcb(����������), DBA�û���, Sm4EncryptEcb(DBA����), Sm4EncryptEcb(gclsBase.Serialize(��ʷ����Ϣ��¼��), G_APP_KEY), ��ʷ��ű���¼��, ͳ����Ϣ�ռ���¼��, �ӳ�ִ�нű���¼��)
        '--[�ű�����]:
        Set objTxt = gobjFSO.OpenTextFile(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL", ForReading)
        Do While Not objTxt.AtEndOfStream
            strLine = objTxt.ReadLine
            If strLine Like "--[[]SCRIPT[]]:*" Then
                arrLine = UnSerializeMulti(Mid(strLine, Len("--[SCRIPT]:*")))
                Set arrLine(ESS_��ʷ��) = UnSerialize(Sm4DecryptEcb(arrLine(ESS_��ʷ��), G_APP_KEY))
                arrLine(ESS_����������) = Sm4DecryptEcb(arrLine(ESS_����������), G_APP_KEY)
                arrLine(ESS_DBA����) = Sm4DecryptEcb(arrLine(ESS_DBA����), G_APP_KEY)
                arrLine(ESS_����������) = Sm4DecryptEcb(arrLine(ESS_����������), G_APP_KEY)
                ReDim Preserve arrScript(lngCurScriptNo)
                arrScript(lngCurScriptNo) = arrLine
                lngCurScriptNo = lngCurScriptNo + 1
            End If
        Loop

        '��ʷ���û���֤
        For i = UBound(arrScript) To LBound(arrScript) Step -1
            If Not arrScript(i)(ESS_��ʷ��) Is Nothing Then
                Set rsTmpHis = arrScript(i)(ESS_��ʷ��)
                Set rsTmpHisAfter = arrScript(i)(ESS_��ʷ��ű�)
                '���ڽű��ҽű�δ��ȫִ�л�δִ��
                If Not rsTmpHisAfter Is Nothing And (i > lngScriptNo Or i = lngScriptNo And lngOjbectNo < 3) Then
                    If i = lngScriptNo Then
                        strFileter = "ID>" & lngSQLID
                    Else
                        strFileter = ""
                    End If
                    rsTmpHisAfter.Filter = strFileter
                    '����δִ�еĽű�
                    If rsTmpHisAfter.RecordCount <> 0 Then
                        Do While Not rsTmpHis.EOF
                            '���������ʷ�⣬��һ�������ľͼ��룬֮��Ĳ��ټ���
                            If Not InCollection(mcllHistory, "K_" & rsTmpHis!ϵͳ��� & "_" & rsTmpHis!����) Then
                                rsTmpHisAfter.Filter = strFileter & IIf(strFileter <> "", " And ", "") & "ϵͳ=" & rsTmpHis!ϵͳ��� & " And BAKDBName='" & rsTmpHis!���� & "'"
                                '��ǰ��ʷ�����δִ�еĽű��������ýű��Լ�����ʷ��
                                If rsTmpHisAfter.RecordCount <> 0 Then
                                    If mrsHistory Is Nothing Then '��ʼ����¼��
                                        Set mrsHistory = CopyNewRec(rsTmpHis, True)
                                        Set mrsHisScript = CopyNewRec(rsTmpHisAfter, True, , Array("ScriptNO", adInteger, Empty, Empty, "DDLParallel", adInteger, Empty, Empty))
                                    End If
                                    Call RecDataAppend(mrsHistory, rsTmpHis, 1, , , True)
                                    Call RecDataAppend(mrsHisScript, rsTmpHisAfter, , "-ScriptNO,DDLParallel", , , Array("ScriptNO", i, "DDLParallel", Val(arrScript(i)(ESS_����))))
                                    '������֤
                                    Set conTmp = gobjRegister.GetConnection(rsTmpHis!������, rsTmpHis!������, rsTmpHis!����, False, MSODBC, "", False)
                                    If conTmp.State = adStateClosed Then
                                        mrsHistory.Update Array("ID", "��ǰ�汾", "Ŀ��汾", "��֤"), Array(mrsHistory.RecordCount, Null, Null, 0)
                                        mcllHistory.Add 0, "K_" & rsTmpHis!ϵͳ��� & "_" & rsTmpHis!����
                                    Else
                                        mrsHistory.Update Array("ID", "Ŀ��汾", "��֤"), Array(mrsHistory.RecordCount, Null, 1)
                                        mcllHistory.Add 1, "K_" & rsTmpHis!ϵͳ��� & "_" & rsTmpHis!����
                                    End If
                                    If conTmp.State = adStateOpen Then
                                        conTmp.Close
                                    End If
                                    Set conTmp = Nothing
                                End If
                            End If
                            rsTmpHis.MoveNext
                        Loop
                    End If
                End If
            End If
            If Not arrScript(i)(ESS_ͳ����Ϣ) Is Nothing Then
                If Not IsDBAOK Then
                    If DBAUser <> arrScript(i)(ESS_DBA�û���) Or DBAPWD <> arrScript(i)(ESS_DBA����) Then
                        DBAUser = arrScript(i)(ESS_DBA�û���)
                        DBAPWD = arrScript(i)(ESS_DBA����)
                        Set conTmp = gobjRegister.GetConnection(Server, DBAUser, DBAPWD, False, MSODBC, "", False)
                        If conTmp.State = adStateOpen Then
                            IsDBAOK = True
                            conTmp.Close
                        End If
                        Set conTmp = Nothing
                    End If
                End If
                Set rsTmpSta = arrScript(i)(ESS_ͳ����Ϣ)
                If mrsStatistics Is Nothing Then
                    Set mrsStatistics = CopyNewRec(rsTmpSta, True, , Array("ScriptNO", adInteger, Empty, Empty))
                End If
                '���ڽű��ҽű�δ��ȫִ�л�δִ��
                If i >= lngScriptNo Then
                    If i = lngScriptNo And lngOjbectNo = 3 Then
                        rsTmpSta.Filter = "ID>" & lngSQLID
                    Else
                        rsTmpSta.Filter = ""
                    End If
                    '����δִ�еĽű�
                    If rsTmpSta.RecordCount <> 0 Then
                        Do While Not rsTmpSta.EOF
                            If Not InCollection(cllStatictics, "K_" & rsTmpSta!Owner & "." & rsTmpSta!TableName) Then
                                '����ͳ����Ϣ���ռ�����δ��ӣ�����ӣ���ӵ�ǰ�в����α�ع�
                                Call RecDataAppend(mrsStatistics, rsTmpSta, 1, "-ScriptNO", , True, Array("ScriptNO", i))
                                cllStatictics.Add 1, "K_" & rsTmpSta!Owner & "." & rsTmpSta!TableName
                            End If
                            rsTmpSta.MoveNext
                        Loop
                    End If
                End If
            End If
        Next
        mlngCurFileLen = FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL")
        mlngCurMaxScriptNo = UBound(arrScript)
        ReadRunAfter = True
        objTxt.Close
        Set objTxt = Nothing
        Kill IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL"
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Private Sub SetSessionParallel(ByRef cnInput As ADODB.Connection, Optional ByVal blnEnabled As Boolean, Optional ByVal intDDLParallel As Integer)
'���û����DDL
    Dim strSQL As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    If intDDLParallel <= 1 Then Exit Sub
    If blnEnabled Then
        strSQL = "Alter Session FORCE PARALLEL DDL PARALLEL " & intDDLParallel
        cnInput.Execute strSQL
    Else
        strSQL = "ALTER Session DISABLE PARALLEL DDL "
        cnInput.Execute strSQL
        strSQL = "Select 'alter index ' || Index_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Indexes" & vbNewLine & _
                    "Where Degree Not In ('0', '1') and index_type='NORMAL' And temporary='N'" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 'alter table ' || Table_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Tables" & vbNewLine & _
                    "Where Degree != ('         1')"
        Set rsTmp = gobjRegister.OpenSQLRecord(cnInput, strSQL, App.Title)
        On Error Resume Next
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                cnInput.Execute rsTmp!SQL, , adCmdText
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Sub

'--------------------------------------------------------------------------------------------------
'�ӿ�           RunAterIdentifyUsers
'����           -��֤RunAter�ű�ִ�е���ʷ���û���Ϣ��ͬʱ��֤DBA�û�������ͳ����Ϣ��
'����ֵ         ADODB.Recordset            ���ص���ʷ����֤��Ϣ
'����б�:
'������         ����                        ˵��
'blnDo          Boolean                     �Ƿ�ִ���ӳ�ִ�нű�
'-------------------------------------------------------------------------------------------------
Private Function RunAterIdentifyUsers() As Boolean
    Dim rsTmp       As ADODB.Recordset
    '�ÿ�ĵڶ�����֤
    If mblnExecAgain Then
        RunAterIdentifyUsers = True
        Exit Function
    End If
    On Error GoTo errH
    If Not mrsHistory Is Nothing Then
        mrsHistory.Filter = "��֤=0"
        If mrsHistory.RecordCount <> 0 Then
            '��δ��֤ͨ���ĵ���������������ֹ���������������
            Set rsTmp = CopyNewRec(mrsHistory)
            Call RecDelete(mrsHistory, "��֤=0")
            Call RecUpdate(rsTmp, "", "����", "")
        End If
    End If
    If Not mrsStatistics Is Nothing And Not IsDBAOK Then
        Call frmUsers.ShowMe(rsTmp, True)
    ElseIf Not rsTmp Is Nothing Then
        Call frmUsers.ShowMe(rsTmp)
    Else
        RunAterIdentifyUsers = True
        Exit Function
    End If
    If Not rsTmp Is Nothing Then
        rsTmp.Filter = ""
        Call RecDataAppend(mrsHistory, rsTmp)
    End If
    RunAterIdentifyUsers = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

''���ø÷�����ע��ر�
Public Function SaveOrReadExecuteRunAfterInfo(intFileNo As Integer, Optional lngScriptNo As Long, Optional lngOjbectNo As Long, Optional lngSQLID As Long, Optional ByVal blnForceClaose As Boolean) As Boolean
'���ܣ�����RunAfter�Ľű�ִ�����
    '[�ű�λ��]�ű����,�������,SQL���
    '�ֽ�����10  4    1   4    1  4
    On Error GoTo errH
    If intFileNo = 0 Then
        intFileNo = FreeFile()
        Open IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini" For Binary Access Read Write Lock Read Write As intFileNo
        If FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini") < 24 Then
            Put #intFileNo, 1, StrConv("[�ű�λ��]", vbFromUnicode)
            Put #intFileNo, 11, lngScriptNo
            Put #intFileNo, 15, CByte(44)
            Put #intFileNo, 16, lngOjbectNo
            Put #intFileNo, 20, CByte(44)
            Put #intFileNo, 21, lngSQLID
        Else
            Get #intFileNo, 11, lngScriptNo
            Get #intFileNo, 16, lngOjbectNo
            Get #intFileNo, 21, lngSQLID
        End If
    ElseIf intFileNo > 0 Then
        Put #intFileNo, 11, lngScriptNo
        Put #intFileNo, 16, lngOjbectNo
        Put #intFileNo, 21, lngSQLID
    Else
        If Not blnForceClaose Then
            Put #Abs(intFileNo), 11, lngScriptNo
            Put #Abs(intFileNo), 16, lngOjbectNo
            Put #Abs(intFileNo), 21, lngSQLID
        End If
        Close #Abs(intFileNo)
    End If
    SaveOrReadExecuteRunAfterInfo = True
    Exit Function
errH:
    '������д����ֹ�����ͬʱִ�С�
    Err.Clear
End Function
'




