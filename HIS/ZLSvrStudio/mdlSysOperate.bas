Attribute VB_Name = "mdlSysOperate"
Option Explicit
'�ļ�����,������˳�����ļ�ִ��˳����ͬ
'X.X.X����Ϊ4λ�汾��X.X.X.X,��ʱΪ����SP�ű���
Public Enum FileType
    'FT_Before �ű���FT_DBA�ű�ִ��ִ��˳����Ի���
    FT_DBA = 0 '��ҪDBA�û�ִ�еĽű�(System�û�):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_Before = 1 '��ǰִ�нű���ZLUPgradeX.X.X_Before.sql.sql(�����ߣ�,ZL*_X.X.X_History_Before.sql (Ӧ��ϵͳ��ʷ��)ZL*_X.X.X_Before.sql(Ӧ��ϵͳ���߿�) *����ϵͳ��\100
    FT_Standard = 2 '��ͨ�����ű���ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(����).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(����).sql,ZL*_X.X.X_History.sql
    FT_Optional = 3 '��ѡִ�нű�:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql��ZL*_X.X.X__HISTORY_Optional.sql
    FT_Deferred = 4 '�ӳ�ִ�нű�:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'�ļ�����ϵͳ
Public Enum SysType
    ST_Tools = 0 '�����߽ű�,�����ļ����ͣ�FT_Before,FT_DBA,FT_Standard,FT_Optional
    ST_App = 1 'Ӧ��ϵͳ���߿�,�����ļ����ͣ�FT_Before,FT_DBA,FT_Standard,FT_Optional��FT_Deferred
    ST_History = 2 'Ӧ��ϵͳ��ʷ�⣬�����ļ����ͣ�FT_Before,FT_Standard,FT_Deferred��FT_Optional
End Enum
'�汾����
Public Enum VersionType
    VT_Normal = 0 '�����汾
    VT_Supple = 1 '���䷢���汾����һ����汾������ǰһ���汾�·�����SP���ǲ���汾
End Enum

Public Enum UserCheckType
    UCT_ZLTOOLS = 0 '�������û���֤
    UCT_DBAUser = 1 'DBA�û���֤
    '��ǰ��������Ϊ1�����ڵ���Ϊ2����Ҫ���������⼸�����Ͷ���ͨ��ֱ�ӵ��ô�����ʹ�õ�
    UCT_CurZLBAK = 2 '��ǰ��ʷ����֤
    UCT_NormalUser = 3 '��ͨ�û���֤
    UCT_SysOwner = 4 '����Ա��¼��֤
    UCT_RACInsUser = 5 'RACʵ���û���֤
    UCT_AuditLog = 6   '��¼��Ҫ��־
End Enum

Public gcllMustObj As Collection '��Ҫ������
Public gobjLog As TextStream
Private mstrStSysOwner As String '��׼��������
Public Function CheckAndAdjustMustTable(ByVal strTable As String, Optional ByVal strColumn As String, Optional ByVal blnMsg As Boolean, Optional ByVal strOwner As String = "ZLTOOLS", Optional ByVal blnCache As Boolean = True) As Boolean
'���ܣ���鲢������Ҫ�����ݽṹ
'������strTable=����
'         strColumn=����
'         blnMsg=��鲢�޸�ʧ���Ƿ���ʾ
'         strOwner=����������
'         blnCache=�ж��Ƿ񻺴����������һЩ���������Ҫ���棬������ͨ���󲻻��棬��ֹ���򻺴����ݽ϶�
'���أ���鲢�޸��Ƿ�ɹ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnHaveTable As Boolean, blnHaveColumn As Boolean
    Dim strFMT As String
    Dim objGetTmp As New clsObjectInfo, objParent As clsObjectInfo, objCurent As clsObjectInfo
    Dim blnHaveData As Boolean
    
    strTable = UCase(strTable): strColumn = UCase(strColumn): strOwner = UCase(strOwner)
    '�������ö������޸�����
    If gcllMustObj Is Nothing Then
        Set gcllMustObj = New Collection
        'ZLUpgrade�����Լ���ǰ�м��
        Set objParent = objGetTmp.GetObject("ZLUPGRADE", OT_Table, _
                                        "CREATE TABLE ZLTOOLS.zlUpgrade(ϵͳ NUMBER(5),ԭʼ�汾 VARCHAR2(10),Ŀ��汾 VARCHAR2(10),��Ǩʱ�� DATE,��Ǩ��� NUMBER(1)" & _
                                        ",����汾 VARCHAR2(10),��ֹ��� VARCHAR2(200),��ǰִ�� number(1))PCTFREE 5|" & _
                                        "ALTER TABLE ZLTOOLS.zlUpgrade ADD CONSTRAINT  zlUpgrade_UQ_��Ǩʱ�� Unique (ϵͳ,��Ǩʱ��)   USING INDEX PCTFREE 5|" & _
                                        "ALTER TABLE ZLTOOLS.zlUpgrade ADD CONSTRAINT  zlUpgrade_FK_ϵͳ FOREIGN KEY (ϵͳ) REFERENCES zlSystems(���) ON DELETE CASCADE")
        Set objCurent = objGetTmp.GetObject("��ǰִ��", OT_Column, "alter Table ZLTOOLS.ZLUPGRADE add ��ǰִ�� number(1)", , objParent)
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLUPGRADE|��ǰִ��")
        'ZLBAKTABLES����
        Set objCurent = objGetTmp.GetObject("ZLBAKTABLES", OT_Table, _
                                        "Create Table ZLTOOLS.zlBakTables(ϵͳ Number(5),���� Varchar2(30),��� Number(2),��� Number(3),ֱ��ת�� Number(1),ͣ�ô����� number(1))|" & _
                                        "Alter Table ZLTOOLS.zlBakTables    Add Constraint zlBakTables_PK Primary Key (ϵͳ,����) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakTables Add Constraint zlBakTables_FK_ϵͳ Foreign Key (ϵͳ) References zlSystems(���) On Delete Cascade")
        
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLBAKTABLES|")
        'ZLBAKSPACES����
        Set objCurent = objGetTmp.GetObject("ZLBAKSPACES", OT_Table, _
                                        "Create Table ZLTOOLS.zlBakSpaces(ϵͳ Number(5),��� Number(18),���� Varchar2(30),������ Varchar2(30),DB���� Varchar2(128),��ǰ Number(1),ֻ�� Number(1))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces Add Constraint zlBakSpaces_PK Primary Key (ϵͳ,���) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces    Add Constraint zlBakSpaces_UQ_���� Unique (ϵͳ,����) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces Add Constraint zlBakSpaces_FK_ϵͳ Foreign Key (ϵͳ) References zlSystems(���) On Delete Cascade")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLBAKSPACES|")
        'zlUpgradeConfig����
        Set objCurent = objGetTmp.GetObject("zlUpgradeConfig", OT_Table, _
                                        "Create Table ZLTOOLS.zlUpgradeConfig(��Ŀ varchar2(50),���� varchar2(4000))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlUpgradeConfig Add Constraint zlUpgradeConfig_PK Primary Key (��Ŀ) USING INDEX PCTFREE 5|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('�ͻ���״̬',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('�û�״̬',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('��̨��ҵ״̬',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('������״̬',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('���õ�ϵͳ����',Null)")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|zlUpgradeConfig|")
        'ZLTriggers����
        Set objCurent = objGetTmp.GetObject("ZLTriggers", OT_Table, _
                                        "Create Table ZLTOOLS.ZLTriggers(���� varChar2(100),������ varChar2(100))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.ZLTriggers Add Constraint ZLTriggers_UQ_���� Unique (����,������) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.ZLTriggers Modify ����  constraint ZLTriggers_NN_����   not  null")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLTriggers|")
        'ZLClient�����Լ�ϵͳ���������м��
        Set objCurent = objGetTmp.GetObject("ϵͳ��������", OT_Column, "alter Table ZLTOOLS.ZLCLIENTS add ϵͳ�������� number(1)", , objGetTmp.GetObject("ZLCLIENTS", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLCLIENTS|ϵͳ��������")
        'ZLAutoJob�����Լ�ϵͳ����ͣ���м��
        Set objCurent = objGetTmp.GetObject("ϵͳ����ͣ��", OT_Column, "alter Table ZLTOOLS.ZLAutoJobs add ϵͳ����ͣ�� number(1)", , objGetTmp.GetObject("ZLAutoJobs", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLAutoJobs|ϵͳ����ͣ��")
        '�ϻ���Ա������Լ�ϵͳ����ͣ���м��
        Set objCurent = objGetTmp.GetObject("ϵͳ��������", OT_Column, "alter Table " & gstrUserName & ".�ϻ���Ա�� add ϵͳ�������� number(1)", gstrUserName, objGetTmp.GetObject("�ϻ���Ա��", OT_Table, , gstrUserName))
        gcllMustObj.Add objCurent, UCase(gstrUserName & "|�ϻ���Ա��|ϵͳ��������")
        'Zlsvrtools�����Լ������м��
        Set objCurent = objGetTmp.GetObject("����", OT_Column, "alter Table ZLTOOLS.Zlsvrtools add ���� number(3)", , objGetTmp.GetObject("Zlsvrtools", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|Zlsvrtools|����")
        'zlParameters�����Լ������м��
        Set objCurent = objGetTmp.GetObject("����", OT_Column, "alter table Zltools.zlParameters add ���� NUMBER(1)", , objGetTmp.GetObject("zlParameters", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|zlParameters|����")
    End If
    
    On Error Resume Next
    Set objCurent = gcllMustObj(strOwner & "|" & strTable & "|" & strColumn)
    If err.Number <> 0 Then
        err.Clear
        If strColumn = "" Then
            Set objCurent = objGetTmp.GetObject(strTable, OT_Table, , strOwner)
        Else
            Set objParent = gcllMustObj(strOwner & "|" & strTable & "|")
            If err.Number <> 0 Then
                err.Clear
                Set objParent = objGetTmp.GetObject(strTable, OT_Table, , strOwner)
            Else
                If blnCache Then gcllMustObj.Remove strOwner & "|" & strTable & "|" '�ϲ�������
            End If
            Set objCurent = objGetTmp.GetObject(strColumn, OT_Column, , strOwner, objParent)
        End If
        '���������
        If blnCache Then gcllMustObj.Add objCurent, UCase(strOwner & "|" & strTable & "|" & strColumn)
    End If
    If Not objCurent.ObjectCheck(blnMsg) Then
        Exit Function
    End If
    CheckAndAdjustMustTable = True
    On Error GoTo errh
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function GetConnection(ByVal strUserName As String, Optional ByVal blnValidate As Boolean = True) As ADODB.Connection
'���ܣ���ȡ����
'������strUserName=ZLTOOLS,�����ߣ�DBA-DBA�û�ȷ�ϣ������������û�ȷ��
'          blnValidate=True:������ܻ�ȡ��������Ҫ�������룬��������,False=������ܻ�ȡ������ֱ���˳�
    Dim uctType As UserCheckType
    Dim cnTmp As ADODB.Connection
    Dim blnNew As Boolean, blnViladate As Boolean
    
    Select Case UCase(strUserName)
        Case "ZLTOOLS"
            If gcnTools Is Nothing Then
                blnNew = True
            ElseIf gcnTools.State = adStateClosed Then
                blnNew = True
            End If
            If blnNew Then
                Set gcnTools = gobjRegister.GetConnection(gstrServer, "ZLTOOLS", IIf(gstrToolsPwd = "", "ZLTOOLS", gstrToolsPwd), False, MSODBC, "", False)
                If gcnTools.State = adStateOpen Then
                    Call SetSQLTrace(gstrServer, "ZLTOOLS", gcnTools)
                    gstrToolsPwd = IIf(gstrToolsPwd = "", "ZLTOOLS", gstrToolsPwd)
                    Set GetConnection = gcnTools: Exit Function
                ElseIf gstrToolsPwd = "" Then
                    Set gcnTools = gobjRegister.GetConnection(gstrServer, "ZLTOOLS", "ZLSOFT", False, MSODBC, "", False)
                    If gcnTools.State = adStateOpen Then
                        Call SetSQLTrace(gstrServer, "ZLTOOLS", gcnTools)
                        gstrToolsPwd = "ZLSOFT"
                        Set GetConnection = gcnTools: Exit Function
                    End If
                End If
            Else
                Set GetConnection = gcnTools: Exit Function
            End If
            uctType = UCT_ZLTOOLS
        Case "DBA", "SYSTEM", "SYS"
            If gcnSystem Is Nothing Then
                blnNew = True
            ElseIf gcnSystem.State = adStateClosed Then
                blnNew = True
            End If
            If gstrSysPwd <> "" And blnNew Then
                Set gcnSystem = gobjRegister.GetConnection(gstrServer, gstrSysUser, gstrSysPwd, False, MSODBC, "", False)
                If gcnSystem.State = adStateOpen Then
                    Call SetSQLTrace(gstrServer, gstrSysUser, gcnSystem)
                    Set GetConnection = gcnSystem: Exit Function
                End If
            ElseIf Not blnNew Then
                Set GetConnection = gcnSystem: Exit Function
            End If
            uctType = UCT_DBAUser
            If UCase(strUserName) = "DBA" Then strUserName = "SYSTEM"
        Case Else
            uctType = UCT_NormalUser
    End Select
    If blnValidate Then
        If Not frmUserCheckLogin.ShowLogin(uctType, cnTmp, strUserName) Then Exit Function
        
        Call SetSQLTrace(gstrServer, strUserName, cnTmp)
        Set GetConnection = cnTmp
        If uctType = UCT_ZLTOOLS Then
            Set gcnTools = cnTmp
        ElseIf uctType = UCT_DBAUser Then
            Set gcnSystem = cnTmp
        End If
    End If
End Function

Public Sub RecToLog(ByVal rsInput As ADODB.Recordset, Optional ByVal strSort As String, Optional ByVal strName As String)
'����¼��ת��Ϊ�ַ������������ټ�¼��־
    Dim i As Long
    Dim lngShort As Long
    Dim strLine As String
    
    If Not gblnTrace Then Exit Sub
    If rsInput Is Nothing Then
        WriteTraceLog "===============" & strName & "========================="
        WriteTraceLog "Nothing"
    End If
    rsInput.Filter = ""
    rsInput.Sort = strSort
    '������������־
    
   WriteTraceLog "===============" & strName & "========================="
    For i = 0 To rsInput.Fields.Count - 1
        strLine = strLine & RPAD(rsInput.Fields(i).name, 12)
    Next
    WriteTraceLog strLine
    Do While Not rsInput.EOF
        lngShort = 0: strLine = ""
        For i = 0 To rsInput.Fields.Count - 1
            If Len(rsInput.Fields(i).value & "") < 9 And lngShort <> 0 Then
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 9)
                lngShort = IIf(lngShort - 3 <= 0, 0, lngShort - 3)
            ElseIf Len(rsInput.Fields(i).value & "") > 12 Then
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 12)
                lngShort = lngShort + Len(rsInput.Fields(i).value & "") - 12
            Else
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 12)
            End If
        Next
        WriteTraceLog strLine
        rsInput.MoveNext
    Loop
End Sub

Public Function GetToolsVersion() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strVer As String, blnUpdate As Boolean
    
    On Error GoTo errh
    '��ȡ�����߰汾
    strSQL = "Select ���� From Zlreginfo Where ��Ŀ = '�汾��'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
    If Not rsTmp.EOF Then strVer = Trim(rsTmp!���� & "")
    '�������߰汾����Ч�汾�����Զ�����
    If strVer = "" Then
        blnUpdate = Not rsTmp.EOF '��Ҫ���°汾
        On Error Resume Next
        strSQL = "Select ����汾 From Zlupgrade Where ϵͳ Is Null And Nvl(��ǰִ��, 0) = 0 Order By ��Ǩʱ�� Desc"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
        If Not rsTmp.EOF Then strVer = Trim(rsTmp!����汾 & "")
        If err.Number <> 0 Then err.Clear
        On Error GoTo errh
        If strVer <> "" Then
            If blnUpdate Then
                gcnOracle.Execute "Update ZLreginfo set ����='" & strVer & "' where ��Ŀ='�汾��'"
            Else
                gcnOracle.Execute "Insert Into zlRegInfo(��Ŀ,�к�,����) Values('�汾��',1,'" & strVer & "')"
            End If
        End If
    End If
    GetToolsVersion = strVer
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'���ܣ���ָ��INI�����ļ������ݶ�ȡ����¼����
'���أ�Nothing�����"��Ŀ,����"�ļ�¼��,����ͬһ��Ŀ�����ж�������
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "��Ŀ", adVarChar, 100
    rsTmp.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = Null
                rsTmp.Update
            End If
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))

            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!��Ŀ = strItem
            rsTmp!���� = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!��Ŀ = strItem
        rsTmp!���� = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Public Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'���ܣ�����Ӧ�������ļ���ʽ�Ƿ���ȷ
'������rsINI=��������ļ����ݵļ�¼��������"��Ŀ,����"�ֶ�
'      strItem=�����ļ��б���Ҫ�������ݵ���Ŀ��,��"��Ŀ1|��Ŀ2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "��Ŀ='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If rsINI!���� & "" = "" Then Exit Function
        If arrItem(i) Like "*�汾��" Then
            If Not IsVerSion(rsINI!����) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function SplitLine(ByVal strSQL As String) As Variant
'���ܣ���SQL���л��в�֣�ͬʱ��¼���з�
    Dim arrLine As Variant, arrReturn() As Variant
    Dim i As Long, j As Long, lngStart As Long, lngEx As Long, lngCur As Long
    Dim strTmp As String
    arrReturn = Array()
    If strSQL = "" Then SplitLine = arrReturn: Exit Function
    arrLine = Split(Replace(Replace(strSQL, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim Preserve arrReturn(UBound(arrLine) * 2)
    lngStart = 1
    For i = LBound(arrLine) To UBound(arrLine)
        If i <> 0 Then
            strTmp = Mid(strSQL, lngStart, 2)
            If strTmp = vbCrLf Then
                arrReturn(i * 2 - 1) = vbCrLf
                lngStart = lngStart + 2
            Else
                arrReturn(i * 2 - 1) = Mid(strSQL, lngStart, 1)
                lngStart = lngStart + 1
            End If
        End If
        arrReturn(i * 2) = arrLine(i)
        lngStart = lngStart + Len(arrLine(i))
    Next
    SplitLine = arrReturn
End Function

Public Function TrimCommentLossless(ByVal strSQL As String) As String
'���ܣ�����ȥ��ע�ͣ���TrimComment�Ƚϣ����㷨��������ʵ���ݡ�
    Dim arrLine As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    Dim blnStr As Boolean, blnMultiCom As Boolean
    Dim lngPos1 As Long, lngPos2 As Long, lngPos3 As Long
    Dim blnAddLine As Boolean
    Dim strTmp As String, strFMT As String
    
    On Error GoTo errh
    'ȥ������ע�͡�
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
            blnAddLine = True
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                blnAddLine = j <> UBound(arrLine)
                If blnMultiCom Then '�Ѿ����ڶ���ע�ͷ�Χ�������Ȳ��ҽ�����
                    lngPos2 = InStr(strTmp, "*/")
                    If lngPos2 > 0 Then
                        strTmp = Mid(strTmp, lngPos2 + 2)
                        blnMultiCom = False
                    Else
                        strTmp = "": blnAddLine = False
                    End If
                End If
                If Not blnMultiCom Then '���/* -- */ ��/*   */--����
                    lngPos2 = InStr(strTmp, "/*")
                    lngPos1 = InStr(strTmp, "--")
                    'ȥ����Ч�Ķ���ע������'/* --*/ ,/* */ ����� --/* */
                    '1������--,����--�ڶ��п�ʼ��֮��
                    '2��������--�����ڶ��п�ʼ��
                    Do While Not blnMultiCom And (lngPos2 > 0 And lngPos2 < lngPos1 Or lngPos1 = 0 And lngPos2 > 0)
                        lngPos3 = InStr(lngPos2, strTmp, "*/")
                        If lngPos3 > 0 Then
                            strTmp = Left(strTmp, lngPos2 - 1) & Mid(strTmp, lngPos3 + 2)
                        Else
                            strTmp = Left(strTmp, lngPos2 - 1)
                            blnMultiCom = True
                        End If
                        lngPos2 = InStr(strTmp, "/*")
                        lngPos1 = InStr(strTmp, "--")
                    Loop
                End If
                'ע���еĿ��У���������
                If blnAddLine Then
                    strFMT = strFMT & strTmp & arrLine(j + 1)
                Else
                    strFMT = strFMT & strTmp
                End If
            Next
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then '�Ƕ���ע�ͣ�������ַ����߽�
            blnStr = Not blnStr '��ʼ�����ַ����߽�
        End If
    Next
    
    'ȥ������ע��
    arrTmp = Split(strFMT, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
'            blnMultiCom = False
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                If j = LBound(arrLine) And blnMultiCom Then
                    blnMultiCom = UBound(arrLine) = 0
                Else
                    blnAddLine = j <> UBound(arrLine)
                    lngPos1 = InStr(strTmp, "--")
                    If lngPos1 > 0 Then
                        strTmp = Left(strTmp, lngPos1 - 1)
                        blnMultiCom = UBound(arrLine) = j
                    End If
                    If blnAddLine Then
                        strFMT = strFMT & strTmp & arrLine(j + 1)
                    Else
                        strFMT = strFMT & strTmp
                    End If
                End If
            Next
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then
            blnStr = Not blnStr '��ʼ�����ַ����߽�
        End If
    Next
    TrimCommentLossless = strFMT
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetFMTSQLStr(ByVal strSQL As String, ByRef cllStrs As Collection) As String
'���ܣ���ȡSQL�е��ַ���������ռλ��ռλ�����ظ�ʽ����SQL
    Dim arrTmp As Variant
    Dim i As Long, j As Long, intIndex As Integer
    Dim strFMT As String, strTmp As String
    Dim blnStr As Boolean
    
    Set cllStrs = New Collection
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            strFMT = strFMT & arrTmp(i)
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            intIndex = intIndex + 1
            '����ַ���
            strFMT = strFMT & "[S" & intIndex & "]"
            cllStrs.Add strTmp, "S" & intIndex
        End If
        blnStr = Not blnStr '��ʼ�����ַ����߽�
    Next
    arrTmp = SplitLine(strFMT)
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(arrTmp(i))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    'ȥ���������еĿո�
    arrTmp = SplitLine(strFMT)
    strFMT = ""
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(TrimBesideOperator(arrTmp(i)))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    GetFMTSQLStr = UCase(strFMT)
End Function

Public Function TrimBesideOperator(ByVal strText As String) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long
    
    strText = Replace(Replace(strText, " :", ":"), ": ", ":")
    strText = Replace(Replace(strText, " =", "="), "= ", "=")
    strText = Replace(Replace(strText, " .", "."), ". ", ".")
    strText = Replace(Replace(strText, " )", ")"), ") ", ")")
    strText = Replace(Replace(strText, " (", "("), "( ", "(")
    strText = Replace(Replace(strText, " %", "("), "% ", "%")
    strText = Replace(Replace(strText, " \", "\"), "\ ", "\")
    TrimBesideOperator = strText
End Function

Public Function GetInfoInsideBracket(ByVal strInfo As String, Optional ByVal strLeftChar As String, Optional ByVal strRightChar As String) As String
'����������ȡ����
'����������������ݣ�ֻȡ�����
    Dim lngSart As Long, lngEnd As Long
    If strRightChar = "" Then strRightChar = ")"
    If strLeftChar = "" Then strLeftChar = "("
    lngEnd = InStrRev(strInfo, strRightChar) - Len(strRightChar) + 1 '��ͷ����β�����Բ���һ
    lngSart = InStr(strInfo, strLeftChar) + Len(strLeftChar)
    If lngEnd < lngSart Then
        GetInfoInsideBracket = ""
    Else
        GetInfoInsideBracket = Mid(strInfo, lngSart, lngEnd - lngSart)
    End If
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

Public Function TrimComment(ByVal strSQL As String) As String
'���ܣ�ȥ��д�ڵ���strSQL�������"--"ע��
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim blnStr As Boolean
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'���ܣ�ȡ";"��βǰ��ĵ�SQL���,����";"�ź���"--"ע�͡�
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long, K As Long
    
    '��ȥ��ע�Ͳ���
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            K = i: Exit For
        End If
    Next
    If K > 0 Then strSQL = Left(strSQL, K - 1)
    
    SplitSQL = strSQL
End Function

Public Function RemoveMark(ByVal strText As String) As String
'���ܣ�ȥ��һ�������е�ǰ��"--"ע�ͱ��
    Dim arrText As Variant, strTemp As String, i As Long
    
    arrText = Split(strText, vbCrLf)
    
    strText = ""
    For i = 0 To UBound(arrText)
        strTemp = arrText(i)
        If Left(strTemp, 2) = "--" And Replace(strTemp, "-", "") <> "" Then
            strText = strText & vbCrLf & Mid(strTemp, 3)
        End If
    Next
    RemoveMark = Mid(strText, 3)
End Function

Public Function GetLogSQL(objSQL As clsSQLInfo) As String
'���ܣ���ȡ��ҪSQL��䣬������д��־
    Dim strSQL As String
    
    If objSQL.Block Then
        If objSQL.BlockName <> "" Then
            strSQL = Trim(Split(objSQL.SQL, vbCrLf)(0))
            If InStr(strSQL, "(") > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            If InStr(1, strSQL, " as", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " is", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " is", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " Return", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " Return", vbTextCompare) - 1))
            End If
        Else '������
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 255)
        End If
    ElseIf UCase(LTrim(objSQL.SQL)) Like "CREATE * VIEW" Then
        '��ͼ���⴦��
        strSQL = Split(objSQL.SQL, vbCrLf)(0)
        If InStr(1, strSQL, " as", vbTextCompare) > 0 Then '��ͼֻ����as
            strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
        End If
    Else
        If InStr(objSQL.SQL, vbCrLf) > 0 Then
            '����SQL
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 255)
        Else
            strSQL = ActualStr(objSQL.SQL, 255)
        End If
    End If
    GetLogSQL = strSQL
End Function

Public Function CheckInitFile(ByVal lngSys As Long, ByVal strFile As String, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef rsReturnINI As ADODB.Recordset, Optional ByVal blnUpgradeCheck As Boolean = True) As Boolean
'������blnUpgradeCheck=�����Ǩ����ļ�
   Dim strSysPath As String, strTmp As String
   Dim rsINI As ADODB.Recordset
   If Not gobjFile.FileExists(strFile) Then
        If Not blnOnlyCheck Then MsgBox "��װ�����ļ�""" & strFile & """�����ڡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    If UCase(gobjFile.GetFileName(strFile)) <> IIf(lngSys = 0, "ZLSERVER.SQL", "ZLSETUP.INI") Then
        If Not blnOnlyCheck Then MsgBox "��װ�����ļ�������ȷ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If lngSys = 0 Then '������
        '��������������麯���ļ��Ƿ���ڡ�
        If blnUpgradeCheck Then
            strSysPath = gobjFile.GetParentFolderName(strFile)
            strTmp = strSysPath & "\zlUpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "��������������ļ�""" & strTmp & """�����ڡ�", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    Else 'Ӧ��ϵͳ
        Set rsINI = ReadINIToRec(strFile)
        If Not CheckINIValid(rsINI, "ϵͳ��|�汾��|��ռ�|�����߰汾��") Then
            If Not blnOnlyCheck Then MsgBox "��װ�����ļ���ʽ����ȷ��", vbExclamation, gstrSysName
            Exit Function
        End If
        '�����ļ�ϵͳ�Ų�ƥ��
        rsINI.Filter = "��Ŀ='ϵͳ��'"
        If Val(rsINI!����) <> lngSys \ 100 Then
            If Not blnOnlyCheck Then MsgBox "��ѡ�����ļ����Ǳ�ϵͳ�İ�װ�����ļ���", vbExclamation, gstrSysName
            Exit Function
        End If
        strSysPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strFile))
        'ϵͳ��ǨĿ¼���
        If Not gobjFile.FolderExists(strSysPath & "\�����ű�") Then
            If Not blnOnlyCheck Then MsgBox "ϵͳ��ǨĿ¼""" & strSysPath & "\�����ű�""�����ڡ�", vbExclamation, gstrSysName
            Exit Function
        End If
        If blnUpgradeCheck Then
            '���Ӧ��ϵͳ������麯���ļ��Ƿ���ڡ�
            strTmp = strSysPath & "\�����ű�\zl" & lngSys \ 100 & "_UpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "ϵͳ��������ļ�""" & strTmp & """�����ڡ�", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        '��Ӧ�İ�װ�ű��ļ��Ƿ����,����Ҫ��飬��Ϊ�Ѿ�ȡ���˿�ѡ�ű�ִ��
    End If
    Set rsReturnINI = rsINI
    CheckInitFile = True
End Function

Public Function GetUpgradeFiles(ByVal rsUpgradeFiles As ADODB.Recordset, ByVal lngSys As Long, ByVal strCurVer As String, ByVal strIniPath As String, _
                                                        Optional ByVal strNoramlBreak As String, Optional ByVal strBeforeBreak As String, _
                                                        Optional ByRef strMaxVer As String, Optional ByRef strCurMaxVer As String, Optional ByVal strBakDB As String, _
                                                        Optional ByVal blnReadByMax As Boolean) As ADODB.Recordset
'���ܣ���ȡ����Ҫִ�е��ļ�
'������rsUpgradeFiles=�����ļ���¼���������Ƕ��ϵͳ�������ļ���¼��
'          lngSys=ϵͳ��,=-1��ʾֻ��ʼ����¼��
'          strIniPath=��װ�����ļ�
'          strBreakVers=��Ǩ�����ļ��Ķϵ�汾
'          strBakDB=��ʷ������
'          strMaxVer=���İ汾
'          strCurMaxVer=������Ǩ��Ŀ��汾
'          blnReadByMax=�������汾strMaxVer��ȡ�ű�����Ҫ����ϵͳ��װʱ�����߰汾�ϵ͹����ߵ�������ʱʹ�ã�
'                                   �ò���ΪTrueʱ��������жϵ㴦�����������Ӧ��ϵͳ�ű�����һ��
'����:�����ļ���¼
'        strMaxVer=����Ŀ��汾,����ǰ�ű�������Ǩ��������汾
'        strCurMaxVer=������Ǩ��Ŀ��汾��ϵͳ��Ǩ��������ĳЩ�汾����������Ǩ��������Ҫ�ֶ����Ǩ���ܵ�����Ŀ��汾��
'                               û�в���������Ǩ�İ汾ʱ,�ð汾��strMaxVer��ͬ
'˵����
'        strBakDB="":��ȡ���нű�����ʱ���²�������
'                            strNoramlBreak�����߿⣨lngSys=0��Ϊ�����ߣ�����������ֹ��Ϣ
'                            strBeforeBreak:���߿⣨lngSys=0��Ϊ�����ߣ���ǰ������ֹ��Ϣ
'                            strMaxVer:������������Ǩ������Ŀ��汾
'                            strCurMaxVer:������������Ǩ�ı���Ŀ��汾
'                            ���ص��ļ���¼���и��ڱ�����ǨĿ��汾�Ľű�ȫ���޳���
'        strBakDB<>"":��ȡ����strCurVer���Ҳ�����strMaxVer�Ľű�������������ʷ��Ľű��ļ���¼����
'                             ����ʷ��ǵ�����Ǩʱ�����ɵĽű��ļ���¼����Ҫ��������Ӧ��ϵͳ��ǰ�汾��Ӧ��ϵͳ����Ŀ��汾֮�����ʷ��ű�
'                             ��ʱ���²������壺
'                            strNoramlBreak����ʷ�ⳣ��������ֹ��Ϣ
'                            strBeforeBreak:��ʷ����ǰ������ֹ��Ϣ
'                            strMaxVer:���߿�ĵ�ǰ�汾
    Dim rsCurFiles As ADODB.Recordset, arrFields As Variant, blnNew As Boolean
    Dim strCurPriFull As String, strCurFull As String, strMaxFull As String, strMaxPriFull As String
    Dim cllFolder As New Collection, objFolder As Folder, objFile As File
    Dim strBreak As String, strTmp As String, arrTmp As Variant, strFilter As String
    Dim strFileVer As String, stFile As SysType, ftFile As FileType, vtFile As VersionType, strSetupVer As String, blnSpecial As Boolean
    Dim strFileNameRule As String, stJudge As SysType
    Dim cllSuppleVers As New Collection, Item As Variant
    Dim i As Long
    Dim strFirstBreak As String, strSecdBreak As String
    Dim strBaseSupple As String
    
    On Error GoTo errh
    
    strCurPriFull = VerFull(GetPrimaryVer(strCurVer))
    strCurFull = VerFull(strCurVer)
    strMaxFull = VerFull(strMaxVer, True) '�մ�������9999.9999.9999.9999
    strMaxPriFull = VerFull(GetPrimaryVer(strMaxFull)) '��ֹ�մ�����ʧ�ܣ���˲���strMaxVer����
    If rsUpgradeFiles Is Nothing Then
        blnNew = True
    ElseIf rsUpgradeFiles.State = adStateClosed Then
        blnNew = True
    End If
    
    If blnNew Or lngSys = -1 Then
        '���ð汾:����ǰִ�нű�Ϊ���Ҫ��汾����ӦӦ��ϵͳ���߿���ͨ�����ű�Ϊ��Ӧ�����߽ű�
        Set rsUpgradeFiles = CopyNewRec(Nothing, True, , _
                                                                Array("ϵͳ���", adInteger, 5, Empty, "������", adVarChar, 100, Empty, "SysType", adInteger, 1, Empty, _
                                                                        "FileName", adVarChar, 50, Empty, "FilePath", adVarChar, 1000, Empty, "FileType", adInteger, 1, Empty, _
                                                                        "SPVer", adVarChar, 20, Empty, "FullSPVer", adVarChar, 20, Empty, "VerType", adInteger, 1, Empty, _
                                                                        "Optional", adVarChar, 2000, Empty, "AbortLine", adInteger, 10, Empty, "Special", adInteger, 1, Empty, _
                                                                        "���ð汾", adVarChar, 20, Empty, "�ϵ�", adInteger, 1, Empty))
    End If
    If lngSys = -1 Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    '��ȡ��ǰϵͳ�Ľű�
    rsUpgradeFiles.Filter = "ϵͳ���=" & lngSys & IIf(strBakDB <> "", " And ������='" & UCase(strBakDB) & "'", "")
    '�ű��Ѿ����ڣ��������¶�ȡ��
    '��ʷ���ȡ���������汾��Ϊ�ա���Ϊ��ʷ�ⵥ����Ǩ��Ŀ��汾Ϊ���߿⵱ǰ�汾���ǵ�������ʱ�����߿⵱ǰ�汾֮�ϵ���ʷ�ű��Ѿ���ȡ
    If Not rsUpgradeFiles.EOF Or strBakDB <> "" And strMaxVer = "" Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    Set rsCurFiles = CopyNewRec(rsUpgradeFiles, strBakDB = "")
    '////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////          1����Ǩ�ļ���ȡ            ///////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '��ȡ��Ҫ�Ѽ��ű����ļ���
    If lngSys = 0 Then
        cllFolder.Add gobjFile.GetFile(strIniPath).ParentFolder
        strFileNameRule = "ZLUPGRADE*.*.*.SQL"
    Else
        strFileNameRule = "ZL" & lngSys \ 100 & "_*.*.*.SQL"
        For Each objFolder In gobjFile.GetFolder(gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strIniPath)) & "\�����ű�\").SubFolders
            If IsVerSion(objFolder.name) And objFolder.name Like "*.*.0" Then
                If VerFull(objFolder.name) >= strCurPriFull And VerFull(objFolder.name) <= strMaxPriFull Then
                    cllFolder.Add objFolder
                End If
            End If
        Next
    End If
    arrFields = Array("ϵͳ���", "SysType", "FileName", "FilePath", "FileType", "SPVer", "FullSPVer", "VerType", "Special", "���ð汾")
    '����,��ȡ�ļ�
    For Each objFolder In cllFolder
        If lngSys <> 0 And strBakDB = "" Then '��ȡzlUpgrade.ini
            '��ȡ��Ч�Ķϵ�汾
            strTmp = GetUpgradeIniBreak(objFolder.Path & "\zlUpgrade.ini", IIf(VerFull(objFolder.name) >= strCurPriFull, strCurVer, objFolder.name), GetPrimaryVer(objFolder.name, True))
            If strTmp <> "" Then
                strBreak = strBreak & "," & strTmp
            End If
        End If
        '��ȡ�ļ�
        For Each objFile In objFolder.Files
            If UCase(objFile.name) Like strFileNameRule Then '�����ļ��Ĺ���ĲŽ������ƽ���
                If AnalysisFileName(objFile.name, lngSys, strFileVer, ftFile, stFile, vtFile, blnSpecial) Then
                    If VerFull(strFileVer) > strCurFull And VerFull(strFileVer) <= strMaxFull Then
                        If vtFile = VT_Supple Then
                            On Error Resume Next
                            'ȷ�ϸô�汾�Ѿ���ǵĲ���汾
                            strBaseSupple = cllSuppleVers("K_" & GetPrimaryVer(strFileVer))
                            If err.Number <> 0 Then
                                err.Clear
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            '�Ѿ���ǵĲ���汾С�ڵ�ǰ�汾���򽲱���޸�Ϊ��ǰ�汾
                            ElseIf VerFull(strBaseSupple) > VerFull(strFileVer) Then
                                cllSuppleVers.Remove "K_" & GetPrimaryVer(strFileVer)
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            End If
                            On Error GoTo errh
                        End If
                        '��ȡ���ð汾
                        If ftFile = FT_Before Or ftFile = FT_Standard And stFile = ST_App And VerFull(strFileVer) > VerFull("10.32.0") Then
                            arrTmp = Split(GetUpgradeCtrolInfo(objFile.Path, ftFile = FT_Before) & "|", "|")
                            strSetupVer = VerFull(arrTmp(IIf(ftFile = FT_Before, 0, 1))) '����Ϊ��׼�汾������Ƚ�;    ��ǰִ�з��أ����Ҫ��汾�����������ű����أ���������|��Ӧ�����߰汾
                            '10.34.0֮�󣬹����ߣ�Ӧ��ϵͳ�汾�Ѿ�һһ��Ӧ����û�нű��İ汾�ÿ��ļ�����
                            If ftFile = FT_Standard Then
                                 If VerFull(strFileVer) >= VerFull("10.34.0") Then
                                    strSetupVer = VerFull(strFileVer) '����Ϊ��׼�汾������Ƚ�
                                ElseIf strSetupVer = VerFull("0") Then  '��ȡӦ�ö�Ӧ���߰汾ʧ�ܣ����Զ�����һ��
                                    strSetupVer = VerFull(GetContractVersion(strFileVer, True))
                                End If
                            End If
                            If Val(arrTmp(0)) <> 1 And ftFile = FT_Standard And strBakDB = "" Then strBreak = strBreak & "," & strFileVer
                        Else
                            strSetupVer = ""
                        End If
                        rsCurFiles.AddNew arrFields, Array(lngSys, stFile, objFile.name, objFile.Path, ftFile, strFileVer, VerFull(strFileVer), vtFile, IIf(blnSpecial, 1, 0), strSetupVer)
                    End If
                End If
            End If
        Next
    Next
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////   2.�ϴ���Ǩ��Ϣ���޳�������汾�ϵ���  ///////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '��ǲ���汾
    For Each Item In cllSuppleVers
        '���ڸô�汾����С�Ĳ���汾����С����һ���汾
        Call RecUpdate(rsCurFiles, "FullSPVer>='" & VerFull(Item) & "' And FullSPVer<'" & VerFull(GetPrimaryVer(Item, True)) & "'", "VerType", VT_Supple)
    Next
    stJudge = IIf(lngSys = 0, ST_Tools, IIf(strBakDB = "", ST_App, ST_History))
    strFilter = "SysType=" & stJudge & " And FileType<>" & FT_Deferred
    '�޳���ǰ��ֹ���֮ǰ���ļ�
    arrTmp = Split(strBeforeBreak & "||", "|")
    'û����ֹ�ļ�����С�ڵ�����ֹ�汾����ǰִ�нű���Ҫɾ��������ֻɾ��С����ֹ�汾����ǰ�ű�
    Call RecDelete(rsCurFiles, strFilter & " And FileType=" & FT_Before & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '����ֹ�ļ�����¼��ֹ��
        Call RecUpdate(rsCurFiles, strFilter & "And FileType=" & FT_Before & " And SPVer='" & arrTmp(0) & "'", "AbortLine", Val(arrTmp(2)))
    End If
    arrTmp = Split(strNoramlBreak & "||", "|")
    '�޳�������ֹ���֮ǰ���ļ�
    Call RecDelete(rsCurFiles, strFilter & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '����ֹ�ļ�
        'ɾ����ֹ��ֹ�汾��ִ��˳������ֹ�ļ�֮ǰ���ļ�
        Call RecDelete(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType<" & Val(arrTmp(1)))
        '��¼��ֹ��
        Call RecUpdate(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType=" & Val(arrTmp(1)), "AbortLine", Val(arrTmp(2)))
    End If
    '����������Ǩ�汾�ı��
    strBreak = Mid(strBreak, 2): arrTmp = Split(strBreak, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Call RecUpdate(rsCurFiles, "SPVer='" & arrTmp(i) & "'", "�ϵ�", 1)
    Next
    '���Խű�ִ�У����޳�����汾�벻��Ҫ������SP
    If Not gblnTestUpgrade Then
        '�޳�����汾�����汾���򣬵�һ���ǲ���汾֮ǰ�����в���汾ȫ��ɾ����
        rsCurFiles.Filter = "VerType=" & VT_Normal: rsCurFiles.Sort = "FullSPVer Desc"
        If Not rsCurFiles.EOF Then Call RecDelete(rsCurFiles, "VerType=" & VT_Supple & " And FullSPVer<'" & rsCurFiles!FullSPVer & "'")
    
        '�޳�����SP�ű������汾���򣬵�һ��������SP�汾֮ǰ����������SPȫ��ɾ��
        '�����ж��и����⣬����һ���汾û����ʽ�ű�������������SP�ű�����˲������ִ���
        rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
        If Not rsCurFiles.EOF Then
            strTmp = VerFull(VerSpecialNormal(rsCurFiles!SPVer))
            Call RecDelete(rsCurFiles, "Special=1 And FullSPVer<'" & strTmp & "'")
        End If
    End If
    '////////////////////////////////////////////////////////////////////////////////////
    '/////////////// 3������Ŀ��汾������Ŀ��汾���Լ���ʷ��ű��Ķ�ȡ ////////////
    '///////////////////////////////////////////////////////////////////////////////////
    If strBakDB = "" Then
        If blnReadByMax Then '�������汾��ȡ
            '��ȡʵ�ʿ��������������汾
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strCurMaxVer = rsCurFiles!SPVer & ""
            End If
        Else
            '��ȡ����Ŀ��汾�Լ�����Ŀ��汾
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strMaxVer = "": strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strMaxVer = rsCurFiles!SPVer & ""
                rsCurFiles.Filter = "�ϵ�=1": rsCurFiles.Sort = "FullSPVer"
                If Not rsCurFiles.EOF Then
                    strFirstBreak = rsCurFiles!SPVer
                    If rsCurFiles.RecordCount > 1 Then
                        rsCurFiles.MoveNext: strSecdBreak = rsCurFiles!SPVer
                    End If
                    rsCurFiles.Filter = "FullSPVer<'" & VerFull(strFirstBreak) & "'"
                    strCurMaxVer = IIf(rsCurFiles.EOF, strSecdBreak, strFirstBreak)
                End If
            End If
            If strCurMaxVer = "" Then
                strCurMaxVer = strMaxVer
            Else 'ɾ������Ҫ������Ǩ����Ҫִ�еĽű�
                Call RecDelete(rsCurFiles, "FullSPVer>'" & VerFull(strCurMaxVer) & "'")
            End If
        End If
    Else
    '��ȡ��ʷ����Ǩ��¼
        'ɾ��С����ʷ�⵱ǰ�汾�Ľű�����ʷ��汾���ܸ������߿⣬�����Ҫ��������
        Call RecDelete(rsCurFiles, "FullSPVer<='" & VerFull(strCurVer) & "'")
        'ɾ�����߿�ű�
        Call RecDelete(rsCurFiles, "SysType<>" & ST_History)
        '�����ļ���¼����������
        Call RecUpdate(rsCurFiles, "", "������", UCase(strBakDB))
    End If
    '�ϲ���¼���������ζ�ȡ���ļ��ϲ������м�¼����
    rsCurFiles.Filter = ""
    Call RecDataAppend(rsUpgradeFiles, rsCurFiles)
    Set GetUpgradeFiles = rsUpgradeFiles
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function FormatUpgradeBreak(ByVal lngSys As Long, ByVal strResultVer As String, Optional ByVal strUpgradeBreak As String) As String
'���ܣ�������ֹ��Ϣ������ֹ����׼�� ��ʽ���ļ��汾|�ļ�����|�����к�
'������
'     strResultVer:ZLUpgrade�еĽ���汾
'     strUpgradeBreak=��Ǩ��ֹ���
'���أ��ļ��Ĳ���·�����ļ���
    Dim arrTmp As Variant
    Dim lngPos As Long
    Dim strTmp As String
    Dim strFileName As String
    Dim lngAbort As Long
    Dim strFileVer As String '���ļ����϶�ȡ�İ汾��Ϣ
    Dim ftReturn As FileType
    Dim strReturn As String
    
    strReturn = strResultVer & "||"
    If strUpgradeBreak <> "" Then
        '��ʷ�����ֹ������Ϊ�汾��
        If Not IsVerSion(strUpgradeBreak) Then
            strUpgradeBreak = strUpgradeBreak & "||"
            arrTmp = Split(strUpgradeBreak, "|")
            If gobjFile.FileExists(arrTmp(0)) Then
                strFileName = gobjFile.GetFileName(arrTmp(0))
            Else '�����ǲ���汾�Ѿ�ɾ����
                strTmp = StrReverse(arrTmp(0))
                lngPos = InStr(strTmp, "\")
                '��ȡ���һ��\�������
                If lngPos <> 0 Then
                    strFileName = StrReverse(Mid(strTmp, lngPos - 1))
                Else
                    strFileName = ""
                End If
            End If
            lngAbort = Val(arrTmp(1))
            If strFileName <> "" Then
                If AnalysisFileName(strFileName, lngSys, strFileVer, ftReturn) Then
                    strReturn = strFileVer & "|" & ftReturn & "|" & lngAbort
                End If
            End If
        Else '��ʷ����ǰ������ŵ��ǰ汾��
            strReturn = strUpgradeBreak & "||"
        End If
    End If
    FormatUpgradeBreak = strReturn
End Function

Public Function GetUpgradeIniBreak(ByVal strFile As String, Optional ByVal strMinVer As String, Optional ByVal strMaxVer As String)
'���ܣ���ȡ��Ǩ�����ļ��Ķϵ�
'������strFile=��Ǩ�����ļ�·��
'          strMinVer=��Ǩ�����ļ�Ŀ��汾����Сֵ
'          strMaxVer=��Ǩ�����ļ�Ŀ��汾�����ֵ
    Dim rsSub As ADODB.Recordset
    Dim strBreakVer As String
    
    If Not gobjFile.FileExists(strFile) Then Exit Function
    Set rsSub = ReadINIToRec(strFile)
    If rsSub Is Nothing Then Exit Function
    rsSub.Filter = "��Ŀ='��������'" '���������ļ���Ŀ��汾�Ƿ�����������
    If rsSub.EOF Then Exit Function
    If Val(rsSub!���� & "") = 1 Then Exit Function '�����������ô���
    rsSub.Filter = "��Ŀ='Ŀ��汾'" '���������ļ���Ŀ��汾
    If rsSub.EOF Then Exit Function
    strBreakVer = Trim(rsSub!���� & "")
    If Not IsVerSion(strBreakVer) Then Exit Function
    If strMinVer <> "" Then 'С����С�汾����öϵ���Ч
        If VerFull(strBreakVer) <= VerFull(strMinVer) Then Exit Function
    End If
    If strMaxVer <> "" Then '������С�汾����öϵ���Ч
        If VerFull(strBreakVer) > VerFull(strMaxVer) Then Exit Function
    End If
    GetUpgradeIniBreak = strBreakVer
End Function

Public Function GetUpgradeCtrolInfo(ByVal strFile As String, Optional ByVal blnBefore As Boolean) As String
'���ܣ���ȡ�ļ��еĿ�����Ϣ
'      strFile=�����жϵĽű��ļ�·��
'      blnBefore=�ļ��Ƿ�������ִ�нű�
'����: blnBefore=false: ��������|�����߰汾��
'        blnBefore=True: ��Ͱ汾��

    Dim objStream As Scripting.TextStream
    Dim strLine As String, arrFind() As Variant, i As Long, strTmp As String, arrTmp As Variant
    Dim strContinue As String, strToolVer As String, strBreakVer As String, strReqVer As String
    Dim rsSub As ADODB.Recordset
    
    On Error GoTo errh
    
    Set objStream = gobjFile.OpenTextFile(strFile, ForReading)
    If blnBefore Then
        arrFind = Array("[[]��Ͱ汾��[]]")
    Else
        arrFind = Array("[[]��������[]]", "[[]�����߰汾��[]]")
    End If
    Do While Not objStream.AtEndOfStream
        strLine = TrimEx(objStream.ReadLine, True)
        If strLine Like "--" & arrFind(i) & "*" Then
            strTmp = Trim(Mid(strLine, Len("--" & arrFind(i)) - 4 + 1))
            If Not blnBefore Then
                If i = 0 Then
                    strContinue = strTmp
                Else
                    strToolVer = strTmp
                End If
            Else
                strReqVer = strTmp
            End If
        End If
        If i = UBound(arrFind) Then Exit Do
        i = i + 1
    Loop
    objStream.Close
    
    If blnBefore Then
        GetUpgradeCtrolInfo = Trim(strReqVer)
    Else
        If Trim(strContinue) = "" Then strContinue = "1"
        GetUpgradeCtrolInfo = Trim(strContinue) & "|" & Trim(strToolVer)
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
'    Debug.Print err.Source & "\" & Me.name & "\GetCtrolInfo:" & err.Description
End Function


Public Function AnalysisFileName(ByVal strFileName As String, ByVal lngSys As Long, Optional ByRef strVersion As String, Optional ByRef ftReturn As FileType, _
                                                        Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType = VT_Normal, Optional ByRef blnSpecial As Boolean) As Boolean
'����:tͨ���ļ�����ȡ�ļ���Ϣ
'������
'   strFile=������·�����ļ���,����չ��
'   lngSys=ϵͳ��
'����:
'       True=�ɹ���ȡ��False=��ȡʧ�ܣ��ļ�����ϵͳ�����ű���
'       strVerReturn=�ļ��汾
'       ftReturn=�ļ�����
'       stReturn=ϵͳ����
'       vtReturn=�汾����
    Dim strSysString As String, strSuffix As String
    Dim arrVer As Variant
    vtReturn = VT_Normal
    blnSpecial = False
    strVersion = ""
    ftReturn = FT_Before
    stReturn = ST_Tools
    If Not UCase(strFileName) Like "*.SQL" Then Exit Function
    strFileName = UCase(Left(strFileName, Len(strFileName) - 4))
    arrVer = Split(strFileName, ".")
    '�汾�ļ����ļ�������2������(����SP����3����
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    '��ȡ�ű�ϵͳǰ׺
    If arrVer(0) Like "ZLUPGRADE*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf arrVer(0) Like "ZL" & lngSys \ 100 & "_*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function 'û��ϵͳ��ʶǰ׺������ϵͳ�ű�
    End If
    'ϵͳ��ʶ����������ǰ汾
    arrVer(0) = Mid(arrVer(0), Len(strSysString) + 1) '��ȡ���屾
    arrVer(UBound(arrVer)) = GetPrefixNumber(arrVer(UBound(arrVer)), strSuffix) '��ȡ�μ��汾
    '��ȡ�����屾����汾�Լ��μ��汾����Ϊ���֣����˳�
    If Not IsNumeric(arrVer(0)) Or Not IsNumeric(arrVer(1)) Or Not IsNumeric(arrVer(2)) Or Not IsNumeric(arrVer(UBound(arrVer))) Then Exit Function
    strVersion = arrVer(0) & "." & arrVer(1) & "." & arrVer(2) & IIf(UBound(arrVer) = 2, "", "." & arrVer(UBound(arrVer)))
    If Not IsVerSion(strVersion) Then Exit Function
    '��λ�汾�ž�������SP
    blnSpecial = strVersion Like "*.*.*.*"
    '�汾�����ļ�������Ϣ
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_History
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    If strSuffix Like "*(����)" Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(����)", "") '��ֹ������Ϣλ�ò��̶�
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_Standard
        Case "_DBA"
            If stReturn = ST_History Then Exit Function '��ʷ�ⲻ֧��DBA�ű�
            ftReturn = FT_DBA
        Case "_OPTIONAL"
            ftReturn = FT_Optional
        Case "_BEFORE"
            ftReturn = FT_Before
        Case "_DEFERRED"
            If stReturn = ST_Tools Then Exit Function '�����߲�֧���ӳ�ִ�нű�
            ftReturn = FT_Deferred
        Case Else '������������Χ�ڵģ����ȡʧ��
            Exit Function
    End Select
    AnalysisFileName = True
End Function

Public Function GetPrefixNumber(ByVal strInput As String, Optional ByRef strOther As String) As String
'���ܣ���ȡһ���ַ���������ǰ׺���Լ�ʣ�ಿ��
'������strInput=������ַ���
'          strOther =ȥ������ǰ׺��ʣ�ಿ��
    Dim i As Long
    
    For i = 1 To Len(strInput)
        If Not IsNumeric(Mid(strInput, i, 1)) Then
            Exit For
        End If
    Next
    strOther = Mid(strInput, i)
    GetPrefixNumber = Mid(strInput, 1, i - 1)
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999.9999,��С�汾��0000.0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '����һ�Σ��Լ�������SP�汾��
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function VerPAD(ByVal strVer As String) As String
'���ܣ�ʹ�汾�ŵ����汾�������Ϊ4λ����֤���汾��ԭ������������汾�Ŷ���
'������strVer=��ǰ�汾��
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        Exit Function
    End If
    arrVer = Split(strVer & ".", ".")
    VerPAD = RPAD(Lpad(arrVer(0), 2) & "." & arrVer(1) & "." & arrVer(2) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000")), 20)
End Function

Public Function GetPrimaryVer(ByVal strVer As String, Optional ByVal blnNext As Boolean)
'���ܣ���ȡһ���汾�����汾
'������strVer=��ǰ�汾
'          blnNext=�Ƿ��ȡ��һ�����汾
'���أ����汾
    Dim arrVer As Variant
    
    arrVer = Split(strVer & "..", ".")
    If blnNext Then
        GetPrimaryVer = Val(arrVer(0)) & "." & (Val(arrVer(1)) + 1) & "." & 0
        '������û��9.45.0��ֱ�Ӻ�Ӧ��ϵͳͬһ��ţ�Ϊ10.34.0
        If GetPrimaryVer = "9.45.0" Then GetPrimaryVer = "10.34.0"
    Else
        GetPrimaryVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & 0
    End If
End Function

Public Function GetContractVersion(ByVal strVer As String, Optional ByVal blnGetTools As Boolean = True)
'���ܣ���ȡӦ��ϵͳ��Ӧ�����ߵ����汾�����߹����߶�ӦӦ��ϵͳ�汾����Ҫ����
'������strVer=��ǰӦ��ϵͳ�汾
'          blnGetTools=True-��ȡ��Ӧ�Ĺ����߰汾,False-��ȡ��Ӧ��Ӧ��ϵͳ�汾
'���أ���Ӧ�汾��Ӧ��ϵͳ10.34.0֮ǰ��ֻ���Ӧ��汾�������嵽SP�汾
'                          ������10.34.0֮ǰ��ֻ���Ӧ��汾�������嵽SP�汾
    Dim arrVer As Variant
    Dim lngDistance As Long
    If strVer = "" Then strVer = "9.1.0"
    If blnGetTools Then
        If VerFull(strVer) >= VerFull("10.34.0") Then '10.34.0  �Ժ�����ߺ�Ӧ��ϵͳ�汾ͳһ
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 33 - Val(arrVer(1)) '��ȡӦ��ϵͳ��10.33.0�汾�Ĵ�汾���
            '������9.44.0��ȥ��Ӧ��汾�����Ϊ��Ӧ�����߰汾
            GetContractVersion = "9." & (44 - lngDistance) & ".0"
        End If
    Else
        If VerFull(strVer) >= VerFull("10.34.0") Then  '  �Ժ�����ߺ�Ӧ��ϵͳ�汾ͳһ
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 44 - Val(arrVer(1)) '��ȡ��������9.44.0�汾�Ĵ�汾���
            'Ӧ��ϵͳ10.33.0��ȥ��Ӧ��汾�����Ϊ��ӦӦ��ϵͳ�İ汾
            GetContractVersion = "10." & (33 - lngDistance) & ".0"
        End If
    End If
End Function

Public Function VerNormal(ByVal strVer As String) As String
'���ܣ���VB���֧�ֵİ汾����ʽ:9999.9999.9999ת��Ϊ�����汾����ʽ����0010.0034.0000.0000��ת��Ϊ10.34.0
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000"))
End Function

Public Function VerSpecialNormal(ByVal strVer As String) As String
'��ȡһ������sp��Ӧ����ʽ�汾�������һ����ʽ�汾���򷵻�������
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerSpecialNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

Public Function ReadHisUpgrade(ByVal cnHistory As ADODB.Connection, ByVal strOwner As String, Optional ByVal blnMsg As Boolean, Optional ByVal lngSys As Long, Optional ByVal blnDB_LINK As Boolean) As ADODB.Recordset
'����:��ȡ��ʷ��ռ�ĸ�ϵͳ��Ϣ��Ǩ��Ϣ
'������ cnHistory=��ʷ������
'           strOwner=������
'           lngSys=ϵͳ���=0����ȡ����ʷ�������ϵͳ��<>0:����ø�ϵͳ��ʷ��
'           blnDB_LINK=�Ƿ���DBLINK����
    Dim rsReturn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim objTmp As New clsObjectInfo, objParent As clsObjectInfo, objCur As clsObjectInfo
    Dim strSQL As String
    
    On Error GoTo errh
    Set rsReturn = CopyNewRec(Nothing, True, , Array("ϵͳ���", adInteger, Empty, Empty, "��ǰ�汾", adVarChar, 20, Empty, "��ֹ��Ϣ", adVarChar, 2000, Empty, _
                                                                                    "��ǰ��ֹ��Ϣ", adVarChar, 2000, Empty, "��ǰִ��", adInteger, 1, Empty))
    'zlbakinfo���
    Set objParent = objTmp.GetObject("zlbakinfo", OT_Table, , strOwner, , cnHistory)
    '��ǰִ���м��
    Set objCur = objTmp.GetObject("��ǰִ��", OT_Column, "alter Table zlbakinfo add ��ǰִ�� number(1)", strOwner, objParent, cnHistory)
    If Not objCur.ObjectCheck(blnMsg) Then
        GoTo ExitCode
    End If
    '��ǰ��ֹ����м��
    Set objCur = objTmp.GetObject("��ǰ��ֹ���", OT_Column, "alter Table zlbakinfo add ��ǰ��ֹ��� VarChar2(500)", strOwner, objParent, cnHistory)
    If Not objCur.ObjectCheck(blnMsg) Then
        GoTo ExitCode
    End If
'    '����ZLBAKInfo��ͼ
'    strSQL = "create or replace view " & strOwner & ".zlbakinfo as" & vbNewLine & _
'        "Select ""ϵͳ"",""�汾��"",""��������"",""���ת������"",""���������"",""��ֹ���"",""��ǰִ��"",""��ǰ��ֹ���"" From " & strOwner & ".ZLBAKINFO"
'    cnHistory.Execute strSQL
    '��Ȩ
    If Not blnDB_LINK Then
        strSQL = "Grant Select On  " & strOwner & ".zlbakinfo To " & gstrUserName
        cnHistory.Execute strSQL
    End If
    '���ɼ�¼����Ϣ
    strSQL = "Select ϵͳ,�汾��,��ֹ���,��ǰִ��,��ǰ��ֹ���  from zlbakinfo " & IIf(lngSys = 0, "", "Where ϵͳ=" & lngSys) & " order by ϵͳ"
    Set rsTmp = gclsBase.OpenSQLRecord(cnHistory, strSQL, "��ȡ��ʷ����ϵͳ��Ϣ")
    Do While Not rsTmp.EOF
        rsReturn.AddNew Array("ϵͳ���", "��ǰ�汾", "��ֹ��Ϣ", "��ǰ��ֹ��Ϣ", "��ǰִ��"), _
                                    Array(rsTmp!ϵͳ, rsTmp!�汾��, FormatUpgradeBreak(rsTmp!ϵͳ, rsTmp!�汾�� & "", rsTmp!��ֹ��� & ""), FormatUpgradeBreak(rsTmp!ϵͳ, rsTmp!�汾�� & "", rsTmp!��ǰ��ֹ��� & ""), rsTmp!��ǰִ��)
        rsTmp.MoveNext
    Loop
    Set ReadHisUpgrade = rsReturn
    Exit Function
errh:
    Set ReadHisUpgrade = rsReturn
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
    Exit Function
ExitCode:
    If 0 = 1 Then
        Resume
    End If
    Set ReadHisUpgrade = rsReturn
End Function

Public Function CheckHavHistory(ByVal lngSys As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:����Ƿ���Ҫ������ʷ�ռ䣨�����ڴ�ת����
    '����:lngSys-ϵͳ��
    '����:��Ҫ����,��true,����False
    '--------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    strSQL = "Select 1 from zltools.zlbakTables where ϵͳ=[1] and rownum<=1"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡbak����", lngSys)
    If rsTmp.EOF Then
       '����False,��ʾ��ϵͳû����ʷ���ݿռ�,û��Ҫ������ʷ���ݿռ�
       Exit Function
    End If
    CheckHavHistory = True
End Function

Public Function GrantBakToUser(ByVal cnOracle As ADODB.Connection, ByVal strToOwner As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ����
    '����:strTableName-����
    '     cnoracle-���ݿ�������
    '     strOwNer-������
    '����:���ڸñ���true,����False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    err = 0: On Error GoTo ErrHand:
    strSQL = "Select TABLE_NAME from user_all_tables" & _
            " Union All Select View_Name From User_Views"
    Call OpenRecordset(rsTemp, strSQL, "������Ȩ", , , cnOracle)
    With rsTemp
        Do While Not .EOF
            strSQL = "Grant ALL on " & Nvl(!Table_Name) & " to " & strToOwner & " With Grant Option"
            cnOracle.Execute strSQL
            .MoveNext
        Loop
    End With
    GrantBakToUser = True
    Exit Function
ErrHand:
    If MsgBox("����Ȩʱ�������´���,����!" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then
        Resume
    End If
    GrantBakToUser = False
End Function

Public Sub ReGrantToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String, Optional ByRef objProcess As Object, Optional ByRef objlblPer As Object, Optional ByRef lngRoleCount As Long)
'���ܣ��Խ�ɫ����������Ȩ��
'������cnOracle=����
'      strRoleNames=������Ȩ�Ľ�ɫ����Ϊ�գ���Ϊ���н�ɫ������Ȩ����Ϊ��ʱ�������ɫ�Զ��÷ָ��ɫ������15����
'      blnGrantBase=�Ƿ������ֵ������Ȩ��
'      strOwners=��Ȩ��ϵͳ��������
'      objProcess=����
    Dim rsPrivs As ADODB.Recordset, rsRoles As ADODB.Recordset
    Dim strRolePars As String
    Dim strSQL As String, i As Long
    Dim lngMax As Long, lngCur As Long
    Dim blnProcess As Boolean
    
    On Error GoTo errh
    blnProcess = Not objProcess Is Nothing
    If strRoleNames = "" Then
    
        '��ǰ��SQL�����Բ�ѯ��47811������,���ھ��Ż�����ѯ��3057�����ݣ���Distinct��Ҫ��һ����ɫ�ж�����ܣ����ܷ��ʵı�֮�����ص�
        '��ɫ����31���汾��10.29.30���Ż�ǰ������ɫ��Ȩ��ʱ138�룬�Ż���19��
        strSQL = "Select Ȩ��, ����, ������, f_List2str(Cast(Collect(��ɫ) As t_Strlist)) As ��ɫ" & vbNewLine & _
                "From (Select ����, ������, Ȩ��, ��ɫ, Floor(Row_Number() Over(Partition By Ȩ��, ����, ������ Order By ��ɫ) / 10) Rn" & vbNewLine & _
                "       From (Select Distinct Upper(p.����) ����, p.������, Upper(p.Ȩ��) Ȩ��, r.��ɫ" & vbNewLine & _
                "              From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
                "              Where Nvl(P.ϵͳ, 0) = Nvl(R.ϵͳ, 0) And P.������ = User And P.��� = R.��� And P.���� = R.���� And R.��ɫ = U.Granted_Role))" & vbNewLine & _
                "Group By Ȩ��, ����, ������, Rn"
    '    'ԭ��SQL
    '    strSQL = "Select P.����, P.������, P.Ȩ��, R.��ɫ" & vbNewLine & _
    '            "From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
    '            "Where Nvl(P.ϵͳ, 0) = Nvl(R.ϵͳ, 0) And P.������ = User And P.��� = R.��� And P.���� = R.���� And R.��ɫ = U.Granted_Role"
        Set rsPrivs = gclsBase.OpenSQLRecord(cnOracle, strSQL, "��ɫ��Ȩ")
        
        strSQL = "Select F_List2str(Cast(Collect(��ɫ) As T_Strlist)) ��ɫ, ��ɫ��" & vbNewLine & _
                "From (Select Floor(Rownum / 10) Rn, ��ɫ, Count(1) Over(Partition By ����) ��ɫ��" & vbNewLine & _
                "       From (Select Distinct R.��ɫ, 1 ����" & vbNewLine & _
                "              From zlRoleGrant R" & vbNewLine & _
                "              Where Exists" & vbNewLine & _
                "               (Select 1" & vbNewLine & _
                "                     From zlProgPrivs P" & vbNewLine & _
                "                     Where Nvl(P.ϵͳ, 0) = Nvl(R.ϵͳ, 0) And P.������ = User And P.��� = R.��� And P.���� = R.����) And Exists" & vbNewLine & _
                "               (Select 1 From User_Role_Privs U Where R.��ɫ = U.Granted_Role)" & vbNewLine & _
                "              Order By R.��ɫ))" & vbNewLine & _
                "Group By Rn, ��ɫ��"
'    'ԭ��SQL
'    strSQL = "Select Distinct R.��ɫ" & vbNewLine & _
'            "From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
'            "Where Nvl(P.ϵͳ, 0) = Nvl(R.ϵͳ, 0) And P.������ = User And P.��� = R.��� And P.���� = R.���� And R.��ɫ = U.Granted_Role" & vbNewLine & _
'            "Order By R.��ɫ"
        Set rsRoles = gclsBase.OpenSQLRecord(cnOracle, strSQL, "��ɫ��Ȩ")
    Else
        strRolePars = "'" & Replace(UCase(strRoleNames), ",", "','") & "'"
        strSQL = "Select Ȩ��, ����, ������, f_List2str(Cast(Collect(��ɫ) As t_Strlist)) As ��ɫ" & vbNewLine & _
                "From (Select ����, ������, Ȩ��, ��ɫ, Floor(Row_Number() Over(Partition By Ȩ��, ����, ������ Order By ��ɫ) / 10) Rn" & vbNewLine & _
                "       From (Select Distinct Upper(p.����) ����, p.������, Upper(p.Ȩ��) Ȩ��, r.��ɫ" & vbNewLine & _
                "              From Zlprogprivs p, Zlrolegrant r" & vbNewLine & _
                "              Where Nvl(p.ϵͳ, 0) = Nvl(r.ϵͳ, 0) And p.������ = User And p.��� = r.��� And p.���� = r.���� And r.��ɫ in(" & strRolePars & ")))" & vbNewLine & _
                "Group By Ȩ��, ����, ������, Rn"
        Set rsPrivs = gclsBase.OpenSQLRecord(cnOracle, strSQL, "��ɫ��Ȩ")
        
        strSQL = "Select f_List2str(Cast(Collect(��ɫ) As t_Strlist)) ��ɫ, ��ɫ��" & vbNewLine & _
                "From (Select Floor(Rownum / 10) Rn, ��ɫ, Count(1) Over(Partition By ����) ��ɫ��" & vbNewLine & _
                "       From (Select Distinct r.��ɫ, 1 ����" & vbNewLine & _
                "              From Zlrolegrant r" & vbNewLine & _
                "              Where Exists (Select 1" & vbNewLine & _
                "                     From Zlprogprivs p" & vbNewLine & _
                "                     Where Nvl(p.ϵͳ, 0) = Nvl(r.ϵͳ, 0) And p.������ = User And p.��� = r.��� And p.���� = r.����) And" & vbNewLine & _
                "                    r.��ɫ In (" & strRolePars & ")" & vbNewLine & _
                "              Order By r.��ɫ))" & vbNewLine & _
                "Group By Rn, ��ɫ��"
        Set rsRoles = gclsBase.OpenSQLRecord(cnOracle, strSQL, "��ɫ��Ȩ")
    End If
    On Error Resume Next
    lngMax = rsPrivs.RecordCount + 25 * rsRoles.RecordCount
    For lngCur = 1 To rsPrivs.RecordCount
        If blnProcess Then
            objProcess.value = lngCur / lngMax * 100
            objlblPer.Caption = Format(objProcess.value / 100, "0%")
        End If
        DoEvents
        cnOracle.Execute "Grant " & rsPrivs!Ȩ�� & " on " & rsPrivs!������ & "." & rsPrivs!���� & " to " & rsPrivs!��ɫ
        rsPrivs.MoveNext
    Next
    If rsRoles.RecordCount <> 0 Then
        lngRoleCount = Val(rsRoles!��ɫ�� & "")
    End If
    For i = 1 To rsRoles.RecordCount
        lngCur = i * 25
        If blnProcess Then
            objProcess.value = lngCur / lngMax * 100
            objlblPer.Caption = Format(objProcess.value / 100, "0%")
        End If
        DoEvents
        Call GrantSpecialToRole(cnOracle, rsRoles!��ɫ & "", True, strOwners, True)
        rsRoles.MoveNext
    Next
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub



Public Sub ReGrantForTools(ByVal cnTools As ADODB.Connection, Optional ByVal strSysOwner As String, Optional ByVal blnSysGrant As Boolean)
    '----------------------------------------------------------------------------------------------------------
    '����:�Թ����ߵĶ������������Ȩ������ͬ���
    '����:cnTools�����������ӡ�strSysOwnerΪ��ʱ�����Դ�Ӧ��ϵͳ���ӣ���ʱΪӦ��ϵͳת��Ȩ�ޡ�
    '     strSysOwner:Ӧ��ϵͳ�����ߡ�Ϊ���Ƿ������������ã�ֻ��������ͬ����Լ���Public��Ȩ���ǿ�ʱ�Ը��û�����ZLTOOLS����Ȩ������
    '     blnSysGrant:ϵͳ������ת�ڹ�����Ȩ�ޣ����ǰ׺ZLTOOLS.,��strSysOwnerΪ���Ҹò���ΪTrueʱ������ϵͳ�����߽�����Ȩ
    '����:
    '----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim strSysSQL As String
    
    On Error Resume Next
    '����ͬ���ȱʧ�Ķ���Ĭ�ϴ��ֶ����Ȩ��Ҳȱʧ
    strSQL = "Select Object_Name" & vbNewLine & _
                    "From ((Select Object_Name" & vbNewLine & _
                    "        From User_Objects" & vbNewLine & _
                    "        Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                    "              Instr(Object_Name, 'BIN$') <= 0) Minus" & vbNewLine & _
                    "       (Select Synonym_Name From All_Synonyms Where Owner = 'PUBLIC' And Table_Owner = 'ZLTOOLS'))"
    Call OpenRecordset(rsTemp, strSQL, "�����߹���ͬ���ȱʧ����", , , cnTools)
    For i = 1 To rsTemp.RecordCount
        cnTools.Execute "Create Public Synonym " & rsTemp!Object_Name & " For ZLTOOLS." & rsTemp!Object_Name
        If err.Number <> 0 Then '���ܴ��������û���ͬ��ʣ���������ZLtools,���ɾ����
            err.Clear
            cnTools.Execute "Drop Public Synonym " & rsTemp!Object_Name
            cnTools.Execute "Drop  Synonym " & rsTemp!Object_Name
            cnTools.Execute "Create Public Synonym " & rsTemp!Object_Name & " For ZLTOOLS." & rsTemp!Object_Name
            If err.Number <> 0 Then err.Clear
        End If
        rsTemp.MoveNext
    Next
    '����PublicȨ��ȱʧ
    strSQL = "Select Object_Name,Privilege" & vbNewLine & _
                    "From ((Select Object_Name," & vbNewLine & _
                    "               Decode(Object_Type, 'SEQUENCE', 'SELECT', 'TABLE', 'SELECT', 'VIEW', 'SELECT', 'EXECUTE') Privilege" & vbNewLine & _
                    "        From User_Objects" & vbNewLine & _
                    "        Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                    "              Instr(Object_Name, 'BIN$') <= 0 And" & vbNewLine & _
                    "              Object_Name Not In ('B_ROLEGROUPMGR', 'ZL_ZLROLEGRANT_BATCHDELETE', 'ZL_ZLROLEGRANT_BATCHINSERT')) Minus" & vbNewLine & _
                    "       (Select Table_Name, Privilege" & vbNewLine & _
                    "        From User_Tab_Privs" & vbNewLine & _
                    "        Where Grantee = 'PUBLIC' And Grantor = 'ZLTOOLS' And Instr(Table_Name, 'BIN$') <= 0))"
    Call OpenRecordset(rsTemp, strSQL, "������PublicȨ��ȱʧ����", , , cnTools)
    For i = 1 To rsTemp.RecordCount
        cnTools.Execute "Grant " & rsTemp!Privilege & " On ZLTOOLS." & rsTemp!Object_Name & " To Public"
        rsTemp.MoveNext
    Next
    
    If err.Number <> 0 Then err.Clear
    
    If strSysOwner <> "" Then
        strSysSQL = "Select '" & Trim(UCase(strSysOwner)) & "' ������ FROM Dual"
    ElseIf blnSysGrant Then
        '�����߶���������Ȩ����������ͬ���,��Ӧ��ϵͳ����ʷ�ⶼ��Ȩ
        strSysSQL = "Select Distinct ������" & vbNewLine & _
                "From Zlsystems" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select Distinct ������" & vbNewLine & _
                "From Zlbakspaces a" & vbNewLine & _
                "Where Exists (Select 1 From All_Users b Where b.Username = Upper(a.������))"
    End If
    If strSysSQL <> "" Then
        'Ӧ��ϵͳ������ȱʧ����Ȩ�ޣ��򻯴���
        strSQL = "Select b.������ Grantee, a.Object_Name, 'ALL' Privilege" & vbNewLine & _
                        "From User_Objects a, (" & strSysSQL & ") b" & vbNewLine & _
                        "Where a.Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                        "      Instr(a.Object_Name, 'BIN$') <= 0" & vbNewLine & _
                        "Minus" & vbNewLine & _
                        "Select Grantee, Table_Name, 'ALL' Privilege" & vbNewLine & _
                        "From User_Tab_Privs a, (" & strSysSQL & ") b" & vbNewLine & _
                        "Where Grantee =b.������  And Grantor = 'ZLTOOLS' And Grantable = 'YES' And Instr(Table_Name, 'BIN$') <= 0 And" & vbNewLine & _
                        "      Privilege In ('EXECUTE', 'INSERT')"
        Call OpenRecordset(rsTemp, strSQL, "Ӧ��ϵͳ�����߶���Ȩ������", , , cnTools)
        
        For i = 1 To rsTemp.RecordCount
            cnTools.Execute "Grant " & rsTemp!Privilege & " On ZLTOOLS." & rsTemp!Object_Name & " To " & rsTemp!Grantee & " With Grant Option"
            rsTemp.MoveNext
        Next
    End If
    If err.Number <> 0 Then err.Clear: cnTools.Errors.Clear
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Public Function GrantSpecialToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String, Optional ByVal blnCreateRole As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:�Թ����ߵĶ����Ӧ�ó���һЩ���������Ȩ������Ķ���
    '����:cnOracle��Ӧ��ϵͳ����
    '     strRoleNames:����Ȩ�Ľ�ɫ�������ɫ�Զ��ŷָһ�㲻����15����ɫ
    '     blnGrantBase:�Ƿ��Ӧ��ϵͳ�����������Ȩ
    '     strOwners��Ӧ��ϵͳ������
    '     blnCreateRole=�Ƿ񴴽���ɫ��������ɫ�����蹫��������������SQL���½����������û�ҵ��ͣ��
    '����:
    '----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnsysSt As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errh
    strSQL = "Select ������ From zlSystems Where Floor(��� / 100) = 1"
    If UBound(strOwners) <> -1 And mstrStSysOwner = "" Then
        OpenRecordset rsTmp, strSQL, "��ȡ��׼��ϵͳ������", , , cnOracle
        Do While Not rsTmp.EOF
            mstrStSysOwner = mstrStSysOwner & "," & rsTmp!������
            rsTmp.MoveNext
        Loop
        If mstrStSysOwner <> "" Then mstrStSysOwner = mstrStSysOwner & ","
    End If
    On Error Resume Next
    For i = LBound(strOwners) To UBound(strOwners)
        If strOwners(i) <> "" Then
            If blnCreateRole Then
                cnOracle.Execute "grant select on " & strOwners(i) & ".���ű� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա�� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".������Ա to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".�ϻ���Ա�� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա����˵�� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա���ʷ��� to " & strRoleNames
            End If
            
            If InStr(mstrStSysOwner, "," & strOwners(i) & ",") > 0 Then
                If blnCreateRole Then
                    '��Ϣƽ̨����
                    cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ���� to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ�嵥 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ���Ѳ��� to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ������Ա to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ״̬ to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".������������Ŀ¼ to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zlpub_ҵ����Ϣ�嵥_insert to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_ҵ����Ϣ�嵥_insert to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_ҵ����Ϣ�嵥_read to " & strRoleNames
                End If
            End If
            If blnGrantBase Then
                cnOracle.Execute "grant execute on " & strOwners(i) & ".zl_�ֵ����_execute to " & strRoleNames
            End If
        End If
    Next
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    If blnCreateRole Then
        '�Է������ļ��������������Ȩ
        '------------------------------------------------------------------------------------------------------------------
        '������������쳣
        cnOracle.Execute "grant delete                on ZLTOOLS.zluserparas to " & strRoleNames
        '�ͻ��������ݴ���
        cnOracle.Execute "grant update                on ZLTOOLS.zlclients   to " & strRoleNames
        cnOracle.Execute "grant insert,update         on ZLTOOLS.zlDiaryLog to " & strRoleNames
        cnOracle.Execute "grant insert                on ZLTOOLS.zlErrorLog to " & strRoleNames
        cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMessages to " & strRoleNames
        cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMsgState to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientScheme to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientParaSet to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientparaList to " & strRoleNames
        cnOracle.Execute "grant Select on sys.dba_role_privs to " & strRoleNames
    End If
    GrantSpecialToRole = True
    Exit Function
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function CheckCBOPars() As Boolean
'���ܣ����ɱ�����������ṩ�޸Ĺ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String
    Dim cnTmp As ADODB.Connection
    
    On Error GoTo errh
    strSQL = "Select Name,Value,Decode(Name,'optimizer_index_cost_adj','20','80') suggestivevalue" & vbNewLine & _
                    "From V$parameter" & vbNewLine & _
                    "Where Name = 'optimizer_index_cost_adj' And Value = '100' Or Name = 'optimizer_index_caching' And Value = '0'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ɱ�����������")
    
    If rsTmp.RecordCount <> 0 Then
        If rsTmp.RecordCount = 1 Then
            strMsg = "���ݿ����""" & rsTmp!name & """��ʼֵΪ" & rsTmp!value & "�����ܻ�" & vbNewLine & _
                            "���²�Ʒ�������⡣ �����޸�Ϊ" & rsTmp!suggestivevalue & "���Ƿ��޸ģ�"
        Else
            strMsg = "�����������ݿ�����ĳ�ʼֵ���ܻᵼ�µĲ�Ʒ�������⣺" & vbNewLine & _
                            "   ����""" & rsTmp!name & """��ʼֵΪ" & rsTmp!value & "�������޸�Ϊ" & rsTmp!suggestivevalue & "��"
            rsTmp.MoveNext
            strMsg = strMsg & vbNewLine & _
                            "   ����""" & rsTmp!name & """��ʼֵΪ" & rsTmp!value & "�� �����޸�Ϊ" & rsTmp!suggestivevalue & "��" & vbNewLine & _
                            "   �Ƿ��޸ģ�"
        End If
        If MsgBox(strMsg, vbInformation + vbYesNo, App.Title) = vbYes Then
            If Not gcnSystem Is Nothing Then
                Set cnTmp = gcnSystem
            ElseIf gblnDBA Then
                Set cnTmp = gcnOracle
            Else
                Set cnTmp = GetConnection("SYSTEM")
            End If
            '��������
            If Not cnTmp Is Nothing Then
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    strSQL = "alter system set " & rsTmp!name & "= " & rsTmp!suggestivevalue
                    cnTmp.Execute strSQL
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    CheckCBOPars = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Function

Public Sub CompileAllInvalidObject(ByRef cnThis As ADODB.Connection, ByRef strErrInfor As String, ByRef objPanel As Panel, ByRef objProgressBar As ProgressBar)
'���ܣ�����ָ�����������ߵ���Ч����
'������cnThis=����������,����������Բ�ͬ�����ߵ���
'      objPanel=������ʾ��ǰ����Ķ�������
'      objProgressBar=������ʾ�������
    Dim rsObjects As New ADODB.Recordset
    Dim rsDepends As New ADODB.Recordset
    Dim arrObjects As Variant, strCompile As String
    Dim strSQL As String, i As Long
    Dim strUser As String
    
    On Error GoTo errHandle
    strErrInfor = ""
  
    strSQL = _
        "Select User, Object_Name, Object_Type" & vbNewLine & _
        "From User_Objects" & vbNewLine & _
        "Where Object_Type In" & vbNewLine & _
        "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
        "      Object_Name Not Like 'BIN$%' And Status = 'INVALID'" & vbNewLine & _
        "Order By Object_Type, Object_Name"

    rsObjects.CursorLocation = adUseClient
    rsObjects.Open strSQL, cnThis, adOpenKeyset '���Կ��������û����ĵ�����
    
    objProgressBar.Max = 100
    objProgressBar.value = 0
      
    If Not rsObjects.EOF Then
        strUser = rsObjects!User
        strSQL = _
            "Select Name, Type, Referenced_Name, Referenced_Type" & vbNewLine & _
            "From User_Dependencies" & vbNewLine & _
            "Where Referenced_Owner = User And Type In ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE'," & vbNewLine & _
            "       'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Referenced_Type In" & vbNewLine & _
            "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Not(Name=Referenced_Name And Type=Referenced_Type) And" & vbNewLine & _
            "      Name Not Like 'BIN$%' And Referenced_Name Not Like 'BIN$%'"
        rsDepends.CursorLocation = adUseClient
        rsDepends.Open strSQL, cnThis, adOpenKeyset '���Կ��������û����ĵ�����
        
        ReDim arrObjects(rsObjects.RecordCount - 1) As String
        For i = 1 To rsObjects.RecordCount
            arrObjects(i - 1) = rsObjects!Object_Name & "," & rsObjects!Object_Type
            rsObjects.MoveNext
        Next
        
        '������Ч����
        DoEvents
        For i = 0 To UBound(arrObjects)
            objPanel.Text = Split(arrObjects(i), ",")(0)    '��ʾ��ǰ��������
            objProgressBar.value = (i + 1) / (UBound(arrObjects) + 1) * 100
            DoEvents    'Ϊ��ˢ�½���
            Call CompileInvalidObject(cnThis, Split(arrObjects(i), ",")(0), Split(arrObjects(i), ",")(1), rsObjects, rsDepends, strCompile, strErrInfor)
        Next
    End If
    If strErrInfor <> "" Then strErrInfor = "������Ч����������:" & vbCrLf & strErrInfor
    
    Exit Sub
    
errHandle: '�����ڲ�������δ֪�쳣
    If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub

Private Sub CompileInvalidObject(ByRef cnThis As ADODB.Connection, ByVal strName As String, ByVal strType As String, _
    ByRef rsObjects As ADODB.Recordset, ByRef rsDepends As ADODB.Recordset, ByRef strCompile As String, ByRef strErrInfor As String)
'���ܣ�����ָ������Ч����
'������strCompile=�Ѿ�����Ķ�������
'˵����CompileAllnvalidObject���Ӻ���
    Dim arrObjRef As Variant, strErr As String
    Dim strSQL As String, i As Long
        
    If InStr(strCompile & ",", "," & strName & ",") > 0 Then Exit Sub
    
    '�ݹ���뵱ǰ���������õĶ���
    rsDepends.Filter = "Name='" & strName & "' And Type='" & strType & "'" '�������Ϳ�������ݹ����(ͬ��BODY)
    If Not rsDepends.EOF Then
        ReDim arrObjRef(rsDepends.RecordCount - 1) As String
        For i = 1 To rsDepends.RecordCount
            arrObjRef(i - 1) = rsDepends!Referenced_Name & "," & rsDepends!Referenced_Type
            rsDepends.MoveNext
        Next
        For i = 0 To UBound(arrObjRef)
            rsObjects.Filter = "Object_Name='" & Split(arrObjRef(i), ",")(0) & "' And Object_Type='" & Split(arrObjRef(i), ",")(1) & "'"
            If Not rsObjects.EOF Then '���ö���Ҳ����Ч����ʱ
                Call CompileInvalidObject(cnThis, Split(arrObjRef(i), ",")(0), Split(arrObjRef(i), ",")(1), rsObjects, rsDepends, strCompile, strErrInfor)
            End If
        Next
    End If
    '���뵱ǰ����
    Select Case strType
    Case "PROCEDURE"
        strSQL = "ALTER PROCEDURE " & strName & " COMPILE"
    Case "FUNCTION"
        strSQL = "ALTER FUNCTION " & strName & " COMPILE"
    Case "VIEW"
        strSQL = "ALTER VIEW " & strName & " COMPILE"
    Case "MATERIALIZED VIEW"
        strSQL = "ALTER MATERIALIZED VIEW " & strName & " COMPILE"
    Case "TRIGGER"
        strSQL = "ALTER TRIGGER " & strName & " COMPILE"
    Case "PACKAGE"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE"
    Case "PACKAGE BODY"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE BODY"
    Case "TYPE"
        strSQL = "ALTER TYPE " & strName & " COMPILE"
    Case "TYPE BODY"
        strSQL = "ALTER TYPE " & strName & " COMPILE BODY"
    End Select
    If strSQL <> "" Then
        strErr = ""
        err.Clear: On Error Resume Next
        cnThis.Execute strSQL
        If cnThis.Errors.Count > 0 Then
            '�������(δ����):Err.Number=0,NativeError=0
            '[Microsoft][ODBC driver for Oracle]�����Ĺ��̻���������б������
            'û�и���Ľ����
            If Not (cnThis.Errors(0).NativeError = 0 And cnThis.Errors.Count = 1) Then
                If cnThis.Errors(0).NativeError <> 0 Then
                    strErr = cnThis.Errors(0).Description
                    strErrInfor = strErrInfor & vbCrLf & strName & ":" & strErr
                Else
                    strErrInfor = strErrInfor & vbCrLf & strName
                End If
            End If
        End If
        err.Clear: On Error GoTo 0
        strCompile = strCompile & "," & strName
    End If
End Sub

Public Function GetDetailParas(ByVal lngParID As Long, Optional ByRef rsSysTems As ADODB.Recordset, Optional ByRef int���� As Integer, Optional ByRef int���� As Integer, Optional ByRef int˽�� As Integer, Optional ByRef strOwner As String) As Recordset
'���ܣ���ȡ��������ϸ����
'������lngParID=����ID
'          rsSystems=ϵͳ�б���Ҫ�����ֶΣ�ϵͳ�������ߣ�
'���أ�����id, վ��, ����id,����, ���ż���,�û���, ��Աid,��Ա,��Ա����,������,����������,����ֵ
    Dim rsParInfo  As ADODB.Recordset, rsDetailParas As ADODB.Recordset
    Dim strSQL As String
    Dim lngSys As Long
    Dim StrDefaultSQL As String
    
    On Error GoTo errh
    strSQL = "Select Nvl(ϵͳ, 0) ϵͳ, Nvl(˽��, 0) ˽��, Nvl(����, 0) ����, Nvl(����, 0) ���� From zlParameters A Where ID =[1]"
    Set rsParInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDetailParas", lngParID)
    If Not rsParInfo.EOF Then
        int���� = Val(rsParInfo!����)
        int���� = Val(rsParInfo!����)
        int˽�� = Val(rsParInfo!˽��)
        lngSys = Val(rsParInfo!ϵͳ)
    End If
    If Not (int���� = 0 And int���� = 0 And int˽�� = 0) And rsSysTems Is Nothing Then
    '����ģ��͹���ȫ��,��������ϸ�������ݣ���˲��û�ȡϵͳ
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    End If
    If lngSys <> 0 Then '�ǹ����߲���
        rsSysTems.Filter = "���=" & lngSys
        If Not rsSysTems.EOF Then strOwner = rsSysTems!������
    Else '�����߲���Ϊ˽��ȫ�ֻ򹫹�ȫ��
        rsSysTems.Filter = "���=100"
        If Not rsSysTems.EOF Then
            strOwner = rsSysTems!������
        Else
             rsSysTems.Filter = ""
             rsSysTems.Sort = "���"
             If Not rsSysTems.EOF Then strOwner = rsSysTems!������
        End If
    End If
    If int���� = 1 Then '���Ų���
        strSQL = "Select a.����id, b.վ��, a.����id, b.���� ����, zlSpellCode(b.����) ���ż���, Null �û���, Null ��Աid, Null ��Ա, Null ��Ա����, Null ������," & vbNewLine & _
                    "       Null ����������, a.����ֵ" & vbNewLine & _
                    "From Zldeptparas A, " & strOwner & ".���ű� B" & vbNewLine & _
                    "Where a.����id =[1] And a.����id = b.Id"
        StrDefaultSQL = "Select a.����id, Null վ��, a.����id, Null ����, Null ���ż���, Null �û���, Null ��Աid, Null ��Ա, Null ��Ա����, Null ������, Null ����������, a.����ֵ" & vbNewLine & _
                                    "From Zldeptparas A" & vbNewLine & _
                                    "Where a.����id = [1]"
    Else
        If int���� = 1 Then '�������Ͳ���
            StrDefaultSQL = "Select a.����id, Null վ��, Null ����id, Null ����, Null ���ż���, a.�û���, Null ��Աid, Null ��Ա, Null ��Ա����, a.������, Null ����������, a.����ֵ" & vbNewLine & _
                                    "From zlUserParas A" & vbNewLine & _
                                    "Where a.����id = [1]"
            If int˽�� = 1 Then '����˽��ģ��
                strSQL = "Select a.����id, d.վ��, e.Id ����id, e.���� ����, zlSpellCode(e.����) ���ż���, a.�û���, b.��Աid, c.���� ��Ա, zlSpellCode(c.����) ��Ա����, a.������," & vbNewLine & _
                                "       zlSpellCode(a.������) ����������, a.����ֵ" & vbNewLine & _
                                "From zlUserParas A, " & strOwner & ".�ϻ���Ա�� B, " & strOwner & ".��Ա�� C, zlClients D, " & strOwner & ".���ű� E" & vbNewLine & _
                                "Where a.����id =[1] And a.�û��� = b.�û���(+) And a.������ = d.����վ(+) And b.��Աid = c.Id(+) And d.���� = e.����(+) And" & vbNewLine & _
                                "      a.�û��� Is Not Null And a.������ Is Not Null"
            Else '��������ģ��
                strSQL = "Select a.����id, d.վ��, e.Id ����id, e.���� ����, zlSpellCode(e.����) ���ż���, a.�û���, Null ��Աid, Null ��Ա, Null ��Ա����, a.������," & vbNewLine & _
                            "       zlSpellCode(a.������) ����������, a.����ֵ" & vbNewLine & _
                            "From zlUserParas A, zlClients D, " & strOwner & ".���ű� E" & vbNewLine & _
                            "Where a.����id =[1] And a.������ = d.����վ(+) And d.���� = e.����(+) And a.�û��� Is Null And a.������ Is Not Null"
            End If
        Else
            If int˽�� = 1 Then '˽��ģ���˽��ȫ��
                strSQL = "Select a.����id, e.վ��, e.Id ����id, e.���� ����, zlSpellCode(e.����) ���ż���, a.�û���, b.��Աid, c.���� ��Ա, zlSpellCode(c.����) ��Ա����, a.������," & vbNewLine & _
                            "       zlSpellCode(a.������) ����������, a.����ֵ" & vbNewLine & _
                            "From zlUserParas A, " & strOwner & ".�ϻ���Ա�� B, " & strOwner & ".��Ա�� C, " & strOwner & ".���ű� E, " & strOwner & ".������Ա F" & vbNewLine & _
                            "Where a.����id =[1] And a.�û��� = b.�û���(+) And b.��Աid = c.Id(+) And c.Id = f.��Աid(+) And f.ȱʡ = 1 And f.����id = e.Id(+) And" & vbNewLine & _
                            "      a.�û��� Is Not Null And a.������ Is Null"
            Else '����ģ��͹���ȫ��,��������ص�����
                strSQL = "Select a.����id, Null վ��, Null ����id, Null ����, Null ���ż���, a.�û���, Null ��Աid, Null ��Ա, Null ��Ա����, a.������," & vbNewLine & _
                            "       zlSpellCode(a.������) As ����������, a.����ֵ" & vbNewLine & _
                            "From zlUserParas A" & vbNewLine & _
                            "Where ����id = [1] And 1 = 2"
            End If
        End If
    End If
    If strOwner = "" Then 'û�а�װ�κ�Ӧ��ϵͳ
        Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, StrDefaultSQL, "GetDetailParas", lngParID)
    Else
        On Error Resume Next '���ܸ�ϵͳû�в�����Ա�Ȼ�����
        Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDetailParas", lngParID)
        If err.Number <> 0 Then
            err.Clear
            On Error GoTo errh
            strOwner = ""
            Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, StrDefaultSQL, "GetDetailParas", lngParID)
        Else
            On Error GoTo errh
        End If
    End If
    Set GetDetailParas = rsDetailParas
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function


Public Function KillSessions(Optional ByVal strSesionInfo As String, Optional ByRef cllRacConn As Collection) As Boolean
'���ܣ���ȡɱ���Ự��SQL
'blnKill=�Ƿ�ֱ��ִ��
'strSesionInfo=ɱ���ƶ��ػ�ID,һ��Ϊ"Sid,Serial"
'cllRacConn=����RAC����������Ҫ�����ӣ�Ϊ�գ�����Ҫ����Rac���⴦��
    Dim rsTmp As ADODB.Recordset, strSQL As String, strAdjustSQL As String, strKillProcess As String
    Dim strTmp As String, bln10g As Boolean, strPre As String
    Dim rsIns As ADODB.Recordset, cnnTmp As ADODB.Connection
    Dim strUser As String, rsSysTems As ADODB.Recordset
    
    On Error GoTo errh
    '���ܴ�ʱgstrUserName��δ��ȡ
    If gstrUserName <> "" Then
        strUser = gstrUserName
    Else
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
        If Not rsSysTems.EOF Then
            rsSysTems.Sort = "���"
            strUser = rsSysTems!������
        End If
    End If
    
    If strSesionInfo <> "" Then
        gcnOracle.Execute "alter system kill session '" & strSesionInfo & "' immediate"
        KillSessions = True
        Exit Function
    End If
    '��ȡ���ݿ�汾
    bln10g = GetOracleVersion(True, True) < 11
    '���ֱ��ɱ��֮ǰ����Ҫ���ÿͻ���
    strSQL = "Select ��Ŀ, ���� From Zlupgradeconfig Where ��Ŀ =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, "�ͻ���״̬")
    If rsTmp.EOF Then
        strAdjustSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('�ͻ���״̬',1)"
    ElseIf Val(rsTmp!���� & "") = 0 Then
        strAdjustSQL = "Update ZLTOOLS.zlUpgradeConfig Set  ����='1'  Where ��Ŀ ='�ͻ���״̬'"
    End If
    If strAdjustSQL <> "" Then '����Ѿ�ʹ�ý��ÿͻ��˲�ɱ���Ự���ܣ���������Ǩ�����е����ť����
        gcnOracle.Execute strAdjustSQL
    End If
    On Error Resume Next
    '�������ٳ�����PL/SQL��ִ�еò�����Ҫ������
    strSQL = "Select Distinct A.Username, A.Program, A.Audsid, B.Ip, B.����վ" & vbNewLine & _
                    "From v$session a, Zlclients b" & vbNewLine & _
                    "Where A.Terminal = B.����վ And Upper(A.Program) =Upper( [1] ) And A.Audsid = Userenv('SessionID') And" & vbNewLine & _
                    "      B.Ip = Sys_Context('USERENV', 'IP_ADDRESS')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, IIf(gblnInIDE, "vb6.exe", App.EXEName & ".exe"))
    If Not rsTmp.EOF Then strTmp = rsTmp!����վ & ""
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    '���ÿͻ���
    strAdjustSQL = "Update Zlclients Set ��ֹʹ�� = 1, ϵͳ�������� = 1 Where Nvl(��ֹʹ��, 0) = 0 " & IIf(strTmp <> "", "  And ����վ <> '" & strTmp & "'", "")
    gcnOracle.Execute strAdjustSQL
    '�ж��Ƿ����ZLkillProcess��
    If CheckAndAdjustMustTable("zlkillprocess", , False) Then
        strKillProcess = "zlkillprocess"
        On Error Resume Next
        If err.Number <> 0 Then err.Clear
        strSQL = "Select Count(1) ���� From Zltools.Zlkillprocess Where Rownum < 2"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "zlkillprocess�����ж�")
        If rsTmp!���� = 0 Then
            strKillProcess = ""
        End If
        '����û�в�ѯȨ��
        If err.Number <> 0 Then
            err.Clear
            strKillProcess = ""
        End If
        On Error GoTo errh
    End If
    If strKillProcess <> "" Then
        strKillProcess = "Select Upper(����) From Zltools.Zlkillprocess Union All" & vbNewLine & _
                        "Select 'VB6.EXE' From Zltools.Zlkillprocess"
    Else
        strKillProcess = "'ZL9LABPRINTSVR.EXE','ZL9LABRECEIV.EXE','ZL9LABTCPSVR.EXE','ZL9LISCOMM.EXE'," & _
                        "'ZL9WIZARDMAIN.EXE','ZLACTMAIN.EXE','ZLHIS+.EXE','ZLHISCRUST.EXE','ZLLISRECEIVESEND.EXE'," & _
                        "'ZLNEWQUERY.EXE','ZLORCLCONFIG.EXE','ZLPACSBROWSERSTATION.EXE','ZLPACSSRV.EXE'," & _
                        "'ZLPEISAUTOANALYSE.EXE','ZLRPTSQLADJUST.EXE','ZLRUNAS.EXE','ZLSVRNOTICE.EXE'," & _
                        "'ZLSVRSTUDIO.EXE','ZLWIZARDSTART.EXE','VB6.EXE'"
    End If
    strTmp = ""
    
    '11gR2����ֱ��ɱ�ỰALTER system KILL SESSION '73,15625,@1'
    '10g��Ҫ��¼����Ӧ��Racʵ��
    strSQL = "Select 'alter system kill session ' || Chr(39) || a.Sid || ',' || a.Serial# || " & IIf(bln10g, "", "',@' || a.INST_ID || ") & " Chr(39) || ' immediate' SQL," & vbNewLine & _
            "       a.Program, b.Ip," & IIf(bln10g, " a.INST_ID,  Decode(INST_ID, userenv('instance'), 1, 0) ��ǰ��־", "userenv('instance') INST_ID,1 ��ǰ��־") & vbNewLine & _
            "  From Gv$session a, Zlclients b" & vbNewLine & _
                "Where a.Terminal = b.����վ And Upper(a.Program) In" & vbNewLine & _
                "(" & strKillProcess & ") And" & vbNewLine & _
                "      (a.Terminal <> userenv('terminal') Or" & vbNewLine & _
                "      a.Terminal= userenv('terminal')  And Upper(a.Program) Not In ('VB6.EXE', 'ZLSVRSTUDIO.EXE'))" & vbNewLine & _
                "Order By a.Terminal,a.Program"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�Ự")
    
    Set cllRacConn = New Collection
    If bln10g Then
        rsTmp.Filter = "��ǰ��־=0"
        If Not rsTmp.EOF Then
            strSQL = "select a.inst_ID, a.Instance_Name, a.Host_name, b.NAME, b.DBID" & vbNewLine & _
                    "  from gv$instance a, gv$database b" & vbNewLine & _
                    " where a.INST_ID = b.INST_ID" & vbNewLine & _
                    "   and a.INST_ID <> userenv('instance')" & vbNewLine & _
                    "   and a.STATUS = 'OPEN'"
            Set rsIns = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡʵ����Ϣ")
            '������rsTmp����Ҫ��Ϊ�˽�Լʱ��
            Do While Not rsIns.EOF
                strTmp = rsIns!INST_ID & "," & rsIns!DBID & "," & rsIns!Instance_Name & "(" & rsIns!name & ")"
                If frmUserCheckLogin.ShowLogin(UCT_RACInsUser, cnnTmp, strUser, "", "", strTmp) Then
                    cllRacConn.Add cnnTmp, "K_" & rsIns!INST_ID
                End If
                rsIns.MoveNext
            Loop
        End If
        rsTmp.Filter = ""
    End If
    rsTmp.Sort = "��ǰ��־ desc,INST_ID"
    On Error Resume Next
    strTmp = "": strPre = ""
    Do While Not rsTmp.EOF
        If rsTmp!��ǰ��־ = 0 Then
            If strPre <> rsTmp!INST_ID Then
                strPre = rsTmp!INST_ID & ""
                Set cnnTmp = cllRacConn("K_" & rsIns!INST_ID)
            End If
            cnnTmp.Execute rsTmp!SQL
        Else
            gcnOracle.Execute rsTmp!SQL
        End If
        If err.Number <> 0 Then
            strTmp = strTmp & vbNewLine & rsTmp!Program & "(" & rsTmp!IP & ")"
            err.Clear
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then
        MsgBox "���³���ĻỰ�޷���ֹ��" & strTmp & "�����ֹ�����", vbInformation, gstrSysName
    End If
    KillSessions = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function LockAppUser() As Boolean
'���ܣ�����Ӧ��ϵͳ�û�
'blnKill=�Ƿ�ֱ��ִ��
'strSesionInfo=ɱ���ƶ��ػ�ID,һ��Ϊ"Sid,Serial"
'cllRacConn=����RAC����������Ҫ�����ӣ�Ϊ�գ�����Ҫ����Rac���⴦��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strUser As String, rsSysTems As ADODB.Recordset
    
    On Error Resume Next
    If Not CheckAndAdjustMustTable("Zlclients", "ϵͳ��������", True) Then
        MsgBox "û�б�Ҫ�Ľṹ֧�֣��޷������û������ֹ�����", vbInformation, gstrSysName
        Exit Function
    End If
    '���ܴ�ʱgstrUserName��δ��ȡ
    If gstrUserName <> "" Then
        strUser = gstrUserName
    Else
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
        If Not rsSysTems.EOF Then
            rsSysTems.Sort = "���"
            strUser = rsSysTems!������
        End If
    End If
    '����������û�
    strSQL = "Update " & strUser & ".�ϻ���Ա�� b" & vbNewLine & _
                        "Set ϵͳ�������� = 1" & vbNewLine & _
                        "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.�û���)) And Upper(B.�û���)<>'" & UCase(strUser) & "'"
    gcnOracle.Execute strSQL
    '����Ѿ����ù��ܣ���������Ǩ�����е����ť����
    strSQL = "Select ��Ŀ, ���� From Zlupgradeconfig Where ��Ŀ =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, "�û�״̬")
    strSQL = ""
    If rsTmp.EOF Then
        strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(��Ŀ,����) values('�û�״̬',1)"
    ElseIf Val(rsTmp!���� & "") = 0 Then
        strSQL = "Update ZLTOOLS.zlUpgradeConfig Set  ����='1'  Where ��Ŀ ='�û�״̬'"
    End If
    If strSQL <> "" Then
        gcnOracle.Execute strSQL
    End If
    '�����û��˻�
    strSQL = "Select 'alter user ' || �û��� || ' account lock '  SQL" & vbNewLine & _
                "From " & strUser & ".�ϻ���Ա�� b" & vbNewLine & _
                "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.�û���)) And Upper(B.�û���)<>[1] " & vbNewLine & _
                "Order By �û���"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡӦ��ϵͳ�û�", UCase(strUser))
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            gcnOracle.Execute rsTmp!SQL
            If err.Number <> 0 Then err.Clear
            rsTmp.MoveNext
        Loop
    End If
    LockAppUser = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function CreateTable(ByVal cnCurrentSelect As ADODB.Connection, ByVal strOwner As String, ByVal strTableSpaces As String, ByVal strBakName As String, ByVal strTable As String, ByVal strTbsNameLob As String, Optional ByVal cnDBACreate As ADODB.Connection) As String
    '--------------------------------------------------------------------------------------------------------------------
    '����:����table.sql
    '����:strTableSpaces-��ռ�,strTbsNameLob-������ռ�
    '     strBakName-�ռ�������
    '     strOwner-���������
    '     strTableName-����
    '     cnDBACreate=���б�ṹ����������
    '     cnCurrentSelect=���б�ṹ��ѯ������
    '���أ��ɹ�����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------
    Dim rsTable As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim strTemp As String, strLobs As String
    Dim strSQL As String, blnHaveLob As Boolean
    
    On Error GoTo ErrHand
    strSQL = "Select a.Table_Name,a.Column_Name, " & _
             "               a.Data_Type, " & _
             "               a.Data_Length, " & _
             "               a.Data_Precision, " & _
             "               a.Data_Scale, " & _
             "               a.Nullable, " & _
             "               a.Data_Default " & _
             "   From Sys.all_Tab_Columns a  " & _
             "   Where a.Owner = [1] and table_Name=[2]" & _
             "   Order By a.Table_Name,a.Column_Id"
    Set rsColumn = gclsBase.OpenSQLRecord(cnCurrentSelect, strSQL, "CreateTable", strOwner, strTable)
    
    strSQL = "select table_name, tablespace_name, pct_free, TEMPORARY,DURATION ���� " & _
             " from sys.all_tables where owner = [1] and table_name=[2]"
    Set rsTable = gclsBase.OpenSQLRecord(cnCurrentSelect, strSQL, "CreateTable", strOwner, strTable)
    
    strSQL = ""
    With rsTable
        Do While Not .EOF
            rsColumn.Filter = "Table_Name='" & !Table_Name & "'"
            If rsColumn.EOF Then
                If Not cnDBACreate Is Nothing Then
                    MsgBox "��:" & !Table_Name & "������������!", vbInformation + vbDefaultButton1
                End If
                Exit Function
            End If
            If Nvl(!Temporary) = "Y" Then
                    strSQL = "CREATE GLOBAL TEMPORARY TABLE " & strBakName & "." & !Table_Name & "("
            Else
                   strSQL = "CREATE TABLE " & strBakName & "." & !Table_Name & "("
            End If
            
            strLobs = ""
            Do While Not rsColumn.EOF
                Select Case rsColumn!DATA_TYPE
                Case "NUMBER"
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE & _
                            IIf(Nvl(rsColumn!Data_Precision) = "", "", "(" & Nvl(rsColumn!Data_Precision) & IIf(Val(Nvl(rsColumn!Data_Scale)) = 0, "", "," & Val(Nvl(rsColumn!Data_Scale))) & ")")
                Case "DATE", "FLOAT", "TIMESTAMP(6)" 'TIMESTAMP(6)����ȷ��С����
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    
                Case "BLOB", "CLOB", "BFILE", "XMLTYPE"
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    blnHaveLob = True
                    
                    If rsColumn!DATA_TYPE = "BLOB" Or rsColumn!DATA_TYPE = "CLOB" Then
                        strLobs = IIf(strLobs = "", "", strLobs & ",") & rsColumn!Column_Name
                    End If
                Case Else
                    If Val(Nvl(rsColumn!Data_Length)) = 0 Then
                        strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    Else
                        strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE & "(" & Nvl(rsColumn!Data_Length) & ")"
                    End If
                End Select
                If rsColumn.AbsolutePosition = rsColumn.RecordCount Then
                    strSQL = strSQL & " " & strTemp & IIf(Nvl(rsColumn!DATA_DEFAULT) = "", "", " DEFAULT " & Replace(Nvl(rsColumn!DATA_DEFAULT), Chr(10), "")) & ")"
                Else
                    strSQL = strSQL & " " & strTemp & IIf(Nvl(rsColumn!DATA_DEFAULT) = "", "", " DEFAULT " & Replace(Nvl(rsColumn!DATA_DEFAULT), Chr(10), "")) & ","
                End If
                
                rsColumn.MoveNext
            Loop
            
            If Nvl(!Temporary) = "Y" Then
                If InStr(1, Nvl(!����), "TRANSACTION") > 0 Then
                    strSQL = strSQL & " ON COMMIT DELETE ROWS;"
                Else
                    strSQL = strSQL & " ON COMMIT PRESERVE ROWS;"
                End If
            Else
                If strLobs <> "" Then
                    If GetOracleVersion(True, True) > 10 Then
                        strLobs = " Lob(" & strLobs & ") Store as Securefile(NOCache LOGGING)"
                    Else
                        strLobs = " Lob(" & strLobs & ") Store as (Cache)"  '���Ա���Basefile��ʽ��Cache LOGGING��д�����,10G��֧��Cache LOGGING�ؼ��֣�ȱʡ��LOGGING
                    End If
                End If
                strSQL = strSQL & strLobs & " TABLESPACE " & IIf(blnHaveLob, strTbsNameLob, strTableSpaces)
                If Nvl(!pct_free) <> "" Then
                    '������ֻ�����ݣ�Ϊ��ߴ洢Ч�ʣ��̶�pctfreeΪ5
                    strSQL = strSQL & " PCTFREE 5"
                End If
            End If
            .MoveNext
        Loop
    End With
    If Not cnDBACreate Is Nothing Then
        cnDBACreate.Execute strSQL
    End If
    CreateTable = strSQL
    Exit Function
ErrHand:
    If Not cnDBACreate Is Nothing Then
        If MsgBox("�ڻ�ȡ��صı�ṹʱ����������ϸ��������:" & vbCrLf & err.Description & _
            vbCrLf & strSQL & vbCrLf & "�Ƿ������˱�Ĵ���?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            CreateTable = IIf(strSQL = "", "Skip", strSQL)
        End If
    End If
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub CreateTempTabForBakTable(ByVal strTable As String)
'���ܣ�Ϊ��ʷת��������ʱ��ͬ���
'������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTableH As String
    
    On Error GoTo ErrHand
    strTableH = "H" & strTable
    
    '����Ѵ�����ʱ���򲻱��ؽ�
    strSQL = "Select 1 From User_Tables Where Table_Name = [1] And Temporary = 'Y'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�����ʷ����ͼ", strTableH)
    If rsTmp.RecordCount = 0 Then

        '��ǰ�����Ĺ���ͬ�����Ȼָ�����H��ͼ�����ڸ�ΪH��󣬲�Ӱ��ʹ�ã���ֻ��ͨ�����ƹ����ģ�û�йܶ������ͣ����Բ���ɾ�����ؽ�
        '1.ɾ����ǰָ����ʷ��ռ����ͼ
        '�����10.35.70֮ǰ�İ汾��������ͼ
        strSQL = "Select 1 From User_Views Where View_Name = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�����ʷ����ͼ", strTableH)
        If rsTmp.RecordCount > 0 Then
            strSQL = "Drop View " & strTableH
            gcnOracle.Execute strSQL
        End If
        
        '2.������ʱ��
        strSQL = "Create Global Temporary Table " & strTableH & " On Commit Delete Rows as select * from " & strTable & " where 1=0"
        gcnOracle.Execute strSQL
    End If
    
    Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub DropTempTabForBakTable(ByVal strTable As String)
'���ܣ���鲢ɾ����ʱ��ʷ��
'������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTableH As String
    
    On Error GoTo ErrHand
    strTableH = "H" & strTable
    
    strSQL = "Select 1 From User_Tables Where Table_Name = [1] And Temporary = 'Y'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�����ʷ����ͼ", strTableH)
    If rsTmp.RecordCount > 0 Then
        '��Dblink��ʷ���л�Ϊ������ʷ��ʱ����ɾ��ͬ������ʱ�����ܴ�����ͼ
        strSQL = "Drop Table " & strTableH
        gcnOracle.Execute strSQL
    End If
    
    Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Public Function CreateAppView(ByVal strOwner As String, ByVal strBakOwner As String, _
    ByVal lngϵͳ As Long, ByVal strDbLink As String, _
    Optional ByRef pgbState As ProgressBar, Optional ByRef clsScript As clsRunScript) As Boolean
    '----------------------------------------------------------------------------------------------------------------------
    '����:����H�����ͼ
    '����:gcnOracle-��ǰӦ��ϵͳ����������
    '     lngϵͳ-ϵͳ���
    '     strOwner-Ӧ��ϵͳ��������
    '     strBakOwner-��ʷ���ݿռ��������
    '     strDbLink-���ݿ��������ƣ���һ���ַ���@
    '     pgbState-�������ؼ�
    '     clsScript-������־�����
    '����:�����ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------
    Dim rsTables As ADODB.Recordset
    Dim strBakTableName As String
    Dim i As Long
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select ���� from zlBakTables where ϵͳ=[1]"
    Set rsTables = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ȡת����", lngϵͳ)
    
    If Not pgbState Is Nothing Then
        pgbState.Max = 100
        pgbState.value = 0
        DoEvents
    End If
    
    On Error Resume Next
    With rsTables
        Do While Not .EOF
            If strDbLink = "" Then '���֮ǰ�Ƿ������Զ����ʷ�����������ʱH��
                Call DropTempTabForBakTable(!����)
            End If
            
            strBakTableName = strBakOwner & "." & !���� & strDbLink
            
            gstrSQL = "Create or replace view  " & strOwner & ".H" & !���� & " as Select * From " & strBakTableName
            gcnOracle.Execute gstrSQL
            
            '����LOB�ֶεı��޷�����ָ��Զ�̷���������ͼ������ORA-22992: �޷�ʹ�ô�Զ�̱�ѡ��� LOB ��λ��
            '��Ϊ������ʱ����ȡʱ����ʱ���������,��Ϊͨ��dblink����lob��֧��insert into ...select ��ʽ
            If strDbLink <> "" Then
                If InStr(err.Description, "ORA-22992") > 0 Then
                    err.Clear
                    Call CreateTempTabForBakTable(!����)
                End If
            End If
            
            If err.Number <> 0 Then
                If Not clsScript Is Nothing Then
                    clsScript.ErrCount = clsScript.ErrCount + 1
                    clsScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(strOwner & "." & "H" & !���� & "����ʧ��", 30) & "������" & err.Description
                Else
                    If MsgBox("������ͼ""H" & !���� & """����," & vbCrLf & " ������Ϣ:(" & err.Number & ") " & err.Description & vbCrLf & "�Ƿ���Ըô�����ִ�У�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
                err.Clear
            End If
            If Not pgbState Is Nothing Then
                i = i + 1
                pgbState.value = i / rsTables.RecordCount * 100
                DoEvents
            End If
            .MoveNext
        Loop
    End With
    
    '��������zlbakInfo��ͼ
    If rsTables.RecordCount <> 0 Then
        strBakTableName = strBakOwner & ".ZLBAKINFO" & strDbLink
        gstrSQL = "Create or replace view  " & strOwner & "." & "ZLBAKINFO as Select * From  " & strBakTableName
        gcnOracle.Execute gstrSQL
        
        If err.Number <> 0 Then
            If Not clsScript Is Nothing Then
                clsScript.ErrCount = clsScript.ErrCount + 1
                clsScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(strOwner & "." & "ZLBAKINFO" & "����ʧ��", 30) & "������" & err.Description & _
                        IIf(strDbLink = "", "", vbCrLf & "Զ�����ӿ��ܲ�����(" & strDbLink & ")")
            Else
                MsgBox strOwner & "." & "ZLBAKINFO" & "����ʧ��." & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
    End If
    
    CreateAppView = True
    Exit Function
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function IsCanInstallPLJson(ByVal strToolsFloder As String, Optional ByRef blnInstallRemain As Boolean) As Boolean
'���ܣ��ж��Ƿ���԰�װPLJSON��1���Ѿ���װ�����ð�װ��2��û�а�װ�����PLJSON�ű��Ƿ���ڣ���������԰�װ
'������strToolsFolder=APPSOFT\TOOLS\Ŀ¼λ��
'
'���أ�True-���԰�װPLJSON��False-���ܰ�װPLJSON
'      blnInstallRemain=�Ƿ���ڰ�װ����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    'TYPE BODY,TYPE,SYNONYM
    strSQL = "Select Count(1) ����" & vbNewLine & _
            "From All_Objects a" & vbNewLine & _
            "Where a.Object_Name In ('JSON', 'JSON_VALUE','JSON_VALUE_ARRAY', 'JSON_LIST','JSON_HELPER','JSON_PARSER','JSON_EXT','JSON_AC')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�Ƿ�װPLJSON")
    blnInstallRemain = False
    '���ܴ���˽��ͬ��ʵ����������ڵ���23��JSON_VALUE_ARRAYû��TYPE BODY��������
    If rsTmp!���� < 23 Then
        If gobjFile.FileExists(strToolsFloder & "\PLJSON1.0.6install.SQL") And gobjFile.FileExists(strToolsFloder & "\PLJSON1.0.6uninstall.SQL") Then
            IsCanInstallPLJson = True
        End If
        blnInstallRemain = rsTmp!���� <> 0
    End If
    Exit Function
errh:
    Call WriteTraceLog("IsConInstallPLJson:" & err.Description)
    err.Clear
End Function

Public Function InstallPLJSON(ByVal cnDBA As ADODB.Connection, ByVal strToolsFloder As String, ByRef objRunScript As clsRunScript, Optional ByVal blnUninstallFirst As Boolean) As Boolean
'���ܣ���װPLJSON,��װʧ�ܣ����Զ�����װ
'������strToolsFolder=APPSOFT\TOOLS\Ŀ¼λ��
'      cnDBA=��װ�����DBA����
'      objRunScript=�ű��ļ���������
'      blnUninstallFirst=�Ƿ��Ƚ��з���װ�����ܰ�װ�жϣ����²��ֶ������
'���أ�True-��װ�ɹ���False-��װʧ��
    Dim blnInstallErr As Boolean

    On Error GoTo errh
    '�����ڰ�װ���������Ƚ��з���װ
    If blnUninstallFirst Then
        Call UninstallPLJSON(cnDBA, strToolsFloder, objRunScript)
    End If
    If Not objRunScript.OpenFile(strToolsFloder & "\PLJSON1.0.6install.SQL") Then
        objRunScript.WriteLog String(9, " ") & "�����PLJSON��װʧ��"
        Exit Function
    End If
    blnInstallErr = True
    Do While Not objRunScript.EOF
        cnDBA.Execute objRunScript.SQLInfo.SQL
        objRunScript.ReadNextSQL
    Loop
    Exit Function
errh:
    If blnInstallErr Then objRunScript.WriteLog Format(objRunScript.SQLInfo.FileLine, "0000000") & ":" & GetLogSQL(objRunScript.SQLInfo), 2
    objRunScript.WriteLog String(17, " ") & "����" & err.Description
    err.Clear
    If blnInstallErr Then
        objRunScript.WriteLog String(17, " ") & "�����Զ���ֹ��װ�����з���װ"
        objRunScript.WriteLog String(9, " ") & "�����PLJSON��װʧ��"
        Call UninstallPLJSON(cnDBA, strToolsFloder, objRunScript)
    Else
        objRunScript.WriteLog String(17, " ") & "�����Զ���ֹ��װ"
        objRunScript.WriteLog String(9, " ") & "�����PLJSON��װʧ��"
    End If
End Function

Public Function UninstallPLJSON(ByVal cnDBA As ADODB.Connection, ByVal strToolsFloder As String, ByRef objRunScript As clsRunScript) As Boolean
'���ܣ�����װPLJSON
'������strToolsFolder=APPSOFT\TOOLS\Ŀ¼λ��
'      cnDBA=����װ�����DBA����
'      objRunScript=�ű��ļ���������
'���أ�True-����װ�ɹ���False-����װʧ��
    Dim blnUninstallErr As Boolean

    On Error GoTo errh
    If Not objRunScript.OpenFile(strToolsFloder & "\PLJSON1.0.6uninstall.SQL") Then
        objRunScript.WriteLog String(9, " ") & "�����PLJSON����װʧ��"
        Exit Function
    End If
    blnUninstallErr = True
    Do While Not objRunScript.EOF
        cnDBA.Execute objRunScript.SQLInfo.SQL
        objRunScript.ReadNextSQL
    Loop
    Exit Function
errh:
    If blnUninstallErr Then objRunScript.WriteLog Format(objRunScript.SQLInfo.FileLine, "0000000") & ":" & GetLogSQL(objRunScript.SQLInfo), 2
    objRunScript.WriteLog String(17, " ") & "����" & err.Description
    objRunScript.WriteLog String(17, " ") & "�����Զ���ֹ����װ"
    err.Clear
End Function

