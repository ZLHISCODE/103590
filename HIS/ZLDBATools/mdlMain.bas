Attribute VB_Name = "mdlMain"
Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gblnDBA As Boolean                   '�Ƿ�DBA
Public gstrSQL    As String                 'ͨ�õ�SQL������
Public gstrSysName As String                'ϵͳ����
Public gstrUserName As String               '�û���
Public gstrPassword As String               '�û�����
Public gstrServer As String                 '��������
Public gstrFilePath As String           '�ļ������Ĭ��·��

'�����ݿ����ӹ�������
Public gblnHadInit As Boolean        '�Ƿ��Ѿ���ʼ��
Public gblnIsZlhis As Boolean           '�Ƿ�ΪZLHIS����
Public gstrVerNum As String           '���ݿ�汾��
Public gstrBigVer As String              '���ݿ��汾
Public gblnRAC As Boolean               '�Ƿ�ΪRac����
Public gintCpuCount  As Integer, gintCpuAdvise As Integer, gintCpuMax As Integer   'CPU��״�Լ����鲢�ж�
Public gintInstId As Integer            'RAC������,��ǰʵ��ID
Public gblnHasBigtables As Boolean '��¼�Ƿ���Bigtables���ű�
Public gblnHasZltables As Boolean '��¼�Ƿ���zltable���ű�

'API���
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_RESTORE = &HF120&
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'������ɫ
Public Enum rowColor
    FULL_��ɫ = &HB3DEF5
    BackAlterNate_��ɫ = &HF0FFF0
    Back_��ɫ = &H80000005
    Used_��ɫ = &HEEEEE0
    OFF_��ɫ = &HB3DEF5
End Enum

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long


Public Function OpenSQLRecordByArray(ByVal strSql As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
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
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String

    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                'strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
        End If
    End If
    
    
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

    '���ԭ�в���:��Ȼ�����ظ�ִ��
'    cmdData.CommandText = "" '��Ϊ����ʱ�����������
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
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
    Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
 
    cmdData.CommandText = strSql
    
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
    
End Function


Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(strSql, strTitle, arrPars)
End Function
Public Sub InitTable(vsf As VSFlexGrid, strCol As String)
'����: ��ʼ����ͷ
    Dim arrHead As Variant
    Dim i As Long
    
    arrHead = Split(strCol, ";")
   
    With vsf
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDNone
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Public Function IsInstallExcel() As Boolean
'���ܣ��жϱ�����װ��EXCELû��
'������
'���أ����򷵻�True
    Dim objTemp  As Object
    
    On Error GoTo errH
    Set objTemp = CreateObject("Excel.Application") '��һ��EXCEL����
    Set objTemp = Nothing
    IsInstallExcel = True
    Exit Function
errH:
    Set objTemp = Nothing
    IsInstallExcel = False
    Err.Clear
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

Public Sub ErrCenter(Optional strErr As String)
    MsgBox Err.Description & vbCrLf & strErr, vbExclamation, "����"
End Sub


Public Sub CreateStr2list()
    '���ܣ�����Ƿ����F_STR2LIST2������������������ӡ�
    Dim strSql As String, rsData As ADODB.Recordset
    
    If mblnZlhis Then Exit Sub  'zlhis����ֱ���˳�
    
    On Error Resume Next
    '�ж��Ƿ���Ҫ��������
    strSql = "select  1 from all_objects where object_name ='F_STR2LIST2' and OWNER='PUBLIC'  and Object_type ='SYNONYM'"
    Set rsData = OpenSQLRecord(strSql, "CreateStr2list")
    If rsData.RecordCount > 0 Then Exit Sub
    
    '����-1.��������
    strSql = "CREATE OR REPLACE Type t_StrObj2 as object (C1 Varchar2(4000),C2 Varchar2(4000))"
    gcnOracle.Execute strSql
    strSql = "CREATE OR REPLACE Type t_StrList2 as table of t_StrObj2"
    gcnOracle.Execute strSql
    
    '����-2.��������
    strSql = "Create Or Replace Function f_Str2list2" & vbNewLine & _
                    "(" & vbNewLine & _
                    "  Str_In      In Varchar2,Split_In    In Varchar2 := ',', Subsplit_In In Varchar2 := ':'" & vbNewLine & _
                    ") Return t_Strlist2" & vbNewLine & _
                    "  Pipelined As" & vbNewLine & _
                    "  v_Str   Long; P       Number; v_Tmp   Varchar2(4000);" & vbNewLine & _
                    "   Out_Rec t_Strobj2 := t_Strobj2(Null, Null);" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  If Str_In Is Null Then" & vbNewLine & _
                    "    Return;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "  v_Str := Str_In || Split_In;" & vbNewLine & _
                    "  Loop" & vbNewLine & _
                    "    P := Instr(v_Str, Split_In);Exit When(Nvl(P, 0) = 0);v_Tmp      := Substr(v_Str, 1, P - 1);Out_Rec.C1 := Substr(v_Tmp, 1, Instr(v_Tmp, Subsplit_In) - 1);" & vbNewLine & _
                    "    Out_Rec.C2 := Substr(v_Tmp, Instr(v_Tmp, Subsplit_In) + 1); Pipe Row(Out_Rec);v_Str := Substr(v_Str, P + 1);" & vbNewLine & _
                    "  End Loop;" & vbNewLine & _
                    "  Return;" & vbNewLine & _
                    "End;"
    gcnOracle.Execute strSql
    
    '����-3.���ͬ���
    strSql = "create or replace synonym F_STR2LIST2 for f_Str2list2"
    gcnOracle.Execute strSql
    
    '����-4.ͬ�����Ȩ
    strSql = " grant execute on  F_STR2LIST2 to public"
    gcnOracle.Execute strSql
    
    If Err.Number > 0 Then
        MsgBox Err.Description
    End If
End Sub

Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '���ܣ����ݱ����жϱ��Ƿ����
    '������strTableName - Ҫ��ѯ�ı���
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "select 1 from dba_all_tables where table_name =[1] "
    Set rsData = OpenSQLRecord(strSql, "CheckTblExist", strTableName)
    CheckTblExist = (rsData.RecordCount > 0)
    
    Exit Function
errH:
    MsgBox Err.Description
End Function

Public Function GetPrevSQLID(ByRef strChildNum As String) As String
'���ܣ���ȡ��ǰ�Ự���һ��ִ�е�SQLID,��������CHILD_NUMBER��ֵ�����������
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "select prev_sql_id,PREV_CHILD_NUMBER from V$session where AUDSID=UserENV('SessionID')"
    Set rsTmp = OpenSQLRecord(strSql, "GetPrevSQLID")
 
    If rsTmp.RecordCount = 0 Then Exit Function
    GetPrevSQLID = rsTmp!prev_sql_id & ""
    strChildNum = rsTmp!PREV_CHILD_NUMBER & ""
    
    Exit Function
errH:
    MsgBox Err.Description
End Function

Public Function GetCurrentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    GetCurrentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    GetCurrentdate = 0
    Err = 0
End Function



Public Function GetTimeString(ByVal datBegin As Date, ByVal datEnd As Date) As String
'���ܣ���ȡ����ʱ��ֵ��ĸ�ʽ�ַ���
'   datBegin=��ʼʱ��
'   datEnd=��ֹʱ��
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim datTmp As Date

    intH = DateDiff("h", datBegin, datEnd)
    datTmp = DateAdd("h", intH, datBegin)
    intM = DateDiff("n", datTmp, datEnd)
    datTmp = DateAdd("n", intM, datTmp)
    intS = DateDiff("s", datTmp, datEnd)
    
    If intS < 0 Then
        intM = intM - 1
        intS = 60 + intS
    End If
    
    If intM < 0 Then
        intH = intH - 1
        intM = 60 + intM
    End If
    GetTimeString = IIf(intH <> 0, intH & "Сʱ", "") & IIf(intM <> 0, intM & "��", "") & intS & "��"
End Function

Public Function getVersion() As String
'���ܣ���ȡ���ݿ�Ĵ�汾��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = OpenSQLRecord(strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            getVersion = Mid(arrTmp(1), 1, InStr(1, arrTmp(1), ".") - 1)
        End If
    End If
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, "����"
End Function

Public Function GetOracleVersion(Optional ByVal blnGetVerNum As Boolean = False) As String
    '���ܣ���ȡ���ݿ�İ汾�ţ�Ĭ�Ϸ������ݿ��汾��
    '������blnGetVerNum-�Ƿ񷵻����ݿ������汾��
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrTmp As Variant
        
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = OpenSQLRecord(strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            strTmp = arrTmp(1)
        End If
    End If
    
    '10.2.0.3.0
    If Not blnGetVerNum Then
        arrTmp = Split(strTmp, ".")
        strTmp = Val(arrTmp(0))
    End If
    
    GetOracleVersion = strTmp
    Exit Function
errH:
    ErrCenter "��ȡ���ݿ�汾ʧ�ܣ����ֹ��ܽ��޷�ʹ�á�"
End Function

Public Function CheckRAC(ByRef intInstID As Integer) As Boolean
'���ܣ�����Ƿ�ΪRAC����
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "select 1 from gv$active_instances"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, "CheckRAC")
    
    If rsTmp.RecordCount > 0 Then
        CheckRAC = True
        
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rsTmp = OpenSQLRecord(strSql, "CheckRAC")
        intInstID = "" & rsTmp!Inst_ID
    Else
        CheckRAC = False
    End If
    
    Exit Function
errH:
    ErrCenter
End Function
Public Function GetCpuCount(ByRef intAdvise As Integer, ByRef intMax As Integer) As String
'���ܣ�����ͳ����Ϣ�ռ��Լ�����DDL�Ĳ��ж�
'����ֵ�� ������CPU������ intDefault ���鲢�жȣ�inxMax ����ж�
    Dim strSql As String, rsTmp As ADODB.Recordset
    
     '�����ΪCPU������ֹ���ߣ�ʵ��ΪCPU����*����CPU�ϲ��н���
    On Error GoTo errH
    strSql = "Select Nvl(Max(Value),0) CPU From " & IIf(gblnRAC, "G", "") & "V$parameter Where Name = 'cpu_count'" & IIf(gblnRAC, "And INST_ID = " & gintInstId & " ", "") & " "
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ����CUP��")
    
    If rsTmp!cpu <= 4 Then
        intAdvise = 1
        intMax = IIf(rsTmp!cpu = 0, 1, rsTmp!cpu)
    ElseIf rsTmp!cpu <= 8 Then
        intAdvise = 4
        intMax = rsTmp!cpu
    ElseIf rsTmp!cpu <= 12 Then
        intAdvise = 8
        intMax = rsTmp!cpu
    Else
        intAdvise = 12
        intMax = rsTmp!cpu
    End If
    
    GetCpuCount = rsTmp!cpu
    Exit Function
errH:
    ErrCenter "��ȡ������CPU����ʧ�ܣ����ֹ����޷�ʹ�á�"
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ��к���
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub ClearVsf(vsfGrid As VSFlexGrid, strNone As String)
'���ܣ���ձ�񣬲���fixedRow����һ�������ʾ��Ϣ
    With vsfGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        If strNone = "" Then
            .Rows = .FixedRows
        Else
            .Rows = .FixedRows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = strNone
        End If
        
        .MergeCells = flexMergeRestrictRows
        .MergeRow(-1) = True
        .AutoResize = True
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDDirect
        .Row = .Rows - 1
    End With
End Sub

Public Sub CreateList2str()
    '�޸Ļ��½�f_List2str����s
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    
    '����������ڣ����ô���
    strSql = "Select 1 From " & IIf(gblnIsZlhis, "DBA_ARGUMENTS", "USER_ARGUMENTS") & " Where object_name = 'F_LIST2STR' And argument_name = 'P_MAXLENGTH'"
    Set rsTmp = OpenSQLRecord(strSql, "CreateList2str")
    If rsTmp.RecordCount > 0 Then Exit Sub
    
    '�������ZLHIS��������Ҫ�������� create or replace type zltools.t_StrList as Table of Varchar2(4000)
    If Not gblnIsZlhis Then
        strSql = "Select 1 From user_types  Where type_name = 'T_STRLIST'"
        Set rsTmp = OpenSQLRecord(strSql, "CreateList2str")
        If rsTmp.RecordCount = 0 Then
            strSql = "create or replace type t_StrList as Table of Varchar2(4000)"
            gcnOracle.Execute strSql
        End If
    End If
    
    '��������
    strSql = "CREATE OR REPLACE Function " & IIf(gblnIsZlhis, "Zltools.", "") & "f_List2str" & vbNewLine & _
                    "(" & vbNewLine & _
                    "  p_Strlist   In t_Strlist,p_Delimiter In Varchar2 Default ',', p_Distinct  In Number Default 1, p_Maxlength In Number Default 0" & vbNewLine & _
                    ") Return Varchar2 Is" & vbNewLine & _
                    "l_String Long;l_Add    Number;" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  If p_Strlist.Count > 0 Then" & vbNewLine & _
                    "    For I In p_Strlist.First .. p_Strlist.Last Loop" & vbNewLine & _
                    "      l_Add := 0;" & vbNewLine & _
                    "      If p_Distinct = 1 Then" & vbNewLine & _
                    "        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then l_Add := 1; End If;" & vbNewLine & _
                    "      Else l_Add := 1; End If;" & vbNewLine & _
                    "      If l_Add = 1 Then If I != p_Strlist.First Then  l_String := l_String || p_Delimiter;End If;" & vbNewLine & _
                    "        l_String := l_String || p_Strlist(I);If p_Maxlength <> 0 And Length(l_String) > p_Maxlength Then" & vbNewLine & _
                    "        l_String := Substr(l_String, 1, p_Maxlength); Return l_String; End If;" & vbNewLine & _
                    "      End If;" & vbNewLine & _
                    "    End Loop;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "  Return l_String;" & vbNewLine & _
                    "End f_List2str;"
    gcnOracle.Execute strSql
    
    '�����ZLHIS���ʹ���ͬ��ʣ��������ZLHIS����ô�ͽ������������û��¡�
    If gblnIsZlhis Then
        '����-3.���ͬ���
        strSql = "create or replace public synonym F_LIST2STR for ZLTOOLS.f_List2str"
        gcnOracle.Execute strSql
        
        '����-4.ͬ�����Ȩ
        strSql = " grant execute on  F_LIST2STR to public"
        gcnOracle.Execute strSql
    End If
        
    Exit Sub
errH:
    ErrCenter
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    '���ܵ���OpenSQLRecord,��ΪOpenSQLRecordҲʹ���˸÷���
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    CurrentDate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    If MsgBox(Err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
    CurrentDate = 0
    Err = 0
End Function

Public Sub DrawPicture(pctDraw As PictureBox, rsData As ADODB.Recordset, intStart As Integer, intEnd As Integer, Optional blnEscape0 As Boolean)
'����:���ݴ�������ݼ����л�ͼ
'����:pctDraw - ��ͼ�Ŀؼ�  rsData-��ͼ�����ݼ�     intStart -���ݼ���ʼ��   intEnd - ���ݼ�������  blnEscape0-�Ƿ�����0
    Dim i As Integer, j As Integer
    Dim intBaseX As Integer, intBaseY As Integer    'ԭ������
    Dim intXwidth As Integer, intYheight As Integer '�����᳤��
    Dim intCountX As Integer '����������
    Dim lngMaxY As Long  '���������ֵ
    Dim intLastX As Integer, intLastY As Integer
    Dim intNowX As Integer, intNowY As Integer
    Dim ����ϵColor As Long
    
    Dim dateNow As Date
    
    dateNow = CurrentDate
    pctDraw.Cls
    If rsData.State = 0 Then Exit Sub
    If rsData.RecordCount = 0 Then Exit Sub  '���û�����ݣ����˳�
    rsData.MoveFirst
    
    '��ɫ����
    ����ϵColor = &H454545
    
    '��������ϵ
    intBaseX = 720 'ԭ��
    intBaseY = pctDraw.Height - 1000
    intXwidth = pctDraw.ScaleWidth - 240 - intBaseX 'X��Y����
    intYheight = intBaseY - 600
    
    With pctDraw
        .DrawWidth = 1
        .FontSize = 9
        
        'X��
        intCountX = intEnd - intStart + 1
        pctDraw.Line (intBaseX, intBaseY)-(intBaseX + intXwidth, intBaseY - 15), ����ϵColor, B
        .CurrentX = intBaseX: .CurrentY = intBaseY + 45
        pctDraw.Print "0"
        For i = intStart To intEnd
            .CurrentX = intBaseX + intXwidth / (intCountX + 1) * (i - intStart + 1):
            .CurrentY = intBaseY + 45
            pctDraw.Print UCase(rsData.Fields(i).Name)
            .CurrentX = intBaseX + intXwidth / (intCountX + 1) * (i - intStart + 1)
            .CurrentY = intBaseY
            pctDraw.Line (.CurrentX, .CurrentY)-(.CurrentX, .CurrentY - 60), ����ϵColor, B
        Next
        
        'Y��
        pctDraw.Line (intBaseX, intBaseY)-(intBaseX + 15, intBaseY - intYheight), ����ϵColor, B
        
        'ȡY�����ֵȷ���̶�
        rsData.MoveFirst
        lngMaxY = 1
        Do While Not rsData.EOF
            lngMaxY = IIf(lngMaxY < Val(rsData!MaxValue), Val(rsData!MaxValue), lngMaxY)
            rsData.MoveNext
        Loop
        lngMaxY = (lngMaxY + 5) / 10
        lngMaxY = lngMaxY * 10          '��������ȡ��
        
        For i = 0 To 4
            .CurrentX = intBaseX
            .CurrentY = intBaseY - intYheight / 6 * (i + 1)
            pctDraw.Line (.CurrentX, .CurrentY)-(.CurrentX + 60, .CurrentY), ����ϵColor, B
            .CurrentX = intBaseX - 600
            .CurrentY = intBaseY - intYheight / 6 * (i + 1) - 90
            pctDraw.Print (i + 1) * (lngMaxY / 5)
        Next
        
    End With
    
    rsData.MoveFirst
    With rsData
        pctDraw.FillStyle = 0
        i = 1
        Do While Not rsData.EOF
        '����ͼ��
            If i <= 5 Then
                pctDraw.FillColor = RGB(255, 255 - i * 51, 255 - i * 51)
            ElseIf i <= 10 And i > 5 Then
                pctDraw.FillColor = RGB(255 - (i - 5) * 51, 255 - (i - 5) * 51, 255 - (i - 5) * 51)
            Else
                pctDraw.FillColor = RGB(255 - (i - 10) * 51, 255, 255 - (i - 10) * 51)
            End If
            
            pctDraw.DrawWidth = 1.5
            pctDraw.FillStyle = 0
            
            pctDraw.Line (intBaseX + (i - 1) * 1600, intBaseY + 500)-(intBaseX + (i - 1) * 1600 + 500, intBaseY + 700), pctDraw.FillColor, BF
            pctDraw.CurrentX = intBaseX + (i - 1) * 1600 + 550
            pctDraw.CurrentY = intBaseY + 500
            pctDraw.Print .Fields(0).Value
            
            '��������
            pctDraw.DrawMode = 3
            pctDraw.DrawWidth = 1
            For j = intStart To intEnd
                intLastX = intNowX: intNowX = intBaseX + intXwidth / 25 * (j - 2)
                intLastY = intNowY: intNowY = intBaseY - IIf(IsNull(.Fields(j).Value), 0, .Fields(j).Value) / lngMaxY * intYheight * 5 / 6
                pctDraw.Circle (intNowX, intNowY), 25, pctDraw.FillColor
                
                If blnEscape0 Then
                    If j <> intStart Then
                        If CDate(Format(.Fields(0).Value, "yyyy/mm/dd")) < CDate(Format(dateNow, "yyyy/mm/dd")) Then
                            pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                        ElseIf CDate(Format(.Fields(0).Value, "yyyy/mm/dd")) = CDate(Format(dateNow, "yyyy/mm/dd")) _
                        And j - 3 < Val(Format(dateNow, "hh")) Then
                            pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                        End If
                    End If
                Else
                    If j <> intStart Then
                        pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                    End If
                End If
            Next
            
            i = i + 1
            rsData.MoveNext
        Loop
    End With
End Sub

Public Function ChangeSQL(ByVal intMod As Integer, ByVal strOldID As String, ByVal strSqlText As String, ByRef strChildNum As String, ByVal strInstID, Optional strOptVersion As String) As String
    '���ܣ��޸�SQL�ı��������µ�SQLID�����޸�ִ�мƻ�
    '����˵����intModΪ�޸Ĳ���, 1-���RULE��ʾ��2-ɾ��RULE��ʾ��3-����Ż����汾��ʾ �� 4-ɾ���Ż����汾��ʾ ��5-�Զ��������ʾ��
    '                   strOldID-��Ҫ�޸ĵ�SQLID; strOptVersion �Ż��������汾;strSqlText-�Զ����SQL���
    '����ֵ:����1˵��SQL������RULE��ʾ������2˵��û��RULE��ʾ������3˵���Ż����汾��ȡʧ�� ������4˵��û���Ż�������������5˵�����ִ��ʧ�� ,���򷵻��µ�SQL���
    Dim strNewSQL As String, strTemp As String
    Dim arrPar() As Variant, strSql As String, rsData As ADODB.Recordset
    Dim intHintStart As Integer, intHintEnd As Integer, strHints  As String
    
    On Error GoTo errH
    strNewSQL = strSqlText
    '������еĿո�ȥ��������ƥ������ַ�
    strTemp = Replace(UCase(strNewSQL), " ", "")
    
    '���RULE��ʾ
    If intMod = 1 Then
        If InStr(1, strTemp, "/*+RULE*/") > 0 Then ChangeSQL = "1": Exit Function
        strNewSQL = Left(strNewSQL, InStr(1, strNewSQL, "SELECT")) + Replace(strNewSQL, " ", " /*+ RULE*/ ", InStr(1, strNewSQL, "SELECT") + 1, 1)
    'ɾ��RULE��ʾ
    ElseIf intMod = 2 Then
        If InStr(1, strTemp, "/*+RULE*/") = 0 Then ChangeSQL = "2": Exit Function
        
        Do While InStr(1, strTemp, "/*+RULE*/") > 0
            intHintStart = InStr(1, strNewSQL, "/")
            intHintEnd = InStr(intHintStart + 1, strNewSQL, "/")
            strHints = Mid(strNewSQL, intHintStart, intHintEnd - intHintStart + 1)
            strNewSQL = Replace(strNewSQL, strHints, " ")
            strTemp = Replace(strNewSQL, " ", "")
        Loop
    
    '����Ż����汾��ʾ
    ElseIf intMod = 3 Then
        If strOptVersion = "" Then
            ChangeSQL = "3"
            Exit Function
        End If
        strNewSQL = Replace(strNewSQL, "", "/*+ optimizer_features_enable('" & strOptVersion & "') */", 1, 1)
    ' ɾ���Ż����汾��ʾ
    ElseIf intMod = 4 Then
        If InStr(1, strTemp, "/*+OPTIMIZER_FEATURES_ENABLE") = 0 Then ChangeSQL = "4": Exit Function
        
        Do While InStr(1, strTemp, "/*+OPTIMIZER_FEATURES_ENABLE") > 0
            intHintStart = InStr(1, strNewSQL, "/")
            intHintEnd = InStr(intHintStart + 1, strNewSQL, "/")
            strHints = Mid(strNewSQL, intHintStart, intHintEnd - intHintStart + 1)
            strNewSQL = Replace(strNewSQL, strHints, " ")
            strTemp = Replace(strNewSQL, " ", "")
        Loop
        
    '�Զ�����ʾ
    Else
        strNewSQL = TrimEx(strSqlText)
    End If
    
    '�鿴����Ƿ��а󶨱������а󶨱������޸ĺ�ִ��
    strSql = "select POSITION,NAME,VALUE_STRING ,last_captured ,DataType from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where  SQL_ID= [1]  " & _
                 "and CHILD_NUMBER in (select max(CHILD_NUMBER)  from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where SQL_ID= [1]  " & IIf(gblnRAC, "And INST_ID = " & strInstID & " ", "") & "  ) order by POSITION"
    Set rsData = OpenSQLRecord(strSql, "ChangeSQL", strOldID)
    
    '�����ѯ���Ϊ�գ����Դ�dba_hist_sqlbind��ͼ�в�ѯ
    If rsData.RecordCount = 0 Then
        strSql = "Select  Distinct Position, Name, Value_String, Last_Captured, Datatype From dba_hist_sqlbind  A Where Sql_Id = [1] order by POSITION"
        Set rsData = OpenSQLRecord(strSql, "ChangeSQL", strOldID)
    End If
    
    
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        ReDim arrPar(rsData.RecordCount - 1)
    End If
    Do While Not rsData.EOF
        
        '�滻�󶨱�����ʽΪ ��[1]��[2]
        strNewSQL = Replace(strNewSQL, rsData!Name, "[" & rsData!Position & "]", 1, 1)
        
        '��Ӳ���
        If rsData!DataType = 12 Or rsData!DataType = 180 Then '�����Ͳ���
            arrPar(rsData!Position - 1) = CDate(Format(rsData!value_String, "mm-dd-yyyy hh:mm:ss"))
        ElseIf rsData!DataType = 2 Then  '������
            arrPar(rsData!Position - 1) = Int(rsData!value_String)
        Else '�ַ���
            arrPar(rsData!Position - 1) = "" & rsData!value_String
        End If
        rsData.MoveNext
    Loop
    

    If rsData.RecordCount = 0 Then
        Call OpenSQLRecord(strNewSQL, "ChangeSQL")
    Else
        Call OpenSQLRecordByArray(strNewSQL, "ChangeSQL", arrPar)
    End If
    
    '������һ��ִ�е�SQLID
    ChangeSQL = GetPrevSQLID(strChildNum)
    If ChangeSQL = "" Then
        ChangeSQL = "5"
        Exit Function
    End If

    Exit Function
errH:
    ChangeSQL = "5"
    If 0 = 1 Then
        Resume
    End If
End Function


Public Function CreateSqlProfiles(strBdSqlID As String, strGdSqlID As String, strChildNum As String) As Boolean
    '���ܣ� ���ݴ����SQLID����SQL PROFILES���ɹ�����True
    '����˵����strBdSqlID-��Ҫ�޸�ִ�мƻ���SQLID��strGdSqlID-�õ�ִ�мƻ���SQLID
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Declare" & vbNewLine & _
                "  Ar_Hint_Table    Sys.Dbms_Debug_Vc2coll; Ar_Profile_Hints Sys.Sqlprof_Attr := Sys.Sqlprof_Attr(); Cl_Sql_Text      Clob;  I                Pls_Integer;" & vbNewLine & _
                "Begin" & vbNewLine & _
                "  With A As(Select Rownum As r_No, a.* From Table( Dbms_Xplan.Display_Cursor('" & strGdSqlID & "', " & strChildNum & ", 'OUTLINE') ) A)," & vbNewLine & _
                "  B As (Select Min(r_No) As Start_r_No From A Where a.Plan_Table_Output = 'Outline Data')," & vbNewLine & _
                "  C As (Select Min(r_No) As End_r_No From A, B Where a.r_No > b.Start_r_No And a.Plan_Table_Output = '  */')," & vbNewLine & _
                "  D As (Select Instr(a.Plan_Table_Output, 'BEGIN_OUTLINE_DATA') As Start_Col From A, B Where r_No = b.Start_r_No + 4)" & vbNewLine & _
                "  Select Substr(a.Plan_Table_Output, d.Start_Col) As Outline_Hints Bulk Collect" & vbNewLine & _
                "  Into Ar_Hint_Table From A, B, C, D Where a.r_No >= b.Start_r_No + 4 And a.r_No <= c.End_r_No - 1 Order By a.r_No;" & vbNewLine & _
                "  Select Sql_FullText Into Cl_Sql_Text From GV$sql Where Sql_Id = '" & strBdSqlID & "' And Rownum<2;" & vbNewLine & _
                "  I := Ar_Hint_Table.First;" & vbNewLine & _
                "  While I Is Not Null Loop" & vbNewLine & _
                "    If Ar_Hint_Table.Exists(I + 1) Then" & vbNewLine & _
                "      If Substr(Ar_Hint_Table(I + 1), 1, 1) = ' ' Then" & vbNewLine & _
                "        Ar_Hint_Table(I) := Ar_Hint_Table(I) || Trim(Ar_Hint_Table(I + 1)); Ar_Hint_Table.Delete(I + 1);" & vbNewLine & _
                "      End If;" & vbNewLine & _
                "    End If;" & vbNewLine & _
                "    I := Ar_Hint_Table.Next(I);" & vbNewLine & _
                "  End Loop;" & vbNewLine & _
                "  I := Ar_Hint_Table.First;" & vbNewLine & _
                "  While I Is Not Null Loop" & vbNewLine & _
                "    Ar_Profile_Hints.Extend;Ar_Profile_Hints(Ar_Profile_Hints.Count) := Ar_Hint_Table(I); I := Ar_Hint_Table.Next(I);" & vbNewLine & _
                "  End Loop;" & vbNewLine & _
                "  Dbms_Sqltune.Import_Sql_Profile(Sql_Text => Cl_Sql_Text, Profile => Ar_Profile_Hints, Name => 'PROFILE_" & strBdSqlID & "'  , Force_Match => True);" & vbNewLine & _
                "End;"
    gcnOracle.Execute strSql
    CreateSqlProfiles = True
    Exit Function
errH:
    MsgBox "���SQL PROFILESʧ�ܣ�����ϵDBA��" & vbNewLine & Err.Description
    CreateSqlProfiles = False
End Function



Public Sub CheckSqlPlan(vsfPlanTbl As VSFlexGrid, ByVal intOptCol As Integer, ByVal intObjCol As Integer, _
                                            rsBigtbl As ADODB.Recordset, rsBigIdx As ADODB.Recordset, rsLowIdx As ADODB.Recordset)
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
    
    On Error GoTo errH
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
                        strObject = Split(Replace(TrimEx(.TextMatrix(i, intOptCol)), strTmp(j), ""), " ")(0)
                        Exit For
                    End If
                Next
            End If
            
            If strOperation <> "" And strObject <> "" Then
                If strOperation = "TABLE ACCESS FULL" Then '��ȡȫ��ɨ��
                    blnTmp = CheckRs(rsBigtbl, "���� = '" & strObject & "'") Or gcnOracle = ""
                ElseIf InStr(1, "INDEX FULL SCAN/INDEX SKIP SCAN/INDEX FAST FULL SCAN", strOperation) > 0 Then '����ȫɨ��\������ɨ��
                    blnTmp = CheckRs(rsBigIdx, "������ = '" & strObject & "'") Or gcnOracle = ""
                ElseIf strOperation = "INDEX RANGE SCAN" And gcnOracle <> "" Then  '������Χɨ��:��Ч����
                    blnTmp = CheckRs(rsLowIdx, "Լ����= '" & GetFkByIdx(strObject) & "'")
                End If
            End If
                
            If blnTmp Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = FULL_��ɫ
            
            strOperation = "": strObject = ""
            blnTmp = False
        Next

    End With
    Exit Sub
errH:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub


Public Sub GetMidTabSize(ByRef lngMinSize As Long, ByRef lngMaxSize As Long)
    '����:��ȡ���ͱ��С
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    lngMinSize = 3000: lngMaxSize = 1000000
    
    On Error GoTo errH
    strSql = "Select A.������,Nvl(A.����ֵ,A.ȱʡֵ) As ����ֵ " & _
                 "From zlParameters A " & _
                 "Where A.������ = '������ͱ�' And a.ϵͳ is null And a.ģ�� is null"
    Set rsTmp = OpenSQLRecord(strSql, "GetMidTabSize")
    
    If rsTmp.EOF Then Exit Sub
    lngMinSize = Split(rsTmp!����ֵ, ",")(0)
    lngMaxSize = Split(rsTmp!����ֵ, ",")(1)
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCheckObj(ByVal intMod As Integer, Optional ByVal lngMinSize As Long, Optional ByVal lngMaxSize As Long) As ADODB.Recordset
'����:��ȡ�漰��������ı�/��������,����һ����¼��
'����intMod: 1-��,2-����,3-��Ч����
'lngMinSize,lngMaxSize - �ж����ͱ������,������Ĭ��Ϊ3000-1000000

    Dim strSql As String
    
    On Error GoTo errH
    
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
errH:
    Set GetCheckObj = Nothing
    MsgBox Err.Description
End Function

Public Function CheckRs(rsData As ADODB.Recordset, ByVal strFilter As String) As Boolean
'����:�Դ���ļ�¼����ӹ���,�����ƥ�����򷵻�True
    
    If rsData Is Nothing Then Exit Function
    rsData.Filter = strFilter
    CheckRs = Not rsData.EOF
    rsData.Filter = 0
End Function

Public Function GetFkByIdx(ByVal strIdxName As String) As String
'����:���ݴ�����������ض�Ӧ�����Լ������
    
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errH:
    
    strSql = "Select Distinct a.Constraint_Name" & vbNewLine & _
                    "From Dba_Cons_Columns A, Dba_Ind_Columns B" & vbNewLine & _
                    "Where a.Owner = b.Table_Owner And a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name And a.Position = b.Column_Position And" & vbNewLine & _
                    "      b.Index_Name = [1]"

    Set rsData = OpenSQLRecord(strSql, "GetFkByIdx", strIdxName)
        
    If Not rsData.EOF Then
        GetFkByIdx = rsData!Constraint_Name & ""
    End If
    Exit Function
errH:
    GetFkByIdx = ""
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Function
