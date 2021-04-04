VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private gstrSysName As String
Private mcnOracle As ADODB.Connection
Private mstrSQL As String
Private mobjRegister As Object

'######################################################################################################################

Public Property Get DatabaseState() As Long
    DatabaseState = mcnOracle.State
End Property

Public Property Get DatabaseConnection() As ADODB.Connection
    Set DatabaseConnection = mcnOracle
End Property

Public Sub BeginTrans()
    mcnOracle.BeginTrans
End Sub

Public Sub CommitTrans()
    mcnOracle.CommitTrans
End Sub

Public Sub RollbackTrans()
    mcnOracle.RollbackTrans
End Sub

Public Function UpdateUserPassword(ByVal strUserName As String, ByVal strPasswd As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Call CreateRegister
    Call mobjRegister.UpdateUserPassword(mcnOracle, strUserName, strPasswd, True)
End Function

Public Sub SetRunWithRBO()
    '******************************************************************************************************************
    '���ܣ���ǰ�Ự��RBO�Ż���ģʽ����SQL���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    
    mstrSQL = "alter session set optimizer_mode=rule"
    On Error Resume Next
    mcnOracle.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "�����õ�ǰ�Ự��RBO�Ż���ģʽ���У�", vbInformation, gstrSysName
    End If
End Sub

Public Function SetSQLTrace(ByVal strUserName As String, ByVal strServerName As String) As String
    '******************************************************************************************************************
    '���ܣ�����100046�¼�����SQL Trace����
    '������
    '���أ�Trc�ļ���
    '******************************************************************************************************************
    Dim strSQL As String, strLevel As String, strFile As String
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        mcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size=10M"
        mcnOracle.Execute strSQL
        Err.Clear
        
        '������һ�������8.1.7���Ժ��֧��
        strFile = "ZL_" & strUserName
        strSQL = "alter session set tracefile_identifier='" & strFile & "'"
        mcnOracle.Execute strSQL
        If Err.Number <> 0 Then strFile = "*.trc": Err.Clear
        
        strLevel = "1"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        mcnOracle.Execute strSQL
        If Err.Number = 0 Then SetSQLTrace = strFile
    End If
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional ByVal blnOLEDB As Boolean = False) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ�������ݿ�
    '������ strServerName�������ַ���
    '       strUserName���û���
    '       strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '******************************************************************************************************************
    Dim strSQL As String
    Dim strError As String
    Dim blnTransPassword As Boolean
    Call CreateRegister
    On Error Resume Next
    blnTransPassword = Not (UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Or UCase(strUserName) = "ZLTOOLS")
    Set mcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, IIf(blnOLEDB, 1, 0))
    If mcnOracle.State = adStateClosed Then
        Exit Function
    End If
    Err = 0
    OraDataOpen = True
    Exit Function
    
errHand:
    OraDataOpen = False
    Err = 0
End Function

Public Function OpenCursor(ByVal strFormCaption As String, _
                           ByVal strOwner As String, _
                           ByVal strPackagesName As String, _
                           ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'���ܣ����ô洢���̷��ؼ�¼��
'��Σ�strPackagesName ����ʽΪ ��.������
'-----------------------------------------
    Dim cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, i As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '��¼��������
    Dim varOutPar As Variant
    Dim strNode As String
    '����������
    If strOwner <> "" Then
        strPackagesName = strOwner & "." & strPackagesName
    End If
    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdPackage.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdPackage.Parameters.Count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN ����
    strNode = ""
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
        strNode = strNode & CStr(varValue) & ","
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = mcnOracle
    End If
    
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strPackagesName, strNode)
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    OpenCursor.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
    cmdPackage.Properties("PLSQLRSet") = False

End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    '******************************************************************************************************************
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
    '******************************************************************************************************************
    
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
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
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
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
        Set cmdData.ActiveConnection = mcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Function ExecuteSQL(strSQL As String, ByVal strFormCaption As String, Optional ByVal blnProc As Boolean = False) As Boolean
    
    On Error GoTo errHand
    
    If blnProc Then
        mcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        mcnOracle.Execute strSQL
    End If
    
    ExecuteSQL = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function

Public Function ExecuteProcedure(strSQL As String, ByVal strFormCaption As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
    '������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
    '˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
    '  1.���������Ǳ���ʽ,��ʱ�����޷������󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
    '  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷������󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
    '  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    '******************************************************************************************************************
    
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '���ԭ�в���:��Ȼ�����ظ�ִ��
'        cmdData.CommandText = "" '��Ϊ����ʱ�����������
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ�����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULLֵ�������ִ����ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ������ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL�������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı���ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
            Exit Function
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = mcnOracle '���Ƚ���
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    
    ExecuteProcedure = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    mcnOracle.Execute strSQL, , adCmdText
    
    ExecuteProcedure = True
    
'    Call gobjComLib.SQLTest
End Function

Public Function Currentdate() As Date
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
        .Open "SELECT SYSDATE FROM DUAL", mcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

'######################################################################################################################
Private Sub Class_Initialize()
    Set mcnOracle = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    If Not (mcnOracle Is Nothing) Then
        If mcnOracle.State = adStateOpen Then mcnOracle.Close
        Set mcnOracle = Nothing
    End If
End Sub


Private Sub CreateRegister()
    '���������ڵ�¼֮ǰ�Ѵ������Լ�ҽ�������ȿ����ٴε��ã������ظ���������
    If mobjRegister Is Nothing Then
        On Error Resume Next
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        
        If mobjRegister Is Nothing Then
            '����֧��δͨ������̨����������prjMain�����ñ������������
            '��������£�����û�е���GetConnection,�޷�ͨ���ö����ȡ�����������û������л�����
            Set mobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If mobjRegister Is Nothing Then
                MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
    End If
End Sub