Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Const GSTR_MESSAGE = "��ʾ��Ϣ"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
     
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public gstrUser As String, gstrUserNameNew As String
Public gintUserID As Integer, gintDeptID As Integer
Public gbytЧ�� As Byte

Public gobjComLib As Object                         'zl9Comlib����
Public gcnOutside As New ADODB.Connection           '�ⲿ���ݿ�����

Public Const GSTR_SYSNAME = "�Զ��ְ����ӿ�"
Public Const GSTR_REGEDIT_PATH = "����ģ��\DrugPackerDBServer"
Public Const MSTR_SERVER = "localhost"
Public Const MSTR_DBNAME = "atf"
Public Const MSTR_USER = "sa"
Public Const MSTR_PASSWORD = ""

Public gintMode As Integer                          '����ҵ��0-ZLHISҵ����ã�1-����ҵ�����
Public gblnSys As Boolean
Public gcnWireless As ADODB.Connection

Public gintCount As Integer                         '����ⲿ���ʵ��������Nothing������ȫ�����ͷź���ͷ��ڲ���ȫ�ֱ���

Public Function MSSQLServerOpen(ByVal strServerName As String, ByVal strDBName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ����MS SQL Server ���ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    If Len(Trim(strUserName)) = 0 Then
        MSSQLServerOpen = False
        MsgBox "�������������ݿ���Ϣ��", vbInformation, GSTR_MESSAGE
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .ConnectionTimeout = 5
        .Open "Driver={SQL Server};Server=" & strServerName & ";Database=" & strDBName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ������������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Or Err.Number = -2147467259 Then
                MsgBox "ҩƷ�ְ������ݿ�����ʧ�ܣ�", vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            MSSQLServerOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    MSSQLServerOpen = True
    Exit Function
    
errHand:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter() = 1 Then Resume
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    MSSQLServerOpen = False
    Err = 0
End Function


Public Function OraDataOpenTest(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
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
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ������������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Then
                MsgBox Mid(strError, InStr(1, strError, "[SQL Server]"), Len(strError)), vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            OraDataOpenTest = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpenTest = True
    Exit Function
    
errHand:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter() = 1 Then Resume
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    OraDataOpenTest = False
    Err = 0
End Function

Public Function StringEnDeCodecn(strSource As String, MA) As String
'�ú���ֻ���������𵽼�������
'����Ϊ��Դ�ļ�������
    On Error GoTo ErrEnDeCode
    Dim X As Single, i As Integer
    Dim CHARNUM As Long, RANDOMINTEGER As Integer
    Dim SINGLECHAR As String * 1
    Dim strTmp As String
    
    If MA < 0 Then
        MA = MA * (-1)
    End If
    
    X = Rnd(-MA)
    For i = 1 To Len(strSource) Step 1                 'ȡ���ֽ�����
        SINGLECHAR = Mid(strSource, i, 1)
        CHARNUM = Asc(SINGLECHAR)
g:
        RANDOMINTEGER = Int(127 * Rnd)
        If RANDOMINTEGER < 30 Or RANDOMINTEGER > 100 Then GoTo g
        CHARNUM = CHARNUM Xor RANDOMINTEGER
        strTmp = strTmp & Chr(CHARNUM)
    Next i
    StringEnDeCodecn = strTmp
    Exit Function

ErrEnDeCode:
    StringEnDeCodecn = ""
    MsgBox Err.Number & "\" & Err.Description
End Function

'Public Function GetUserNameInfo() As Boolean
''��ȡ�û���Ϣ
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL As String
'
'    Set rsTmp = gobjComLib.GetUserInfo
'
'    With rsTmp
'        If Not .EOF Then
'            gintUserID = IIf(IsNull(!Id), 0, !Id)
'            gintDeptID = IIf(IsNull(!����id), 0, !����id)
'            gstrUserNameNew = IIf(IsNull(!����), "", !����) '��ǰ�û�����
'            GetUserNameInfo = True
'        Else
'            gintUserID = 0
'            gintDeptID = 0
'            gstrUserNameNew = "" '��ǰ�û�����
'        End If
'    End With
'    rsTmp.Close
'
'    strSQL = "Select ������, ����ֵ, ȱʡֵ From Zlparameters Where ϵͳ = [1] And Nvl(˽��, 0) = 0 And ģ�� Is Null and ������=[2] "
'    Set rsTmp = gobjComLib.OpenSQLRecord(strSQL, "ȡϵͳ����", 100, 149)
'    With rsTmp
'        If Not .EOF Then
'            gbytЧ�� = IIf(IsNull(rsTmp!����ֵ), rsTmp!ȱʡֵ, rsTmp!����ֵ)
'        Else
'            gbytЧ�� = 0
'        End If
'    End With
'
'End Function

'Public Function CheckProvider(ByVal intProvider As Integer) As String
''��˹�Ӧ��ID
'    Dim rsTmp As New ADODB.Recordset
'    Set rsTmp = zlDatabase.OpenSQLRecord("select ���� from ��Ӧ�� where id=[1]", "��˹�Ӧ��ID", intProvider)
'    If rsTmp.RecordCount = 1 Then
'        CheckProvider = rsTmp!����
'    End If
'    rsTmp.Close
'End Function

Public Sub SelText(ByVal ctlVal As Control)
    If TypeOf ctlVal Is TextBox Then
        ctlVal.SelStart = 0
        ctlVal.SelLength = Len(ctlVal.Text)
    End If
End Sub










'---------------------------------

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
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

    '������ʹ���˶�̬�ڴ��������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    
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
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
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
    'If cmdData.ActiveConnection Is Nothing Then
'        If gblnSys = True Then
'            Set cmdData.ActiveConnection = gcnSysConn
'        Else
            Set cmdData.ActiveConnection = gcnWireless '���Ƚ���(���ִ��1000��Լ0.5x��)
'        End If
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
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

Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    Static srsInfo As New ADODB.Recordset
    Static sblnTemp As Boolean
    Dim strInfo As String, aryInfo() As String
    Dim strSQL As String
    
    On Error GoTo errHand
    If blnTemp Or sblnTemp <> blnTemp Or (srsInfo.State <> adStateOpen) Then
        sblnTemp = blnTemp
        strSQL = "Select Item,Text From Table(Cast(zltools.f_Reg_Info([1]) As zlTools.t_Reg_Rowset))"
        Set srsInfo = mdlDrugPacker.OpenSQLRecord(strSQL, "zlRegInfo", IIf(blnTemp, 1, 0))
    End If
    
    srsInfo.Filter = "Item='" & strItem & "'"
    If srsInfo.RecordCount <> 1 Then zlRegInfo = "": Exit Function
    strInfo = "" & srsInfo!Text
    If (strItem = "��λ����" Or strItem = "��Ʒ������" Or strItem = "����֧����") And intBits <> -1 Then
        aryInfo = Split(strInfo, ";")
        If intBits > UBound(aryInfo) Then
            strInfo = ""
        Else
            strInfo = aryInfo(intBits)
        End If
    End If
    zlRegInfo = strInfo
    Exit Function
errHand:
    zlRegInfo = ""
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ���ʽ,��ʱ�����޷������󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷������󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
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
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnWireless '���Ƚ���
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnWireless.Execute strSQL, , adCmdText
    
'    Call gobjComLib.SQLTest
End Sub

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
        .Open "SELECT SYSDATE FROM DUAL", gcnWireless, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function

errH:
    MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    Currentdate = 0
    Err = 0
End Function

Public Function GetINIInfo(ByVal strKey As String) As String
    Dim strReturn As String * 1024
    Dim strValue As String
    
    GetPrivateProfileString "CONFIG", strKey, "", strReturn, Len(strReturn), App.Path & "\zlDrugPacker.ini"
    
    strValue = Left(strReturn, InStr(strReturn, vbNullChar) - 1)
    GetINIInfo = strValue
End Function