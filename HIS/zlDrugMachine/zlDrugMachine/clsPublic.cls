VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_MSG As String = "clsPublic"
Private mcnOracle As ADODB.Connection
Private mrsParas As ADODB.Recordset
Private mrsDeptParas As ADODB.Recordset
Private mrsUserParas As ADODB.Recordset

Public Sub InitCommon(cnMain As ADODB.Connection)
    Set mcnOracle = cnMain
End Sub

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
            Set cmdData.ActiveConnection = mcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
'        End If
    'End If
    cmdData.CommandText = strSQL
    
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
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
            Set cmdData.ActiveConnection = mcnOracle    '���Ƚ���
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
    
NoneVarLine:
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    mcnOracle.Execute strSQL, , adCmdText
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
        .Open "SELECT SYSDATE FROM DUAL", mcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function

errH:
'    MsgBox Err.Description, vbInformation, MSTR_MSG
    Currentdate = 0
    Err = 0
End Function

Private Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
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

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, _
    Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
    
'���ܣ���ȡָ���Ĳ���ֵ
'������varPara=�����Ż�������������ֻ��ַ����ʹ�������
'      lngSys=ʹ�øò�����ϵͳ��ţ���100
'      lngModual=ʹ�øò�����ģ��ţ���1230
'      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
'      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
'      arrControl=�ؼ����飬��Array(Me.Text1, Me.CheckBox1)�����ں����ڲ��Զ�������Ӧ�ؼ�����ʾ��ɫ���Ƿ��ֹ���á�
'      blnSetup=����ģ���Ƿ��в�������Ȩ��
'      intType=���ز��������ز�������
'      lngDeptID=���ż����������ݲ��Ŷ�ȡ
'���أ�����ֵ���ַ�����ʽ

    Dim strSQL As String, i As Integer
    Dim blnNew As Boolean, blnEnabled As Boolean, blnNewRow As Boolean, blnNotExists As Boolean
    Dim strSqlFilter As String, strResFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim blnNewVersion As Boolean
    
    On Error GoTo errH
    
    intType = 0
    
    '��������
    If TypeName(varPara) = "String" Then
        strResFilter = "������='" & CStr(varPara) & "' And ģ��=" & lngModual & " And ϵͳ=" & lngSys
        strSqlFilter = "������=[5] And Nvl(ģ��,0)=[3] And Nvl(ϵͳ,0)= [4] "
    Else
        strResFilter = "������=" & Val(varPara) & " And ģ��=" & lngModual & " And ϵͳ=" & lngSys
        strSqlFilter = "������=[6] And Nvl(ģ��,0)=[3] And Nvl(ϵͳ,0)=[4] "
    End If
    
    '���������ж�
    If mrsParas Is Nothing Then
        blnNew = True
    ElseIf mrsParas.State = 0 Then
        blnNew = True
    Else
        mrsParas.Filter = strResFilter
        blnNewRow = mrsParas.EOF
    End If
    
    If blnNew Or blnNewRow Then
        '��ȡ�汾��
        strSQL = "Select �汾�� From zlSystems Where ��� = 100 "
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡZLHIS���ݿ�汾")
        If rsTmp.EOF = False Then
            blnNewVersion = NVL(rsTmp!�汾��) >= "10.35.0"
        End If
        rsTmp.Close
        
        '����������ȡ��������
        If blnNewVersion Then
            '10.35.0���ϰ汾
            strSQL = "Select ID,Nvl(ϵͳ,0) as ϵͳ,Nvl(ģ��,0) as ģ��,Nvl(˽��,0) as ˽��,Nvl(����,0) as ����,NVL(����,0) as ����,Nvl(��Ȩ,0) as ��Ȩ,������,������," & _
                " Nvl(����ֵ,ȱʡֵ) as ����ֵ,[1] as �û���,[2] as ������ " & vbNewLine & _
                "From zlParameters " & vbNewLine & _
                "Where " & strSqlFilter
        Else
            '10.35.0���°汾
            strSQL = "Select ID,Nvl(ϵͳ,0) as ϵͳ,Nvl(ģ��,0) as ģ��,Nvl(˽��,0) as ˽��,Nvl(����,0) as ����,0 as ����,Nvl(��Ȩ,0) as ��Ȩ,������,������," & _
                " Nvl(����ֵ,ȱʡֵ) as ����ֵ,[1] as �û���,[2] as ������ " & vbNewLine & _
                "From zlParameters " & vbNewLine & _
                "Where " & strSqlFilter
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "GetPara", CStr("ZLHIS"), CStr("P"), lngModual, lngSys, CStr(varPara), Val(varPara))
    
        If rsTmp.EOF Then
            blnNotExists = True
        Else
            If blnNewRow Then
                Call Me.Append(mrsParas, rsTmp)
            Else
                Set mrsParas = CopyNewRec(rsTmp)
            End If
            mrsParas.Filter = strResFilter '�ٴι���
            '��ȡ���ż�����
            If mrsParas!���� = 1 Then
                strSQL = "Select ����id, ����ID, ����ֵ" & vbNewLine & _
                        "From Zldeptparas" & vbNewLine & _
                        "Where ����id = [1]"
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", Val(mrsParas!ID), lngDeptID)
                
                If mrsDeptParas Is Nothing Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                ElseIf mrsDeptParas.State = 0 Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                Else
                    Call Me.Append(mrsDeptParas, rsTmp)
                End If
            '��ȡ�û��򱾻�����
            ElseIf mrsParas!˽�� = 1 Or mrsParas!���� = 1 Then
                strSQL = "Select ����id, Nvl(�û���, 'NullUser') As �û���, Nvl(������, 'NullMachine') As ������, ����ֵ" & vbNewLine & _
                        "From zlUserParas" & vbNewLine & _
                        "Where ����id = [3]"
                        
                If mrsParas!˽�� = 1 And mrsParas!���� = 1 Then
                    strSQL = strSQL & " And �û���=[1] And ������=[2]"
                ElseIf mrsParas!˽�� = 1 Then
                    strSQL = strSQL & " And �û���=[1] "
                Else
                    strSQL = strSQL & " And ������=[2]"
                End If
                
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", "", "", Val(mrsParas!ID))
                
                If mrsUserParas Is Nothing Then
                    Set mrsUserParas = CopyNewRec(rsTmp)
                ElseIf mrsUserParas.State = 0 Then
                    Set mrsUserParas = CopyNewRec(rsTmp)
                Else
                    Call Append(mrsUserParas, rsTmp)
                End If
            End If
        End If
    End If

    If blnNotExists Then
        GetPara = strDefault
    Else
        '��ȡ����ֵ
        If mrsParas!���� = 1 Then
            mrsDeptParas.Filter = "����ID=" & mrsParas!ID & _
                 " And ����ID=" & lngDeptID
            If Not mrsDeptParas.EOF Then
                GetPara = NVL(mrsDeptParas!����ֵ, strDefault)
            Else
                GetPara = NVL(mrsParas!����ֵ, strDefault)
            End If
        ElseIf mrsParas!˽�� = 1 Or mrsParas!���� = 1 Then
            mrsUserParas.Filter = "����ID=" & mrsParas!ID & _
                IIf(mrsParas!˽�� = 1, " And �û���='" & mrsParas!�û��� & "'", " And �û���='NullUser'") & _
                IIf(mrsParas!���� = 1, " And ������='" & mrsParas!������ & "'", " And ������='NullMachine'")
            If Not mrsUserParas.EOF Then
                GetPara = NVL(mrsUserParas!����ֵ, strDefault)
            Else
                GetPara = NVL(mrsParas!����ֵ, strDefault)
            End If
        Else
            GetPara = NVL(mrsParas!����ֵ, strDefault)
        End If
        
        '���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
        If mrsParas!���� = 1 Then
            intType = 3   '���ż�����ģ�鹫����������
        ElseIf mrsParas!ϵͳ <> 0 And mrsParas!ģ�� = 0 And mrsParas!˽�� = 0 And mrsParas!���� = 0 Then
            intType = 1
        ElseIf mrsParas!ģ�� = 0 And mrsParas!˽�� = 1 And mrsParas!���� = 0 Then
            intType = 2
        ElseIf mrsParas!ϵͳ <> 0 And mrsParas!ģ�� <> 0 And mrsParas!˽�� = 0 And mrsParas!���� = 0 Then
            intType = 3
        ElseIf mrsParas!ϵͳ <> 0 And mrsParas!ģ�� <> 0 And mrsParas!˽�� = 1 And mrsParas!���� = 0 Then
            intType = 4
        ElseIf mrsParas!ϵͳ <> 0 And mrsParas!ģ�� <> 0 And mrsParas!˽�� = 0 And mrsParas!���� = 1 Then
            intType = IIf(mrsParas!��Ȩ = 1, 15, 5)
        ElseIf mrsParas!ϵͳ <> 0 And mrsParas!ģ�� <> 0 And mrsParas!˽�� = 1 And mrsParas!���� = 1 Then
            intType = 6
        End If
        
        '������Ӧ�Ŀؼ���ɫ���ɿ�״̬
        If IsArray(arrControl) And (intType = 3 Or (intType Mod 10) = 5) Then
            blnEnabled = Not ((intType = 3 Or (intType Mod 10) = 5 And mrsParas!��Ȩ = 1) And Not blnSetup)
            For i = 0 To UBound(arrControl)
                Select Case TypeName(arrControl(i))
                Case "Label"
                    arrControl(i).ForeColor = vbBlue
                Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "CommandButton", "DTPicker"
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "MSHFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "VSFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Editable = 0
                Case Else
                    On Error Resume Next
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                    Err.Clear: On Error GoTo errH
                End Select
            Next
        End If
    End If
    
    Exit Function
    
errH:
End Function

Private Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim i As Long
    
    On Error GoTo errH
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '������¼���ṹ
        If Not rsClone Is Nothing Then
            If strFields = "" Then '��¼��ȫ����ģʽ
                arrFieldsName = Array()
                If rsClone.Fields.Count > 0 Then
                    ReDim arrFieldsName(rsClone.Fields.Count - 1)
                Else
                    arrFieldsName = Array()
                End If
                For intFields = 0 To rsClone.Fields.Count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).Name & ""
                    .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
                Next
            Else '��¼�����ָ���ģʽ
                If rsClone.Fields.Count > 0 Then
                    arrFieldsName = Split(strFields, ",")
                    For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                        '�а�������
                        arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                        strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                        If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).Name & ""
                        '��ȡ�ֶ�ԭ������������
                        arrFieldsName(intFields) = strFieldName
                        '�����ֶ�,�������ڱ������������е�����Ϊ����
                        .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
                    Next
                End If
            End If
        End If
        '׷���ֶ�����
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '��������
        If Not blnOnlyStructure And Not rsClone Is Nothing Then
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳�����ӣ���˿�������
                    .Fields(intFields).Value = rsClone.Fields(arrFieldsName(intFields)).Value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
    Exit Function
    
errH:
End Function

Function Append(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'���ܣ���ָ����¼�����������ӵ���һ����¼����
'������rsSource=Ŀ���¼��
'      rsAppend=���ݼ�¼��
'      arrInput=�ֶζ�Ӧ���򣬸ò�������ʱ��Ĭ������¼���ṹ��ͬ����ʽ��arrInput(0):[��¼��1].�ֶ�1,�ֶ�2...��arrInput(1)��[��¼��2].�ֶ�1,�ֶ�2...
'���أ��Ƿ�ɹ�
'      rsSource=�������ݺ�ļ�¼��
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then Append = True: Exit Function
    If rsAppend.RecordCount = 0 Then Append = True: Exit Function
    If rsSource Is Nothing Then Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).Value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.Count - 1
                rsSource.Fields(i).Value = rsAppend.Fields(i).Value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    Append = True
    Exit Function
    
errH:
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function