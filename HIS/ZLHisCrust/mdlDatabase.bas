Attribute VB_Name = "mdlDatabase"
Option Explicit
Public Function ErrCenter() As Byte
'���ܣ� �����������������
'������
'���أ� cancel      ���� 0
'       resume      ���� 1
    Dim strNote As String
    
        '---------------VB��׼����--------------------------
    Select Case Err.Number
        Case 3, 3 - 2146828288
            strNote = "δ���ñ�׼���ع���"
        Case 5, 5 - 2146828288
            strNote = "��Ч�Ĺ��̻����"
        Case 6, 6 - 2146828288
            strNote = "�������"
        Case 7, 7 - 2146828288
            strNote = "�ڴ����"
        Case 9, 9 - 2146828288
            strNote = "�±곬��"
        Case 10, 10 - 2146828288
            strNote = "�����ǹ̶��������ʱ����"
        Case 11, 11 - 2146828288
            strNote = "����Ϊ��̫С"
        Case 13, 13 - 2146828288
            strNote = "���Ͳ�ƥ��"
        Case 14, 14 - 2146828288
            strNote = "�����ַ���������"
        Case 16, 16 - 2146828288
            strNote = "���ʽ̫����"
        Case 17, 17 - 2146828288
            strNote = "��֧��Ҫ��Ĳ���"
        Case 18, 18 - 2146828288
            strNote = "�������û��ж�"
        Case 20, 20 - 2146828288
            strNote = "�޴��󷵻�"
        Case 28, 28 - 2146828288
            strNote = "��ջ�ռ����"
        Case 35, 35 - 2146828288
            strNote = "���̻���δ����"
        Case 47, 47 - 2146828288
            strNote = " ̫��Ķ�̬����⣨DLL��Ӧ�ÿͻ�"
        Case 48, 48 - 2146828288
            strNote = " ���ö�̬����⣨DLL������"
        Case 49, 49 - 2146828288
            strNote = " ��̬����⣨DLL��Լ������"
        Case 51, 51 - 2146828288
            strNote = "�ڲ�����"
        Case 52, 52 - 2146828288
            strNote = "������ļ������ļ���"
        Case 53, 53 - 2146828288
            strNote = "�ļ�δ�ҵ�"
        Case 54, 54 - 2146828288
            strNote = "�ļ���ʽ����"
        Case 55, 55 - 2146828288
            strNote = "�ļ��Ѿ���"
        Case 57, 57 - 2146828288
            strNote = "�豸���� / �������"
        Case 58, 58 - 2146828288
            strNote = "�ļ��Ѿ�����"
        Case 59, 59 - 2146828288
            strNote = "����ļ�¼����"
        Case 61, 61 - 2146828288
            strNote = "������"
        Case 62, 62 - 2146828288
            strNote = "���볬���ļ�β"
        Case 63, 63 - 2146828288
            strNote = "����ļ�¼��"
        Case 67, 67 - 2146828288
            strNote = "�ļ�̫��"
        Case 68, 68 - 2146828288
            strNote = "�豸��Ч��֧��"
        Case 70, 70 - 2146828288
            strNote = "�ܾ�����"
        Case 71, 71 - 2146828288
            strNote = "����δ׼����"
        Case 74, 74 - 2146828288
            strNote = "��������Ϊ��ͬ��������"
        Case 75, 75 - 2146828288
            strNote = "·�� / �ļ����ʴ���"
        Case 76, 76 - 2146828288
            strNote = "·��δ�ҵ�"
        Case 91, 91 - 2146828288
            strNote = "�������������Ϊ����(δ�½�ʵ��)"
        Case 92, 92 - 2146828288
            strNote = "ѭ��δ��ʼ��"
        Case 93, 93 - 2146828288
            strNote = "�����ģʽ�ַ���"
        Case 94, 94 - 2146828288
            strNote = "�����ʹ�ÿ�(Null)"
        Case 96, 96 - 2146828288
            strNote = " �����Ѿ�ʹ�õĶ���ʱ�䳬���������õ����Ԫ�غţ����²����ܽ����¼�"
        Case 97, 97 - 2146828288
            strNote = "���ܵ���һ��δ����ʵ�����������"
        Case 98, 98 - 2146828288
            strNote = " ����ʹ��һ��˽�ж�������Ժͷ���?�����ͷ���ֵ"
        Case 321, 321 - 2146828288
            strNote = "������ļ���ʽ"
        Case 322, 322 - 2146828288
            strNote = "���ܴ�����Ҫ����ʱ�ļ�"
        Case 325, 325 - 2146828288
            strNote = "��Դ�ļ��д���ĸ�ʽ"
        Case 380, 380 - 2146828288
            strNote = "���������ֵ"
        Case 381, 381 - 2146828288
            strNote = "�����������������"
        Case 382, 382 - 2146828288
            strNote = "��֧�ֵ�����ʱ����"
        Case 383, 383 - 2146828288
            strNote = "��֧�ֵ�ֻ����������"
        Case 385, 384 - 2146828288
            strNote = "��Ҫ������������"
        Case 387, 387 - 2146828288
            strNote = "�����������"
        Case 393, 393 - 2146828288
            strNote = "��֧�ֵ�����ʱ��ȡ"
        Case 394, 394 - 2146828288
            strNote = "��֧�ֵ�ֻд���Զ�ȡ"
        Case 422, 422 - 2146828288
            strNote = "�����ڵ�����"
        Case 423, 423 - 2146828288
            strNote = "�����ڵ����Ի򷽷�"
        Case 424, 424 - 2146828288
            strNote = "Ҫ��һ������"
        Case 429, 429 - 2146828288
            strNote = "ActiveX���ܴ�������"
        Case 430, 430 - 2146828288
            strNote = "�಻֧�ֵ��Զ���������֧�ֵĽ���"
        Case 432, 432 - 2146828288
            strNote = "���Զ������ڼ�δ�ҵ��ļ�����������"
        Case 438, 438 - 2146828288
            strNote = "����֧�ָ����Ի򷽷�"
        Case 440, 440 - 2146828288
            strNote = "�Զ����������"
        Case 442, 442 - 2146828288
            strNote = "��Զ��������������ᶪʧ����OK����Ի���ȥ����"
        Case 443, 443 - 2146828288
            strNote = "�Զ�������û��ȱʡֵ"
        Case 445, 445 - 2146828288
            strNote = "����֧�����ֲ���"
        Case 446, 446 - 2146828288
            strNote = "����֧����������"
        Case 447, 447 - 2146828288
            strNote = "����֧�ֵ�ǰ��������"
        Case 448, 448 - 2146828288
            strNote = "��������δ�ҵ�"
        Case 449, 449 - 2146828288
            strNote = "�������ǿ�ѡ��"
        Case 450, 450 - 2146828288
            strNote = "����Ĳ������������Է���"
        Case 451, 451 - 2146828288
            strNote = "���Ը�ֵ(Let)���̺Ͷ�ȡ(Get)���̲����ض���"
        Case 452, 452 - 2146828288
            strNote = "��Ч�����"
        Case 453, 453 - 2146828288
            strNote = "ָ����DLL����δ�ҵ�"
        Case 454, 454 - 2146828288
            strNote = "������Դδ�ҵ�"
        Case 455, 455 - 2146828288
            strNote = "������Դ��������"
        Case 457, 457 - 2146828288
            strNote = "�ùؼ�ֵ�Ѿ��뼯�ϵ���һԪ�ؽ��"
        Case 458, 458 - 2146828288
            strNote = "VB��֧�ֵĿɱ��Զ�������"
        Case 459, 459 - 2146828288
            strNote = "������಻֧�ֵ��¼���"
        Case 460, 460 - 2146828288
            strNote = "����ļ������ʽ"
        Case 461, 461 - 2146828288
            strNote = "���������ݳ�Աδ�ҵ�"
        Case 462, 462 - 2146828288
            strNote = "Զ�̷����������ڻ���Ч"
        Case 463, 463 - 2146828288
            strNote = "��û���ڱ���ע��"
        Case 481, 481 - 2146828288
            strNote = "��Ч��ͼƬ��ʽ"
        Case 482, 482 - 2146828288
            strNote = "��ӡ������"
        Case 735, 735 - 2146828288
            strNote = "���ܽ��洢Ϊ��ʱ�ļ�"
        Case 744, 744 - 2146828288
            strNote = "δ�ҵ�����������"
        Case 746, 746 - 2146828288
            strNote = "̫���ĸ���"
        '------------------ADO����-------------------
        Case 3001
            strNote = "�������ʹ��󣬻���ֵ������Χ�������ͻ��"
        Case 3021
            strNote = "��¼����(EOF/BOF)�����ߵ�ǰ��¼��ɾ������ǰӦ�ò�����Ҫ��λ��ǰ��¼��"
        Case 3219
            strNote = "�����Ļ���������ǰӦ�ò����������Ǵ�����δ���������񣩡�"
        Case 3246
            strNote = "������ִ���У����ܹر�һ���������"
        Case 3251
            strNote = "��ǰ������֧����һӦ�ò�����"
        Case 3265
            strNote = "ADOû�ҵ�Ӧ�ó���Ҫ��Ķ�Ӧ���ƻ���š�"
        Case 3367
            strNote = "�����Ѿ����ڣ�������ӡ�"
        Case 3420
            strNote = "����δ���á�"
        Case 3421
            strNote = "��ǰ����ʹ���˴������ֵ���͡�"
        Case 3704
            strNote = "����ر�ʱ����ǰ��������ִ�С�"
        Case 3705
            strNote = "������ʱ����ǰ��������ִ�С�"
        Case 3706
            strNote = "ADOû�ҵ�ָ����֧�֡�"
        Case 3707
            strNote = "���ܲ����������ı�һ����¼���Ļ����Դ�����ԡ�"
        Case 3708
            strNote = "Ӧ�ó�����ִ���Ĳ������塣"
        Case 3709
            strNote = "Ӧ�ó���Ҫ��һ���رյ����ö������Ч���������"
        Case Else
            strNote = Err.Description
    End Select
    
    ErrCenter = frmErrAsk.ShowForm(Err.Number, strNote)
    Err.Clear
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
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
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = ActualLen(varValue)
            
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
                intMax = ActualLen(varValue(lngLeft))
                            
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
    Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    cmdData.CommandText = strSQL
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
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
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
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
                        intMax = ActualLen(strPar)
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
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
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
        
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
        cmdData.CommandType = adCmdText
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
    gcnOracle.Execute strSQL, , adCmdText
End Sub


Public Function Currentdate() As Date
'���ܣ���ȡ�������ϵ�ǰ����
'������
'���أ�����Oracle���ڸ�ʽ�����⣬����
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    Set rsTemp = OpenSQLRecord("SELECT SYSDATE FROM DUAL", App.Title)
    Currentdate = rsTemp!SYSDATE
    Exit Function
ErrH:
    Currentdate = 0
End Function


Public Function IsWriteRunErrLog() As Boolean
'����:�Ƿ��¼���д���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    On Error GoTo ErrH
    '3   �Ƿ��¼���д���(�Ƿ��¼ʹ�ù����з����ĸ��ִ���)
    strSQL = "select ����ֵ from ZLTOOLS.ZlOptions where ������=3"
    Set rsTmp = OpenSQLRecord(strSQL, "������")

    If Not rsTmp.EOF Then
         strReturn = NVL(rsTmp!����ֵ, "0")
    Else
         strReturn = "0"
    End If
    IsWriteRunErrLog = strReturn <> "0"
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetNoteLength() As Long
'���ܣ���ȡһ�������ֶζ��峤��
'������strTable=����
'        strColumn=����
'���أ������г���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLength As Long
    
    lngLength = 300
    strSQL = "Select Column_Name,Nvl(Data_Precision, Data_Length) Collen ,Owner" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name =[2]"
    On Error GoTo ErrH
    Set rsTmp = OpenSQLRecord(strSQL, "FieldsLength", "ZLCLIENTS", "˵��")
    If Not rsTmp.EOF Then
        rsTmp.Filter = "Owner='ZLTOOLS'"
        If Not rsTmp.EOF Then
            lngLength = Val(rsTmp!collen)
        Else
            rsTmp.Filter = ""
            rsTmp.Sort = "Owner"
            lngLength = Val(rsTmp!collen)
        End If
    End If
    GetNoteLength = lngLength
    Exit Function
ErrH:
    GetNoteLength = lngLength
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function IsSampleFTP() As Boolean
'����:�Ƿ�ʹ�ü���FTP
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrH
    strSQL = "Select Nvl(Max(����), 0) As ʹ�ü���ftp���� From Zlreginfo Where ��Ŀ = 'FTP������ļ�����'"
    Set rsTmp = OpenSQLRecord(strSQL, "ʹ�ü���ftp����")
    IsSampleFTP = Val(rsTmp!ʹ�ü���ftp���� & "") <> 0
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsHaveVersion() As Boolean
'���ܣ���ȡһ�������ֶζ��峤��
'������strTable=����
'        strColumn=����
'���أ������г���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLength As Long
    
    lngLength = 300
    strSQL = "Select Column_Name,Owner" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name =[2] And Owner='ZLTOOLS'"
    On Error GoTo ErrH
    Set rsTmp = OpenSQLRecord(strSQL, "FieldsLength", "ZLFILESUPGRADE", "�ļ��汾��")
    If Not rsTmp.EOF Then
        IsHaveVersion = True
    End If
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
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
    
    On Error GoTo ErrH
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
                        '����ֶ�,�������ڱ������������е�����Ϊ����
                        .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
                    Next
                End If
            End If
        End If
        '׷���ֶ����
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
                    '�¼�¼�����а�˳����ӣ���˿�������
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
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdateRec(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'���ܣ�����ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'      arrInput=������ֶ����Լ�ֵ����ʽ���ֶ���1,ֵ1, �ֶ���2,ֵ2,....
'���أ��Ƿ�ɹ�
'      rsInput=�������º�ļ�¼��
'˵����arrInput���ֶ�ֵ�����ü�¼���е������ֶ������¸��ֶΣ���ʱ��ʽΪ��!�ֶ���
    Dim strFiledName As String, strFileValue As String
    Dim blnFiled As Boolean, i As Long

    On Error GoTo ErrH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If IsNull(arrInput(i + 1)) Then
                    rsInput(strFiledName).Value = Null
                Else
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFileValue = rsInput(Mid(arrInput(i + 1), 2)).Value & ""
                        If Err.Number <> 0 Then Err.Clear: blnFiled = False
                        On Error GoTo ErrH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).Value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).Value = rsInput(Mid(arrInput(i + 1), 2)).Value
                    End If
                End If
                blnFiled = False
                Call rsInput.Update
            Next
            .MoveNext
        Loop
    End With
    UpdateRec = True
    Exit Function
ErrH:

End Function
