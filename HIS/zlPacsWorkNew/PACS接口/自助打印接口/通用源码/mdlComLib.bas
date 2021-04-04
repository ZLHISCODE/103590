Attribute VB_Name = "mdlComLib"

Option Explicit

'######################################################################################################################

Private mblnTrans As Boolean

'������־������ر���
Private mlngErrNum As Long
Private mstrErrInfo As String
Private mbytErrType As Byte
Private mstrRecentSQL As String  '���ִ�е�SQL���

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
    Dim arrPar As Variant, I As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '��¼��������
    Dim varOutPar As Variant
    Dim strNode As String
    '���������
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
    For I = 0 To UBound(varParValue)
        varValue = varParValue(I)
        If IsNull(varValue) Then Exit For
        
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adDBTimeStamp, adParamInput, , varValue)
            
        End Select
        strNode = strNode & CStr(varValue) & ","
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = gcnOracle
    End If
    
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    
    cmdPackage.Properties("PLSQLRSet") = False

End Function

Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
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
    Dim strSeq As String, intMax As Integer, I As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            I = CInt(strSeq)
            strPar = strPar & "," & I
            If I > intMax Then intMax = I
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSql
    For I = 1 To intMax
        strSql = Replace(strSql, "[" & I & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(I - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & I & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & I & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & I & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
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
    For I = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(I) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(I) <> lngRight Then lngLeft = 0
            lngRight = arrPar(I)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    cmdData.CommandText = strSql

    Set OpenSQLRecord = cmdData.Execute

End Function

Public Sub ExecuteProcedure(strSql As String, ByVal strFormCaption As String)
    '���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
    '������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
    '˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
    '  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
    '  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
    '  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, I As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSql), 1) = ")" Then
        '���ԭ�в���:��Ȼ�����ظ�ִ��
'        cmdData.CommandText = "" '��Ϊ����ʱ�����������
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        'ִ�еĹ�����
        strTemp = Trim(strSql)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For I = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, I, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, I, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, I, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, I, 1) = "," And Not blnStr And intBra = 0 Then
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
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ��2000���ַ�����ȷ
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
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
                        If datCur = CDate(0) Then datCur = CurrentDate
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
                strPar = strPar & Mid(strTemp, I, 1)
            End If
        Next
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSql
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For I = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
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
    strSql = "Call " & strSql
    If InStr(strSql, "(") = 0 Then strSql = strSql & "()"
    gcnOracle.Execute strSql, , adCmdText

End Sub

Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function Nvl(ByVal varValue As Variant, Optional defaultvalue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), defaultvalue, varValue)
End Function


Public Function VarNvl(ByVal strValue As String, Optional defaultvalue As Variant = "") As Variant
    VarNvl = IIf(strValue = "", Default, strValue)
End Function


Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ��������ϵ�ǰ����ʱ��
    '-------------------------------------------------------------
    
    Dim rsTmp  As ADODB.Recordset
    
    Err = 0
    On Error GoTo errHandle
    
    Set rsTmp = OpenCursor("clsDataBase", "ZLTOOLS", "B_Public.Get_Current_Date")
    
    If rsTmp.RecordCount > 0 Then
        CurrentDate = rsTmp.Fields(0)
    Else
        CurrentDate = 0
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
    CurrentDate = 0
    Err = 0
End Function

Public Function ErrCenter() As Byte
'------------------------------------------------
'���ܣ� �����������������
'������
'���أ� cancel      ���� 0
'       resume      ���� 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    
    bytReturnType = 1
    If gcnOracle.Errors.Count <> 0 Then
        'PL/SQL�洢���̴���(����Ƕ�׵Ĺ��̵���)
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '��־����
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If

'        If gcnOracle.Errors(0).NativeError >= 20000 And gcnOracle.Errors(0).NativeError <= 20200 Then
'            '��־����
'            mbytErrType = 1
'            mlngErrNum = gcnOracle.Errors(0).NativeError
'            mstrErrInfo = gcnOracle.Errors(0).Description
'
'            strNote = gcnOracle.Errors(0).Description
'            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
'            Exit Function
'        End If
        
        'ORACLE��������
        '��־����
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "�Ѿ�������ͬ���ݵ����ݣ�Ҫ��Ψһ������[���š����Ƶ�]���ظ�����"
            bytReturnType = 0
        Case 903
            strNote = "�����ƴ���"
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
        Case 904, 920
            strNote = "�����ƴ���" & vbCrLf & vbCrLf & "SQL�����ʹ���˲����ڵ��л�������."
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
        Case 942
            strNote = "�����ͼ�����ڣ��ܿ������㲻�߱�ʹ�øò������ݵ�Ȩ�ޡ�"
            bytReturnType = 0
            
            strTemp = GetInvalidTable(mstrRecentSQL)
            If strTemp <> "" Then
                mstrErrInfo = "������ж�����м�飺" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
            End If
        Case 1000
            strNote = "�򿪵����ݱ�̫�࣬��Ҫʱ��ϵͳ����Ա�޸����ݿ��Open_Cursors���á�"
        Case 1005
            strNote = "������û��������롣"
        Case 1017
            strNote = "������û��������롣"
            bytReturnType = 0
        Case 1031
            strNote = "û���㹻��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1045
            strNote = "û���������ݿ��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1400
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ����������ʧ�ܡ�"
            bytReturnType = 0
        Case 1401
            strNote = "���ڸ����ֵ�������п����ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1402
            strNote = "���ڸ����ֵ��������ͼ���������ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1403
            strNote = "����δ���������ݣ����º�������ʧ�ܡ�"
        Case 1404
            strNote = "�޸��в�����������ص�����̫��"
        Case 1405
            strNote = "ȡ�õ���ֵΪ�ա�"
        Case 1406
            strNote = "ȡ�õ���ֵ���ж϶������ˡ�"
        Case 1407
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ�����¸���ʧ�ܡ�"
            bytReturnType = 0
        Case 1408
            strNote = "ָ�������Ѿ�������������"
        Case 1409
            strNote = "���ܽ�����˳�����(NoSort)����Ϊ�����û����"
        Case 1410
            strNote = "�������ID(ROWID)����ID���������ֺ��ַ���ɵ�16���Ƹ�ʽ��"
        Case 1411
            strNote = "��ǰ�в��ܴ洢����64K�����ݡ�"
            bytReturnType = 0
        Case 1412
            strNote = "��ǰ���������Ͳ��ܴ洢�㳤���ַ�����"
            bytReturnType = 0
        Case 1413
            strNote = "�����С��λ��������ʧ�ܡ�"
            bytReturnType = 0
        Case 1415
            strNote = "���ܶ�һ����ǩα��ָ��������[Outer-Join(+)]"
        Case 1416
            strNote = "���ű���ͬʱָ��һ��������[Outer-Join(+)]"
        Case 1417
            strNote = "һ�ű�ֻ��ָ��ָ�򲻳���һ�ű��������[Outer-Join(+)]"
        Case 1418
            strNote = "ָ�������������ڡ�"
        Case 1424
            strNote = "�������Ч�Ļ����ַ�(ͨ�����ֻ����'%'��'_')��"
        Case 1425
            strNote = "�����ַ������ǳ���Ϊ1���ַ���"
        Case 1426
            strNote = "��ֵ���ʽ���������(̫���̫С)��"
        Case 1427
            strNote = "�����Ӳ�ѯ�����˶��С�"
        Case 1428
            strNote = "�����Ĳ�������򳬽硣"
        Case 1429
            strNote = "һ�����������ڸ�ʽ���硣"
        Case 1430
            strNote = "ϣ�����ӵ����Ѿ����ڡ�"
        Case 1431
            strNote = "��Ȩ����(GRANT)�������ڵĲ�һ�¡�"
        Case 1432
            strNote = "ϣ��ɾ���Ĺ���ͬ����Ѿ������ڡ�"
        Case 1433
            strNote = "ϣ��������ͬ����Ѿ����ڡ�"
        Case 1434
            strNote = "ϣ��ɾ����ͬ����Ѿ������ڡ�"
        Case 1435
            strNote = "ָ�����û������ڡ�"
            bytReturnType = 0
        Case 1438
            strNote = "��ֵ������������ľ�ȷ�̶ȡ�"
        Case 1439, 1440, 1441
            strNote = "ֻ�п�ֵ�в����޸��������͡������Ȼ�ߴ��С"
        Case 1536
            strNote = "ĳ��������ռ�Ŀռ�������"
        Case 2290
            strNote = "������Ŀֵ��������ķ�Χ��Υ���˼��Լ�������������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 2291
            strNote = "����δ��д��ر��д��ڵ���Ŀֵ(Υ�������Լ��)���������ӻ����ʧ�ܡ�"
        Case 2091, 2292
            strNote = "��Ϊ�ü�¼�Ѿ�ʹ�ã�����ɾ�������ʧ�ܡ�"
            bytReturnType = 0
        Case 2391
            strNote = "�û��Ѵﵽ���ݿ������������¼����"
        Case 12203
            strNote = "������������д�����û���������⣬�����������ӡ�"
            bytReturnType = 0
        Case 20003
            strNote = "�洢������Ч�����ʧЧ�Ĵ洢���̽��б��롣"
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "���ڷ����������ߵĽ�ɫ������������ӶԹ��̡�" & strTemp & "������Ȩ��"
            Else
                strNote = "δ֪���󣬷�����" & gcnOracle.Errors(0).Source
            End If
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
        End Select
        
    Else
        'VB��׼����
        '��־����
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
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
                strNote = "̫���ĸ�������"
                
            'ADO����
            Case -2147483647
                strNote = "δʵ��"
            Case -2147483646
                strNote = "�ڴ治��"
            Case -2147483645
                strNote = "һ������������Ч"
            Case -2147483644
                strNote = "��֧�������Ľӿ�"
            Case -2147483643
                strNote = "��Чָ��"
            Case -2147483642
                strNote = "��Ч���"
            Case -2147483641
                strNote = "������ֹ"
            Case -2147483640
                strNote = "��ȷ���Ĵ���"
            Case -2147483639
                strNote = "һ����ʾܾ�����"
            Case -2147483638
                strNote = "��ɲ�������������ݲ��ٿ���"
            Case -2147467263
                strNote = "δʵ��"
            Case -2147467262
                strNote = "��֧�������Ľӿ�"
            Case -2147467261
                strNote = "��Чָ��"
            Case -2147467260
                strNote = "������ֹ"
            Case -2147467259
                strNote = "��ȷ���Ĵ���"
            Case -2147467258
                strNote = "�̱߳��ش洢ʧ��"
            Case -2147467257
                strNote = "��ȡ������ڴ�������ʧ��"
            Case -2147467256
                strNote = "��ȡ�ڴ�������ʧ��"
            Case -2147467255
                strNote = "���ܳ�ʼ����ĸ��ٻ���"
            Case -2147467254
                strNote = "���ܳ�ʼ��RPC����"
            Case -2147467253
                strNote = "���������̱߳��ش洢ͨ������"
            Case -2147467252
                strNote = "���ܷ����̱߳��ش洢ͨ������"
            Case -2147467251
                strNote = "�û��ṩ���ڴ������򲻿ɽ���"
            Case -2147467250
                strNote = "OLE���񻥳����Ѵ���"
            Case -2147467249
                strNote = "OLE�����ļ�ӳ���Ѵ���"
            Case -2147467247
                strNote = "��ͼ����OLE����ʧ��"
            Case -2147467246
                strNote = "�ڵ��߳�ģ������ͼ��һ�ε���CoInitialize"
            Case -2147467245
                strNote = "��Ҫһ��Զ�̼�����ǲ�����"
            Case -2147467244
                strNote = "��Ҫһ��Զ�̼�������ṩ�ķ�����������Ч"
            Case -2147467243
                strNote = "���������õİ�ȫid������߲�ͬ"
            Case -2147467242
                strNote = "ʹ��OLE1���������DDE���ڱ���ֹ"
            Case -2147467241
                strNote = "RunAsָ���ı���������\�û�����ֻ���û���"
            Case -2147467240
                strNote = "������̲�������������·��������ȷ"
            Case -2147467239
                strNote = "�����ñ�ʶʱ������̲���������·�������ܲ���ȷ����Ч"
            Case -2147467238
                strNote = "�������ñ�ʶ����ȷ��������̲�������������û����Ϳ���"
            Case -2147467237
                strNote = "������ͻ��������������"
            Case -2147467236
                strNote = "�ṩ�������ķ�������������"
            Case -2147467235
                strNote = "����������ܺͷ������ṩ�����������ͨ��"
            Case -2147467234
                strNote = "��������������Ӧ"
            Case -2147467233
                strNote = "��������ע����Ϣ��һ�»�����"
            Case -2147467232
                strNote = "����ӿڵ�ע����Ϣ��һ�»�����"
            Case -2147467231
                strNote = "��֧����ͼִ�еĲ���"
            Case -2147418113
                strNote = "������ʧ��"
            Case -2147024891
                strNote = "һ����ʾܾ�����"
            Case -2147024890
                strNote = "��Ч���"
            Case -2147024882
                strNote = "�ڴ治��"
            Case -2147024809
                strNote = "һ������������Ч"
            Case 3000
                strNote = "�ṩ��ִ������Ķ���ʧ��"
            Case 3001
                strNote = "�������ʹ��󣬻���ֵ������Χ�������������ͻ����ͻ��"
            Case 3002
                strNote = "����������ļ�ʱ����������"
            Case 3003
                strNote = "��ָ�����ļ�ʱ����"
            Case 3004
                strNote = "д�ļ�ʱ�д���"
            Case 3021
                strNote = "BOF��EOF��һ��ΪTrue�����ߵ�ǰ��¼�ѱ�ɾ����Ӧ�ó�������������Ҫ��ǰ��¼"
            Case 3219
                strNote = "�����Ļ���������ǰӦ�ò����������Ǵ�����δ���������񣩡�"
            Case 3220
                strNote = "���ܸı��ṩ��"
            Case 3246
                strNote = "������ִ���У����ܹر�һ���������"
            Case 3251
                strNote = "�ṩ�߲�֧�ָ�Ӧ�ó�������Ĳ�����"
            Case 3265
                strNote = "ADOû�ҵ�Ӧ�ó���Ҫ��Ķ�Ӧ���ƻ���ţ������������ƴ��󣩡�"
            Case 3367
                strNote = "�������ڼ����У�����׷��"
            Case 3420
                strNote = "����δ���û����õĶ�������Ч��"
            Case 3421
                strNote = "��ǰ����ʹ���˴������ֵ���͡�"
            Case 3704
                strNote = "��������ѹرգ�������Ӧ�ó�������Ĳ���"
            Case 3705
                strNote = "��������Ѵ򿪣�������Ӧ�ó�������Ĳ���"
            Case 3706
                strNote = "ADO�����ҵ�ָ�����ṩ��"
            Case 3707
                strNote = "���ܲ����������ı�һ����¼���Ļ����Դ�����ԡ�"
            Case 3708
                strNote = "Ӧ�ó�����ִ���Ĳ������塣"
            Case 3709
                strNote = "Ӧ�ó��������һ������Ĳ���ʱʹ����һ�����ã���������ָ����һ���رյĻ���Ч��Connection����"
            Case 3710
                strNote = "������������ִ��"
            Case 3711
                strNote = "������Ȼ��ִ��"
            Case 3712
                strNote = "������ȡ��"
            Case 3713
                strNote = "������Ȼ��������"
            Case 3714
                strNote = "������Ч"
            Case 3715
                strNote = "��������ִ�й�����"
            Case 3716
                strNote = "��������������в���ȫ"
            Case 3717
                strNote = "��������һ����ȫ�Ի�"
            Case 3718
                strNote = "��������һ����ȫ�Ի�ͷ"
            Case 3719
                strNote = "Υ�����ݵ������ԣ�����ʧ�ܡ�"
            Case 3720
                strNote = "�û�û���㹻��Ȩ����ɲ���������ʧ�ܡ�"
            Case 3721
                strNote = "���ݳ����������������͵ķ�Χ"
            Case 3722
                strNote = "����Υ����ģʽ"
            Case 3723
                strNote = "���ʽ������ƥ��ķ���"
            Case 3724
                strNote = "����ת��ֵ���ܴ�����Դ"
            Case 3726
                strNote = "��һ���в�����ָ������"
            Case 3727
                strNote = "URL������"
            Case 3728
                strNote = "û�в鿴Ŀ¼����Ȩ��"
            Case 3729
                strNote = "�ṩ��URL��Ч"
            Case 3730
                strNote = "��Դ������"
            Case 3731
                strNote = "��Դ�Ѿ�����"
            Case 3732
                strNote = "������ɶ���"
            Case 3733
                strNote = "�ļ��汾��Ϣû�ҵ�"
            Case 3734
                strNote = "�������ò����㹻�Ŀռ���ɲ���������ʧ��"
            Case 3735
                strNote = "��Դ������Χ"
            Case 3736
                strNote = "�������"
            Case 3737
                strNote = "�����������е�URL������"
            Case 3738
                strNote = "����ɾ����Դ���ⳬ��������Χ"
            Case 3739
                strNote = "����ѡ����У����������Ч"
            Case 3740
                strNote = "�������ṩ��һ����Ч��ѡ��"
            Case 3741
                strNote = "�������ṩ��һ����Ч��ֵ"
            Case 3742
                strNote = "�������������ɺ��������Գ�ͻ"
            Case 3743
                strNote = "�������е����Զ��ܱ�����"
            Case 3744
                strNote = "����û�б�����"
            Case 3745
                strNote = "���Բ��ܱ�����"
            Case 3746
                strNote = "���Բ���֧��"
            Case 3747
                strNote = "���û���������Զ�������ִ��"
            Case 3748
                strNote = "���ܸı�����"
            Case 3749
                strNote = "Fields���ϵ�Update����ʧ��"
            Case 3750
                strNote = "��������DenyȨ�ޣ���Ϊ�ṩ�߲�֧��"
            Case 3751
                strNote = "�ṩ�߲�֧�������Deny����"
                
            Case Else
                strNote = "����δ֪�Ľ������"
        End Select
        bytReturnType = 0
    End If

    If bytReturnType = 1 Then
        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
    Else
        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
        ErrCenter = 0
    End If
    
    '�������
    Err.Clear
End Function

Public Function GetInvalidTable(ByVal strRecentSQL As String) As String
'���ܣ��õ������ʹ�õ�SQL����в��ܷ��ʵı����ͼ
'
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(strRecentSQL), ",")
    
    On Error Resume Next
    
    For lngCount = 0 To UBound(varTables)
        strTable = varTables(lngCount)
        
        '���Ըö����Ƿ����
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        'ȥ����һ������
        GetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function

Public Function SQLObject(ByVal strSql As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim I As Integer, J As Integer
    
    On Error GoTo errH
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSql))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
        intL = 1: intR = 0
        For I = intB + 1 To Len(strAnal)
            If Mid(strAnal, I, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, I, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = I
                If intE - intB - 1 <= 0 Then
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����
    arrFrom = Split(strAnal, "FROM")
    For I = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
        strCur = arrFrom(I)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        Else
            strMulti = strCur
        End If
        For J = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(J))
            If InStr(strObject, "," & strTrue) = 0 And strTrue <> "Ƕ�ײ�ѯ" Then
                strObject = strObject & "," & strTrue
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Private Function TrimChar(Str As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim I As Long, J As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    I = InStr(strTmp, "  ")
    Do While I > 0
        strTmp = Left(strTmp, I) & Mid(strTmp, I + 2)
        I = InStr(strTmp, "  ")
    Loop
    
    I = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While I > 0
        strTmp = Left(strTmp, I + 1) & Mid(strTmp, I + 4)
        I = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function TrueObject(ByVal strObject As String) As String
'���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim I As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For I = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, I, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, I)
    'Ѱ�Һ����һ���������ַ�
    For I = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, I, 1)) > 0 Then Exit For
    Next
    If I <= Len(strObject) Then strObject = Left(strObject, I - 1)
    TrueObject = strObject
End Function

