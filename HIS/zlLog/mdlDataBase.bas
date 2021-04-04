Attribute VB_Name = "mdlDataBase"
'@ģ�� mdlDataBase-2019/9/17
'@��д lshuo
'@����
'   ���ݿ��ʵ��
'@����
'
'@��ע
'
Option Explicit

Public Function CallProcedure(cnInput As ADODB.Connection, ByVal strProcName As String, ByVal strFormCaption As String _
    , blnUseLog As Boolean, ParamArray arrProcParas() As Variant) As Variant
    
    Dim arrPars()   As Variant
    Dim arrRet      As Variant
    
    arrPars = arrProcParas
    arrRet = mdlDataBase.CallProcedureByArray(cnInput, strProcName, strFormCaption, blnUseLog, arrPars, True)
    If UBound(arrRet) = 0 Then
        If IsObject(arrRet(0)) Then
            Set CallProcedure = arrRet(0)
        Else
            CallProcedure = arrRet(0)
        End If
    Else
        CallProcedure = arrRet
    End If
End Function

Public Function CallProcedureByArray(cnInput As ADODB.Connection, ByVal strProcName As String _
    , ByVal strFormCaption As String, blnUseLog As Boolean, arrProcParas() As Variant _
    , Optional ByVal blnAsSubCall As Boolean) As Variant
    
    Dim cmdData As New ADODB.Command
    Dim rsReturn As New ADODB.Recordset
    Dim i As Long, lngAdjust As Long, dtSize As Long, lngMax As Long, lngParaUbound As Long
    Dim pdCur As ParameterDirectionEnum, dtCur As DataTypeEnum
    Dim varValue As Variant, arrRet() As Variant
    Dim arrIntOut() As Integer
    Dim blnOELDB As Boolean
    
    If blnAsSubCall Then
        CallProcedureByArray = Array(False)
    Else
        CallProcedureByArray = False
    End If
    Const MAX_STRING_SIZE = 32767
    
    ReDim Preserve arrIntOut(0)
    For i = LBound(arrProcParas) To UBound(arrProcParas)
        '1�����������Ķ��峤�ȡ����͡������Լ�����ֵ
        pdCur = adParamUnknown: dtCur = adEmpty: dtSize = -1: varValue = Empty: lngMax = -1
        '���Σ�Empty
        If IsEmpty(arrProcParas(i)) Then
            pdCur = adParamOutput: dtCur = adLongVarChar
        ElseIf IsArray(arrProcParas(i)) Then
            '���Σ�Array(Empty[,�ֶ�����,�ֶγ���])
            '����ֵ��Array(Empty,Empty[,�ֶ�����,�ֶγ���])
            varValue = arrProcParas(i)(0)
            pdCur = adParamOutput
            lngParaUbound = UBound(arrProcParas(i))
            '����ֵ��Array(Empty,Empty[,�ֶ�����,�ֶγ���])
            If lngParaUbound > 0 Then
                If IsEmpty(arrProcParas(i)(0)) And IsEmpty(arrProcParas(i)(1)) Then
                    pdCur = adParamReturnValue
                    If i <> 0 Then
                        Err.Raise vbObjectError + 2, strFormCaption, "�����ķ���ֵ(λ��0)�����ں�������֮ǰ���ݣ���ǰ����ֵλ�ã�" & i
                    End If
                    If lngParaUbound > 1 Then
                        dtCur = arrProcParas(i)(2)
                        If lngParaUbound > 2 Then
                            dtSize = arrProcParas(i)(3)
                        End If
                    End If
                End If
            End If
            '���Σ�Array(Empty[,�ֶ�����,�ֶγ���])
            '����Σ�Array(ֵ[,�ֶ�����,�ֶγ���])
            If pdCur <> adParamReturnValue Then
                If Not IsEmpty(arrProcParas(i)(0)) Then
                    pdCur = adParamInputOutput
                End If
                If lngParaUbound > 0 Then
                    dtCur = arrProcParas(i)(1)
                    If lngParaUbound > 1 Then
                        dtSize = arrProcParas(i)(2)
                    End If
                End If
            End If
        Else
            varValue = arrProcParas(i)
            pdCur = adParamInput
        End If
        '2\�ռ�����λ�ã��Է�������ռ�����ֵ��
        If pdCur > adParamInput Then
            ReDim Preserve arrIntOut(UBound(arrIntOut) + 1)
            arrIntOut(UBound(arrIntOut)) = i
        End If
        '3�����Ӳ���
        Select Case VarType(varValue)
            Case vbString
                lngMax = dtSize
                If dtSize = -1 Then         'δ���峤�ȣ����ȡԭʼ����
                    lngMax = LenB(varValue) 'LenB(StrConv(varValue, vbFromUnicode)) '�����ַ�������ת����ʱ
                Else
                    lngMax = arrProcParas(i)(1)
                    If lngMax < LenB(varValue) Then     '���峤��С��ʵ�ʳ��ȣ�����Ϊʵ�ʳ���
                        lngMax = LenB(varValue)  'LenB(StrConv(varValue, vbFromUnicode)) '�����ַ�������ת����ʱ
                    End If
                End If
                'ȡOLEDB���ַ�����
                If lngMax <= 4000 Then
                    If lngMax < 32 Then
                        lngMax = 32
                    ElseIf lngMax < 128 Then
                        lngMax = 128
                    ElseIf lngMax < 2000 Then
                        lngMax = 2000
                    Else
                        lngMax = 4000
                    End If
                ElseIf dtSize = -1 And pdCur > adParamInput Then 'ԭʼ���ȵĳ���>4000 ,��û�ж��峤�ȣ�ͳһʹ����󳤶�
                    lngMax = MAX_STRING_SIZE
                End If
                
                 If lngMax <= 2000 Then
                    If dtCur = adEmpty Then dtCur = adVarChar   'С��2000��û��ָ�����ͣ�ʹ��adVarChar
                Else
                    If dtCur = adEmpty Then dtCur = adLongVarChar   '����2000��û��ָ�����ͣ�ʹ��adLongVarChar
                End If
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, lngMax, varValue)
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal '����
                If dtCur = adEmpty Then dtCur = adVarNumeric
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, 38, varValue) '��ǰ30�޸�Ϊ38
            Case vbDate
                If dtCur = adEmpty Then dtCur = adDBTimeStamp
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, , varValue)
            Case vbNull, vbEmpty    '����ֵ��NULL����EMPTY��NULL��˵��������NULL,EMPTY��˵�������ǳ��Ρ�
                If dtCur = adEmpty Then
                    If dtSize = -1 Then dtSize = MAX_STRING_SIZE    'û�ж��峤�ȣ�ͳһʹ����󳤶�
                    If dtSize <= 2000 Then
                        dtCur = adVarChar   'С��2000��û��ָ�����ͣ�ʹ��adVarChar
                    Else
                        dtCur = adLongVarChar
                    End If
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize, varValue)
                Else
                    Select Case dtCur
                        Case adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adBSTR, adVarBinary, adLongVarBinary
                            If dtSize = -1 Then dtSize = MAX_STRING_SIZE    'û�ж��峤�ȣ�ͳһʹ����󳤶�
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                            If dtSize = -1 Then dtSize = 38
                    End Select
                    If VarType(varValue) <> vbNull Then
                        If dtSize = -1 Then
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, , varValue)
                        Else
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize, varValue)
                        End If
                    Else
                        If dtSize = -1 Then
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur)
                        Else
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize)
                        End If
                    End If
                End If
            Case Else
                Err.Raise vbObjectError + 1, strFormCaption, "�洢���̴��ݵĲ��������޷�ʶ�𣬲���λ�ã�" & i
        End Select
    Next
    Set cmdData.ActiveConnection = cnInput   '���Ƚ���
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = strProcName
    blnOELDB = IsOLEDBConnection(cnInput)
    If blnOELDB Then
        cmdData.Properties("PLSQLRSet") = True
    End If
    
    'Set rsReturn = cmdData.Execute
    Set rsReturn = mdlDataBase.CommandExecuteStoredProc(cmdData)
    
    If blnOELDB Then
        cmdData.Properties("PLSQLRSet") = False
    End If
    If rsReturn.State = adStateClosed Then
        arrIntOut(0) = -1       '����޷����α�
    End If
    If UBound(arrIntOut) > 0 Or arrIntOut(0) <> -1 Then
        'ֻ��1����ͨ�����Լ���ͨ����ֵ
        If UBound(arrIntOut) = 1 And arrIntOut(0) = -1 Then
            If blnAsSubCall Then
                CallProcedureByArray = Array(cmdData.Parameters(arrIntOut(1)).Value & "")
            Else
                CallProcedureByArray = cmdData.Parameters(arrIntOut(1)).Value & ""
            End If
        'ֻ�������α�
        ElseIf UBound(arrIntOut) = 0 Then
            If blnAsSubCall Then
                CallProcedureByArray = Array(rsReturn)
            Else
                Set CallProcedureByArray = rsReturn
            End If
        Else
            '�����α꣬�ܹ�������������1
            If arrIntOut(0) = -1 Then
                ReDim Preserve arrRet(UBound(arrIntOut) - 1)
                lngAdjust = 1
            Else
                ReDim Preserve arrRet(UBound(arrIntOut))
                Set arrRet(0) = rsReturn
                lngAdjust = 0
            End If
            For i = 1 To UBound(arrIntOut)
                arrRet(i - lngAdjust) = cmdData.Parameters(arrIntOut(i)).Value & ""
            Next
            CallProcedureByArray = arrRet
        End If
    Else
        '���κη�����Ϣ
        If blnAsSubCall Then
            CallProcedureByArray = Array(True)
        Else
            CallProcedureByArray = True
        End If
    End If
End Function

Public Function OpenSQLRecord(cnInput As ADODB.Connection, ByVal strSql As String, ByVal strTitle As String _
    , ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = mdlDataBase.OpenSQLRecordByArray(cnInput, strSql, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(cnInput As ADODB.Connection, ByVal strSql As String, ByVal strTitle As String _
    , arrInput() As Variant, Optional intLobOprate As Integer = 0) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      intLobOprate=0:��ͨSQL,1:LOB���Ͷ�ȡSQL,2:LOB����SQL
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim strError As String
    Dim lngPos     As Long
    Dim cnOLEDB     As ADODB.Connection
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrStr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrStr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrStr) Then
            If Not Replace(strSQLTmp, " ", "") Like "*/[*]+CARDINALITY*[*]/*" Then '�����ж��CARDINALITY���磺/*+cardinality(c,10) cardinality(d,10)*/
                strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
            End If
        End If
    End If
    
'    If Replace(strSQLTmp, " ", "") Like "*/[*]+DRIVING_SITE*[*]/*" Then
'        If Not CheckDatamoveRemote Then
'            arrStr = Split(strSql, "/*")
'            strSql = arrStr(LBound(arrStr))
'            For i = LBound(arrStr) To UBound(arrStr)
'                If i <> UBound(arrStr) Then
'                    lngPos = InStr(arrStr(i + 1), "*/")
'                    lngLeft = 0
'                    If lngPos <> 0 Then
'                        lngLeft = InStr(1, arrStr(i + 1), "DRIVING_SITE", vbTextCompare)
'                        If lngLeft > 0 Then
'                            If Trim(Mid(arrStr(i + 1), 1, lngLeft - 1)) <> "+" Then
'                                lngLeft = 0
'                            End If
'                        End If
'                    End If
'
'                    If lngLeft > 0 And lngLeft < lngPos Then
'                        strSql = strSql & Mid(arrStr(i + 1), lngPos + 2)
'                    Else
'                        strSql = strSql & "/*" & arrStr(i + 1)
'                    End If
'                End If
'            Next
'        End If
'    End If
    
    Call AdjustSQL(strSql)
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        If lngRight = 0 Then Exit Do
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

'    If intLobOprate = 0 Then
        Set cmdData.ActiveConnection = cnInput '���Ƚ���(���ִ��1000��Լ0.5x��)
'    Else
'        Set cnOLEDB = gcnOracleOLEDB
'        If cnOLEDB Is Nothing Then
'            If Not IsOLEDBConnection(cnInput) Then
'                Set cnOLEDB = gobjRegister.ReGetConnection(Val("1-OraOLEDB"), strError, cnInput)
'            Else
'                Set cnOLEDB = cnInput
'            End If
'            If cnInput Is gcnOracle Then Set gcnOracleOLEDB = cnOLEDB
'        End If
'        Set cmdData.ActiveConnection = cnOLEDB
'    End If

    cmdData.CommandText = strSql
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
'    If intLobOprate > 0 Then '����LOB,��ȡLOBҲҪʹ�øò��������������Լ10�����
'        Set OpenSQLRecordByArray = New ADODB.Recordset
'        OpenSQLRecordByArray.Open cmdData, , adOpenStatic, adLockOptimistic
'    Else
        Set OpenSQLRecordByArray = cmdData.Execute
        On Error Resume Next
        Set OpenSQLRecordByArray.ActiveConnection = Nothing
        On Error GoTo 0
'    End If
'    Call gobjComLib.SQLTest
End Function

Private Sub AdjustSQL(ByRef strSQLIn As String)
'���ܣ�����SQL��д��ʽ������ADOִ���쳣

    Const STR_SELECT As String = "select ", STR_FROM As String = "from"
    
    Dim i As Long, lngPos As Long
    Dim intLen As Integer
    Dim strSql As String
    Dim blnDo As Boolean
    
    intLen = Len(STR_SELECT)
    lngPos = 1
    
    On Error GoTo hErr
    
    '1.����*/'�������ִ�
    If strSQLIn Like "*/[*]*+*[*]/'*" Then
        strSql = strSQLIn
        Do While True
            lngPos = InStr(lngPos, LCase$(strSql), STR_SELECT)
            If lngPos > 0 Then
                If Trim(Mid$(strSql, lngPos + intLen, 20)) Like "/[*]*+*" Then
                    i = InStr(Mid$(strSql, lngPos + intLen), "/*+")
                    If i <= 0 Then i = InStr(Mid$(strSql, lngPos + intLen), "/* ")
                    If i > 0 Then
                        '���ڡ� */'�������ִ�
                        For i = lngPos + intLen + 3 To Len(strSql)
                            If Mid$(strSql, i, 3) = "*/'" Then
                                strSql = Left$(strSql, i - 1) & "*/ '" & Mid$(strSql, i + 3)
                            ElseIf LCase(Mid$(strSql, i, Len(STR_FROM))) = STR_FROM Then
                                Exit For
                            End If
                        Next
                        lngPos = i
                    End If
                Else
                    lngPos = lngPos + intLen
                End If
            Else
                Exit Do
            End If
        Loop
        blnDo = True
    End If
    
    '2.����
    '...
    
    If blnDo Then strSQLIn = strSql
    Exit Sub
    
hErr:
    '...
End Sub

Public Function CommandExecuteStoredProc(ByRef cmdVar As ADODB.Command) As ADODB.Recordset
    If cmdVar Is Nothing Then Exit Function
    
    On Error GoTo hErr
    Set CommandExecuteStoredProc = cmdVar.Execute
    
    If Not cmdVar.ActiveConnection.Errors Is Nothing Then
        If cmdVar.ActiveConnection.Errors.Count > 0 Then
            If VBA.Err.Number = 0 _
                And (cmdVar.ActiveConnection.Errors(0).Number = CLng(&H40EC9) _
                    Or UCase(cmdVar.ActiveConnection.Errors(0).Description) Like "*ORAOLEDB*40EC9*") Then
                '��������������Ӷ����Errors����ֹ�������뱨����󽫸��쳣�׳�
                cmdVar.ActiveConnection.Errors.Clear
            End If
        End If
    End If
    Exit Function
    
hErr:
    Err.Raise Err.Number, , Err.Description
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
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
    FromatSQL = strText
End Function

'Public Function CheckDatamoveRemote(Optional ByVal lngSys As Long = 100) As Boolean
''���ܣ����ϵͳ����ʷ���Ƿ���DBLinK
'    Dim rsTmp As ADODB.Recordset, strSql As String
'
'    On Error GoTo ErrH
'    If lngSys <> 100 Or mbytCheckDatamoveRemote = 0 Then
'        strSql = "Select 1 From zlBakSpaces Where ϵͳ = [1] And ��ǰ = 1 And Db���� Is Not Null"
'        Set rsTmp = OpenSQLRecord(strSql, "CheckDatamoveRemote", lngSys)
'        CheckDatamoveRemote = rsTmp.RecordCount > 0
'        If CheckDatamoveRemote Then
'            mbytCheckDatamoveRemote = 1
'        Else
'            mbytCheckDatamoveRemote = 2
'        End If
'    Else
'        CheckDatamoveRemote = mbytCheckDatamoveRemote = 1
'    End If
'    Exit Function
'
'ErrH:
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'End Function
