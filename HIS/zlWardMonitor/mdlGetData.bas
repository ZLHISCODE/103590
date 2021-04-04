Attribute VB_Name = "mdlGetData"
Option Explicit
Private Declare Function CEC_DevNo2His Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strInPatient As String) As Boolean
'lngType:1�໤�Ǵ���, 2HIS����, 3�������

Private Declare Function CEC_UpdateDataBase Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal strResult As String) As Boolean
Private Declare Function CEC_HisSetDataToCec Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal strResult As String) As Boolean
Private Declare Function CEC_GetMonitorData Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strResult As String) As Boolean

Public gcnOracle As New ADODB.Connection    '�������ݿ�����
Private mobjRichEPR As Object           '�������Ĳ���
Const SP = "[|]"
Const SPN = "[^]"


Public Function RequestData(ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal obj As Object) As Boolean
'���ܣ����ݼ໤���ϵĲ���ָ��������������
'������lngDevice-�豸�ţ�lngCmd-���������
    Dim strCmd As String, strdata As String, strResult As String * 20
    Dim strBedNO As String, strInPatient As String, strTmp As String
        
    On Error Resume Next
    strCmd = Hex(lngCmd)
        
    Select Case strCmd
        '��������Ϣ
        Case "F0001"
            Call CEC_GetMonitorData(lngDevice, 6, strResult)
            strTmp = Replace(Replace(Trim(strResult), "{", ""), "}", "")
            strBedNO = Split(strTmp, "|")(0)
            strInPatient = Split(strTmp, "|")(1)
                
            strdata = GetPatientInfor(strInPatient)
            strdata = strBedNO & "|" & strdata
            If strdata <> strBedNO & "|" Then Call CEC_UpdateDataBase(lngDevice, 1, strdata)
            
        '���������Ϣ
        Case "A0001", "A0002", "A0003", "A0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetFee(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
            
        '����ҽ����Ϣ
        Case "B0001", "B0002", "B0003", "B0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
        
            strdata = GetAdvice(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
    
        '��������Ϣ
        Case "C0001", "C0002", "C0003", "C0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetCase(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
        
        '���󱨸���Ϣ
        Case "D0001", "D0002", "D0003", "D0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetReport(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
        
    End Select

    RequestData = True
End Function

Public Function GetPatientInfor(ByVal strInPatient As String) As String
'���ܣ���ȡ������Ϣ
'������strInPatient-סԺ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLimit As Long
    
    strSQL = "select Nvl(Zl_Getsysparameter(147), 12) par from dual"    '��ͯ����綨����
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", strSQL)
    lngLimit = rsTmp!par
     
    strSQL = "Select Nvl(b.��Ժ����,' ')||'|'||Nvl(d.����,' ')||'|'||b.סԺ��||'|'||b.��ҳID||'|'||a.����||'|'||a.����||'|'||a.�Ա�||'| | |'||" & vbNewLine & _
        "to_char(b.��Ժ����,'yyyy-mm-dd')||'|'||to_char(a.��������,'yyyy-mm-dd')||'|'||" & vbNewLine & _
        "Decode(sign(zl_to_number(Substr(a.����, 1, Instr(a.����, '��') - 1))-" & lngLimit & "),1,0,1)||'|'||Nvl(b.Ѫ��,' ')||'|'||Nvl(c.��Ϣֵ,' ')||'|'||Nvl(a.���֤��,' ')||'|'||Nvl(b.��ͥ�绰,' ')||'|'||Nvl(b.��ͥ��ַ,' ') as Data" & vbNewLine & _
        "From ������Ϣ a,������ҳ b,������ҳ�ӱ� c,���ű� d" & vbNewLine & _
        "Where a.����id=b.����id And a.סԺ����=b.��ҳid And b.����id=c.����id(+) And b.��ҳid=c.��ҳid(+) And c.��Ϣ��(+) = '����ҽʦ' And a.��ǰ����id=d.id(+) And a.סԺ�� = [1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
    
    '"�໤�Ǵ���|HIS����|����|������[������]|סԺ����|��������|����|�Ա�|���|����|סԺ����|��������|����|Ѫ��|����ҽ��|���֤��|�绰|סַ"
    
    If rsTmp.RecordCount > 0 Then GetPatientInfor = rsTmp!Data
    
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetFee(ByVal strInPatient As String, ByVal strCmd As String) As String
'���ܣ���ȡ���˷�����Ϣ
'������strInPatient-סԺ��
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    Dim strHead As String, curSum As Currency
    
    strSQL = "Select ����,�Ա�,����,��ǰ���� From ������Ϣ Where סԺ��=[1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "����:" & !���� & SP & "�Ա�:" & !�Ա� & SP & "����:" & Replace(!����, "��", "") & SP & _
                 "סԺ��:" & Val(strInPatient) & SP & "����:" & !��ǰ���� & SPN
    End With
    strHead = strHead & "����ʱ��" & SP & "����" & SP & "������Ŀ" & SP & "����" & SP & "����" & SP & "ʵ�ս��" & SPN
    
    Select Case strCmd
        Case "A0001"
            strIF = "And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60"
        Case "A0002"
            strIF = "And A.����ʱ�� Between Trunc(Sysdate - 1) And Trunc(Sysdate ) - 1 / 24 / 60 / 60"
        Case "A0003"
            strIF = "And A.����ʱ�� Between Trunc(Sysdate-180, 'mm') And Sysdate"
        Case "A0004"
            strIF = ""
    End Select
    
    strSQL = "Select To_Char(A.����ʱ��, 'yyyy/mm/dd') ����ʱ��, C.���� ��������, D.���� �շ���Ŀ," & vbNewLine & _
            "       Decode(Nvl(A.����, 1), 1, '', 0, '', A.���� || ' �� �� ') || A.���� || ' ' || A.���㵥λ As ����, Ltrim(To_Char(A.��׼����,'9999999990.00000')) as ��׼����, Ltrim(To_Char(Nvl(Sum(A.ʵ�ս��),0),'9999999990.00')) as  ʵ�ս��" & vbNewLine & _
            "From ���˷��ü�¼ A, ������Ϣ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where A.����id = B.����id And A.��������id = C.ID And A.�շ�ϸĿid = D.ID And A.��¼״̬ > 0 And B.סԺ�� = [1]" & vbNewLine & _
            "      " & strIF & vbNewLine & _
            "Group By A.NO, Mod(A.��¼����, 10), Nvl(A.�۸񸸺�, A.���), A.��¼״̬, To_Char(A.����ʱ��, 'yyyy/mm/dd'), C.����, D.����, A.����, A.����, A.���㵥λ, A.��׼����" & vbNewLine & _
            "Order By ����ʱ��, A.NO"
    
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
    
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            For i = 1 To .RecordCount
                strValue = strValue & vbNewLine & !����ʱ�� & SP & !�������� & SP & !�շ���Ŀ & SP & !���� & SP & !��׼���� & SP & !ʵ�ս�� & SPN
                curSum = curSum + !ʵ�ս��
                .MoveNext
            Next
        End With
        
        GetFee = strHead & strValue & "�ϼ�" & SP & curSum
    Else
        GetFee = strHead
    End If

    Exit Function
errH:
    Debug.Print Err.Description
    Call WriteLog(Err.Description)
End Function

Private Function GetAdvice(ByVal strInPatient As String, ByVal strCmd As String) As String
'���ܣ���ȡ����ҽ����Ϣ
'������strInPatient-סԺ��
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    Dim strHead As String, strDoctor As String
    
     strSQL = "Select A.����, A.�Ա�, A.����, A.��ǰ����, B.��Ϣֵ As ����ҽʦ" & vbNewLine & _
            "From ������Ϣ A, ������ҳ�ӱ� B" & vbNewLine & _
            "Where A.סԺ�� = [1] And A.����id = B.����id(+) And A.סԺ���� = B.��ҳid(+) And B.��Ϣ��(+) = '����ҽʦ'"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "����:" & !���� & SP & "�Ա�:" & !�Ա� & SP & "����:" & Replace(!����, "��", "") & SP & _
                 "סԺ��:" & Val(strInPatient) & SP & "����:" & !��ǰ���� & SPN
        '������У�Ҫ��ֻ������
        strDoctor = " [|] [^]����ҽ��[|]" & IIf(IsNull(!����ҽʦ), "��", !����ҽʦ)
    End With
    
    strHead = strHead & "��Ч" & SP & "��ʼʱ��" & SP & "ҽ������" & SP & "�÷�" & SP & "Ƶ��" & SPN
    Select Case strCmd
        Case "B0001"
            strIF = " And A.����ʱ�� + 0 Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60"
        Case "B0002"
            strIF = " And A.ҽ����Ч = 0"
        Case "B0003"
            strIF = " And A.ҽ����Ч = 1"
    End Select
    
    strSQL = "Select Decode(A.ҽ����Ч, 0, '����', '����') As ��Ч, To_Char(A.��ʼִ��ʱ��, 'MM-DD HH24:MI') As ��ʼʱ��," & vbNewLine & _
            "       Decode(D.�������, '5', D.ҽ������, '6', D.ҽ������, A.ҽ������) ҽ������," & vbNewLine & _
            "       Decode(A.�������, 'E', Decode(Instr('2468', Nvl(E.��������, '0')), 0, Null, E.����), Null) As �÷�, A.ִ��Ƶ�� As Ƶ��" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ D, ������Ϣ B, ������Ŀ��� C, ������ĿĿ¼ E" & vbNewLine & _
            "Where A.����id = B.����id And A.��ҳid = B.סԺ���� And A.������� = C.����(+) And A.ҽ��״̬ <> -1 And B.סԺ�� = [1] And A.ID = D.���id(+) And" & vbNewLine & _
            "      A.���id Is Null And Instr('5,6', D.�������(+)) > 0 And A.������Ŀid = E.ID(+) And A.ҽ��״̬ Not In (4, 8, 9)" & vbNewLine & strIF

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            For i = 1 To .RecordCount
                strValue = strValue & !��Ч & SP & !��ʼʱ�� & SP & !ҽ������ & SP & !�÷� & SP & !Ƶ�� & SPN
                .MoveNext
            Next
        End With
        GetAdvice = strHead & strValue & strDoctor
    Else
        GetAdvice = strHead & strDoctor
    End If

    Debug.Print strSQL
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetCase(ByVal strInPatient As String, ByVal strCmd As String) As String
'���ܣ���ȡ���˲�����Ϣ
'������strInPatient-סԺ��
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strHead As String, strText As String, i As Long
    Dim objTmp As Object
    
    strSQL = "Select ����,�Ա�,����,��ǰ���� From ������Ϣ Where סԺ��=[1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "����:" & !���� & SP & "�Ա�:" & !�Ա� & SP & "����:" & Replace(!����, "��", "") & SP & _
                 "סԺ��:" & Val(strInPatient) & SP & "����:" & !��ǰ���� & SPN
    End With
    
    Select Case strCmd
        Case "C0001"
            strIF = " And A.�������� = '��Ժ��¼'"
        Case "C0002"
            strIF = " And A.�������� = '�״β��̼�¼'"
        Case "C0003"
            strIF = " And A.�������� = '������¼'"
        Case "C0004"
            strIF = " And (A.�������� = '�����¼' or A.�������� = '������¼')"
    End Select
    
    strSQL = "Select A.ID" & vbNewLine & _
        "From ���Ӳ�����¼ A, ������Ϣ B, �����ļ�Ŀ¼ C" & vbNewLine & _
        "Where A.����id = B.����id And B.סԺ�� = [1] And A.�ļ�id = C.ID And C.���� = 2" & strIF & vbNewLine & _
        "Order by ����ʱ�� Desc"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))
    '�ж�������ļ�ʱ��ֻȡ���һ��
    If rsTmp.RecordCount > 0 Then
        If mobjRichEPR Is Nothing Then
            '����д����ĳ�ʼ���У���Ϊ��Ϣ����ʱ����һ�����̻߳�����޷��������еĶ���
            If mobjRichEPR Is Nothing Then Set mobjRichEPR = CreateObject("zlRichEPR.cRichEPR")
            Call mobjRichEPR.InitRichEPR(gcnOracle, objTmp, 100, True)
        End If
                
        strText = "����" & SP & mobjRichEPR.GetDocumentText(Val(rsTmp!ID))
        GetCase = strHead & strText
    Else
        GetCase = strHead & "����" & SP & "��"
    End If
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetReport(ByVal strInPatient As String, ByVal strCmd As String) As String
'���ܣ���ȡ���˱�����Ϣ
'������strInPatient-סԺ��
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    
    strSQL = ""
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������", Val(strInPatient))

    GetReport = ""
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function




Public Sub WriteLog(ByVal strInfo As String)
    '��������Ϣд���ļ���
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.Log"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, 8) '8-ForAppending
    objText.WriteLine Now()
    objText.WriteLine strInfo
    objText.Close
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
            If intMax = 0 Or intMax < 200 Then intMax = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
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
                If intMax = 0 Or intMax < 200 Then intMax = 200
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
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
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    'Debug.Print strLog
    cmdData.CommandText = strSQL
    Set OpenSQLRecord = cmdData.Execute
End Function
