Attribute VB_Name = "mdlThirdSwap"
Option Explicit

Public Function ZlGetForceDelToCashNote(ByRef cllForceDelToCash As Collection) As String
    '��ȡǿ������ժҪ������"����˵��"�Զ��У���ʽ��XXXXǿ������:XXX��;XXX��
    '��Σ�
    '   cllForceDelToCash Array(����Ա,���������)
    Dim str����Ա As String
    Dim strTemp As String, i As Integer
    
    On Error GoTo ErrHandler
    If cllForceDelToCash Is Nothing Then Exit Function
    If cllForceDelToCash.Count = 0 Then Exit Function
    
    str����Ա = cllForceDelToCash(1)(0)
    For i = 1 To cllForceDelToCash.Count
        strTemp = strTemp & ";" & cllForceDelToCash(i)(1)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    ZlGetForceDelToCashNote = str����Ա & "ǿ�����֣�" & strTemp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlThreeBalanceCheck(frmMain As Form, ByVal lngModule As Long, _
    ByVal objCard As Card, ByRef cllForceDelToCash As Collection, _
    ByVal str��������� As String, ByVal bln�������� As Boolean, _
    Optional ByRef blnǿ������ As Boolean, _
    Optional ByVal blnȱʡ���� As Boolean) As Boolean
    '������ǿ�����ּ��
    '��Σ�
    '   objCard ҽ�ƿ���Ϣ
    '   str��������� ���������
    '���Σ�
    '   cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������)
    '���أ�����ǿ�����֣�����True�����򣬷���False
    '105432
    Dim str����Ա As String
    
    On Error GoTo ErrHandler
    blnǿ������ = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    
    If objCard Is Nothing Then
        If bln�������� = False And blnȱʡ���� = False Then
            ShowMsgbox "δ�ҵ���" & str��������� & "����" & _
                "�޷��ж����Ƿ�֧�����֣�����ǿ����Ϊ�������㷽ʽ��"
            Exit Function
        Else
            If MsgBox("δ�ҵ���" & str��������� & "�����޷��ж����Ƿ�֧�����֣�" & _
                "��ȷ��Ҫǿ����Ϊ�������㷽ʽ��", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Else
        If Not (objCard.�ӿ���� > 0 And Not objCard.���ѿ�) Then
            ZlThreeBalanceCheck = True: Exit Function
        End If
        If bln�������� Then ZlThreeBalanceCheck = True: Exit Function
        If blnȱʡ���� Then '���������֣�ͬʱȱʡ�����֣�������ǿ������
            ZlThreeBalanceCheck = True: Exit Function
        Else
            ShowMsgbox "��" & str��������� & "��������ǿ����Ϊ�������㷽ʽ��"
            Exit Function
        End If
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
        If MsgBox("��" & str��������� & "����֧�����֣���ȷ��Ҫ����ǿ��������", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        cllForceDelToCash.Add Array(UserInfo.����, str���������)
    Else
        str����Ա = zlDatabase.UserIdentifyByUser(frmMain, _
            "��" & str��������� & "��ǿ�����֣�Ȩ����֤��", _
            glngSys, lngModule, "�����˿�ǿ������", , True)
        If str����Ա = "" Then Exit Function
        cllForceDelToCash.Add Array(str����Ա, str���������)
    End If
    blnǿ������ = True
    ZlThreeBalanceCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeDetailXML(ByVal lng����ID As Long) As String
    '��ȡ�����������˷ѽӿ�zlRetuenCheck�з����б�
    '��Σ�
    '   lng����ID - ����ID
    '���أ�
    '      <TFLIST> //�˷��б�
    '        <NO></NO> // �˷ѵ���
    '        <TFITEM> //�˷���
    '          <SerialNum></SerialNum> //���
    '          ��
    '        </TFITEM>
    '      </TFLIST>
    '      ...
    Dim strPriorNO As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim strXML As String, strXMLSub As String
    
    On Error GoTo ErrHandler
    
    strSQL = _
        "Select a.NO, a.���, a.ʵ�ս��" & vbNewLine & _
        "From ������ü�¼ A" & vbNewLine & _
        "Where a.����id = [1]" & vbNewLine & _
        "Order By a.NO, a.���"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlPublicThreeSwap", lng����ID)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    strXML = "": strPriorNO = ""
    Do While Not rsRecord.EOF
        If strPriorNO <> Nvl(rsRecord!NO) Then
            If strPriorNO <> "" Then
                strXML = strXML & "    </TFITEM>" & vbCrLf
                strXML = strXML & "  </TFLIST>" & vbCrLf
            End If
            strXML = strXML & "  <TFLIST>" & vbNewLine '�˷��б�
            strXML = strXML & "    <NO>" & Nvl(rsRecord!NO) & "</NO>" & vbCrLf '�˷ѵ���
            strXML = strXML & "    <TFITEM>" & vbCrLf '�˷���
        End If
        
        strXML = strXML & "      <SerialNum>" & Val(Nvl(rsRecord!���)) & "</SerialNum>" & vbCrLf '���
        strPriorNO = Nvl(rsRecord!NO)
        
        rsRecord.MoveNext
    Loop
    
    strXML = strXML & "    </TFITEM>" & vbCrLf
    strXML = strXML & "  </TFLIST>" & vbCrLf
    
    ZlGetDelFeeDetailXML = strXML
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlMakeDelFeeRecord() As ADODB.Recordset
    '�����˷���ϸ
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "���", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set ZlMakeDelFeeRecord = rsTmp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeRecord(ByVal lng����ID As Long) As ADODB.Recordset
    '��ȡ�����˷���Ŀ
    Dim i As Integer
    Dim rsDelFeeRecord As ADODB.Recordset
    Dim strSQL As String, rsRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    
    strSQL = _
        "Select a.NO, a.���, a.ʵ�ս��" & vbNewLine & _
        "From ������ü�¼ A" & vbNewLine & _
        "Where a.����id = [1]" & vbNewLine & _
        "Order By a.NO, a.���"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlPublicThreeSwap", lng����ID)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    Do While Not rsRecord.EOF
        With rsDelFeeRecord
            .AddNew
            !NO = Nvl(rsRecord!NO)
            !��� = Nvl(rsRecord!���)
            !ʵ�ս�� = Nvl(rsRecord!ʵ�ս��)
            .Update
        End With
        
        rsRecord.MoveNext
    Loop
    
    Set ZlGetDelFeeRecord = rsDelFeeRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetBalanceXML(rsBalance As ADODB.Recordset, _
    rsBalanceByNo As ADODB.Recordset, _
    ByVal lngԭ����ID As Long, ByVal lng�����ID As Long, _
    Optional ByRef dblMoneyTotal As Double, _
    Optional ByVal blnȫ�� As Boolean, _
    Optional ByVal rsDelFeeRecord As ADODB.Recordset, _
    Optional ByRef lng��������ID As Long, _
    Optional ByVal dblDelMoney As Double) As String
    '��ȡԭʼ������Ϣ
    '��Σ�
    '   rsBalance - �������ݣ������ڲ���Ԥ����¼�������쳣״̬������
    '   rsBalanceByNo - �ֵ��ݽ������ݣ�������ҽ��������ϸ�������쳣״̬������
    '   lngԭ����ID - ԭʼ����ID
    '   lng�����ID - ҽ�ƿ����ID
    '   blnȫ�� - �Ƿ�ȫ��
    '   rsDelFeeRecord - �����˿���Ϣ��NO,���,ʵ�ս���ԭ����ʱ������
    '   dblDelMoney - ����ʵ��¼���˿����ԭ����ʱ������
    '���Σ�
    '   lng��������ID - �����˿�ԭ��������ID
    '   dblMoneyTotal - �����˿�ϼ�
    '���أ�
    '  <TKLIST>//�˿��б�35.90��ǰ�޴����ݣ�
    '    <TK>
    '      <TKFS>�˿ʽ</TKFS>
    '      <TKJE>�˿���</TKJE>
    '      <JYLSH>ԭ������ˮ��</JYLSH>
    '      <JYSM><ԭ����˵��</JYSM>
    '      <DJH>���ݺ�</DJH>
    '    </TK>
    '    ��
    '  </TKLIST>
    Dim strXML As String, dblCurMoney As Double
    Dim bln�ֵ��� As Boolean, dblMoney As Double
    Dim i As Integer, j As Integer
    Dim cllNo As Collection, cllBalance As Collection, strKey As String
    Dim rsDelFeeRecordByNo As ADODB.Recordset
    Dim blnFind As Boolean
    
    On Error GoTo ErrHandler
    If rsBalance Is Nothing Then Exit Function
    
    dblMoneyTotal = 0
    If blnȫ�� Then
        If Not rsDelFeeRecord Is Nothing Then Set rsDelFeeRecordByNo = rsDelFeeRecord
        Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    Else
        If rsDelFeeRecord Is Nothing Then Exit Function
    End If
    
    rsBalance.Filter = "����ID=" & lngԭ����ID & " And �����ID=" & lng�����ID
    If rsBalance.RecordCount = 0 Then
        '1.����ֻ��ҽ������������
        '2.��ҽ������ֽ��㷽ʽ������ʱ�����������
        rsBalance.Filter = "����=" & Enum_BalanceType.һ��ͨ & " And �����ID=" & lng�����ID & " And �˷�=0"
        If rsBalance.RecordCount = 0 Then Exit Function
        lngԭ����ID = Val(Nvl(rsBalance!����ID))
    End If
    
    lng��������ID = Val(Nvl(rsBalance!��������ID))
    
    If Not rsBalanceByNo Is Nothing Then
        rsBalanceByNo.Filter = "��������ID=" & lng��������ID & " And �����ID=" & lng�����ID
        bln�ֵ��� = Not rsBalanceByNo.EOF
    Else
        bln�ֵ��� = False
    End If
    
    dblMoney = 0: strXML = ""
    If blnȫ�� Then
        '1.ȫ�ˣ����ֵ���
        If bln�ֵ��� = False Then
            With rsBalance
                .Filter = "��������ID=" & lng��������ID & " And �����ID=" & lng�����ID
                Do While Not .EOF
                    dblMoney = dblMoney + Val(Nvl(!��Ԥ��))
                    
                    rsDelFeeRecord.AddNew
                    rsDelFeeRecord!ʵ�ս�� = Val(Nvl(!��Ԥ��))
                    rsDelFeeRecord.Update
                    
                    .MoveNext
                Loop
            End With
        Else '2.ȫ�ˣ��ֵ���(�����ǲ��ֵ���ȫ��)
            With rsBalanceByNo
                .Filter = "��������ID=" & lng��������ID & " And �����ID=" & lng�����ID
                Do While Not .EOF
                    blnFind = True
                    If Not rsDelFeeRecordByNo Is Nothing Then
                        rsDelFeeRecordByNo.Filter = "NO='" & Nvl(!NO) & "'"
                        blnFind = Not rsDelFeeRecordByNo.EOF
                    End If
                    
                    If blnFind Then
                        dblMoney = dblMoney + Val(Nvl(!���))
                        
                        rsDelFeeRecord.AddNew
                        rsDelFeeRecord!NO = Nvl(!NO)
                        rsDelFeeRecord!ʵ�ս�� = Val(Nvl(!���))
                        rsDelFeeRecord.Update
                    End If
                    
                    .MoveNext
                Loop
            End With
        End If
        dblDelMoney = dblMoney
    End If
    
    '3.�����ˣ����ֵ���
    If bln�ֵ��� = False Then
        With rsDelFeeRecord
            .Filter = "": dblMoney = 0
            Do While Not .EOF
                dblMoney = dblMoney + Val(Nvl(!ʵ�ս��))
                .MoveNext
            Loop
        End With
        dblMoneyTotal = dblMoney
        
        dblCurMoney = 0
        Set cllBalance = New Collection
        With rsBalance
            .Filter = "��������ID=" & lng��������ID & " And �����ID=" & lng�����ID
            Do While Not .EOF
                strKey = "_" & Nvl(!���㷽ʽ)
                If CollectionExitsValue(cllBalance, strKey) Then
                    dblCurMoney = cllBalance(strKey)(1) + Val(Nvl(!��Ԥ��))
                    cllBalance.Remove strKey
                Else
                    dblCurMoney = Val(Nvl(!��Ԥ��))
                End If
                If dblCurMoney <> 0 Then
                    cllBalance.Add Array(Nvl(!���㷽ʽ), dblCurMoney), strKey
                End If
                
                .MoveNext
            Loop
        
            For j = 1 To cllBalance.Count
                If dblMoney > cllBalance(j)(1) Then
                    dblCurMoney = cllBalance(j)(1)
                Else
                    dblCurMoney = dblMoney
                End If
                If dblDelMoney < dblCurMoney Then dblCurMoney = dblDelMoney
                If dblCurMoney <= 0 Then Exit For
            
                .Filter = "����ID=" & lngԭ����ID & " And ��������ID=" & lng��������ID & _
                    " And ���㷽ʽ='" & cllBalance(j)(0) & "'" & " And �����ID=" & lng�����ID
                If .EOF = False Then
                    strXML = strXML & "    <TK>" & vbCrLf
                    strXML = strXML & "      <TKFS>" & cllBalance(j)(0) & "</TKFS>" & vbCrLf
                    strXML = strXML & "      <TKJE>" & dblCurMoney & "</TKJE>" & vbCrLf
                    strXML = strXML & "      <JYLSH>" & Nvl(!������ˮ��) & "</JYLSH>" & vbCrLf
                    strXML = strXML & "      <JYSM>" & Nvl(!����˵��) & "</JYSM>" & vbCrLf
                    strXML = strXML & "      <DJH>" & "" & "</DJH>" & vbCrLf
                    strXML = strXML & "    </TK>" & vbCrLf
                End If
                dblMoney = dblMoney - dblCurMoney
                dblDelMoney = dblDelMoney - dblCurMoney
            Next
        End With
        
        If strXML <> "" Then ZlGetBalanceXML = "  <TKLIST>" & vbCrLf & strXML & "  </TKLIST>"
        Exit Function
    End If
    
    '4.�����ˣ��ֵ���
    Set cllNo = New Collection
    With rsDelFeeRecord
        .Filter = "": dblMoney = 0
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(!ʵ�ս��))
            
            strKey = "_" & Nvl(!NO)
            If CollectionExitsValue(cllNo, strKey) Then
                dblCurMoney = cllNo(strKey)(1) + Val(Nvl(!ʵ�ս��))
                cllNo.Remove strKey
            Else
                dblCurMoney = Val(Nvl(!ʵ�ս��))
            End If
            cllNo.Add Array(Nvl(!NO), dblCurMoney), strKey
            
            .MoveNext
        Loop
    End With
    dblMoneyTotal = dblMoney
    
    For i = 1 To cllNo.Count
        dblMoney = cllNo(i)(1): dblCurMoney = 0
        Set cllBalance = New Collection
        With rsBalanceByNo
            .Filter = "��������ID=" & lng��������ID & " And No='" & cllNo(i)(0) & "'" & _
                " And �����ID=" & lng�����ID
            Do While Not .EOF
                strKey = "_" & Nvl(!���㷽ʽ)
                If CollectionExitsValue(cllBalance, strKey) Then
                    dblCurMoney = cllBalance(strKey)(1) + Val(Nvl(!���))
                    cllBalance.Remove strKey
                Else
                    dblCurMoney = Val(Nvl(!���))
                End If
                If dblCurMoney <> 0 Then
                    cllBalance.Add Array(Nvl(!���㷽ʽ), dblCurMoney), strKey
                End If
                
                .MoveNext
            Loop
        
            For j = 1 To cllBalance.Count
                If dblMoney > cllBalance(j)(1) Then
                    dblCurMoney = cllBalance(j)(1)
                Else
                    dblCurMoney = dblMoney
                End If
                If dblDelMoney < dblCurMoney Then dblCurMoney = dblDelMoney
                If dblCurMoney <= 0 Then Exit For
                
                .Filter = "����ID=" & lngԭ����ID & " And ��������ID=" & lng��������ID & _
                    " And No='" & cllNo(i)(0) & "' And ���㷽ʽ='" & cllBalance(j)(0) & "'" & _
                    " And �����ID=" & lng�����ID
                If .EOF = False Then
                    strXML = strXML & "    <TK>" & vbCrLf
                    strXML = strXML & "      <TKFS>" & cllBalance(j)(0) & "</TKFS>" & vbCrLf
                    strXML = strXML & "      <TKJE>" & dblCurMoney & "</TKJE>" & vbCrLf
                    strXML = strXML & "      <JYLSH>" & Nvl(!������ˮ��) & "</JYLSH>" & vbCrLf
                    strXML = strXML & "      <JYSM>" & Nvl(!����˵��) & "</JYSM>" & vbCrLf
                    strXML = strXML & "      <DJH>" & cllNo(i)(0) & "</DJH>" & vbCrLf
                    strXML = strXML & "    </TK>" & vbCrLf
                End If
                dblMoney = dblMoney - dblCurMoney
                dblDelMoney = dblDelMoney - dblCurMoney
            Next
        End With
    Next
    
    If strXML <> "" Then ZlGetBalanceXML = "  <TKLIST>" & vbCrLf & strXML & "  </TKLIST>"
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeRecordFromGrid(ByVal vsfBill As VSFlexGrid, _
    Optional ByVal bln�쳣���� As Boolean) As ADODB.Recordset
    '�ӽ������л�ȡ�����˷���Ŀ
    Dim i As Integer
    Dim rsDelFeeRecord As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Set rsDelFeeRecord = ZlMakeDelFeeRecord()
    
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Or bln�쳣���� Then
                rsDelFeeRecord.AddNew
                rsDelFeeRecord!NO = .TextMatrix(i, .ColIndex("���ݺ�"))
                rsDelFeeRecord!��� = .RowData(i)
                rsDelFeeRecord!ʵ�ս�� = Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                rsDelFeeRecord.Update
            End If
        Next
    End With
    
    Set ZlGetDelFeeRecordFromGrid = rsDelFeeRecord
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelFeeDetailXMLFromGrid(ByVal vsfBill As VSFlexGrid, _
    Optional ByVal bln�쳣���� As Boolean) As String
    '�ӽ������л�ȡ�����������˷ѽӿ�zlRetuenCheck�з����б�
    '��Σ�
    '   lng����ID - ����ID
    '���أ�
    '      <TFLIST> //�˷��б�
    '        <NO></NO> // �˷ѵ���
    '        <TFITEM> //�˷���
    '          <SerialNum></SerialNum> //���
    '          ��
    '        </TFITEM>
    '      </TFLIST>
    '      ...
    Dim i As Integer
    Dim strXML As String, blnFindSelectItem As Boolean
    Dim strNo As String, strPriorNO As String
    
    On Error GoTo ErrHandler
    strXML = "": blnFindSelectItem = False
    
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Or bln�쳣���� Then
                blnFindSelectItem = True
                strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                If strNo <> strPriorNO Then
                    If strPriorNO <> "" Then
                        strXML = strXML & "    </TFITEM>" & vbCrLf
                        strXML = strXML & "  </TFLIST>" & vbCrLf
                    End If
                    strXML = strXML & "  <TFLIST>" & vbNewLine '�˷��б�
                    strXML = strXML & "    <NO>" & strNo & "</NO>" & vbCrLf '�˷ѵ���
                    strXML = strXML & "    <TFITEM>" & vbCrLf '�˷���
                End If
                strXML = strXML & "      <SerialNum>" & .RowData(i) & "</SerialNum>" & vbCrLf '���
                strPriorNO = strNo
            End If
        Next
    End With
    If blnFindSelectItem = False Then Exit Function
    
    strXML = strXML & "    </TFITEM>" & vbCrLf
    strXML = strXML & "  </TFLIST>" & vbCrLf
    
    ZlGetDelFeeDetailXMLFromGrid = strXML
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlCheckThreeSwapValied(frmMain As Object, ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByVal strPatiName As String, ByVal strSex As String, ByVal strOld As String, _
    ByRef objSquareCard As Object, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���㷽ʽ As String, ByRef dblMoney As Double, ByVal strNos As String, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByRef dbl�ʻ���� As Double, Optional ByRef str�½��㷽ʽ As String, _
    Optional ByVal rsClassMoney As ADODB.Recordset, Optional ByVal str������Դ As String, _
    Optional ByRef cllSquareBalance As Collection) As Boolean
    '����:����֧�����׼��
    '���:
    '   str���㷽ʽ-��ǰ���㷽ʽ
    '   dblMoney-֧�����
    '   strNos-����֧�����漰�ĵ���
    '   rsClassMoney-���������ϸ(ʹ�����ѿ�֧��ʱ����)
    '   str������Դ-��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '���Σ�
    '   strCardNo-ˢ������
    '   strPassWord-��������
    '   dbl�ʻ����-�ʻ����
    '   str�½��㷽ʽ-ˢ���ӿڷ��صĽ��㷽ʽ������ʽ�����㷽ʽ|������
    '   cllSquareBalance- �ѿ�֧����ϸ
    '����:���׺Ϸ�����true,���򷵻�False
    Dim strXMLExpend As String
    Dim str���㷽ʽ_Out As String, dbl������_Out As Double
    Dim strExpand As String

    On Error GoTo ErrHandler
    If objSquareCard Is Nothing Then Exit Function
    
    str�½��㷽ʽ = ""
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXMLIn As String = "", _
        Optional ByVal str������Դ As String, _
        Optional ByVal lng����ID As Long, _
        Optional ByRef str���㷽ʽ_Out As String = "", _
        Optional ByRef dbl������_Out As Double = 0) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:����ָ��֧�����,����ˢ������
        '���:rsClassMoney:�շ����,���
        '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
        '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
        '       dblBrushTotaled-������Ч,��ʾ�Ѿ�ˢ���ѿ��ܶ�(��Ҫ���ڶ��ˢ��)
        '       str�ϴ��������-�ϴ�ˢ����ʱ���������(ͬ�ζ��ˢ���ѿ�ʱ,��Ҫ��鱾��ˢ��������ϴ�����Ƿ�һ��,��һ�²�����ˢ������)
        '       varSquareBalance- Collection����,��ǰ�Ѿ�ˢ������Ϣ(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ����� ))
        '       blnԤ��-�Ƿ�תԤ��
        '       blnAllPay-�Ƿ����ȫ֧����true-����δ֧���겻����ɽ��㣬false-����ֻ֧�����ֲ�����
        '       strXMLExpend-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
        '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
        '����:str�������-�������(���ѿ�����)
        '        lng���ѿ�ID-���ѿ���Ϣ.ID(���ѿ�����)
        '       strCardNO-����ˢ���Ŀ���
        '       strPassWord-����ˢ������Ӧ������
        '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        '       str���㷽ʽ_Out-���صĽ��㷽ʽ
        '       dbl������_Out-���صĽ�����
        '����:�ɹ�,����true,���򷵻�False
    strXMLExpend = "<IN><CZLX>0</CZLX></IN>"
    If objSquareCard.zlBrushCard(frmMain, lngModule, rsClassMoney, lng�����ID, bln���ѿ�, _
        strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, False, True, False, False, _
        cllSquareBalance, False, False, strXMLExpend, str������Դ, lng����ID, _
        str���㷽ʽ_Out, dbl������_Out) = False Then Exit Function
    
    If str���㷽ʽ_Out <> "" Then
        If RoundEx(dblMoney, 6) <> Round(dbl������_Out, 6) Then
            MsgBox str���㷽ʽ & " ʵ��ˢ��֧�����(" & Format(dbl������_Out, "0.00") & ")" & _
                "��Ӧ�����(" & Format(dblMoney, "0.00") & ")���ȣ�֧��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        str�½��㷽ʽ = str���㷽ʽ_Out & "|" & dbl������_Out
    End If
    
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNos As String, _
        Optional ByVal strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�ʻ��ۿ�׼��
        '���:frmMain-���õ�������
        '       lngModule-���õ�ģ���
        '       strCardTypeID-�����ID
        '       strCardNo-����
        '       dblMoney-֧�����(�˿�ʱΪ����)
        '       strNos-����֧�����漰�ĵ���
        '       strXMLExpend-(XML��:��֤����:��������)
        '����:
        '   strXMLExpend-(XML��:������Ϣ)
        '����:�ۿ�Ϸ�,����true,���򷵻�Flase
        '����:���˺�
        '����:2011-05-26 16:42:43
        '˵��:
        '   �ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬ �����ٵ��ÿۿ��ǰ�� _
        '   �Ƚ������ݵĺϷ��Լ��,�Ա�������������
    If objSquareCard.zlPaymentCheck(frmMain, lngModule, lng�����ID, bln���ѿ�, _
        strCardNo, dblMoney, strNos) = False Then Exit Function
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        strExpand As String, dblMoney As Double, _
        Optional bln���ѿ� As Boolean = False) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:��ȡ�ʻ����
        '���:frmMain-���õ�������
        '        lngModule-ģ���
        '       strCardNo-����
        '       strExpand-Ԥ����Ϊ��,�Ժ���չ
        '       bln���ѿ�-�Ƿ�Ϊ���ѿ�
        '����:dblMoney-�����ʻ����
        '����:��������    True:���óɹ�,False:����ʧ��
        '����:���˺�
        '����:2011-05-26 16:29:48
        '˵��:
        '       ��������Ҫ�ۿ�ĵط�����Ҫ����ʻ�����Ƿ����,�ʻ�������ʱ������ۿ�.
        '       ���ĳЩ�������ӿڲ��������ӿڣ����Թ̶�����һ���Ľ�
    If objSquareCard.zlGetAccountMoney(frmMain, lngModule, _
        lng�����ID, strCardNo, strExpand, dbl�ʻ����, bln���ѿ�) = False Then Exit Function
    If dbl�ʻ���� <> 0 And dbl�ʻ���� < dblMoney Then
        MsgBox str���㷽ʽ & " �ʻ����㣡", vbInformation, gstrSysName
        Exit Function
    End If

    ZlCheckThreeSwapValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef cllBalance As Collection, ByRef strExpend As String) As Boolean
    '���ܣ���������֧������
    '��Σ�
    '   strXMLExpend:XML��
    '    <OUTPUT>
    '        <JYLIST> //�����б�
    '            <JY> //���浽Ԥ����¼ʱ����������ˮ�ż�����˵�����ܴ���
    '                <JYFS>���׷�ʽ</JYFS> //���׷�ʽ:�����㷽ʽ.����
    '                <JYJE>���׽��</JYJE>
    '                <JYLSH>������ˮ��</JYLSH>
    '                <JYSM>����˵��</JYSM>
    '                <DJH>���ݺ�</DJH> //���ݺ�,�൥���շ�ʱ���� ���洢��"ҽ��������ϸ"����,��Ҫ�Ƿֵ��ݱ���
    '                <SFPTJS>�Ƿ���ͨ����</SFPTJS> //�Ƿ���ͨ����(1-��ͨ����;0-һ��ͨ����):Ϊ1ʱ����Ԥ����¼�в���д�����ID,������һ��ͨ����
    '            </JY>
    '            ...
    '        </JYLIST>
    '        <Expends> //������չ��Ϣ
    '            <Expend> //���浽Ԥ����¼ʱ����������ˮ�ż�����˵�����ܴ���
    '                <XMMC>��Ŀ����</XMMC> //���׷�ʽ:�����㷽ʽ.����
    '                <XMNR>��Ŀ����</XMNR>
    '            </Expend>
    '            ...
    '        </Expends>
    '    </OUTPUT>
    '���Σ�
    '   dblOutMoney - ʵ��֧�����
    '   cllBalance - �������ݣ���ʽ��Array("���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����",������ˮ��,����˵��)
    '   strExpend - ��չ���ݣ���ʽ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str������ˮ�� As String, str����˵�� As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection: strExpend = ""
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '������Ϣ
    Call zlXML_GetRows("JYLIST/JY", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("JYFS", i, strValue)
        strBalance = strValue   '���㷽ʽ
        Call zlXML_GetNodeValue("JYJE", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '������
        dblOutMoney = dblOutMoney + Val(strValue)
        strBalance = strBalance & "|" & " " '�������
        strBalance = strBalance & "|" & " " '����ժҪ
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '���ݺ�
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '�Ƿ���ͨ����
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str������ˮ�� = strValue '������ˮ��
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str����˵�� = strValue   '����˵��
        
        cllBalance.Add Array(strBalance, str������ˮ��, str����˵��)
    Next
    
    '��չ��Ϣ
    Call zlXML_GetRows("Expends/Expend", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("XMMC", i, strValue)
        strExpend = strExpend & "||" & strValue '��Ŀ����
        Call zlXML_GetNodeValue("XMNR", i, strValue)
        strExpend = strExpend & "|" & strValue '��Ŀ����
    Next
    If strExpend <> "" Then strExpend = Mid(strExpend, 3)
    ZLGetThreeSwapXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetThreeSwapBalanceSQL(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal lngҽ�ƿ����ID As Long, ByVal byt�������� As Byte, _
    ByVal strˢ������ As String, ByVal str���㷽ʽ As String, _
    Optional ByVal lng��������ID As Long, Optional ByVal blnɾ��ԭ���� As Boolean, _
    Optional ByVal str������ˮ�� As String, Optional ByVal str����˵�� As String, _
    Optional ByVal bytУ�Ա�־ As Byte = 1) As String
    '��ȡ֧������SQL
    'byt�������� 1-����������,3-���ѿ�����,4-���������ֽ��㷽ʽ����
    Dim strSQL As String
    
    ' Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '  --   0-��ͨ�շѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����֧Ʊ��_In:������
    '  --   3-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
    '  --     ����֧Ʊ��_In:������
    '  --   4-���������㣬���ֽ��㷽ʽ:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    strSQL = strSQL & byt�������� & ","
    '    ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '    ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '    ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '    ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID = 0, "NULL", lngҽ�ƿ����ID) & ","
    '    ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & strˢ������ & "',"
    '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '    �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '    �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '    -- �����_In:��������ʱ,����
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ɽ���_In      Number := 0,
    '    -- ��ɽ���_In:1-����շ�;0-δ����շ�
    strSQL = strSQL & "" & 0 & ","
    '  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��Ԥ������ids_In Varchar2 := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���½������_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    '  ��������id_In    ����Ԥ����¼.��������id%Type := Null
    strSQL = strSQL & IIf(lng��������ID = 0, "NULL", lng��������ID) & ","
    '  ɾ��ԭ����_In    Number := 0,
    strSQL = strSQL & "" & IIf(blnɾ��ԭ����, "1", "0") & ","
    '  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0
    strSQL = strSQL & "" & bytУ�Ա�־ & ")"
    ZlGetThreeSwapBalanceSQL = strSQL
End Function

Public Function ZlCheckThreeSwapDelValied(frmMain As Object, ByVal lngModule As Long, _
    ByVal strPatiName As String, ByVal strSex As String, ByVal strOld As String, _
    ByRef objSquareCard As Object, ByVal lng�����ID As Long, _
    ByVal blnTransfer As Boolean, ByVal dblMoney As Double, ByVal strԭ����ID As String, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByVal str������ˮ�� As String, Optional ByVal str����˵�� As String, _
    Optional ByVal strXMLExpend As String, Optional ByVal bln�Ƿ��˿��鿨 As Boolean) As Boolean
    '����:�����˿�׼��,�������ѿ�
    '���:
    '     dblMoney-�˿���
    '����:���׺Ϸ�����true,���򷵻�False
    Dim strBalanceIDs As String
    Dim strExpend As String
    
    On Error GoTo ErrHandler
    If objSquareCard Is Nothing Then Exit Function
    
    'ת��ģʽ
    If blnTransfer Then
        'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln���ѿ� As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl��� As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln�˷� As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln���� As Boolean = False, _
            Optional ByVal bln�����ֹ As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal blnתԤ�� As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXMLIn As String = "", _
            Optional ByVal str������Դ As String, _
            Optional ByVal lng����ID As Long, _
            Optional ByRef str���㷽ʽ_Out As String = "", _
            Optional ByRef dbl������_Out As Double = 0) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:����ָ��֧�����,����ˢ������
            '���:rsClassMoney:�շ����,���
            '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
            '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
            '       dblBrushTotaled-������Ч,��ʾ�Ѿ�ˢ���ѿ��ܶ�(��Ҫ���ڶ��ˢ��)
            '       str�ϴ��������-�ϴ�ˢ����ʱ���������(ͬ�ζ��ˢ���ѿ�ʱ,��Ҫ��鱾��ˢ��������ϴ�����Ƿ�һ��,��һ�²�����ˢ������)
            '       varSquareBalance- Collection����,��ǰ�Ѿ�ˢ������Ϣ(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ����� ))
            '       blnԤ��-�Ƿ�תԤ��
            '       blnAllPay-�Ƿ����ȫ֧����true-����δ֧���겻����ɽ��㣬false-����ֻ֧�����ֲ�����
            '       strXMLExpend-����������XML���,Ŀǰ��ʽ����:
            '       <IN>
            '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
            '       </IN>
            '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
            '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
            '����:str�������-�������(���ѿ�����)
            '        lng���ѿ�ID-���ѿ���Ϣ.ID(���ѿ�����)
            '       strCardNO-����ˢ���Ŀ���
            '       strPassWord-����ˢ������Ӧ������
            '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
            '       str���㷽ʽ_Out-���صĽ��㷽ʽ
            '       dbl������_Out-���صĽ�����
            '����:�ɹ�,����true,���򷵻�False
        strExpend = "<IN><CZLX>1</CZLX></IN>"
        If objSquareCard.zlBrushCard(frmMain, lngModule, Nothing, lng�����ID, False, _
            strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
            Nothing, False, False, strExpend) = False Then Exit Function

        'zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
            ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, _
            Optional ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:ת�ʼ��
            '���:
            '   frmMain-���õ�������
            '   lngModule-HIS����ģ���
            '   lngCardTypeID-�����ID
            '   strCardNo-����
            '   dblMoney-ת�ʽ��(����ʱΪ����)
            '   strBalanceID-����֧������ID��4-�����˷�ҵ��Ϊԭ����ID
            '   strXMLExpend-XML��:
            '       <IN>
            '           <CZLX >��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��
            '       </IN>
            '����:strXMLExpend-XML��:
            '        <OUT>
            '           <ERRMSG>������Ϣ</ERRMSG >
            '        </OUT>
            '����:�������ݺϷ�,����True:���򷵻�False
            '������:ҽ���������(����ʱ����)
            '˵��:
            '  ��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
            '  ��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
        strExpend = "<IN><CZLX>4</CZLX></IN>"
        If objSquareCard.zltransferAccountsCheck(frmMain, lngModule, lng�����ID, _
            strCardNo, dblMoney, strԭ����ID, strExpend) = False Then Exit Function
    Else
        'zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:�ʻ����˽���ǰ�ļ��
            '���:frmMain-���õ�������
            '       lngModule-���õ�ģ���
            '       lngCardTypeID-�����ID
            '       strCardNo-����
            '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
            '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������
            '       dblMoney-�˿���
            '       strSwapNo-������ˮ��(�˿�ʱ���),���ղ������ʱ�����
            '       strSwapMemo-����˵��(�˿�ʱ����),���ղ������ʱ�����
            '       strXMLExpend    XML IN
            '        <TFDATA>   //�˷�����
            '            <YCTF>�쳣�˷ѱ�־<YCTF> //1-�쳣����;0-�˷Ѵ˽ڵ����û����
            '            <TFLIST>  //�˷��б�
            '                <NO></NO>  // �˷ѵ���
            '                <TFITEM>     //�˷���
            '                    <SerialNum>���</SerialNum>
            '                    ��.
            '                </ TFITEM >
            '            </TFLIST>
            '
            '            <TKLIST>   //�˿��б�35.90��ǰ�޴����ݣ�
            '                <TK>
            '                    <TKFS>�˿ʽ</TKFS>// Varchar2    20
            '                    <TKJE>�˿���</TKJE>//NUMBER
            '                    <JSLSH>ԭ������ˮ��</JSLSH>//   Varchar2    50
            '                    <JYSM><ԭ����˵��</JYSM>//  Varhcar2    500
            '                    <DJH>���ݺ�</DJH> //    Varchar2    8
            '                </TK>
            '                ....
            '            </TKLIST>
            '        </TFDATA>
            '����:�˿�Ϸ�,����true,���򷵻�Flase
            '˵��:
            '    �ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,
            '    �Ա�������������
        strBalanceIDs = "3|" & strԭ����ID
        If objSquareCard.zlReturnCheck(frmMain, lngModule, lng�����ID, False, strCardNo, _
            strBalanceIDs, dblMoney, str������ˮ��, str����˵��, strXMLExpend) = False Then Exit Function
    
        If bln�Ƿ��˿��鿨 Then
           'zlBrushCard(frmMain As Object, _
                ByVal lngModule As Long, _
                ByVal rsClassMoney As ADODB.Recordset, _
                ByVal lngCardTypeID As Long, _
                ByVal bln���ѿ� As Boolean, _
                ByVal strPatiName As String, ByVal strSex As String, _
                ByVal strOld As String, ByRef dbl��� As Double, _
                Optional ByRef strCardNo As String, _
                Optional ByRef strPassWord As String, _
                Optional ByRef bln�˷� As Boolean = False, _
                Optional ByRef blnShowPatiInfor As Boolean = False, _
                Optional ByRef bln���� As Boolean = False, _
                Optional ByVal bln�����ֹ As Boolean = True, _
                Optional ByRef varSquareBalance As Variant, _
                Optional ByVal blnתԤ�� As Boolean = False, _
                Optional ByVal blnAllPay As Boolean = False, _
                Optional ByVal strXMLIn As String = "", _
                Optional ByVal str������Դ As String, _
                Optional ByVal lng����ID As Long, _
                Optional ByRef str���㷽ʽ_Out As String = "", _
                Optional ByRef dbl������_Out As Double = 0) As Boolean
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '����:����ָ��֧�����,����ˢ������
                '���:rsClassMoney:�շ����,���
                '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
                '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
                '       dblBrushTotaled-������Ч,��ʾ�Ѿ�ˢ���ѿ��ܶ�(��Ҫ���ڶ��ˢ��)
                '       str�ϴ��������-�ϴ�ˢ����ʱ���������(ͬ�ζ��ˢ���ѿ�ʱ,��Ҫ��鱾��ˢ��������ϴ�����Ƿ�һ��,��һ�²�����ˢ������)
                '       varSquareBalance- Collection����,��ǰ�Ѿ�ˢ������Ϣ(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ����� ))
                '       blnԤ��-�Ƿ�תԤ��
                '       blnAllPay-�Ƿ����ȫ֧����true-����δ֧���겻����ɽ��㣬false-����ֻ֧�����ֲ�����
                '       strXMLExpend-����������XML���,Ŀǰ��ʽ����:
                '       <IN>
                '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
                '       </IN>
                '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
                '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
                '����:str�������-�������(���ѿ�����)
                '        lng���ѿ�ID-���ѿ���Ϣ.ID(���ѿ�����)
                '       strCardNO-����ˢ���Ŀ���
                '       strPassWord-����ˢ������Ӧ������
                '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                '       str���㷽ʽ_Out-���صĽ��㷽ʽ
                '       dbl������_Out-���صĽ�����
                '����:�ɹ�,����true,���򷵻�False
            strExpend = "<IN><CZLX>2</CZLX></IN>"
            If objSquareCard.zlBrushCard(frmMain, lngModule, Nothing, lng�����ID, False, _
                strPatiName, strSex, strOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
                Nothing, False, False, strExpend) = False Then Exit Function
        End If
    End If
    
    ZlCheckThreeSwapDelValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapDelXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef cllBalance As Collection) As Boolean
    '���ܣ����������˿�����
    '��Σ�
    '   strXMLExpend:XML��
    '    <OUTPUT>
    '        <TKLIST>
    '            <TK>
    '                <TKFS>�˿ʽ</TKFS>
    '                <TKJE>������</TKJE>
    '                <JYLSH>�˿����ˮ��</JYLSH>
    '                <JYSM>�˿��˵��</JYSM>
    '                <DJH>���ݺ�</DJH>
    '                <SFPTJS>�Ƿ���ͨ����</SFPTJS>
    '            </TK>
    '            ��
    '        </TKLIST>
    '    </OUTPUT>
    '   blnDelMoney - �Ƿ�Խ��ȡ�෴��
    '���Σ�
    '   dblOutMoney - ʵ���˿���
    '   cllBalance - �������ݣ���ʽ��Array("���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����",������ˮ��,����˵��)
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str������ˮ�� As String, str����˵�� As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '������Ϣ
    Call zlXML_GetRows("TKLIST/TK", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("TKFS", i, strValue)
        strBalance = strValue '�˿ʽ
        Call zlXML_GetNodeValue("TKJE", i, strValue)
        strBalance = strBalance & "|" & -1 * Val(strValue)    '������
        dblOutMoney = dblOutMoney + -1 * Val(strValue)
        strBalance = strBalance & "|" & " " '�������
        strBalance = strBalance & "|" & " " '����ժҪ
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '���ݺ�
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '�Ƿ���ͨ����
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str������ˮ�� = strValue '������ˮ��
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str����˵�� = strValue   '����˵��
        
        cllBalance.Add Array(strBalance, str������ˮ��, str����˵��)
    Next
    ZLGetThreeSwapDelXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapDelBalanceSQL(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal lngҽ�ƿ����ID As Long, ByVal byt�������� As Byte, _
    ByVal strˢ������ As String, ByVal str���㷽ʽ As String, _
    Optional ByVal lng��������ID As Long, Optional ByVal blnɾ��ԭ���� As Boolean, _
    Optional ByVal str������ˮ�� As String, Optional ByVal str����˵�� As String, _
    Optional ByVal bytУ�Ա�־ As Byte = 1) As String
    '��ȡ�˿����SQL
    'byt�������� 2-����������,4-���ѿ�����,5-���������ֽ��㷽ʽ����
    Dim strSQL As String
    
    'Zl_�����˷ѽ���_Modify(
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --   5.�������˷ѽ��㣬���ֽ��㷽ʽ:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����"
    '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    strSQL = strSQL & "" & byt�������� & ","
    '  ����id_In        ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In        ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In      Varchar2,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �����id_In      ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(lngҽ�ƿ����ID = 0, "NULL", lngҽ�ƿ����ID) & ","
    '  ����_In          ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & strˢ������ & "',"
    '  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����˷�_In      Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  ԭ����id_In      ����Ԥ����¼.����id%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ʣ��תԤ��_In    Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��Ԥ������ids_In Varchar2 := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��������id_In    ����Ԥ����¼.��������id%Type := Null,
    strSQL = strSQL & "" & IIf(lng��������ID = 0, "NULL", lng��������ID) & ","
    '  ɾ��ԭ����_In    Number := 0,
    strSQL = strSQL & "" & IIf(blnɾ��ԭ����, "1", "0") & ","
    '  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0
    strSQL = strSQL & "" & bytУ�Ա�־ & ")"
    ZLGetThreeSwapDelBalanceSQL = strSQL
End Function

Public Function ZLCheckThreeSwapDelToCash(frmMain As Object, ByVal lngModule As Long, _
    ByRef objSquareCard As Object, rsBalance As ADODB.Recordset, _
    ByVal lngԭ����ID As Long, ByVal lngCardTypeID As Long, _
    ByVal dblMoney As Double, ByVal strXMLExpend As String, _
    Optional blnDelDefaultCash_Out As Boolean, _
    Optional strDefaultDelBalance_Out As String) As Boolean
    '�������㽻�����ּ��
    Dim strCardNo As String, lng����ID As Long
    Dim strSwapNO As String, strSwapMemo As String
    
    On Error GoTo ErrHandler
    If rsBalance Is Nothing Then Exit Function
    
    rsBalance.Filter = "����ID=" & lngԭ����ID & " And �����id=" & lngCardTypeID & " And �˷�=0"
    If rsBalance.EOF Then Exit Function
    strCardNo = Nvl(rsBalance!����)
    lng����ID = Nvl(rsBalance!����ID)
    strSwapNO = Nvl(rsBalance!������ˮ��)
    strSwapMemo = Nvl(rsBalance!����˵��)
    
    strXMLExpend = "<INPUT>" & vbCrLf & _
                        strXMLExpend & vbCrLf & _
                    "</INPUT>"
    'zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String, _
        Optional blnDelDefaultCash_Out As Boolean, Optional strDefaultDelBalance_Out As String) As Boolean
    '����:���ֽ��׼��
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID
    '       strCardNo-����
    '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�˿�ʱ���)���ֽ��㷽ʽʱ��������Ϊ��һ�����㷽ʽ�Ľ�����ˮ��
    '       strSwapMemo-����˵��(�˿�ʱ����) ���ֽ��㷽ʽʱ��������Ϊ��һ�����㷽ʽ�Ľ���˵��
    '       strXMLExpend    XML IN  10.35.90���֧��
    '        <INPUT>
    '            <TKLIST>    //�����˿��б�
    '                <TK>
    '                    <TKFS>�˿ʽ</TKFS>
    '                    <TKJE>�˿���</TKJE>
    '                    <JYLSH>ԭ������ˮ��</JYLSH>
    '                    <JYSM>ԭ����˵��</JYSM>
    '                </TK>
    '                ....
    '            </TKLIST>
    '        </INPUT>
    '����:
    '       blnDelDefaultCash_Out-�Ƿ�ȱʡ���֣��ӿڷ���trueʱ��Ч��trueʱ����ʾȱʡ�˳��ֽ�ȱʡ��ʽΪ:strȱʡ���ַ�ʽ_Out����ֵ),����ȱʡ�˻�ԭ�������������Աѡ����Ϊ�ֽ�
    '       strDefaultDelBalance_Out-ȱʡ���ַ�ʽ,���磺֧Ʊ���ֽ��
    '       strXMLExpend:10.35.90���֧��
    '        <OUTPUT>
    '            <SFQSTX>�Ƿ�ȱʡ����<SFQSTX>//NUMBER 1 �Ƿ�ȱʡ����: 1-ȱʡ;0-��ȱʡ��ȱʡ�˻�ԭ�������������Ա��������
    '            <QSTKFS>ȱʡ�����˿ʽ</QSTKFS>//Varchar2 20 ȱʡ�����˿ʽ�����㷽ʽ.����
    '                    1.���������������Ľ��㷽ʽ
    '                    2.Ӧ����ʹ�ã�ҽ������㣬һ��ͨ����Ľ��㷽ʽ�����ѿ���һЩ������㷽ʽ���������෽ʽ��������ʹ����Щ��ʽ
    '        </OUTPUT>
    '����:���ֺϷ�,����true,���򷵻�Flase
    If objSquareCard.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, lng����ID, dblMoney, _
        strSwapNO, strSwapMemo, strXMLExpend, _
        blnDelDefaultCash_Out, strDefaultDelBalance_Out) = False Then Exit Function
    
    ZLCheckThreeSwapDelToCash = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
