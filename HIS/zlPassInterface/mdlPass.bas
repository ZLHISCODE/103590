Attribute VB_Name = "mdlPass"
Option Explicit

Public Function OutAdviceCheckWarn_MK( _
    ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset) As Long
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'���ܣ�����Passϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        1/33-�����Զ����(סԺ/����),2/34-�ύ�Զ����(סԺ/����),3-�ֹ��������
'        6-��ҩ����,12-��ҩ�о�,22-����״̬/����ʷ����(�༭)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0,6ʱ��Ҫ
'   lngRow=��ǰ��
'����
'   blnNoSave=���ڱ���Ƿ񱣴棨���ڽ��汣�水ť�����Կ��ƣ�
'   rsOut=����ҩƷ˵��
'���أ�������˷��ص���߼���ʾֵ,Ϊ-1,-2,-3��ʾû�н������
'      ���PASS�˵�ʱ������>=0��ʾ���Ե����˵�
' rsOut=ҽ���������
'˵������ҩ��飺�漰�����µ�����(������ִ��)����δֹͣ�ĳ���
'      ��ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPatiInfo As New ADODB.Recordset
    Dim rs��ҩ As ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, strƵ�� As String, str�÷�ID As String, str�����λ As String
    Dim str��ϱ��� As String, str������� As String, strTmp As String, strPre�÷� As String
    Dim strҽ��ID As String, str��� As String, str���� As String, str������λ As String
    Dim str�����ʶ As String, str�������� As String, str���ﵥλ As String, str��ҩĿ�� As String, str���� As String
    Dim str���ID As String, str��ҩ��IDs As String
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim lngTmp As Long
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long, j As Long
    Dim lng��ҩ��ID As Long, lngLight As Long
    
    Dim strType As String
    Dim str��� As String, str���� As String
    Dim strCurrentDate As String
    Dim arrLight(0 To 4) As String
    Dim str�������� As String, str����ҽ�� As String
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer
    Dim objDiag As clsDiagItem
    Dim rs��� As ADODB.Recordset
    Dim strҩƷID As String, str����ID As String
    Dim str��ҩ�䷽ As String
    
    Dim arrSQL As Variant
    
    lngMaxWarn = -1
    OutAdviceCheckWarn_MK = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11

    '����3.0
    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036ͬһ�����˶�����ʱ������Ϣÿ�ζ�Ҫ����
    '-------------------------------------------------------------
    strSQL = "Select B.ID as ����ID,B.����,B.�Ա�,A.��������," & _
             " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
             " From ������Ϣ A,���˹Һż�¼ B,���ű� C,��Ա�� E" & _
             " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID" & _
             " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.str�Һŵ�)
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    '������Ϣ
    strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                    "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                    "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
    Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng�Һ�ID)
    rsPatiInfo.Filter = "��Ŀ����='���'"
    If rsPatiInfo.RecordCount <> 0 Then str��� = NVL(rsPatiInfo!��¼����)
    rsPatiInfo.Filter = "��Ŀ����='����'"
    If rsPatiInfo.RecordCount <> 0 Then str���� = NVL(rsPatiInfo!��¼����)

    Call PassSetPatientInfo(gobjPati.lng����ID, rsTmp!����Id, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), str����, str���, _
                            rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), "")

    '���˲��˹���ʷ
    '-------------------------------------------------------
    Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, 0)

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!ҩ��ID & "", rsTmp!ҩ���� & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '���˲���״̬
    '------------------------------------------------------------------
    
    '* �����Ϣ
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If glngModel = PM_����༭ Then
        If Not gobjDiags Is Nothing Then
            With gobjDiags
                For i = 1 To .Count
                    If .Item(i).str������� <> "" Then
                        str��ϱ��� = IIf(.Item(i).str�������� <> "", .Item(i).str��������, .Item(i).str��ϱ���)
                        str������� = .Item(i).str�������
                        Call PassSetMedCond(i & "", str��ϱ���, str�������, "User", strCurrentDate, strCurrentDate)
                    End If
                Next
            End With
        End If
    Else
        Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng�Һ�ID, "1,11")
        For i = 1 To rsTmp.RecordCount
            Call PassSetMedCond(i & "", rsTmp!���� & "", rsTmp!���� & "", "User", strCurrentDate, strCurrentDate)
            rsTmp.MoveNext
        Next
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With gobjAdvice
            If IIf(glngModel = PM_����༭, .RowData(lngRow) <> 0, True) And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                'ȡҩƷ����
                If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                    strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҩƷ����)
                Else
                    strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҽ������)  '��ҩ����
                End If

                'ȡҩƷ��ҩ;��(��ǰ�ɼ��в������в�ҩ) ,������λ
                str�÷� = ""
                If glngModel = PM_����༭ Then
                    k = .FindRow(CLng(.TextMatrix(lngRow, gobjCOL.intCOL���ID)), lngRow + 1)
                    If k <> -1 Then str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                    strTmp = .TextMatrix(lngRow, gobjCOL.intCOL������λ)
                Else
                    str�÷� = .TextMatrix(lngRow, gobjCOL.intCOL�÷�)
                    If InStr(str�÷�, ",") > 0 Then str�÷� = Left(str�÷�, InStr(str�÷�, ",") - 1)
                    strTmp = .TextMatrix(lngRow, gobjCOL.intCOL����)
                    If Mid(strTmp, 1, 2) = "0." Then '������С��������⴦��
                        strTmp = Replace(strTmp, Format(Val(strTmp) & "", "0.####"), "") '�����嵥�������������� & ��������λ����
                    Else
                        strTmp = Replace(strTmp, Val(strTmp) & "", "")    '
                    End If
                End If
                '�����ѯҩƷ��Ϣ
                Call PassSetQueryDrug(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID), strҩƷ, strTmp, str�÷�)

                '���ò˵�����״̬
                
                OutAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
            ElseIf glngModel = PM_����ҽ���嵥 And .TextMatrix(lngRow, gobjCOL.intCOL�������) = "E" And .TextMatrix(lngRow, gobjCOL.intCol��������) = "4" Then
                OutAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If

    '����ʷ/����״̬�༭
    '-------------------------------------------------------------
    If lngCmd = 22 Then
        'lngCmd=21-ֻ��,22-��ǿ�Ʊ༭,23-ǿ�Ʊ༭
        If PassDoCommand(lngCmd) = 2 Then
            '�������ֵΪ2��ʾ"����ʷ/����״̬�༭"�������仯����Ҫ�����Զ����
            lngCmd = 34    'תΪ�Զ��������,����ִ��
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    '�����˽���ҩƷ˵������  �ҳ���Ϊ����༭��鹦��
    If (lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3) And glngModel = PM_����༭ And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_�������)
    End If
    '���벡��ҽ����Ϣ
    '-------------------------------------------------------------
    With gobjAdvice
        If lngCmd = 6 Then
            If glngModel = PM_����༭ Then
                strTmp = .RowData(lngRow)
            Else
                strTmp = .TextMatrix(lngRow, gobjCOL.intCOLID)
            End If
            Call PassSetWarnDrug(strTmp)   '��ҩ����(�Ѿ����ҽ��Ψһ��)
        Else
            '��ҩ��˻���ҩ�о�
            lngCount = 0
            curDate = zlDatabase.Currentdate
            strҩƷ = "": str�÷� = "": strƵ�� = ""
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                    blnDo = blnDo And (lngCmd = 12 Or Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = (InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4"))
                    blnDo = blnDo And (lngCmd = 12 Or Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                End If
                If blnDo Then
                    If glngModel = PM_����ҽ���嵥 And .TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4" Then
                        '��ȡ��ҩҽ����ID
                        str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        'ȡҩƷ����
                        If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                            strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                        Else
                            strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                        End If
                       
                        'ȡҩƷ��ҩ;��
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ
                        If str�÷� = "" Then
                            If glngModel = PM_����༭ Then
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                If k <> -1 Then
                                    If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                        str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                                    Else
                                        str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    End If
                                End If
                            Else
                                If Trim(.TextMatrix(i, gobjCOL.intCOL�÷�)) = "" Then
                                    str�÷� = strPre�÷�
                                Else
                                    str�÷� = Split(.TextMatrix(i, gobjCOL.intCOL�÷�), ",")(0)
                                End If
                                str�÷�ID = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")   '������
                                strPre�÷� = str�÷�
                            End If
                        End If
    
                        'ȡ��ҩƵ��(��/��),��Ϊ������������
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then strƵ�� = ""    'һ����ҩ���ظ�ȡ
                        If strƵ�� = "" Then
                            If glngModel = PM_����༭ Then
                                strƵ�� = GetFrequency(.TextMatrix(i, gobjCOL.intCOL�����λ), .TextMatrix(i, gobjCOL.intCOLƵ�ʴ���), .TextMatrix(i, gobjCOL.intCOLƵ�ʼ��))
                            Else
                                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), "")
                                strƵ�� = GetFrequency(str�����λ, intƵ�ʴ��� & "", intƵ�ʼ�� & "")
                            End If
                            str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                            If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                            str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                        End If
    
                        '����ҽ����Ϣ
                        If glngModel = PM_����༭ Then
                            Call PassSetRecipeInfo(.RowData(i), .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID), strҩƷ, _
                                                   .TextMatrix(i, gobjCOL.intCOL����), .TextMatrix(i, gobjCOL.intCOL������λ), strƵ��, _
                                                   Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd"), "", str�÷�, _
                                                   .TextMatrix(i, gobjCOL.intCOL���ID), 1, str����ҽ��)
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                '��ҩ,�г�ҩ
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = CLng(.RowData(i) & "")
                                    rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                    rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                    rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                                '��ҩ�䷽  ����˵����������ҩ������
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                        rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                        rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                        rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                        rsOut.Update
                                    End If
                                End If
                            End If
                        Else
                            strTmp = .TextMatrix(i, gobjCOL.intCOL����)
                            If Mid(strTmp, 1, 2) = "0." Then
                                strTmp = "0" & Val(strTmp)
                            Else
                                strTmp = Val(strTmp)
                            End If
                            
                            Call PassSetRecipeInfo(.TextMatrix(i, gobjCOL.intCOLID), .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID), strҩƷ, _
                                 strTmp, Replace(.TextMatrix(i, gobjCOL.intCOL����), strTmp, ""), strƵ��, _
                                 Format(.TextMatrix(i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd"), "", str�÷�, _
                                 .TextMatrix(i, gobjCOL.intCOL���ID), 1, str����ҽ��)
                        End If
                        lngCount = lngCount + 1
                    End If
                End If
            Next
            '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
            If glngModel = PM_����ҽ���嵥 Then
                If str��ҩ��IDs <> "" Then
                    Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                    With rs��ҩ
                        For i = 1 To .RecordCount
                            If !���ID & "" <> str���ID Then
                                str����ҽ�� = !����ҽ��
                                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                                strƵ�� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
                                str���ID = !���ID & ""
                            End If
                            Call PassSetRecipeInfo(!id, !ҩƷID & "", !ҩƷ���� & "", !�������� & "", !������λ & "", strƵ��, Format(!��ʼʱ�� & "", "yyyy-MM-dd"), _
                            "", !�÷� & "", !���ID & "", IIf(!ҽ����Ч & "" = "0", "0", "1"), str����ҽ��)
                            
                            lngCount = lngCount + 1
                            .MoveNext
                        Next
                    End With
                End If
            End If
            '�޿�����ҩƷ
            If (lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)

    '��ȡҽ�������,����д��ʾ��
    '-------------------------------------------------------------
    If lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3 Then
        arrSQL = Array()
        '����ֵ˳��0-����,1-�Ƶ�,2-���,3-�ڵ�,4-�ȵ�
        '��ʾ��˳��0-����,1-�Ƶ�,4-�ȵ�,2-���,3-�ڵ�(��ΪPASS������ԭ��)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                End If
                If blnDo Then
                    If glngModel = PM_����༭ Then
                        strҽ��ID = .RowData(i)
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    End If
                    k = PassGetWarn(strҽ��ID)
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        
                        '���þ�ʾ��
                        If k >= 0 And k <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(k + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If glngModel = PM_����༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            '��¼�½���ҩƷ K=3����ڵ� �� ֻ���δУ��ҽ�����н���ҩƷ˵��ԭ��ı��,�Ѿ�У�Է��͵�ҽ��������
                            If k = 3 And Not rsOut Is Nothing Then
                                rsOut.Filter = "ҽ��ID = " & strҽ��ID & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        ElseIf PM_����ҽ���嵥 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '��ҩ�䷽
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                            lngLight = -1 '��ʼ��
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                    
                    '��¼��߼���ʾֵ
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                Else
                    If glngModel = PM_����༭ Then
                        '��ҩ��ʾ�Ƶ�������
                        If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                            '���þ�ʾ��
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                            End If
                            
                            If glngModel = PM_����༭ Then
                                '���������仯,�Ա��������ݿ�
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                    blnNoSave = True    '���Ϊδ����
                                End If
                                '��¼�½���ҩƷ K=3����ڵ�
                                If lngLight = 3 And Not rsOut Is Nothing Then
                                    rsOut.Filter = "ҽ��ID = " & lng��ҩ��ID & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                End If
                            End If
                            lng��ҩ��ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
            If glngModel = PM_����ҽ���嵥 And Not rs��ҩ Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '��ҩ����
                    If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        lngLight = -1
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs��ҩ.Filter = "���ID=" & strҽ��ID
                        
                        For j = 1 To rs��ҩ.RecordCount
                            k = PassGetWarn(rs��ҩ!id & "")
                            '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If arrLevel(k) > arrLevel(lngLight) Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs��ҩ.MoveNext
                        Next
                        
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        '��ʾ�Ƹ��µ����ݿ�
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                        End If
                            
                        '��¼��߼���ʾֵ
                        If lngLight >= 0 Then
                            If lngMaxWarn >= 0 Then
                                If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                    lngMaxWarn = lngLight
                                End If
                            Else
                                lngMaxWarn = lngLight
                            End If
                        End If
                    End If
                Next
            End If
        End With
        
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
        
    End If
    '���������
    OutAdviceCheckWarn_MK = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function OperateAdviceCheckWarn_MK(ByVal lngCmd As Long, ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���,��ʿվУ��ʱ����
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����к�:lngCmd=0ʱ��Ҫ,�ಡ����������ʱ��Ҫ��ǰ������
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String
    Dim strҩƷID As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strSQL As String, i As Long, k As Long
    Dim strCurrentDate As String

    OperateAdviceCheckWarn_MK = -1
    If Not (lngRow >= gobjAdvice.FixedRows) Then Exit Function    '����Ҫȷ������������

    On Error GoTo errH
    Screen.MousePointer = 11
    If gstrVersion = "3.0" Then
    '����3.0
    
        '����PASS����״̬
        '-------------------------------------------------------------
        If PassGetState("PassEnable") = 0 Then
            MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
            Screen.MousePointer = 0: Exit Function
        End If
    
        '114036ͬһ�����˶�����ʱ������Ϣÿ�ζ�Ҫ����
        '-------------------------------------------------------------
        lng����ID = Val(gobjAdvice.TextMatrix(lngRow, gobjCOL.intCOL����ID))
        lng��ҳID = Val(gobjAdvice.TextMatrix(lngRow, gobjCOL.intCOL��ҳID))
      
        strSQL = _
        " Select NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,A.��������,B.���,B.����,B.��Ժ����,B.��Ժ����," & _
                 " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                 " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                 " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                 " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng����ID, lng��ҳID)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

        Call PassSetPatientInfo(lng����ID, lng��ҳID, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), rsTmp!���� & "", rsTmp!��� & "", _
                                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), _
                                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))

        '���˲��˹���ʷ
        '-------------------------------------------------------
        Set rsTmp = Get���˹�����¼(lng����ID, lng��ҳID)

        For i = 1 To rsTmp.RecordCount
            Call PassSetAllergenInfo(i, rsTmp!ҩ��ID & "", rsTmp!ҩ���� & "", "DrugName", "")
            rsTmp.MoveNext
        Next

        '���˲���״̬
        '------------------------------------------------------------------
        Set rsTmp = Get������ϼ�¼(lng����ID, lng��ҳID, "2,12")
        strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

        For i = 1 To rsTmp.RecordCount
            Call PassSetMedCond(i & "", rsTmp!���� & "", rsTmp!���� & "", "User", strCurrentDate, strCurrentDate)
            rsTmp.MoveNext
        Next
     
    
        'PASS�Զ���˵����
        '-------------------------------------------------------------
        If lngCmd = 0 Then
            With gobjAdvice
                If Val(.TextMatrix(lngRow, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                        
                    'ȡҩƷ��ҩ;��(��ǰ�ɼ��в������в�ҩ)
                    str�÷� = .TextMatrix(lngRow, gobjCOL.intCOL�÷�)
                    
                    'ҩƷ����ҽ����Ʒ���´�,������ҩƷID
                    If Val(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                        strҩƷID = GetDrugID(.TextMatrix(lngRow, gobjCOL.intCOL������ĿID))
                    Else
                        strҩƷID = .TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)
                    End If
                    
                    '�����ѯҩƷ��Ϣ
                    Call PassSetQueryDrug(strҩƷID, strҩƷ, .TextMatrix(lngRow, gobjCOL.intCOL������λ), str�÷�)
                    
                    OperateAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
                End If
            End With
            Screen.MousePointer = 0: Exit Function
        End If
    
        'ִ����Ӧ������
        '-------------------------------------------------------------
        Call PassDoCommand(lngCmd)
    ElseIf gstrVersion = "4.0" Then
    '����4.0
        With gobjAdvice
            Select Case lngCmd
            
            Case MK4_���PASS�˵�״̬
               
                If Val(.TextMatrix(lngRow, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                    OperateAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
                End If
                Screen.MousePointer = 0: Exit Function
            Case 1
            
            End Select
       End With
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OutAdviceCheckWarn_DT() As Boolean
'���ܣ����ô�ͨ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim strҩƷ As String, str��ҩ;�� As String, strƵ�ʱ��� As String, strXML As String
    Dim arrDiagName(1 To 3) As String, arrDiagCode(1 To 3) As String
    Dim strRetXML As String
    Dim str���� As String
    Dim str������λ As String
    Dim lngҽ��ID As Long
    
    On Error GoTo errH

    curDate = zlDatabase.Currentdate
    With xmlbase
        .dDoctCode = UserInfo.�û���
        .dDoctName = UserInfo.����
        .dDoctType = UserInfo.רҵ����ְ��
        .dDeptCode = UserInfo.����ID
        .dDeptName = UserInfo.������
        .dInHosCode = ""
        .dBedNo = ""
        .mPresDate = curDate
        .pCaseID = gobjPati.lng����ID
        .pOutID = gobjPati.str�Һŵ�
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(gobjPati.dat��������, vbNull)
        .pPatiName = gobjPati.str����
        .pSex = gobjPati.str�Ա�
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
        
        '* ����Դ
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, 0)
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!ҩ��ID
            .pCaseName1 = rsTmp!ҩ����
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!ҩ��ID
                .pCaseName2 = rsTmp!ҩ����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!ҩ��ID
                    .pCaseName3 = rsTmp!ҩ����
                End If
            End If
        End If
        
        '* �����Ϣ
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        If glngModel = PM_����༭ Then
            k = 1
            If Not gobjDiags Is Nothing Then
                With gobjDiags
                    For i = 1 To .Count
                        If .Item(i).str������� <> "" Then
                            arrDiagCode(i) = IIf(.Item(i).str�������� <> "", .Item(i).str��������, .Item(i).str��ϱ���)
                            arrDiagName(i) = .Item(i).str�������
                            If k = 3 Then Exit For
                            k = k + 1
                        End If
                    Next
                End With
            End If
            .pDiagnose1 = arrDiagCode(1)
            .pDiagnose2 = arrDiagCode(2)
            .pDiagnose3 = arrDiagCode(3)
            .pDiagnoseName1 = arrDiagName(1)
            .pDiagnoseName2 = arrDiagName(2)
            .pDiagnoseName3 = arrDiagName(3)
        ElseIf glngModel = PM_����ҽ���嵥 Then
            Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng�Һ�ID, "1")
            If rsTmp.RecordCount > 0 Then
                .pDiagnose1 = "" & rsTmp!����
                .pDiagnoseName1 = "" & rsTmp!����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose2 = "" & rsTmp!����
                    .pDiagnoseName2 = "" & rsTmp!����
                    rsTmp.MoveNext
                    If Not rsTmp.EOF Then
                        .pDiagnose3 = "" & rsTmp!����
                        .pDiagnoseName3 = "" & rsTmp!����
                    End If
                End If
            End If
        End If
        
        '* ������״̬
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        strTmp = Get���˲��������(gobjPati.lng����ID, 0)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_����༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            ElseIf glngModel = PM_����ҽ���嵥 Then
                blnDo = Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If

            If blnDo Then
                'ȡҩƷ����
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                    strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                Else
                    strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                End If
                
                'ȡҩƷ��ҩ;��
                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str��ҩ;�� = "" 'һ����ҩ���ظ�ȡ
                If str��ҩ;�� = "" Then
                    If glngModel = PM_����༭ Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                        If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                    ElseIf glngModel = PM_����ҽ���嵥 Then
                        str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")  '������
                    End If
                End If
                
                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
                
                If glngModel = PM_����༭ Then
                    str���� = StrToXML(.TextMatrix(i, gobjCOL.intCOL����))
                    str������λ = StrToXML(.TextMatrix(i, gobjCOL.intCOL������λ))
                    lngҽ��ID = .RowData(i)
                ElseIf glngModel = PM_����ҽ���嵥 Then
                    lngҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    str���� = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL����)))
                    If Mid(str����, 1, 2) = "0." Then
                        str���� = "0" & Val(str����)
                    Else
                        str���� = Val(str����)
                    End If
                    str������λ = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL����)))
                    If Mid(str������λ, 1, 2) = "0." Then '������С��������⴦��
                        str������λ = Replace(str������λ, Format(Val(str������λ) & "", "0.####"), "") '�����嵥�������������� & ��������λ����
                    Else
                        str������λ = Replace(str������λ, Val(str������λ) & "", "")    '
                    End If
                End If
                
                xmlpre.PresID = lngҽ��ID
                xmlpre.PresType = "mz"
                xmlpre.Current = 1
                xmlpre.GeneralName = StrToXML(Sys.RowValue("������ĿĿ¼", Val(.TextMatrix(i, gobjCOL.intCOL������ĿID)), "����"))
                xmlpre.HosMediCode = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                xmlpre.MediName = StrToXML(strҩƷ)
                
                xmlpre.DCL = str����
                xmlpre.PCDM = StrToXML(strƵ�ʱ���)
                xmlpre.Days = StrToXML(.TextMatrix(i, gobjCOL.intCOL����))
                
                xmlpre.Unit = str������λ
                xmlpre.GYTJ = str��ҩ;��
                xmlpre.GroupNum = Val(.TextMatrix(i, gobjCOL.intCOL���ID))
                
                strXML = MakePresXML(xmlpre, 0)
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = strXML
            End If
        Next
    End With
    
        
    OutAdviceCheckWarn_DT = True
    If UBound(arrTmp) >= 0 Then
        strXML = MakeXML(xmlbase, arrTmp, 0)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strXML
        
        If gbytSuperVolume = 0 Then
            strTmp = dtywzxUI2(4, 0, strXML, strRetXML)
            WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
            strRetXML = GetAlertFromXml(strRetXML)
            If InStr(strRetXML, ";CJLJJ;") > 0 Then
                MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڳ�����������ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_DT = False
            End If
            strRetXML = ""
        Else
            strTmp = dtywzxUI(4, 0, strXML)
            WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
        End If
        '
        If glngModel = PM_����༭ Then
            If strTmp = "2" And gbytBlackLamp = 0 Then
                MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_DT = False
            ElseIf strTmp = "1" Or strTmp = "2" And gbytBlackLamp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then OutAdviceCheckWarn_DT = False
            End If
            If OutAdviceCheckWarn_DT Then
                If gbytSuperVolume = 0 Then
                    strTmp = dtywzxUI2(13, 0, strXML, strRetXML)
                    WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
                    strRetXML = GetAlertFromXml(strRetXML)
                    If InStr(strRetXML, ";CJLJJ;") > 0 Then
                        MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڳ�����������ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                        OutAdviceCheckWarn_DT = False
                    End If
                    strRetXML = ""
                Else
                    strTmp = dtywzxUI(13, 0, strXML)
                    WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
                End If
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    OutAdviceCheckWarn_DT = False
End Function

Public Function OutAdviceCheckWarn_TYT(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByRef blnNoSave As Boolean, _
    Optional ByRef rsOut As ADODB.Recordset) As Long
'���ܣ�����̫Ԫͨϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'       0-��ҩ�淶
'       1-��ȡҽ�������,����д��ʾ��
'       2-ҩƷ��ʾ
'       3-ҽҩ֪ʶ�⣬4-ϵͳ����;5-��ȡ��ʾ����
'
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=2ʱ��Ҫ
'���Σ�
'      rsOut-����˵��
'����ֵ��ҽ��������ã���Ҫ�÷���ֵ�ж��Ƿ���ڽ�����ҩ
    Dim strDrugCode As String, strҽ������ As String, str����ҽ�� As String, strDescription As String
    Dim str���� As String, str������λ As String
    Dim strҽ����� As String
    Dim strSQL As String, strOrderInfo As String, strƵ�ʱ��� As String, strƵ�� As String
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim str��ҩ;�� As String, strҩƷ As String, str��ҩ��IDs As String
    Dim str��������ID As String, str���ID As String, strҽ��ID As String
    
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim rsPati As ADODB.Recordset, rs��ҩ As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long, j As Long, lngMaxWarn As Long, lngҽ��ID As Long
    Dim lng��ҩ��ID As Long, lngLight As Long
    
    Dim strTmp As String, strOld As String

    Dim arrTmp As Variant, colAuditResult As Collection
    Dim arrLight(1 To 3) As String
    
    On Error GoTo errH
    Screen.MousePointer = 11

    With gobjAdvice
        Select Case lngCmd

        Case 0  '0-��ҩ�淶

            gobjPass.getPdssPrescription

        Case 1  '1-��ȡҽ�������,����д��ʾ��
        
            If glngModel = PM_����༭ And gbytReason = 1 Then
                Set rsOut = InitAdviceRS(FUN_�������)
            End If
                
            If glngModel = PM_����ҽ���嵥 Then
                Set rsTmp = ReadPatient(gobjPati.lng����ID, gobjPati.str�Һŵ�)
                gobjPati.str���� = rsTmp!���� & ""
                gobjPati.str�Ա� = rsTmp!�Ա� & ""
                gobjPati.dat�������� = CDate(rsTmp!�������� & "")
                gobjPati.lng�Һ�ID = rsTmp!����Id
            End If
            
            '������Ϣ
            With udtPatiOrder
                '���˲�����Ϣ:����ID,����,�Ա� 1-Ů, 0-��, 2-���꣬���˳������ڣ���ʽ YYYY-MM-DD ��Ϊ�գ����
                
                .PatientID = gobjPati.lng����ID & ""
                .Pname = gobjPati.str����
                .pSex = IIf(gobjPati.str�Ա� = "��", "0", IIf(gobjPati.str�Ա� = "Ů", "1", "2"))
                .pdateOfBirth = Format(gobjPati.dat��������, "yyyy-MM-dd")

                '������Ϣ
                strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng�Һ�ID)
                rsTmp.Filter = "��Ŀ����='���'"
                If rsTmp.RecordCount <> 0 Then .pHeight = IIf(Val(rsTmp!��¼���� & "") = 0, "", rsTmp!��¼���� & "")
                rsTmp.Filter = "��Ŀ����='����'"
                If rsTmp.RecordCount <> 0 Then .pWeight = IIf(Val(rsTmp!��¼���� & "") = 0, "", rsTmp!��¼���� & "")

                .PvisitID = gobjPati.dbl��ʶ�� & ""

                '���˲����������
                strTmp = Get���˲��������(gobjPati.lng����ID, 0)
                .isLact = IIf(InStr(strTmp, "������") > 0, "1", "0")    '�Ƿ��飬��Ϊ1����Ϊ0 ��Ϊ��
                .isPregnant = IIf(InStr(strTmp, "�и�") > 0, "1", "0")    '�Ƿ��и�����Ϊ1 ����Ϊ0 ��Ϊ��
                .isLiverWhole = IIf(InStr(strTmp, "�ι����쳣") > 0, "1", "0") '�Ƿ�ι��쳣 1-�쳣��0-���� ��Ϊ��
                .isKidneyWhole = IIf(InStr(strTmp, "�������쳣") > 0, "1", "0") '�Ƿ������쳣 1-�쳣��0-���� ��Ϊ��

                '��¼ҽ����Ϣ
                .DoctDeptID = UserInfo.����ID & ""
                .DoctDeptName = UserInfo.������ & ""
                .DoctID = UserInfo.��� & ""
                .DoctName = UserInfo.���� & ""
                .DoctTitleID = GetDoctorTitleType(UserInfo.רҵ����ְ��)
                .DoctTitleName = IIf(UserInfo.רҵ����ְ�� = "", "����ְ��", UserInfo.רҵ����ְ��)
                .SysFlag = "1"  '2-סԺҽ��վ��1-����ҽ��վ
            End With

            'ҩƷ��Ϣ
            curDate = zlDatabase.Currentdate
            arrTmp = Array()
            With gobjAdvice

                For i = .FixedRows To .Rows - 1
                    If glngModel = PM_����༭ Then
                        blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                                And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                        blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4  '��ʱҽ�����ϲ����
                    ElseIf glngModel = PM_����ҽ���嵥 Then
                        blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                                And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")
                        blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4  '��ʱҽ�����ϲ����
                    End If
                    
                    If blnDo Then
                    
                        If glngModel = PM_����ҽ���嵥 And .TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4" Then
                             '��ȡ��ҩҽ����ID
                            str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                        Else
                            'ȡҩƷ����
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                            Else
                                strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                            End If
    
                            If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then  'һ����ҩ���ظ�ȡ
                                '��ҩ;��
                                If glngModel = PM_����༭ Then
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                    If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                                ElseIf glngModel = PM_����ҽ���嵥 Then
                                    str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")  '������
                                End If
                                'ȡ��ҩƵ��(��/��),��Ϊ������������
                                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
    
                                If str�����λ = "��" Then
                                    strƵ�� = intƵ�ʴ��� & "/" & intƵ�ʼ��
                                ElseIf str�����λ = "��" Then
                                    strƵ�� = intƵ�ʴ��� & "/7"
                                ElseIf str�����λ = "Сʱ" Then
                                    If intƵ�ʼ�� <= 24 Then
                                        strƵ�� = Format(24 / intƵ�ʼ�� * intƵ�ʴ���, "0") & "/1"
                                    Else
                                        strƵ�� = intƵ�ʴ��� & "/" & Format(intƵ�ʼ�� / 24, "0")
                                    End If
                                ElseIf str�����λ = "����" Then
                                    strƵ�� = Format((24 * 60) / intƵ�ʼ�� * intƵ�ʴ���, "0") & "/1"
                                End If
    
                                str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                            End If
                            
                            If glngModel = PM_����༭ Then
                                lngҽ��ID = Val(.RowData(i) & "")
                                str���� = .TextMatrix(i, gobjCOL.intCOL����)
                                str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                                str��������ID = .TextMatrix(i, gobjCOL.intCOL��������ID)
                                strҽ����� = .TextMatrix(i, gobjCOL.intCOL���)
                            ElseIf glngModel = PM_����ҽ���嵥 Then
                                str���� = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL����)))
                                If Mid(str����, 1, 2) = "0." Then
                                    str���� = "0" & Val(str����)
                                Else
                                    str���� = Val(str����)
                                End If
                                
                                str������λ = Trim(.TextMatrix(i, gobjCOL.intCOL����))
                                If Mid(str������λ, 1, 2) = "0." Then '������С��������⴦��
                                    str������λ = Replace(str������λ, Format(Val(str������λ) & "", "0.####"), "") '�����嵥�������������� & ��������λ����
                                Else
                                    str������λ = Replace(str������λ, Val(str������λ) & "", "")    '
                                End If
                                Set rsTmp = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOLID)))
                                str��������ID = rsTmp!��������id & ""
                                strҽ����� = rsTmp!��� & ""
                            End If
                            udtDrug.drugID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)    'his ϵͳ��ҩƷ���벻Ϊ��
                            udtDrug.DrugName = StrToXML(strҩƷ)               'his ϵͳ��ҩƷ���Ʋ�Ϊ��
                            udtDrug.recMainNo = .TextMatrix(i, gobjCOL.intCOL���ID)     'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                            udtDrug.recSubNo = strҽ�����       'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                            udtDrug.dosage = str����      'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
    
                            udtDrug.doseUnits = str������λ    'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
                            udtDrug.administrationID = str��ҩ;��              'his ϵͳ��ҽ��;�����벻Ϊ��
                            udtDrug.performFreqDictID = StrToXML(strƵ�ʱ���)   'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
                            udtDrug.performFreqDictText = strƵ��               'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
    
                            udtDrug.startDateTime = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DDHH: MM: SS ��Ϊ��
                            udtDrug.stopDateTime = ""                           'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
                            
                            udtDrug.doctorDept = str��������ID   'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
                            udtDrug.DoctorID = strҽ������                          'his ϵͳ�Ŀ�ҽ��ҽ������
                            udtDrug.Doctor = str����ҽ��                         'his ϵͳ�Ŀ�ҽ��ҽ������,
                            If glngModel = PM_����༭ Then
                                udtDrug.isNew = IIf(.TextMatrix(i, gobjCOL.intCOLEDIT) = "1", "1", "0")    '����ҽ��ֵΪ1������Ϊ0
                            Else
                                udtDrug.isNew = "0"
                            End If
                            
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = lngҽ��ID
                                    rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                    rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                    rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                                    '��ҩ�䷽  ����˵����������ҩ������
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                        rsOut!������� = .TextMatrix(k, gobjCOL.intCOL�������)
                                        rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                        rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                        rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                        rsOut.Update
                                    End If
                                End If
                            End If
                            
                            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                            arrTmp(UBound(arrTmp)) = udtDrug
                        End If
                    End If
                Next
                
                '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
                If glngModel = PM_����ҽ���嵥 Then
                    If str��ҩ��IDs <> "" Then
                        Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                        With rs��ҩ
                            For i = 1 To .RecordCount
                                If !���ID & "" <> str���ID Then
                                    str����ҽ�� = !����ҽ��
                                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                    strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                                    strƵ�� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
                                    Call GetƵ����Ϣ_����(!Ƶ�� & "", CInt(!Ƶ�ʴ��� & ""), CInt(!Ƶ�ʼ�� & ""), !�����λ & "", 2, strƵ�ʱ���)
                                    str���ID = !���ID
                                End If
                                udtDrug.drugID = !ҩƷID & ""    'his ϵͳ��ҩƷ���벻Ϊ��
                                udtDrug.DrugName = !ҩƷ���� & ""              'his ϵͳ��ҩƷ���Ʋ�Ϊ��
                                udtDrug.recMainNo = !���ID & ""     'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                                udtDrug.recSubNo = !��� & ""       'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                                udtDrug.dosage = !�������� & ""      'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
        
                                udtDrug.doseUnits = !������λ & ""    'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
                                udtDrug.administrationID = !�÷�ID & ""              'his ϵͳ��ҽ��;�����벻Ϊ��
                                udtDrug.performFreqDictID = strƵ�ʱ���    'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
                                udtDrug.performFreqDictText = strƵ��               'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
        
                                udtDrug.startDateTime = Format(!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DDHH: MM: SS ��Ϊ��
                                udtDrug.stopDateTime = ""                           'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
                                
                                udtDrug.doctorDept = !��������id & ""  'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
                                udtDrug.DoctorID = strҽ������                          'his ϵͳ�Ŀ�ҽ��ҽ������
                                udtDrug.Doctor = str����ҽ��                         'his ϵͳ�Ŀ�ҽ��ҽ������,
                        
                                udtDrug.isNew = "0"
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtDrug
                                .MoveNext
                            Next
                        End With
                    End If
                End If
            End With
            If UBound(arrTmp) = -1 Then
                Screen.MousePointer = 0: Exit Function
            End If
            udtPatiOrder.PatDrugs = arrTmp

            '���
            arrTmp = Array()
            If glngModel = PM_����༭ Then
                If Not gobjDiags Is Nothing Then
                    With gobjDiags
                        For i = 1 To .Count
                            If .Item(i).str������� <> "" Then
                                udtPatiDiag.diagnosisID = IIf(.Item(i).str�������� <> "", .Item(i).str��������, .Item(i).str��ϱ���) 'his ϵͳ����ϱ���
                                udtPatiDiag.diagnosisName = .Item(i).str�������           'his ϵͳ���������
                                udtPatiDiag.diagnosisType = "�������"                     'ϵͳ��������ͣ���������ϡ���Ժ��ϵ�
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtPatiDiag
                            End If
                        Next
                    End With
                End If
            Else
                Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng�Һ�ID, "1,11")
                For i = 1 To rsTmp.RecordCount
                    udtPatiDiag.diagnosisID = rsTmp!���� 'his ϵͳ����ϱ���
                    udtPatiDiag.diagnosisName = rsTmp!����          'his ϵͳ���������
                    udtPatiDiag.diagnosisType = "�������"                     'ϵͳ��������ͣ���������ϡ���Ժ��ϵ�
                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                    arrTmp(UBound(arrTmp)) = udtPatiDiag
                    rsTmp.MoveNext
                Next
            End If
            
            udtPatiOrder.PatDiagnoses = arrTmp
            '����
            arrTmp = Array()
            Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, 0)
            For i = 0 To rsTmp.RecordCount - 1
                udtPatiSensitive.patOrderDrugSensitiveID = "0"          '�̶�ֵ
                udtPatiSensitive.drugAllergenID = rsTmp!����Դ���� & ""    'ϵͳ�Ĺ�������
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiSensitive
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDrugSensitives = arrTmp
            '֢״
            arrTmp = Array()
            Set rsTmp = GetPatiSymptom(gobjPati.lng����ID, gobjPati.lng�Һ�ID)
            For i = 0 To rsTmp.RecordCount - 1
                UdtPatiSymptom.symptomID = rsTmp!���� & ""    'his ϵͳ��֢״����
                UdtPatiSymptom.symptomName = rsTmp!���� & ""  'his ϵͳ��֢״����
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = UdtPatiSymptom
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatSymptoms = arrTmp

            strOrderInfo = MakePatientOrderXml(udtPatiOrder)

            'ҽ����Ϣ���ӿڵ���"

            strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

            '���������
            '����ֵ˳����ʾ����1�� ���ɣ�������ʾ��ɫ��ʾ�ƣ���2�� ���ã�������ʾ��ɫ��ʾ��ʾ����3�� ��ʾ��������ʾ��ɫ��ʾ�ƣ�
            'ͼ����ɫfrmIcons.imgpassTYT ��1-�죬2-�ƣ�3-��
            arrLight(1) = "��": arrLight(2) = "��": arrLight(3) = "��"
            lngMaxWarn = 4
            If glngModel = PM_����ҽ���嵥 Then arrTmp = Array()
            If strDescription = "" Then
                MsgBox "ҩ����鹦��δִ�У�����̫Ԫͨ�ӿ������Ƿ�����", vbInformation + vbOKOnly, G_STR_PASS
                Screen.MousePointer = 0: Exit Function
            ElseIf strDescription = "-101" Then
                '-101����ʾ�û����Ժ��Ը÷���ֵ������ҵ����
            Else
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                With gobjAdvice
                    For i = .FixedRows To .Rows - 1
                        If glngModel = PM_����༭ Then
                            blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                                    And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4 And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        Else
                            blnDo = Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                                   And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4 And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        End If
                        
                        If blnDo Then

                            '��ȡ��ʾ��
                            If glngModel = PM_����༭ Then
                                strTmp = .TextMatrix(i, gobjCOL.intCOL���ID) & "_" & .TextMatrix(i, gobjCOL.intCOL���)   '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                                lngҽ��ID = CLng(.RowData(i) & "")
                            Else
                                strҽ����� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOLID)), "���")
                                strTmp = .TextMatrix(i, gobjCOL.intCOL���ID) & "_" & strҽ�����   '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                            End If
                            On Error Resume Next
                            udtAuditResult = colAuditResult(strTmp)
                            If Err.Number > 0 Then
                                strTmp = "δ�ҵ�"
                            End If
                            Err.Clear: On Error GoTo 0
                            If strTmp <> "δ�ҵ�" Then  '�ҵ���˾�ʾ��
                                k = Val(udtAuditResult.alertLevel)
                            Else
                                k = 0
                            End If
                            
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                '���þ�ʾ��
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                If k >= 1 And k <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                End If
                                
                                If glngModel = PM_����༭ Then
                                    '���������仯,�Ա��������ݿ�
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                        blnNoSave = True    '���Ϊδ����
                                    End If
                                ElseIf glngModel = PM_����ҽ���嵥 Then
                                     '��ʾ�Ƹ��µ����ݿ�
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                        ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                        arrTmp(UBound(arrTmp)) = "ZL_����ҽ����¼_�������(" & .TextMatrix(i, gobjCOL.intCOLID) & "," & IIf(k >= 1 And k <= 3, k, "NULL") & ")"
                                    End If
                                End If
                                
                                '��¼�½���ҩƷ K=1�����ɫ��ʾ��
                                If gbytReason = 1 And k = 1 And Not rsOut Is Nothing Then
                                    rsOut.Filter = "ҽ��ID = " & lngҽ��ID & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                End If
                            Else
                                 '��ҩ�䷽
                                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                                    lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                                    lngLight = 4 '��ʼ��
                                End If
                                If k > 0 Then
                                    If lngLight > k Then
                                        lngLight = k
                                    End If
                                End If
                            End If
                            '��¼��߼���ʾֵ
                            If k > 0 Then
                                If lngMaxWarn > k Then
                                    lngMaxWarn = k
                                End If
                            End If
                            
                        Else
                            If glngModel = PM_����༭ Then
                                '��ҩ��ʾ�Ƶ�������
                                If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                    '���þ�ʾ��
                                    If lngLight >= 1 And lngLight <= 3 Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                    Else
                                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                    End If
                                    
                                    If glngModel = PM_����༭ Then
                                        '���������仯,�Ա��������ݿ�
                                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                            .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                            blnNoSave = True    '���Ϊδ����
                                        End If
                                        '��¼�½���ҩƷ K=3����ڵ�
                                        If lngLight = 1 And Not rsOut Is Nothing Then
                                            rsOut.Filter = "ҽ��ID = " & lng��ҩ��ID & " And ״̬ < 3 "
                                            If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                        End If
                                    End If
                                    
                                    lng��ҩ��ID = 0
                                    lngLight = 4
                                End If
                            End If
                        End If
                    Next
                    'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
                    If glngModel = PM_����ҽ���嵥 And Not rs��ҩ Is Nothing Then
                        For i = .FixedRows To .Rows - 1
                            '��ҩ����
                            If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                lngLight = 4
                                strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                                rs��ҩ.Filter = "���ID=" & strҽ��ID
                                
                                For j = 1 To rs��ҩ.RecordCount
                                    strTmp = rs��ҩ!���ID & "_" & rs��ҩ!���  '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                                    On Error Resume Next
                                    udtAuditResult = colAuditResult(strTmp)
                                    If Err.Number > 0 Then
                                        strTmp = "δ�ҵ�"
                                    End If
                                    Err.Clear: On Error GoTo 0
                                    If strTmp <> "δ�ҵ�" Then  '�ҵ���˾�ʾ��
                                        k = Val(udtAuditResult.alertLevel)
                                    Else
                                        k = 0
                                    End If
                                    '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                                    '��¼��߼���ʾֵ
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                    rs��ҩ.MoveNext
                                Next
                                
                                '���þ�ʾ��
                                If lngLight >= 1 And lngLight <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                End If
                                '��ʾ�Ƹ��µ����ݿ�
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                    arrTmp(UBound(arrTmp)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 1 And lngLight <= 3, lngLight, "NULL") & ")"
                                End If

                                '��¼��߼���ʾֵ
                                If lngLight > 0 Then
                                    If lngMaxWarn > lngLight Then
                                        lngMaxWarn = lngLight
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                End With
                '�����ύ,����������
                If glngModel = PM_����ҽ���嵥 Then
                    For i = 0 To UBound(arrTmp)
                        Call zlDatabase.ExecuteProcedure(CStr(arrTmp(i)), "������ҩ���")
                    Next
                End If
            End If
        Case 2    ' 2-ҩƷ��ʾ
            If Val(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                '��ȡ��ѡҽ����ҩƷ����
                strDrugCode = .TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)
                '����ҩƷ��ʾ�ӿ�
                gobjPass.getDrugExplain (strDrugCode)
            Else
                MsgBox "��ǰѡ�е�ҽ�����ǰ�����´��ҩƷҽ����", vbInformation + vbOKOnly, "������ҩ���"
            End If
        Case 3    '3-����ҽҩ֪ʶ��
            '��������ҽҩ֪ʶ��
            gobjPass.accessIFMI ("0")  '����ֵ�̶�Ϊ:"0",�޷���ֵ

        Case 4  '4-ϵͳ����
            gobjPass.sysConfig

        Case 5    '5-��ȡ��ʾ����
            gobjPass.getDrugAlertDetail

        End Select
    End With

    OutAdviceCheckWarn_TYT = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InAdviceCheckWarn_MK4(Optional ByVal bytShow As Byte = 0, Optional ByVal bytSubmit As Byte = 0, _
        Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
        Optional ByRef lngResult As Long = 1) As Long
'���ܣ�����Passϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'
'���Σ�
'      rsOut=����ҩƷ˵��
'      ���أ�blnIsHaveOut=�Ƿ������Ժ��ҩ��ҩƷ
'           lngResult-ҩʦ��Ԥϵͳ 0-��ͨ����1-ͨ��
'���أ�������˷��ص���߼���ʾֵ,Ϊ-1,-2,-3��ʾû�н������
'      ���PASS�˵�ʱ������>=0��ʾ���Ե����˵�
'˵������ҩ��飺�漰�����µ�����(������ִ��)����δֹͣ�ĳ���
'      ��ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim rs��ҩ As ADODB.Recordset

    Dim strҽ��ID As String, str���ID As String, strҽ����Ч As String, strҽ����� As String, strҽ��״̬ As String
    Dim str��������ID As String, str�������� As String, str��������IDTag As String
    Dim strҩƷ���� As String, strҩƷID As String, strƵ�� As String, strPre�÷� As String
    Dim str�÷� As String, str�÷�ID As String, str��Ժ��ҩ As String
    Dim str�������� As String, str������λ As String
    Dim strҽ������ As String, str���� As String
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim str����ҽ�� As String, strҽ������ As String, str����ҽ��Tag As String
    Dim str���� As String, str������λ As String, str��ҩĿ�� As String, str���� As String
    Dim str����ʱ�� As String, str����ʱ�� As String, str��ʼʱ�� As String
    Dim str��ҩ��IDs As String, strGroupIDs As String
    Dim strִ�п���ID As String
    Dim strҽ��IDs As String
    
    Dim str��ʾ As String
    Dim str��ʾֵ As String
    Dim str״̬ As String
    Dim str������ĿID As String
    
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim arrLight(0 To 4) As String
    Dim strCurrentDate As String
    Dim i As Long, k As Long, j As Long, lng��ҩ��ID As Long, lngLight As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Dim blnOK As Boolean, blnDo As Boolean
    
    Dim strAdvicesIds As String
   
    
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    
    lngMaxWarn = -1
    InAdviceCheckWarn_MK4 = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
    With gobjAdvice

        'PASS����һ����ҩ�嵥��¼�������ظ����ã�MDC_AddScreenDrug
        lngCount = 0
        curDate = zlDatabase.Currentdate
        '��ʼ��ҩ����Ϣ
        Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ)
        '������ȡ���
        For i = .FixedRows To .Rows - 1
            If InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                str������ĿID = str������ĿID & "," & .TextMatrix(i, gobjCOL.intCOL������ĿID)
            End If
        Next
        If str������ĿID <> "" Then
            Set rs��� = GetDrugID(str������ĿID) 'һ����¼ҲҪ����뷵�ؼ�¼����
        End If
                
        '�����˽���ҩƷ˵������;����ΪסԺ�༭;��鹦��
        If glngModel = PM_סԺ�༭ And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_�������)
        End If
        
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_סԺ�༭ Then
                'סԺ�༭�������ҽ��ʱ�Ѿ����ε�����ҽ����ֹͣ��ȷ��ֹͣ�ĳ���
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            Else
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")
                
                If blnDo Then
                    'һ����ҩ��ֻ��������ʾ��Ч,�����в�������vsAdvice_DrawCell��
                    'һ����ҩ����Чȡ������Ч
                    If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                        strҽ����Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                    Else
                        strҽ����Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                    End If
                    '1-����ҽ����7�������ϵģ�,
                    '2-����δͣ�õĳ���ҽ��(1-�¿�2-����3-У��5-������,6-����ͣ,7-������;��8-ֹͣ,9-ȷ��ֹͣ��ֻ��ֹͣ���ڴ��ڵ������� ),
                    '3-������ʱҽ��
                    str״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                    str����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-mm-dd")
                    blnDo = blnDo And (str״̬ = "4" Or _
                        (strҽ����Ч = "����" And (InStr(",8,9,", str״̬) > 0 And str����ʱ�� > Format(curDate, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str״̬) > 0) Or _
                        strҽ����Ч = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                End If
            End If
            
            If blnDo Then
                '��ȡ��ҩҽ����ID
                If glngModel = PM_סԺҽ���嵥 And (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                    str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                    If glngModel = PM_סԺ�༭ Then
                        '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                        str��Ժ��ҩ = ""
                        If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                            blnIsHaveOut = True: str��Ժ��ҩ = "��Ժ��ҩ"
                        End If
    
                         'ȡҩƷ��ҩ;������ҩ�÷�
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ
                        If str�÷� = "" Then
                            str���� = ""
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                                Else
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    If InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "��/����") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "����/Сʱ") > 0 Then
                                        str���� = .TextMatrix(k, gobjCOL.intcolҽ������)
                                    End If
                                End If
                                str�÷�ID = .TextMatrix(k, gobjCOL.intCOL������ĿID)
                            End If
                        End If
                    Else
                        'ȡҩƷ��ҩ;������ҩ�÷�
                        If Trim(.TextMatrix(i, gobjCOL.intCOL�÷�)) = "" Then
                            str�÷� = strPre�÷�
                        Else
                            str�÷� = Split(.TextMatrix(i, gobjCOL.intCOL�÷�), ",")(0)
                        End If
                        
                        strPre�÷� = str�÷�
                    End If
                    
                    'ȡ��ҩƵ��(��/��),��Ϊ������������
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then strƵ�� = ""    'һ����ҩ���ظ�ȡ
                    If strƵ�� = "" Then
                        strƵ�� = .TextMatrix(i, gobjCOL.intCOLƵ��)
                        
                        str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                        If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                        
                        If str����ҽ��Tag <> str����ҽ�� And str����ҽ�� <> "" Then
                            strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                            str����ҽ��Tag = str����ҽ��
                        End If
                       
                    End If
                    
                    '����ҽ����Ʒ���´�ʱ,����ȡһ��ҩƷId
                    If Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                        rs���.Filter = "ҩ��ID =" & .TextMatrix(i, gobjCOL.intCOL������ĿID)
                        If Not rs���.EOF Then strҩƷID = rs���!ҩƷID & ""
                    Else
                        strҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    End If
                    '������������
                    str��������ID = .TextMatrix(i, gobjCOL.intCOL��������ID)
                    If .TextMatrix(i, gobjCOL.intCOL��������ID) <> str��������IDTag And Val(str��������ID) <> 0 Then
                        str�������� = Sys.RowValue("���ű�", Val(.TextMatrix(i, gobjCOL.intCOL��������ID)), "����")
                        str��������IDTag = .TextMatrix(i, gobjCOL.intCOL��������ID)
                    End If
                    
                    If glngModel = PM_סԺ�༭ Then
                        strҽ��ID = .RowData(i)
                        strҽ����Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                        
                        str�������� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                        
                        str���� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intcol������λ)
                        str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
                        strִ�п���ID = .TextMatrix(i, gobjCOL.intColִ�п���ID)
                        If strҽ����Ч = "1" Then
                            str����ʱ�� = str��ʼʱ��
                        End If
                        
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                '��ҩ,�г�ҩ
                                rsOut.AddNew
                                rsOut!ҽ��ID = CLng(strҽ��ID)
                                rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            '��ҩ�䷽
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                    rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                    rsOut.Update
                                End If
                            End If
                        End If
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        strҽ��IDs = strҽ��IDs & "," & strҽ��ID
                        str�������� = Val(.TextMatrix(i, gobjCOL.intCOL����))
                        str������λ = .TextMatrix(i, gobjCOL.intCOL����)
             
                        str�������� = FormatEx(str��������, 5)
                        str������λ = Replace(str������λ, str��������, "")
                        
                        str���� = Val(.TextMatrix(i, gobjCOL.intCOL����))
                        str������λ = .TextMatrix(i, gobjCOL.intCOL����)
                        str���� = FormatEx(str����, 5)
                        str������λ = Replace(str������λ, str����, "")
                        
                        
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
                        str����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
                        str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
                        strִ�п���ID = ""
                        If strҽ����Ч = "����" Then
                            str����ʱ�� = str����ʱ��
                        End If
                    End If
                    
                    If str����ʱ�� & "" = "" Then str����ʱ�� = " "
                    strҽ����Ч = IIf(strҽ����Ч = "����", 1, 0)
                    str���ID = .TextMatrix(i, gobjCOL.intCOL���ID)
                    strҽ������ = .TextMatrix(i, gobjCOL.intcolҽ������)
                    str��ҩĿ�� = .TextMatrix(i, gobjCOL.intcol��ҩĿ��)
                    
                    If str��ҩĿ�� = "1" Then
                        str��ҩĿ�� = "3"
                    ElseIf str��ҩĿ�� = "2" Then
                        str��ҩĿ�� = "4"
                    Else
                        str��ҩĿ�� = "0"
                    End If
                    '
                    strҽ��״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                    '"0"-���ã�Ĭ�ϣ���"1"-�����ϣ�"2"-��ͣ����"3"-��Ժ��ҩ������ϵͳ���ò�����飩
                    
                    If glngModel = PM_סԺ�༭ Then
                        blnOK = str��Ժ��ҩ = "��Ժ��ҩ"
                    Else
                        blnOK = .TextMatrix(i, gobjCOL.intCOLִ������) = "��Ժ��ҩ"
                        If InStr("," & strGroupIDs & ",", "," & str���ID & ",") = 0 And InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                            strGroupIDs = strGroupIDs & "," & str���ID
                        End If
                    End If
                    If blnOK Then
                        strҽ��״̬ = "3"
                    ElseIf strҽ��״̬ = "4" Then
                        strҽ��״̬ = "1"
                    Else
                        strҽ��״̬ = "0"
                        If glngModel = PM_סԺ�༭ Then
                            strҽ��״̬ = IIf(InStr(",1,2,", "," & .TextMatrix(i, gobjCOL.intCOLEDIT) & ",") > 0, "9", "0")
                        End If
                    End If
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = strҽ��ID
                    rsAdvice!���ID = str���ID
                    rsAdvice!ҽ����Ч = strҽ����Ч
                    rsAdvice!ҽ����� = lngCount + 1
                    rsAdvice!ҽ��״̬ = strҽ��״̬
                    rsAdvice!�������� = str��������
                    rsAdvice!��������id = str��������ID
                    rsAdvice!����ҽ������ = strҽ������
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!ҩƷID = strҩƷID
                    rsAdvice!ҩƷ���� = strҩƷ����
                    rsAdvice!�������� = str��������
                    
                    rsAdvice!������λ = str������λ
                    rsAdvice!Ƶ�� = strƵ��
                    rsAdvice!�÷� = str�÷�
                    rsAdvice!�÷�ID = str�÷�ID
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!��ʼʱ�� = str��ʼʱ��
                    rsAdvice!����ʱ�� = str����ʱ��
                    
                    rsAdvice!���� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!��ҩĿ�� = str��ҩĿ��
                    rsAdvice!ҽ������ = strҽ������
                    rsAdvice!���� = str����
                    rsAdvice!ִ�п���ID = strִ�п���ID
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                    
                    lngCount = lngCount + 1
                End If
            End If
        Next
        strGroupIDs = Mid(strGroupIDs, 2)
        
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        If glngModel = PM_סԺҽ���嵥 Then
            '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
            str���ID = ""
            If str��ҩ��IDs <> "" Then
                Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                With rs��ҩ
                    For i = 1 To .RecordCount
                        If !���ID & "" <> str���ID Then
                            str����ҽ�� = !����ҽ��
                            If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                            str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                            str�������� = Sys.RowValue("���ű�", Val(!��������id & ""), "����")
                            
                            str����ʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                            str����ʱ�� = Format(!��ֹʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                            str��ʼʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                            If !ҽ����Ч & "" = "1" Then
                                str����ʱ�� = str����ʱ��
                            End If
                            
                            If !��ҩĿ�� & "" = "1" Then
                                str��ҩĿ�� = "3"
                            ElseIf !��ҩĿ�� & "" = "2" Then
                                str��ҩĿ�� = "4"
                            Else
                                str��ҩĿ�� = "0"
                            End If
                            
                            If !��ִ������ & "" = "5" And !ִ������ <> "5" Then
                                strҽ��״̬ = "3"
                            ElseIf !ҽ��״̬ & "" = "4" Then
                                strҽ��״̬ = "1"
                            Else
                                strҽ��״̬ = "0"
                            End If
                            str���ID = !���ID & ""
                        End If
                        '----------------------------------------------------------
                        rsAdvice.AddNew
                        rsAdvice!ҽ��ID = !id
                        rsAdvice!���ID = !���ID & ""
                        rsAdvice!ҽ����Ч = !ҽ����Ч & ""
                        rsAdvice!ҽ����� = lngCount + 1
                        rsAdvice!ҽ��״̬ = strҽ��״̬
                        rsAdvice!�������� = str��������
                        rsAdvice!��������id = !��������id & ""
                        rsAdvice!����ҽ������ = strҽ������
                        rsAdvice!����ҽ�� = str����ҽ��
                        rsAdvice!ҩƷID = !ҩƷID & ""
                        rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                        rsAdvice!�������� = !�������� & ""
                        
                        rsAdvice!������λ = !������λ & ""
                        rsAdvice!Ƶ�� = !Ƶ�� & ""
                        rsAdvice!�÷� = !�÷� & ""
                        rsAdvice!�÷�ID = !�÷�ID & ""
                        rsAdvice!����ʱ�� = str����ʱ��
                        rsAdvice!��ʼʱ�� = str��ʼʱ��
                        rsAdvice!����ʱ�� = str����ʱ��
                        
                        rsAdvice!���� = !�ܸ����� & ""
                        rsAdvice!������λ = !������λ & ""
                        rsAdvice!��ҩĿ�� = str��ҩĿ��
                        rsAdvice!ҽ������ = !ҽ������ & ""
                        rsAdvice!ִ�п���ID = !ִ�п���ID & ""
                        rsAdvice.Update
                        '----------------------------------------------------------------------------
                        lngCount = lngCount + 1
                        .MoveNext
                    Next
                End With
            End If

            '�����ݿ���ȡ���ϵ�ҽ��
            ' ֻ��������������
            If strAdvicesIds <> "" Then
                strAdvicesIds = strAdvicesIds & ","
            End If
            strSQL = "Select a.Id As ҽ��id, a.���id, a.��� As ҽ�����, a.�������, a.ҽ����Ч, a.ҽ��״̬, a.������Ŀid, NVL(a.�շ�ϸĿid,f.ҩƷID) as ҩƷID , Decode(a.�������||'','7',a.ҽ������,a.�걾��λ) as ҩƷ����, a.ִ��Ƶ�� as Ƶ��, a.��������, a.�ܸ�����," & vbNewLine & _
                "       a.ִ�б��, a.��ʼִ��ʱ��, a.����ʱ��,a.��ʼִ��ʱ�� as ��ʼʱ��,a.ִ����ֹʱ�� as ����ʱ��, a.ҽ������, a.��������id, e.���� As ��������, a.����ҽ��, a.��ҩĿ��, a.ִ�п���ID, b.���㵥λ as ������λ, c.סԺ��λ as ������λ," & vbNewLine & _
                "       a.��ҩĿ��,a.ҽ������,d.ҽ������ As �÷�, d.������Ŀid As �÷�id " & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ��� C, ����ҽ����¼ D, ���ű� E,ҩƷ��� F " & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2] And a.������Ŀid = b.Id(+) And a.�շ�ϸĿid = c.ҩƷid(+) And a.������ĿID = f.ҩ��ID(+) And Nvl(a.���id, 0) = d.Id(+) And" & vbNewLine & _
                "      a.��������id = e.Id(+) And a.������� In ('5', '6', '7') And Nvl(a.ִ�б��, 0) <> -1 And" & vbNewLine & _
                "      (a.ҽ��״̬ = 4 And a.����ʱ�� Between Trunc(Sysdate) - 7 And Trunc(Sysdate + 1) Or" & vbNewLine & _
                "      (a.ҽ��״̬ In (8, 9) And Trunc(a.ִ����ֹʱ��) > Trunc(Sysdate))) And Not Instr([3], ',' || a.Id || ',') > 0"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng��ҳID, strAdvicesIds)
            For i = 1 To rsTmp.RecordCount
                '----------------------------------------------------------
                strҽ��ID = rsTmp!ҽ��ID & ""
                str���ID = rsTmp!���ID & ""
                strҽ����Ч = rsTmp!ҽ����Ч & ""
                strҽ����� = rsTmp!ҽ����� & ""
                strҽ��״̬ = rsTmp!ҽ��״̬ & ""
                str�������� = rsTmp!�������� & ""
                str��������ID = rsTmp!��������id & ""
                str����ҽ�� = rsTmp!����ҽ�� & ""
                strҩƷID = rsTmp!ҩƷID & ""
                strҩƷ���� = rsTmp!ҩƷ���� & ""
                str�������� = rsTmp!�������� & ""
                
                str������λ = rsTmp!������λ & ""
                strƵ�� = rsTmp!Ƶ�� & ""
                str�÷� = rsTmp!�÷� & ""
                str�÷�ID = rsTmp!�÷�ID & ""
                str����ʱ�� = Format(rsTmp!��ʼʱ�� & "", "yyyy-mm-dd HH:MM:ss")
                str��ʼʱ�� = Format(rsTmp!��ʼʱ�� & "", "yyyy-mm-dd HH:MM:ss")
                str����ʱ�� = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd HH:MM:ss")
                
                str���� = rsTmp!�ܸ����� & ""
                str������λ = rsTmp!������λ & ""
                str��ҩĿ�� = rsTmp!��ҩĿ�� & ""
                strҽ������ = rsTmp!ҽ������ & ""
                strִ�п���ID = rsTmp!ִ�п���ID & ""
                '-------------------------
                
                If str����ҽ��Tag <> str����ҽ�� And str����ҽ�� <> "" Then
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                    str����ҽ��Tag = str����ҽ��
                End If
                
                If strҽ����Ч = "1" Then
                    str����ʱ�� = str��ʼʱ��
                End If
                
                If str����ʱ�� & "" = "" Then str����ʱ�� = " "
                '"0"-���ã�Ĭ�ϣ���"1"-�����ϣ�"2"-��ͣ����"3"-��Ժ��ҩ������ϵͳ���ò�����飩
                If strҽ��״̬ = "4" Then
                    strҽ��״̬ = "1"
                ElseIf strҽ��״̬ = "8" Or strҽ��״̬ = "9" Then
                    strҽ��״̬ = "2"
                Else
                    strҽ��״̬ = "0"
                End If
                
                If str��ҩĿ�� = "1" Then
                    str��ҩĿ�� = "3"
                ElseIf str��ҩĿ�� = "2" Then
                    str��ҩĿ�� = "4"
                Else
                    str��ҩĿ�� = "0"
                End If
                
                '----------------------------------------------------------
                rsAdvice.AddNew
                rsAdvice!ҽ��ID = strҽ��ID
                rsAdvice!���ID = str���ID
                rsAdvice!ҽ����Ч = strҽ����Ч
                rsAdvice!ҽ����� = strҽ�����
                rsAdvice!ҽ��״̬ = strҽ��״̬
                rsAdvice!�������� = str��������
                rsAdvice!��������id = str��������ID
                rsAdvice!����ҽ������ = strҽ������
                rsAdvice!����ҽ�� = str����ҽ��
                rsAdvice!ҩƷID = strҩƷID
                rsAdvice!ҩƷ���� = strҩƷ����
                rsAdvice!�������� = str��������
                
                rsAdvice!������λ = str������λ
                rsAdvice!Ƶ�� = strƵ��
                rsAdvice!�÷� = str�÷�
                rsAdvice!�÷�ID = str�÷�ID
                rsAdvice!����ʱ�� = str����ʱ��
                rsAdvice!��ʼʱ�� = str��ʼʱ��
                rsAdvice!����ʱ�� = str����ʱ��
                
                rsAdvice!���� = str����
                rsAdvice!������λ = str������λ
                rsAdvice!��ҩĿ�� = str��ҩĿ��
                rsAdvice!ҽ������ = strҽ������
                rsAdvice!ִ�п���ID = strִ�п���ID
                rsAdvice.Update
                lngCount = lngCount + 1
                '-------------------------
                rsTmp.MoveNext
            Next
            
            If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
            'ȡִ�п���ID
            If strҽ��IDs <> "" Then
                Set rsTmp = GetDrugInfo_MK4("", strҽ��IDs, gobjPati.lng����ID, gobjPati.lng��ҳID)
                rsAdvice.Filter = ""
                For i = 1 To rsAdvice.RecordCount
                    rsTmp.Filter = "ID=" & rsAdvice!ҽ��ID
                    If Not rsTmp.EOF Then
                        rsAdvice!ִ�п���ID = rsTmp!ִ�п���ID & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
            '��ȡ����
            If strGroupIDs <> "" Then
                Set rsTmp = Get����("," & strGroupIDs & ",")
                For i = 1 To rsTmp.RecordCount
                    rsAdvice.Filter = "���ID =" & rsTmp!id
                    Do While Not rsAdvice.EOF
                        rsAdvice!���� = rsTmp!ҽ������ & ""
                        rsAdvice.MoveNext
                    Loop
                    rsTmp.MoveNext
                Next
                rsAdvice.Filter = ""
            End If
        End If
        '�޿�����ҩƷl
        If lngCount = 0 Then
            Screen.MousePointer = 0: Exit Function
        End If
        
        'PASS��麯��MDC_DoCheck
        Call AdviceCheckWarn_MK4(gobjPati.lng����ID, "", gobjPati.lng��ҳID, bytShow, bytSubmit, rsAdvice, str��ʾ, lngResult)
        
        arrSQL = Array()
        '��ȡҽ�������,����д��ʾ��
        '-------------------------------------------------------------
        '����ֵ˳��0-����,1-�ڵ�,2-���,3-�ȵ�,4-�Ƶ�
        '��ʾ��˳��0-����,4-�Ƶ�,3-�ȵ�,2-���,1-�ڵ�(��ΪPASS������ԭ��)
        arrLevel(0) = 0: arrLevel(1) = 4: arrLevel(2) = 3: arrLevel(3) = 2: arrLevel(4) = 1
        arrLight(0) = "��_4": arrLight(1) = "��_4": arrLight(2) = "��_4": arrLight(3) = "��_4": arrLight(4) = "��_4"
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_סԺ�༭ Then
                'סԺ�༭�������ҽ��ʱ�Ѿ����ε�����ҽ����ֹͣ��ȷ��ֹͣ�ĳ���
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            Else
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")
                
                If blnDo Then
                    'һ����ҩ��ֻ��������ʾ��Ч,�����в�������vsAdvice_DrawCell��
                    'һ����ҩ����Чȡ������Ч
                    If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                        strҽ����Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                    Else
                        strҽ����Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                    End If
                    '1-����ҽ����7�������ϵģ�,
                    '2-����δͣ�õĳ���ҽ��(1-�¿�2-����3-У��5-������,6-����ͣ,7-������;��8-ֹͣ,9-ȷ��ֹͣ��ֻ��ֹͣ���ڴ��ڵ������� ),
                    '3-������ʱҽ��
                    str״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                    str����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-mm-dd")
                    blnDo = blnDo And (str״̬ = "4" Or _
                        (strҽ����Ч = "����" And (InStr(",8,9,", str״̬) > 0 And str����ʱ�� > Format(curDate, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str״̬) > 0) Or _
                        strҽ����Ч = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                End If
            End If
            If blnDo Then
                If glngModel = PM_סԺ�༭ Then
                    strҽ��ID = .RowData(i) & ""
                Else
                    strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                End If
                rsAdvice.Filter = "ҽ��ID='" & strҽ��ID & "'"
                
                If rsAdvice.RecordCount > 0 Then
                    k = CLng(rsAdvice!��ʾ & "")
                Else
                    k = -1
                End If
                
                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                    '��ҩ������ҩ'���þ�ʾ��
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                    End If
                    
                    If glngModel = PM_סԺ�༭ Then
                        '���������仯,�Ա��������ݿ�
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                            blnNoSave = True    '���Ϊδ����
                        End If
                        
                        If Not rsOut Is Nothing And k = 1 Then
                            rsOut.Filter = "ҽ��ID=" & CLng(strҽ��ID) & " And ״̬ < 3 "
                            If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                        End If
                    ElseIf PM_סԺҽ���嵥 = glngModel Then
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                        End If
                    End If
                ElseIf .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                    '��ҩ�䷽
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                        lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                        lngLight = -1 '��ʼ��
                    End If
                    '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                    If k >= 0 Then
                        If lngLight >= 0 Then
                            If arrLevel(k) > arrLevel(lngLight) Then
                                lngLight = k
                            End If
                        Else
                            lngLight = k
                        End If
                    End If
                End If
    
                 '��¼��߼���ʾֵ
                If k >= 0 Then
                    If lngMaxWarn >= 0 Then
                        If arrLevel(k) > arrLevel(lngMaxWarn) Then
                            lngMaxWarn = k
                        End If
                    Else
                        lngMaxWarn = k
                    End If
                End If
            Else
                If glngModel = PM_סԺ�༭ Then
                    If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If glngModel = PM_סԺ�༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            
                            If Not rsOut Is Nothing And lngLight = 1 Then
                                rsOut.Filter = "ҽ��ID=" & lng��ҩ��ID & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        End If
                        lng��ҩ��ID = 0
                        lngLight = -1
                    End If
                End If
            
            End If
        Next
        'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
        If glngModel = PM_סԺҽ���嵥 And Not rs��ҩ Is Nothing Then
            For i = .FixedRows To .Rows - 1
                '��ҩ����
                If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                    lngLight = -1
                    strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    rs��ҩ.Filter = "���ID=" & strҽ��ID
                    
                    For j = 1 To rs��ҩ.RecordCount
                        rsAdvice.Filter = "ҽ��ID='" & rs��ҩ!id & "'"
                        If rsAdvice.RecordCount > 0 Then
                            k = CLng(rsAdvice!��ʾ & "")
                        Else
                            k = -1
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                        rs��ҩ.MoveNext
                    Next
                    
                    '���þ�ʾ��
                    If lngLight >= 0 And lngLight <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                    End If
                    
                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                    End If
                    
                    '��¼��߼���ʾֵ
                    If lngLight >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = lngLight
                            End If
                        Else
                            lngMaxWarn = lngLight
                        End If
                    End If
                End If
            Next
        End If
            
'        '���ڽ������Ҳ�����ҽ��,ͨ��SQLǿ��ˢ��
'        If strAdvicesIds <> "" Then
'            strAdvicesIds = Mid(strAdvicesIds, 2)
'            arrTmp = Split(strAdvicesIds, ",")
'            For i = LBound(arrTmp) To UBound(arrTmp)
'                strҽ��ID = Split(arrTmp(i), ":")(0)
'                str��ʾֵ = Split(arrTmp(i), ":")(1)
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
'            Next
'        End If
'
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
            
    End With

    '���������
    InAdviceCheckWarn_MK4 = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function InAdviceCheckWarn_MK(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, _
    Optional ByVal bytFunc As Byte = 0, Optional ByRef rsOut As ADODB.Recordset) As Long
'���ܣ�����Passϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        1/33-�����Զ����(סԺ/����),2/34-�ύ�Զ����(סԺ/����),3-�ֹ��������
'        6-��ҩ����,12-��ҩ�о�,22-����״̬/����ʷ����(�༭)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0,6ʱ��Ҫ
'      ���أ�blnIsHaveOut=�Ƿ������Ժ��ҩ��ҩƷ
'����:
'   rsOut=����˵��
'���أ�������˷��ص���߼���ʾֵ,Ϊ-1,-2,-3��ʾû�н������
'      ���PASS�˵�ʱ������>=0��ʾ���Ե����˵�
'˵������ҩ��飺�漰�����µ�����(������ִ��)����δֹͣ�ĳ���
'      ��ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim rs��ҩ As ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, strƵ�� As String, strPre�÷� As String, str��Ч As String
    Dim strҩƷID As String, str�÷�ID As String, strTmp As String, strType As String
    Dim str��ҩ��IDs As String
    Dim str���ID As String
    Dim lngMaxWarn As Long, strOld As String, lng��ҩ��ID As Long
    Dim strSQL As String, blnDo As Boolean, blnLight As Boolean
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim arrLight(0 To 4) As String
    Dim strCurrentDate As String
    Dim i As Long, k As Long, j As Long, lngLight As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim str����ҽ�� As String, strҽ������ As String, str�������� As String, strҽ��ID As String
    Dim str���� As String, str������λ As String
    Dim strסԺ���� As String, strסԺ��λ As String, str��ҩĿ�� As String, str���� As String
    Dim str����ʱ�� As String, str��ֹʱ�� As String, strִ��ʱ�� As String
    Dim lngBegin As Long, lngEnd As Long, lngGroupMax As Long
    Dim rsAdvice As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim strAdvicesIds As String, strAll As String, strFaceID As String
    Dim strִ��ʱ�䷽�� As String
    Dim strֹͣ�� As String
    
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    
    lngMaxWarn = -1
    InAdviceCheckWarn_MK = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
    '����3.0
    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036ͬһ�����˶�����ʱ������Ϣÿ�ζ�Ҫ����
    '-------------------------------------------------------------
    strSQL = _
    " Select Nvl(B.����,A.����) ����,Nvl(B.�Ա�,A.�Ա�) �Ա�,A.��������,B.���,B.����,B.��Ժ����,B.��Ժ����," & _
             " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
             " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
             " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
             " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng��ҳID)
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

    Call PassSetPatientInfo(gobjPati.lng����ID, gobjPati.lng��ҳID, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), rsTmp!���� & "", rsTmp!��� & "", _
                            rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), _
                            IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))

    '���˲��˹���ʷ
    '-------------------------------------------------------
    Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, gobjPati.lng��ҳID)

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!ҩ��ID & "", rsTmp!ҩ���� & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '���˲���״̬
    '------------------------------------------------------------------
    Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng��ҳID, "2,12")
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

    For i = 1 To rsTmp.RecordCount
        Call PassSetMedCond(i & "", rsTmp!���� & "", rsTmp!���� & "", "User", strCurrentDate, strCurrentDate)
        rsTmp.MoveNext
    Next

    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With gobjAdvice
            If IIf(glngModel = PM_סԺ�༭, .RowData(lngRow) <> 0, True) And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                'ȡҩƷ����
                If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 Then
                    strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҩƷ����)
                Else
                    strҩƷ = .TextMatrix(lngRow, gobjCOL.intCOLҽ������) '��ҩ����
                End If
                
                'ȡҩƷ��ҩ;��(��ǰ�ɼ��в������в�ҩ)
                If glngModel = PM_סԺ�༭ Then
                    str�÷� = ""
                    k = .FindRow(CLng(.TextMatrix(lngRow, gobjCOL.intCOL���ID)), lngRow + 1)
                    If k <> -1 Then str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                Else
                    str�÷� = .TextMatrix(lngRow, gobjCOL.intCOL�÷�)
                    If InStr(str�÷�, ",") > 0 Then str�÷� = Left(str�÷�, InStr(str�÷�, ",") - 1)
                End If
                
                'ҩƷ����ҽ����Ʒ���´�ʱ,�շ�ϸĿIDΪ��,������ҩƷID
                If Val(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                    strҩƷID = GetDrugID(.TextMatrix(lngRow, gobjCOL.intCOL������ĿID))
                Else
                    strҩƷID = .TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)
                End If
                
                '�����ѯҩƷ��Ϣ
                Call PassSetQueryDrug(strҩƷID, strҩƷ, .TextMatrix(lngRow, gobjCOL.intCOL������λ), str�÷�)
                
                '���ò˵�����״̬����zlPASSPopupCommandBars������
                InAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
            ElseIf glngModel = PM_סԺҽ���嵥 And .TextMatrix(lngRow, gobjCOL.intCOL�������) = "E" And .TextMatrix(lngRow, gobjCOL.intCol��������) = "4" Then
                 InAdviceCheckWarn_MK = 1    '��ʾ���Ե����˵�
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    If glngModel = PM_סԺ�༭ Then
        '����ʷ/����״̬�༭
        '-------------------------------------------------------------
        If lngCmd = 22 Then
            'lngCmd=21-ֻ��,22-��ǿ�Ʊ༭,23-ǿ�Ʊ༭
            If PassDoCommand(lngCmd) = 2 Then
                '�������ֵΪ2��ʾ"����ʷ/����״̬�༭"�������仯����Ҫ�����Զ����
                lngCmd = 2    'תΪ�Զ��������,����ִ��
            Else
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    '�����˽���ҩƷ˵������  �ҳ���ΪסԺ�༭��鹦��
    If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And glngModel = PM_סԺ�༭ And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_�������)
    End If
    
    '���벡��ҽ����Ϣ
    '-------------------------------------------------------------
    With gobjAdvice
        If lngCmd = 6 Then
            If glngModel = PM_סԺ�༭ Then
                strTmp = .RowData(lngRow)
            Else
                strTmp = .TextMatrix(lngRow, gobjCOL.intCOLID)
            End If
            Call PassSetWarnDrug(strTmp)    '��ҩ����(�Ѿ����ҽ��Ψһ��)
        Else
            '��ҩ��˻���ҩ�о�
            lngCount = 0
            curDate = zlDatabase.Currentdate
            strҩƷ = "": str�÷� = "": strƵ�� = ""
            '�����ȡ���������ҩƷID
            For i = .FixedRows To .Rows - 1
                If InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                    strҩƷ = strҩƷ & "," & .TextMatrix(i, gobjCOL.intCOL������ĿID)
                End If
            Next
            If strҩƷ <> "" Then
                Set rs��� = GetDrugID(strҩƷ) 'һ����¼ҲҪ����뷵�ؼ�¼����
                strҩƷ = ""
            End If
            
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    blnDo = blnDo And (lngCmd = 12 Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                            Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")
                    If blnDo Then
                        'һ����ҩ��ֻ��������ʾ��Ч,�����в�������vsAdvice_DrawCell��
                        'һ����ҩ����Чȡ������Ч
                        If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                            str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                        Else
                            str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                        End If
                        '��ֹͣ�ĳ���ҲҪ����
                        blnDo = (lngCmd = 12 Or .TextMatrix(i, gobjCOL.intCOL״̬) <> "4" And _
                                 (str��Ч = "����" Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                        '���������ϵ�ҽ��,�����������
                    End If
                End If
                
                If blnDo Then
                    If glngModel = PM_סԺҽ���嵥 And (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        '��ȡ��ҩҽ����ID
                        str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        'ȡҩƷ����
                        If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 1 Then
                            strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                        Else
                            strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                        End If
                        
                        If glngModel = PM_סԺ�༭ Then
                            '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                            If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                                blnIsHaveOut = True
                            End If
    
                            'ȡҩƷ��ҩ;������ҩ�÷�
                            If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ
                            If str�÷� = "" Then
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                If k <> -1 Then
                                    If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                        str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                                    Else
                                        str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    End If
                                End If
                            End If
                        Else
                            If Trim(.TextMatrix(i, gobjCOL.intCOL�÷�)) = "" Then
                                str�÷� = strPre�÷�
                            Else
                                str�÷� = Split(.TextMatrix(i, gobjCOL.intCOL�÷�), ",")(0)
                            End If
                            strPre�÷� = str�÷�
                        End If
                        'ȡ��ҩƵ��(��/��),��Ϊ������������
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then strƵ�� = ""    'һ����ҩ���ظ�ȡ
                        If strƵ�� = "" Then
                            If glngModel = PM_סԺ�༭ Then
                                strƵ�� = GetFrequency(.TextMatrix(i, gobjCOL.intCOL�����λ), .TextMatrix(i, gobjCOL.intCOLƵ�ʴ���), .TextMatrix(i, gobjCOL.intCOLƵ�ʼ��))
                            Else
                                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), "")
    
                                strƵ�� = GetFrequency(str�����λ, intƵ�ʴ��� & "", intƵ�ʼ�� & "")
                            End If
                            str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                            If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                            str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                        End If
                        '����ҽ����Ʒ���´�ʱ,���⴫һ��ҩƷID
                        If Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                            rs���.Filter = "ҩ��ID =" & .TextMatrix(i, gobjCOL.intCOL������ĿID)
                            If Not rs���.EOF Then strҩƷID = rs���!ҩƷID & ""
                        Else
                            strҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                        End If
                        '����ҽ����Ϣ
                        If glngModel = PM_סԺ�༭ Then
                            strҽ��ID = CStr(.RowData(i))
                        Else
                            strҽ��ID = CStr(.TextMatrix(i, gobjCOL.intCOLID))
                        End If
                        '������������λ
                        str���� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                        str���� = Replace(str����, str������λ, "")
                        
                        Call PassSetRecipeInfo(strҽ��ID, strҩƷID, strҩƷ, _
                                             str����, str������λ, strƵ��, _
                                              Format(IIf(glngModel = PM_סԺ�༭, .Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), .TextMatrix(i, gobjCOL.intCOL����ʱ��)), "yyyy-MM-dd"), _
                                              Format(IIf(glngModel = PM_סԺ�༭, .Cell(flexcpData, i, gobjCOL.intCOL��ֹʱ��), .TextMatrix(i, gobjCOL.intCOL��ֹʱ��)), "yyyy-MM-dd"), _
                                              str�÷�, .TextMatrix(i, gobjCOL.intCOL���ID), IIf(glngModel = PM_סԺ�༭, IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", 0, 1), IIf(str��Ч = "����", 0, 1)), str����ҽ��)
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                '��ҩ,�г�ҩ
                                rsOut.AddNew
                                rsOut!ҽ��ID = CLng(strҽ��ID)
                                rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            '��ҩ�䷽
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                    rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                    rsOut.Update
                                End If
                            End If
                        End If
                        
                        lngCount = lngCount + 1
                    End If
                End If
            Next
            '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
            If glngModel = PM_סԺҽ���嵥 Then
                If str��ҩ��IDs <> "" Then
                    Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                    With rs��ҩ
                        For i = 1 To .RecordCount
                            If !���ID & "" <> str���ID Then
                                str����ҽ�� = !����ҽ��
                                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                                strƵ�� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
                                str���ID = !���ID & ""
                            End If
                            Call PassSetRecipeInfo(!id, !ҩƷID & "", !ҩƷ���� & "", !�������� & "", !������λ & "", strƵ��, Format(!����ʱ�� & "", "yyyy-MM-dd"), _
                            Format(!ͣ��ʱ�� & "", "yyyy-MM-dd"), !�÷� & "", !���ID & "", IIf(!ҽ����Ч & "" = "0", "0", "1"), str����ҽ��)
                            
                            lngCount = lngCount + 1
                            .MoveNext
                        Next
                    End With
                End If
            End If
            '�޿�����ҩƷ
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)

    '��ȡҽ�������,����д��ʾ��
    '-------------------------------------------------------------
    If lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3 Then
        arrSQL = Array()
        '����ֵ˳��0-����,1-�Ƶ�,2-���,3-�ڵ�,4-�ȵ�
        '��ʾ��˳��0-����,1-�Ƶ�,4-�ȵ�,2-���,3-�ڵ�(��ΪPASS������ԭ��)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 1 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                            Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 1
                
                    If blnDo Then
                        'һ����ҩ��ֻ��������ʾ��Ч,�����в�����vsAdvice_DrawCell��
                        'һ����ҩ����Чȡ������Ч
                        If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                            str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                        Else
                            str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                        End If
                        '��ֹͣ�ĳ���ҲҪ����'���������ϵ�ҽ��,�����������
                        blnDo = .TextMatrix(i, gobjCOL.intCOL״̬) <> "4" And (str��Ч = "����" _
                               Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                    End If
                End If
                
                If blnDo Then
                    If glngModel = PM_סԺ�༭ Then
                        strҽ��ID = .RowData(i) & ""
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If

                    k = PassGetWarn(strҽ��ID)
                    
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 1 Then
                        '��ҩ������ҩ'���þ�ʾ��
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        If k >= 0 And k <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(k + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If glngModel = PM_סԺ�༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            
                            If Not rsOut Is Nothing And k = 3 Then
                                rsOut.Filter = "ҽ��ID=" & CLng(strҽ��ID) & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        ElseIf glngModel = PM_סԺҽ���嵥 Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                        
                    Else
                        '��ҩ�䷽
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                            lngLight = -1 '��ʼ��
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                    '��¼��߼���ʾֵ
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                Else
                    If glngModel = PM_סԺ�༭ Then
                        If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                            '���þ�ʾ��
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                            End If
                            
                            If glngModel = PM_סԺ�༭ Then
                                '���������仯,�Ա��������ݿ�
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                    blnNoSave = True    '���Ϊδ����
                                End If
                                
                                If Not rsOut Is Nothing And lngLight = 3 Then
                                    rsOut.Filter = "ҽ��ID=" & lng��ҩ��ID & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                End If
                            End If

                            lng��ҩ��ID = 0
                            lngLight = -1
                        End If
                        
                    End If
                End If
            Next
            'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
            If glngModel = PM_סԺҽ���嵥 And Not rs��ҩ Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '��ҩ����
                    If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        lngLight = -1
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs��ҩ.Filter = "���ID=" & strҽ��ID
                        
                        For j = 1 To rs��ҩ.RecordCount
                            k = PassGetWarn(rs��ҩ!id & "")
                            '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If arrLevel(k) > arrLevel(lngLight) Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs��ҩ.MoveNext
                        Next
                        
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                        End If
                        
                        '��¼��߼���ʾֵ
                        If lngLight >= 0 Then
                            If lngMaxWarn >= 0 Then
                                If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                    lngMaxWarn = lngLight
                                End If
                            Else
                                lngMaxWarn = lngLight
                            End If
                        End If
                    End If
                Next
            End If
        End With
        
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
    End If
    '���������
    InAdviceCheckWarn_MK = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InAdviceCheckWarn_DT() As Boolean
'���ܣ����ô�ͨ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim strҩƷ As String, str��ҩ;�� As String, strƵ�ʱ��� As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '�ж��Ƿ����Ժ��ִ�е�ҩƷ

    Set rsPati = GetPatiInfo(gobjPati.lng����ID, gobjPati.lng��ҳID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    curDate = zlDatabase.Currentdate
    With xmlbase
        .dDoctCode = UserInfo.�û���
        .dDoctName = UserInfo.����
        .dDoctType = UserInfo.רҵ����ְ��
        .dDeptCode = UserInfo.����ID
        .dDeptName = UserInfo.������
        .dInHosCode = rsPati!סԺ�� & ""
        .dBedNo = "" & rsPati!��ǰ����
        .mPresDate = curDate
        .pCaseID = gobjPati.lng����ID
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(rsPati!��������, vbNull)
        .pPatiName = rsPati!����
        .pSex = rsPati!�Ա�
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
        
        '* ����Դ
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, gobjPati.lng��ҳID)
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!ҩ��ID
            .pCaseName1 = rsTmp!ҩ����
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!ҩ��ID
                .pCaseName2 = rsTmp!ҩ����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!ҩ��ID
                    .pCaseName3 = rsTmp!ҩ����
                End If
            End If
        End If
        
        '* �����Ϣ
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng��ҳID, "2")
        If rsTmp.RecordCount > 0 Then
            .pDiagnose1 = "" & rsTmp!����
            .pDiagnoseName1 = "" & rsTmp!����
            rsTmp.MoveNext
            If Not rsTmp.EOF Then
                .pDiagnose2 = "" & rsTmp!����
                .pDiagnoseName2 = "" & rsTmp!����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose3 = "" & rsTmp!����
                    .pDiagnoseName3 = "" & rsTmp!����
                End If
            End If
        End If
        
        '* ������״̬
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        strTmp = Get���˲��������(gobjPati.lng����ID, gobjPati.lng��ҳID)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
           If glngModel = PM_סԺ�༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                        Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            ElseIf glngModel = PM_סԺҽ���嵥 Then
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                '���������ϵ�ҽ��,ֹͣ��ȷ��ֹͣ�ĳ���;�������������
                If blnDo Then
                    blnDo = .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                            Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                            And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4
                End If
            End If
            
            If blnDo Then
                'ȡҩƷ����
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                    strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                Else
                    strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                End If

                'ȡҩƷ��ҩ;��
                If glngModel = PM_סԺ�༭ Then
                    '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                    If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                        blnIsHaveOut = True
                    End If
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str��ҩ;�� = "" 'һ����ҩ���ظ�ȡ
                    If str��ҩ;�� = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                        If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                    End If
                ElseIf glngModel = PM_סԺҽ���嵥 Then
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then  'һ����ҩ���ظ�ȡ
                        str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")  '������
                    End If
                End If
                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
            
                xmlpre.PresID = gobjPati.lng����ID  'û��ҽ��ID������ID
                xmlpre.PresType = IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", "L", "T")
                xmlpre.GeneralName = StrToXML(Sys.RowValue("������ĿĿ¼", Val(.TextMatrix(i, gobjCOL.intCOL������ĿID)), "����"))
                xmlpre.HosMediCode = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                xmlpre.MediName = StrToXML(strҩƷ)
                xmlpre.DCL = FormatEx(Val(.TextMatrix(i, gobjCOL.intCOL����)), 5)
                xmlpre.PCDM = StrToXML(strƵ�ʱ���)
                xmlpre.Unit = StrToXML(.TextMatrix(i, gobjCOL.intCOL������λ))
                xmlpre.GYTJ = str��ҩ;��
                xmlpre.GroupNum = Val(.TextMatrix(i, gobjCOL.intCOL���ID))
                xmlpre.BTime = Format(IIf(glngModel = PM_סԺ�༭, .TextMatrix(i, gobjCOL.intCOL��ʼʱ��), .Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��)), "yyyy-MM-dd HH:mm:ss")
                
                xmlpre.ETime = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
                xmlpre.PresTime = Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")
                
                strXML = MakePresXML(xmlpre, 1)
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = strXML
            End If
        Next
    End With
    
        
    InAdviceCheckWarn_DT = True
    If UBound(arrTmp) >= 0 Then
        On Error GoTo errH
        strXML = MakeXML(xmlbase, arrTmp, 1)
        WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strXML
        If gbytSuperVolume = 0 Then
            strTmp = dtywzxUI2(28676, 1, strXML, strRetXML)
            WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
            strRetXML = GetAlertFromXml(strRetXML)
            If InStr(strRetXML, ";CJLJJ;") > 0 Then
                MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڳ�����������ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                InAdviceCheckWarn_DT = False: Exit Function
            End If
            strRetXML = ""
        Else
            strTmp = dtywzxUI(28676, 1, strXML) '��������
            WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
        End If
        
        If glngModel = PM_סԺ�༭ Then
            If strTmp = "2" And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("��ҩ���ϵͳ������Ժ��ִ�е�ҩƷ���ڽ�����ҩ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        InAdviceCheckWarn_DT = False
                    End If
                Else
                    MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                    InAdviceCheckWarn_DT = False: Exit Function
                End If
            ElseIf strTmp = "1" Or strTmp = "2" And gbytBlackLamp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_DT = False
            End If
            If InAdviceCheckWarn_DT Then
                If gbytSuperVolume = 0 Then
                    strTmp = dtywzxUI2(28685, 1, strXML, strRetXML)
                    WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
                    strRetXML = GetAlertFromXml(strRetXML)
                    If InStr(strRetXML, ";CJLJJ;") > 0 Then
                        MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڳ�����������ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                        InAdviceCheckWarn_DT = False
                        Exit Function
                    End If
                    strRetXML = ""
                Else
                    strTmp = dtywzxUI(28685, 1, strXML)
                    WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
                End If
            End If
        Else
            If strTmp = "2" Then
                'MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�������������⣬�������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                InAdviceCheckWarn_DT = False
            ElseIf strTmp = "1" Then
                'If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ������һ�����⣬�Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_DT = False
            End If
            '�����ñ��洦���ӿ�28685
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InAdviceCheckWarn_DT = False
End Function

Public Function InAdviceCheckWarn_TYT(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional blnIsHaveOut As Boolean, _
    Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset) As Long
'���ܣ�����̫Ԫͨϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'       0-��ҩ�淶;1-��ȡҽ�������,����д��ʾ��
'       2-ҩƷ��ʾ
'       3-ҽҩ֪ʶ��;4-ϵͳ����;5-�����ʾ�ƣ���ȡ��ʾ����
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=2ʱ��Ҫ
'����:
'   rsOut-����ҩƷ˵��
'      ���أ�blnIsHaveOut=�Ƿ������Ժ��ҩ��ҩƷ
'����ֵ��ҽ��������ã���Ҫ�÷���ֵ�ж��Ƿ���ڽ�����ҩ
    Dim strDrugCode As String, strҽ������ As String, str����ҽ�� As String, strDescription As String
    Dim strSQL As String, strOrderInfo As String, strƵ�ʱ��� As String, strƵ�� As String
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String, str���ID As String
    Dim str��ҩ;�� As String, strҩƷ As String, str��Ч As String, str��ҩ��IDs As String
    Dim strҽ��ID As String
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, rs��ҩ As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long, j As Long, lngMaxWarn As Long, lng��ҩ��ID As Long
    Dim lngBegin As Long, lngEnd As Long, lngLight As Long
    Dim strTmp As String, strOld As String
    Dim arrTmp As Variant, colAuditResult As Collection
    Dim arrLight(1 To 3) As String

    On Error GoTo errH
    Screen.MousePointer = 11

    With gobjAdvice
        Select Case lngCmd
        Case 0   '0-��ҩ�淶

            gobjPass.getPdssPrescription

        Case 1  '1-��ȡҽ�������,����д��ʾ��
            If gbytReason = 1 And glngModel = PM_סԺ�༭ Then
                Set rsOut = InitAdviceRS(FUN_�������)
            End If
            strSQL = _
            " Select A.סԺ��,Nvl(B.����,A.����) ����,Nvl(B.�Ա�,A.�Ա�) �Ա� ,A.��������,B.���,B.����  " & _
                     " From ������Ϣ A,������ҳ B" & _
                     " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng��ҳID)
            If rsPati.EOF Then Screen.MousePointer = 0: Exit Function

            '������Ϣ
            With udtPatiOrder
                '���˲�����Ϣ:����ID,����,�Ա� 1-Ů, 0-��, 2-���꣬���˳������ڣ���ʽ YYYY-MM-DD ��Ϊ�գ����

                .PatientID = gobjPati.lng����ID & ""
                .Pname = rsPati!���� & ""
                .pSex = IIf(rsPati!�Ա� & "" = "��", "0", IIf(rsPati!�Ա� & "" = "Ů", "1", "2"))
                .pdateOfBirth = Format(rsPati!��������, "yyyy-MM-dd")
                .pHeight = IIf(Val(rsPati!��� & "") = 0, "", rsPati!��� & "")
                .pWeight = IIf(Val(rsPati!���� & "") = 0, "", rsPati!���� & "")
                .PvisitID = rsPati!סԺ�� & ""

                '���˲����������
                strTmp = Get���˲��������(gobjPati.lng����ID, gobjPati.lng��ҳID)
                .isLact = IIf(InStr(strTmp, "������") > 0, "1", "0")    '�Ƿ��飬��Ϊ1����Ϊ0 ��Ϊ��
                .isPregnant = IIf(InStr(strTmp, "�и�") > 0, "1", "0")    '�Ƿ��и�����Ϊ1 ����Ϊ0 ��Ϊ��
                .isLiverWhole = IIf(InStr(strTmp, "�ι����쳣") > 0, "1", "0") '�Ƿ�ι��쳣 1-�쳣��0-���� ��Ϊ��
                .isKidneyWhole = IIf(InStr(strTmp, "�������쳣") > 0, "1", "0") '�Ƿ������쳣 1-�쳣��0-���� ��Ϊ��

                '��¼ҽ����Ϣ
                .DoctDeptID = UserInfo.����ID & ""
                .DoctDeptName = UserInfo.������ & ""
                .DoctID = UserInfo.��� & ""
                .DoctName = UserInfo.���� & ""
                .DoctTitleID = GetDoctorTitleType(UserInfo.רҵ����ְ��)
                .DoctTitleName = IIf(UserInfo.רҵ����ְ�� = "", "����ְ��", UserInfo.רҵ����ְ��)
                .SysFlag = "2"  '2-סԺҽ��վ��1-����ҽ��վ
            End With

            'ҩƷ��Ϣ
            curDate = zlDatabase.Currentdate
            arrTmp = Array()
            With gobjAdvice

                For i = .FixedRows To .Rows - 1
                    If glngModel = PM_סԺҽ���嵥 Then
                        blnDo = (InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")
                        'һ����ҩ����Чȡ������Ч
                        If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                            str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                        Else
                            str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                        End If
                        If blnDo Then
                            '���������ϵģ���ֹͣ�ģ�ȷ��ֹͣ��ҽ��;�������������
                            blnDo = (str��Ч = "����" Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")) _
                                    And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0
                        End If
                    Else
                        'ҩ��״̬�����ϡ�ֹͣ��ȷ��ֹͣ �������
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 And InStr("4,8,9", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0
    
                        If blnDo Then
                            blnDo = .RowData(i) <> 0 And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                                    And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                            blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                                               Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        End If
                    End If
                    
                    If blnDo Then
                        If glngModel = PM_סԺҽ���嵥 And (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                            '��ȡ��ҩҽ����ID
                            str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                        Else
                            'ȡҩƷ����
                            If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                                strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                            Else
                                strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                            End If
    
                            If glngModel = PM_סԺ�༭ Then
                                '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                                If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                                    blnIsHaveOut = True
                                End If
                            End If
                            
                            If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then  'һ����ҩ���ظ�ȡ
                                '��ҩ;��
                                If glngModel = PM_סԺ�༭ Then
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                    If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                                Else
                                    str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID") '������
                                End If
                                'ȡ��ҩƵ��(��/��),��Ϊ������������
                                Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
    
                                If str�����λ = "��" Then
                                    strƵ�� = intƵ�ʴ��� & "/" & intƵ�ʼ��
                                ElseIf str�����λ = "��" Then
                                    strƵ�� = intƵ�ʴ��� & "/7"
                                ElseIf str�����λ = "Сʱ" Then
                                    If intƵ�ʼ�� <= 24 Then
                                        strƵ�� = Format(24 / intƵ�ʼ�� * intƵ�ʴ���, "0") & "/1"
                                    Else
                                        strƵ�� = intƵ�ʴ��� & "/" & Format(intƵ�ʼ�� / 24, "0")
                                    End If
                                ElseIf str�����λ = "����" Then
                                    strƵ�� = Format((24 * 60) / intƵ�ʼ�� * intƵ�ʴ���, "0") & "/1"
                                End If
    
                                str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                            End If
    
                            udtDrug.drugID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)    'his ϵͳ��ҩƷ���벻Ϊ��
                            udtDrug.DrugName = StrToXML(strҩƷ)               'his ϵͳ��ҩƷ���Ʋ�Ϊ��
                            udtDrug.recMainNo = .TextMatrix(i, gobjCOL.intCOL���ID)     'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                            udtDrug.recSubNo = .TextMatrix(i, gobjCOL.intCOL���)        'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                            udtDrug.dosage = Val(.TextMatrix(i, gobjCOL.intCOL����))     'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
    
                            udtDrug.doseUnits = .TextMatrix(i, gobjCOL.intCOL������λ)    'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
                            udtDrug.administrationID = str��ҩ;��              'his ϵͳ��ҽ��;�����벻Ϊ��
                            udtDrug.performFreqDictID = StrToXML(strƵ�ʱ���)   'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
                            udtDrug.performFreqDictText = strƵ��               'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
    
                            udtDrug.startDateTime = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DDHH: MM: SS ��Ϊ��
                            udtDrug.stopDateTime = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
                            udtDrug.doctorDept = .TextMatrix(i, gobjCOL.intCOL��������ID)                 'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
                            udtDrug.DoctorID = strҽ������                          'his ϵͳ�Ŀ�ҽ��ҽ������
                            udtDrug.Doctor = str����ҽ��                         'his ϵͳ�Ŀ�ҽ��ҽ������,
                            If glngModel = PM_סԺҽ���嵥 Then
                                udtDrug.isNew = "0"                             '����ҽ��ֵΪ1������Ϊ0
                            Else
                                udtDrug.isNew = IIf(.TextMatrix(i, gobjCOL.intCOLEDIT) = "1", "1", "0")
                            End If
                            '
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                    '��ҩ,�г�ҩ
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = CLng(CStr(.RowData(i)))
                                    rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                    rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                                '��ҩ�䷽
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!ҽ��ID = CLng(CStr(.RowData(k)))
                                        rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                        rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                        rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                        rsOut.Update
                                    End If
                                End If
                                
                            End If
                            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                            arrTmp(UBound(arrTmp)) = udtDrug
                        End If
                    End If
                Next
                
                '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
                If glngModel = PM_סԺҽ���嵥 Then
                    If str��ҩ��IDs <> "" Then
                        Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                        With rs��ҩ
                            For i = 1 To .RecordCount
                                If !���ID & "" <> str���ID Then
                                    str����ҽ�� = !����ҽ��
                                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                                    strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                                    strƵ�� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
                                    Call GetƵ����Ϣ_����(!Ƶ�� & "", Val(!Ƶ�ʴ��� & ""), Val(!Ƶ�ʼ�� & ""), !�����λ & "", 2, strƵ�ʱ���)
                                    str���ID = !���ID & ""
                                End If
        
                                udtDrug.drugID = !ҩƷID & ""                      'his ϵͳ��ҩƷ���벻Ϊ��
                                udtDrug.DrugName = !ҩƷ���� & ""             'his ϵͳ��ҩƷ���Ʋ�Ϊ��
                                udtDrug.recMainNo = !���ID & ""             'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                                udtDrug.recSubNo = !��� & ""      'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                                udtDrug.dosage = !�������� & ""     'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
        
                                udtDrug.doseUnits = !������λ & ""     'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
                                udtDrug.administrationID = !�÷�ID & ""              'his ϵͳ��ҽ��;�����벻Ϊ��
                                udtDrug.performFreqDictID = StrToXML(strƵ�ʱ���)   'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
                                udtDrug.performFreqDictText = strƵ��               'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
         
                                udtDrug.startDateTime = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")     'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DDHH: MM: SS ��Ϊ��
                                udtDrug.stopDateTime = Format(!��ֹʱ�� & "", "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
                                udtDrug.doctorDept = !��������id & ""               'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
                                udtDrug.DoctorID = strҽ������                          'his ϵͳ�Ŀ�ҽ��ҽ������
                                udtDrug.Doctor = str����ҽ��                         'his ϵͳ�Ŀ�ҽ��ҽ������,
                                udtDrug.isNew = "0"
                                
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtDrug
                                .MoveNext
                            Next
                        End With
                    End If
                End If
            End With
           
            If UBound(arrTmp) = -1 Then
                Screen.MousePointer = 0: Exit Function
            End If
            udtPatiOrder.PatDrugs = arrTmp

            '���
            arrTmp = Array()
            Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng��ҳID, "2,12")   '��ҽסԺ����ҽסԺ

            For i = 0 To rsTmp.RecordCount - 1
                udtPatiDiag.diagnosisID = rsTmp!���� & ""       'his ϵͳ����ϱ���
                udtPatiDiag.diagnosisName = rsTmp!���� & ""     'his ϵͳ���������
                udtPatiDiag.diagnosisType = "��Ժ���"          'ϵͳ��������ͣ���������ϡ���Ժ��ϵ�
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiDiag
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDiagnoses = arrTmp
            '����
            arrTmp = Array()
            Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, gobjPati.lng��ҳID)
            For i = 0 To rsTmp.RecordCount - 1
                udtPatiSensitive.patOrderDrugSensitiveID = "0"          '�̶�ֵ
                udtPatiSensitive.drugAllergenID = rsTmp!����Դ���� & ""    'ϵͳ�Ĺ�������
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiSensitive
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDrugSensitives = arrTmp

            '֢״
            arrTmp = Array()
            Set rsTmp = GetPatiSymptom(gobjPati.lng����ID, gobjPati.lng��ҳID)
            For i = 0 To rsTmp.RecordCount - 1
                UdtPatiSymptom.symptomID = rsTmp!���� & ""              'his ϵͳ��֢״����
                UdtPatiSymptom.symptomName = rsTmp!���� & ""            'his ϵͳ��֢״����

                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = UdtPatiSymptom
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatSymptoms = arrTmp

            strOrderInfo = MakePatientOrderXml(udtPatiOrder)

            'ҽ����Ϣ���ӿڵ���"

            strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

            '���������
            '����ֵ˳����ʾ����(�ߵ���)��1�� ���ɣ�������ʾ��ɫ��ʾ�ƣ���2�� ���ã�������ʾ��ɫ��ʾ��ʾ����3�� ��ʾ��������ʾ��ɫ��ʾ�ƣ�
            lngMaxWarn = 4
            If strDescription = "" Then
                MsgBox "ҩ����鹦��δִ�У�����̫Ԫͨ�ӿ������Ƿ�����", vbInformation + vbOKOnly, G_STR_PASS
                Screen.MousePointer = 0: Exit Function

            ElseIf strDescription = "-101" Then
                '-101����ʾ�û����Ժ��Ը÷���ֵ������ҵ����
            Else
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                
                If glngModel = PM_סԺҽ���嵥 Then arrTmp = Array()
                
                With gobjAdvice
                    '��ȡ��ʾ��
                    'ͼ����ɫfrmIcons.imgpass ��1-�죬2-�ƣ�3-��
                    arrLight(1) = "��": arrLight(2) = "��": arrLight(3) = "��"
                    For i = .FixedRows To .Rows - 1
                        If glngModel = PM_סԺҽ���嵥 Then
                            blnDo = InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                            'һ����ҩ��ֻ��������ʾ��Ч,�����в�����vsAdvice_DrawCell��
                            'һ����ҩ����Чȡ������Ч
                            If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                                str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                            Else
                                str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                            End If
                            '���������ϵģ�ֹͣ�ģ�ȷ��ֹͣ��ҽ��,�����������
                            blnDo = InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 And (str��Ч = "����" _
                                    Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        Else
                            '���ϣ�ֹͣ��ȷ��ֹͣ�Ĳ����
                            blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                                    And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                                    And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                            blnDo = blnDo And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                                    Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        End If

                        If blnDo Then
                            strTmp = .TextMatrix(i, gobjCOL.intCOL���ID) & "_" & .TextMatrix(i, gobjCOL.intCOL���)   '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                            On Error Resume Next
                            udtAuditResult = colAuditResult(strTmp)
                            If Err.Number > 0 Then
                                strTmp = "δ�ҵ�"
                            End If
                            Err.Clear: On Error GoTo 0
                            If strTmp <> "δ�ҵ�" Then  '�ҵ���˾�ʾ��
                                k = Val(udtAuditResult.alertLevel)
                            Else
                                k = 0
                            End If
                            
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                                If Not rsOut Is Nothing And glngModel = PM_סԺ�༭ And k = 1 Then
                                    rsOut.Filter = "ҽ��ID=" & CLng(.RowData(i) & "") & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                End If
                                '���þ�ʾ��
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                If k >= 1 And k <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                End If
    
                                '���������仯,�Ա��������ݿ�
                                If glngModel = PM_סԺҽ���嵥 Then
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                        ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                        arrTmp(UBound(arrTmp)) = "ZL_����ҽ����¼_�������(" & .TextMatrix(i, gobjCOL.intCOLID) & "," & _
                                                                 IIf(CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) = "", "NULL", Val(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ))) & ")"
                                    End If
                                Else
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                        blnNoSave = True    '���Ϊδ����
                                    End If
                                End If
                            ElseIf .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                If glngModel = PM_סԺ�༭ Then
                                    '��ҩ�䷽
                                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                                        lng��ҩ��ID = CLng(.TextMatrix(i, gobjCOL.intCOL���ID))          '��ҩ�䷽��ID
                                        lngLight = 4 '��ʼ��
                                    End If
                                    '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                End If
                            End If
                            
                            '��¼��߼���ʾֵ (��ʾֵԽС��ʾ��Խ��)
                            If k > 0 Then
                                If lngMaxWarn > k Then
                                    lngMaxWarn = k
                                End If
                            End If
                        Else
                            If glngModel = PM_סԺ�༭ Then
                                If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                    '���þ�ʾ��
                                    If lngLight >= 1 And lngLight <= 3 Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                    Else
                                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                    End If
                                    
                                    If glngModel = PM_סԺ�༭ Then
                                        '���������仯,�Ա��������ݿ�
                                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                            .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                            blnNoSave = True    '���Ϊδ����
                                        End If
                                        
                                        If Not rsOut Is Nothing And lngLight = 1 Then
                                            rsOut.Filter = "ҽ��ID=" & lng��ҩ��ID & " And ״̬ < 3 "
                                            If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                                        End If
                                    End If
                                    lng��ҩ��ID = 0
                                    lngLight = 4
                                End If
                            End If
                        End If
                    Next
                    
                    'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
                    If glngModel = PM_סԺҽ���嵥 And Not rs��ҩ Is Nothing Then
                        For i = .FixedRows To .Rows - 1
                            '��ҩ����
                            If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                                lngLight = 4
                                strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                                rs��ҩ.Filter = "���ID=" & strҽ��ID
                                
                                For j = 1 To rs��ҩ.RecordCount
                                    strTmp = rs��ҩ!���ID & "_" & rs��ҩ!���   '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                                    On Error Resume Next
                                    udtAuditResult = colAuditResult(strTmp)
                                    If Err.Number > 0 Then
                                        strTmp = "δ�ҵ�"
                                    End If
                                    Err.Clear: On Error GoTo 0
                                    If strTmp <> "δ�ҵ�" Then  '�ҵ���˾�ʾ��
                                        k = Val(udtAuditResult.alertLevel)
                                    Else
                                        k = 0
                                    End If
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                    
                                    rs��ҩ.MoveNext
                                Next
                                
                                '���þ�ʾ��
                                If lngLight >= 1 And lngLight <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                                End If
                                
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                    arrTmp(UBound(arrTmp)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 1 And lngLight <= 3, lngLight, "NULL") & ")"
                                End If
                                
                                '��¼��߼���ʾֵ
                                If lngLight > 0 Then
                                    If lngMaxWarn > lngLight Then
                                        lngMaxWarn = lngLight
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                End With
                 '�����ύ,����������
                If glngModel = PM_סԺҽ���嵥 Then
                    For i = 0 To UBound(arrTmp)
                        Call zlDatabase.ExecuteProcedure(CStr(arrTmp(i)), "������ҩ���")
                    Next
                End If
            End If

        Case 2    ' 2-ҩƷ��ʾ
            If InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                '��ȡ��ѡҽ����ҩƷ����
                strDrugCode = .TextMatrix(lngRow, gobjCOL.intCOL�շ�ϸĿID)
                '����ҩƷ��ʾ�ӿ�
                gobjPass.getDrugExplain (strDrugCode)
            Else
                MsgBox "��ǰѡ�е�ҽ�����ǰ�����´��ҩƷҽ����", vbInformation + vbOKOnly, "������ҩ���"
            End If
        Case 3    '3-����ҽҩ֪ʶ��
            '��������ҽҩ֪ʶ��
            gobjPass.accessIFMI ("0")  '����ֵ�̶�Ϊ:"0",�޷���ֵ
        Case 4  '4-ϵͳ����
            gobjPass.sysConfig
        Case 5    '5-��ȡ��ʾ����
            gobjPass.getDrugAlertDetail
        End Select
    End With
    InAdviceCheckWarn_TYT = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PassInitialize() As Boolean
'���ܣ���PASS�ӿڽ���ע��ͳ�ʼ����ͬʱ���PASS�ӿ�DLL�Ƿ���ȷ��װ
    Dim lngTmp As Long
    Dim strRet As String
    Dim strCheckMode As String
    Dim strDetails As String
    Dim udtBase As YWS_BASE
    Dim udtDTBSBase As DTBS_BASE
    
    
    On Error GoTo errH
    
    If gbytPass = UNPASS Then Exit Function   '83970 PASSMap���������
    
    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            'PASS���ܺ���ע��(����ͻ���ģʽ)
            If RegisterServer <> 0 Then
                MsgBox "PASS�ͻ���ע��ʧ�ܣ���ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            'PASS��ʼ��
            If PassInit(UserInfo.��� & "/" & UserInfo.�û���, UserInfo.������ & "/" & UserInfo.������, 10) <> 1 Then
                MsgBox "PASSϵͳ��ʼ��ʧ�ܣ���ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
                    
            'PASS�Ƿ���ü���ʼ���ɹ���,�����ٵ���״̬�ӿ� ��������ʦ���
            If PassGetState("PassEnable") = 0 Then
                MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
                Call PassQuit: Exit Function
            End If
            
            'PASSӦ��ģʽ����(��Ĭ��ֵ)
            Call PassSetControlParam(1, 1, 0, 2, 1)    '113198
            
        ElseIf gstrVersion = "4.0" Then
            If glngModel = PM_����༭ Or glngModel = PM_����ҽ���嵥 Then
                strCheckMode = "mz"
            ElseIf glngModel = PM_סԺ�༭ Or glngModel = PM_סԺҽ���嵥 Or glngModel = PM_��ʿУ�� Then
                strCheckMode = "zy"
            ElseIf glngModel = PM_PIVA���� Then
                strCheckMode = "pivas"
            ElseIf glngModel = PM_������ҩ Then
                strCheckMode = "mzyf"
            ElseIf glngModel = PM_���ŷ�ҩ Then
                strCheckMode = "zyyf"
            ElseIf glngModel = PM_���ﴦ����� Then
                strCheckMode = "mzsc"
            ElseIf glngModel = PM_סԺҩ����� Then
                strCheckMode = "zysc"
            End If
            '����δ����ҽԺ����ʱȡվ���
            If gstrHOSCODE = "" Then gstrHOSCODE = IIf(gstrNodeNo = "-", "0", gstrNodeNo)
            lngTmp = MDC_Init(strCheckMode, gstrHOSCODE, UserInfo.���)
            
            If lngTmp <= 0 Then   'YWJ��Ҫ����ģʽ
                MsgBox "PASS4.0ϵͳ��ʼ��ʧ�ܣ���ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��" & vbCrLf & _
                        "����ֵ:" & lngTmp & vbTab & _
                        "������Ϣ:��" & MDC_GetLastError() & "��", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    ElseIf gbytPass = DT Then
        If gstrVersion = "3.0" Then  'CS��
            lngTmp = dtywzxUI(0, 0, "") '��ʼ���ӿ�,�򿪴�ͨ����
            lngTmp = dtywzxUI(768, 0, UserInfo.���) '����ҽ������
        ElseIf gstrVersion = "4.0" Then 'BS��
            With udtDTBSBase
                .strHIS = "HIS"
                .strҽԺ���� = gstrHOSCODE
                .strҽ������ = UserInfo.���
                .strҽ��������� = UserInfo.רҵ��������   'Ҫ�����
                .strҽ���������� = UserInfo.רҵ����ְ��
                .strҽ������ = UserInfo.����
                .str���Ҵ��� = UserInfo.������
                .str�������� = UserInfo.������
            End With
            gstrBaseXml = DTBS_MakeBASEXML(udtDTBSBase)
            strDetails = DTBS_MakeDetailXML(DTBS_��¼, "")
            WriteLog "" & glngModel, "PassInitialize", gstrBaseXml
            WriteLog "" & glngModel, "PassInitialize", strDetails
            lngTmp = CRMS_UI(DTBS_��¼, gstrBaseXml, strDetails, "")
            WriteLog "" & glngModel, "PassInitialize", "��¼�ӿڷ���ֵ:" & lngTmp
        End If
    ElseIf gbytPass = TYT Then
        On Error Resume Next
        Set gobjPass = GetObject(, "Midlayer.ComInterface")
        Err.Clear: On Error GoTo 0
        On Error Resume Next
        If gobjPass Is Nothing Then Set gobjPass = CreateObject("Midlayer.ComInterface")
        Err.Clear: On Error GoTo 0
        If gobjPass Is Nothing Then
            MsgBox "̫Ԫͨ�ӿڳ�ʼ��ʧ��,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
                   vbCrLf & "����ȷ��װ�����ú�����ҩ���ϵͳ֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf gbytPass = YWS Then '���ݱ��� ҩ��ʿ
        On Error Resume Next
        Set gobjPass = GetObject(, "YWSUI.YWS")
        Err.Clear: On Error GoTo 0
        On Error Resume Next
        If gobjPass Is Nothing Then Set gobjPass = CreateObject("YWSUI.YWS")
        If Err.Number <> 0 Then
            MsgBox "ҩ��ʿ��¼ʧ��,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
                   vbCrLf & "����ȷ��װ�����ú�����ҩ���ϵͳ֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
            WriteLog "" & glngModel, "PassInitialize", "��̬����������������:" & Err.Number & "|��������:" & Err.Description
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
        With udtBase
            .strHIS = "HIS"
            .strҽԺ���� = ""
            .strҽ������ = UserInfo.���
            .strҽ��������� = UserInfo.רҵ��������   'Ҫ�����
            .strҽ���������� = UserInfo.רҵ����ְ��
            .strҽ������ = UserInfo.����
            .str���Ҵ��� = UserInfo.������
            .str�������� = UserInfo.������
        End With
        gstrBaseXml = YWS_MakeBASEXML(udtBase)
        WriteLog "" & glngModel, "PassInitialize", gstrBaseXml
        strRet = gobjPass.YWS_UI(YWS_��¼, gstrBaseXml, "", "")
        WriteLog "" & glngModel, "PassInitialize", "��¼�ӿڷ���ֵ:" & lngTmp
    ElseIf gbytPass = HZYY Then
        '���Ե�ַ:"http://118.31.246.211:8080/zlcx/data_detail.action?webHisId=11221&hospitalCode=cqzl123"
    ElseIf gbytPass = ZL Then
        Call ZLShowWindow
    End If
    
    PassInitialize = True
    Exit Function
errH:
    If Err.Number = 53 And InStr(UCase(Err.Description), UCase("ShellRunAs")) > 0 Then
        MsgBox "PASS�ӿ��ļ� ShellRunAs.dll ������,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
            vbCrLf & "����ȷ��װ�����ú�����ҩ���ϵͳ֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("DIFPassDll")) > 0 Then
        MsgBox "PASS�ӿ��ļ� DIFPassDll.dll ������,��������Ϊ����ԭ��" & vbCrLf & _
            vbCrLf & "1.PASS�ͻ����ǵ�һ�ε�¼�����˳�֮�������µ�¼��������ʹ�á�" & _
            vbCrLf & "2.������ҩ���ϵͳδ��ȷ��װ�����ã�����ϸ�����ٵ�¼���ԡ�", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("dtywzxUI")) > 0 Then
        MsgBox "PASS�ӿ��ļ� dtywzxUI.dll ������,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
            vbCrLf & "����ȷ��װ������֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("CRMS_UI")) > 0 Then
        MsgBox "PASS�ӿ��ļ� CRMS_UI.dll ������,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
            vbCrLf & "����ȷ��װ������֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("PASS4Invoke")) > 0 Then
          MsgBox "PASS�ӿ��ļ� PASS4Invoke.dll ������,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
            vbCrLf & "����ȷ��װ������֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Public Function InAdviceCheckWarn_YWS(Optional ByRef blnNoSave As Boolean) As Boolean
'���ܣ����ñ���ҩ��ʿ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt����Դ As YWS_ALLERGIC
    Dim udt��� As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   'ҩƷ��Ϣ
    Dim i As Long, j As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lng��ҩ��ID As Long, lngLight As Long
    Dim arrTmp As Variant, arrSQL As Variant
    Dim curDate As Date
    Dim rsTmp As Recordset, rsRet As ADODB.Recordset, rsPati As ADODB.Recordset, rs��ҩ As ADODB.Recordset
    Dim strTmp As String, str���� As String, str������λ As String
    Dim strҽ��ID As String, strRetXML As String, str���ID As String
    Dim strҩƷ As String, str��ҩ;�� As String, strƵ�ʱ��� As String, strXML As String
    Dim str��Ч As String, strOld As String, str��ҩ��IDs As String
    Dim arrLight(0 To 4) As String
    
    Dim blnDo As Boolean, blnIsHaveOut As Boolean  '�ж��Ƿ����Ժ��ִ�е�ҩƷ

    '������Ϣ
    Set rsPati = GetPatiInfo(gobjPati.lng����ID, gobjPati.lng��ҳID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    With udtPati
        .str���� = rsPati!����
        .str�������� = rsPati!��������
        .str�Ա� = rsPati!�Ա�
        .str���� = rsPati!��� & ""
        .str��� = rsPati!���� & ""
        .str���֤�� = rsPati!���֤�� & ""
        .str�������� = rsPati!סԺ�� & ""
        .str���� = ""
        .str������ = ""
        .str����ʱ�� = ""
        .str����ʱ�䵥λ = ""
        '����Դ
        Set colTmp = New Collection
        Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, gobjPati.lng��ҳID, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!ҩ��ID <> "" Then
                With udt����Դ
                    .str�������� = "5"   '1=ҩ��ʿҩƷ���� 2=ҩ��ʿҩƷ�ɷ�
                    .str����Դ���� = rsTmp!ҩ����
                    .str����Դ���� = "" & rsTmp!ҩ��ID
                End With
                colTmp.Add udt����Դ, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col����Դs = colTmp
        
        '��ϼ�¼
        Set colTmp = New Collection
        Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng��ҳID, "2,12")
        For i = 1 To rsTmp.RecordCount
            With udt���
                If rsTmp!����ID & "" <> "" Then
                     .str������� = "2" '2=IDC10����
                Else
                    .str������� = "0" '0=����
                End If
                .str��ϴ��� = "" & rsTmp!����
                .str������� = "" & rsTmp!����
            End With
            colTmp.Add udt���, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        '������
        strTmp = Get���˲��������(gobjPati.lng����ID, gobjPati.lng��ҳID)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt���
                    .str������� = "1" '1=������״̬
                    .str��ϴ��� = Sys.RowValue("���������", arrTmp(i), "����", "����")
                    .str������� = arrTmp(i)
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col���s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    
    With udtDetail
        .strHISϵͳʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        .str����סԺ��ʶ = "ip"    'סԺ��ʶ

        .str�������� = YWS_GetTreatType(2, gobjPati.lng����ID, gobjPati.lng��ҳID)
        .str����� = rsPati!סԺ�� & ""
        .str��λ�� = "" & rsPati!��ǰ����
        '������Ϣ
        .udt������Ϣ = udtPati
        '������Ϣ
        .udt������Ϣ = udtPres
    End With
    'ҩƷ��Ϣ
    Set colTmp = New Collection
    
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
           If glngModel = PM_סԺ�༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                        Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                        And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4)
                        
            ElseIf glngModel = PM_סԺҽ���嵥 Then
                blnDo = ((InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4"))
                
                If blnDo Then
                    'һ����ҩ��ֻ��������ʾ��Ч,�����в�������vsAdvice_DrawCell��
                    'һ����ҩ����Чȡ������Ч
                    If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                        str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                    Else
                        str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                    End If
                    '���������ϵ�ҽ��,ֹͣ��ȷ��ֹͣ�ĳ���;�������������
                    blnDo = str��Ч = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                            Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                            And .TextMatrix(i, gobjCOL.intCOL״̬) <> "4"
                End If
            End If

            If blnDo Then
                '��ȡ��ҩҽ����ID
                If glngModel = PM_סԺҽ���嵥 And (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                    str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
    
                    'ȡҩƷ��ҩ;��
                    If glngModel = PM_סԺ�༭ Then
                         strҽ��ID = CStr(.RowData(i))
                        '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                        If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                            blnIsHaveOut = True
                        End If
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str��ҩ;�� = "" 'һ����ҩ���ظ�ȡ
                        If str��ҩ;�� = "" Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                        End If
                    ElseIf glngModel = PM_סԺҽ���嵥 Then
                        strҽ��ID = CStr(.TextMatrix(i, gobjCOL.intCOLID))
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then  'һ����ҩ���ظ�ȡ
                            str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")   '������
                        End If
                    End If
                    
                    Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
                    
                    udtMedic.strҩƷ���� = YWS_GetDrugType(.TextMatrix(i, gobjCOL.intCOL�������))
                    udtMedic.str������ = strҽ��ID    '��ҽ��ID
                    udtMedic.Strҽ������ = IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", "L", "T")
                    udtMedic.str����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")          '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
                    udtMedic.str��Ʒ�� = YWS_StrToXML(Sys.RowValue("������ĿĿ¼", Val(.TextMatrix(i, gobjCOL.intCOL������ĿID)), "����"))
                    udtMedic.strҽԺҩƷ���� = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    udtMedic.strҽ������ = ""
                    udtMedic.str��׼�ĺ� = ""
                    udtMedic.str��ҩ��ʼʱ�� = Format(IIf(glngModel = PM_סԺ�༭, .TextMatrix(i, gobjCOL.intCOL��ʼʱ��), .Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��)), "yyyy-MM-dd HH:mm:ss")
                    udtMedic.str��ҩ����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:mm:ss")
                    udtMedic.str��� = ""
                    udtMedic.str��� = .TextMatrix(i, gobjCOL.intCOL���ID)
                    udtMedic.str��ҩ���� = ""
                    '������������λ
                    str���� = .TextMatrix(i, gobjCOL.intCOL����)
                    str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                    str���� = Replace(str����, str������λ, "")
                    
                    udtMedic.str��������λ = str������λ
                    udtMedic.str������ = str����
                    udtMedic.strƵ�δ��� = strƵ�ʱ���
                    udtMedic.str��ҩ;������ = str��ҩ;��
                    udtMedic.str��ҩ���� = .TextMatrix(i, gobjCOL.intCOL����)   'OP ���ﴦ����Ч
                    
                    colTmp.Add udtMedic, "_" & colTmp.Count + 1
                End If
            End If
        Next
        '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
        If glngModel = PM_סԺҽ���嵥 Then
            If str��ҩ��IDs <> "" Then
                Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                With rs��ҩ
                    For i = 1 To .RecordCount
                        If !���ID & "" <> str���ID Then
                            Call GetƵ����Ϣ_����(!Ƶ�� & "", 0, 0, "", IIf(!������� & "" = "7", 2, 1), strƵ�ʱ���)
                            str���ID = !���ID & ""
                        End If
                        udtMedic.strҩƷ���� = YWS_GetDrugType(!������� & "")
                        udtMedic.str������ = !id & ""    '��ҽ��ID
                        udtMedic.Strҽ������ = IIf(!ҽ����Ч & "" = "0", "L", "T")
                        udtMedic.str����ʱ�� = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str��Ʒ�� = !ҩƷ���� & ""
                        udtMedic.strҽԺҩƷ���� = !ҩƷID & ""
                        udtMedic.strҽ������ = ""
                        udtMedic.str��׼�ĺ� = ""
                        udtMedic.str��ҩ��ʼʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str��ҩ����ʱ�� = Format(!��ֹʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str��� = ""
                        udtMedic.str��� = !���ID & ""
                        udtMedic.str��ҩ���� = ""
                        udtMedic.str��������λ = !������λ & ""
                        udtMedic.str������ = !�������� & ""
                        udtMedic.strƵ�δ��� = strƵ�ʱ���
                        udtMedic.str��ҩ;������ = !�÷�ID & ""
                        udtMedic.str��ҩ���� = !���� & ""   'OP ���ﴦ����Ч
                        colTmp.Add udtMedic, "_" & colTmp.Count + 1
                        .MoveNext
                    Next
                End With
            End If
        End If
        With udtPres
            Set .colҩƷ��Ϣ = colTmp
            .str������ = "0"
            .str�������� = ""
            .str����ʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str�Ƿ�ǰ���� = "1" '0 ��ʷ���� 1 ��ǰ����������Ĭ�ϵ�ǰ�������Ժ����䣩
            .Strҽ������ = "L"
        End With
    End With
    
    udtDetail.udt������Ϣ = udtPres
    
    InAdviceCheckWarn_YWS = True
    If udtPres.colҩƷ��Ϣ.Count > 0 Then
        On Error GoTo errH
        
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", strXML
        strTmp = gobjPass.YWS_UI(YWS_��������, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", "����ֵ:" & strTmp & vbCrLf & strRetXML
        Set rsRet = YWS_ReturnRS(strRetXML)
        '���þ�ʾ��
        With gobjAdvice
            arrSQL = Array()
            'ͼƬ�±꣺1-����,2-�Ƶ�,5-�ȵ�,3-���
            '��ʾ��˳��0-����,1-�Ƶ�,2-�ȵ�,3-���
            arrLight(0) = "��": arrLight(1) = "��": arrLight(2) = "��": arrLight(3) = "��"
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                            Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                ElseIf glngModel = PM_סԺҽ���嵥 Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0

                    If blnDo Then
                        'һ����ҩ��ֻ��������ʾ��Ч,�����в�������vsAdvice_DrawCell��
                        'һ����ҩ����Чȡ������Ч
                        If RowInһ����ҩ(i, lngBegin, lngEnd) Then
                            str��Ч = .TextMatrix(lngBegin, gobjCOL.intCOL��Ч)
                        Else
                            str��Ч = .TextMatrix(i, gobjCOL.intCOL��Ч)
                        End If
                        '���������ϵ�ҽ��,ֹͣ��ȷ��ֹͣ�ĳ���;�������������
                        blnDo = str��Ч = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                                Or str��Ч = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL״̬) <> "4"
                    End If
                End If
                If blnDo Then
                    If glngModel = PM_סԺ�༭ Then
                        strҽ��ID = .RowData(i) & ""
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If

                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    If rsRet.RecordCount > 0 Then
                        k = CLng(rsRet!��ʾֵ & "")
                    Else
                        k = 0
                    End If
               
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 1 Then
                        '��ҩ������ҩ'���þ�ʾ��
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        If k >= 0 And k <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If

                        If glngModel = PM_סԺ�༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                        ElseIf glngModel = PM_סԺҽ���嵥 Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 3, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '��ҩ�䷽
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                            lngLight = -1 '��ʼ��
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If k > lngLight Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                Else
                    If glngModel = PM_סԺ�༭ Then
                        If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                            '���þ�ʾ��
                            If lngLight >= 0 And lngLight <= 3 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                            End If

                            If glngModel = PM_סԺ�༭ Then
                                '���������仯,�Ա��������ݿ�
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                    blnNoSave = True    '���Ϊδ����
                                End If
                            End If

                            lng��ҩ��ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
            If glngModel = PM_סԺҽ���嵥 And Not rs��ҩ Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '��ҩ����
                    If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        lngLight = -1
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs��ҩ.Filter = "���ID=" & strҽ��ID

                        For j = 1 To rs��ҩ.RecordCount
                            rsRet.Filter = "ҽ��ID ='" & rs��ҩ!id & "" & "'"
                            If rsRet.RecordCount > 0 Then
                                k = CLng(rsRet!��ʾֵ & "")
                            Else
                                k = 0
                            End If
                            '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If k > lngLight Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs��ҩ.MoveNext
                        Next
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 3, lngLight, "NULL") & ")"
                        End If
                    End If
                Next
            End If
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End With
        
        If glngModel = PM_סԺ�༭ Then
            'YWS_��������=6 ������������ֵ��0��1��2��3��8�ֱ����û�����⣬�������⣬һ�����⣬�������⣬��д�������� ����ǰ����
            '1��2 ���������е���ʾ,����ƿ��Բ鿴��ϸ��Ϣ
            '8-��ʱû���ṩ
            If strTmp = "3" And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("��ҩ���ϵͳ������Ժ��ִ�е�ҩƷ���ڽ�����ҩ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        InAdviceCheckWarn_YWS = False
                        Exit Function
                    End If
                Else
                    MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                    InAdviceCheckWarn_YWS = False
                    Exit Function
                End If
            ElseIf strTmp = "3" And gbytBlackLamp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_YWS = False: Exit Function
            End If
            
            '���洦��
            If strTmp <> "0" Then
                strTmp = gobjPass.YWS_UI(YWS_�ϴ�����, gstrBaseXml, strXML, strRetXML)
                WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", "���洦��:����ֵ" & strTmp
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    InAdviceCheckWarn_YWS = False
End Function

Public Function OutAdviceCheckWarn_YWS(Optional ByRef blnNoSave As Boolean) As Boolean
'���ܣ����ñ���ҩ��ʿ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt����Դ As YWS_ALLERGIC
    Dim udt��� As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   'ҩƷ��Ϣ
    Dim curDate As Date
    Dim rsTmp As Recordset, rsRet As Recordset, rsPati As Recordset, rsPatiInfo As Recordset
    Dim rs��ҩ As ADODB.Recordset
    Dim arrTmp As Variant, arrSQL As Variant
    Dim i As Long, j As Long, k As Long
    Dim lng��ҩ��ID As Long, lngLight As Long
    
    Dim strҩƷ As String, str��ҩ;�� As String, strƵ�ʱ��� As String, strXML As String
    Dim str���� As String, str������λ As String, str��� As String, str���� As String
    Dim strҽ��ID As String, str����ʱ�� As String, strOld As String, str��ҩ��IDs As String
    Dim strRetXML As String, strSQL As String, strTmp As String, str���ID As String
    Dim arrLight(0 To 4) As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    Set rsPati = ReadPatient(gobjPati.lng����ID, gobjPati.str�Һŵ�)
    If rsPati.EOF Then Screen.MousePointer = 0: Exit Function
    
    If glngModel = PM_����ҽ���嵥 Then
        gobjPati.lng�Һ�ID = rsPati!����Id
    End If
    
    '������Ϣ
    strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                    "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                    "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
    Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng����ID, gobjPati.lng�Һ�ID)
    rsPatiInfo.Filter = "��Ŀ����='���'"
    If rsPatiInfo.RecordCount <> 0 Then str��� = NVL(rsPatiInfo!��¼����)
    rsPatiInfo.Filter = "��Ŀ����='����'"
    If rsPatiInfo.RecordCount <> 0 Then str���� = NVL(rsPatiInfo!��¼����)
            
    With udtPati
        .str���� = rsPati!����
        .str�������� = rsPati!�������� & ""
        .str�Ա� = rsPati!�Ա� & ""
        .str���� = str���
        .str��� = str����
        .str���֤�� = rsPati!���֤�� & ""
        .str�������� = gobjPati.lng�Һ�ID
        .str���� = ""
        .str������ = ""
        .str����ʱ�� = ""
        .str����ʱ�䵥λ = ""
        '����Դ
        Set colTmp = New Collection
        Set rsTmp = Get���˹�����¼(gobjPati.lng����ID, 0, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!ҩ��ID <> "" Then
                With udt����Դ
                    .str�������� = "5"   '1=ҩ��ʿҩƷ���� 2=ҩ��ʿҩƷ�ɷ� 5-��hisҩƷ
                    .str����Դ���� = rsTmp!ҩ����
                    .str����Դ���� = "" & rsTmp!ҩ��ID
                End With
                colTmp.Add udt����Դ, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col����Դs = colTmp
        
        '��ϼ�¼
        Set colTmp = New Collection
        If glngModel = PM_����༭ Then
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udt���
                        If gobjDiags.Item(i).str������� <> "" Then
                            If gobjDiags.Item(i).str�������� <> "" Then
                                .str������� = "2" '2=IDC10����
                                .str��ϴ��� = gobjDiags.Item(i).str��������
                            Else
                                .str������� = "0"
                                .str��ϴ��� = gobjDiags.Item(i).str��ϱ���
                            End If
                            .str������� = gobjDiags.Item(i).str�������
                        End If
                    End With
                    colTmp.Add udt���, "_" & colTmp.Count + 1
                Next
            End If
        Else
            Set rsTmp = Get������ϼ�¼(gobjPati.lng����ID, gobjPati.lng�Һ�ID, "1,11")
            For i = 1 To rsTmp.RecordCount
                With udt���
                    If rsTmp!����ID & "" <> "" Then
                         .str������� = "2" '2=IDC10����
                    Else
                        .str������� = "0" '0=����
                    End If
                    .str��ϴ��� = "" & rsTmp!����
                    .str������� = "" & rsTmp!����
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
        End If
        '������
        strTmp = Get���˲��������(gobjPati.lng����ID, 0)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt���
                    .str������� = "1" '1=������״̬
                    .str��ϴ��� = Sys.RowValue("���������", arrTmp(i), "����", "����")
                    .str������� = arrTmp(i)
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col���s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    With udtDetail
        .strHISϵͳʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        .str����סԺ��ʶ = "op"    'סԺ��ʶ
    
        .str�������� = YWS_GetTreatType(1, gobjPati.lng�Һ�ID)
        .str����� = gobjPati.lng�Һ�ID
        .str��λ�� = ""
        '������Ϣ
        .udt������Ϣ = udtPati
        '������Ϣ
        .udt������Ϣ = udtPres
    End With
    'ҩƷ��Ϣ
    Set colTmp = New Collection
    
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_����༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                    
            Else
                blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")) _
                    And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL״̬)) <> 4 '���ϵ�ҽ��������
            If blnDo Then
                If glngModel = PM_����ҽ���嵥 And .TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4" Then
                    '��ȡ��ҩҽ����ID
                    str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str��ҩ;�� = "" 'һ����ҩ���ظ�ȡ
                    If str��ҩ;�� = "" Then
                        If glngModel = PM_����༭ Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                        Else
                            str��ҩ;�� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "������ĿID")  '������
                        End If
                    End If
                    
                    Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
                    
                    If glngModel = PM_����༭ Then
                        strҽ��ID = .RowData(i)
                         '������������λ
                        str���� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                        str���� = Replace(str����, str������λ, "")
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")        '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        str���� = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL����)))
                        If Mid(str����, 1, 2) = "0." Then
                            str���� = "0" & Val(str����)
                        Else
                            str���� = Val(str����)
                        End If
                        str������λ = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL����)))
                        If Mid(str������λ, 1, 2) = "0." Then '������С��������⴦��
                            str������λ = Replace(str������λ, Format(Val(str������λ) & "", "0.####"), "") '�����嵥�������������� & ��������λ����
                        Else
                            str������λ = Replace(str������λ, Val(str������λ) & "", "")    '
                        End If
                        
                        str����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")          '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
                    End If
                    udtMedic.strҩƷ���� = YWS_GetDrugType(.TextMatrix(i, gobjCOL.intCOL�������))       '��ҩ/�г�ҩ/��ҩ
                    udtMedic.str������ = strҽ��ID    '��ҽ��ID
                    udtMedic.Strҽ������ = "T"
                    udtMedic.str����ʱ�� = str����ʱ��
                    udtMedic.str��Ʒ�� = YWS_StrToXML(Sys.RowValue("������ĿĿ¼", Val(.TextMatrix(i, gobjCOL.intCOL������ĿID)), "����"))
                    udtMedic.strҽԺҩƷ���� = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    udtMedic.strҽ������ = ""
                    udtMedic.str��׼�ĺ� = ""
                    udtMedic.str��ҩ��ʼʱ�� = ""
                    udtMedic.str��ҩ����ʱ�� = ""
                    udtMedic.str��� = ""
                    udtMedic.str��� = .TextMatrix(i, gobjCOL.intCOL���ID)
                    udtMedic.str��ҩ���� = ""
                    udtMedic.str��������λ = str������λ
                    udtMedic.str������ = str����
                    udtMedic.strƵ�δ��� = strƵ�ʱ���
                    udtMedic.str��ҩ;������ = str��ҩ;��
                    
                    udtMedic.str��ҩ���� = .TextMatrix(i, gobjCOL.intCOL����)   'OP ���ﴦ����Ч
            
                    colTmp.Add udtMedic, "_" & colTmp.Count + 1
                End If
            End If
        Next
        '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
        If glngModel = PM_����ҽ���嵥 Then
            If str��ҩ��IDs <> "" Then
                Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                With rs��ҩ
                    For i = 1 To .RecordCount
                        If !���ID & "" <> str���ID Then
                            Call GetƵ����Ϣ_����(!Ƶ�� & "", 0, 0, "", IIf(!������� & "" = "7", 2, 1), strƵ�ʱ���)
                            str���ID = !���ID & ""
                        End If
                        udtMedic.strҩƷ���� = YWS_GetDrugType(!������� & "")
                        udtMedic.str������ = !id & ""    '��ҽ��ID
                        udtMedic.Strҽ������ = "T"
                        udtMedic.str����ʱ�� = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str��Ʒ�� = !ҩƷ���� & ""
                        udtMedic.strҽԺҩƷ���� = !ҩƷID & ""
                        udtMedic.strҽ������ = ""
                        udtMedic.str��׼�ĺ� = ""
                        udtMedic.str��ҩ��ʼʱ�� = ""
                        udtMedic.str��ҩ����ʱ�� = ""
                        udtMedic.str��� = ""
                        udtMedic.str��� = !���ID & ""
                        udtMedic.str��ҩ���� = ""
                        udtMedic.str��������λ = !������λ & ""
                        udtMedic.str������ = !�������� & ""
                        udtMedic.strƵ�δ��� = strƵ�ʱ���
                        udtMedic.str��ҩ;������ = !�÷�ID & ""
                        udtMedic.str��ҩ���� = !���� & ""   'OP ���ﴦ����Ч
                
                        colTmp.Add udtMedic, "_" & colTmp.Count + 1
                        .MoveNext
                    Next
                End With
            End If
        End If
        With udtPres
            Set .colҩƷ��Ϣ = colTmp
            .str������ = "0"
            .str�������� = ""
            .str����ʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str�Ƿ�ǰ���� = "1" '0 ��ʷ���� 1 ��ǰ����������Ĭ�ϵ�ǰ�������Ժ����䣩
            .Strҽ������ = "T"  '������ʱ
        End With
    End With
    
    udtDetail.udt������Ϣ = udtPres
    
    OutAdviceCheckWarn_YWS = True
    If udtPres.colҩƷ��Ϣ.Count > 0 Then
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", strXML
        'YWS_��������=6 ������������ֵ��0��1��2��3��8�ֱ����û�����⣬�������⣬һ�����⣬�������⣬��д�������� ����ǰ����
        strTmp = gobjPass.YWS_UI(YWS_��������, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", "����ֵ:" & strTmp & vbCrLf & strRetXML
        '���þ�ʾ��
        Set rsRet = YWS_ReturnRS(strRetXML)
        With gobjAdvice
            arrSQL = Array()
            'ͼƬ�±꣺1-����,2-�Ƶ�,5-�ȵ�,3-���
            '��ʾ��˳��0-����,1-�Ƶ�,2-�ȵ�,3-���
            arrLight(0) = "��": arrLight(1) = "��": arrLight(2) = "��": arrLight(3) = "��"
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                    And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                End If
                If blnDo Then
                    If glngModel = PM_����༭ Then
                        strҽ��ID = .RowData(i)
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    End If
                    'ȡҩ�����ʾֵ
                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    If rsRet.RecordCount > 0 Then
                        k = CLng(rsRet!��ʾֵ & "")
                    Else
                        k = 0
                    End If
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        '���þ�ʾ��
                        If k >= 0 And k <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If glngModel = PM_����༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                        ElseIf PM_����ҽ���嵥 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 3, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '��ҩ�䷽
                        If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                            lngLight = -1 '��ʼ��
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If k > lngLight Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                Else
                    If glngModel = PM_����༭ Then
                        '��ҩ��ʾ�Ƶ�������
                        If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                            '���þ�ʾ��
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                            End If
                            
                            If glngModel = PM_����༭ Then
                                '���������仯,�Ա��������ݿ�
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                    blnNoSave = True    '���Ϊδ����
                                End If
                            End If
                            lng��ҩ��ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
            If glngModel = PM_����ҽ���嵥 And Not rs��ҩ Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '��ҩ����
                    If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        lngLight = -1
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs��ҩ.Filter = "���ID=" & strҽ��ID
                        
                        For j = 1 To rs��ҩ.RecordCount
                            'ȡҩ�����ʾֵ
                            rsRet.Filter = "ҽ��ID ='" & rs��ҩ!id & "'"
                            If rsRet.RecordCount > 0 Then
                                k = CLng(rsRet!��ʾֵ & "")
                            Else
                                k = 0
                            End If
                            '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If k > lngLight Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs��ҩ.MoveNext
                        Next
                        
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        '��ʾ�Ƹ��µ����ݿ�
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 3, lngLight, "NULL") & ")"
                        End If
                    End If
                Next
            End If
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End With

        If glngModel = PM_����༭ Then
            If strTmp = "3" And gbytBlackLamp = 0 Then
                MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_YWS = False
                Exit Function
            ElseIf strTmp = "3" And gbytBlackLamp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then OutAdviceCheckWarn_YWS = False: Exit Function
            End If
            
            '�ϴ�����
            If strTmp <> "0" Then
                strTmp = gobjPass.YWS_UI(YWS_�ϴ�����, gstrBaseXml, strXML, strRetXML)
                WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", "���洦��:����ֵ" & strTmp
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    OutAdviceCheckWarn_YWS = False
End Function


Private Function ReadPatient(ByVal lngPatiID As Long, ByVal strNo As String) As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select b.ID as ����ID,B.����,B.�Ա�,A.��������,A.����,A.���֤�� " & _
         " From ������Ϣ A,���˹Һż�¼ B " & _
         " Where A.����ID=B.����ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
    On Error GoTo errH
    Set ReadPatient = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, strNo)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Function AdviceCheckWarn_MK_YF(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, _
            ByVal lngCmd As Long, Optional ByVal lngCurrID As Long, Optional ByVal strҽ��IDs As String, _
            Optional str��ʾ As String) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'        1/33-�����Զ����(סԺ/����),2/34-�ύ�Զ����(סԺ/����),3-�ֹ��������
'        6=��ҩ����
'      lngCurrID=��ǰҩƷҽ�����кţ�lngCmd=0ʱ��Ҫ
'      strҽ��IDs ҽ��ID����ID1,ID2,ID3...
'      str��ʾ-ҽ��ID:��ʾֵ,ҽ��ID2:��ʾֵ2
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset, rs��� As ADODB.Recordset, rs����ҽ�� As ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, str������λ As String, strƵ�� As String
    Dim str����ҽ�� As String, str����ҽ���� As String, strҩƷID As String
    Dim strSQL As String, i As Long, k As Long
    Dim lng��ʶ��  As Long
    Dim lngCount As Long
    Dim blnDo As Boolean
    Dim strCurrentDate As String

    AdviceCheckWarn_MK_YF = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036ͬһ�����˶�����ʱ������Ϣÿ�ζ�Ҫ����
    '-------------------------------------------------------------
    Set rsTmp = GetPatiInfo_YF(lng����ID, str�Һŵ�, lng��ҳID)
    If str�Һŵ� <> "" Then               '���ﲡ��
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        Call PassSetPatientInfo(lng����ID, rsTmp!����Id, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
            rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), "")
        lng��ʶ�� = NVL(rsTmp!����Id, 0)
    Else                                    'סԺ����
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        Call PassSetPatientInfo(lng����ID, lng��ҳID, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
            rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), _
            IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        lng��ʶ�� = lng��ҳID
    End If
    '���˲��˹���ʷ
    '-------------------------------------------------------
    Set rsTmp = Get���˹�����¼(lng����ID, IIf(str�Һŵ� <> "", 0, lng��ҳID))

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!ҩ��ID & "", rsTmp!ҩ���� & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '���˲���״̬
    '------------------------------------------------------------------
    Set rsTmp = Get������ϼ�¼(lng����ID, lng��ʶ��, IIf(str�Һŵ� <> "", "1,11", "2,12"))
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

    For i = 1 To rsTmp.RecordCount
        Call PassSetMedCond(i & "", rsTmp!���� & "", rsTmp!���� & "", "User", strCurrentDate, strCurrentDate)
        rsTmp.MoveNext
    Next

    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = MK_���PASS�˵�״̬ Then
        If lngCurrID = 0 Then: Exit Function
        strSQL = "Select Nvl(a.�걾��λ, a.ҽ������) As ҩƷ����,a.������Ŀid,a.�շ�ϸĿid As ҩƷid, c.���㵥λ As ������λ, b.ҽ������ As �÷�" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "Where a.Id = [1] And a.������Ŀid = c.Id And a.���id = b.Id(+)"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngCurrID)
        
        If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        'ȡҩƷ����
        strҩƷ = rsTmp!ҩƷ���� & ""
        strҩƷID = rsTmp!ҩƷID & ""
        str������λ = rsTmp!������λ & ""
        If strҩƷID = "" Then strҩƷID = GetDrugID(rsTmp!������ĿID & "")
        'ȡҩƷ��ҩ;��
        str�÷� = rsTmp!�÷� & ""
        '�����ѯҩƷ��Ϣ
        Call PassSetQueryDrug(strҩƷID, strҩƷ, str������λ, str�÷�)
    
        AdviceCheckWarn_MK_YF = 1 '��ʾ���Ե����˵�

        Screen.MousePointer = 0: Exit Function
    ElseIf lngCmd = MK_��ҩ���� Then
        Call PassSetWarnDrug(lngCurrID)    '��ҩ����(�Ѿ����ҽ��Ψһ��)
    ElseIf lngCmd = MK_����״̬����ʷ�鿴 Then
        Call PassDoCommand(lngCmd)  '21-�鿴����ʷ
        Screen.MousePointer = 0
        Exit Function
    Else
        Set rsTmp = GetAdviceInfo_YF(lng����ID, lng��ҳID, str�Һŵ�, strҽ��IDs)
        If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        
        '��ҩ��˻���ҩ�о�
        With rsTmp
            lngCount = 0
            strҩƷ = "": str����ҽ���� = ""
            For i = 1 To .RecordCount
                If Val(!�շ�ϸĿid & "") = 0 Then
                    strҩƷ = strҩƷ & "," & !������ĿID
                End If
                '��ȡ����ҽ��
                If NVL(!����ҽ��) <> "" Then
                    str����ҽ�� = NVL(!����ҽ��)
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    If InStr("," & str����ҽ���� & ",", "," & str����ҽ�� & ",") = 0 And str����ҽ�� <> "" Then
                        str����ҽ���� = str����ҽ���� & "," & str����ҽ��
                    End If
                End If
                .MoveNext
            Next
            
            If strҩƷ <> "" Then
                Set rs��� = GetDrugID(strҩƷ)
            End If
            
            If str����ҽ���� <> "" Then
                str����ҽ���� = Mid(str����ҽ����, 2)
                Set rs����ҽ�� = Sys.RowValue("��Ա��", str����ҽ����, "���,����", "����")
            End If
            
            str����ҽ���� = "": strҩƷID = ""
            .MoveFirst
            
            For i = 1 To .RecordCount
                'ȡ��ҩƵ��(��/��),��Ϊ������������
                strƵ�� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
            
                '����ҽ����Ʒ���´�ʱ,ȡ����ҩƷID
                If Val(!�շ�ϸĿid & "") = 0 Then
                    rs���.Filter = "ҩ��ID =" & !������ĿID
                    If Not rs���.EOF Then strҩƷID = rs���!ҩƷID & ""
                Else
                    strҩƷID = !�շ�ϸĿid & ""
                End If
                '����ҽ��
                str����ҽ�� = NVL(!����ҽ��)
                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                
                If str����ҽ�� <> Mid(str����ҽ����, InStr(str����ҽ����, "/") + 1) Then
                    If Not rs����ҽ�� Is Nothing Then
                        rs����ҽ��.Filter = "����='" & str����ҽ�� & "'"
                        If Not rs����ҽ��.EOF Then str����ҽ���� = rs����ҽ��!��� & "/" & str����ҽ��
                    End If
                End If
                
                '����ҽ����Ϣ
                Call PassSetRecipeInfo(!ҽ��ID & "", strҩƷID, !ҩƷ���� & "", !�������� & "", !������λ & "", strƵ��, _
                    Format(!��ʼʱ�� & "", "yyyy-MM-dd"), Format(!����ʱ�� & "", "yyyy-MM-dd"), !�÷� & "", _
                    !���ID & "", !ҽ����Ч & "", str����ҽ����)

                lngCount = lngCount + 1

                .MoveNext
            Next
    
            '�޿�����ҩƷ
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End With
    End If

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    
    If str��ʾ <> "-1" And lngCount > 0 Then
        str��ʾ = ""
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            k = PassGetWarn(rsTmp!ҽ��ID & "")
            str��ʾ = str��ʾ & "," & rsTmp!ҽ��ID & ":" & k
            rsTmp.MoveNext
        Next
        If str��ʾ <> "" Then str��ʾ = Mid(str��ʾ, 2)
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AdviceCheckWarn_MK4_YF(ByVal lngPatiID As Long, ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String, _
                Optional ByVal strҽ��IDs As String, Optional str��ʾ As String, Optional rsRet As ADODB.Recordset) As Long
'���ܣ�����Pass4ϵͳ��鹦��
'����:
'    strҽ��IDs-��Ҫ����ҽ��ID
'    str��ʾ-ҽ��ID:��ʾֵ,ҽ��ID2:��ʾֵ2
'���أ���������
'   rsRet-���﷢��ʱ�������״̬
'˵����
'
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim strSQL As String, i As Long, k As Long
    Dim str����ҽ�� As String
    Dim str����ҽ������ As String
    Dim strҽ�� As String
    Dim str��ҩĿ�� As String
    Dim bytSubmit As Byte
    
    AdviceCheckWarn_MK4_YF = -1

    On Error GoTo errH
    Screen.MousePointer = 11
    'ҩ����Ϣ��ȡ
    If glngModel = PM_���﷢�� Then
        Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng��ҳID, str�Һŵ�, strҽ��IDs, 2)
        bytSubmit = 1
    Else
        Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng��ҳID, str�Һŵ�, strҽ��IDs)
        bytSubmit = 0
    End If
            
    If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
    '
    Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ)
    
    With rsTmp
        For i = 1 To .RecordCount
            rsAdvice.AddNew
            rsAdvice!ҽ��ID = !ҽ��ID & ""
            rsAdvice!���ID = !���ID & ""
            rsAdvice!ҽ����Ч = !ҽ����Ч & ""
            rsAdvice!ҽ����� = !ҽ����� & ""
            rsAdvice!���� = !���� & ""
            If glngModel = PM_������ҩ Or glngModel = PM_���ŷ�ҩ Or glngModel = PM_PIVA���� Then
                '"0"-���ã�Ĭ�ϣ���"1"-�����ϣ�"2"-��ͣ����"3"-��Ժ��ҩ������ϵͳ���ò�����飩
                rsAdvice!ҽ��״̬ = "0"
            ElseIf glngModel = PM_���﷢�� Then
                rsAdvice!���״̬ = 1
                rsAdvice!ҽ��״̬ = IIf(!ҽ��״̬ & "" = 1, "0", "-1")
            End If
            '
            str����ҽ�� = !����ҽ�� & ""
            If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
            If strҽ�� <> str����ҽ�� And str����ҽ�� <> "" Then
                strҽ�� = str����ҽ��   '����ҽ��ͬһ������ҽ��,ֻ�����һ��
                str����ҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
            End If
            
            rsAdvice!�������� = !�������� & ""
            rsAdvice!��������id = !��������id & ""
            
            rsAdvice!����ҽ������ = str����ҽ������
            rsAdvice!����ҽ�� = str����ҽ��
            rsAdvice!ҩƷID = !�շ�ϸĿid & ""
            
            rsAdvice!ҩƷ���� = !ҩƷ���� & ""
            rsAdvice!�������� = FormatEx(NVL(!��������), 5)
            rsAdvice!������λ = !������λ & ""
            rsAdvice!Ƶ�� = !Ƶ�� & ""
            rsAdvice!�÷� = !�÷� & ""
            rsAdvice!�÷�ID = !�÷�ID & ""    '����4.0�÷�ID���÷�����
            '��ʱҽ����ʼʱ��ͽ���ʱ����ͬ
            If !ҽ����Ч & "" = "0" Then  '����
                rsAdvice!����ʱ�� = Format(!����ʱ�� & "", "YYYY-MM-dd hh:mm:ss")
            Else '��ʱҽ��
                rsAdvice!����ʱ�� = Format(!����ʱ�� & "", "YYYY-MM-dd hh:mm:ss")
            End If
            
            rsAdvice!����ʱ�� = Format(!����ʱ�� & "", "YYYY-MM-dd hh:mm:ss")
            rsAdvice!��ʼʱ�� = Format(!��ʼʱ�� & "", "YYYY-MM-dd hh:mm:ss")
            
            If str�Һŵ� <> "" Then
                '����
                If InStr(",5,6,", "," & !������� & ",") > 0 Then
                    '��ҩ����������,�����۵�λ���,���ﵥλ��ʾ
                    If Not IsNull(!����) And Not IsNull(!�����װ) Then
                        rsAdvice!���� = FormatEx(!���� / !�����װ, 5)
                    End If
                End If
                rsAdvice!������λ = !���ﵥλ & ""
            Else
                rsAdvice!���� = !���� & ""
                rsAdvice!������λ = !סԺ��λ & ""
            End If
            
            '��ҩĿ��(0Ĭ��, 1����Ԥ����2�������ƣ�3Ԥ����4���ƣ�5Ԥ��+����)
            str��ҩĿ�� = !��ҩĿ�� & ""
            If str��ҩĿ�� = "1" Then
                str��ҩĿ�� = "3"
            ElseIf str��ҩĿ�� = "2" Then
                str��ҩĿ�� = "4"
            Else
                str��ҩĿ�� = "0"
            End If
            rsAdvice!��ҩĿ�� = str��ҩĿ��
            rsAdvice!ҽ������ = !ҽ������ & ""
            rsAdvice!������� = !������� & ""
            rsAdvice!������ = !������ & ""
            rsAdvice!ִ�п���ID = !ִ�п���ID & ""
            rsAdvice.Update
            .MoveNext
        Next
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    End With
    Call AdviceCheckWarn_MK4(lngPatiID, str�Һŵ�, lng��ҳID, 1, bytSubmit, rsAdvice, str��ʾ)   '��ʾ������,���ɼ�����
    If glngModel = PM_���﷢�� Then
        Set rsRet = rsAdvice
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetFrequency(ByVal str�����λ As String, ByVal strƵ�ʴ��� As String, ByVal strƵ�ʼ�� As String) As String
'����:����Ƶ���ַ���
    Dim strƵ�� As String
    
    If str�����λ = "��" Then
        strƵ�� = strƵ�ʴ��� & "/" & strƵ�ʼ��
    ElseIf str�����λ = "��" Then
        strƵ�� = strƵ�ʴ��� & "/7"
    ElseIf str�����λ = "Сʱ" Then
        If Val(strƵ�ʼ��) <= 24 Then
            strƵ�� = Format(24 / Val(strƵ�ʼ��) * Val(strƵ�ʴ���), "0") & "/1"
        Else
            strƵ�� = Val(strƵ�ʴ���) & "/" & Format(Val(strƵ�ʼ��) / 24, "0")
        End If
    ElseIf str�����λ = "����" Then
        strƵ�� = Format((24 * 60) / Val(strƵ�ʼ��) * Val(strƵ�ʴ���), "0") & "/1"
    End If
    GetFrequency = strƵ��
End Function

Public Function AdviceCheckWarn_YWS_YF(ByVal lngPatiID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, _
            Optional ByVal strҽ��IDs As String, Optional str��ʾ As String) As Boolean
'���ܣ����ñ���ҩ��ʿ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt����Դ As YWS_ALLERGIC
    Dim udt��� As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   'ҩƷ��Ϣ
    
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Long, lngFunc As Long
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim arrTmp As Variant, curDate As Date
    
    Dim rsTmp As ADODB.Recordset
    Dim rsPati As ADODB.Recordset, rsRet As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset

    Dim str��� As String, str���� As String, strƵ�ʱ��� As String, strXML As String
    Dim str��Ч As String
    Dim lng�Һ�ID As String
    Dim str�������� As String
    Dim str�������� As String
    

    Dim k As Long, blnDo As Boolean

    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '�ж��Ƿ����Ժ��ִ�е�ҩƷ
    '��ȡ������Ϣ
    If str�Һŵ� <> "" Then
        Set rsPati = ReadPatient(lngPatiID, str�Һŵ�)
        If rsPati.RecordCount = 0 Then Exit Function
        lng�Һ�ID = Val(rsPati!����Id & "")
        strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
                        
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng�Һ�ID)
        rsPatiInfo.Filter = "��Ŀ����='���'"
        If rsPatiInfo.RecordCount <> 0 Then str��� = NVL(rsPatiInfo!��¼����)
        rsPatiInfo.Filter = "��Ŀ����='����'"
        If rsPatiInfo.RecordCount <> 0 Then str���� = NVL(rsPatiInfo!��¼����)
        str�������� = lng�Һ�ID & ""
    Else
        Set rsPati = GetPatiInfo(lngPatiID, lng��ҳID)
        If rsPati.RecordCount = 0 Then Exit Function
        str��� = rsPati!��� & ""
        str���� = rsPati!���� & ""
        str�������� = rsPati!סԺ�� & ""
    End If
    
    With udtPati
        .str���� = rsPati!����
        .str�������� = rsPati!�������� & ""
        .str�Ա� = rsPati!�Ա� & ""
        .str���� = str���
        .str��� = str����
        .str���֤�� = rsPati!���֤�� & ""
        .str�������� = str��������
        .str���� = ""
        .str������ = ""
        .str����ʱ�� = ""
        .str����ʱ�䵥λ = ""
        '����Դ
        Set colTmp = New Collection
        Set rsTmp = Get���˹�����¼(lngPatiID, lng��ҳID, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!ҩ��ID <> "" Then
                With udt����Դ
                    .str�������� = "5"   '1=ҩ��ʿҩƷ���� 2=ҩ��ʿҩƷ�ɷ�
                    .str����Դ���� = rsTmp!ҩ����
                    .str����Դ���� = "" & rsTmp!ҩ��ID
                End With
                colTmp.Add udt����Դ, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col����Դs = colTmp
        
        '��ϼ�¼
        Set colTmp = New Collection
        Set rsTmp = Get������ϼ�¼(lngPatiID, IIf(str�Һŵ� <> "", lng�Һ�ID, lng��ҳID), IIf(str�Һŵ� <> "", "1,11", "2,12"))
        For i = 1 To rsTmp.RecordCount
            With udt���
                If rsTmp!����ID & "" <> "" Then
                     .str������� = "2" '2=IDC10����
                Else
                    .str������� = "0" '0=����
                End If
                .str��ϴ��� = "" & rsTmp!����
                .str������� = "" & rsTmp!����
            End With
            colTmp.Add udt���, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        '������
        strTmp = Get���˲��������(lngPatiID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt���
                    .str������� = "1" '1=������״̬
                    .str��ϴ��� = Sys.RowValue("���������", arrTmp(i), "����", "����")
                    .str������� = arrTmp(i)
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col���s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    
    With udtDetail
        .strHISϵͳʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        If str�Һŵ� <> "" Then
            .str����סԺ��ʶ = "op"  '����סԺ��ʶ
            .str�������� = YWS_GetTreatType(1, lng�Һ�ID)
            .str����� = lng�Һ�ID & ""
            .str��λ�� = ""
        Else
            .str����סԺ��ʶ = "ip" '����סԺ��ʶ
            .str�������� = YWS_GetTreatType(2, lngPatiID, lng��ҳID)
            .str����� = rsPati!סԺ�� & ""
            .str��λ�� = "" & rsPati!��ǰ����
        End If
        
        '������Ϣ
        .udt������Ϣ = udtPati
        '������Ϣ
        .udt������Ϣ = udtPres
    End With
    'ҩƷ��Ϣ
    Set colTmp = New Collection
    
    Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng��ҳID, str�Һŵ�, strҽ��IDs)
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        For i = 1 To rsTmp.RecordCount

            Call GetƵ����Ϣ_����(!Ƶ�� & "", 0, 0, "", IIf(!������� & "" = "7", 2, 1), strƵ�ʱ���)
            udtMedic.strҩƷ���� = YWS_GetDrugType(!������� & "")
            udtMedic.str������ = !ҽ��ID & ""    '��ҽ��ID
            udtMedic.Strҽ������ = IIf(!ҽ����Ч & "" = "����", "L", "T")
            udtMedic.str����ʱ�� = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")          '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
            udtMedic.str��Ʒ�� = YWS_StrToXML(!ҩƷ���� & "")
            udtMedic.strҽԺҩƷ���� = !�շ�ϸĿid & ""
            udtMedic.strҽ������ = ""
            udtMedic.str��׼�ĺ� = ""
            udtMedic.str��ҩ��ʼʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str��ҩ����ʱ�� = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str��� = ""
            udtMedic.str��� = !���ID & ""
            udtMedic.str��ҩ���� = ""
            '������������λ
            udtMedic.str��������λ = !������λ & ""
            udtMedic.str������ = !�������� & ""
            udtMedic.strƵ�δ��� = strƵ�ʱ���
            udtMedic.str��ҩ;������ = !�÷�ID & ""
            udtMedic.str��ҩ���� = !���� & ""   'OP ���ﴦ����Ч
            
            colTmp.Add udtMedic, "_" & colTmp.Count + 1
            .MoveNext
        Next
        
        With udtPres
            Set .colҩƷ��Ϣ = colTmp
            .str������ = "0"
            .str�������� = ""
            .str����ʱ�� = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str�Ƿ�ǰ���� = "1" '0 ��ʷ���� 1 ��ǰ����������Ĭ�ϵ�ǰ�������Ժ����䣩
            .Strҽ������ = IIf(str�Һŵ� <> "", "T", "L")
            
        End With
    End With
    
    udtDetail.udt������Ϣ = udtPres
    
    AdviceCheckWarn_YWS_YF = True
    If udtPres.colҩƷ��Ϣ.Count > 0 Then
        On Error GoTo errH
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "AdviceCheckWarn_YWS_YF", strXML
        If glngModel = PM_PIVA���� And strҽ��IDs <> "" Then
            lngFunc = YWS_��������������
        Else
            lngFunc = YWS_��������
        End If
        strTmp = gobjPass.YWS_UI(lngFunc, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "AdviceCheckWarn_YWS_YF", "����ֵ:" & strTmp & vbCrLf & strRetXML
        If lngFunc = YWS_�������������� Then
            '�������⼶���������������鴰��
            Set rsRet = YWS_ReturnRS(strRetXML, 1)
            If rsRet.RecordCount > 0 Then
                frmPassResultYWS.ShowMe rsRet
            End If
        End If
        If str��ʾ <> "-1" Then
            Set rsRet = YWS_ReturnRS(strRetXML)
            str��ʾ = ""
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rsRet.Filter = "ҽ��ID ='" & rsTmp!ҽ��ID & "" & "'"
                If rsRet.RecordCount > 0 Then
                    k = CLng(rsRet!��ʾֵ & "")
                Else
                    k = 0
                End If
                str��ʾ = str��ʾ & "," & rsTmp!ҽ��ID & ":" & k
                rsTmp.MoveNext
            Next
            If str��ʾ <> "" Then str��ʾ = Mid(str��ʾ, 2)
        Else
            str��ʾ = ""
        End If
        
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    AdviceCheckWarn_YWS_YF = False
End Function

Private Function CheckAdvice_YF(ByVal strNo As String, ByVal int���� As Integer, lngPatiID As Long, str�Һŵ� As String, lng��ҳID As Long) As Boolean
'���ܣ���鲡���Ƿ�����
'����:
'���أ�T-���ز���ID,�Һŵ�,��ҳID (�ҵ�ҽ��);F-δ��ҽ��
'
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRet As Boolean
    
     '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
    strSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] " & _
        " Union All " & _
        " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, strNo, int����)

    If rsTmp.RecordCount = 0 Then
        blnRet = False
    Else
        lngPatiID = rsTmp!����ID
        str�Һŵ� = NVL(rsTmp!�Һŵ�)
        lng��ҳID = rsTmp!��ҳID
        blnRet = True
    End If

    CheckAdvice_YF = blnRet
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiInfo_YF(ByVal lngPatiID As Long, ByVal str�Һŵ� As String, ByVal lng��ҳID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If gbytPass = MK And gstrVersion = "3.0" Then
        If str�Һŵ� <> "" Then               '���ﲡ��
            strSQL = "Select B.ID as ����ID,B.����,B.�Ա�,A.��������," & _
                     " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                     " From ������Ϣ A,���˹Һż�¼ B,���ű� C,��Ա�� E" & _
                     " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID" & _
                     " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else                                    'סԺ����
            strSQL = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[3]"
           
        End If
        
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        If str�Һŵ� <> "" Then
            strSQL = "Select A.�����,B.ID as ����ID,B.����,B.����,B.�Ա�,A.��������," & _
                 " C.ID As ����ID,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                 " From ������Ϣ A,���˹Һż�¼ B,���ű� C,��Ա�� E" & _
                 " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID" & _
                 " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else
            strSQL = _
                " Select Nvl(B.����,A.����) ����,Nvl(B.�Ա�,A.�Ա�) �Ա�,A.��������,A.סԺ��,B.���,B.����,B.��Ժ����,B.��Ժ����," & _
                         " C.ID as ����ID,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                         " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                         " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                         " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[3]"
        End If
    ElseIf gbytPass = TYT Then
        If str�Һŵ� <> "" Then
            strSQL = "Select b.ID as ����ID,A.�����,B.����,B.�Ա�,A.��������,A.����,A.���֤�� " & _
            " From ������Ϣ A,���˹Һż�¼ B " & _
            " Where A.����ID=B.����ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else
            strSQL = _
            " Select A.סԺ��,Nvl(B.����,A.����) ����,Nvl(B.�Ա�,A.�Ա�) �Ա� ,A.��������,B.���,B.����  " & _
                     " From ������Ϣ A,������ҳ B" & _
                     " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[3]"
        End If
    ElseIf gbytPass = DT Then
        If str�Һŵ� <> "" Then
            strSQL = "Select b.ID as ����ID,A.�����,B.����,B.�Ա�,A.��������,A.����,A.���֤�� " & _
            " From ������Ϣ A,���˹Һż�¼ B " & _
            " Where A.����ID=B.����ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else
            strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, A.�����, A.������,A.���֤��,B.���,B.����" & vbNewLine & _
                "From ������Ϣ A, ������ҳ B" & vbNewLine & _
                "Where A.����id = B.����id And A.����id = [1] And B.��ҳid = [3]"
        End If
    ElseIf gbytPass = HZYY Then
        If str�Һŵ� <> "" Then
            strSQL = "Select b.ID as ����ID,A.�����,B.����,B.����,B.�Ա�,A.��������,A.����,A.���֤��,A.����,A.����,A.ְҵ,A.����״��,B.ִ�в���ID As �Һſ���ID,C.���� As �Һſ���" & _
            ",A.�ֻ���,A.��ͥ��ַ,D.���� As ҽ�Ƹ��ʽ, B.ִ��ʱ�� As ����ʱ�� " & _
            " From ������Ϣ A,���˹Һż�¼ B,���ű� C, ҽ�Ƹ��ʽ D " & _
            " Where A.����ID=B.����ID And B.ִ�в���ID =C.ID(+) And B.ҽ�Ƹ��ʽ = D.����(+) And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else
            strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) As ����" & vbNewLine & _
                ", A.�����, A.������,A.���֤��,A.����,A.����,B.���,B.����,B.��Ժ����ID,D.���� As ��Ժ����,NVL(B.��Ժ����,0) As ��Ժ����," & vbNewLine & _
                "B.��Ժ����ID,C.���� As ��Ժ����,B.��Ժ����, E.���� AS ҽ�Ƹ��ʽ " & vbNewLine & _
                "From ������Ϣ A, ������ҳ B, ���ű� C, ���ű� D,ҽ�Ƹ��ʽ E " & vbNewLine & _
                "Where A.����id = B.����id  And B.��Ժ����ID =C.ID And B.��Ժ����ID =D.ID(+) And B.ҽ�Ƹ��ʽ = E.����(+) And A.����id = [1] And B.��ҳid = [3]"
        End If
    ElseIf gbytPass = ZL Then
        If str�Һŵ� <> "" Then
            strSQL = "Select b.ID as ����ID,A.�����,A.��Ժʱ��,A.����ʱ��,A.����״��,B.����,B.�Ա�,A.��������,A.����,A.ְҵ,Decode(A.��������, Null, '', Round(Sysdate - A.��������, 2))as �������� " & _
            "   ,A.��ǰ����,B.ִ�в���ID AS ��ǰ����ID,B.ִ����,C.���� AS ��ǰ����" & vbNewLine & _
            " From ������Ϣ A,���˹Һż�¼ B, ���ű� C " & _
            " Where A.����ID=B.����ID And B.ִ�в���ID =C.ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
        Else
            strSQL = "Select A.סԺ��,A.��������,A.��Ժʱ��,A.����״��, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�,Nvl(B.����, A.����) As ����,A.ְҵ," & _
                "Decode(A.��������, Null, '', Round(Sysdate - A.��������, 2))as ��������,B.���,B.���� " & vbNewLine & _
                "   ,A.��ǰ����,A.��ǰ����ID,B.��ǰ����ID,C.���� AS ��ǰ����,D.���� AS ��ǰ���� " & vbNewLine & _
                "From ������Ϣ A, ������ҳ B,���ű� C,���ű� D" & vbNewLine & _
                "Where B.����id = A.����id And A.��ǰ����ID =C.ID(+) And B.��ǰ����ID =D.ID(+) And B.����id = [1] And B.��ҳid = [3]"
        End If
    End If
    Set GetPatiInfo_YF = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, str�Һŵ�, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAdviceInfo_YF(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, _
    Optional ByVal strҽ��IDs As String, Optional ByVal bytFunc As Byte = 0) As ADODB.Recordset
'����:��ȡҩ����Ϣ
    Dim strSQL As String
    
    If str�Һŵ� <> "" Then
        strSQL = " And a.�Һŵ� = [3]"
    Else
        strSQL = " And  a.����ID =[1] And a.��ҳID = [2]"
    End If
    If bytFunc = 0 Then
        If strҽ��IDs <> "" Then
            strҽ��IDs = "," & strҽ��IDs & ","
            strSQL = "Select a.Id As ҽ��id, a.���id,a.ҽ����Ч, a.��� As ҽ�����,a.������־ as ��־,a.����ҽ��, a.ҽ��״̬, a.����ʱ��, a.��ʼִ��ʱ�� As ��ʼʱ��, a.ִ����ֹʱ�� As ����ʱ��, a.�������," & vbNewLine & _
            "       a.������Ŀid, a.�շ�ϸĿid, a.ִ��Ƶ�� As Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��������, a.�ܸ����� As ����, e.�����װ, e.���ﵥλ, e.סԺ��λ, a.����," & vbNewLine & _
            "       b.���� As ҩƷ����, b.���㵥λ As ������λ, b.ִ��Ƶ��, g.���� As �÷�, c.������Ŀid As �÷�id, a.��������id, a.ִ�п���id, d.���� As ��������, a.��ҩĿ��,a.��ҩ����,a.ҽ������,F.���, " & vbNewLine & _
            "  Decode(G.���||'_'||G.��������||'_'||G.ִ�з���,'E_2_1',C.ҽ������,'') As ����,NULL As ������" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ����¼ C, ���ű� D, ҩƷ��� E,�շ���ĿĿ¼ F, ������ĿĿ¼ G " & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.���id = c.Id(+) And a.��������id = d.Id(+) And a.�շ�ϸĿid = e.ҩƷid(+) And a.�շ�ϸĿid=F.ID(+) And " & vbNewLine & _
            "      c.������Ŀid = g.Id(+) And a.������� In ('5', '6', '7') " & strSQL & " And inStr([4],','|| a.ID ||',')>0 " & vbNewLine & _
            "Order By a.���"
        Else
            strSQL = "Select a.Id As ҽ��id, a.���id, a.ҽ����Ч, a.��� As ҽ�����,a.������־ as ��־,a.����ҽ��, a.ҽ��״̬, a.����ʱ��, a.��ʼִ��ʱ�� As ��ʼʱ��, a.ִ����ֹʱ�� As ����ʱ��, a.�������," & vbNewLine & _
                     "       a.������Ŀid, a.�շ�ϸĿid, a.ִ��Ƶ�� As Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��������, a.�ܸ����� As ����, e.�����װ, e.���ﵥλ, e.סԺ��λ, a.����," & vbNewLine & _
                     "       b.���� As ҩƷ����, b.���㵥λ As ������λ, b.ִ��Ƶ��, g.���� As �÷�, c.������Ŀid As �÷�id, a.��������id, a.ִ�п���id, d.���� As ��������, a.��ҩĿ��,a.��ҩ����,a.ҽ������,F.���,Decode(G.���||'_'||G.��������||'_'||G.ִ�з���,'E_2_1',C.ҽ������,'') As ����, NULL As ������" & vbNewLine & _
                     "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ����¼ C, ���ű� D, ҩƷ��� E,�շ���ĿĿ¼ F, ������ĿĿ¼ G " & vbNewLine & _
                     "Where a.������Ŀid = b.Id And a.���id = c.Id(+) And a.��������id = d.Id(+) And a.�շ�ϸĿid = e.ҩƷid(+) And a.�շ�ϸĿid=F.ID(+) And " & vbNewLine & _
                     "      c.������Ŀid = g.Id(+) And a.������� In ('5', '6', '7') " & strSQL & vbNewLine & _
                     "      And ((a.ҽ����Ч = 1 And Trunc(a.����ʱ��) = Trunc(Sysdate) And a.ҽ��״̬ = 8) Or" & vbNewLine & _
                     "      (a.ҽ����Ч = 0 And (a.ҽ��״̬ In (8, 9) And a.ִ����ֹʱ�� >= Sysdate Or a.ҽ��״̬ In (3, 5, 7))))" & vbNewLine & _
                     "Order By a.���"
        End If
    ElseIf bytFunc = 1 Then
        strSQL = "Select a.Id As ҽ��id, a.���id, a.ҽ����Ч, a.��� As ҽ�����,a.������־ as ��־,a.����ҽ��, a.ҽ��״̬, a.����ʱ��, a.��ʼִ��ʱ�� As ��ʼʱ��, a.ִ����ֹʱ�� As ����ʱ��, a.�������," & vbNewLine & _
        "       a.������Ŀid, a.�շ�ϸĿid, a.ִ��Ƶ�� As Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��������, a.�ܸ����� As ����, e.�����װ, e.���ﵥλ, e.סԺ��λ, a.����," & vbNewLine & _
        "       b.���� As ҩƷ����, b.���㵥λ As ������λ, b.ִ��Ƶ��, g.���� As �÷�, c.������Ŀid As �÷�id, a.��������id, a.ִ�п���id, d.���� As ��������, a.��ҩĿ��, a.��ҩ����, a.ҽ������,F.���," & vbNewLine & _
        "       Decode(G.���||'_'||G.��������||'_'||G.ִ�з���,'E_2_1',C.ҽ������,'') As ����,NULL As ������,a.������� AS ����ID,A.ִ������ AS Aִ������,C.ִ������ AS Bִ������,G.���,G.��������,G.ִ�з���  " & vbNewLine & _
        "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ����¼ C, ���ű� D, ҩƷ��� E,�շ���ĿĿ¼ F, ������ĿĿ¼ G " & vbNewLine & _
        "Where a.������Ŀid = b.Id And a.���id = c.Id(+) And a.��������id = d.Id(+) And a.�շ�ϸĿid = e.ҩƷid(+) And a.�շ�ϸĿid=F.ID(+) And " & vbNewLine & _
        "      c.������Ŀid = g.Id(+) And a.������� In ('5', '6', '7') And Nvl(A.ִ�б��,0)<>-1 " & strSQL & vbNewLine & _
        "Order By a.���"
    ElseIf bytFunc = 2 Then
        strSQL = "Select a.����id, a.�Һŵ�, a.Id As ҽ��id, a.���id, a.ҽ����Ч, a.��� As ҽ�����, a.������־ As ��־, a.����ҽ��, a.ҽ��״̬, a.����ʱ��, a.��ʼִ��ʱ�� As ��ʼʱ��," & vbNewLine & _
            "       a.ִ����ֹʱ�� As ����ʱ��, a.�������, a.������Ŀid, a.�շ�ϸĿid, a.ִ��Ƶ�� As Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��������, a.�ܸ����� As ����, e.�����װ," & vbNewLine & _
            "       e.���ﵥλ, e.סԺ��λ, a.����, b.���� As ҩƷ����, b.���㵥λ As ������λ, b.ִ��Ƶ��, g.���� As �÷�, c.������Ŀid As �÷�id, a.��������id, a.ִ�п���id," & vbNewLine & _
            "       d.���� As ��������, a.��ҩĿ��, a.��ҩ����, a.ҽ������, f.���," & vbNewLine & _
            "       Decode(g.��� || '_' || g.�������� || '_' || g.ִ�з���, 'E_2_1', c.ҽ������, '') As ����, Nvl(a.�������, 0) As ������" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ����¼ C, ���ű� D, ҩƷ��� E, �շ���ĿĿ¼ F, ������ĿĿ¼ G" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.���id = c.Id(+) And a.��������id = d.Id(+) And a.�շ�ϸĿid = e.ҩƷid(+) And a.�շ�ϸĿid = f.Id(+) And" & vbNewLine & _
            "      c.������Ŀid = g.Id(+) And a.������� In ('5', '6', '7') And" & vbNewLine & _
            "      (a.ҽ��״̬ = 1 And Instr([4], ',' || a.���id || ',') > 0 Or a.ҽ��״̬ = 8) " & strSQL & vbNewLine & _
            "Order By a.���"
    End If
    On Error GoTo errH
    Set GetAdviceInfo_YF = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng����ID, lng��ҳID, str�Һŵ�, "," & strҽ��IDs & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function AdviceCheckWarn_MK4(ByVal lngPatiID As Long, ByVal str�Һŵ� As String, ByVal lng��ҳID As Long, ByVal bytShow As Byte, _
    ByVal bytSubmit As Byte, ByRef rsAdvice As ADODB.Recordset, Optional str��ʾ As String = "-1", Optional lngResult As Long = 1) As String
'���ܣ�����4.0�Զ���ӿ�
'����:
'   lngCmd=0 MK4_���PASS�˵�״̬
'       1-�ֶ����
'       2-�������
'   bytShow-0-����ʾ����,1-��ʾ����
'   bytSubmit-0-���ɼ�����,1-�ɼ�����
'   rsAdvice-����ҽ����¼(��������,���ڷ��ؾ�ʾ��Ϣ)
'   lngResult ��ʾҩʦ��Ԥ�����1-ͨ����0-����ͨ��
'����

'str��ʾ-���ؾ�ʾ������ʽ��ҽ��ID1:��ʾֵ1,ҽ��ID2:��ʾֵ2    ��ȱʡ�����ؾ�ʾֵ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCount As Long, strInHospNo As String, strVisitCode As String
    Dim i As Long, k As Long
    Dim str��� As String, str���� As String, str�����ʶ As String
    Dim lng���� As Long, lng���� As Long, lng�ι� As Long, lng���� As Long
    Dim str�������� As String, strTmp As String, strJson As String
    Dim lng�Һ�ID As Long
    Dim strҩƷID As String
    Dim rs��� As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim strҩ��IDs As String
    Dim strPharmacyName As String
    
    Dim colTemp As New Collection
    
    On Error GoTo errH
    Screen.MousePointer = 11
    '����4.0

     '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
     '-------------------------------------------------------------
     Set rsTmp = GetPatiInfo_YF(lngPatiID, str�Һŵ�, lng��ҳID)
     If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    '�����������
    strTmp = Get���˲��������(lngPatiID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
    Call PASS4���������(strTmp, lng����, lng����, lng�ι�, lng����, str��������)
    'ֱ�Ӵ�lng����,lng���ﴫ�������ӿڣ������ڲ�����ʱֵ�ᱻת���ɼ���ֵ
    colTemp.Add -1, "K" & "-1"
    colTemp.Add 0, "K" & "0"
    colTemp.Add 1, "K" & "1"
    colTemp.Add 2, "K" & "2"
    colTemp.Add 3, "K" & "3"
    colTemp.Add 4, "K" & "4"
     
    'PASS����һ�����˵Ļ�����ϢMDC_SetPatient
    If str�Һŵ� <> "" Then
        lng�Һ�ID = rsTmp!����Id
        '������Ϣ
        strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng�Һ�ID)
        rsPatiInfo.Filter = "��Ŀ����='���'"
        If rsPatiInfo.RecordCount > 0 Then str��� = NVL(rsPatiInfo!��¼����)
        rsPatiInfo.Filter = "��Ŀ����='����'"
        If rsPatiInfo.RecordCount > 0 Then str���� = NVL(rsPatiInfo!��¼����)
        
        str�����ʶ = IIf(NVL(rsTmp!����, 0) = 0, 2, 3)
        'ҩʦ��Ԥϵͳ
        strInHospNo = rsTmp!����� & "/" & str�Һŵ�
        strVisitCode = rsTmp!����� & ""
         'A.PASS����һ�����˵Ļ�����ϢMDC_SetPatient
        Call MDC_SetPatient(lngPatiID, strInHospNo, strVisitCode, rsTmp!���� & "", NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), _
                     str���, str����, rsTmp!����ID & "", rsTmp!������ & "", rsTmp!ҽ���� & "", rsTmp!ҽ���� & "", str�����ʶ, CLng(colTemp("K" & lng����)), CLng(colTemp("K" & lng����)), str��������, CLng(colTemp("K" & lng�ι�)), CLng(colTemp("K" & lng����)))
     
     Else
        strInHospNo = rsTmp!סԺ�� & ""
        strVisitCode = lng��ҳID & ""
        Call MDC_SetPatient(lngPatiID & "", rsTmp!סԺ�� & "", lng��ҳID & "", rsTmp!���� & "", rsTmp!�Ա� & "", _
             Format(rsTmp!��������, "yyyy-MM-dd"), rsTmp!��� & "", rsTmp!���� & "", _
             rsTmp!����ID & "", rsTmp!������ & "", rsTmp!ҽ���� & "", rsTmp!ҽ���� & "", 1, CLng(colTemp("K" & lng����)), CLng(colTemp("K" & lng����)), str��������, CLng(colTemp("K" & lng�ι�)), CLng(colTemp("K" & lng����)))
     End If
     
     '���˲��˹���ʷ
     '-------------------------------------------------------
     Set rsTmp = Get���˹�����¼(lngPatiID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
     
     'PASS����һ����׼���Ĺ�����¼�������ظ����ã�MDC_AddAller
     '����ȡһ��ҩƷID����
     For i = 1 To rsTmp.RecordCount
         strҩƷID = ""
         If rsTmp!ҩ��ID & "" <> "" Then
             strSQL = "select ҩƷID from ҩƷ��� where ҩ��id=[1] and rownum <2"
             Set rs��� = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, rsTmp!ҩ��ID)
             If Not rs���.EOF Then strҩƷID = rs���!ҩƷID & ""
         End If
         Call MDC_AddAller(i, strҩƷID, rsTmp!ҩ���� & "", rsTmp!������Ӧ & "")
         rsTmp.MoveNext
     Next

     '���˲������
     '------------------------------------------------------------------
     'PASS����һ����׼������ϼ�¼�������ظ����ã�MDC_AddMedCond
    If glngModel = PM_����༭ Then
        If Not gobjDiags Is Nothing Then
            With gobjDiags
                For i = 1 To .Count
                    If .Item(i).str������� <> "" Then
                        Call MDC_AddMedCond(i & "", IIf(.Item(i).str�������� <> "", .Item(i).str��������, .Item(i).str��ϱ���), .Item(i).str�������, "")
                    End If
                Next
            End With
        End If
    Else
        Set rsTmp = Get������ϼ�¼(lngPatiID, IIf(str�Һŵ� <> "", lng�Һ�ID, lng��ҳID), IIf(str�Һŵ� <> "", "1,11", "2,12"))
        If lng��ҳID <> 0 Then
           Set rsTmp = zlDatabase.CopyNewRec(rsTmp, , "����,����") '����Ϊ�ɱ༭�ļ�¼��117045
           If CreatePlugInOK Then
               On Error Resume Next
               Call gobjPlugIn.SetPassDiag(lngPatiID, lng��ҳID, rsTmp)
               Call zlPlugInErrH(Err, "SetPassDiag")
               If Not rsTmp Is Nothing Then rsTmp.Filter = ""
               Err.Clear: On Error GoTo 0
           End If
        End If
        For i = 1 To rsTmp.RecordCount
            Call MDC_AddMedCond(i & "", rsTmp!���� & "", rsTmp!���� & "", "")
            rsTmp.MoveNext
        Next
    End If
    '���벡��������¼MDC_AddOperation
    Set rsTmp = GetPatiOperation(lngPatiID, lng��ҳID, str�Һŵ�)
    
    On Error Resume Next
    For i = 1 To rsTmp.RecordCount
        Call MDC_AddOperation(rsTmp!id & "", rsTmp!���� & "", rsTmp!���� & "", "", Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), "")
        rsTmp.MoveNext
    Next
    Err.Clear: On Error GoTo 0
    
    On Error GoTo errH
     'PASS����һ����ҩ�嵥��¼�������ظ����ã�MDC_AddScreenDrug
     lngCount = 0
     
     With rsAdvice
         strҩ��IDs = ""
         For i = 1 To .RecordCount
            If InStr("," & strҩ��IDs & ",", "," & !ִ�п���ID & ",") = 0 Then
                strҩ��IDs = strҩ��IDs & "," & !ִ�п���ID
            End If
            .MoveNext
         Next
         If strҩ��IDs <> "" Then Set rsTmp = GetRS("���ű�", "ID,����", strҩ��IDs)
         If .RecordCount > 0 Then .MoveFirst
         For i = 1 To .RecordCount
             '����ҽ����Ϣ
             'ҽ��ID,���ID,ҽ����Ч,ҽ�����,ҽ��״̬,��������,��������ID,����ҽ������,����ҽ��,ҩƷID,ҩƷ����,��������,������λ,Ƶ��,�÷�,�÷�ID,����ʱ��,��ʼʱ��,����ʱ��,����,������λ,��ҩĿ��,ҽ������
             If glngModel = PM_���﷢�� Then
                strTmp = !������
             Else
                strTmp = ""
             End If
            If Val(!ҽ��״̬ & "") = -1 Then
                '���﷢�ʹ�����ʷҽ��
                strJson = FuncGetOtherRecipInfo(!ҽ��ID, strTmp, !ҩƷID, !ҩƷ����, !�÷�, !Ƶ��, !������λ, !��������, !����, !������λ, !����)
                If strJson <> "" Then Call MDC_AddJsonInfo(strJson)
            Else
                Call MDC_AddScreenDrug(!ҽ��ID, !ҽ�����, !ҩƷID, !ҩƷ���� & "", !�������� & "", !������λ & "", !Ƶ�� & "", !�÷� & "", !�÷� & "", !����ʱ�� & "", _
                        !����ʱ�� & "", !��ʼʱ�� & "", !���ID & "", !ҽ����Ч & "", !ҽ��״̬ & "", !��������id & "", !�������� & "", !����ҽ������ & "", _
                        !����ҽ�� & "", strTmp, !���� & "", !������λ & "", !��ҩĿ�� & "", "", "", !ҽ������ & "")
            End If
            '����\ִ�п���
            rsTmp.Filter = "ID=" & Val(!ִ�п���ID): strPharmacyName = ""
            If Not rsTmp.EOF Then strPharmacyName = rsTmp!���� & ""
            strJson = FuncGetDripInfo(!ҽ��ID & "", !���� & "", Val(!ִ�п���ID), strPharmacyName, !����)
            If strJson <> "" Then Call MDC_AddJsonInfo(strJson)
            lngCount = lngCount + 1
            
             .MoveNext
         Next
     End With
     
     
     '�޿�����ҩƷl
     If lngCount = 0 Then
         Screen.MousePointer = 0: Exit Function
     End If
     
    If gblnTEST Then bytShow = 0
     'PASS��麯��MDC_DoCheck
    If bytShow = 0 And bytSubmit = 0 Then
        Call MDC_DoCheck(G_INT_MODEL_0, G_INT_MODEL_0)  '����ʾ����,���ɼ�
    ElseIf bytShow = 0 And bytSubmit = 1 Then
        Call MDC_DoCheck(G_INT_MODEL_0, G_INT_MODEL_1) '����ʾ����,Ҫ�ɼ�
    ElseIf bytShow = 1 And bytSubmit = 0 Then
        Call MDC_DoCheck(G_INT_MODEL_1, G_INT_MODEL_0) '��ʾ����,��Ҫ�ɼ�
    ElseIf bytShow = 1 And bytSubmit = 1 Then
        Call MDC_DoCheck(G_INT_MODEL_1, G_INT_MODEL_1) '��ʾ����,Ҫ�ɼ�
    End If
    
    If gblnPharmReview And glngModel = PM_סԺ�༭ Then
       On Error Resume Next
       lngResult = MDC_GetTaskStatus(lngPatiID, strInHospNo, strVisitCode, "", 1)  '����ֵ:1-ͨ��
       WriteLog "" & glngModel, "AdviceCheckWarn_MK4", "MDC_GetTaskStatus ����ֵ:" & lngResult
       Err.Clear: On Error GoTo 0
    ElseIf gblnPharmReview And glngModel = PM_���﷢�� Then
        With rsAdvice
            .Filter = "ҽ��״̬='0'": strTmp = "": lngResult = 0
            For i = 1 To .RecordCount
                If strTmp <> !������ & "" Then
                    strTmp = !������ & ""
                    On Error Resume Next
                    lngResult = MDC_GetTaskStatus(lngPatiID, strInHospNo, strVisitCode, strTmp, 2)
                    WriteLog "" & glngModel, "AdviceCheckWarn_MK4", "MDC_GetTaskStatus ����ֵ:" & lngResult
                    Err.Clear: On Error GoTo 0
                    !���״̬ = lngResult
                Else
                    !���״̬ = lngResult
                End If
                .MoveNext
            Next
            .Filter = ""
        End With
    Else
       lngResult = 1
    End If

     'PASS��龯ʾֵ����
    If str��ʾ <> "-1" And lngCount > 0 Then
        str��ʾ = ""
        rsAdvice.MoveFirst
        For i = 1 To rsAdvice.RecordCount
            k = MDC_GetWarningCode(rsAdvice!ҽ��ID & "")
            str��ʾ = str��ʾ & "," & rsAdvice!ҽ��ID & ":" & k
            rsAdvice!��ʾ = k
            rsAdvice.MoveNext
        Next
        If str��ʾ <> "" Then str��ʾ = Mid(str��ʾ, 2)
        
    Else
        str��ʾ = ""
    End If
            
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function OutAdviceCheckWarn_MK4(Optional ByVal bytShow As Byte = 0, Optional ByVal bytSubmit As Byte = 0, _
        Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
        Optional ByRef lngResult As Long = 1) As Long
'���ܣ�����Passϵͳ�ж�ҽ�����к�����ҩ������ع���
'������bytShow=0-����ʾ���������,1-��ʾ���������
'       0-���˵������ԣ�1-���ӿ�
'       0-�������PASS�˵�״̬,1-���ӿ�
'       bytSubmit=0-�����ϴ�����,1-�ϴ�����
'���Σ�
'       rsOut-���ؽ���ҩƷ˵��
'       lngResult-ҩʦ��Ԥϵͳ 0-��ͨ����1-ͨ��
'���أ�������˷��ص���߼���ʾֵ,Ϊ-1,-2,-3��ʾû�н������
'      ���PASS�˵�ʱ������>=0��ʾ���Ե����˵�
'˵������ҩ��飺�漰�����µ�����(������ִ��)����δֹͣ�ĳ���
'      ��ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset, rsPatiInfo As New ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rs��ҩ As ADODB.Recordset
    
    Dim strҩƷ���� As String, str�÷� As String, strƵ�� As String, str�÷�ID As String, str�����λ As String
    Dim str��ϱ��� As String, str������� As String, strTmp As String, strPre�÷� As String, str������� As String
    Dim strҽ��ID As String, str���ID As String, strҽ����� As String, str�������� As String, str������λ As String
    Dim str�����ʶ As String, str���� As String, str������λ As String, str��ҩĿ�� As String, strҽ������ As String
    Dim strҽ����Ч As String, strҽ��״̬ As String, strҽ������ As String
    Dim strҩƷID As String, str��������ID As String, str��������Tag As String
    Dim str�������� As String, str����ҽ�� As String, str����ҽ��Tag As String
    Dim str����ʱ�� As String, str��ʼʱ�� As String, str����ʱ�� As String
    Dim str��ʾ As String, str��ʾֵ As String, str���� As String
    Dim str��ҩ��IDs As String, strGroupIDs As String
    Dim strҽ��IDs As String
    Dim strִ�п���ID As String
    
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim lngTmp As Long, lng��ҩ��ID As Long, lngLight As Long
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long, j As Long
    Dim arrTmp As Variant
    
    Dim strType As String
    Dim str��� As String, str���� As String
    Dim arrLight(0 To 4) As String
    
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer
    Dim objDiag As clsDiagItem
    Dim rs��� As ADODB.Recordset
    
    Dim arrSQL As Variant
    
    lngMaxWarn = -1
    OutAdviceCheckWarn_MK4 = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
     '���벡��ҽ����Ϣ
    '-------------------------------------------------------------
    '�����˽���ҩƷ˵������;����Ϊ����༭;��鹦��
    If glngModel = PM_����༭ And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_�������)
    End If
    
    With gobjAdvice
        lngCount = 0
        curDate = zlDatabase.Currentdate
        '��ʼ��ҩ����Ϣ
        Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ)
        
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_����༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            Else
                blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) _
                Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4"))
                blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
            
            If blnDo Then
                If glngModel = PM_����ҽ���嵥 And (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                    '��ȡ��ҩҽ����ID
                    str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                    strҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ
                    
                    If str�÷� = "" Then
                        str���� = ""
                        If glngModel = PM_����༭ Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                                Else
                                    str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    If InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "��/����") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcolҽ������), "����/Сʱ") > 0 Then
                                        str���� = .TextMatrix(k, gobjCOL.intcolҽ������)
                                    End If
                                End If
                            End If
                        Else
                            str�÷� = Sys.RowValue("����ҽ����¼", Val(.TextMatrix(i, gobjCOL.intCOL���ID)), "ҽ������")
                        End If
                    End If
    
                    'ȡ��ҩƵ��(��/��),��Ϊ������������
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then strƵ�� = ""    'һ����ҩ���ظ�ȡ
                    If strƵ�� = "" Then
                        strƵ�� = .TextMatrix(i, gobjCOL.intCOLƵ��)
                    End If
    
                    '������������
                    str��������ID = .TextMatrix(i, gobjCOL.intCOL��������ID)
                    If str��������ID <> str��������Tag And str��������ID <> "" Then
                        str�������� = Sys.RowValue("���ű�", Val(str��������ID), "����")
                        str��������Tag = str��������ID
                    End If
                    
                    '����ҽ��
                    str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    
                    If str����ҽ��Tag <> str����ҽ�� And str����ҽ�� <> "" Then
                        strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
                        str����ҽ��Tag = str����ҽ��
                    End If
    
                    strType = .TextMatrix(i, gobjCOL.intCOL״̬)
                    '"0"-���ã�Ĭ�ϣ���"1"-�����ϣ�"2"-��ͣ����"3"-��Ժ��ҩ������ϵͳ���ò�����飩
                    If strType = "4" Then '4-����
                        strҽ��״̬ = "1"
                    Else
                        strҽ��״̬ = "0"
                    End If
                    
                    'PASS����һ����ҩ�嵥��¼�������ظ����ã�MDC_AddScreenDrug
                    If glngModel = PM_����༭ Then
                        strҽ��ID = .RowData(i)
                        strҽ����� = .TextMatrix(i, gobjCOL.intCOL���)
                        str�������� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                        str���� = .TextMatrix(i, gobjCOL.intCOL����)
                        str������λ = .TextMatrix(i, gobjCOL.intcol������λ)
                        
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
                        str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
                        str����ʱ�� = "" '����ҽ������վ������ҩ�� ����ֵ���Ϳ��������ظ���ҩ
                        strִ�п���ID = .TextMatrix(i, gobjCOL.intColִ�п���ID)
                        
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                            '��ҩ,�г�ҩ
                                rsOut.AddNew
                                rsOut!ҽ��ID = CLng(.RowData(i) & "")
                                rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                                rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                                rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                            '��ҩ�䷽  ����˵����������ҩ������
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                    rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                    rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                    rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                    rsOut.Update
                                End If
                            End If
                        End If
                            
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        strҽ��IDs = strҽ��IDs & "," & strҽ��ID
                        strҽ����� = "-1"  'ϵͳ�Զ����
                        str�������� = Val(.TextMatrix(i, gobjCOL.intCOL����))
                        str������λ = .TextMatrix(i, gobjCOL.intCOL����)
                        str�������� = FormatEx(str��������, 5)
                        
                        str������λ = Replace(str������λ, str��������, "")
                        
                        str���� = Val(.TextMatrix(i, gobjCOL.intCOL����))
                        str������λ = .TextMatrix(i, gobjCOL.intCOL����)
                        str���� = FormatEx(str����, 5)
                        str������λ = Replace(str������λ, str����, "")
                        
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
                        str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
                        str����ʱ�� = ""  '����ҽ������վ������ҩ�� ����ֵ���Ϳ��������ظ���ҩ
                        strִ�п���ID = ""
                        If InStr(strGroupIDs & ",", "," & .TextMatrix(i, gobjCOL.intCOL���ID) & ",") = 0 Then
                            strGroupIDs = strGroupIDs & "," & .TextMatrix(i, gobjCOL.intCOL���ID)
                        End If
                    End If
                    str���ID = .TextMatrix(i, gobjCOL.intCOL���ID)
                    strҽ������ = .TextMatrix(i, gobjCOL.intcolҽ������)
                    str��ҩĿ�� = .TextMatrix(i, gobjCOL.intcol��ҩĿ��)
                    str������� = .TextMatrix(i, gobjCOL.intCOL�������)
                    If str��ҩĿ�� = "1" Then
                        str��ҩĿ�� = "3"
                    ElseIf str��ҩĿ�� = "2" Then
                        str��ҩĿ�� = "4"
                    Else
                        str��ҩĿ�� = "0"
                    End If
                    
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = strҽ��ID
                    rsAdvice!���ID = str���ID
                    rsAdvice!ҽ����Ч = "1" '����
                    rsAdvice!ҽ����� = strҽ�����
                    rsAdvice!ҽ��״̬ = strҽ��״̬
                    rsAdvice!�������� = str��������
                    rsAdvice!��������id = str��������ID
                    rsAdvice!����ҽ������ = strҽ������
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!ҩƷID = strҩƷID
                    rsAdvice!ҩƷ���� = strҩƷ����
                    rsAdvice!�������� = str��������
                    
                    rsAdvice!������λ = str������λ
                    rsAdvice!Ƶ�� = strƵ��
                    rsAdvice!�÷� = str�÷�
                    rsAdvice!�÷�ID = ""
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!��ʼʱ�� = str��ʼʱ��
                    rsAdvice!����ʱ�� = str����ʱ��
                    
                    rsAdvice!���� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!��ҩĿ�� = str��ҩĿ��
                    rsAdvice!ҽ������ = strҽ������
                    rsAdvice!������� = str�������
                    rsAdvice!���� = str����
                    rsAdvice!ִ�п���ID = strִ�п���ID
                    rsAdvice.Update
                    '---------------------------------------------------------------------------
                    lngCount = lngCount + 1
                End If
            End If
        Next
        '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
        If glngModel = PM_����ҽ���嵥 Then
            If str��ҩ��IDs <> "" Then
                Set rs��ҩ = Get��ҩ�䷽(str��ҩ��IDs)
                With rs��ҩ
                    For i = 1 To .RecordCount
                        If !���ID & "" <> str���ID Then
                            str����ҽ�� = !����ҽ�� & ""
                            If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                            str����ҽ�� = Sys.RowValue("��Ա��", str����ҽ��, "���", "����") & "/" & str����ҽ��
                            str�������� = Sys.RowValue("���ű�", Val(!��������id & ""), "����")

                            str����ʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                            str����ʱ�� = ""
                            str��ʼʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                            If !ҽ����Ч & "" = "1" Then
                                str����ʱ�� = str����ʱ��
                            End If
                            
                            If !��ҩĿ�� & "" = "1" Then
                                str��ҩĿ�� = "3"
                            ElseIf !��ҩĿ�� & "" = "2" Then
                                str��ҩĿ�� = "4"
                            Else
                                str��ҩĿ�� = "0"
                            End If
                            
                            If !ҽ��״̬ & "" = "4" Then '����
                                strҽ��״̬ = "1"
                            Else
                                strҽ��״̬ = "0"
                            End If
                            str���ID = !���ID & ""
                        End If
                        '----------------------------------------------------------
                        rsAdvice.AddNew
                        rsAdvice!ҽ��ID = !id
                        strҽ��IDs = strҽ��IDs & "," & !id
                        rsAdvice!���ID = !���ID & ""
                        rsAdvice!ҽ����Ч = "1"
                        rsAdvice!ҽ����� = lngCount + 1
                        rsAdvice!ҽ��״̬ = strҽ��״̬
                        rsAdvice!�������� = str��������
                        rsAdvice!��������id = !��������id & ""
                        rsAdvice!����ҽ������ = strҽ������
                        rsAdvice!����ҽ�� = str����ҽ��
                        rsAdvice!ҩƷID = !ҩƷID & ""
                        rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                        rsAdvice!�������� = !�������� & ""
                        
                        rsAdvice!������λ = !������λ & ""
                        rsAdvice!Ƶ�� = !Ƶ�� & ""
                        rsAdvice!�÷� = !�÷� & ""
                        rsAdvice!�÷�ID = ""
                        rsAdvice!����ʱ�� = str����ʱ��
                        rsAdvice!��ʼʱ�� = str��ʼʱ��
                        rsAdvice!����ʱ�� = str����ʱ��
                        
                        rsAdvice!���� = !�ܸ����� & ""
                        rsAdvice!������λ = !���ﵥλ & ""
                        rsAdvice!��ҩĿ�� = str��ҩĿ��
                        rsAdvice!ҽ������ = !ҽ������ & ""
                        rsAdvice!������� = !������� & ""
                        rsAdvice.Update
                        '----------------------------------------------------------------------------
                        lngCount = lngCount + 1
                        .MoveNext
                    Next
                End With
            End If
            If strҽ��IDs <> "" Then
                Set rsTmp = GetDrugInfo_MK4(gobjPati.str�Һŵ�, strҽ��IDs)
                rsAdvice.Filter = ""
                For i = 1 To rsAdvice.RecordCount
                    rsTmp.Filter = "ID=" & rsAdvice!ҽ��ID
                    If Not rsTmp.EOF Then
                        rsAdvice!������� = rsTmp!������� & ""
                        rsAdvice!�÷� = rsTmp!�÷� & ""
                        rsAdvice!ִ�п���ID = rsTmp!ִ�п���ID & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
            'Drip ����
            If Mid(strGroupIDs, 2) <> "" Then
                Set rsTmp = Get����(strGroupIDs)
                For i = 1 To rsTmp.RecordCount
                    rsAdvice.Filter = "���ID =" & rsTmp!id
                    Do While Not rsAdvice.EOF
                        rsAdvice!���� = rsTmp!ҽ������ & ""
                        rsAdvice.MoveNext
                    Loop
                    rsTmp.MoveNext
                Next
                rsAdvice.Filter = ""
            End If
        End If
        '�޿�����ҩƷ
        If lngCount = 0 Then
            Screen.MousePointer = 0: Exit Function
        End If
        
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        Call AdviceCheckWarn_MK4(gobjPati.lng����ID, gobjPati.str�Һŵ�, 0, bytShow, bytSubmit, rsAdvice, str��ʾ, lngResult)
        
        arrSQL = Array()
        '��ȡ��ʾ����
        '����ֵ˳��0-����,1-�ڵ�,2-���,3-�ȵ�,4-�Ƶ�
        '��ʾ��˳��0-����,4-�Ƶ�,3-�ȵ�,2-���,1-�ڵ�(��ΪPASS������ԭ��)
        arrLevel(0) = 0: arrLevel(1) = 4: arrLevel(2) = 3: arrLevel(3) = 2: arrLevel(4) = 1
        arrLight(0) = "��_4": arrLight(1) = "��_4": arrLight(2) = "��_4": arrLight(3) = "��_4": arrLight(4) = "��_4"
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_����༭ Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            Else
                blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) _
                Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4"))
                blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
                
            If blnDo Then
                If glngModel = PM_����ҽ���嵥 Then
                    strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    strҽ��ID = .RowData(i)
                End If
                rsAdvice.Filter = "ҽ��ID = '" & strҽ��ID & "'"
                If rsAdvice.RecordCount > 0 Then
                     k = CLng(rsAdvice!��ʾ & "")
                Else
                     k = -1 'ҽ���嵥��ҩ�䷽
                End If
               
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = k
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                    End If

                    If PM_����༭ = glngModel Then
                        If strOld <> CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) Then
                            .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                            blnNoSave = True    '���Ϊδ����
                        End If
                        '��¼�½���ҩƷ K=1 ����ڵ� �� ֻ���δУ��ҽ�����н���ҩƷ˵��ԭ��ı��,�Ѿ�У�Է��͵�ҽ��������
                        If k = 1 And Not rsOut Is Nothing Then
                            rsOut.Filter = "ҽ��ID = " & strҽ��ID & " And ״̬ < 3 "
                            If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                        End If
                    ElseIf PM_����ҽ���嵥 = glngModel Then
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                        End If

                    End If
                ElseIf .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                    '��ҩ�䷽
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                        lng��ҩ��ID = .TextMatrix(i, gobjCOL.intCOL���ID)          '��ҩ�䷽��ID
                        lngLight = -1 '��ʼ��
                    End If
                    '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                    If k >= 0 Then
                        If lngLight >= 0 Then
                            If arrLevel(k) > arrLevel(lngLight) Then
                                lngLight = k
                            End If
                        Else
                            lngLight = k
                        End If
                    End If
                End If
                '��¼��߼���ʾֵ
                If k >= 0 Then
                    If lngMaxWarn >= 0 Then
                        If arrLevel(k) > arrLevel(lngMaxWarn) Then
                            lngMaxWarn = k
                        End If
                    Else
                        lngMaxWarn = k
                    End If
                End If
            Else
                If glngModel = PM_����༭ Then
                    '��ҩ��ʾ�Ƶ�������
                    If .RowData(i) = lng��ҩ��ID And .RowData(i) <> 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                        '���þ�ʾ��
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                        End If
                        
                        If glngModel = PM_����༭ Then
                            '���������仯,�Ա��������ݿ�
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            '��¼�½���ҩƷ =1����ڵ�
                            If lngLight = 1 And Not rsOut Is Nothing Then
                                rsOut.Filter = "ҽ��ID = " & lng��ҩ��ID & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        End If
                        lng��ҩ��ID = 0
                        lngLight = -1
                    End If
                End If
            End If
        Next
        'ҽ���嵥��ҩ�䷽��ʾ�ƴ���
        If glngModel = PM_����ҽ���嵥 And Not rs��ҩ Is Nothing Then
            For i = .FixedRows To .Rows - 1
                '��ҩ����
                If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL��ʾ)
                    lngLight = -1
                    strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    rs��ҩ.Filter = "���ID=" & strҽ��ID
                    
                    For j = 1 To rs��ҩ.RecordCount
                        rsAdvice.Filter = "ҽ��ID = '" & rs��ҩ!id & "'"
                        If rsAdvice.RecordCount > 0 Then
                             k = CLng(rsAdvice!��ʾ & "")
                        Else
                             k = -1 'ҽ���嵥��ҩ�䷽
                        End If
                        '���þ�ʾ�� ȡ��ҩ�����ʾֵ
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                        rs��ҩ.MoveNext
                    Next
                    
                    '���þ�ʾ��
                    If lngLight >= 0 And lngLight <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = CStr(lngLight)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL��ʾ) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL��ʾ) = Nothing
                    End If
                    '��ʾ�Ƹ��µ����ݿ�
                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                    End If
                        
                    '��¼��߼���ʾֵ
                    If lngLight >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = lngLight
                            End If
                        Else
                            lngMaxWarn = lngLight
                        End If
                    End If
                    
                End If
            Next
        End If
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
        
    End With
    
    '���������
    OutAdviceCheckWarn_MK4 = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function AdviceCheckWarn_DT(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String, _
                Optional ByVal strҽ��IDs As String) As Boolean
'���ܣ����ô�ͨ��ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim strҩƷ As String, str��ҩ;�� As String, strƵ�ʱ��� As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '�ж��Ƿ����Ժ��ִ�е�ҩƷ
    Dim lng�Һ�ID As Long
    
    Set rsPati = GetPatiInfo_YF(lng����ID, str�Һŵ�, lng��ҳID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    curDate = zlDatabase.Currentdate
    With xmlbase
        If str�Һŵ� = "" Then
            .dInHosCode = rsPati!סԺ�� & ""
            .dBedNo = "" & rsPati!��ǰ����
        Else
            .dInHosCode = ""
            .dBedNo = ""
            .pOutID = str�Һŵ�
            lng�Һ�ID = NVL(rsPati!����Id, 0)
        End If
        .pCaseID = lng����ID
        .dDoctCode = UserInfo.�û���
        .dDoctName = UserInfo.����
        .dDoctType = UserInfo.רҵ����ְ��
        .dDeptCode = UserInfo.����ID
        .dDeptName = UserInfo.������
        .mPresDate = curDate
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(rsPati!��������, vbNull)
        .pPatiName = rsPati!����
        .pSex = rsPati!�Ա�
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
            
        '* ����Դ
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        
        If str�Һŵ� <> "" Then
            Set rsTmp = Get���˹�����¼(lng����ID, 0)
        Else
            Set rsTmp = Get���˹�����¼(lng����ID, lng��ҳID)
        End If
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!ҩ��ID
            .pCaseName1 = rsTmp!ҩ����
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!ҩ��ID
                .pCaseName2 = rsTmp!ҩ����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!ҩ��ID
                    .pCaseName3 = rsTmp!ҩ����
                End If
            End If
        End If
        
        '* �����Ϣ
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        If str�Һŵ� <> "" Then
            Set rsTmp = Get������ϼ�¼(lng����ID, lng�Һ�ID, "1,11")
        Else
            Set rsTmp = Get������ϼ�¼(lng����ID, lng��ҳID, "2,12")
        End If
        If rsTmp.RecordCount > 0 Then
            .pDiagnose1 = "" & rsTmp!����
            .pDiagnoseName1 = "" & rsTmp!����
            rsTmp.MoveNext
            If Not rsTmp.EOF Then
                .pDiagnose2 = "" & rsTmp!����
                .pDiagnoseName2 = "" & rsTmp!����
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose3 = "" & rsTmp!����
                    .pDiagnoseName3 = "" & rsTmp!����
                End If
            End If
        End If
        
        '* ������״̬
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        If str�Һŵ� <> "" Then
            strTmp = Get���˲��������(lng����ID, 0)
        Else
            strTmp = Get���˲��������(lng����ID, lng��ҳID)
        End If
        
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    Set rsTmp = GetAdviceInfo_YF(lng����ID, lng��ҳID, str�Һŵ�, strҽ��IDs)
    With rsTmp
        For i = 1 To rsTmp.RecordCount
            Call GetƵ����Ϣ_����(rsTmp!Ƶ�� & "", 0, 0, "", IIf(rsTmp!������� & "" = "7", 2, 1), strƵ�ʱ���)
        
            xmlpre.PresID = rsTmp!ҽ��ID & ""  'û��ҽ��ID������ID
            If str�Һŵ� <> "" Then
                xmlpre.PresType = "mz"
                xmlpre.Current = 1
                xmlpre.Days = StrToXML(rsTmp!���� & "")
            Else
                xmlpre.PresType = IIf(rsTmp!ҽ����Ч & "" = "0", "L", "T")
                xmlpre.BTime = Format(rsTmp!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                xmlpre.ETime = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
                xmlpre.PresTime = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
            
            End If
            
            xmlpre.GeneralName = StrToXML(rsTmp!ҩƷ���� & "")
            xmlpre.HosMediCode = rsTmp!�շ�ϸĿid & ""
            xmlpre.MediName = StrToXML(rsTmp!ҩƷ���� & "")
            xmlpre.DCL = FormatEx(rsTmp!�������� & "", 5)
            xmlpre.PCDM = StrToXML(strƵ�ʱ���)
            xmlpre.Unit = StrToXML(rsTmp!������λ & "")
            xmlpre.GYTJ = rsTmp!�÷�ID & ""
            xmlpre.GroupNum = rsTmp!���ID & ""
            
            strXML = MakePresXML(xmlpre, 1)
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = strXML
            .MoveNext
        Next
    End With
    
    If UBound(arrTmp) >= 0 Then
        On Error GoTo errH
        strXML = MakeXML(xmlbase, arrTmp, 1)
        WriteLog "" & glngModel, "AdviceCheckWarn_DT", strXML
        
        strTmp = dtywzxUI(28676, 1, strXML) '��������
        WriteLog "" & glngModel, "AdviceCheckWarn_DT", strTmp

    End If
    AdviceCheckWarn_DT = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    AdviceCheckWarn_DT = False
End Function

Public Function AdviceCheckWarn_TYT_YF(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, _
                    ByVal lngCmd As Long, Optional ByVal lngCurrAdviceID As Long, Optional str��ʾ As String = "", Optional ByVal strҽ��IDs As String) As Long
'���ܣ�����̫Ԫͨϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'       0-��ҩ�淶;1-��ȡҽ�������,����д��ʾ��
'       2-ҩƷ��ʾ
'       3-ҽҩ֪ʶ��;4-ϵͳ����;5-�����ʾ�ƣ���ȡ��ʾ����
'
'����
'str��ʾ-���ؾ�ʾ������ʽ��ҽ��ID1:��ʾֵ1,ҽ��ID2:��ʾֵ2
    Dim strҽ������ As String, str����ҽ�� As String, strDescription As String
    Dim strSQL As String, strOrderInfo As String, strƵ�ʱ��� As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long
    Dim lng�Һ�ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strTmp As String, strOld As String
    Dim arrTmp As Variant, colAuditResult As Collection
    
    On Error GoTo errH
    Screen.MousePointer = 11


    Select Case lngCmd
    Case 0   '0-��ҩ�淶

        gobjPass.getPdssPrescription

    Case 1  '1-��ȡҽ�������,����д��ʾ��
        Set rsPati = GetPatiInfo_YF(lng����ID, str�Һŵ�, lng��ҳID)
        If rsPati.EOF Then Screen.MousePointer = 0: Exit Function
        '������Ϣ
        With udtPatiOrder
            '���˲�����Ϣ:����ID,����,�Ա� 1-Ů, 0-��, 2-���꣬���˳������ڣ���ʽ YYYY-MM-DD ��Ϊ�գ����
            .PatientID = lng����ID & ""
            .Pname = rsPati!���� & ""
            .pSex = IIf(rsPati!�Ա� & "" = "��", "0", IIf(rsPati!�Ա� & "" = "Ů", "1", "2"))
            .pdateOfBirth = Format(rsPati!��������, "yyyy-MM-dd")
            
            If str�Һŵ� <> "" Then
                lng�Һ�ID = NVL(rsPati!����Id, 0)
                '������Ϣ
                strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng����ID, lng�Һ�ID)
                rsTmp.Filter = "��Ŀ����='���'"
                If rsTmp.RecordCount <> 0 Then .pHeight = IIf(Val(rsTmp!��¼���� & "") = 0, "", rsTmp!��¼���� & "")
                rsTmp.Filter = "��Ŀ����='����'"
                If rsTmp.RecordCount <> 0 Then .pWeight = IIf(Val(rsTmp!��¼���� & "") = 0, "", rsTmp!��¼���� & "")
                .PvisitID = rsPati!����� & ""
                .SysFlag = "1"  '2-סԺҽ��վ��1-����ҽ��վ
            Else
                .pHeight = IIf(Val(rsPati!��� & "") = 0, "", rsPati!��� & "")
                .pWeight = IIf(Val(rsPati!���� & "") = 0, "", rsPati!���� & "")
                .PvisitID = rsPati!סԺ�� & ""
                .SysFlag = "2"  '2-סԺҽ��վ��1-����ҽ��վ
            End If
            
             '���˲����������
            strTmp = Get���˲��������(lng����ID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
            .isLact = IIf(InStr(strTmp, "������") > 0, "1", "0")    '�Ƿ��飬��Ϊ1����Ϊ0 ��Ϊ��
            .isPregnant = IIf(InStr(strTmp, "�и�") > 0, "1", "0")    '�Ƿ��и�����Ϊ1 ����Ϊ0 ��Ϊ��
            .isLiverWhole = IIf(InStr(strTmp, "�ι����쳣") > 0, "1", "0") '�Ƿ�ι��쳣 1-�쳣��0-���� ��Ϊ��
            .isKidneyWhole = IIf(InStr(strTmp, "�������쳣") > 0, "1", "0") '�Ƿ������쳣 1-�쳣��0-���� ��Ϊ��
                
            '��¼ҽ����Ϣ
            .DoctDeptID = UserInfo.����ID & ""
            .DoctDeptName = UserInfo.������ & ""
            .DoctID = UserInfo.��� & ""
            .DoctName = UserInfo.���� & ""
            .DoctTitleID = GetDoctorTitleType(UserInfo.רҵ����ְ��)
            .DoctTitleName = IIf(UserInfo.רҵ����ְ�� = "", "����ְ��", UserInfo.רҵ����ְ��)
           
        End With

        'ҩƷ��Ϣ
        arrTmp = Array()
        
        Set rsAdvice = GetAdviceInfo_YF(lng����ID, lng��ҳID, str�Һŵ�, strҽ��IDs)
        If rsAdvice.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        With rsAdvice
            If NVL(!����ҽ��) <> "" Then
                str����ҽ�� = NVL(!����ҽ��)
                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                strҽ������ = Sys.RowValue("��Ա��", str����ҽ��, "���", "����")
            End If
            
            For i = 1 To .RecordCount
                udtDrug.drugID = !�շ�ϸĿid & ""    'his ϵͳ��ҩƷ���벻Ϊ��
                udtDrug.DrugName = StrToXML(!ҩƷ���� & "")               'his ϵͳ��ҩƷ���Ʋ�Ϊ��
                udtDrug.recMainNo = !���ID & ""     'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                udtDrug.recSubNo = !ҽ����� & ""      'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψ
                udtDrug.dosage = !�������� & ""     'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
    
                udtDrug.doseUnits = !������λ & ""    'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
                udtDrug.administrationID = !�÷�ID & ""               'his ϵͳ��ҽ��;�����벻Ϊ��
                strƵ�ʱ��� = GetFrequency(!�����λ & "", !Ƶ�ʴ��� & "", !Ƶ�ʼ�� & "")
                udtDrug.performFreqDictID = StrToXML(strƵ�ʱ���)   'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
                udtDrug.performFreqDictText = !Ƶ�� & ""               'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
    
                udtDrug.startDateTime = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DDHH: MM: SS ��Ϊ��
                udtDrug.stopDateTime = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")    'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
                udtDrug.doctorDept = !��������id & ""               'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
                udtDrug.DoctorID = strҽ������                          'his ϵͳ�Ŀ�ҽ��ҽ������
                udtDrug.Doctor = str����ҽ��                         'his ϵͳ�Ŀ�ҽ��ҽ������,
                udtDrug.isNew = "0"                             '����ҽ��ֵΪ1������Ϊ0
               
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtDrug
                .MoveNext
            Next
        End With
           
        If UBound(arrTmp) = -1 Then
            Screen.MousePointer = 0: Exit Function
        End If
        udtPatiOrder.PatDrugs = arrTmp

        '���
        arrTmp = Array()
  
        If str�Һŵ� <> "" Then
            Set rsTmp = Get������ϼ�¼(lng����ID, lng�Һ�ID, "1,11")
            strTmp = "�������"
        Else
            Set rsTmp = Get������ϼ�¼(lng����ID, lng��ҳID, "2,12")   '��ҽסԺ����ҽסԺ
            strTmp = "��Ժ���"
        End If
        
        For i = 0 To rsTmp.RecordCount - 1
            udtPatiDiag.diagnosisID = rsTmp!���� & ""       'his ϵͳ����ϱ���
            udtPatiDiag.diagnosisName = rsTmp!���� & ""     'his ϵͳ���������
            udtPatiDiag.diagnosisType = strTmp      'ϵͳ��������ͣ���������ϡ���Ժ��ϵ�
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = udtPatiDiag
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatDiagnoses = arrTmp
        
        
        '����
        arrTmp = Array()
        If str�Һŵ� <> "" Then
            Set rsTmp = Get���˹�����¼(lng����ID, 0)
        Else
            Set rsTmp = Get���˹�����¼(lng����ID, lng��ҳID)
        End If
        For i = 0 To rsTmp.RecordCount - 1
            udtPatiSensitive.patOrderDrugSensitiveID = "0"          '�̶�ֵ
            udtPatiSensitive.drugAllergenID = rsTmp!����Դ���� & ""    'ϵͳ�Ĺ�������
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = udtPatiSensitive
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatDrugSensitives = arrTmp

        '֢״
        arrTmp = Array()
        If str�Һŵ� <> "" Then
            Set rsTmp = GetPatiSymptom(lng����ID, lng�Һ�ID)
        Else
            Set rsTmp = GetPatiSymptom(lng����ID, lng��ҳID)
        End If
        For i = 0 To rsTmp.RecordCount - 1
            UdtPatiSymptom.symptomID = rsTmp!���� & ""              'his ϵͳ��֢״����
            UdtPatiSymptom.symptomName = rsTmp!���� & ""            'his ϵͳ��֢״����

            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = UdtPatiSymptom
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatSymptoms = arrTmp

        strOrderInfo = MakePatientOrderXml(udtPatiOrder)

        'ҽ����Ϣ���ӿڵ���"

        strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

        '���������
        '����ֵ˳����ʾ����(�ߵ���)��1�� ���ɣ�������ʾ��ɫ��ʾ�ƣ���2�� ���ã�������ʾ��ɫ��ʾ��ʾ����3�� ��ʾ��������ʾ��ɫ��ʾ�ƣ�
        If strDescription = "" Then
            MsgBox "ҩ����鹦��δִ�У�����̫Ԫͨ�ӿ������Ƿ�����", vbInformation + vbOKOnly, G_STR_PASS
            Screen.MousePointer = 0: Exit Function

        ElseIf strDescription = "-101" Then
            '-101����ʾ�û����Ժ��Ը÷���ֵ������ҵ����
        Else
            If str��ʾ <> "-1" Then
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                With rsAdvice
                    .MoveFirst
                    str��ʾ = ""
                    For i = 1 To rsAdvice.RecordCount
                        '��ȡ��ʾ��
                        strTmp = !���ID & "_" & !ҽ�����  '�ؼ��ָ�ʽ:��ҽ����_ҽ�����
                        On Error Resume Next
                        udtAuditResult = colAuditResult(strTmp)
                        If Err.Number > 0 Then
                            strTmp = "δ�ҵ�"
                        End If
                        Err.Clear: On Error GoTo 0
                        If strTmp <> "δ�ҵ�" Then  '�ҵ���˾�ʾ��
                            str��ʾ = str��ʾ & "," & !ҽ��ID & ":" & Val(udtAuditResult.alertLevel)
                        End If

                        .MoveNext
                    Next
                    If str��ʾ <> "" Then str��ʾ = Mid(str��ʾ, 2)
                End With
            Else
                str��ʾ = ""
            End If
        End If

    Case 2    ' 2-ҩƷ��ʾ

        '����ҩƷ��ʾ�ӿ�
        strSQL = "Select �շ�ϸĿid From ����ҽ����¼ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngCurrAdviceID)
        If rsTmp.RecordCount = 0 Then Exit Function
        gobjPass.getDrugExplain (rsTmp!�շ�ϸĿid & "")
      
    Case 3    '3-����ҽҩ֪ʶ��
        '��������ҽҩ֪ʶ��
        gobjPass.accessIFMI ("0")  '����ֵ�̶�Ϊ:"0",�޷���ֵ
    Case 4  '4-ϵͳ����
        gobjPass.sysConfig
    Case 5    '5-��ȡ��ʾ����
        gobjPass.getDrugAlertDetail
    End Select

    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PASS4���������(ByVal strData As String, ByRef lng���� As Long, ByRef lng���� As Long, _
        ByRef lng�ι� As Long, ByRef lng���� As Long, ByRef str�������� As String)
'����: ��ȡ���������
'����״̬��ȡֵ��-1�޷���ȡ����״̬��Ĭ�ϣ�;0����;1��
'����״̬��ȡֵ��-1-�޷���ȡ����״̬��Ĭ�ϣ�;0-����;1-��
'���￪ʼ���ڣ���ʽΪyyyy-mm-dd��
'���˸��𺦳̶ȣ�ȡֵ�� -1-��ȷ����Ĭ�ϣ���0-�޸��𺦣�1-�ι��ܲ�ȫ��2-��ȸ��𺦣�3-�жȸ��𺦣�4-�ضȸ���*/
'�������𺦳̶ȣ�ȡֵ�� -1-��ȷ����Ĭ�ϣ���0-�����𺦣�1-�����ܲ�ȫ��2-������𺦣�3-�ж����𺦣�4-�ض�����*/
    Dim i  As Integer
    
    lng���� = 0: lng�ι� = 0: lng���� = 0: lng���� = 0: str�������� = ""
    If strData = "" Then Exit Sub
    For i = LBound(Split(strData, ",")) To UBound(Split(strData, ","))
        If Split(strData, ",")(i) = "����" Then
            lng���� = 1
        ElseIf Split(strData, ",")(i) = "����" Then
            lng���� = 1
        ElseIf InStr("������,�����ܲ�ȫ,�������,�ж�����,�ض�����", Split(strData, ",")(i)) > 0 Then
            If Split(strData, ",")(i) = "������" Then
                lng���� = 0
            ElseIf Split(strData, ",")(i) = "�����ܲ�ȫ" Then
                lng���� = 1
            ElseIf Split(strData, ",")(i) = "�������" Then
                lng���� = 2
            ElseIf Split(strData, ",")(i) = "�ж�����" Then
                lng���� = 3
            ElseIf Split(strData, ",")(i) = "�ض�����" Then
                lng���� = 4
            End If
        ElseIf InStr("�޸���,�ι��ܲ�ȫ,��ȸ���,�жȸ���,�ضȸ���", Split(strData, ",")(i)) > 0 Then
            If Split(strData, ",")(i) = "�޸���" Then
                lng�ι� = 0
            ElseIf Split(strData, ",")(i) = "�ι��ܲ�ȫ" Then
                lng�ι� = 1
            ElseIf Split(strData, ",")(i) = "��ȸ���" Then
                lng�ι� = 2
            ElseIf Split(strData, ",")(i) = "�жȸ���" Then
                lng�ι� = 3
            ElseIf Split(strData, ",")(i) = "�ضȸ���" Then
                lng�ι� = 4
            End If
        ElseIf InStr(Split(strData, ",")(i), "��������|") > 0 Then
            str�������� = Split(Split(strData, ",")(i), "|")(1)
        End If
    Next
End Sub

Public Function AdviceCheckWarn_DTBS(ByVal bytFunc As Byte, Optional ByVal blnUpLoad As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
    Optional ByRef objMap As clsPassMap, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, _
    Optional ByVal str�Һŵ� As String, Optional ByVal strҽ��IDs As String) As Boolean
'���ܣ����ô�ͨ��ҩ���ϵͳ(BS��)��ҽ�����к�����ҩ������ع���
'
'������
'bytFunc=1-ҽ������վ;2-ҩ��
'blnUpLoad:�Ƿ��ϴ� T-��;F-��
'
'���Σ�
'      rsOut=����ҩƷ˵��
    Dim udtDetail As DTBS_DETAILS, xmlpre As dt_Pres
    Dim udtPati As DTBS_PATIENT
    Dim udt��� As DTBS_DIAGNOSE
    Dim udt����Դ As DTBS_ALLERGIC
    Dim udtPres As DTBS_PRESCRIPTION
    Dim udtMedic As DTBS_MEDICINE
    
    Dim colTmp As Collection, colPres As Collection
    Dim str��� As String, str���� As String
    Dim strTmp As String, arrTmp As Variant, curDate As Date

    Dim i As Long, j As Long, blnDo As Boolean
    Dim lngTmp As Long, lngPos As Long
    Dim strҩƷ As String, str��ҩ;�� As String, strXML As String
    Dim str���ID As String
    Dim strSQL As String
    
    Dim rsPati As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '�ж��Ƿ����Ժ��ִ�е�ҩƷ
    Dim lng�Һ�ID As Long
    Dim byt���� As Byte
    
    If bytFunc = 1 Then
        lng����ID = gobjPati.lng����ID
        lng��ҳID = gobjPati.lng��ҳID
        str�Һŵ� = gobjPati.str�Һŵ�
    End If
    
    Set rsPati = GetPatiInfo_YF(lng����ID, str�Һŵ�, lng��ҳID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    If str�Һŵ� <> "" Then
        lng�Һ�ID = Val(rsPati!����Id & "")
        strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
                        
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng����ID, lng�Һ�ID)
        rsPatiInfo.Filter = "��Ŀ����='���'"
        If rsPatiInfo.RecordCount <> 0 Then str��� = NVL(rsPatiInfo!��¼����)
        rsPatiInfo.Filter = "��Ŀ����='����'"
        If rsPatiInfo.RecordCount <> 0 Then str���� = NVL(rsPatiInfo!��¼����)
    Else
        str��� = rsPati!��� & ""
        str���� = rsPati!���� & ""
    End If
    
    curDate = zlDatabase.Currentdate
    
    With udtPati
        .str���� = rsPati!���� & ""
        .str�Ƿ�Ӥ�� = 0  '0:  ��Ӥ�׶��� 1�� ��Ӥ�׶�
        .str�������� = rsPati!�������� & ""
        .str�Ա� = rsPati!�Ա� & ""
        .str���� = str����
        .str��� = str���
        .str���֤�� = rsPati!���֤�� & ""
        .str������ = ""
        .str���� = ""
        .str����ʱ�䵥λ = ""
        .str����ʱ�� = ""
        
        '����Դ
        Set colTmp = New Collection
        Set rsTmp = Get���˹�����¼(lng����ID, lng��ҳID)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!ҩ��ID <> "" Then
                With udt����Դ
                    .str�������� = "5"   '1=��ͨҩƷ���� 2=��ͨҩƷ�ɷ� 5-HISҩƷ����
                    .str����Դ���� = rsTmp!ҩ����
                    .str����Դ���� = "" & rsTmp!ҩ��ID
                End With
                colTmp.Add udt����Դ, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col����Դs = colTmp
        
        '��ϼ�¼
        Set colTmp = New Collection
        Select Case glngModel
        Case PM_����༭
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udt���
                        If gobjDiags.Item(i).str������� <> "" Then
                            If gobjDiags.Item(i).str�������� <> "" Then
                                .str������� = "2" '2=IDC10����
                                .str��ϴ��� = gobjDiags.Item(i).str��������
                            Else
                                .str������� = "0"
                                .str��ϴ��� = gobjDiags.Item(i).str��ϱ���
                            End If
                            .str������� = gobjDiags.Item(i).str�������
                        End If
                    End With
                    colTmp.Add udt���, "_" & colTmp.Count + 1
                Next
            End If
        Case Else
            Set rsTmp = Get������ϼ�¼(lng����ID, IIf(str�Һŵ� <> "", lng�Һ�ID, lng��ҳID), IIf(str�Һŵ� <> "", "1,11", "2,12"))
            For i = 1 To rsTmp.RecordCount
                With udt���
                    If rsTmp!����ID & "" <> "" Then
                         .str������� = "2" '2=IDC10����
                    Else
                        .str������� = "0" '0=����
                    End If
                    .str��ϴ��� = "" & rsTmp!����
                    .str������� = "" & rsTmp!����
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
            If colTmp.Count = 0 Then
                colTmp.Add udt���, "_" & colTmp.Count + 1
            End If
        End Select
        '������
        strTmp = Get���˲��������(lng����ID, IIf(str�Һŵ� <> "", 0, lng��ҳID))
        If strTmp <> "" Then
            Set rsSub = GetRS("���������", "����,����", strTmp, "����", 0, 1)
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt���
                    .str������� = "1" '1=������״̬
                    .str������� = arrTmp(i)
                    rsSub.Filter = "����='" & arrTmp(i) & "'"
                    If Not rsSub.EOF Then .str��ϴ��� = rsSub!���� & ""
                End With
                colTmp.Add udt���, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col���s = colTmp
        '�����ⵥ�ڵ�
        'Set colTmp = New Collection
        'strTmp = ""
    End With
    
    With udtDetail
        .str�Ƿ��ϴ� = IIf(blnUpLoad, "1", "0")  'Ĭ�ϲ��ϴ���������
        .strHISϵͳʱ�� = Format(curDate, "YYYY-MM-dd hh:mm:ss")
        If str�Һŵ� <> "" Then
            .str����סԺ��ʶ = "op"
            .str�������� = DTBS_GetTreatType(1, lng�Һ�ID)
            .str����� = lng�Һ�ID & ""
        Else
            .str����סԺ��ʶ = "ip"
            .str�������� = DTBS_GetTreatType(2, lng����ID, lng��ҳID)
            .str����� = rsPati!סԺ�� & ""
        End If
        .udt������Ϣ = udtPati
    End With

    'ҩƷ��Ϣ
    Select Case glngModel
    Case PM_����༭, PM_����ҽ���嵥, PM_סԺ�༭, PM_סԺҽ���嵥
        Set rsTmp = CreateAdviceRS(rsOut)
    Case PM_���ŷ�ҩ, PM_������ҩ, PM_PIVA����
        Set rsTmp = CreateAdviceRS(, lng����ID, lng��ҳID, str�Һŵ�, strҽ��IDs)
        byt���� = 1
    End Select
    
    If rsTmp.RecordCount = 0 Then AdviceCheckWarn_DTBS = True: Exit Function    'ҽ���´����û���´�ҩƷʱ������
    
    With rsTmp
        Set colPres = New Collection
        .MoveFirst
        For i = 1 To rsTmp.RecordCount
            udtPres.str������ = !ҽ��ID & ""
            udtPres.str�������� = ""
            udtPres.str����ҽ������ = !����ҽ������ & ""
            udtPres.str����ҽ������ = !����ҽ�� & ""
            udtPres.str�������Ҵ��� = !��������id & ""
            udtPres.str������������ = !�������� & ""
            udtPres.str����ʱ�� = Format(!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
            udtPres.str�Ƿ�������� = IIf(!��־ & "" = "1", "1", "0")
            udtPres.str�Ƿ��¿����� = IIf(Val(!ҽ��״̬ & "") < 2, "1", "0") '�ݴ�,�¿� ����1
            udtPres.str�Ƿ�ǰ���� = IIf(byt���� = 1, 1, IIf(Val(!ҽ��״̬ & "") < 2, "1", "0")) '0 ��ʷ���� 1 ��ǰ�����¿�����(=1ʱ���������Ż᷵���������)
            udtPres.Strҽ������ = IIf(!ҽ����Ч & "" = "1", "T", "L") 'סԺ������Ч(hosp_flag = ip)L:����ҽ�� T: ��ʱҽ��
            
            Set colTmp = New Collection
            udtMedic.str��Ʒ�� = DTBS_StrToXML(!ҩƷ���� & "")
            udtMedic.strҽԺҩƷ���� = !ҩƷID & ""
            udtMedic.str��Һ���� = ""
            udtMedic.str��Һ����� = ""
            udtMedic.strҽ������ = ""
            udtMedic.str��� = !��� & ""
            udtMedic.str��� = !���ID & ""
            udtMedic.str��ҩ���� = !��ҩ���� & ""   '�ǿ���ҩ��Ϊ��
            '������������λ
            udtMedic.str��������λ = !������λ & ""
            udtMedic.str������ = !�������� & ""
            udtMedic.strƵ�δ��� = !Ƶ�ʱ��� & ""
            udtMedic.str��ҩ;������ = !�÷�ID & ""
            udtMedic.str��ҩ��ʼʱ�� = Format(!��ʼʱ�� & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str��ҩ����ʱ�� = Format(!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str��ҩ���� = !���� & ""   'OP ���ﴦ����Ч
            udtMedic.str�Ƿ�Ԥ����ҩ = IIf(!��ҩĿ�� & "" = "1", 1, 0)
            udtMedic.str�������� = ""
            udtMedic.strǩ��ҽʦ���� = ""
            udtMedic.str��Ȩʱ�� = ""
            udtMedic.str������ҩʱ�� = ""
            udtMedic.str������ҩ���� = ""
            colTmp.Add udtMedic, "_" & colTmp.Count

            Set udtPres.colҩƷ��Ϣ = colTmp
            colPres.Add udtPres, "_" & colPres.Count + 1
            .MoveNext
        Next
        .Filter = "��Ժ��ҩ=1"
        blnIsHaveOut = .RecordCount > 0
    End With
    
    Set udtDetail.col������Ϣ = colPres
    
    If udtPres.colҩƷ��Ϣ.Count > 0 Then
        On Error GoTo errH
        strXML = DTBS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "���ܺ�:" & DTBS_��������
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "DetailXML:" & strXML
    
        lngTmp = CRMS_UI(DTBS_��������, gstrBaseXml, strXML, strRetXML)
        strRetXML = StrConvToNormal(strRetXML)
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "������������ֵ:" & lngTmp & vbCrLf & "RetXML:" & strRetXML
      
        If blnUpLoad Then AdviceCheckWarn_DTBS = True: Exit Function
       
        If glngModel = PM_����༭ Then
            '���ݷ�Ϊ���֣�0��1��2��3�ֱ����û�����⣬�������⣬һ���������������
            '1�� ���洦����Ϣ�Ĺ����У��Ӵ�ͨ��ҩ��ȫ���ϵͳ�õ��ķ���ֵ�����0����1����ʾ��ǰ������û���������⡣��0��ʾû�����⣻1��ʾ�������⣬�������ⶼ�Ƕ�ҽ������ʾ��
            '2�� �õ��ķ���ֵ�����2��3����ʾ��ǰ�������������⣬��Ҫ�Ըô����������ء�
           If lngTmp = 3 And gbytBlackLamp = 0 Then
                MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            ElseIf ((lngTmp = 2 Or lngTmp = 3) And gbytBlackLamp = 1) Then
                If gbytReason = 1 Then
                    '��¼�½���ҩƷ ��ʾֵ=2\3������� �� ֻ���δУ��ҽ�����н���ҩƷ˵��ԭ��ı��,�Ѿ�У�Է��͵�ҽ��������
                    Set rsRet = ReadXML(strRetXML)
                    If Not rsRet Is Nothing Then
                        For i = 1 To rsRet.RecordCount
                            If Not rsOut Is Nothing Then
                                If rsRet!��ʾֵ >= 2 Then
                                    rsOut.Filter = "ҽ��ID = " & rsRet!ҽ��ID & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then
                                        rsOut!�Ƿ���� = 1
                                    End If
                                End If
                            End If
                            rsRet.MoveNext
                        Next
                    End If
                    If Not AddDrugReason(objMap, rsOut) Then Exit Function
                Else
                    If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf lngTmp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�������������⣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        ElseIf glngModel = PM_סԺ�༭ Then
            If lngTmp = 3 And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("��ҩ���ϵͳ������Ժ��ִ�е�ҩƷ���ڽ�����ҩ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���������ܼ���!", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf (lngTmp = 2 Or lngTmp = 3) And gbytBlackLamp = 1 Then
                If gbytReason = 1 Then
                    Set rsRet = ReadXML(strRetXML)
                    If Not rsRet Is Nothing Then
                        For i = 1 To rsRet.RecordCount
                            If Not rsOut Is Nothing Then
                                If rsRet!��ʾֵ >= 2 Then
                                    rsOut.Filter = "ҽ��ID = " & rsRet!ҽ��ID & " And ״̬ < 3 "
                                    If rsOut.RecordCount = 1 Then
                                        rsOut!�Ƿ���� = 1
                                    End If
                                End If
                            End If
                            rsRet.MoveNext
                        Next
                    End If
                    If Not AddDrugReason(objMap, rsOut) Then Exit Function
                Else
                    If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�����ڽ�����ҩ���Ƿ����?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf lngTmp = 1 Then
                If MsgBox("��ҩ���ϵͳ���ֵ�ǰҽ�������������⣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If

    AdviceCheckWarn_DTBS = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    AdviceCheckWarn_DTBS = False
End Function

Public Function CreateAdviceRS(Optional ByRef rsOut As ADODB.Recordset, Optional ByVal lng����ID As String, _
    Optional ByVal lng��ҳID As String, Optional ByVal str�Һŵ� As String, _
    Optional ByVal strҽ��IDs As String) As ADODB.Recordset
'����;����ҽ����¼��
    Dim i As Long, k As Long, lngCount As Long, lngPos As Long
    Dim blnDo As Boolean, blnIsHaveOut As Boolean
    Dim strҩƷ As String, strҽ��ID As String, str���ID As String
    Dim str����ʱ�� As String
    Dim str��Ч As String, str���� As String, str������λ As String, strƵ�� As String
    Dim str��ҩ;�� As String, strƵ�ʱ��� As String, str�÷� As String, str�÷�ID As String, str��ʼʱ�� As String, str����ʱ�� As String
    Dim str��������Tag As String, str��������ID As String, str������ĿIDs As String, strҩƷID As String
    Dim str����ҽ��Tag As String, str����ҽ�� As String
    Dim str���� As String, str������λ As String, str״̬ As String
    Dim str����ID, str�շ�ϸĿID As String
    
    Dim rsAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsƵ�� As ADODB.Recordset
    Dim rs����ҽ�� As ADODB.Recordset
    Dim rs��������  As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim rsҩƷ As ADODB.Recordset
    
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ_DTBS)
    
    Select Case glngModel
    Case PM_����༭, PM_סԺ�༭
        '�����˽���ҩƷ˵������;����Ϊ����༭\סԺ�༭;��鹦��
        If (glngModel = PM_����༭ Or glngModel = PM_סԺ�༭) And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_�������)
        End If
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" _
                            Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                End If
                
                If blnDo Then
                    str����ID = .TextMatrix(i, gobjCOL.intCOL������ĿID)
                    If InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                        If InStr("," & str������ĿIDs & ",", "," & str����ID & ",") = 0 Then
                            str������ĿIDs = str������ĿIDs & "," & str����ID
                        End If
                    End If
                    strҽ��ID = CStr(.RowData(i))
                    
                    'ȡҩƷ����
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 Then
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҩƷ����)
                    Else
                        strҩƷ = .TextMatrix(i, gobjCOL.intCOLҽ������) '��ҩ����
                    End If
                    
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str�÷� = ""    'һ����ҩ���ظ�ȡ
                    If str�÷� = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                        If k <> -1 Then
                            If .TextMatrix(i, gobjCOL.intCOL�������) = "7" Then
                                str�÷� = .TextMatrix(k, gobjCOL.intCOL�÷�)
                            Else
                                str�÷� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                            End If
                            str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                        End If
                    End If
    
                    '������������
                    str��������ID = .TextMatrix(i, gobjCOL.intCOL��������ID)
                    If InStr("," & str��������Tag & ",", "," & str��������ID & ",") = 0 Then
                        str��������Tag = str��������Tag & "," & str��������ID
                    End If
                   
                    '����ҽ��
                    str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                        str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                    End If
                    
                    str��ʼʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd HH:MM:SS")
'
                    str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:mm:ss")         '����ʱ�䣨YYYY-MM-DD HH:mm:SS��
                    '������������λ
                    str���� = .TextMatrix(i, gobjCOL.intCOL����)
                    str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                    str���� = .TextMatrix(i, gobjCOL.intCOL����)
                    str������λ = .TextMatrix(i, gobjCOL.intcol������λ)
                    
                    strҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    
                    If glngModel = PM_����༭ Then
                        str����ʱ�� = ""
                        str��Ч = "1"
                    ElseIf glngModel = PM_סԺ�༭ Then
                        str����ʱ�� = Format(.Cell(flexcpData, i, gobjCOL.intCOL��ֹʱ��), "yyyy-MM-dd HH:MM:SS")
                        '�ж��Ƿ���Ժ��ִ�е�ҩƷ
                        If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                            blnIsHaveOut = True
                        End If
                        str��Ч = IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", 0, 1)
                    End If
                    
                    If InStr(";" & strƵ�� & ";", ";" & .TextMatrix(i, gobjCOL.intCOLƵ��) & "," & IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1) & ";") = 0 Then
                        strƵ�� = strƵ�� & ";" & .TextMatrix(i, gobjCOL.intCOLƵ��) & "," & IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1)
                    End If
                    
                    '����˵��
                    If Not rsOut Is Nothing Then
                        If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 Then
                        '��ҩ,�г�ҩ
                            rsOut.AddNew
                            rsOut!ҽ��ID = CLng(strҽ��ID)
                            rsOut!����ҩƷ˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                            rsOut!״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                            rsOut!ҩƷ���� = .TextMatrix(i, gobjCOL.intCOLҽ������)
                            rsOut.Update
                        ElseIf Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then
                        '��ҩ�䷽  ����˵����������ҩ������
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                            If k <> -1 Then
                                rsOut.AddNew
                                rsOut!ҽ��ID = CLng(.RowData(k) & "")
                                rsOut!����ҩƷ˵�� = .TextMatrix(k, gobjCOL.intCol����ҩƷ˵��)
                                rsOut!״̬ = .TextMatrix(k, gobjCOL.intCOL״̬)
                                rsOut!ҩƷ���� = .TextMatrix(k, gobjCOL.intCOLҽ������)
                                rsOut.Update
                            End If
                        End If
                    End If
                        
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = strҽ��ID
                    rsAdvice!���ID = .TextMatrix(i, gobjCOL.intCOL���ID)
                    rsAdvice!ҽ����Ч = str��Ч
                    rsAdvice!ҽ����� = lngCount + 1
                    rsAdvice!��������id = str��������ID
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!������ĿID = str����ID
                    rsAdvice!ҩƷID = strҩƷID
                    rsAdvice!ҩƷ���� = strҩƷ
                    rsAdvice!ҽ��״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                    rsAdvice!�������� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!Ƶ�� = .TextMatrix(i, gobjCOL.intCOLƵ��)
                    rsAdvice!�÷� = str�÷�
                    rsAdvice!�÷�ID = str��ҩ;��
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!��ʼʱ�� = str��ʼʱ��
                    rsAdvice!����ʱ�� = str����ʱ��
                    rsAdvice!���� = str����
                    rsAdvice!������λ = str������λ
                    rsAdvice!���� = .TextMatrix(i, gobjCOL.intCOL����)
                    rsAdvice!ҽ������ = .TextMatrix(i, gobjCOL.intcolҽ������)
                    rsAdvice!��ҩĿ�� = .TextMatrix(i, gobjCOL.intcol��ҩĿ��)
                    rsAdvice!��ҩ���� = .TextMatrix(i, gobjCOL.intcol��ҩ����)
                    rsAdvice!������� = .TextMatrix(i, gobjCOL.intCOL�������)
                    rsAdvice!��־ = .TextMatrix(i, gobjCOL.intCol��־)
                    rsAdvice!��Ժ��ҩ = IIf(blnIsHaveOut, 1, 0)
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                End If
            Next
        End With
    Case PM_����ҽ���嵥, PM_סԺҽ���嵥
        Set rsTmp = GetAdviceInfo_YF(gobjPati.lng����ID, gobjPati.lng��ҳID, gobjPati.str�Һŵ�, , 1)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS = rsAdvice: Exit Function
            For i = 1 To .RecordCount
                If glngModel = PM_����ҽ���嵥 Then
                    blnDo = InStr(",5,6,7,", "," & !������� & ",") > 0 And Val(!�շ�ϸĿid & "") <> 0 And Format(!����ʱ�� & "", "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                ElseIf glngModel = PM_סԺҽ���嵥 Then
                    blnDo = InStr(",5,6,7,", "," & !������� & ",") > 0 And Not InStr(",4,8,9,", "," & !ҽ��״̬ & ",") > 0
                End If
                If blnDo Then
              
                    If InStr(",5,6,7,", "," & !������� & ",") > 0 And Not InStr(",4,8,9,", "," & !ҽ��״̬ & ",") > 0 And Val(!�շ�ϸĿid & "") = 0 Then
                        If InStr("," & str������ĿIDs & ",", "," & !������ĿID & ",") = 0 Then
                            str������ĿIDs = str������ĿIDs & "," & !������ĿID
                        End If
                    End If
                    '����ҽ��
                    str����ҽ�� = !����ҽ�� & ""
                    If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                    If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                        str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                    End If
                    
                    If gobjPati.str�Һŵ� <> "" Then
                        str������λ = !���ﵥλ & ""
                    Else
                        str������λ = !סԺ��λ & ""
                    End If
                    
                    If InStr(";" & strƵ�� & ";", ";" & !Ƶ�� & "," & IIf(!������� & "" = "7", 2, 1) & ";") = 0 Then
                        strƵ�� = strƵ�� & ";" & !Ƶ�� & "," & IIf(!������� = "7", 2, 1)
                    End If
                    
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = !ҽ��ID & ""
                    rsAdvice!���ID = !���ID & ""
                    rsAdvice!ҽ����Ч = !ҽ����Ч & ""
                    rsAdvice!ҽ����� = lngCount + 1
                    rsAdvice!��������id = !��������id & ""
                    rsAdvice!�������� = !�������� & ""
                    rsAdvice!����ҽ�� = str����ҽ��
                    rsAdvice!������ĿID = !������ĿID & ""
                    rsAdvice!ҩƷID = !�շ�ϸĿid & ""
                    rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                    rsAdvice!ҽ��״̬ = !ҽ��״̬ & ""
                    rsAdvice!�������� = !�������� & ""
                    rsAdvice!������λ = !������λ & ""
                    rsAdvice!Ƶ�� = !Ƶ�� & ""
                    rsAdvice!�÷� = !�÷� & ""
                    rsAdvice!�÷�ID = !�÷�ID & ""
                    rsAdvice!����ʱ�� = !����ʱ�� & ""
                    rsAdvice!��ʼʱ�� = !��ʼʱ�� & ""
                    rsAdvice!����ʱ�� = !����ʱ�� & ""
                    rsAdvice!���� = !���� & ""
                    rsAdvice!������λ = str������λ
                    rsAdvice!���� = !���� & ""
                    rsAdvice!ҽ������ = !ҽ������ & ""
                    rsAdvice!��ҩĿ�� = !��ҩĿ�� & ""
                    rsAdvice!��ҩ���� = !��ҩ���� & ""
                    rsAdvice!������� = !������� & ""
                    rsAdvice!��� = !��� & ""
                    rsAdvice!��־ = !��־ & ""
                    rsAdvice.Update
                End If
                .MoveNext
            Next
        End With
    Case PM_PIVA����, PM_���ŷ�ҩ, PM_������ҩ
        Set rsTmp = GetAdviceInfo_YF(lng����ID, lng��ҳID, str�Һŵ�)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS = rsAdvice: Exit Function
            For i = 1 To .RecordCount
            
                If Val(!�շ�ϸĿid & "") = 0 Then
                    If InStr("," & str������ĿIDs & ",", "," & !������ĿID & ",") = 0 Then
                        str������ĿIDs = str������ĿIDs & "," & !������ĿID
                    End If
                End If
                '����ҽ��
                str����ҽ�� = !����ҽ�� & ""
                If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                    str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                End If
                
                If str�Һŵ� <> "" Then
                    str������λ = !���ﵥλ & ""
                Else
                    str������λ = !סԺ��λ & ""
                End If
                
                If InStr(";" & strƵ�� & ";", ";" & !Ƶ�� & "," & IIf(!������� & "" = "7", 2, 1) & ";") = 0 Then
                    strƵ�� = strƵ�� & ";" & !Ƶ�� & "," & IIf(!������� = "7", 2, 1)
                End If
                
                rsAdvice.AddNew
                rsAdvice!ҽ��ID = !ҽ��ID & ""
                rsAdvice!���ID = !���ID & ""
                rsAdvice!ҽ����Ч = !ҽ����Ч & ""
                rsAdvice!ҽ����� = lngCount + 1
                rsAdvice!��������id = !��������id & ""
                rsAdvice!�������� = !�������� & ""
                rsAdvice!����ҽ�� = str����ҽ��
                rsAdvice!������ĿID = !������ĿID & ""
                rsAdvice!ҩƷID = !�շ�ϸĿid & ""
                rsAdvice!ҩƷ���� = !ҩƷ���� & ""
                rsAdvice!ҽ��״̬ = !ҽ��״̬ & ""
                rsAdvice!�������� = !�������� & ""
                rsAdvice!������λ = !������λ & ""
                rsAdvice!Ƶ�� = !Ƶ�� & ""
                rsAdvice!�÷� = !�÷� & ""
                rsAdvice!�÷�ID = !�÷�ID & ""
                rsAdvice!����ʱ�� = !����ʱ�� & ""
                rsAdvice!��ʼʱ�� = !��ʼʱ�� & ""
                rsAdvice!����ʱ�� = !����ʱ�� & ""
                rsAdvice!���� = !���� & ""
                rsAdvice!������λ = str������λ
                rsAdvice!���� = !���� & ""
                rsAdvice!ҽ������ = !ҽ������ & ""
                rsAdvice!��ҩĿ�� = !��ҩĿ�� & ""
                rsAdvice!��ҩ���� = !��ҩ���� & ""
                rsAdvice!������� = !������� & ""
                rsAdvice!��� = !��� & ""
                rsAdvice!��־ = !��־ & ""
                rsAdvice.Update

                .MoveNext
            Next
        End With
    End Select
    
    '����������ȡ
    If rsAdvice.RecordCount > 0 Then
        
        rsAdvice.MoveFirst
        Select Case glngModel
        
        Case PM_����༭, PM_����ҽ���嵥, PM_סԺ�༭, PM_סԺҽ���嵥, PM_PIVA����, PM_���ŷ�ҩ, PM_������ҩ
            If str������ĿIDs <> "" Then
                str������ĿIDs = Mid(str������ĿIDs, 2)
                Set rsҩƷ = GetRS("ҩƷ���", "ҩ��id,ҩƷid", str������ĿIDs, "ҩ��id")
            End If
            If strƵ�� <> "" Then Set rsƵ�� = GetRS("����Ƶ����Ŀ", "����, ����, ���÷�Χ", strƵ��, "����, ���÷�Χ", 1, 2)
            If str��������Tag <> "" Then Set rs�������� = GetRS("���ű�", "ID,����", str��������Tag)
            If str����ҽ��Tag <> "" Then Set rs����ҽ�� = GetRS("��Ա��", "���,����", str����ҽ��Tag, "����", 0, 1)
            For i = 1 To rsAdvice.RecordCount
                 '����ҽ����Ʒ���´�ʱ,����ȡһ��ҩƷId
                If Val(rsAdvice!ҩƷID & "") = 0 And Val(rsAdvice!ҽ����Ч & "") = 0 Then
                    If Not rsҩƷ Is Nothing Then
                        rsҩƷ.Filter = "ҩ��ID =" & rsAdvice!������ĿID
                        If Not rsҩƷ.EOF Then rsAdvice!ҩƷID = rsҩƷ!ҩƷID & ""
                    End If
                End If
                
                If InStr("," & str�շ�ϸĿID & ",", "," & rsAdvice!ҩƷID & ",") = 0 Then
                    str�շ�ϸĿID = str�շ�ϸĿID & "," & rsAdvice!ҩƷID
                End If
                
                If Not rsƵ�� Is Nothing Then
                    rsƵ��.Filter = "���� ='" & rsAdvice!Ƶ�� & "' And ���÷�Χ=" & IIf(rsAdvice!������� & "" = "7", 2, 1)
                    If Not rsƵ��.EOF Then rsAdvice!Ƶ�ʱ��� = rsƵ��!���� & ""
                End If
                
                If Not rs����ҽ�� Is Nothing Then
                    rs����ҽ��.Filter = "����='" & rsAdvice!����ҽ�� & "'"
                    If Not rs����ҽ��.EOF Then rsAdvice!����ҽ������ = rs����ҽ��!��� & ""
                End If
                If Not rs�������� Is Nothing Then
                    rs��������.Filter = "ID =" & rsAdvice!��������id
                    If Not rs��������.EOF Then rsAdvice!�������� = rs��������!���� & ""
                End If
                
                rsAdvice.MoveNext
            Next
            
            If str�շ�ϸĿID <> "" Then
                str�շ�ϸĿID = Mid(str�շ�ϸĿID, 2)
                Set rs��� = GetRS("�շ���ĿĿ¼", "ID,���", str�շ�ϸĿID)
                rsAdvice.MoveFirst
                For i = 1 To rsAdvice.RecordCount
                    If Not rs��� Is Nothing Then
                        rs���.Filter = "ID =" & rsAdvice!ҩƷID
                        If Not rs���.EOF Then rsAdvice!��� = rs���!��� & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
        End Select
        rsAdvice.MoveFirst
    End If
    Set CreateAdviceRS = rsAdvice
End Function

Private Function GetDrugInfo_MK4(ByVal str�Һŵ� As String, ByVal strAdvice As String, Optional ByVal lngPatiID As Long, Optional ByVal lng��ҳID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.���id, a.�������, d.���� As �÷�,a.ִ�п���ID " & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ D" & vbNewLine & _
            "Where a.���id = b.Id(+) And b.������Ŀid = d.Id(+) " & IIf(lngPatiID = 0, "And A. �Һŵ� = [1]", "And A.����ID = [3] And A.��ҳID =[4]") & " And a.���id <> 0 And Instr([2], ',' || a.Id || ',') > 0" & vbNewLine & _
            "Order By a.���"
    Set GetDrugInfo_MK4 = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", str�Һŵ�, "," & strAdvice & ",", lngPatiID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
