Attribute VB_Name = "mdlPassDefine_ZL"
Option Explicit

Public gobjFrm As frmPass      '����������
Public grsRet       As ADODB.Recordset   '���浱ǰ�������һ�������,�л��������;�����ϴ�ʱ�ṩ��ʾ����
Public Const conMenu_EditPopup = 3    '�༭
Public Const conMenu_Drug_View = 30821 '�鿴ҩƷ˵����
Public Const conMenu_Drug_Match = 5 '�������
Public Const conMenu_PAR_SET = 6    '��������
Public Const conMenu_FRM_VISIBLE = 7 '����\��ʾ������
Public Const conCOLOR_BULE As Long = &HD48A00
Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)
Public Const conCOLOR_BULELIGHT As Long = &HE4B440

Public Const conSTR_Key_Tip     As String = "��Ӧ֢,�÷�����,������Ӧ,����֢,ע������,�и���ҩ,��ͯ��ҩ,��������ҩ,�໥����,ҩ�����"
Public gstrParaTip As String

'Public gobjAir As zl9ComLib.clsAirBubble      'zl9ComLib.clsAirBubble
Public gobjAir As Object        '������ʾ

Private mstrPharmDept   As String    '����ҩʦ�����ÿ���
Private mstrPassDept    As String    '������ҩ�����鹦�����ÿ���
Private mstrHosName     As String    'ҽԺ����

Public Sub ZLShowWindow()
    If gobjFrm Is Nothing Then Set gobjFrm = New frmPass
    Call gobjFrm.Show
End Sub

Public Sub ZLCloseWindow()
    Unload gobjFrm
    Set gobjFrm = Nothing
End Sub
 
Public Function ZLGetDrugCode(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String, _
    Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef blnIsHaveOut As Boolean, _
    Optional ByRef rsOut As ADODB.Recordset, Optional ByVal bytFunc As Byte = 0) As String
'����:������Ҫ����ҩƷ��λ��
'   bytFunc=3 �ϴ��������
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer
    Dim blnDo As Boolean, blnAsk As Boolean
    
    Dim str�շ�ϸĿIDs As String
    Dim str������ĿIDs As String
    Dim str��ҩ;��    As String, str��Ч As String
    Dim strƵ�ʱ���    As String, str�����λ As String
    Dim str��ҩ��IDs    As String, str���ID As String
    Dim strҽ��ID       As String, str���� As String, str������λ As String
    Dim strҽ��IDs      As String
    
    Dim str����ҽ�� As String, str����ҽ��Tag As String
    Dim str������ID As String   '��¼���ID��������ȡ��Ӧ���
    Dim rsDoct As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    
    Dim curDate As Date

    On Error GoTo errH
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_ҽ����Ϣ_ZL)
    '�����˽���ҩƷ˵������  �ҳ���Ϊ����༭��鹦��
    If (glngModel = PM_סԺ�༭ Or glngModel = PM_����༭) And (gbytReason = 1 Or gbytReason = 0 And InStr("," & mstrPharmDept & ",", "," & gobjPati.lngDeptID & ",") > 0) Then
        Set rsOut = InitAdviceRS(FUN_�������)
    End If
    Select Case glngModel
    Case PM_����༭, PM_סԺ�༭
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_סԺ�༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    If blnDo Then
                        blnDo = (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL״̬)) = 0 _
                                Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL״̬) <> "4")
                    End If
                End If
                '����������Ҫ����������ĿID �ų�ҩƷ
                If gstrIP <> "" And bytFunc = 0 Then
                    blnAsk = Not InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL������ĿID)) <> 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1
                Else
                    blnAsk = False
                End If
                If blnDo Then
                    If Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                        str�շ�ϸĿIDs = str�շ�ϸĿIDs & IIf(str�շ�ϸĿIDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    ElseIf Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) = 0 Then
                        If InStr("," & str������ĿIDs & ",", "," & .TextMatrix(i, gobjCOL.intCOL������ĿID) & ",") = 0 Then
                            str������ĿIDs = str������ĿIDs & IIf(str������ĿIDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL������ĿID)
                        End If
                    End If
                    If glngModel = PM_סԺ�༭ Then
                        If Val(.TextMatrix(i, gobjCOL.intCOLִ������)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID))), gobjCOL.intCOLִ������)) = 5 Then
                            blnIsHaveOut = True
                        End If
                    End If
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL���ID)) Then str��ҩ;�� = "" 'һ����ҩ���ظ�ȡ
                    If str��ҩ;�� = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL���ID)), i + 1)
                        If k <> -1 Then str��ҩ;�� = Val(.TextMatrix(k, gobjCOL.intCOL������ĿID))   '������
                    End If
                    Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
                    
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = .RowData(i)
                    rsAdvice!������ = .TextMatrix(i, gobjCOL.intCOL����)
                    rsAdvice!������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                    rsAdvice!������ĿID = .TextMatrix(i, gobjCOL.intCOL������ĿID)
                    rsAdvice!ҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                    rsAdvice!��Һ��� = .TextMatrix(i, gobjCOL.intCOL���ID)
                    rsAdvice!��ҩƵ�� = strƵ�ʱ���
                    rsAdvice!��ҩƵ������ = .TextMatrix(i, gobjCOL.intCOLƵ��)
                    rsAdvice!��ҩ;�� = str��ҩ;��
                    rsAdvice!ÿ���� = Getÿ����(.TextMatrix(i, gobjCOL.intCOL����), str�����λ, intƵ�ʴ���, intƵ�ʼ��, .TextMatrix(i, gobjCOL.intCOLƵ��))
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
                    If bytFunc = 3 Then
                        strҽ��IDs = strҽ��IDs & "," & .RowData(i)
                        rsAdvice!����ʱ�� = Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd HH:MM:SS")
                        rsAdvice!������־ = IIf(.TextMatrix(i, gobjCOL.intCol��־) = "1", "1", "0") '0-��ͨ,1-������2-��¼
                        rsAdvice!ҽ��״̬ = .TextMatrix(i, gobjCOL.intCOL״̬)
                        
                        '����ҽ��
                        str����ҽ�� = .TextMatrix(i, gobjCOL.intCOL����ҽ��)
                        If InStr(str����ҽ��, "/") > 0 Then str����ҽ�� = Mid(str����ҽ��, 1, InStr(str����ҽ��, "/") - 1)
                        If InStr("," & str����ҽ��Tag & ",", "," & str����ҽ�� & ",") = 0 Then
                            str����ҽ��Tag = str����ҽ��Tag & "," & str����ҽ��
                        End If
                        rsAdvice!����ҽ�� = str����ҽ��
                        rsAdvice!ҽ������ = .TextMatrix(i, gobjCOL.intCOLҽ������)
                        If glngModel = PM_����༭ Then
                            rsAdvice!��ҩ���� = .TextMatrix(i, gobjCOL.intCOL����)   'OP ���ﴦ����Ч
                            '�������
                            If InStr("," & str������ID & ",", "," & .TextMatrix(i, gobjCOL.intCOL���ID) & ",") = 0 Then
                                str������ID = str������ID & "," & .TextMatrix(i, gobjCOL.intCOL���ID)
                            End If
                            rsAdvice!ҽ����Ч = "1"
                        Else
                            rsAdvice!ҽ����Ч = IIf(.TextMatrix(i, gobjCOL.intCOL��Ч) = "����", "1", "0")
                        End If

                        '����ҩ���Ԥ��������
                        If .TextMatrix(i, gobjCOL.intcol��ҩĿ��) = "1" Then
                            rsAdvice!��ҩĿ�� = "Ԥ��"
                        ElseIf .TextMatrix(i, gobjCOL.intcol��ҩĿ��) = "2" Then
                            rsAdvice!��ҩĿ�� = "����"
                        Else
                            rsAdvice!��ҩĿ�� = ""
                        End If
                        '�ϴ����ʱ��������
                        If Not grsRet Is Nothing Then
                            grsRet.Filter = "OrderId =" & .RowData(i)
                            If Not grsRet.EOF Then
                                rsAdvice!ҩƷ���ɵȼ� = IIf(grsRet!Level & "" = "", "����", grsRet!Level & "")
                                rsAdvice!ҩƷ�������� = grsRet!Type & ""
                            End If
                        End If
                        rsAdvice!ҩƷ����˵�� = .TextMatrix(i, gobjCOL.intCol����ҩƷ˵��)
                    End If
                ElseIf blnAsk Then
                    rsAdvice.AddNew
                    rsAdvice!ҽ��ID = .RowData(i)
                    rsAdvice!������ĿID = .TextMatrix(i, gobjCOL.intCOL������ĿID)
                    rsAdvice!���� = 1 '��ʶ��������
                End If
            Next
            If rsAdvice.RecordCount > 0 Then rsAdvice.UpdateBatch
        End With
    Case PM_����ҽ���嵥, PM_סԺҽ���嵥
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
               If glngModel = PM_����ҽ���嵥 Then
                    blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 _
                        Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4")) _
                        And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                            
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
                    If (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4") Then
                        str��ҩ��IDs = str��ҩ��IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
        
                        Call GetƵ����Ϣ_����(.TextMatrix(i, gobjCOL.intCOLƵ��), intƵ�ʴ���, intƵ�ʼ��, str�����λ, IIf(.TextMatrix(i, gobjCOL.intCOL�������) = "7", 2, 1), strƵ�ʱ���)
                        If Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0 Then
                            str�շ�ϸĿIDs = str�շ�ϸĿIDs & IIf(str�շ�ϸĿIDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                        End If
                        '������������λ
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                        strҽ��IDs = strҽ��IDs & "," & strҽ��ID
                        If glngModel = PM_����ҽ���嵥 Then
                            str���� = Val(.TextMatrix(i, gobjCOL.intCOL����))
                            str������λ = .TextMatrix(i, gobjCOL.intCOL����)
                            str���� = FormatEx(str����, 5)
                            str������λ = Replace(str������λ, str����, "")
                        Else
                            str���� = .TextMatrix(i, gobjCOL.intCOL����)
                            str������λ = .TextMatrix(i, gobjCOL.intCOL������λ)
                            str���� = Replace(str����, str������λ, "")
                        End If
                        rsAdvice.AddNew
                        rsAdvice!ҽ��ID = strҽ��ID
                        rsAdvice!������ = str����
                        rsAdvice!������λ = str������λ
                        rsAdvice!������ĿID = .TextMatrix(i, gobjCOL.intCOL������ĿID)
                        rsAdvice!ҩƷID = .TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)
                        rsAdvice!��Һ��� = .TextMatrix(i, gobjCOL.intCOL���ID)
                        rsAdvice!��ҩƵ�� = strƵ�ʱ���
                        rsAdvice!��ҩƵ������ = .TextMatrix(i, gobjCOL.intCOLƵ��)
                        rsAdvice!��ҩ;�� = ""
                        rsAdvice!ÿ���� = Getÿ����(str����, str�����λ, intƵ�ʴ���, intƵ�ʼ��, .TextMatrix(i, gobjCOL.intCOLƵ��))
                        rsAdvice.Update
                        
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
                    End If
                End If
            Next
            '����ҽ���嵥�䷽��������,��Ҫ�����ݿ���ȡ��ҩ����
            If glngModel = PM_סԺҽ���嵥 Or glngModel = PM_����ҽ���嵥 Then
                If str��ҩ��IDs <> "" Then
                    Set rsDrug = Get��ҩ�䷽(str��ҩ��IDs)
                    With rsDrug
                        For i = 1 To .RecordCount
                            If !���ID & "" <> str���ID Then
                                Call GetƵ����Ϣ_����(!Ƶ�� & "", 0, 0, "", IIf(!������� & "" = "7", 2, 1), strƵ�ʱ���)
                                str���ID = !���ID & ""
                            End If
                            rsAdvice.AddNew
                            rsAdvice!ҽ��ID = !id & ""
                            rsAdvice!������ = !�������� & ""
                            rsAdvice!������λ = !������λ & ""
                            rsAdvice!������ĿID = !������ĿID & ""
                            rsAdvice!ҩƷID = !ҩƷID & ""
                            rsAdvice!��Һ��� = !���ID & ""
                            rsAdvice!��ҩƵ�� = strƵ�ʱ���
                            rsAdvice!��ҩ;�� = !�÷�ID & ""
                            rsAdvice!ÿ���� = Getÿ����(str����, str�����λ, intƵ�ʴ���, intƵ�ʼ��, !Ƶ�� & "")
                            rsAdvice.Update
                            .MoveNext
                        Next
                    End With
                End If
            End If
     
        End With
    End Select
    If strҽ��IDs <> "" Then
        Set rsDrug = GetDrugPlus(lng����ID, lng��ҳID, str�Һŵ�, strҽ��IDs)
        rsAdvice.Filter = "����=0"
        For i = 1 To rsAdvice.RecordCount
            rsDrug.Filter = "ID=" & rsAdvice!ҽ��ID
            If Not rsDrug.EOF Then
                rsAdvice!��ҩ;�� = rsDrug!��ҩ;��ID & ""
                If bytFunc = 3 Then
                    rsAdvice!��ҩ;������ = rsDrug!��ҩ;������ & ""
                    rsAdvice!���� = rsDrug!ҩƷ���� & ""
                    rsAdvice!������� = rsDrug!������� & ""
                    rsAdvice!����˵�� = rsDrug!����˵�� & ""
                    rsAdvice!ҩƷ����ҩ��ȼ� = rsDrug!������ & ""
                End If
            End If
            rsAdvice.MoveNext
        Next
    End If
    If str�շ�ϸĿIDs <> "" Then
        Set rsDrug = GetRS("ҩƷ��� A", "A.ҩƷID,A.��λ��", str�շ�ϸĿIDs, "A.ҩƷID")
        rsAdvice.Filter = "����=0"
        For i = 1 To rsAdvice.RecordCount
            rsDrug.Filter = "ҩƷID=" & Val(rsAdvice!ҩƷID)
            If Not rsDrug.EOF Then
                rsAdvice!��λ�� = rsDrug!��λ�� & ""
            End If
            rsAdvice.MoveNext
        Next
    End If
    If str������ĿIDs <> "" Then
        Set rsDrug = GetRS("ҩƷ��� A,�շ���ĿĿ¼ B", "A.ҩ��ID,A.��λ��", str������ĿIDs, "A.ҩƷID=B.ID And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� IS NULL) And A.ҩ��ID")
        rsAdvice.Filter = "����=0"
        For i = 1 To rsAdvice.RecordCount
            If rsAdvice!��λ�� = "" Then
                rsDrug.Filter = "ҩ��ID=" & Val(rsAdvice!������ĿID)
                If Not rsDrug.EOF Then
                    rsAdvice!��λ�� = rsDrug!��λ�� & ""
                End If
            End If
            rsAdvice.MoveNext
        Next
    End If
    
    If str����ҽ��Tag <> "" And bytFunc = 3 Then
        str����ҽ��Tag = Mid(str����ҽ��Tag, 2)
        Set rsDoct = GetDoctorInfo(str����ҽ��Tag, IIf(glngModel = PM_����༭, 2, 1))
        rsAdvice.Filter = "����=0"
        For i = 1 To rsAdvice.RecordCount
            rsDoct.Filter = "����='" & rsAdvice!����ҽ�� & "'"
            If Not rsDoct.EOF Then
                rsAdvice!ҽ��ְ�� = rsDoct!Ƹ�μ���ְ�� & ""
                rsAdvice!ҽ������ҩ��ȼ� = rsDoct!���� & ""
            End If
            rsAdvice.MoveNext
        Next
    End If
    
    If glngModel = PM_����༭ And bytFunc = 3 And str������ID <> "" Then
        str������ID = Mid(str������ID, 2)
        Set rsDoct = GetOutAdviceDiagsInfo(str������ID)
        rsAdvice.Filter = "����=0"
        For i = 1 To rsAdvice.RecordCount
            rsDoct.Filter = "ҽ��ID=" & rsAdvice!��Һ���
            Do While Not rsDoct.EOF
                rsAdvice!������� = rsAdvice!������� & "|" & rsDoct!�������
                rsDoct.MoveNext
            Loop
            rsAdvice!������� = Mid(rsAdvice!������� & "", 2)
            rsAdvice.MoveNext
        Next
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AdviceCheckWarn_ZL(ByVal lngPatiID As Long, ByVal str�Һŵ� As String, ByVal lng��ҳID As String, _
     Optional ByVal bytFunc As Byte, Optional rsOut As ADODB.Recordset, Optional ByRef objMap As clsPassMap) As Boolean
'���ܣ�����������ҩ���ϵͳ��ҽ�����к�����ҩ������ع���
'
'������bytFunc=0 �༭�������;1-ҽ��վ���;3-ҽ���´���汣��ҽ�������,�������ϴ��������ϵͳ
'����ֵ:
    Dim strRet As String
    Dim strPara As String, strParaAsk As String
    Dim strUrl  As String
    Dim strҽ����Ч As String, str״̬ As String, str����ʱ�� As String, strҽ��ID As String
    Dim strOld As String
    
    Dim rsAdvice    As ADODB.Recordset
    
    Dim bytRet      As Byte
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Dim blnIsHaveOut As Boolean
    Dim blnDo As Boolean, blnNoSave As Boolean
    Dim arrSQL As Variant
    Dim arrLight(0 To 4) As String
    Dim datCurr As Date
    
    On Error GoTo errH
    If gblnBreak Then Exit Function
    strPara = ZL_MakeDetailXML(lngPatiID, lng��ҳID, str�Һŵ�, rsAdvice, rsOut, blnIsHaveOut, bytFunc)
    If strPara = "" Then AdviceCheckWarn_ZL = True: Exit Function
    '��������
LineAsk:
    If (glngModel = PM_����༭ Or glngModel = PM_סԺ�༭) And gstrIP <> "" And bytFunc = 0 Then
        strParaAsk = strPara
        Call AskPatiStatus(rsAdvice, strPara, lngPatiID)
    End If
    '������ҩ���
    strUrl = "http://" & gstrDrugIP & ":" & gstrDrugPort & "/DrugCorrect/CheckContent"
    WriteLog "" & glngModel, "AdviceCheckWarn_ZL", "���URL:" & strUrl & ",���XML:" & strPara
    strRet = HttpPost(strUrl, strPara, responseText, "text/plain", , , gblnBreak)
    strRet = Replace(strRet, """", "")
    '<errormsg>������Ϣ</errormsg>
    WriteLog "" & glngModel, "AdviceCheckWarn_ZL", "�����:" & strRet
    If strRet = "" Or InStr(strRet, "<errormsg>") > 0 Or gblnBreak Then
        gobjFrm.SetNotifyIcon
        gsngCheckLinkTime = Timer
        AdviceCheckWarn_ZL = True '�쳣�Ͽ�,������ҽ��
        Exit Function
    End If
    '�ϴ����ϵͳ
    If bytFunc = 3 Then Exit Function
    
    Set grsRet = ZL_ParseXML(strRet)
    If grsRet.RecordCount > 0 Then
        Call frmPassResultZL.ShowMe(gfrmMain, grsRet, bytFunc, bytRet, blnIsHaveOut)
        If bytRet = 3 Then
            strPara = strParaAsk
            GoTo LineAsk
        End If
    End If

    If bytFunc > 1 Then Exit Function
    With gobjAdvice
        arrSQL = Array()
        datCurr = zlDatabase.Currentdate
        '��ȡҽ�������,����д��ʾ��
        '-------------------------------------------------------------
        '����ֵ˳��0-����(Ĭ��),1-�ȵ�(���� �� ��),2-���(����),3-�Ƶ�(ע��),4-�ڵ�(��ֹ)
        '��ʾ��˳��0-����,3-�Ƶ�,1-�ȵ�,2-���,4-�ڵ�
        arrLight(0) = "��_4":    arrLight(1) = "��_4":  arrLight(2) = "��_4": arrLight(3) = "��_4": arrLight(4) = "��_4"
        If glngModel = PM_����༭ Or glngModel = PM_����ҽ���嵥 Then
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_����༭ Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")
                Else
                    blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL�������) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL�շ�ϸĿID)) <> 0) _
                    Or (.TextMatrix(i, gobjCOL.intCOL�������) = "E" And .TextMatrix(i, gobjCOL.intCol��������) = "4"))
                    blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL����ʱ��), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")
                End If
                    
                If blnDo Then
                    If glngModel = PM_����ҽ���嵥 Then
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        strҽ��ID = .RowData(i)
                    End If
                    grsRet.Filter = "OrderId = '" & strҽ��ID & "'"
                    grsRet.Sort = "WarnLevel DESC"
                    If grsRet.RecordCount > 0 Then
                         k = CLng(grsRet!Light & "")
                    Else
                         k = 0 'ҽ���嵥��ҩ�䷽
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
                            If strOld <> CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) Or Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            '��¼�½���ҩƷ K=2 ������ �� ֻ���δУ��ҽ�����н���ҩƷ˵��ԭ��ı��,�Ѿ�У�Է��͵�ҽ��������
                            If k = 2 And Not rsOut Is Nothing Then
                                rsOut.Filter = "ҽ��ID = " & strҽ��ID & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        ElseIf PM_����ҽ���嵥 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    End If
                End If
            Next
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        ElseIf glngModel = PM_סԺ�༭ Or glngModel = PM_סԺҽ���嵥 Then
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_סԺ�༭ Then
                    'סԺ�༭�������ҽ��ʱ�Ѿ����ε�����ҽ����ֹͣ��ȷ��ֹͣ�ĳ���
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL�������)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOLӤ��)) = gobjPati.intӤ�� _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOLѡ��) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL��Ч) = "����" Or .TextMatrix(i, gobjCOL.intCOL��Ч) = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd"))
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
                            (strҽ����Ч = "����" And (InStr(",8,9,", str״̬) > 0 And str����ʱ�� > Format(datCurr, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str״̬) > 0) Or _
                            strҽ����Ч = "����" And Format(.Cell(flexcpData, i, gobjCOL.intCOL��ʼʱ��), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")))
                    End If
                End If
                If blnDo Then
                    If glngModel = PM_סԺ�༭ Then
                        strҽ��ID = .RowData(i) & ""
                    Else
                        strҽ��ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If
                    grsRet.Filter = "OrderId='" & strҽ��ID & "'"
                    grsRet.Sort = "WarnLevel DESC"
                    If grsRet.RecordCount > 0 Then
                        k = CLng(grsRet!Light & "")
                    Else
                        k = 0
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
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Or Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL���) = 1
                                blnNoSave = True    '���Ϊδ����
                            End If
                            
                            If Not rsOut Is Nothing And k = 2 Then
                                rsOut.Filter = "ҽ��ID=" & CLng(strҽ��ID) & " And ״̬ < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!�Ƿ���� = 1
                            End If
                        ElseIf PM_סԺҽ���嵥 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL��ʾ)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & strҽ��ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    End If
                End If
            Next
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End If
    End With

    If bytRet = 1 Then  '�޸Ĵ���
        Exit Function
    ElseIf bytRet = 2 Then '������
        If bytFunc = 0 Then
            grsRet.Filter = "Light = 2"
            If grsRet.RecordCount > 0 Then
                If gbytBlackLamp = 1 Then
                    If (gbytReason = 1 Or gbytReason = 0 And InStr("," & mstrPharmDept & ",", "," & gobjPati.lngDeptID & ",") > 0) Then
                        If Not AddDrugReason(objMap, rsOut) Then Exit Function
                    Else
                        If MsgBox("��鷢�ֽ�����ҩ����ȷ��Ҫ������", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Exit Function
                        End If
                    End If
                Else
                    If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                        If MsgBox("����Ժ��ִ�е�ҩƷ��鷢�ֽ�����ҩ����ȷ��Ҫ������", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Exit Function
                        End If
                    End If
                End If
            Else
                If MsgBox("��鷢������ҩƷ����ȷ��Ҫ������", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                    Exit Function
                End If
            End If
        End If
    End If
    AdviceCheckWarn_ZL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function JSONParse(ByVal strJSONPath As String, ByVal strJSONData As String) As Variant
    Dim objJSON As Object
    Dim strValue As String
    
    On Error GoTo errH
    Set objJSON = CreateObject("MSScriptControl.ScriptControl")
    objJSON.Language = "JScript"
    strValue = NVL(objJSON.eval("JSON=" & strJSONData & ";JSON." & strJSONPath & ";"))
    JSONParse = JSONReplace(strValue)
    Set objJSON = Nothing
    Exit Function
errH:
    MsgBox "JSONParse �����:" & Err.Number & "��������:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function JSONReplace(ByVal strJson As String)
'����:JSON�������ַ���ת��
    If strJson <> "" Then
        strJson = Replace(strJson, "\n", vbLf)
        strJson = Replace(strJson, "\r", vbCr)
        strJson = Replace(strJson, "\t", vbTab)
        strJson = Trim(strJson)
    End If
    JSONReplace = strJson
End Function

Public Function ZL_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
10      On Error GoTo errH
20      strPara = zlDatabase.GetPara(90001, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
        '��ʽ������IP&&�������˿ں�
30      If strPara = "" Then Exit Function
40      arrList = Split(strPara, G_STR_SPLIT)
50      If UBound(arrList) >= 3 Then
60          gstrDrugIP = arrList(0)
70          gstrDrugPort = arrList(1)
80          If Val(arrList(2)) > 10 Then
90              gsngWaitTime = 10
100         ElseIf Val(arrList(2)) < 1 Then
110             gsngWaitTime = 1   '���ʵȴ�3s
120         Else
130             gsngWaitTime = Val(arrList(2))
140         End If
150         If Val(arrList(3)) > 10 Then
160             gsngAutoLinkTime = 10   '10����
170         ElseIf Val(arrList(3)) < 0 Then
180             gsngAutoLinkTime = 1   ' 1����
190         Else
200             gsngAutoLinkTime = Val(arrList(3))
210         End If
220     Else
230         gstrDrugIP = ""
240         gstrDrugPort = ""
250         gsngWaitTime = 3
260         gsngAutoLinkTime = 5
270         Exit Function
280     End If
290     mstrHosName = zlRegInfo("��λ����", , 0)
300     gstrIP = GetParaURL("֪ʶ��", "��������")
310     gstrStatusEdit = GetParaURL("֪ʶ��", "����״̬�༭")
320     gstrStatusGet = GetParaURL("֪ʶ��", "����״̬��ѯ")
330     gstrStatusSave = GetParaURL("֪ʶ��", "����״̬����")
340     gstrParaTip = zlDatabase.GetPara(299, glngSys)
350     strPara = GetParaURL("ҩʦ�������", "���ÿ��Ҳ�ѯ")
360     If strPara <> "" Then
370         strPara = HttpGet(strPara, responseText, 1)
            '{"items":[{"dept_ids":"168,143,159,149,151,156,148,158,"}],"first":{"$ref":"http://192.168.0.231:8080/ords/zlrecipe/recipe/getenabledept"}}
380         WriteLog "" & glngModel, "ZL_GetPara", "���ÿ��Ҳ�ѯ:" & strPara
390         If strPara <> "" Then
400             mstrPharmDept = JSONParse("items[0].dept_ids", strPara)
410         End If
420     End If
430     strPara = GetParaURL("֪ʶ��", "�����Ҳ�ѯ")
440     If strPara <> "" Then
450         strPara = HttpGet(strPara, responseText, 1)
            '{"items":[{"dept_ids":"132,122,433,473,138,148,149,129,151,168,515,147,144,143,159,514,152,146,157,141,135,145,155,557,150,163,154,106,219,235,236,226,230,237,228,238,513,224,229,231,234,"}],"hasMore":false,"limit":1000,"offset":0,"count":1,"links":[{"rel":"self","href":"http://192.168.0.231:8080/ords/rudrug/para/getenableddeptlist"},{"rel":"describedby","href":"http://192.168.0.231:8080/ords/rudrug/metadata-catalog/para/item"},{"rel":"first","href":"http://192.168.0.231:8080/ords/rudrug/para/getenableddeptlist"}]}
460         WriteLog "" & glngModel, "ZL_GetPara", "�����Ҳ�ѯ:" & strPara
470         If strPara <> "" Then
480             mstrPassDept = JSONParse("items[0].dept_ids", strPara)
490         End If
500     End If
510     ZL_GetPara = True
520     Exit Function
errH:
530     MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "ZL_GetPara:��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function ZL_SetPara() As String
    ZL_SetPara = IIf(gstrDrugIP = "", "192.168.6.17", gstrDrugIP) & G_STR_SPLIT & IIf(gstrDrugPort = "", "80", gstrDrugPort) & _
        G_STR_SPLIT & IIf(gsngWaitTime = 0, 3, gsngWaitTime) & G_STR_SPLIT & IIf(gsngAutoLinkTime = 0, 5, gsngAutoLinkTime)
End Function

Public Function ZL_MakeDetailXML(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, _
    Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef rsOut As ADODB.Recordset, Optional ByRef blnIsHaveOut As Boolean, _
    Optional ByVal bytFunc As Byte) As String
'���ܣ�����details XML�ַ���
    Dim strXML As String
    Dim strTmp As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    
    Dim colPati  As Collection
    Dim lng�Һ�ID As Long
    Dim i As Long
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    Set rsTmp = GetPatiInfo_YF(lng����ID, str�Һŵ�, lng��ҳID)
    If rsTmp.EOF Then Exit Function
    
    gobjPati.lngDeptID = Val(rsTmp!��ǰ����ID & "")
    If mstrPassDept <> "" And InStr("," & mstrPassDept & ",", "," & gobjPati.lngDeptID & ",") = 0 Then Exit Function
    
    Set colPati = New Collection
    If str�Һŵ� <> "" Then
        lng�Һ�ID = rsTmp!����Id
        '������Ϣ
        strSQL = "Select b.��Ŀ����, b.��¼����" & vbNewLine & _
                        "From ���˻����¼ A, ���˻������� B" & vbNewLine & _
                        "Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng����ID, lng�Һ�ID)
        rsPatiInfo.Filter = "��Ŀ����='���'"
        If rsPatiInfo.RecordCount > 0 Then
            colPati.Add rsPatiInfo!��¼���� & "", "���"
        Else
            colPati.Add "", "���"
        End If
        rsPatiInfo.Filter = "��Ŀ����='����'"
        If rsPatiInfo.RecordCount > 0 Then
            colPati.Add rsPatiInfo!��¼���� & "", "����"
        Else
            colPati.Add "", "����"
        End If
    Else
        colPati.Add rsTmp!��� & "", "���"
        colPati.Add rsTmp!���� & "", "����"
    End If
    
    colPati.Add rsTmp!���� & "", "����"
    colPati.Add rsTmp!�������� & "", "��������"
    colPati.Add Get��������(Val(rsTmp!�������� & "")), "��������"
    colPati.Add NVL(rsTmp!�Ա�), "�Ա�"
    colPati.Add NVL(rsTmp!ְҵ), "ְҵ"

    '�����������
    strTmp = "," & Get���˲��������(lng����ID, IIf(str�Һŵ� <> "", 0, lng��ҳID)) & ","
    If InStr(strTmp, ",����,") > 0 Then
        colPati.Add "1", "����"
    Else
        colPati.Add "0", "����"
    End If
    If InStr(strTmp, ",����,") > 0 Then
        colPati.Add "1", "����"
    Else
        colPati.Add "0", "����"
    End If
    If InStr(strTmp, ",�ι��ܲ�ȫ,") > 0 Then
        colPati.Add "1", "�ι��ܲ�ȫ"
    Else
        colPati.Add "0", "�ι��ܲ�ȫ"
    End If
    If InStr(strTmp, ",���ظι��ܲ�ȫ,") > 0 Then
        colPati.Add "1", "���ظι��ܲ�ȫ"
    Else
        colPati.Add "0", "���ظι��ܲ�ȫ"
    End If
    If InStr(strTmp, ",�����ܲ�ȫ,") > 0 Then
        colPati.Add "1", "�����ܲ�ȫ"
    Else
        colPati.Add "0", "�����ܲ�ȫ"
    End If
    If InStr(strTmp, ",���������ܲ�ȫ,") > 0 Then
        colPati.Add "1", "���������ܲ�ȫ"
    Else
        colPati.Add "0", "���������ܲ�ȫ"
    End If
    
    If bytFunc = 3 Then
        '���������Ϣ
        colPati.Add "2", "�ύ����"
        'colPati.Add lng����ID, "����ID"
        If str�Һŵ� <> "" Then
            'colPati.Add lng�Һ�ID, "����ID"
            colPati.Add rsTmp!����� & "", "�����"
            colPati.Add "1", "������Դ"   '1-����;2-סԺ
        Else
            'colPati.Add lng��ҳID, "����ID"
            colPati.Add rsTmp!סԺ�� & "", "סԺ��"
            colPati.Add rsTmp!��ǰ����ID & "", "���ﲡ��ID"
            colPati.Add rsTmp!��ǰ���� & "", "���ﲡ��"
            colPati.Add "2", "������Դ"   '1-����;2-סԺ
        End If
        colPati.Add rsTmp!���� & "", "����"
        colPati.Add Format(NVL(rsTmp!��������), "YYYY-MM-DD HH:MM:SS"), "��������"
        colPati.Add Format(NVL(rsTmp!��Ժʱ��), "YYYY-MM-DD HH:MM:SS"), "��Ժ����"
        colPati.Add rsTmp!��ǰ���� & "", "��ǰ����"
        colPati.Add rsTmp!��ǰ���� & "", "�������"
        colPati.Add rsTmp!��ǰ����ID & "", "�������ID"
        colPati.Add "0", "Ӥ��"     '1-Ӥ��;0-��Ӥ��
        colPati.Add "100", "HIS_NO"
    Else
        colPati.Add "1", "�ύ����"
    End If
    
    '�����Ϣ
    Set rsTmp = Get������ϼ�¼(lng����ID, IIf(str�Һŵ� <> "", lng�Һ�ID, lng��ҳID), IIf(str�Һŵ� <> "", "1,11", "2,12"))
    strTmp = "": strXML = ""
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & IIf(i = 1, "", ",") & rsTmp!����
        strXML = strXML & IIf(i = 1, "", ",") & rsTmp!����
        rsTmp.MoveNext
    Next
    colPati.Add strXML, "�������"
    colPati.Add strTmp, "���"
    
    'ҽ����Ϣ
    Call ZLGetDrugCode(lng����ID, lng��ҳID, str�Һŵ�, rsAdvice, blnIsHaveOut, rsOut, bytFunc)
    strXML = ZL_GET_Details(colPati, rsAdvice, IIf(str�Һŵ� <> "", 1, 2), bytFunc)
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "������ʱ��XML:" & strXML
    
    gcnOracle.BeginTrans: blnTran = True
    'XMLд����ʱ��:����������ҩ����
    Call sys.SaveLob(glngSys, 30, "", strXML, 1)
    '���ù���:��д����ʱ��Ĳ������ݽ��и���
    strSQL = "Zl_����������ҩ����_Update(" & lng����ID & "," & lng��ҳID & "," & lng�Һ�ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, G_STR_PASS)
    '��ȡ��ʱ������:����������ҩ����
    strXML = ReadLobForPASS()
    gcnOracle.CommitTrans: blnTran = False
    
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "������ʱ��XML:" & strXML
    strXML = "<root><����ID_IN></����ID_IN><ҽԺ����_IN>" & mstrHosName & "</ҽԺ����_IN>" & strXML & "</root>"
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "�������ӿ�XML:" & strXML
    ZL_MakeDetailXML = strXML
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZL_GET_PatiInfo(ByVal colPati As Collection, ByVal byt���� As Byte) As String
 
    Dim strXML As String
'    <patient_info>
'        <info name=���ύ���͡� value=��1��/> --1-�¿���2-����
'        <info name="��������" value="28114.45"/> --��������-sysdate
'        <info name="��������" value="����"/>
'        <info name="�Ա�" value="Ů"/>
'        <info name="ְҵ" value="�˶�Ա"/>
'        <info name="����" value="1"/>
'        <info name="����" value="1"/>
'        <info name="�ι��ܲ�ȫ" value="1">
'        <info name="���ظι��ܲ�ȫ" value="1">
'        <info name="�����ܲ�ȫ" value="1">
'        <info name="���������ܲ�ȫ" value="1">
'        <info name="���" value="J18.000"/> --��ϴ����룬�������Զ��ŷָ�
'    </patient_info>
    strXML = _
    "<patient_info>" & vbNewLine & _
            "    <info name=""�ύ����"" value=""" & colPati("�ύ����") & """/>" & vbNewLine & _
            "    <info name=""��������"" value=""" & colPati("��������") & """  unit=""��""/>" & vbNewLine & _
            "    <info name=""��������"" value=""" & colPati("��������") & """/>" & vbNewLine & _
            "    <info name=""����"" value=""" & colPati("����") & """/>" & vbNewLine & _
            "    <info name=""���"" value=""" & colPati("���") & """/>" & vbNewLine & _
            "    <info name=""����"" value=""" & colPati("����") & """/>" & vbNewLine & _
            "    <info name=""�Ա�"" value=""" & colPati("�Ա�") & """/>" & vbNewLine & _
            "    <info name=""ְҵ"" value=""" & colPati("ְҵ") & """/>" & vbNewLine & _
            "    <info name=""����"" value=""" & colPati("����") & """/>" & vbNewLine & _
            "    <info name=""����"" value=""" & colPati("����") & """/>" & vbNewLine & _
            "    <info name=""�ι��ܲ�ȫ"" value=""" & colPati("�ι��ܲ�ȫ") & """/>" & vbNewLine & _
            "    <info name=""���ظι��ܲ�ȫ"" value=""" & colPati("���ظι��ܲ�ȫ") & """/>" & vbNewLine & _
            "    <info name=""�����ܲ�ȫ"" value=""" & colPati("�����ܲ�ȫ") & """/>" & vbNewLine & _
            "    <info name=""���������ܲ�ȫ"" value=""" & colPati("���������ܲ�ȫ") & """/>" & vbNewLine & _
            "    <info name=""�������"" value=""" & colPati("�������") & """/>" & vbNewLine & _
            "    <info name=""���"" value=""" & colPati("���") & """/>"
        If colPati("�ύ����") = "2" Then
             strXML = strXML & vbNewLine & _
                "    <info name=""����"" value=""" & colPati("����") & """/>" & vbNewLine & _
                "    <info name=""��������"" value=""" & colPati("��������") & """/>" & vbNewLine & _
                "    <info name=""��Ժ����"" value=""" & colPati("��Ժ����") & """/>" & vbNewLine & _
                "    <info name=""��ǰ����"" value=""" & colPati("��ǰ����") & """/>" & vbNewLine & _
                "    <info name=""�������"" value=""" & colPati("�������") & """/>" & vbNewLine & _
                "    <info name=""�������ID"" value=""" & colPati("�������ID") & """/>" & vbNewLine & _
                "    <info name=""Ӥ��"" value=""" & colPati("Ӥ��") & """/>" & vbNewLine & _
               "    <info name=""HIS_NO"" value=""" & colPati("HIS_NO") & """/>" & vbNewLine
                If byt���� = 1 Then
                    strXML = strXML & "    <info name=""�����"" value=""" & colPati("�����") & """/>" & vbNewLine & _
                                "    <info name=""������Դ"" value=""" & colPati("������Դ") & """/>" & vbNewLine
                    
                Else
                    strXML = strXML & "    <info name=""סԺ��"" value=""" & colPati("סԺ��") & """/>" & vbNewLine & _
                                    "    <info name=""���ﲡ��ID"" value=""" & colPati("���ﲡ��ID") & """/>" & vbNewLine & _
                                    "    <info name=""���ﲡ��"" value=""" & colPati("���ﲡ��") & """/>" & vbNewLine & _
                                    "    <info name=""������Դ"" value=""" & colPati("������Դ") & """/>" & vbNewLine
                End If
        End If
    strXML = strXML & "</patient_info>"
    ZL_GET_PatiInfo = strXML
End Function

Public Function ZL_GET_Medicine(ByVal rsAdvice As ADODB.Recordset, ByVal bytFunc As Byte) As String
    Dim strXML As String
    '����:bytFunc=3 ���������Ϣ�ڵ�
    '<medicine_info>
    '  <medicine>
    '    <info name="ҽ��ID" value="2"/>
    '    <info name="��λ��" value="86903291000301" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/>
    '    <info name="������ĿID" value="67231" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
    '    <info name="��Һ���" value="1"/>
    '    <info name="������λ" value="ml"/>
    '    <info name="������" value="60"/>
    '    <info name="������-������" value="1.25"/>  '�����������Բ�������
    '    <info name="������-�����" value="40.87"/> '������-�����=trunc(��������/(0.0061*�������+0.0128*��������-0.1529),2)
    '    <info name="ÿ����" value="60"/>
    '    <info name="ÿ����-������" value="1.25"/>
    '    <info name="ÿ����-�����" value="40.87"/>
    '    <info name="��ҩƵ��" value="ÿ��һ��"/>
    '    <info name="��ҩ;��" value="������Һ"/>
    '  </medicine> ���ҩƷ���Medicine
    '</medicine_info>
    rsAdvice.Filter = "����=0"
    strXML = "<medicine_info>"
    Do While Not rsAdvice.EOF
        strXML = strXML & _
        "  <medicine>" & vbNewLine & _
        "    <info name=""ҽ��ID"" value=""" & rsAdvice!ҽ��ID & """/>" & vbNewLine & _
        "    <info name=""��λ��"" value=""" & rsAdvice!��λ�� & """ main=""46d64420-8319-4768-9a11-f4b0f5e4ce7a""/>" & vbNewLine & _
        "    <info name=""������ĿID"" value=""" & rsAdvice!������ĿID & """ main=""4e19df1c-c1b9-4a43-a83d-0741a19961ab""/>" & vbNewLine & _
        "    <info name=""��Һ���"" value=""" & rsAdvice!��Һ��� & """/>" & vbNewLine & _
        "    <info name=""������λ"" value=""" & rsAdvice!������λ & """/>" & vbNewLine & _
        "    <info name=""������"" value=""" & rsAdvice!������ & """/>" & vbNewLine & _
        "    <info name=""ÿ����"" value=""" & rsAdvice!ÿ���� & """/>" & vbNewLine & _
        "    <info name=""��ҩƵ��"" value=""" & rsAdvice!��ҩƵ�� & """/>" & vbNewLine & _
        "    <info name=""��ҩƵ������"" value=""" & rsAdvice!��ҩƵ������ & """/>" & vbNewLine & _
        "    <info name=""��ҩ;��"" value=""" & rsAdvice!��ҩ;�� & """/>" & vbNewLine
        If bytFunc = 3 Then
            strXML = strXML & _
                "    <info name=""��ҩ;������"" value=""" & rsAdvice!��ҩ;������ & """/>" & vbNewLine & _
                "    <info name=""ҽ����Ч"" value=""" & rsAdvice!ҽ����Ч & """/>" & vbNewLine & _
                "    <info name=""����ʱ��"" value=""" & rsAdvice!����ʱ�� & """/>" & vbNewLine & _
                "    <info name=""ҽ��״̬"" value=""" & rsAdvice!ҽ��״̬ & """/>" & vbNewLine & _
                "    <info name=""������־"" value=""" & rsAdvice!������־ & """/>" & vbNewLine & _
                "    <info name=""����ҽ��"" value=""" & rsAdvice!����ҽ�� & """/>" & vbNewLine & _
                "    <info name=""ҽ��ְ��"" value=""" & rsAdvice!ҽ��ְ�� & """/>" & vbNewLine & _
                "    <info name=""ҽ������ҩ��ȼ�"" value=""" & rsAdvice!ҽ������ҩ��ȼ� & """/>" & vbNewLine & _
                "    <info name=""ҽ������"" value=""" & rsAdvice!ҽ������ & """/>" & vbNewLine
            strXML = strXML & _
                "    <info name=""��ҩ����"" value=""" & rsAdvice!��ҩ���� & """/>" & vbNewLine & _
                "    <info name=""����"" value=""" & rsAdvice!���� & """/>" & vbNewLine & _
                "    <info name=""ҩƷ����ҩ��ȼ�"" value=""" & rsAdvice!ҩƷ����ҩ��ȼ� & """/>" & vbNewLine & _
                "    <info name=""�������"" value=""" & rsAdvice!������� & """/>" & vbNewLine & _
                "    <info name=""����˵��"" value=""" & rsAdvice!����˵�� & """/>" & vbNewLine & _
                "    <info name=""��ҩĿ��"" value=""" & rsAdvice!��ҩĿ�� & """/>" & vbNewLine & _
                "    <info name=""ҩƷ���ɵȼ�"" value=""" & rsAdvice!ҩƷ���ɵȼ� & """/>" & vbNewLine & _
                "    <info name=""ҩƷ��������"" value=""" & rsAdvice!ҩƷ�������� & """/>" & vbNewLine & _
                "    <info name=""ҩƷ����˵��"" value=""" & rsAdvice!ҩƷ����˵�� & """/>" & vbNewLine & _
                "    <info name=""�������"" value=""" & rsAdvice!������� & """/>" & vbNewLine
        End If
        strXML = strXML & "  </medicine>"
        rsAdvice.MoveNext
    Loop
    strXML = strXML & "</medicine_info>" & vbNewLine
    ZL_GET_Medicine = strXML
End Function

Public Function ZL_GET_Cusrules(ByVal rsAdvice As ADODB.Recordset) As String
    Dim strXML As String
 
    rsAdvice.Filter = "����=1"
    strXML = strXML & "<cusrules><diatreat>" & vbNewLine
    Do While Not rsAdvice.EOF
        strXML = strXML & "<info name=""������ĿID"" value=""" & rsAdvice!������ĿID & """ main=""4e19df1c-c1b9-4a43-a83d-0741a19961ab""/>" & vbNewLine
        rsAdvice.MoveNext
    Loop
    strXML = strXML & "</diatreat></cusrules>"  '</diatreat></cusrules>�滻ʱ��Ϊ�ؼ���
  
    ZL_GET_Cusrules = strXML
End Function

Public Function ZL_GET_Details(ByVal colPati As Collection, ByVal rsAdvice As ADODB.Recordset, _
    ByVal byt���� As Byte, ByVal bytFunc As Byte) As String
    Dim strXML As String
'byt����=1 ����;2-סԺ
'<details_xml>
'    <patient_info>
'    </patient_info>
'    <medicine_info>
'    </medicine_info>
'</details_xml>
    strXML = "<details_xml>"
            strXML = strXML & ZL_GET_PatiInfo(colPati, byt����)
            strXML = strXML & ZL_GET_Medicine(rsAdvice, bytFunc)
            If gstrIP <> "" And bytFunc = 0 Then strXML = strXML & ZL_GET_Cusrules(rsAdvice)
            strXML = strXML & "</details_xml>"
    ZL_GET_Details = strXML
End Function
 
Public Function ReadLobForPASS() As String
'���ܣ���ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'������
'���أ�������ݵ��ļ�����ʧ���򷵻��㳤��""
    Dim rsLob As ADODB.Recordset
    Dim lngCount As Long
    Dim strText As String
    Dim strSQL As String
    Dim strFile As String
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select Zl_Read_����������ҩ����([1]) as Ƭ�� From Dual"
    lngCount = 0
    strFile = ""
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(strSQL, "ReadLobForPASS", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        strFile = strFile & strText
        lngCount = lngCount + 1
    Loop
     
    ReadLobForPASS = strFile
    Exit Function
Errhand:
    Err.Clear
End Function

Private Function Getÿ����(ByVal str������ As String, ByVal str�����λ As String, ByVal intƵ�ʴ��� As Integer, _
    ByVal intƵ�ʼ�� As Integer, ByVal strƵ�� As String) As String
'����:
'1.ÿ����=������*��Ƶ��
'2.��Ƶ�μ��㣺
'                    b.�����λ=�� and Ƶ�ʼ��=1����Ƶ��=Ƶ�ʴ���
'                    c.�����λ=�� and Ƶ�ʼ��>1 and Ƶ�ʴ���=1����Ƶ��=1
'                    d.�����λ=Сʱ and Ƶ�ʼ��<=24,��Ƶ��=24/Ƶ�ʼ��*Ƶ�ʴ���
'                    e.�����λ=Сʱ and Ƶ�ʼ��>24 and Ƶ�ʴ���=1����Ƶ��=1
'                    f.�����λ=�� and Ƶ�ʴ���=1����Ƶ��=1
'                   strƵ��=һ����  ��Ƶ��=1

    Dim strÿ���� As String
    
    If str�����λ = "��" And intƵ�ʼ�� = 1 Then
        strÿ���� = Val(str������) * intƵ�ʴ���
    ElseIf str�����λ = "��" And intƵ�ʼ�� > 1 And intƵ�ʴ��� = 1 Then
        strÿ���� = Val(str������) * 1
    ElseIf str�����λ = "Сʱ" And intƵ�ʼ�� <= 24 Then
        strÿ���� = Val(str������) * (24 / intƵ�ʼ�� * intƵ�ʴ���)
    ElseIf str�����λ = "Сʱ" And intƵ�ʼ�� > 24 And intƵ�ʴ��� = 1 Then
        strÿ���� = Val(str������) * 1
    ElseIf str�����λ = "��" And intƵ�ʴ��� = 1 Then
        strÿ���� = Val(str������) * 1
    ElseIf strƵ�� = "һ����" Then
        strÿ���� = Val(str������) * 1
    End If
    Getÿ���� = FormatEx(strÿ����, 2)
End Function

Private Function Get��������(ByVal dbl�������� As Double) As String
'����:
'������:<=28��
'Ӥ��:>28 �� ;<=365��
'�׶�:>365 ��; <= 6��
'��ͯ:>6��;<=14��
'����:>14�� and <=18
'����:>18   and  <=60
'����:>60��

    Dim str�������� As String
    
    If dbl�������� <= 28 Then
        str�������� = "������"
    ElseIf dbl�������� > 28 And dbl�������� <= 365 Then
        str�������� = "Ӥ��"
    ElseIf dbl�������� > 365 And dbl�������� <= (365 * 6) Then
        str�������� = "�׶�"
    ElseIf dbl�������� > (365 * 6) And dbl�������� <= (365 * 14) Then
        str�������� = "��ͯ"
    ElseIf dbl�������� > (365 * 14) And dbl�������� <= (365 * 18) Then
        str�������� = "����"
    ElseIf dbl�������� > (365 * 18) And dbl�������� <= (365 * 60) Then
        str�������� = "����"
    ElseIf dbl�������� > (365 * 60) Then
        str�������� = "����"
    Else
        str�������� = ""
    End If
    Get�������� = str��������
End Function

Private Function ZL_ParseXML(ByVal strData As String) As ADODB.Recordset
          '����:����XML�ַ���
          '����XML:
          '1.  �ڵ���ͣ�
          '<order>��ͷβ�ڵ�
          '<orderid>��ҽ��ID
          '<drugcode>:��λ��
          '<type>���������ͣ�������Ϣ��������������ǰ���ص�������һ������
          '<level>����ʾ�ȼ���������Ϣ��������/����/��
          '<describ>��������������ǰ������Ϣ����Ҫ����
          '<remaks>����ע��Ϣ
          '<order><order_id>1</order_id> '�໥���á�ע������顢�ظ���ҩ�������඼�Ƕ��ҩһ����ʾ��,�ʷ���ҽ��ID��,��Ƕ��ŷָ�
          '<drugcode>86900967000160</drugcode>
          '<type>��Ӧ֢</type><level></level><describ>���Ȼ���ע��Һ��ֻ�ʺ����������������ע������</describ>
          '<remaks>��ҩ;��</remaks></order><order>
          '<order_id>2</order_id><drugcode>86903291000301</drugcode>
          '<type>��Ӧ֢</type><level></level><describ>������ע��Һ��ֻ�ʺ����������������ע������ע��</describ>
          '<remaks>��ҩ;��</remaks></order>
        Dim xmlDoc As New DOMDocument
        Dim xNode As IXMLDOMNode
        Dim xNodeList As IXMLDOMNodeList
        Dim rsRet As ADODB.Recordset
        Dim arrTemp As Variant
        Dim i As Long
          
10      On Error GoTo errH
20      Set rsRet = InitAdviceRS(FUN_�����_ZL)
        '��ȡ������Ӧ���ݣ�XML��ʽ��
30      xmlDoc.loadXML (strData)
40      Set xNodeList = xmlDoc.selectNodes(".//order")
50      For Each xNode In xNodeList
60          arrTemp = Split(xNode.selectSingleNode(".//order_id").Text, ",")
70          For i = LBound(arrTemp) To UBound(arrTemp)
80              rsRet.AddNew
90              rsRet!OrderId = arrTemp(i)
100             rsRet!DrugCode = xNode.selectSingleNode(".//drugcode").Text
110             rsRet!Type = xNode.selectSingleNode(".//type").Text
120             If Mid(rsRet!Type & "", 1, Len("�����")) = "�����" Then
130                 rsRet!Category = 1
140             Else
150                 rsRet!Category = 0
160             End If
170             rsRet!Level = xNode.selectSingleNode(".//level").Text
180             rsRet!describ = xNode.selectSingleNode(".//describ").Text
190             rsRet!remaks = xNode.selectSingleNode(".//remaks").Text
200             If rsRet!Level = "��ֹ" Then
210                 rsRet!Light = 4
220                 rsRet!WarnLevel = 4   '��ʾ��������
230             ElseIf rsRet!Level = "����" Then
240                 rsRet!Light = 2
250                 rsRet!WarnLevel = 3   '��ʾ��������
260             ElseIf rsRet!Level = "ע��" Then
270                 rsRet!Light = 3
280                 rsRet!WarnLevel = 1
290             ElseIf rsRet!Level = "����" Or rsRet!Level = "" Then
300                 rsRet!Light = 1
310                 rsRet!WarnLevel = 2
320             Else
330                 rsRet!Light = 0
340                 rsRet!WarnLevel = 0
350             End If
360             rsRet!Tag = IIf(i = LBound(arrTemp), 0, 1)    '1-����ظ�����
                '��������ȥ��
370             If rsRet!Category = 1 Then
380                 rsRet.Filter = "Category=1 And Type='" & rsRet!Type & "' And describ='" & rsRet!describ & "'"
390                 If rsRet.RecordCount > 1 Then rsRet.Delete
400             End If
410             rsRet.Update
420         Next
430     Next
440     Set ZL_ParseXML = rsRet
450     Exit Function
errH:
460     MsgBox Err.Description & vbCrLf & "ZL_ParseXML" & "�� " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function ZL_ParseXMLCusRules(ByVal strData As String) As ADODB.Recordset
    '����:����XML�ַ���
    '����XML:
    '    "<cusrules>" & vbNewLine & _
    '    "  <result>" & vbNewLine & _
    '    "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
    '    "    <info name=""�Ƿ�����"" type=""ƽ�浥ѡ"" index=""2"" value=""��|��"" class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
    '    "    <info name=""�Ա�"" type=""ƽ�浥ѡ"" index=""3"" value=""��|Ů"" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
    '    "    <info name=""����Դ"" type=""������ѡ"" index=""4"" value=""����|ͷ��|����ù��|��Ī����|��˾ƥ��"" class=""�����Ŀ"" obsid=""fac26638-6d75""/>" & vbNewLine & _
    '    "    <info name=""����"" type=""��������"" index=""5"" value=""Ӥ��|ѧǰ|��ͯ|����|����|����|����"" class=""�����Ŀ"" obsid=""fac26638-6d75""/>" & vbNewLine & _
    '    "    <info name=""����ʷ����"" type=""�ı�"" index=""475"" value="""" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
    '    "  </result>" & vbNewLine & _
    '    "</cusrules>"
        Dim xmlDoc As New DOMDocument
        Dim xNode As IXMLDOMNode
        Dim xNodeList As IXMLDOMNodeList
        Dim rsRet As ADODB.Recordset
        Dim strNodeValue As String
        
        Dim i As Long
    
        On Error GoTo errH
100     Set rsRet = InitAdviceRS(FUN_��������_ZL)
        '��ȡ������Ӧ���ݣ�XML��ʽ��
102     xmlDoc.loadXML (strData)
104     Set xNodeList = xmlDoc.selectNodes(".//cusrules/result/info")
106     For Each xNode In xNodeList
108             rsRet.AddNew
110             For i = 0 To xNode.Attributes.length - 1
112                 strNodeValue = xNode.Attributes(i).nodeValue
114                 Select Case xNode.Attributes(i).baseName
                    Case "name"
116                     rsRet!Name = strNodeValue
118                 Case "type"
120                     rsRet!Type = strNodeValue
122                 Case "index"
124                     rsRet!Index = strNodeValue
126                 Case "value"
128                     rsRet!Value = strNodeValue
130                 Case "default"
132                     rsRet!Default = strNodeValue
134                 Case "class"
136                     rsRet!Class = strNodeValue
138                 Case "obsid"
140                     rsRet!Obsid = strNodeValue
142                 Case "proid"
144                     rsRet!Proid = strNodeValue
                    End Select
                Next
146             rsRet.Filter = "Name='" & rsRet!Name & "' And Type ='" & rsRet!Type & "'"
148             If rsRet.RecordCount > 1 Then 'ȥ���ظ���Ŀ
150                 rsRet.Delete
152             Else
154                 rsRet.Update
                End If
        Next
156     rsRet.Filter = ""
         
158     Set ZL_ParseXMLCusRules = rsRet
        Exit Function
errH:
160     MsgBox Err.Description & vbCrLf & "ZL_ParseXMLCusRules" & "�� " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetTestXML(ByVal bytFunc As Byte, Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef strAsk As String) As String
    Dim strPar As String
    Dim i As Long

    If bytFunc = 0 Then
        strPar = strPar & "{""ҽԺID_IN"":1," & vbNewLine & _
                " ""ҩƷҽ��XML_IN"":""<details_xml>" & vbNewLine & _
                "  <patient_info>" & vbNewLine & _
                "    <info name=\""��������\"" value=\""28114.45\""/>" & vbNewLine & _
                "    <info name=\""��������\"" value=\""����\""/>" & vbNewLine & _
                "    <info name=\""�Ա�\"" value=\""Ů\""/>" & vbNewLine & _
                "    <info name=\""ְҵ\"" value=\""�˶�Ա\""/>" & vbNewLine & _
                "    <info name=\""����\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""����\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""�ι��ܲ�ȫ\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""���ظι��ܲ�ȫ\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""�����ܲ�ȫ\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""���������ܲ�ȫ\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""���\"" value=\""J18.000\""/>" & vbNewLine & _
                "  </patient_info>"
            strPar = strPar & "<medicine_info>" & vbNewLine & _
                "    <medicine>" & vbNewLine & _
                "      <info name=\""ҽ��ID\"" value=\""1\""/>" & vbNewLine & _
                "      <info name=\""��λ��\"" value=\""86900967000160\"" main=\""46d64420-8319-4768-9a11-f4b0f5e4ce7a\""/>" & vbNewLine & _
                "      <info name=\""������ĿID\"" value=\""67232\"" main=\""4e19df1c-c1b9-4a43-a83d-0741a19961ab\""/>" & vbNewLine & _
                "      <info name=\""��Һ���\"" value=\""1\""/>" & vbNewLine & _
                "      <info name=\""������λ\"" value=\""ml\""/>" & vbNewLine & _
                "      <info name=\""������\"" value=\""250\""/>" & vbNewLine & _
                "      <info name=\""������-������\"" value=\""5.21\""/>" & vbNewLine & _
                "      <info name=\""������-�����\"" value=\""170.3\""/>" & vbNewLine & _
                "      <info name=\""ÿ����\"" value=\""250\""/>" & vbNewLine & _
                "      <info name=\""ÿ����-������\"" value=\""5.21\""/>" & vbNewLine & _
                "      <info name=\""ÿ����-�����\"" value=\""170.3\""/>" & vbNewLine & _
                "      <info name=\""��ҩƵ��\"" value=\""ÿ��һ��\""/>" & vbNewLine & _
                "      <info name=\""��ҩ;��\"" value=\""������Һ\""/>" & vbNewLine & _
                "    </medicine>"

            strPar = strPar & " <medicine>" & vbNewLine & _
        "      <info name=\""ҽ��ID\"" value=\""2\""/>" & vbNewLine & _
        "      <info name=\""��λ��\"" value=\""86903291000301\"" main=\""46d64420-8319-4768-9a11-f4b0f5e4ce7a\""/>" & vbNewLine & _
        "      <info name=\""������ĿID\"" value=\""67231\"" main=\""4e19df1c-c1b9-4a43-a83d-0741a19961ab\""/>" & vbNewLine & _
        "      <info name=\""��Һ���\"" value=\""1\""/>" & vbNewLine & _
        "      <info name=\""������λ\"" value=\""ml\""/>" & vbNewLine & _
        "      <info name=\""������\"" value=\""60\""/>" & vbNewLine & _
        "      <info name=\""������-������\"" value=\""1.25\""/>" & vbNewLine & _
        "      <info name=\""������-�����\"" value=\""40.87\""/>" & vbNewLine & _
        "      <info name=\""ÿ����\"" value=\""60\""/>" & vbNewLine & _
        "      <info name=\""ÿ����-������\"" value=\""1.25\""/>" & vbNewLine & _
        "      <info name=\""ÿ����-�����\"" value=\""40.87\""/>" & vbNewLine & _
        "      <info name=\""��ҩƵ��\"" value=\""ÿ��һ��\""/>" & vbNewLine & _
        "      <info name=\""��ҩ;��\"" value=\""������Һ\""/>" & vbNewLine & _
        "    </medicine>" & vbNewLine & _
        "  </medicine_info>" & vbNewLine & _
        "</details_xml>""}"

    ElseIf bytFunc = 1 Then
        strPar = """[{\""ͨ������\"":\""������Ƭ\"",\""��Ʒ��\"":null,\""����ƴ��\"":null,\""Ӣ������\"":\""\\n������Ƭ\\nIsoniazid Tablets\"",\""ҩ����\"":\""0.1g\"",\""ҩ�����\"":\""Ƭ��\"",\""������ҵ\"":\""��������ҩҵ���޹�˾\""" & _
                    ",\""��׼�ĺ�\"":\""��ҩ׼��H33021636\"",\""��ѧ����\"":\""4-��़�����\"",\""��״\"":\""��ƷΪ��ɫƬ�������ɫƬ\"",\""ҩ����\"":\""��Ʒ��һ�־���ɱ�����õĺϳɿ���ҩ����Ʒֻ�Է�֦�˾�����Ҫ��������ֳ�ڵ�ϸ����Ч�������û�����δ������������������ϸ����֦���ᣨmycolicacid���ĺϳɶ�ʹϸ�������ѡ�\"","
        strPar = strPar & "\""ҩ������ѧ\"":\""��Ʒ�ڷ���Ѹ����θ�������գ����ֲ���ȫ����֯����Һ�У������Լ�Һ����ˮ����ˮ��Ƥ�������⡢��֭�͸�������֯��" & _
                "���ɴ���̥�����ϡ����׽���ʽ�0~10%���ڷ�1~2СʱѪҩŨ�ȿɴ��ֵ����4~6Сʱ��ѪҩŨ�ȸ��ݻ��ߵ���������������һ�����������ߣ�T1/2Ϊ0.5~1.6Сʱ��" & _
                "����������Ϊ2~5Сʱ���Ρ����������߿����ӳ�����л��Ҫ�ڸ����������������޻��Դ�л��������еľ��иζ��ԡ����������������Ŵ���������" & _
                "���������߳��и���N-����ת��øȱ����δ�������������¿ɱ����ֽ�ϡ���Ʒ��Ҫ������й��Լ70%������24Сʱ���ų����󲿷�Ϊ�޻��Դ�л�" & _
                "������������93%��������������Һ���ų�������������Ϊ63%��������������Һ��7%�������³���������ͣ���������������Ϊ37%����Ʒ��ͨ��Ѫ�����ϣ�"

        strPar = strPar & "��ɴ���֭�ų�������������Һ��̵Һ�ͷ�����ų����൱���������¿ɾ�ѪҺ͸���븹Ĥ͸�������\""," & _
                "\""��Ӧ֢\"":\""1�������������������ҩ���ϣ������ڸ��ͽ�˲������ƣ������������Ĥ���Լ�������֦�˾���Ⱦ��" & _
                "��2�������µ��������ڸ��ͽ�˲���Ԥ�������½�ȷ��Ϊ��˲����ߵļ�ͥ��Ա�����нӴ��ߣ��ڽ�˾��ش��������������飨PPD��" & _
                "ǿ����ͬʱ�ز�X���߼����Ϸǽ����Խ�˲���̵�����ԣ���ȥδ���ܹ����濹��������ߣ������ڽ����������Ƽ����ڼ������ƵĻ��ߣ�" & _
                "ĳЩѪҺ������״��Ƥϵͳ���������Ѫ����������ϲ��������򲡡���֢�����λ�θ�г����Ȼ��ߣ����˾��ش�������������������Է�Ӧ��" & _
                "����35�����½�˾��ش������������������ԵĻ��ߣ�����֪����ΪHIV��Ⱦ�ߣ����˾��ش�������������������Է�Ӧ�ߣ������Էν�˻��������нӴ��ߡ�\"","

                strPar = strPar & _
                "\""�÷�����\"":\""�ڷ���Ԥ��������һ��0.3g���ٷ���С��ÿ�հ�����10mg/kg��һ������������0.3g���ٷ������ƣ����������������ҩ���ã�������ÿ�տڷ�5mg/kg�����0.3g����ÿ��15mg/kg�����900mg��ÿ��2~3�Ρ�С��������ÿ��10~20mg/kg��ÿ�ղ�����0.3g���ٷ���ĳЩ���ؽ�˲���������������Ĥ�ף���ÿ�հ����ؿɸߴ�30mg/kg��һ�������500mg������Ҫע��ι����𺦺���Χ���׵ķ�����\""," & _
                "\""������Ӧ\"":\""�����ʽ϶����в�̬���Ȼ���ľ��̸С����Ƹл���ָ��ʹ����Χ���ף�����ɫ���ۻ�Ƥ����Ⱦ���ζ��ԣ�35�����ϻ��߸ζ��Է��������ߣ���ʳ�����ѡ��쳣���������������Ļ�Ż�£��ζ��Ե�ǰ��֢״���������ʼ�����������ģ�����������ˣ��ϲ��򲻺ϲ���ʹ�������ף������ȡ�Ƥ�Ѫϸ�����ټ������鷿�����ȡ���Ʒż�����񾭶�������ĳ鴤��\""," & _
                "\""����֢\"":\""�ι��ܲ������ߣ����񲡻��ߺ���ﲡ�˽��á�\""," & _
                "\""ע������\"":\""��1�����������Ӧ�����������̰�����������������������ѧ�ṹ�й�ҩ�������Ҳ���ܶԱ�Ʒ������" & _
                "��2������ϵĸ��ţ�������ͭ���������ǲⶨ�ɳʼ����Է�Ӧ������Ӱ��ø���ⶨ�Ľ���������¿�ʹѪ�嵨���ء������ᰱ��ת��ø���Ŷ����ᰱ��ת��ø�Ĳⶨֵ���ߡ�"

                strPar = strPar & "��3���о��񲡡���ﲡʷ�ߡ���������������Ӧ���á���4�����Ƴ��г���������֢״��Ӧ���������۲���飬�����ڸ��顣" & _
                "��5���������ж�ʱ���ô����ά����B6�Կ���\""," & _
                "\""�и���ҩ\"":\""��1����Ʒ�ɴ���̥�̣�����̥��ѪҩŨ�ȸ���ĸѪҩŨ�ȡ�����ʵ��֤ʵ�����¿�������̥��������������δ֤ʵ���и�Ӧ��ʱ������Ȩ�����ס�" & _
                "������������ҩ������ʱ��̥����������δ���������⣬����������ҩʱӦ���й۲첻����Ӧ����2������������֭��Ũ�ȿɴ�12mg/L,��ѪҩŨ�������" & _
                "��Ȼ����������δ֤ʵ�����⣬�����ڼ�Ӧ����Ӧ���Ȩ�����ס�����ҩ����ֹͣ���顣\"","

                strPar = strPar & _
                "\""��ͯ��ҩ\"":\""�ϸ��ն�ͯ�÷�����ʹ��\""," & _
                "\""�໥����\"":\""1������������ʱÿ�����ƣ�������Ʒ�շ��ĸ��඾�Է�Ӧ�������������µĴ�л���������������µļ����������й۲�ζ�������" & _
                "ӦȰ�滼�߷�ҩ�ڼ����ƾ����ϡ���2����������ҩ���ӻ������������¿ڷ�������գ�ʹѪҩŨ�ȼ��ͣ���Ӧ��������ͬʱ���ã����ڿڷ������ǰ����1Сʱ���������¡�" & _
                "��3������Ѫҩ�����㶹�ػ�����˫ͪ�������������ͬʱӦ��ʱ�����������˿���ҩ��ø��л��ʹ����������ǿ��" & _
                "��4���뻷˿����ͬ��ʱ������������ϵͳ������Ӧ����ͷ�����˯��������������������й۲�������ϵͳ��������������ڴ�����Ҫ�����ȽϸߵĹ����Ļ��ߡ�"

                strPar = strPar & "��5������ƽ�������º���ʱ�����Ӹζ��Ե�Σ���ԣ����������иι������߻�Ϊ�����¿��������ߣ�������Ƴ̵�ͷ3����Ӧ����������޸ζ���������֡�" & _
                "��6��������Ϊά����B6���׿���,������ά����B6�����ų���,������ܵ�����Χ����,����������ʱά����B6����Ҫ�����ӡ�" & _
                "��7����������Ƥ�ʼ���(������������)����ʱ,�������������ڸ��ڵĴ�л����й,���º���ѪҩŨ�ȼ��Ͷ�Ӱ����Ч,�ڿ��������߸�Ϊ����,Ӧ�ʵ�����������" & _
                "��8���밢��̫�ᣨalfentanil������ʱ������������Ϊ��ҩø���Ƽ������ӳ�����̫������ã���" & _
                "˫����(disulfiram)���ÿ���ǿ��������ϵͳ���ã�����ѣ�Ρ�������Э�����׼��ǡ�ʧ�ߵȣ��밲���Ѻ��ÿ����Ӿ��������Ե��޻�����л����γɡ�" & _
                "��9�����������̰������������ҩ���ã��ɼ��غ���ߵĲ�����Ӧ���������ζ���ҩ���ÿ����ӱ�Ʒ�ĸζ��ԣ�����˾������⡣��10�������²�����ͪ������俵����ã�" & _
                "���ʹ�����ߵ�ѪҩŨ�Ƚ��͡���11���뱽��Ӣ�ƻ򰱲�����ʱ�����ƶ����ڸ����еĴ�л�������±���Ӣ�ƻ򰱲��ѪҩŨ������" & _
                "�����������������Ⱥ�Ӧ�û����ʱ������Ӣ�ƻ򰱲��ļ���Ӧ�ʵ���������12��������������Ӻ���ʱ�����������¿��յ���ϸ��ɫ��P-450��" & _
                "ʹǰ���γɶ��Դ�л��������ӣ������Ӹζ��Լ������ԡ���13���뿨����ƽͬʱӦ��ʱ�������¿��������л��ʹ������ƽ��ѪҩŨ�����ߣ�" & _
                "�������Է�Ӧ��������ƽ���յ������µ�΢�����л���γɾ��иζ��Ե��м��л�����ӡ���14����Ʒ�����������񾭶�ҩ����ã����������񾭶��ԡ�\""," & _
                "\""ҩ�����\"":\""δ���и����������޲ο�����\""," & _
                "\""��������\"":\""�ڹ⣬�ܷ⣬�ڸ��ﴦ���档\""}]"""
    ElseIf bytFunc = 2 Then
        strPar = "<details_xml>"
        strPar = strPar & "<order><order_id>1</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>��Ӧ֢</type><level>����</level><describ>���Ȼ���ע��Һ��ֻ�ʺ����������������ע������</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>" & _
                            "<order><order_id>2</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>��Ӧ֢</type><level>����</level><describ>������ע��Һ��ֻ�ʺ����������������ע������ע�� ���Ȼ���ע��Һ��ֻ�ʺ����������������ע�����á��Ȼ���ע��Һ��ֻ�ʺ����������������ע�����á�����ע��Һ��ֻ�ʺ����������������ע������ע�� ���Ȼ���ע��Һ��ֻ�ʺ����������������ע�����á��Ȼ���ע��Һ��ֻ�ʺ����������������ע������</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>"
        strPar = strPar & "<order><order_id>3</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>��Ӧ֢</type><level>����</level><describ>���Ȼ���ע��Һ��ֻ�ʺ����������������ע������</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>" & _
                            "<order><order_id>4</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>��Ӧ֢</type><level>����</level><describ>������ע��Һ��ֻ�ʺ����������������ע������ע��</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>"
        strPar = strPar & "<order><order_id>5</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>��Ӧ֢</type><level></level><describ>���Ȼ���ע��Һ��ֻ�ʺ����������������ע������</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>" & _
                            "<order><order_id>6</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>��Ӧ֢</type><level></level><describ>������ע��Һ��ֻ�ʺ����������������ע������ע��</describ>" & _
                            "<remaks>��ҩ;��</remaks></order>"
        strPar = strPar & "<order><order_id>7,8</order_id><drugcode>86900967000160</drugcode>" & _
                    "<type>ҩƷ�໥����</type><level>����</level><describ>����ŵ������ע��Һ���͡�ά����Cע��Һ�����໥���ã�" & vbCrLf & "�����ά����C�ɸ��ſ���ҩ�Ŀ���Ч����</describ>" & _
                    "<remaks>��ҩ;��</remaks></order>" & _
                    "<order><order_id>9,10</order_id><drugcode>86903291000301</drugcode>" & _
                    "<type>ҩƷ�໥����</type><level>����</level><describ>������ע��Һ��ֻ�ʺ����������������ע������ע��</describ>" & _
                    "<remaks>��ҩ;��</remaks></order>"
        strPar = strPar & "</details_xml>"



        strPar = " <details_xml><order><order_id>1211848</order_id><drugcode>86900002000018</drugcode><type>��ҩ;��</type>" & vbNewLine & _
            " <level></level><describ>����������Ƭ��ֻ�ʺ�����������������̷�</describ>" & vbNewLine & _
            " <remaks>��ҩ;��</remaks></order><order><order_id>1211850,1211852</order_id>" & vbNewLine & _
            " <drugcode></drugcode><type>��ҩƷ�໥����</type><level>����</level>" & vbNewLine & _
            " <describ>�������ע��Һ���͡���˾ƥ�ֳ���Ƭ�����໥���ã�����ѡ���Ի�����ø��2���Ƽ���COX��2���Ƽ���" & vbNewLine & _
            " ���ڵķ������࿹��ҩ��NSAIDs���ή�����������������Ѫѹҩ���Ч������ˣ�Ѫ�ܽ����آ������׿�������" & vbNewLine & _
            " ��Ҳ�ᱻ����ѡ����COX��2���Ƽ����ڵ�NSAIDs��ҩ�������</describ><remaks></remaks></order><order>" & vbNewLine & _
            " <order_id>1211848,1211850</order_id><drugcode></drugcode><type>�����ظ���ҩ</type><level>����</level>" & vbNewLine & _
            " <describ>����������Ƭ����˾ƥ�ֳ���Ƭ��ͬ���ڡ��н���Ѫ�����õ�ҩƷ���������ظ���ҩ��</describ><remaks></remaks>" & vbNewLine & _
            " </order><order><order_id>1211850,1211852,1211854</order_id><drugcode></drugcode><type>�����ظ���ҩ</type>" & vbNewLine & _
            " <level>����</level><describ>����˾ƥ�ֳ���Ƭ�������ע��Һ��ע������ø��ͬ���ڡ���" & vbNewLine & _
            "��ҩ���������ظ���ҩ��</describ><remaks></remaks></order>"
        
        strPar = strPar & "<order><order_id>1211850,1211854</order_id><drugcode></drugcode><type>�����ظ���ҩ</type>" & _
            "<level>����</level><describ>����˾ƥ�ֳ���Ƭ��ע������ø��ͬ���ڡ�Ӱ��ֹѪ��ҩ��������ظ���ҩ��</describ>" & _
            "<remaks></remaks></order><order><order_id>1211854</order_id><drugcode>86901576000121</drugcode><type>��ҩ;��</type><level></level><describ>��ע������ø��ֻ�ʺ������������ϴ��������ע��������ע</describ><remaks>��ҩ;��</remaks></order><order><order_id>1211848</order_id><drugcode>86900002000018</drugcode><type>����֢</type><level>����</level><describ>����������Ƭ������С��18������</describ><remaks></remaks></order><order><order_id>1211854,1211850</order_id><drugcode></drugcode><type>��ҩƷ�໥����</type><level>����</level><describ>����˾ƥ�ֳ���Ƭ���͡�ע������ø�����໥���ã�NSAIDs����ѪС��ۼ�����θ����ճĤ�������ӿ���ҩ��Ļ��ԣ�����ʹ�ÿ���ҩ�Ĳ���θ������Ѫ�ķ��ա����ǿ��Խ������еļ�⣬���ȷ���Ӧ�������㶹����ڷ�����Ѫҩ������ƥ����Ѫ˨�ܽ�" & _
            "�������غ��á�</describ><remaks></remaks></order><order><order_id>1211848,1211850</order_id><drugcode></drugcode><type>��ҩƷ�໥����</type><level>����</level><describ>����˾ƥ�ֳ���Ƭ���͡���������Ƭ�����໥���ã�������ҩ�������ȵ��ء��������ࣺ" & vbNewLine & _
            "�߼�����˾ƥ�־��н�Ѫ�����ö���ǿ����Ч������������������ྺ�����Ѫ�����ס�</describ><remaks></remaks></order><order><order_id>1211850,1211852</order_id><drugcode></drugcode><type>�����ظ���ҩ</type><level>����</level><describ>����˾ƥ�ֳ���Ƭ�������ע��Һ��ͬ���ڡ���״��������ҩ���������ظ���ҩ��</describ>" & _
            "<remaks></remaks></order><order><order_id>1211850</order_id><drugcode>86979489000088</drugcode>" & _
            "<type>��ҩ;��</type><level></level><describ>����˾ƥ�ֳ���Ƭ��ֻ�ʺ�����������ڷ���ҩ</describ>" & _
            "<remaks>��ҩ;��</remaks></order><order><order_id>1211848,1211850</order_id><drugcode></drugcode>" & _
            "<type>�����ظ���ҩ</type><level>����</level><describ>����������Ƭ����˾ƥ�ֳ���Ƭ��ͬ���ڡ���Ѫ�����׸߶Ƚ�ϵ�ҩ��������ظ���ҩ��</describ><remaks></remaks></order>" & _
            "<order><order_id>1211850,1211854</order_id><drugcode></drugcode><type>�����ظ���ҩ</type><level>����</level><describ>����˾ƥ�ֳ���Ƭ��ע������ø��ͬ���ڡ����඾��ҩ��������ظ���ҩ��</describ>" & _
            "<remaks></remaks></order><order><order_id>1211850,1211854</order_id><drugcode></drugcode>" & _
            "<type>�����ظ���ҩ</type><level>����</level><describ>����˾ƥ�ֳ���Ƭ��ע������ø��ͬ���ڡ�ѪҺѧҩ��������ظ���ҩ��</describ><remaks></remaks></order>" & _
            "<order><order_id>1211854,1211852</order_id><drugcode></drugcode><type>��ע����������</type><level>����</level><describ>��ע������ø���������ע��Һ�������飺��ҩ��Ϻ��������ҩ��ҩ��ѧ��ҩЧѧ�ȷ���������ɡ�</describ><remaks></remaks></order></details_xml>"

        '����ģ�⾯ʾ������
        rsAdvice.Filter = ""
        For i = 1 To rsAdvice.RecordCount
            strPar = Replace(strPar, "<order_id>" & i & "</order_id>", "<order_id>" & rsAdvice!ҽ��ID & "</order_id>")
            If i = 6 Then Exit For
            rsAdvice.MoveNext
        Next
    ElseIf bytFunc = 3 Then
        '��������֧������:ƽ���ѡ��ƽ�浥ѡ���ı���������ѡ����������
        strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""�Ƿ�����adsfasdfasdf���߰�˹�ٷ���������˹�ٷҷ��ʹ������ط���˹�ٷҷ��͵ط�as�ط���˹�ٷҰ�˹�ٷҰ�˹�ٷҷ��͵���ɵɵ�ķ�"" type=""ƽ�浥ѡ"" index=""2"" value=""��|��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""�Ա�"" type=""ƽ�浥ѡ"" index=""3"" value=""��|Ů""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""����Դ"" type=""������ѡ"" index=""4"" value=""����|ͷ��|����ù��|��Ī����|��˾ƥ��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""����ù��""/>" & vbNewLine & _
                "    <info name=""����"" type=""��������"" index=""5"" value=""Ӥ��|ѧǰ|��ͯ|����|����|����|����""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""ѧǰ""/>" & vbNewLine & _
                "    <info name=""����ʷ����"" type=""�ı�"" index=""475"" value="""" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
                "    <info name=""�Ƿ�����"" type=""ƽ�浥ѡ"" index=""2"" value=""��|��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""�Ա�"" type=""ƽ�浥ѡ"" index=""3"" value=""��|Ů"" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""����Դ"" type=""������ѡ"" index=""4"" value=""����|ͷ��|����ù��|��Ī����|��˾ƥ��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""����ù��""/>" & vbNewLine & _
                "    <info name=""����"" type=""��������"" index=""5"" value=""Ӥ��|ѧǰ|��ͯ|����|����|����|����"" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""����ʷ����"" type=""�ı�"" index=""475"" value="""" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
                "    <info name=""�Ƿ�����"" type=""ƽ�浥ѡ"" index=""2"" value=""��|��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""�Ա�"" type=""ƽ�浥ѡ"" index=""3"" value=""��|Ů"" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""����Դ"" type=""������ѡ"" index=""4"" value=""����|ͷ��|����ù��|��Ī����|��˾ƥ��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""����ù��""/>" & vbNewLine & _
                "    <info name=""����"" type=""��������"" index=""5"" value=""Ӥ��|ѧǰ|��ͯ|����|����|����|����"" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""����ʷ����"" type=""�ı�"" index=""475"" value="""" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                "</cusrules>"
                
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""�Ƿ�����adsfasdfasdf���߰�˹�ٷ���������˹�ٷҷ��ʹ������ط���˹�ٷҷ��͵ط�as�ط���˹�ٷҰ�˹�ٷҰ�˹�ٷҷ��͵���ɵɵ�ķ�"" type=""ƽ�浥ѡ"" index=""2"" value=""��|��""  class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�AAAAAAAAAAA"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""�Ա�"" type=""ƽ�浥ѡ"" index=""3"" value=""��|Ů""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""��""/>" & vbNewLine & _
                "    <info name=""����Դ"" type=""������ѡ"" index=""4"" value=""����|ͷ��|����ù��|��Ī����|��˾ƥ��""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""����ù��""/>" & vbNewLine & _
                "    <info name=""����"" type=""��������"" index=""5"" value=""Ӥ��|ѧǰ|��ͯ|����|����|����|����""  class=""�����Ŀ"" obsid=""fac26638-6d75"" default=""ѧǰ""/>" & vbNewLine & _
                "    <info name=""����ʷ����"" type=""�ı�"" index=""475"" value="""" class=""�����Ŀ"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            'ƽ���ѡ
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����|��������|���ظι���ȫ|����������ȫ|����"" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�asdfasdfasdfasdfasdf"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫXXXX|����|��������asdfasdf|���ظι���ȫasdfasdf|����������ȫ|����asdfasf"" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAAA"" obsid=""fac26638-6d75"" default=""����|������ȫ""/>" & vbNewLine & _
                "    <info name=""���������"" type=""ƽ���ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAA"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '��������
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷҳ��ȳ����Զ�����"" type=""��������"" index=""1"" value=""�ι���ȫ|������ȫ|����|��������|���ظι���ȫ|����������ȫ|����"" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""��������"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAA"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""��������"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAA"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '������ѡ
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷҳ��ȳ����Զ�����"" type=""������ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����|��������|���ظι���ȫ|����������ȫ|����"" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""������ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAA"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""������ѡ"" index=""1"" value=""�ι���ȫ|������ȫ|����"" class=""AAAA"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '�ı�
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷҳ��ȳ����Զ�����"" type=""�ı�"" index=""1"" value="""" class=""��˹�ٷ��������ķ��͵����͵����͵����͵����͵����͵����͵����͵����͵����͵����ǵķ��͵���˹�ٷ�"" obsid=""fac26638-6d75"" default=""����""/>" & vbNewLine & _
                "    <info name=""���������"" type=""�ı�"" index=""1"" value="""" class=""AAAA"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""���������"" type=""�ı�"" index=""1"" value="""" class=""AAAA"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
        strAsk = strPar
    ElseIf bytFunc = 4 Then
        '���ͨ��
        strPar = "{""recipes"":[]}"
        '���δͨ��
        strPar = "{""recipes"":[{""ORDER_ID"":2296,""NO_PASS_REASON"":""δͨ��""},{""ORDER_ID"":1924,""NO_PASS_REASON"":""δͨ��""},{""ORDER_ID"":2202,""NO_PASS_REASON"":""δͨ��""}]}"
    End If
    GetTestXML = strPar
End Function

Public Sub SetFormTranslucency(hWnd As Long, crKey As Long, bAlpha As Byte, dwFlags As Long) 'ʵ�ְ�͸������
'����:���ô���͸����
'hwnd,  ���ھ��
'crKey:ָ����Ҫ͸���ı�����ɫֵ������RGB()��
'bAlpha:����͸���ȣ�0��ʾ��ȫ͸����255��ʾ��͸��
'dwFlags: ͸����ʽdwFlags������ȡ����ֵ��
'       LWA_ALPHA=&H2ʱ��crKey������Ч��bAlpha������Ч��
'       LWA_COLORKEY=&H1�������е�������ɫΪcrKey�ĵط�����Ϊ͸����bAlpha������Ч���䳣��ֵΪ1��
'       LWA_ALPHA | LWA_COLORKEY��crKey�ĵط�����Ϊȫ͸�����������ط�����bAlpha����ȷ��͸���ȡ�
   Dim lngRet As Long
   
    lngRet = GetWindowLong(hWnd, GWL_EXSTYLE)
    lngRet = lngRet Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, lngRet
    SetLayeredWindowAttributes hWnd, crKey, bAlpha, dwFlags
 End Sub

Public Sub GetDrugInstructions(objfrmMain As Object, ByRef frmDrug As frmPassDrug, ByVal bytStyle As Byte, _
        ByVal strDrugCode As String, Optional ByVal strDrugName As String, Optional ByVal blnTip As Boolean)
'����:ҩƷ˵����
    Dim strRet As String
    
    If gblnBreak Then Exit Sub
    
    If frmDrug Is Nothing Then Set frmDrug = New frmPassDrug
    If strDrugCode <> "" Then
        If Not GetDrugInfo_ZL(strDrugCode, strRet) Then Exit Sub
    Else
        If blnTip Then Exit Sub
        strRet = """[{\""ͨ������\"":\""" & strDrugName & "\"",\""��Ʒ��\"":null,\""����ƴ��\"":null,\""Ӣ������\"":null,\""ҩ����\"":null,\""ҩ�����\"":null,\""������ҵ\"":null" & _
                        ",\""��׼�ĺ�\"":null,\""��ѧ����\"":null,\""��״\"":null,\""ҩ����\"":null,\""ҩ������ѧ\"":null," & _
                    "\""��Ӧ֢\"":null,\""�÷�����\"":null,\""������Ӧ\"":null,\""����֢\"":null," & _
                    "\""ע������\"":null,\""�и���ҩ\"":null,\""��ͯ��ҩ\"":null," & _
                    "\""�໥����\"":null,\""ҩ�����\"":null,\""��������\"":null}]"""
    End If
    If bytStyle = 1 Then
        Call gobjFrm.CloseGetDrugInstructions
    End If
 
    frmDrug.ShowMe objfrmMain, strRet, bytStyle, blnTip
    
End Sub

Public Function GetDrugInfo_ZL(ByVal strDrugCode As String, ByRef strDrugInfo As String) As Boolean
    Dim strUrl As String
    Dim strRet As String
    
    strUrl = "http://" & gstrDrugIP & ":" & gstrDrugPort & "/api/DrugInstructions/" & strDrugCode
    strRet = HttpGet(strUrl, responseText, , gblnBreak)
    WriteLog "" & glngModel, "GetDrugInstructions", "˵����URL:" & strUrl & ",����ֵ:" & strRet
    If strRet = "" Or gblnBreak Then
        Call gobjAir.OpenTransparentAirBubble(gobjFrm, "������ҩ���������쳣�Ͽ���", 2, 3, 0, vbWhite, vbRed, , 3, , , ����, True)
        gobjFrm.SetNotifyIcon
        gsngCheckLinkTime = Timer
        Exit Function
    End If
    If InStr(strRet, "errormsg") > 0 Then
        strRet = Replace(strRet, """{", "{")
        strRet = Replace(strRet, "}""", "}")
        strRet = Replace(strRet, "\""", """")
        strUrl = JSONParse("errormsg", strRet)
        If strUrl <> "" Then
            'MsgBox "ҩƷ˵����:" & vbCrLf & strURL, vbInformation + vbOKOnly, gstrSysName
            Call gobjAir.OpenTransparentAirBubble(gobjFrm, "ҩƷ˵����:" & strUrl, 2, 3, 0, vbWhite, vbRed, , 3, , 3000, ����, True)
            Exit Function
        End If
    End If
    strDrugInfo = strRet
    GetDrugInfo_ZL = True
End Function

Private Function GetDrugPlus(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal strAdvice As String) As ADODB.Recordset
    Dim strSQL As String
    Dim strPati As String
    
    On Error GoTo errH
    If str�Һŵ� <> "" Then
        strPati = " And A.�Һŵ� = [3] "
    Else
        strPati = " And A.����ID =[1] And A.��ҳID = [2] "
    End If
    strSQL = "Select a.Id, a.���id, a.�������, d.���� As ��ҩ;������, D.ID as ��ҩ;��ID" & vbNewLine & _
            " , a.����˵��,E.ҩƷ����,E.�������,E.������ " & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ D,ҩƷ���� E " & vbNewLine & _
            "Where a.���id = b.Id(+) And b.������Ŀid = d.Id(+) And A.������ĿID =E.ҩ��ID(+) " & strPati & " And a.���id <> 0 And Instr([4], ',' || a.Id || ',') > 0" & vbNewLine & _
            "Order By a.���"
    Set GetDrugPlus = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", lng����ID, lng��ҳID, str�Һŵ�, "," & strAdvice & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDoctorInfo(ByVal strDoctorName As String, ByVal byt���� As Byte) As ADODB.Recordset
    Dim strSQL As String
    'byt����-1סԺ;2-����
    'Ƹ�μ���ְ�� "1.����", "2.����", "3.�м�", "4.����/ʦ��", "5.Ա/ʿ", "9.��Ƹ"
    On Error GoTo errH
    If InStr(strDoctorName, ",") > 0 Then
        strSQL = "Select a.����, a.רҵ����ְ��, Decode(a.Ƹ�μ���ְ��, 1, '����', 2, '����', 3, '�м�', 4, '����/ʦ��', 5, 'Ա/ʿ', 9, '��Ƹ') As Ƹ�μ���ְ��, b.���� " & vbNewLine & _
                "From ��Ա�� A, ��Ա����ҩ��Ȩ�� B" & vbNewLine & _
                "Where a.Id = b.��Աid(+) And a.���� In (Select /*+cardinality(C,10)*/" & vbNewLine & _
                "                                  Column_Value" & vbNewLine & _
                "                                 From Table(f_Str2list([1])) C) And b.��¼״̬(+) = 1 And b.����(+)=[2]"
    Else
        strSQL = "Select a.����, a.רҵ����ְ��, Decode(a.Ƹ�μ���ְ��, 1, '����', 2, '����', 3, '�м�', 4, '����/ʦ��', 5, 'Ա/ʿ', 9, '��Ƹ') As Ƹ�μ���ְ��, b.����, b.����" & vbNewLine & _
                "From ��Ա�� A, ��Ա����ҩ��Ȩ�� B" & vbNewLine & _
                "Where a.Id = b.��Աid(+) And a.���� =[1] And b.��¼״̬(+) = 1  And b.����(+)=[2]"
    End If

    Set GetDoctorInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strDoctorName, byt����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutAdviceDiagsInfo(ByVal strAdviceIDs As String) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
     strSQL = "Select a.ҽ��id, b.�������" & vbNewLine & _
                "From �������ҽ�� A, ������ϼ�¼ B" & vbNewLine & _
                "Where a.���id = b.Id And a.ҽ��id In (Select /*+cardinality(C,10)*/" & vbNewLine & _
                "                                    Column_Value" & vbNewLine & _
                "                                   From Table(f_Num2list([1])) C)"


    Set GetOutAdviceDiagsInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strAdviceIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetParaURL(ByVal strSysName As String, ByVal strServiceName As String) As String
'����:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strUrl As String
    On Error GoTo errH
    strSQL = "Select �����ַ From ������������Ŀ¼ Where ϵͳ��ʶ = [1] And �������� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strSysName, strServiceName)
    If Not rsTmp.EOF Then strUrl = Trim(rsTmp!�����ַ & "")
    GetParaURL = strUrl
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



'------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------����ҩʦ��ϵͳ�ӿ�--------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetReviewResult(ByVal lngPatiID As Long, ByVal lngVisitID As Long, Optional ByRef rsRet As ADODB.Recordset, _
    Optional ByRef strAdvice As String = "0", Optional ByRef strAdviceID As String = "0") As Boolean
    '����:�������ѯ
    '     �󷽽����ѯ
    '     �̶����� http://192.168.0.231:8080//ords/zlrecipe/recipe/result
    '     �������� ?pid=20800808&pvid=1
    '     strURL = "http://192.168.0.231:8080//ords/zlrecipe/recipe/result?pid=20800808&pvid=1";
    '     ����ֵ:{""recipes"":[{""ORDER_ID"":123,""ORDER_GROUP_ID"":321,""NO_PASS_REASON"":""δͨ��""}]}

        Dim strUrl As String
        Dim strRet As String
        Dim strUnPass As String
        Dim strID As String
        
        Dim lngLength As Long
        Dim i As Long
        
        On Error GoTo errH
100     strUrl = GetParaURL("ҩʦ�������", "�������ѯ")
102     If strUrl = "" Then Exit Function
104     Set rsRet = InitAdviceRS(FUN_ҩʦ���_ZL)
106     strUrl = strUrl & "?pid=" & lngPatiID & "&pvid=" & lngVisitID
108     WriteLog "" & glngModel, "GetReviewResult", "ҩʦ����ѯURL:" & strUrl
110     strRet = HttpGet(strUrl, responseText, 1)
112     WriteLog "" & glngModel, "GetReviewResult", "ҩʦ����ѯ���:" & strRet
114     If strRet <> "" Then
116         lngLength = JSONParse("recipes.length", strRet)
118         If lngLength > 0 Then
120             For i = 0 To lngLength - 1
122                 rsRet.AddNew
124                 rsRet!ҽ��ID = JSONParse("recipes[" & i & "].ORDER_ID", strRet)
126                 rsRet!���ID = JSONParse("recipes[" & i & "].ORDER_GROUP_ID", strRet)
128                 If strAdvice <> "0" Then
130                     If InStr("," & strUnPass & ",", "," & rsRet!���ID & ",") = 0 Then
132                         strUnPass = strUnPass & "," & rsRet!���ID
                        End If
                    End If
134                 If strAdviceID <> "0" Then
136                     If InStr("," & strID & ",", "," & rsRet!ҽ��ID & ",") = 0 Then
138                         strID = strID & "," & rsRet!ҽ��ID
                        End If
                    End If
140                 rsRet!������� = JSONParse("recipes[" & i & "].NO_PASS_REASON", strRet)
142                 rsRet.Update
                Next
144             If strUnPass <> "" Then strAdvice = Mid(strUnPass, 2)
146             If strID <> "" Then strAdviceID = Mid(strID, 2)
            End If
        End If
148     GetReviewResult = True
150     If rsRet.RecordCount > 0 Then rsRet.MoveFirst
        Exit Function
errH:
152     MsgBox Err.Description & vbCrLf & "GetReviewResult" & "�� " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function EditPatiStatus() As Boolean
'����:�༭����״̬
'����:
    Dim strPath         As String
    Dim strUrl          As String
    Dim strPENVR       As String
    Dim strPvid         As String
    Dim bytPType         As Byte
    Dim i As Long
    
    If gstrStatusEdit = "" Then Exit Function
    If glngModel = PM_����ҽ���嵥 Then
        strPENVR = "10"
    ElseIf glngModel = PM_סԺҽ���嵥 Then
        strPENVR = "11"
    End If
    
    If gobjPati.str�Һŵ� <> "" Then
        strPvid = gobjPati.str�Һŵ�
        bytPType = 1
    Else
        strPvid = gobjPati.lng��ҳID & ""
        bytPType = 2
    End If
    strPath = Replace(UCase(App.Path), UCase("\Public"), "") & "\ZTHL"
    If Dir(strPath & "\nw.exe") <> "" Then
        strUrl = gstrStatusEdit & "?p=113:1:::::P1_ENVR_IN,P1_PID_IN,P1_PVID_IN,P1_RECORDER_IN,P1_RECORDER_ID_IN,P1_VISIT_TYPE_IN:" & _
                strPENVR & "," & gobjPati.lng����ID & "," & strPvid & "," & zlStr.Base64Encode(UserInfo.����) & "," & UserInfo.id & "," & bytPType
        Shell "cmd /c rd " & strPath & "\userData /s/q"
        WriteLog "" & glngModel, "EditPatiStatus", "����״̬�༭URL:" & strUrl & ",�ļ�·��:" & strPath
        i = ShellExecute(0, "open", "nw.exe", strUrl, strPath, SW_SHOWMAXIMIZED)
    Else
        MsgBox "����״̬�����ļ���" & strPath & "\nw.exe�������ڡ�" & vbCrLf & "�������ϵҽԺϵͳ����Ա��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    EditPatiStatus = True
End Function

Public Function GetXMLResult(ByVal rsRec As ADODB.Recordset)
'����:���췴��������ӦXML
    Dim i As Long
    Dim strXML As String
    For i = 1 To rsRec.RecordCount
        strXML = strXML & "    <info name=""" & rsRec!Name & """ type=""" & rsRec!Type & """ index=""" & _
            rsRec!Index & """ value=""" & rsRec!Default & """ obsid=""" & rsRec!Obsid & """/>" & vbNewLine
        rsRec.MoveNext
    Next
    GetXMLResult = Replace(strXML, """", "\""")
End Function

Private Sub AskPatiStatus(ByVal rsAdvice As ADODB.Recordset, ByRef strPara As String, ByVal lngPatiID As Long)
      '����:������״̬�Ѿ������,��������ʱ���ſ��Ѿ����������
      '����ֵ:strPara ���ط�����������

          Dim strUrl      As String
          Dim strData     As String
          Dim strRet      As String
          Dim strResult   As String
          Dim rsStatus    As ADODB.Recordset
          Dim rsAsk       As ADODB.Recordset
          Dim lngLength   As Long
          Dim i           As Long
          '���Ե�ַhttp://192.168.32.201:8888/bizdomain/6f73f15d-3718-4570-8cea-cf6282a6f6f6
10       On Error GoTo errH

20        strData = "{""ҽ���´�XML_IN"":""" & Replace(ZL_GET_Cusrules(rsAdvice), """", "\""") & """}"
30        WriteLog "" & glngModel, "AskPatiStatus", "��������URL:" & gstrIP & ",��������XML:" & strData
40        strRet = HttpPost(gstrIP, strData, responseText, , "Basic " & zlStr.Base64Encode("xxx:xxx"), , gblnBreak)
50        WriteLog "" & glngModel, "AskPatiStatus", "����������:" & strRet
60        If strRet <> "" Then
              '������������
              'Call GetTestXML(3, rsAdvice, strRet)
70            Set rsAsk = ZL_ParseXMLCusRules(strRet)
80            If gstrStatusGet <> "" Then
                  'strURL = "http://192.168.0.231:8080/ords/patstatus/pat/getpatstatus?pati_id_in=4989"
90                strUrl = gstrStatusGet & "?pati_id_in=" & lngPatiID
100               strRet = HttpGet(strUrl, responseText, 1)
110               WriteLog "" & glngModel, "AskPatiStatus", "����״̬��ѯURL:" & strUrl & vbCrLf & _
                                                            "����״̬��ѯ���:" & strRet
120               If strRet <> "" Then
130                   lngLength = JSONParse("patient_status.length", strRet)
140                   If lngLength > 0 Then
150                       Set rsStatus = InitAdviceRS(FUN_����״̬_ZL)
160                       For i = 0 To lngLength - 1
170                           rsStatus.AddNew
180                           rsStatus!STATUS_ID = JSONParse("patient_status[" & i & "].STATUS_ID", strRet)
190                           rsStatus!status_name = JSONParse("patient_status[" & i & "].STATUS_NAME", strRet)
200                           rsStatus!STATUS_SITUATION = JSONParse("patient_status[" & i & "].STATUS_SITUATION", strRet)
210                           rsStatus.Update
220                       Next
230                   End If
240               End If
250           End If
              '����ѯ�ʽ���
260           If Not rsAsk Is Nothing Then
270               If rsAsk.RecordCount > 0 Then
280                   If Not rsStatus Is Nothing Then
290                       rsStatus.Filter = "": rsAsk.Filter = ""
300                       If rsStatus.RecordCount > 0 Then
310                           For i = 1 To rsAsk.RecordCount
320                               rsStatus.Filter = "STATUS_NAME='" & rsAsk!Index & "'"
330                               If rsStatus.RecordCount > 0 Then
340                                   rsAsk!Default = IIf(rsStatus!STATUS_SITUATION & "" = "3", "��", "��")
350                               End If
360                               rsAsk.MoveNext
370                           Next
380                       End If
390                   End If
400                   rsAsk.Filter = ""
410                   If rsAsk.RecordCount > 0 Then
420                       If Not frmPassAsk.ShowMe(gfrmMain, rsAsk, strResult) Then Exit Sub
430                       If strResult <> "" Then strPara = Replace(strPara, "</cusrules>", "<result>" & strResult & "</result></cusrules>")
440                   End If
450
460               End If
470           End If
480       End If

490      Exit Sub
errH:
500       MsgBox "AskPatiStatus ������:" & Erl() & " �����:" & Err.Number & "��������:" & Err.Description, vbExclamation + vbOKOnly, gstrSysName
End Sub

