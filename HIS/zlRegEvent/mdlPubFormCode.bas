Attribute VB_Name = "mdlPubFormCode"
Option Explicit
'*********************************************************************************************************************************************
'ģ��˵��:����ҺŴ���ͨ�ù���
'1.�ؼ���ϣ�ȷ���ؼ������봰���еĿؼ�����һ�£��������
'   1)GetPayFromList:��֧���б��е����ݸ��µ�֧�����󼯺���
'   2)AddPayToList:��֧�����ݸ��µ�֧���б���
'   3)zlGetBalanceSQLByVsf:�Һ�ʱ���б��л�ȡδ�����sql
'   4)SelectMemo:ѡ��Һ�ժҪ
'   5)SetDelMemo:ѡ���˺�ժҪ
'   6)Load��סַ:������סַ
'   7)SetTxtTop:�ı��ı��򶨵㵫��������������Ӧ�仯
'       7.1)SetTxtLeft
'       7.2)SetTxtWidth
'2.ͨ�ù���
'   1)GetPayInfo:����ѡ���֧����ʽ��ȡ֧����Ϣ
'   2)zlIsAllowPatiChargeFeeMode:����Ƿ�����ı䲡���շ�ģʽ
'   3)zlCheckBackCard:����˺�ʱ���˿������Ƿ�Ϸ�
'   4)zlGetBackInvoice-��ȡ�˺ŷ�Ʊ
'   5)zlGetBalanceInfor:��ȡ�˺Ž�����Ϣ
'   6)GetAllҽ��-��ȡҽ���б�
'   7)zlCheckBackCard-�˺�ʱ���˿����
'   8)GetPatiIDByComminuty-���������Ż�ȡ����ID
'   9)GetColItem-��ȡ������ָ���Ľڵ�ֵ
'   10)GetRoom-��ȡ�Һŵ�����
'       10.1)GetRoomVisit
'  12)zlGetʧԼ��-��ȡ������ĳһ��.ԤԼʧԼ��
'3.��Һ���
'   1)Plug_PatiValiedCheck:��鲡���Ƿ���Ч����Ч���ֹ�Һ�
'*********************************************************************************************************************************************
Private mrsDoctor As ADODB.Recordset
'�ؼ����
'*********************************************************************************************************************************************
Public Function GetPayFromList(objRegInfor As clsRegEventInfor, ByVal vsfPay As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֧���б������ݱ��浽���󼯺�
    '���:objRegInfor-�Һ���Ϣ��vsfPay֧���б�
    '����:
    '����:���ϴ�
    '����:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim dbl֧����� As Double
    Dim lng�����id As Long, bln���ѿ� As Boolean
    Dim objPay As clsPayInfo, objPayReg As clsSubPayInfo, objCard As Card
    Dim str���㷽ʽ As String, strRows As String   '�Ѿ�����ͳ���˵�����
    
    On Error GoTo Errhand
    objRegInfor.objPayInfos.ReMoveAll
    With vsfPay
        For i = 1 To vsfPay.Rows - 1
            If InStr(strRows & ",", "," & i & ",") = 0 And .RowHidden(i) = False And (.RowData(i) <> 11 Or Val(.TextMatrix(i, .ColIndex("���"))) > 0) Then
                strRows = strRows & "," & i
                dbl֧����� = Val(.TextMatrix(i, .ColIndex("���"))) - Val(.TextMatrix(i, .ColIndex("��֧��")))
                lng�����id = Val(.TextMatrix(i, .ColIndex("�����ID")))
                bln���ѿ� = Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1
                str���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                '�����������֧������,�򲻴�ֵ
                If Not (RoundEx(dbl֧�����, 6) <= RoundEx(objRegInfor.objPayInfos.Card_������, 6) And _
                   lng�����id = objRegInfor.objPayInfos.Card_�����ID And _
                   bln���ѿ� = objRegInfor.objPayInfos.Card_���ѿ� And _
                   str���㷽ʽ = objRegInfor.objPayInfos.Card_���㷽ʽ And _
                   (objRegInfor.Card_�䶯���� = CP_���� Or objRegInfor.Card_�䶯���� = CP_�˿�)) Then
                   
                    Set objPay = New clsPayInfo 'clsPayInfo
                    objPay.���� = .TextMatrix(i, .ColIndex("֧����ʽ"))
                    
                    If lng�����id = objRegInfor.objPayInfos.Card_�����ID And _
                        bln���ѿ� = objRegInfor.objPayInfos.Card_���ѿ� And _
                        str���㷽ʽ = objRegInfor.objPayInfos.Card_���㷽ʽ And _
                        (objRegInfor.Card_�䶯���� = CP_���� Or objRegInfor.Card_�䶯���� = CP_�˿�) Then
                        dbl֧����� = dbl֧����� - objRegInfor.objPayInfos.Card_������
                        objRegInfor.objPayInfos.Card_У�Ա�־ = Val(.TextMatrix(i, .ColIndex("У�Ա�־")))
                        objRegInfor.objPayInfos.Card_����ɹ� = Val(.Cell(flexcpData, i, .ColIndex("У�Ա�־"))) = 1
                    End If
                    objPay.֧����� = dbl֧�����
                    objPay.���㷽ʽ = str���㷽ʽ
                    objPay.�������� = .RowData(i)
                    objPay.�ӿ���� = lng�����id
                    objPay.���ѿ� = bln���ѿ�
                    objPay.���ѿ�ID = Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                    objPay.���� = .TextMatrix(i, .ColIndex("����"))
                    objPay.У�Ա�־ = Val(.TextMatrix(i, .ColIndex("У�Ա�־")))
                    objPay.����ɹ� = Val(.Cell(flexcpData, i, .ColIndex("У�Ա�־"))) = 1 '�̶�=1
                    objPay.������ˮ�� = .TextMatrix(i, .ColIndex("������ˮ��"))
                    objPay.����˵�� = .TextMatrix(i, .ColIndex("����˵��"))
                    objPay.��������ID = Val(.TextMatrix(i, .ColIndex("��������ID")))
                    objPay.PayRow = i
                    
                    
                    Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                    objPayReg.PayRow = i
                    objPayReg.���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                    objPayReg.������ = dbl֧�����
                    objPayReg.������� = .TextMatrix(i, .ColIndex("�������"))
                    objPayReg.������ˮ�� = .TextMatrix(i, .ColIndex("������ˮ��"))
                    objPayReg.����˵�� = .TextMatrix(i, .ColIndex("����˵��"))
                    objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                    
                    If objPay.�ӿ���� > 0 Then
                        For j = i + 1 To vsfPay.Rows - 1
                            If objPay.�ӿ���� = Val(.TextMatrix(j, .ColIndex("�����ID"))) And objPay.���ѿ� = (Val(.TextMatrix(j, .ColIndex("���ѿ�"))) = 1) Then
                                str���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                                dbl֧����� = Val(.TextMatrix(j, .ColIndex("���"))) - Val(.TextMatrix(j, .ColIndex("��֧��")))
                                If lng�����id = objRegInfor.objPayInfos.Card_�����ID And _
                                    bln���ѿ� = objRegInfor.objPayInfos.Card_���ѿ� And _
                                    str���㷽ʽ = objRegInfor.objPayInfos.Card_���㷽ʽ And _
                                    (objRegInfor.Card_�䶯���� = CP_���� Or objRegInfor.Card_�䶯���� = CP_�˿�) Then
                                    dbl֧����� = dbl֧����� - objRegInfor.objPayInfos.Card_������
                                End If
                                
                                strRows = strRows & "," & j
                                Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                                objPayReg.PayRow = j
                                objPayReg.���㷽ʽ = .TextMatrix(j, .ColIndex("���㷽ʽ"))
                                objPayReg.������ = dbl֧�����
                                objPayReg.������� = .TextMatrix(j, .ColIndex("�������"))
                                objPayReg.������ˮ�� = .TextMatrix(j, .ColIndex("������ˮ��"))
                                objPayReg.����˵�� = .TextMatrix(j, .ColIndex("����˵��"))
                                objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                            End If
                        Next
                    ElseIf .RowData(i) = 3 Or .RowData(i) = 4 Then 'ҽ��֧�ֶ��ֽ��㷽ʽ��Ҳ����ͬ�ķ�ʽ����
                        For j = i + 1 To vsfPay.Rows - 1
                            If (.RowData(j) = 3 Or .RowData(j) = 4) And Val(.TextMatrix(j, .ColIndex("�����ID"))) = 0 Then
                                dbl֧����� = Val(.TextMatrix(j, .ColIndex("���"))) - Val(.TextMatrix(j, .ColIndex("��֧��")))
                                strRows = strRows & "," & j
                                Set objPayReg = New clsSubPayInfo 'clsSubPayInfo
                                objPayReg.PayRow = j
                                objPayReg.���㷽ʽ = .TextMatrix(j, .ColIndex("���㷽ʽ"))
                                objPayReg.������ = dbl֧�����
                                objPayReg.������� = .TextMatrix(j, .ColIndex("�������"))
                                objPayReg.������ˮ�� = .TextMatrix(j, .ColIndex("������ˮ��"))
                                objPayReg.����˵�� = .TextMatrix(j, .ColIndex("����˵��"))
                                objPay.AddItem objPayReg, "K" & objPayReg.PayRow
                            End If
                        Next
                    End If
                    
                    If objPay.�ӿ���� <> 0 Then ' ���ڲ�����������
                        If objPay.���ѿ� Then
                            objPay.֧������ = Pay_SquarePay
                        Else
                            objPay.֧������ = Pay_ThreePay
                        End If
                        objRegInfor.objPayInfos.AddItem objPay, IIf(objPay.���ѿ�, "X", "K") & objPay.�ӿ����
                    Else
                        objPay.֧������ = Decode(vsfPay.RowData(i), 11, Pay_AccountPay, 2, Pay_CashPay, 3, Pay_InsurePay, 4, Pay_InsurePay, Pay_CashPay)
                        objRegInfor.objPayInfos.AddItem objPay, "PAY" & objPay.֧������ & "_" & objPay.PayRow
                    End If
                    If objPay.֧������ = Pay_AccountPay Then
                        objRegInfor.objPayInfos.Ԥ���� = objPay.֧�����
                    End If
                Else
                    objRegInfor.objPayInfos.Card_У�Ա�־ = Val(.TextMatrix(i, .ColIndex("У�Ա�־")))
                    objRegInfor.objPayInfos.Card_����ɹ� = Val(.Cell(flexcpData, i, .ColIndex("У�Ա�־"))) = 1
                    objRegInfor.objPayInfos.Card_PayRow = i
                End If
            End If
        Next
    End With
    GetPayFromList = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdatePayToList(objPayInfos As clsPayInfos, ByVal vsfPay As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �����б��еĽ��㷽ʽ,��Ҫ�����˿�ɹ������У�Ա�־=2
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/3/23 09:13
    '---------------------------------------------------------------------------------------
    Dim objPay As clsPayInfo, objSubPay As clsSubPayInfo
    On Error GoTo Errhand
    For Each objPay In objPayInfos
        If objPay.Count > 0 Then
            For Each objSubPay In objPay
                If objSubPay.PayRow > 0 Then
                    With vsfPay
                        .TextMatrix(objSubPay.PayRow, .ColIndex("У�Ա�־")) = objPay.У�Ա�־
                        If objPay.У�Ա�־ = 2 Then
                            .Cell(flexcpForeColor, objSubPay.PayRow, 0, objSubPay.PayRow, .Cols - 1) = 0
                        End If
                    End With
                End If
            Next
        Else
            If objPay.PayRow > 0 Then
                With vsfPay
                    .TextMatrix(objPay.PayRow, .ColIndex("У�Ա�־")) = objPay.У�Ա�־
                    If objPay.У�Ա�־ = 2 Then
                        .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
                    End If
                End With
            End If
        End If
    Next
    If objPayInfos.Card_PayRow > 0 Then
        With vsfPay
            .TextMatrix(objPayInfos.Card_PayRow, .ColIndex("У�Ա�־")) = objPayInfos.Card_У�Ա�־
            If objPayInfos.Card_У�Ա�־ = 2 Then
                .Cell(flexcpForeColor, objPayInfos.Card_PayRow, 0, objPayInfos.Card_PayRow, .Cols - 1) = 0
            End If
        End With
    End If
    UpdatePayToList = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddPayToList(ByVal objPay As clsPayInfo, ByVal vsfPay As VSFlexGrid, _
                Optional ByVal bln�쳣���� As Boolean, Optional ByVal byt֧������ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϣ���µ�֧���б���
    '���:objPayInfo-������Ϣ
    '     byt֧������ - 0-δʵ��֧���������һ����ɣ�1-��֧����ɣ�2-֧�����ڽ����У�δ��ȷ���
    '����:
    '����:���ϴ�
    '����:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim objSubPay As clsSubPayInfo
    Dim str���㷽ʽ As String
    
    If objPay Is Nothing Then AddPayToList = True: Exit Function
    If objPay.֧����� = 0 Then AddPayToList = True: Exit Function
    With vsfPay
        If objPay.Count > 0 Then
            .RemoveItem objPay.PayRow
            For Each objSubPay In objPay
                If str���㷽ʽ <> objSubPay.���㷽ʽ Then str���㷽ʽ = objSubPay.���㷽ʽ: .Rows = .Rows + 1
                .RowData(.Rows - 1) = objPay.��������
                .TextMatrix(.Rows - 1, .ColIndex("֧����ʽ")) = objSubPay.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(Val(.TextMatrix(.Rows - 1, .ColIndex("���"))) + objSubPay.������, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = objSubPay.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("�������")) = objSubPay.�������
                .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = objPay.�ӿ����
                .TextMatrix(.Rows - 1, .ColIndex("���ѿ�")) = IIf(objPay.���ѿ�, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("���ѿ�ID")) = objPay.���ѿ�ID
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objPay.����
                .TextMatrix(.Rows - 1, .ColIndex("��������ID")) = objPay.��������ID
                .TextMatrix(.Rows - 1, .ColIndex("�޸�")) = IIf(byt֧������ = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("����޸�")) = IIf(byt֧������ = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 2, 1)
                .Cell(flexcpData, .Rows - 1, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 1, 0) '�̶�
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = IIf(objPay.��������, 1, 0)
                If byt֧������ = 2 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 0
                End If
            Next
        Else
            '���ָ������Ч������뵽���һ����
            If objPay.PayRow = 0 Or objPay.PayRow > .Rows - 1 Then
                objPay.PayRow = 0
                For lngRow = 1 To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("֧����ʽ"))) = "" Then
                        objPay.PayRow = lngRow: Exit For
                    End If
                Next
                If objPay.PayRow = 0 Then
                    objPay.PayRow = .Rows
                    .Rows = .Rows + 1
                End If
            ElseIf bln�쳣���� Then
                .RemoveItem objPay.PayRow
                objPay.PayRow = .Rows
                .Rows = .Rows + 1
            End If
            .RowData(objPay.PayRow) = objPay.��������
            .TextMatrix(objPay.PayRow, .ColIndex("֧����ʽ")) = objPay.����
            .TextMatrix(objPay.PayRow, .ColIndex("���")) = Format(Val(.TextMatrix(objPay.PayRow, .ColIndex("���"))) + objPay.֧�����, "0.00")
            .TextMatrix(objPay.PayRow, .ColIndex("���㷽ʽ")) = objPay.���㷽ʽ
            .TextMatrix(objPay.PayRow, .ColIndex("�������")) = objPay.�������
            .TextMatrix(objPay.PayRow, .ColIndex("�����ID")) = objPay.�ӿ����
            .TextMatrix(objPay.PayRow, .ColIndex("���ѿ�")) = IIf(objPay.���ѿ�, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("���ѿ�ID")) = objPay.���ѿ�ID
            .TextMatrix(objPay.PayRow, .ColIndex("����")) = objPay.����
            .TextMatrix(objPay.PayRow, .ColIndex("��������ID")) = objPay.��������ID
            .TextMatrix(objPay.PayRow, .ColIndex("�޸�")) = IIf(byt֧������ = 1, 0, 1)
            .TextMatrix(objPay.PayRow, .ColIndex("����޸�")) = IIf(byt֧������ = 0, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 2, 1)
            .Cell(flexcpData, objPay.PayRow, .ColIndex("У�Ա�־")) = IIf(byt֧������ = 1, 1, 0) '�̶�
            .TextMatrix(objPay.PayRow, .ColIndex("��������")) = IIf(objPay.��������, 1, 0)
            If byt֧������ = 2 Then
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
            End If
        End If
    End With
End Function

Public Function zlGetBalanceSQLByVsf(ByVal objRegInfor As clsRegEventInfor, ByVal vsfPay As VSFlexGrid, _
                                    cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� :
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim dbl��� As Double
    Dim str���㷽ʽ As String, str������Ϣ As String, strSQL As String
    Dim PayType As gPagePay
    Dim i As Long
    On Error GoTo errH
    If cllPro Is Nothing Then Set cllPro = New Collection
    With vsfPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("У�Ա�־"))) = 1 And (Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1 Or Val(.TextMatrix(i, .ColIndex("�����ID"))) = 0) Then
                dbl��� = Val(.TextMatrix(i, .ColIndex("���"))) - Val(.TextMatrix(i, .ColIndex("��֧��")))
                str���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                If str���㷽ʽ = objRegInfor.objPayInfos.Card_���㷽ʽ And objRegInfor.objPayInfos.Card_������ <> 0 Then
                    dbl��� = dbl��� - objRegInfor.objPayInfos.Card_������
                    strSQL = zlGetCardFeeModifySQL(False, objRegInfor.objPayInfos.Card_���ݺ�, objRegInfor.objPayInfos.Card_����ID, str���㷽ʽ, _
                                objRegInfor.objPayInfos.Card_������, , , Val(.TextMatrix(i, .ColIndex("�����ID"))), _
                                IIf(Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1, True, False), .TextMatrix(i, .ColIndex("����")), , , , .TextMatrix(i, .ColIndex("�������")))
                    Call zlAddArray(cllPro, strSQL)
                End If
                If RoundEx(dbl���, 6) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("���ѿ�"))) = 1 Then
                        str������Ϣ = str���㷽ʽ & "," & dbl���
                        PayType = Pay_SquarePay
                    Else
                        str������Ϣ = str���㷽ʽ & "," & dbl��� & "," & .TextMatrix(i, .ColIndex("�������")) & ", "
                        PayType = Pay_CashPay
                    End If
                    strSQL = zlGetRegFeeModifySQL(False, objRegInfor.objPayInfos.Reg_���ݺ�, objRegInfor.objPayInfos.Reg_����ID, str������Ϣ, PayType, , , , , _
                                Val(.TextMatrix(i, .ColIndex("�����ID"))), .TextMatrix(i, .ColIndex("����")))
                    Call zlAddArray(cllPro, strSQL)
                End If
            End If
        Next
    End With
'    '����Ϊ0ʱ��������,ȱʡΪ�ֽ𣬽��㷽ʽ���ܲ����б���
'    If objRegInfor.Card_�䶯���� = CP_���� And objRegInfor.objPayInfos.Card_������ = 0 Then
'        strSql = zlGetCardFeeModifySQL(False, objRegInfor.objPayInfos.Card_���ݺ�, objRegInfor.objPayInfos.Card_����ID, objRegInfor.objPayInfos.Card_���㷽ʽ, 0)
'        Call zlAddArray(cllPro, strSql)
'    End If
    zlGetBalanceSQLByVsf = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SelectMemo(frmMain As Form, cbo��ע As ComboBox, ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ����ժҪ
    '���:strInput-���봮;Ϊ��ʱ,��ʾȫ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(cbo��ע.Text) Then
             strWhere = " And  ���� like [1] "
        ElseIf zlCommFun.IsNumOrChar(cbo��ע.Text) Then
             strWhere = " And (���� like upper([1]) or ���� like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,����,����,����  " & _
     "   From ���ùҺ�ժҪ " & _
     "   Where 1=1 " & strWhere & _
     "   Order by ȱʡ��־"
     vRect = zlControl.GetControlRect(cbo��ע.Hwnd)
     On Error GoTo Hd
     Set rsInfo = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "���ùҺ�ժҪ", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cbo��ע.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "û�����ó��ùҺ�ժҪ,�����ֵ����������", vbInformation, gstrSysName
        End If
        zlCommFun.PressKey vbKeyTab: Exit Function
     End If
     Call zlControl.CboSetText(cbo��ע, Nvl(rsInfo!����))
     cbo��ע.Tag = Nvl(rsInfo!����)
     If cbo��ע.Visible And cbo��ע.Enabled Then cbo��ע.SetFocus
     zlCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function SetDelMemo(ByVal cbo��ע As ComboBox, ByVal strInput As String) As Boolean
    Dim rsMemo As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    cbo��ע.Clear
    If strInput = "" Then
        strSQL = "Select ����,ȱʡ��־ From �����˺�ԭ�� Order By ȱʡ��־ Desc,����"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, "SetDelMemo")
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo��ע.AddItem rsMemo!����
                If Val(Nvl(rsMemo!ȱʡ��־)) = 1 Then
                    '������Click�¼�
                    Call cbo.SetIndex(cbo��ע.Hwnd, cbo��ע.NewIndex): cbo��ע.Tag = cbo��ע.Text
                End If
                rsMemo.MoveNext
            Loop
        End If
    Else
        strSQL = "Select ����,ȱʡ��־,����,���� From �����˺�ԭ�� Order By ȱʡ��־ Desc,����"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, "SetDelMemo")
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo��ע.AddItem rsMemo!����
                If Nvl(rsMemo!����) Like UCase(strInput) & "*" Or Nvl(rsMemo!����) Like UCase(strInput) & "*" Or Nvl(rsMemo!����) Like strInput & "*" Then
                    '������Click�¼�
                    Call cbo.SetIndex(cbo��ע.Hwnd, cbo��ע.NewIndex): cbo��ע.Tag = cbo��ע.Text
                End If
                rsMemo.MoveNext
            Loop
            If cbo��ע.Text = "" Then
                MsgBox "û���ҵ���Ӧ���˺�ԭ��,����������", vbInformation, gstrSysName
                SetDelMemo = False
                Exit Function
            End If
        End If
    End If
    SetDelMemo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function Load��סַ(ByVal frmMain As Form) As ADODB.Recordset
    Dim strSQL As String, strFile As String
    Dim fld As Field, rsCheck As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim rsCopy As ADODB.Recordset
    Dim rsNew As ADODB.Recordset, rs��סַ As New ADODB.Recordset
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        rs��סַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
        If rs��סַ.RecordCount > 0 Then
            rs��סַ!���� = rs��סַ!���� + 1
        End If
        If Err <> 0 Then
            rs��סַ.Close
        Else
            rs��סַ!���� = rs��סַ!���� - 1
        End If
    End If
    Err.Clear: On Error GoTo errH
    
    If rs��סַ.State = 0 Then
        strSQL = "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ����"
        Set rsCopy = zlDatabase.OpenSQLRecord(strSQL, "Load��סַ")     '������adUseClient���ܽ�����
        Set rs��סַ = zlDatabase.CopyNewRec(rsCopy)
        If Not rs��סַ.EOF Then
            '��������:����,����
            Set fld = rs��סַ.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = rs��סַ.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            rs��סַ.Save strFile, adPersistADTG
        End If
        rs��סַ.Close
        rs��סַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
    Else
        strSQL = "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ���� Where 1 = 0"
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Load��סַ")
        If rsCheck.Fields(1).DefinedSize > rs��סַ.Fields(1).DefinedSize Or rsCheck.Fields(2).DefinedSize > rs��סַ.Fields(2).DefinedSize Then
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            strSQL = "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ����"
            Set rsCopy = zlDatabase.OpenSQLRecord(strSQL, "Load��סַ")
            Set rsNew = zlDatabase.CopyNewRec(rsCopy)
            rsNew.Save strFile, adPersistXML
            rs��סַ.Close
            rs��סַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
        End If
    End If
    
    frmMain.lbl��סַ.ToolTipText = "�붨�ڱ��ݱ���[������סַ]�����ļ�:" & strFile
    Set Load��סַ = rs��סַ
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set Load��סַ = rs��סַ
End Function

Public Function SetTxtTop(ByVal txtThis As TextBox, ByVal lngTop As Long)
    '---------------------------------------------------------------------------------------
    ' ���� : �����ؼ�λ�ã���������ؼ��������������Ӧ�仯
    ' ��� : txtThis-��Ҫ�����Ŀؼ�
    '        lngLeft-��߾���
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Top = lngTop
    
    txtThis.FontSize = objFont.Size
End Function

Public Function SetTxtLeft(ByVal txtThis As TextBox, ByVal lngLeft As Long)
    '---------------------------------------------------------------------------------------
    ' ���� : �����ؼ�λ�ã���������ؼ��������������Ӧ�仯
    ' ��� : txtThis-��Ҫ�����Ŀؼ�
    '        lngLeft-��߾���
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Left = lngLeft
    
    txtThis.FontSize = objFont.Size
End Function

Public Function SetTxtWidth(ByVal txtThis As TextBox, ByVal lngWidth As Long)
    '---------------------------------------------------------------------------------------
    ' ���� : �����ؼ�λ�ã���������ؼ��������������Ӧ�仯
    ' ��� : txtThis-��Ҫ�����Ŀؼ�
    '        lngWidth-���
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/5/17 13:52
    '---------------------------------------------------------------------------------------
    Dim objFont As StdFont
    On Error Resume Next
    Set objFont = New StdFont
    objFont.Size = txtThis.FontSize
    
    txtThis.FontSize = 1
    txtThis.Width = lngWidth
    
    txtThis.FontSize = objFont.Size
End Function

'ͨ�ù���
'*********************************************************************************************************************************************
Public Function GetPayInfo(ByVal colCardPayMode As Collection, ByVal str���㷽ʽ As String, _
                            objPayInfo As clsPayInfo) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ݽ������ƻ�ȡ������Ϣ
    ' ��� : colCardPayMode:֧����Ϣ���ϣ��������֧����ʽʱ��ʼ��
    '      : str���㷽ʽ :��Ҫ��ȡ�Ľ��㷽ʽ
    ' ���� : objPayInfo�������������ʡ��������ʡ����㷽ʽ���ӿ���š��Ƿ����ѿ���֧�����͡��Ƿ��������
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2018/11/20 09:30
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    If objPayInfo Is Nothing Then Set objPayInfo = New clsPayInfo
    If str���㷽ʽ = "" Then GetPayInfo = True: Exit Function
    
    
    If str���㷽ʽ = "Ԥ����" Or str���㷽ʽ = "��Ԥ��" Then
        objPayInfo.���� = str���㷽ʽ
        objPayInfo.֧������ = Pay_AccountPay
        objPayInfo.�������� = 11
        GetPayInfo = True: Exit Function
    End If
    
    '���ȳ伯���в���,������һ�����⣬ҽ�ƿ����ƺ����ѿ�������ͬ
    'colCardPayMode:���ƣ����ʣ����㷽ʽ�������ID���Ƿ����ѿ����Ƿ��������
    For i = 1 To colCardPayMode.Count
        If colCardPayMode(i)(0) = str���㷽ʽ Then
            objPayInfo.���� = str���㷽ʽ
            objPayInfo.�������� = Val(colCardPayMode(i)(1))
            objPayInfo.���㷽ʽ = colCardPayMode(i)(2)
            objPayInfo.�ӿ���� = Val(colCardPayMode(i)(3))
            objPayInfo.���ѿ� = Val(colCardPayMode(i)(4)) = 1
            objPayInfo.�������� = Val(colCardPayMode(i)(5)) = 1
            If objPayInfo.�ӿ���� > 0 Then
                If objPayInfo.���ѿ� Then
                    objPayInfo.֧������ = Pay_SquarePay
                Else
                    objPayInfo.֧������ = Pay_ThreePay
                End If
            Else
                objPayInfo.֧������ = Pay_CashPay
            End If
            GetPayInfo = True
            Exit Function
        End If
    Next
    ' ʲôʱ��ѡ���˽��㷽ʽ�����ǻ�����û�е�
    MsgBox "��Ч�Ľ��㷽ʽ", vbInformation, gstrSysName
    Exit Function
    
    strSQL = "Select ���� From ���㷽ʽ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str���㷽ʽ)
    If Not rsTemp.EOF Then
        objPayInfo.���� = str���㷽ʽ
        objPayInfo.�������� = Val(rsTemp!����)
        objPayInfo.���㷽ʽ = objPayInfo.����
    Else 'nexttodo
'        strSql = "Select 1 As ����, a.Id, b.����, b.����, A.�Ƿ��������" & vbNewLine & _
'                "From ҽ�ƿ���� a, ���㷽ʽ b" & vbNewLine & _
'                "Where a.���㷽ʽ = b.���� And a.���� = [1]" & vbNewLine & _
'                "Union" & vbNewLine & _
'                "Select 2 As ����, c.��� As Id, d.����, d.����, 0 as �Ƿ��������" & vbNewLine & _
'                "From ���ѿ����Ŀ¼ c, ���㷽ʽ d" & vbNewLine & _
'                "Where c.���㷽ʽ = d.���� And c.���� = [1]"
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetPayInfo", str���㷽ʽ)
'        If rsTemp.EOF Then
'            MsgBox str���㷽ʽ & "����Ч�Ľ��㷽ʽ����ѡ������֧����ʽ��", vbInformation, gstrSysName
'            Exit Function
'        End If
'        objPayInfo.���� = str���㷽ʽ
'        objPayInfo.�������� = Val(Nvl(rsTemp!����))
'        objPayInfo.���㷽ʽ = Nvl(rsTemp!����)
'        objPayInfo.�ӿ���� = Val(Nvl(rsTemp!ID))
'        objPayInfo.���ѿ� = Val(Nvl(rsTemp!����)) = 2
'        objPayInfo.�������� = Val(Nvl(rsTemp!�Ƿ��������)) = 1
'        If objPayInfo.�ӿ���� > 0 Then
'            If objPayInfo.���ѿ� Then
'                objPayInfo.֧������ = Pay_SquarePay
'            Else
'                objPayInfo.֧������ = Pay_ThreePay
'            End If
'        Else
'            objPayInfo.֧������ = Pay_CashPay
'        End If
    End If
    GetPayInfo = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function zlGetBackInvoice(ByVal strNO As String, ByRef strBackInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���շ�Ʊ
    ' ��� : strNO-�˺ŵ���
    '        strBackInvoice-�˺��漰�ķ�Ʊ��
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    
    On Error GoTo errH
    '��ȡ�ջ�Ʊ��
    strSQL = _
    "   Select A.����" & vbNewLine & _
    "   From Ʊ��ʹ����ϸ A" & vbNewLine & _
    "   Where A.���� = 1 And a.ԭ�� <> 6 " & vbNewLine & _
    "           And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
    "Minus" & vbNewLine & _
    "Select A.����" & vbNewLine & _
    "From Ʊ��ʹ����ϸ A" & vbNewLine & _
    "Where A.���� = 2 And a.ԭ�� <> 6 " & vbNewLine & _
    "   And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
    "Order By ����"
    Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ջ�Ʊ��", strNO, 4)
    Do While Not rsInvoice.EOF
        strBackInvoice = strBackInvoice & "," & rsInvoice!����
        rsInvoice.MoveNext
    Loop
    If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
    zlGetBackInvoice = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBookingNO(ByVal strInput As String) As String
    Dim objInterCard As clsInterFaceCard
    Dim lngԤԼʧԼ���� As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPatiIds As String
    
    If zlCreateOneCardObject(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser, objInterCard) = False Then Exit Function
    lngԤԼʧԼ���� = Val(zlDatabase.GetPara("ԤԼʧԼ����", glngSys, glngModul, 0))
    If Len(strInput) = 8 And InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(strInput, 1, 1))) > 0 And IsNumeric(Mid(strInput, 2)) Then
        strInput = UCase(strInput)
        strSQL = "" & _
                "Select Min(A.NO) NO" & vbNewLine & _
                "From ������ü�¼ A" & vbNewLine & _
                "Where A.��¼���� = 4 And A.��¼״̬ = 0 And A.No = [1] " & _
                IIf(lngԤԼʧԼ���� > 0, " And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
                "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [2]) or  (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [3])) ")
    Else
        If objInterCard.GetPatiIdsByRange(strInput, strPatiIds) = False Then Exit Function
        If strPatiIds = "" Then Exit Function
        strInput = strPatiIds
        strSQL = "" & _
            "Select /*+cardinality(B,10) */ Min(A.NO) NO" & vbNewLine & _
            "From ������ü�¼ A, Table(f_num2list([1])) B" & vbNewLine & _
            "Where A.��¼���� = 4 And A.��¼״̬ = 0 And A.����id = B.Column_Value(+) " & _
            IIf(lngԤԼʧԼ���� > 0, " And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [2]) or (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [3])) ")
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetBookingNO", strInput, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    GetBookingNO = "" & rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAbNormalNO(ByVal strInput As String) As String
    Dim objInterCard As clsInterFaceCard
    Dim lngԤԼʧԼ���� As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPatiIds As String
    
    If zlCreateOneCardObject(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser, objInterCard) = False Then Exit Function
    lngԤԼʧԼ���� = Val(zlDatabase.GetPara("ԤԼʧԼ����", glngSys, glngModul, 0))
    If Len(strInput) = 8 And InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(strInput, 1, 1))) > 0 And IsNumeric(Mid(strInput, 2)) Then
        strInput = UCase(strInput)
        strSQL = " And A.NO = [1]"
        strSQL = "" & _
                "Select Min(A.NO) NO" & vbNewLine & _
                "From ������ü�¼ A" & vbNewLine & _
                "Where A.��¼���� = 4 And A.��¼״̬ = 1 And A.No = [1] " & _
                IIf(lngԤԼʧԼ���� > 0, " And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
                "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [2]) or  (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [3])) ")
    Else
        If objInterCard.GetPatiIdsByRange(strInput, strPatiIds) = False Then Exit Function
        If strPatiIds = "" Then Exit Function
        strInput = strPatiIds
        strSQL = "" & _
            "Select /*+cardinality(B,10) */ Min(A.NO) NO" & vbNewLine & _
            "From ������ü�¼ A, Table(f_num2list([1])) B" & vbNewLine & _
            "Where A.��¼���� = 4 And A.��¼״̬ = 1 And A.����id = B.Column_Value(+) " & _
            IIf(lngԤԼʧԼ���� > 0, " And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [2]) or (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [3])) ")
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAbNormalNO", strInput, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    GetAbNormalNO = "" & rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetBalanceInfor(ByVal bytMode As Byte, ByVal objRegInfor As clsRegEventInfor, _
                        ByVal strAdvance As String, ByRef strȱʡ���� As String, _
                        str���� As String, strBalance As String, str�����ѿ� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�˺ŵĽ�����Ϣ
    ' ��� : bytMode-�˺�ģʽ��4�������˷ѣ�5-���ϣ�6-����
    '        objRegInfor-���ιҺ���Ϣ
    '        strAdvance- ҽ����֧�����ֵĽ��㷽ʽ
    ' ���� : str���� - ��������
    '        strBalance - �˽ӿ�֧��,��Ҫ�ȱ���У������
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim objPay As clsPayInfo, objSubPay As clsSubPayInfo
    Dim str�ֽ� As String, dbl�ֽ� As Double
    
    On Error GoTo errH
    For Each objPay In objRegInfor.objPayInfos
        If bytMode = 4 Or (objPay.У�Ա�־ <> 2 And bytMode = 6) Or (bytMode = 5 And objPay.У�Ա�־ <> 1) Then
            If objPay.֧������ = Pay_CashPay Then
                If objPay.�������� = 1 Then
                    str�ֽ� = objPay.���㷽ʽ: strȱʡ���� = objPay.���㷽ʽ
                    dbl�ֽ� = dbl�ֽ� + objPay.֧�����
                ElseIf objPay.֧����� <> 0 Then
                    str���� = str���� & "|" & objPay.���㷽ʽ & "," & objPay.֧����� & "," & objPay.������� & ", "
                End If
            ElseIf objPay.֧������ = Pay_InsurePay Then
                For Each objSubPay In objPay
                    If InStr(strAdvance, objSubPay.���㷽ʽ) <> 0 Then
                        dbl�ֽ� = dbl�ֽ� + objSubPay.������
                        If str�ֽ� = "" Then str�ֽ� = strȱʡ����
                    Else
                        strBalance = strBalance & "|" & objSubPay.���㷽ʽ & "," & objSubPay.������ & "," & 0
                    End If
                Next
            ElseIf objPay.֧������ = Pay_SquarePay Then
                str�����ѿ� = str�����ѿ� & "|" & objPay.���㷽ʽ & "," & objPay.֧�����
            ElseIf objPay.֧������ = Pay_ThreePay Then
                For Each objSubPay In objPay
                    strBalance = strBalance & "|" & objSubPay.���㷽ʽ & "," & objSubPay.������ & "," & 1
                Next
            End If
        End If
    Next
    If RoundEx(dbl�ֽ�, 6) <> 0 Then
        str���� = str���� & "|" & str�ֽ� & "," & dbl�ֽ�
    End If
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If str�����ѿ� <> "" Then str�����ѿ� = Mid(str�����ѿ�, 2)
    If str���� <> "" Then str���� = Mid(str����, 2)
    
    zlGetBalanceInfor = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAllҽ��() As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    If mrsDoctor Is Nothing Then
        '��Ա���ʹ̶���ֱ��д��sql�У����԰󶨱�������
        strSQL = "Select a.Id, a.����, Upper(a.����) As ����,b.����id,a.���" & _
                " From ��Ա�� a, ������Ա b, ��Ա����˵�� c" & _
                " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = 'ҽ��' And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order By a.���� Desc"
        Set mrsDoctor = zlDatabase.OpenSQLRecord(strSQL, "GetAllҽ��")
    End If
    Set GetAllҽ�� = mrsDoctor
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckRegistAppointment(ByVal strNO As String) As Boolean
    '���ԤԼ��¼�Ƿ񱻽���
    'True-ԤԼ��¼δ����;False-ԤԼ��¼�ѱ�����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    strSQL = "Select 1 From ���˹Һż�¼ Where NO = [1] And ����ʱ�� Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckRegistAppointment", strNO)
    If Not rsTmp.EOF Then
        CheckRegistAppointment = True
    Else
        CheckRegistAppointment = False
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    Err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    Err.Clear: On Error GoTo 0
End Function

Public Function GetRoom(ByVal lng����ID As Long, ByVal lng�ƻ�ID As Long) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSQL As String, str�ű� As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    '139670��2019/6/27�����ϴ����շ�ԤԼδ���ĹҺ�ʱ�����ݼƻ�ID��ȡ���﷽ʽ�Լ�ָ�����������
    '�������﷽ʽ��Ϊ����ͳ�ƺ��������Լ�ͳ�Ƶļ���ҽԺ��ûʹ�ã����β�����
    If lng�ƻ�ID = 0 Then
        strSQL = "Select ����,Nvl(���﷽ʽ,0) as ���� From �ҺŰ��� Where ID=[1]"
    Else
        strSQL = "Select ����,Nvl(���﷽ʽ,0) as ���� From �ҺŰ��żƻ� Where ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng����ID, lng�ƻ�ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    str�ű� = Nvl(rsTmp!����)
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        If lng�ƻ�ID = 0 Then
            strSQL = "Select �������� From �ҺŰ������� Where �ű�ID=[1]"
        Else
            strSQL = "Select �������� From �Һżƻ����� Where �ƻ�ID=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng����ID, lng�ƻ�ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select ��������,0 as NUM From �ҺŰ������� Where �ű�ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And ��¼����=1 and ��¼״̬=1 and  ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű�=[2]" & _
                " And ���� IN(Select �������� From �ҺŰ������� Where �ű�ID=[1])" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoom", lng����ID, str�ű�)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select �ű�ID,��������,��ǰ���� From �ҺŰ������� Where �ű�ID=" & lng����ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "GetRoom", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    GetRoom = rsTmp!��������
                    rsTmp!��ǰ���� = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!��������
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRoomVisit(ByVal lng��¼ID As Long, ByVal blnԤԼ������������ As Boolean) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    
    On Error GoTo errH
    
    If blnԤԼ������������ Then
        strSQL = "Select a.Id" & vbNewLine & _
                "From �ٴ������¼ a, �ٴ������¼ b" & vbNewLine & _
                "Where a.��Դid = b.��Դid And a.�Ƿ��ʱ�� = b.�Ƿ��ʱ�� And a.�Ƿ���ſ��� = b.�Ƿ���ſ��� And a.����id = b.����id And" & vbNewLine & _
                "      Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And a.�ϰ�ʱ�� = b.�ϰ�ʱ�� And Nvl(a.�Ƿ񷢲�, 0) = 1 And a.�������� = Trunc(Sysdate) And" & vbNewLine & _
                "      b.Id = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��������ID", lng��¼ID)
        If Not rsTmp.EOF Then
            lng��¼ID = Val(Nvl(rsTmp!id))
        End If
    End If
            
    strSQL = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ٴ������¼ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", lng��¼ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        strSQL = "Select B.���� As �������� From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID=B.ID And A.��¼ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", CLng(rsTmp!id))
        If Not rsTmp.EOF Then GetRoomVisit = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select B.���� As ��������,0 as NUM From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID = B.ID And ��¼ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And ��¼����=1 and ��¼״̬=1 and  ����ʱ�� Between Trunc(Sysdate) And Sysdate And �����¼ID = [2]" & _
                " And ���� IN (Select D.���� As �������� From �ٴ��������Ҽ�¼ C,�������� D Where C.��¼ID=[1] And C.����ID = D.ID )" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", CLng(rsTmp!id), lng��¼ID)
        If Not rsTmp.EOF Then GetRoomVisit = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select * From �ٴ��������Ҽ�¼ Where ��¼ID=" & rsTmp!id
'        strSQL = "Select A.��¼ID,B.���� As ��������,A.��ǰ���� From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID=B.ID And A.��¼ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "GetRoomVisit", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    strRoomIDs = rsTmp!����ID
                    rsTmp!��ǰ���� = 0
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!����ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select ���� From �������� Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, "GetRoomVisit", strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoomVisit = rsRoom!����
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckCanModifyName(ByVal strNO As String, ByVal str����ʱ�� As String, ByVal lng����ID As Long) As Boolean
'����:���Һŵ��Ƿ�����޸�����,������ǹҺ�ʱ���ĵ�,�Ͳ����޸�.
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From ������ü�¼ A" & vbNewLine & _
            "Where A.NO = [1] And A.��¼���� = 4 And A.�Ǽ�ʱ�� = To_Date([2],'YYYY-MM-DD HH24:MI:SS') And A.����id = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�Һ�ʱ����", strNO, str����ʱ��, lng����ID)
    CheckCanModifyName = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetʧԼ��(ByVal bytRegMode As Byte, ByVal varUDID As Variant, ByVal lngԤԼ��Чʱ�� As Long, datThis As Date) As Long
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������ĳһ��.ԤԼʧԼ��
    ' ��� : bytRegMode:0-�ƻ�����ģʽ��1-�����Ű�ģʽ
    '        var��ʶ:�Һ�Ψһ��ʶ��bytRegMode=0�Ǻű�bytRegMode=1 �Ǽ�¼id
    '        datThis-ָ������
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/16 17:17
    '---------------------------------------------------------------------------------------
    Dim strSQL  As String, strWhere As String
    Dim rsTmp   As ADODB.Recordset
    Dim strBegin  As String, strEnd As String
    
    If bytRegMode = 0 Then
        strWhere = "�ű� = [1]"
    Else
        strWhere = "�����¼ID = [1]"
    End If
    
    strSQL = "Select Count(1) As ʧԼ��" & vbNewLine & _
            " From ���˹Һż�¼" & vbNewLine & _
            " Where " & strWhere & " And ��¼���� = 2 And ��¼״̬ = 1 And ����ʱ�� - [2] / 24 / 60 < Sysdate And ����ʱ�� Between to_Date([3],'YYYY-MM-DD') And to_Date([4],'YYYY-MM-DD') - 1/24/60/60"
    strBegin = Format(datThis, "yyyy-MM-dd")
    strEnd = Format(datThis + 1, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ʧԼ����", varUDID, lngԤԼ��Чʱ��, strBegin, strEnd)
    If rsTmp.EOF Then
        zlGetʧԼ�� = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    zlGetʧԼ�� = Val(Nvl(rsTmp!ʧԼ��, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Public Function zlCheckRegistVaild(ByVal objService As zlPublicExpense.clsService, _
                ByVal byt�������� As Byte, ByVal lng����ID As Long, ByVal str���� As String, _
                ByVal datԤԼʱ�� As Date, ByVal blnר�Һ� As Boolean, _
                Optional ByVal str�����¼ID As String, Optional ByVal strԤԼ��ʽ As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : Zl_Fun_���˹Һż�¼_Check
    ' ��� : byt�������ͣ�0-�Һ�;1-ԤԼ;2-ԤԼ����
    '        str�����¼ID�����Ű�ģʽ���������Ű�>=0
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/26 18:43
    '---------------------------------------------------------------------------------------
    Dim rsCheck As ADODB.Recordset, strSQL As String
    Dim strResult As String
    On Error GoTo Errhand

    If objService Is Nothing Then
        Call ShowMsgbox("����ӿ�(zlPublicExpense.clsService)δ���ã�")
        Exit Function
    End If
    If byt�������� = 1 Then
        If objService.zlPatisvr_GetPatiBlackInfo(lng����ID, "ԤԼ", strԤԼ��ʽ) = False Then Exit Function
    End If
    
    strSQL = "Select Zl_Fun_���˹Һż�¼_Check_S([1],[2],[3],[4],[5],[6]) As ����� From Dual"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "zlCheckRegistVaild", byt��������, lng����ID, str����, str�����¼ID, datԤԼʱ��, IIf(blnר�Һ�, 1, 0))
    If Not rsCheck.EOF Then
        strResult = Nvl(rsCheck!�����)
        If Val(Mid(strResult, 1, 1)) <> 0 Then
            MsgBox Mid(strResult, 3), vbInformation, gstrSysName
            Exit Function
        End If
    Else
        MsgBox "��Ч�Լ��ʧ��,�޷�������", vbInformation, gstrSysName
        Exit Function
    End If
    zlCheckRegistVaild = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Check����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
'����:�жϲ����Ƿ��ٴε�����ͬ�ٴ����ʵ��ٴ����ҡ��Һ�
'     �����ҹ��ŵ�,��ס��Ժ��,���ﲻ��ȷ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As �����־ From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������", lng����ID, lngִ�в���ID)
    Check���� = Val(Nvl(rsTmp!�����־)) = 1
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCisNewVisitRec(ByVal objService As zlPublicExpense.clsService, ByVal strNO As String, _
                Optional ByRef str����ʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �Һ���ɺ�֪ͨ�ٴ�
    ' ��� :
    ' ���� : str����ʱ��-����ʱ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand

    strSQL = "Select ��¼����,�ű�,����,ԤԼ,����id,�����,����,�Ա�,����,�ѱ�,����,����,����,ִ�в���ID," & _
                    "ִ����,����ʱ��,����ģʽ,�����¼ID " & vbNewLine & _
            "From ���˹Һż�¼ Where No  = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���¾���Ǽ�", strNO)
    If rsTmp.EOF Then Exit Function
    Set cllVisit = New Collection
    With rsTmp
        cllVisit.Add strNO, "�Һŵ���"
        cllVisit.Add Val(Nvl(!��¼����)), "��������"
        cllVisit.Add Val(Nvl(!ԤԼ)), "ԤԼ��־"
        cllVisit.Add Nvl(!�ű�), "�ű�"
        cllVisit.Add Val(Nvl(!�����¼ID)), "�����¼ID"
        cllVisit.Add Nvl(!����), "����"
        cllVisit.Add Val(Nvl(!����ID)), "����ID"
        cllVisit.Add Nvl(!�����), "�����"
        cllVisit.Add Nvl(!����), "����"
        cllVisit.Add Nvl(!�Ա�), "�Ա�"
        cllVisit.Add Nvl(!����), "����"
        cllVisit.Add Val(Nvl(!����)), "����"
        cllVisit.Add Nvl(!�ѱ�), "�ѱ�"
        cllVisit.Add Val(Nvl(!����)), "����"
        cllVisit.Add Nvl(!����), "����"
        cllVisit.Add Val(Nvl(!ִ�в���ID)), "ִ�в���ID"
        cllVisit.Add Nvl(!ִ����), "ִ����"
        cllVisit.Add Nvl(!����ʱ��), "����ʱ��"
        cllVisit.Add Val(Nvl(!����ģʽ)), "����ģʽ"
        str����ʱ�� = Nvl(!����ʱ��)
    End With
    zlCisNewVisitRec = objService.zlCISSvr_NewOutPatiVisitRec(cllVisit)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVisitRoom(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal strִ���� As String, ByVal strRoom As String, _
                Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �����ٴ���������
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("�Һŵ���", strNO)
    cllVisit.Add Array("����", strRoom)
    cllVisit.Add Array("ִ����", strִ����)
    
    zlCisUpdateVisitRoom = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVisitState(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal intִ��״̬ As Integer, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �����ٴ�����״̬
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("�Һŵ���", strNO)
    cllVisit.Add Array("ִ��״̬", intִ��״̬)
    
    zlCisUpdateVisitState = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisUpdateVistPatiBase(ByVal objService As zlPublicExpense.clsService, _
                ByVal blnNewPati As Boolean, ByVal strNO As String, ByVal lng����ID As Long, ByVal str����� As String, _
                ByVal str���� As String, ByVal str�Ա� As String, ByVal str���� As String, _
                ByVal str�ѱ� As String, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : �����ٴ�������Ϣ�в��˼�¼��Ϣ
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand

    Set cllVisit = New Collection
    cllVisit.Add Array("�Һŵ���", strNO)
    If blnNewPati Then
        cllVisit.Add Array("����ID", lng����ID)
        cllVisit.Add Array("����", str����)
        cllVisit.Add Array("�Ա�", str�Ա�)
        cllVisit.Add Array("����", str����)
    End If
    cllVisit.Add Array("�����", str�����)
    cllVisit.Add Array("�ѱ�", str�ѱ�)
    
    zlCisUpdateVistPatiBase = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisDoneVisit(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal strִ���� As String, ByVal strRoom As String, ByVal str���ʱ�� As String, _
                ByVal strժҪ As String, ByVal bln��ʿִ�� As Boolean, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ɾ���
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand
    
    Set cllVisit = New Collection
    cllVisit.Add Array("�Һŵ���", strNO)
    cllVisit.Add Array("ִ��״̬", 1)
    cllVisit.Add Array("����", strRoom)
    cllVisit.Add Array("���ʱ��", str���ʱ��)
    If strִ���� <> "" Then
        cllVisit.Add Array("ִ����", strִ����)
    End If
    If strժҪ <> "" Then
        cllVisit.Add Array("ժҪ", strժҪ)
    End If
    
    zlCisDoneVisit = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCisCancelVisit(ByVal objService As zlPublicExpense.clsService, _
                ByVal strNO As String, ByVal bln��ʿִ�� As Boolean, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ȡ������
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/20 17:02
    '---------------------------------------------------------------------------------------
    Dim cllVisit As Collection
    On Error GoTo Errhand
    
    Set cllVisit = New Collection
    cllVisit.Add Array("�Һŵ���", strNO)
    cllVisit.Add Array("ִ��״̬", IIf(bln��ʿִ��, 0, 2))
    cllVisit.Add Array("���ʱ��", "")
    
    zlCisCancelVisit = objService.zlCISSvr_UpdateOutPatiVisitRec(cllVisit, True, strErrMsg)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddAdressInfo(ByRef cllPati As Collection, padd��סַ As Object, padd���ڵ�ַ As Object) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�ṹ����ַ��Ϣ
    ' ��� : padd��סַ:�ֵ�ַ�ؼ�
    '        padd���ڵ�ַ:���ڵ�ַ�ؼ�
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/30 21:59
    '---------------------------------------------------------------------------------------
    Dim cllTemp As Collection, cllSubTemp As Collection
    On Error GoTo Errhand

    If cllPati Is Nothing Then Set cllPati = New Collection
    Set cllTemp = New Collection
    
    If padd��סַ.Value <> "" Then
        If zlGetAdressCol(cllSubTemp, 1, 3, padd��סַ.valueʡ, padd��סַ.value��, padd��סַ.value����, padd��סַ.value����, _
            padd��סַ.value��ϸ��ַ, padd��סַ.Code) = False Then Exit Function
    Else
        If zlGetAdressCol(cllSubTemp, 2, 3) = False Then Exit Function
    End If
    cllTemp.Add cllSubTemp, "��ͥ��ַ"
    
    If padd���ڵ�ַ.Value <> "" Then
        If zlGetAdressCol(cllSubTemp, 1, 4, padd���ڵ�ַ.valueʡ, padd���ڵ�ַ.value��, padd���ڵ�ַ.value����, padd���ڵ�ַ.value����, _
            padd���ڵ�ַ.value��ϸ��ַ, padd���ڵ�ַ.Code) = False Then Exit Function
    Else
        If zlGetAdressCol(cllSubTemp, 2, 4) = False Then Exit Function
    End If
    cllTemp.Add cllSubTemp, "���ڵ�ַ"
    
    If cllTemp.Count > 0 Then cllPati.Add cllTemp, "��ַ��Ϣ"
    AddAdressInfo = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiCol(ByRef cllPati_Out As Collection, _
                ByVal lng����ID As Long, ByVal str���� As String, ByVal str�Ա� As String, _
                ByVal str���� As String, ByVal str�������� As String, ByVal str���֤�� As String, ByVal str����� As String, _
                ByVal str�ѱ� As String, ByVal strҽ�Ƹ��ʽ���� As String, ByVal str���� As String, _
                ByVal str���� As String, ByVal str����״�� As String, ByVal strְҵ As String, _
                ByVal str��� As String, ByVal str������λ As String, ByVal str��λ�绰 As String, _
                ByVal lng��ͬ��λid As Long, ByVal str��λ�ʱ� As String, _
                ByVal str��ͥ��ַ As String, ByVal str��ͥ�绰 As String, _
                ByVal str��ͥ��ַ�ʱ� As String, ByVal str���� As String, ByVal str�����ص� As String, _
                ByVal str���ڵ�ַ As String, ByVal str���ڵ�ַ�ʱ� As String, ByVal str��ϵ������ As String, _
                ByVal str��ϵ�����֤�� As String, ByVal str��ϵ�˵绰 As String, ByVal str��ϵ�˹�ϵ As String, _
                ByVal str�໤�� As String, ByVal str�ֻ��� As String, ByVal strҽ���� As String, ByVal int���� As Integer, _
                ByVal str�Ǽ�ʱ�� As String, Optional ByVal lng������� As Long, Optional ByVal str�������� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������Ϣ����
    ' ��� : ���˻�����Ϣ
    ' ���� : ������Ϣ����
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    On Error GoTo Errhand
    Set cllPati_Out = New Collection
    
    cllPati_Out.Add Array("����ID", lng����ID), "����ID"
    cllPati_Out.Add Array("����", str����), "����"
    cllPati_Out.Add Array("�Ա�", str�Ա�), "�Ա�"
    cllPati_Out.Add Array("����", str����), "����"
    cllPati_Out.Add Array("��������", str��������), "��������"
    cllPati_Out.Add Array("���֤��", str���֤��), "���֤��"
    cllPati_Out.Add Array("�����", str�����), "�����"
    cllPati_Out.Add Array("�ѱ�", str�ѱ�), "�ѱ�"
    cllPati_Out.Add Array("ҽ�Ƹ��ʽ����", strҽ�Ƹ��ʽ����), "ҽ�Ƹ��ʽ����"
    cllPati_Out.Add Array("����", str����), "����"
    cllPati_Out.Add Array("����", str����), "����"
    cllPati_Out.Add Array("����״��", str����״��), "����״��"
    cllPati_Out.Add Array("ְҵ", strְҵ), "ְҵ"
    cllPati_Out.Add Array("���", str���), "���"
    cllPati_Out.Add Array("������λ", str������λ), "������λ"
    cllPati_Out.Add Array("��λ�绰", str��λ�绰), "��λ�绰"
    cllPati_Out.Add Array("��ͬ��λID", lng��ͬ��λid), "��ͬ��λID"
    cllPati_Out.Add Array("��λ�ʱ�", str��λ�ʱ�), "��λ�ʱ�"
    cllPati_Out.Add Array("��ͥ��ַ", str��ͥ��ַ), "��ͥ��ַ"
    cllPati_Out.Add Array("��ͥ�绰", str��ͥ�绰), "��ͥ�绰"
    cllPati_Out.Add Array("��ͥ��ַ�ʱ�", str��ͥ��ַ�ʱ�), "��ͥ��ַ�ʱ�"
    cllPati_Out.Add Array("����", str����), "����"
    cllPati_Out.Add Array("�����ص�", str�����ص�), "�����ص�"
    cllPati_Out.Add Array("���ڵ�ַ", str���ڵ�ַ), "���ڵ�ַ"
    cllPati_Out.Add Array("���ڵ�ַ�ʱ�", str���ڵ�ַ�ʱ�), "���ڵ�ַ�ʱ�"
    cllPati_Out.Add Array("�໤��", str�໤��), "�໤��"
    cllPati_Out.Add Array("�ֻ���", str�ֻ���), "�ֻ���"
    cllPati_Out.Add Array("ҽ����", strҽ����), "ҽ����"
    cllPati_Out.Add Array("����", int����), "����"
    cllPati_Out.Add Array("�Ǽ�ʱ��", str�Ǽ�ʱ��), "�Ǽ�ʱ��"
    cllPati_Out.Add Array("����Ա����", UserInfo.����), "����Ա����"
    cllPati_Out.Add Array("����Ա���", UserInfo.���), "����Ա���"
    If str��ϵ������ <> "" Then
        Set cllTmp = New Collection
        cllTmp.Add str��ϵ������, "��ϵ������"
        cllTmp.Add str��ϵ�����֤��, "��ϵ�����֤��"
        cllTmp.Add str��ϵ�˵绰, "��ϵ�˵绰"
        cllTmp.Add str��ϵ�˹�ϵ, "��ϵ�˹�ϵ"
        cllPati_Out.Add cllTmp, "��ϵ��"
    End If
    
    If str�������� <> "" Then
        Set cllTmp = New Collection
        cllTmp.Add lng�������, "�������"
        cllTmp.Add str��������, "��������"
        cllTmp.Add 1, "������������"
        cllPati_Out.Add cllTmp, "������Ϣ"
    End If
    
    zlGetPatiCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiMecCol(ByRef cllPati_Out As Collection, _
                ByVal lng����ID As Long, ByVal str���� As String, ByVal str�Ա� As String, _
                ByVal str���� As String, ByVal str�������� As String, ByVal str���֤�� As String, _
                ByVal str����� As String, ByVal str�ѱ� As String, ByVal strҽ�Ƹ��ʽ���� As String, _
                ByVal str���� As String, ByVal str���� As String, ByVal str����״�� As String, _
                ByVal strְҵ As String, ByVal str��� As String, _
                ByVal str������λ As String, ByVal str��λ�绰 As String, _
                ByVal lng��ͬ��λid As Long, ByVal str��λ�ʱ� As String, _
                ByVal str��ͥ��ַ As String, ByVal str��ͥ�绰 As String, ByVal str��ͥ��ַ�ʱ� As String, _
                ByVal str���ڵ�ַ As String, ByVal str���ڵ�ַ�ʱ� As String, _
                ByVal str�໤�� As String, ByVal str�ֻ��� As String, ByVal strҽ���� As String, ByVal int���� As Integer, _
                ByVal str�Ǽ�ʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ������Ϣ����
    ' ��� : ���˻�����Ϣ
    ' ���� : ������Ϣ����
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    On Error GoTo Errhand
    Set cllPati_Out = New Collection
    
    cllPati_Out.Add Array("����ID", lng����ID)
    cllPati_Out.Add Array("����", str����)
    cllPati_Out.Add Array("�Ա�", str�Ա�)
    cllPati_Out.Add Array("����", str����)
    cllPati_Out.Add Array("��������", str��������)
    cllPati_Out.Add Array("���֤��", str���֤��)
    cllPati_Out.Add Array("�����", str�����)
    cllPati_Out.Add Array("�ѱ�", str�ѱ�)
    cllPati_Out.Add Array("ҽ�Ƹ��ʽ����", strҽ�Ƹ��ʽ����)
    cllPati_Out.Add Array("����", str����)
    cllPati_Out.Add Array("����", str����)
    cllPati_Out.Add Array("����״��", str����״��)
    cllPati_Out.Add Array("ְҵ", strְҵ)
    cllPati_Out.Add Array("���", str���)
    cllPati_Out.Add Array("������λ", str������λ)
    cllPati_Out.Add Array("��λ�绰", str��λ�绰)
    cllPati_Out.Add Array("��ͬ��λID", lng��ͬ��λid)
    cllPati_Out.Add Array("��λ�ʱ�", str��λ�ʱ�)
    cllPati_Out.Add Array("��ͥ��ַ", str��ͥ��ַ)
    cllPati_Out.Add Array("��ͥ�绰", str��ͥ�绰)
    cllPati_Out.Add Array("��ͥ��ַ�ʱ�", str��ͥ��ַ�ʱ�)
    cllPati_Out.Add Array("���ڵ�ַ", str���ڵ�ַ)
    cllPati_Out.Add Array("���ڵ�ַ�ʱ�", str���ڵ�ַ�ʱ�)
    cllPati_Out.Add Array("�໤��", str�໤��)
    cllPati_Out.Add Array("�ֻ���", str�ֻ���)
    cllPati_Out.Add Array("ҽ����", strҽ����)
    cllPati_Out.Add Array("����", int����)
    cllPati_Out.Add Array("�Ǽ�ʱ��", str�Ǽ�ʱ��)
    cllPati_Out.Add Array("����Ա����", UserInfo.����)
    cllPati_Out.Add Array("����Ա���", UserInfo.���)
    
    zlGetPatiMecCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetAdressCol(ByRef cllAdress_Out As Collection, ByVal byt�������� As Byte, _
                ByVal byt��ַ��� As Byte, Optional ByVal str��ַ_ʡ As String, Optional ByVal str��ַ_�� As String, _
                Optional ByVal str��ַ_�� As String, Optional ByVal str��ַ_�� As String, Optional ByVal str��ַ_���� As String, _
                Optional ByVal str�������� As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�ṹ����ַ��Ϣ����
    ' ��� : �ṹ����ַ��Ϣ
    ' ���� : �ṹ����ַ����
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/30 19:04
    '---------------------------------------------------------------------------------------
    On Error GoTo Errhand
    Set cllAdress_Out = New Collection
    
    cllAdress_Out.Add byt��������, "��������"
    cllAdress_Out.Add byt��ַ���, "��ַ���"
    cllAdress_Out.Add str��ַ_ʡ, "��ַ_ʡ"
    cllAdress_Out.Add str��ַ_��, "��ַ_��"
    cllAdress_Out.Add str��ַ_��, "��ַ_��"
    cllAdress_Out.Add str��ַ_��, "��ַ_��"
    cllAdress_Out.Add str��ַ_����, "��ַ_����"
    cllAdress_Out.Add str��������, "��������"

    zlGetAdressCol = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetȱʡ�ѱ�() As String
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡȱʡ�ѱ�
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/10/31 14:11
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    strSQL = "Select ����  From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȱʡ�ѱ�")
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function

Public Function zlExcMedcCardService(ByVal objService As zlPublicExpense.clsService, _
                ByVal int�������� As Integer, ByVal lng����ID As Long, ByVal lng�����id As Long, _
                ByVal strԭ���� As String, ByVal strҽ�ƿ��� As String, ByVal str���� As String, _
                ByVal str�䶯ԭ�� As String, Optional ByVal strIC���� As String, Optional ByVal str��ʧ��ʽ As String, _
                Optional ByVal str��ά�� As String, Optional ByVal str��ֹʹ��ʱ�� As String, _
                Optional ByVal strNO As String, Optional ByVal dbl���� As Double, _
                Optional ByVal str����ʱ�� As String, Optional ByRef strErrMsg As String) As Boolean
    Dim cllCard As Collection, cllVisit As Collection

    On Error GoTo errHandle
    If objService Is Nothing Then Exit Function
    str���� = zlCommFun.zlStringEncode(str����)
    
    Set cllCard = New Collection
    cllCard.Add Array("��������", int��������)
    cllCard.Add Array("����id", lng����ID)
    cllCard.Add Array("�����ID", lng�����id)
    cllCard.Add Array("ԭ����", strԭ����)
    cllCard.Add Array("ҽ�ƿ���", strҽ�ƿ���)
    cllCard.Add Array("�䶯ԭ��", str�䶯ԭ��)
    cllCard.Add Array("����", str����)
    cllCard.Add Array("IC����", strIC����)
    cllCard.Add Array("��ʧ��ʽ", str��ʧ��ʽ)
    cllCard.Add Array("��ά��", str��ά��)
    cllCard.Add Array("��ֹʹ��ʱ��", str��ֹʹ��ʱ��)
    cllCard.Add Array("����ʱ��", str����ʱ��)
    cllCard.Add Array("����Ա����", UserInfo.����)
    cllCard.Add Array("����Ա���", UserInfo.���)
    cllCard.Add Array("���ݺ�", strNO)
    cllCard.Add Array("����", dbl����)

    If objService.zlPatisvr_SaveMedcCard(cllCard, strErrMsg) = False Then Exit Function
    zlExcMedcCardService = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlPatiUpdVisitService(ByVal objService As zlPublicExpense.clsService, ByVal lng����ID As Long, _
                ByVal strNode As String, Optional ByVal int����״̬ As Integer, Optional ByVal str����ʱ�� As String, _
                Optional ByVal str����� As String, Optional ByVal str�������� As String, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���²��˾�����Ϣ
    ' ��� : strNode-��Ҫ���µĽڵ�
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/5 15:40
    '---------------------------------------------------------------------------------------
    Dim cllPati As Collection
    
    On Error GoTo Errhand
    If objService Is Nothing Then Exit Function
    Set cllPati = New Collection
    cllPati.Add lng����ID, "����ID"
    If InStr("," & strNode & ",", "����״̬") > 0 Then
        cllPati.Add int����״̬, "����״̬"
    End If
    If InStr("," & strNode & ",", "�����") > 0 Then
        cllPati.Add str�����, "�����"
    End If
    If InStr("," & strNode & ",", "����") > 0 Then
        cllPati.Add str��������, "��������"
    End If
    If InStr("," & strNode & ",", "����ʱ��") > 0 Then
        cllPati.Add str����ʱ��, "����ʱ��"
    End If
    
    If objService.ZlPatiSvr_UpdateOutPatiState(cllPati, strErrMsg) = False Then Exit Function
    
    zlPatiUpdVisitService = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function LoadErrUnBalanceInfo(ByRef rsBillAdvance As ADODB.Recordset, _
                objInterCard As clsInterFaceCard, ByVal objPayInfor As clsPayInfos, _
                ByVal strԭ����IDs As String, ByVal str����IDs As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ�쳣����δ������
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/9 10:23
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objCard As Card
    On Error GoTo Errhand
    If objInterCard Is Nothing Then Exit Function
    
    strSQL = "Select Mod(a.��¼����, 10) As ��¼����, b.����,Nvl(d.����, a.���㷽ʽ) as ����, a.���㷽ʽ, Sum(a.��Ԥ��) As ���, a.�������, Nvl(a.�����id, a.���㿨���) As �����id," & vbNewLine & _
            "       Decode(a.���㿨���, Null, 0, 1) As ���ѿ�, a.����, e.���ѿ�id, a.��������id, a.������ˮ��, a.����˵��" & vbNewLine & _
            "From ����Ԥ����¼ a, ���㷽ʽ b, ���ѿ����Ŀ¼ d, ���˿������¼ e" & vbNewLine & _
            "Where a.����id In (Select /* +cardinality(M,10) */" & vbNewLine & _
            "                  m.Column_Value" & vbNewLine & _
            "                 From Table(f_Str2list([1])) m) And a.���㷽ʽ Is Not Null And Nvl(a.У�Ա�־, 0) <> 1 And a.���㷽ʽ = b.����(+) And" & vbNewLine & _
            "      a.Id = e.����id(+) And a.���㿨��� = d.���(+) And a.���㷽ʽ = d.���㷽ʽ(+) And " & vbNewLine & _
            "      (a.��¼���� = 11 Or (Nvl(a.�����id, Nvl(a.���㿨���, 0)) = 0 And b.���� In (1, 2)) Or" & vbNewLine & _
            "      (Nvl(a.�����id, Nvl(a.���㿨���, 0)) > 0 And" & vbNewLine & _
            "      (Nvl(a.�����id, 0), Nvl(a.���㿨���, 0)) Not In" & vbNewLine & _
            "      (Select Nvl(�����id, 0), Nvl(���㿨���, 0)" & vbNewLine & _
            "         From ����Ԥ����¼" & vbNewLine & _
            "         Where ����id In (Select /* +cardinality(M,10) */" & vbNewLine & _
            "                         m.Column_Value" & vbNewLine & _
            "                        From Table(f_Str2list([2])) m) And ��¼״̬ = 2 And ���㷽ʽ Is Not Null)))" & vbNewLine & _
            "Group By a.��¼����, b.����, a.���㷽ʽ,d.����, a.�������, Nvl(a.�����id, a.���㿨���), Decode(a.���㿨���, Null, 0, 1), a.����, e.���ѿ�id, a.��������id," & vbNewLine & _
            "         a.������ˮ��, a.����˵��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡδ������", strԭ����IDs, str����IDs)
    
    If rsBillAdvance Is Nothing Then
        If InitBalanceData(rsBillAdvance) = False Then Exit Function
    End If
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            rsBillAdvance.AddNew
            rsBillAdvance!��¼���� = Val(Nvl(rsTmp!��¼����))
            rsBillAdvance!�������� = Val(Nvl(rsTmp!����))
            rsBillAdvance!�����ID = Val(Nvl(rsTmp!�����ID))
            rsBillAdvance!���ѿ� = Val(Nvl(rsTmp!���ѿ�))
            If Val(Nvl(rsTmp!�����ID)) > 0 And Val(Nvl(rsTmp!���ѿ�)) = 0 Then
                If objInterCard.GetCard(Val(Nvl(rsTmp!�����ID)), False, objCard) Then
                    If Nvl(rsBillAdvance!���㷽ʽ) = objCard.���㷽ʽ Then
                        rsBillAdvance!֧����ʽ = objCard.����
                    Else
                        rsBillAdvance!֧����ʽ = Nvl(rsTmp!����)
                    End If
                Else
                    rsBillAdvance!֧����ʽ = Nvl(rsTmp!����)
                End If
            Else
                 rsBillAdvance!֧����ʽ = Nvl(rsTmp!����)
            End If
            rsBillAdvance!���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
            rsBillAdvance!��� = Val(Nvl(rsTmp!���))
            rsBillAdvance!����ID = IIf(Val(Nvl(rsTmp!��¼����)) = 5, objPayInfor.Card_����ID, objPayInfor.Reg_����ID)
            rsBillAdvance!���ѿ�ID = Val(Nvl(rsTmp!���ѿ�ID))
            rsBillAdvance!������� = Nvl(rsTmp!�������)
            rsBillAdvance!���� = Nvl(rsTmp!����)
            rsBillAdvance!��������ID = Val(Nvl(rsTmp!��������ID))
            rsBillAdvance!������ˮ�� = Nvl(rsTmp!������ˮ��)
            rsBillAdvance!����˵�� = Nvl(rsTmp!����˵��)
            rsBillAdvance!У�Ա�־ = 0
            rsBillAdvance!�̶� = 0
            rsBillAdvance.Update
            rsTmp.MoveNext
        Loop
    End If
    LoadErrUnBalanceInfo = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

    
Public Function zlExcStationErrReceive(ByVal objService As zlPublicExpense.clsService, ByVal objExseSvr As zlPublicExpense.clsExpenceSvr, _
                ByVal strNO As String, ByRef lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ����ҽ��վԤԼ�����쳣����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/1/3 16:56
    '---------------------------------------------------------------------------------------
    Dim cllPro As Collection, cllSwapOther As Collection
    Dim blnTrans As Boolean
    Dim intͬ��״̬ As Integer
    Dim str������Ϣ As String, strErrMsg As String
    
    On Error GoTo Errhand
    lng����ID = 0
    If objService.zlCISSvr_GetErrBillInfo(strNO, intͬ��״̬, lng����ID, , cllPro, cllSwapOther) = False Then Exit Function
    If intͬ��״̬ = 0 Then zlExcStationErrReceive = True: Exit Function
    If intͬ��״̬ <> 2 Then
        If intͬ��״̬ <> -1 Then lng����ID = 0
        If objService.zlCISSvr_delErrBillInfo(strNO, True) = False Then Exit Function
    Else
        gcnOracle.BeginTrans: blnTrans = True
            zlExecuteProcedureArrAy cllPro, "zlExcStationErrReceive", True, True
            If objService.zlCISSvr_delErrBillInfo(strNO, False, strErrMsg) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            zlExecuteProcedureArrAy cllSwapOther, "zlExcStationErrReceive", True, True
            On Error GoTo Errhand
        gcnOracle.CommitTrans: blnTrans = False
    End If
    zlExcStationErrReceive = True
    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlInitPati(ByVal objInterCard As clsInterFaceCard, ByRef rsPatiInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������Ϣ��
    '����:������Ϣ��
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPatiInfor = New ADODB.Recordset
    If objInterCard Is Nothing Then
        Set objInterCard = New clsInterFaceCard
        Call objInterCard.Init(Nothing, glngSys, glngModul, gcnOracle, gstrDBUser)
    End If
    With rsPatiInfor
        If .State = adStateOpen Then .Close
        '����ID,����,�Ա�,����,��������,�����ص�,���֤��,����֤��,���,ְҵ,��ͥ��ַ,��ͥ�绰,��ͥ�ʱ�,
        '������λ,��λ�ʱ�,ҽ����,ҽ�Ƹ��ʽ,�ѱ�,����,����,����״��,����
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, objInterCard.GetPatiInforMaxLen("����"), adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .Fields.Append "�����ص�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���֤��", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "����֤��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "ְҵ", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "��ͥ��ַ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ͥ�绰", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ͥ�ʱ�", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "��ͬ��λID", adDouble, 18, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��λ�绰", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��λ�ʱ�", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "ҽ����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ҽ�Ƹ��ʽ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ѱ�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����״��", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    zlInitPati = True
End Function

Public Function InitRegist(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnOracle As ADODB.Connection, ByVal strDbUser As String, _
                Optional objRegist As zlPublicExpense.clsRegist, _
                Optional objExseSvr As zlPublicExpense.clsExpenceSvr, _
                Optional objService As zlPublicExpense.clsService) As Boolean
    '��ʼ���Һ�
    Dim strDept As String
    On Error GoTo errH:
    Set objRegist = New clsRegist
    If objRegist.zlInitCommon(lngSys, cnOracle, strDbUser) = False Then Exit Function
    
    Set objExseSvr = New clsExpenceSvr
    If objExseSvr.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    Set objService = New zlPublicExpense.clsService
    If objService.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    InitRegist = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiCollect(ByRef cllPati_Out As Collection, ByVal objPati As clsPatientInfo, ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ����
    '���:
    '����:objPati-������Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-29 11:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set cllPati_Out = Nothing
    
    If Not objPati Is Nothing Then
        Set cllPati_Out = New Collection
        With objPati
            cllPati_Out.Add .����ID, "_����ID"
            cllPati_Out.Add 0, "_��ҳID"
            cllPati_Out.Add .����, "_����"
            cllPati_Out.Add .�Ա�, "_�Ա�"
            cllPati_Out.Add .����, "_����"
            cllPati_Out.Add .�����, "_�����"
            cllPati_Out.Add 0, "_סԺ��"
            cllPati_Out.Add intInsure, "_����"
        End With
    End If
    If cllPati_Out Is Nothing Then
        Set cllPati_Out = New Collection
        cllPati_Out.Add 0, "_����ID"
        cllPati_Out.Add 0, "_��ҳID"
        cllPati_Out.Add "", "_����"
        cllPati_Out.Add "", "_�Ա�"
        cllPati_Out.Add "", "_����"
        cllPati_Out.Add "", "_�����"
        cllPati_Out.Add "", "_סԺ��"
        cllPati_Out.Add 0, "_����"
    End If
    
    GetPatiCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
