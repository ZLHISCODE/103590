Attribute VB_Name = "mdlRegEvent"
Option Explicit 'Ҫ���������
Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    'support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��
    
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportҽ���ӿڴ�ӡƱ�� = 46    'HIS��ֻ��Ʊ�ݺŵ�������ӡ��ҽ���ӿ�(����)�д�ӡ
    support�����Һ� = 62            '�ڹҺ�ʱ���Ƿ������������Һ�(������ɿ����Ž���)
    support�ҺŲ���ȡ������ = 81    '�ڹҺ�ʱ����ʹ��ҽ����ȡ������
    support�Һż����Ŀ = 86
End Enum
 Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public gobjSquare As SquareCard  '�����㲿��  42301
Public gobjTax As Object '˰�ش�ӡ�ӿڶ���
Public gblnTax As Boolean '�����Ƿ�ʹ��˰�ش�ӡ
Public gstrTax As String
Public gobjRegist As Object
Public gobjPlugIn As Object, gblnPlugin As Boolean
Public gobjPublicExpense As Object
Public gintPriceGradeStartType As Integer
Public gstrPriceGrade As String

'Ʊ�ݿ���
Public gobjBillPrint As Object '������Ʊ�ݴ�ӡ����
Public gblnBillPrint As Boolean '������Ʊ�ݴ�ӡ�����Ƿ����


Public Function GetMaxLen() As Byte
'���ܣ���ȡ�Һ���Ŀ�ű����󳤶�
'˵������ȡ�Һ���Ŀ�������󳤶�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxLen = 5
    strSQL = "Select Nvl(Max(Length(����)),5) as ���� From �ҺŰ���"
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    
    If Not rsTmp.BOF Then GetMaxLen = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCashMoney(ByVal strNO As String) As Currency
'���ܣ�ҽ����֧���˸����ʻ�ʱ,�����ʻ����ֽ�,��ȡ�ֽ��˿���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select -1*A.��Ԥ�� as �ֽ� From ����Ԥ����¼ A,������ü�¼ B,���㷽ʽ C " & _
            " Where A.����ID=B.����ID And A.���㷽ʽ=C.���� And B.NO=[1] " & _
            " And A.��¼����=4 And A.��¼״̬=2 And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.BOF Then GetCashMoney = rsTmp!�ֽ�
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal lng��Ŀid As Long, ByVal intִ�п������� As Integer) As Long
'���ܣ���ȡ�ҺŸ�����Ŀ(������,���￨��)���շ���Ŀ��ִ�п���
'������
'���أ����������,��ʾ�Һſ���(ҽ�����ڿ���)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Get�շ�ִ�п���ID = UserInfo.����ID
    
    Select Case intִ�п�������
        Case 0 '0-����ȷ����
        Case 1 '1-�������ڿ���
            Get�շ�ִ�п���ID = 0
        Case 2 '2-�������ڲ���
            Get�շ�ִ�п���ID = 0
        Case 3 '3-����Ա����
        Case 4 '4-ָ������
            strSQL = "Select ִ�п���ID From �շ�ִ�п��� Where �շ�ϸĿID=[1] And Nvl(������Դ,1)=1 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng��Ŀid)
            
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
        Case 6 '�����˿���
    End Select
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadRegistPrice(ByVal lng��Ŀid As Long, ByVal bln���� As Boolean, ByVal bln���￨ As Boolean, _
    Optional str�ѱ� As String, Optional rsItems As ADODB.Recordset, Optional rsIncomes As ADODB.Recordset, _
    Optional lng����ID As Long, Optional int���� As Integer, Optional str�ű� As String, Optional bytMode As Integer, _
    Optional lng�Һſ���ID As Long = 0, Optional strPriceGrade As String, Optional strDate As String, _
    Optional ByVal lng����ϸĿID As Long) As Long
'���ܣ���ȡָ���Һ���Ŀ��Ӧ�ķ�����Ϣ����¼����
'������lng��ĿID=��ʾ�Ƿ��ȡ�Һŷ���(Ҫ���ĹҺ���ĿID)
'      bln����=��ʾ�Ƿ��ȡ����������(���ܽ���ȡ������)
'      bln���￨=��ʾ�Ƿ��ȡ���￨����(��Һŷѻ�����һ����ȡ)
'      str�ѱ�=�Һŷѱ�
'      rsItems(Out)=�����Һ���Ŀ��������Ŀ,������New��ʽ����
'      rsInComes(Out)=����������Ŀ���������,������New��ʽ����
'      strPriceGrade=�շ���Ŀ�ļ۸�ȼ�
'      lng����ϸĿID= �����Զ��忨��ʱ����
'���أ���ȡ����Ŀ����,ͬʱrsItems,rsInCome=Nothing
'˵������������Ϊ1,����趨���δ���,��Ϊ�̶�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngԭ��ID As Long
    Dim rsFeeTmp As ADODB.Recordset
    Dim strFee As String
    Dim str������ĿID As String
    Dim strWherePriceGrade As String
    Dim strDateCondition As String
    
    Set rsItems = Nothing
    Set rsIncomes = Nothing
    
    If strDate <> "" Then
        strDateCondition = " [5] "
    Else
        strDateCondition = " Sysdate "
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [4]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [4]" & vbNewLine & _
            "                                   And " & strDateCondition & " Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    
    '��ȡ�Һ���Ŀ��������Ŀ�ķ���
    If lng��Ŀid <> 0 Then
        strSQL = _
            "Select 1 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[1]" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        strSQL = strSQL & " Union ALL " & _
            "Select 2 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=[1]" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '��ȡ���������Ѷ�Ӧ�ķ���
    If bln���� Then
        strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
            "Select 3 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
            " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,A.ִ�п��� as ִ�п�������" & _
            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D" & _
            " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.�շ�ϸĿID And D.�ض���Ŀ='������'" & _
            " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '��ȡ���￨��Ӧ�ķ���(��֧�����ö��������Ŀ,Ϊ�˱��ֺ;��￨������һ��)
    '���������޼�Ϊ��ʱ,���տ���
    If bln���￨ Then
        If lng����ϸĿID = 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                "Select 4 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,A.ִ�п��� as ִ�п�������" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D" & _
                " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.�շ�ϸĿID And D.�ض���Ŀ=[2] And (A.�Ƿ���=1 And Nvl(B.ԭ��,0)<>0 or A.�Ƿ���=0)" & _
                " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum=1" & vbNewLine & _
                strWherePriceGrade
        Else
            strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                "Select 4 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,A.ִ�п��� as ִ�п�������" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
                " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[2] And (A.�Ƿ���=1 And Nvl(B.ԭ��,0)<>0 or A.�Ƿ���=0)" & _
                " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum=1" & vbNewLine & _
                strWherePriceGrade
        End If
    End If
    
    If bytMode <> 1 And bytMode <> 10 And Not (lng��Ŀid = 0 And bln���� = True) Then
        strFee = "Select zl_Fun_CustomRegExpenses([1],[2],[3]) As ���ӷ� From Dual"
        Set rsFeeTmp = zlDatabase.OpenSQLRecord(strFee, "zl_Fun_CustomRegExpenses", lng����ID, int����, str�ű�)
        If Not rsFeeTmp.EOF Then
            str������ĿID = Nvl(rsFeeTmp!���ӷ�)
        End If
        
        If str������ĿID <> "" Then
            If strSQL = "" Then
                strSQL = " " & _
                    "Select /*+cardinality(D,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                    " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,Table(f_str2list([3])) D " & _
                    " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
                    " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
            Else
                strSQL = strSQL & " Union ALL " & _
                    "Select /*+cardinality(D,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                    " 1 as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,Table(f_str2list([3])) D " & _
                    " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.Column_Value " & _
                    " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
            End If
            strSQL = strSQL & " Union ALL " & _
                "Select /*+cardinality(E,10)*/ 5 as ����,A.���,A.ID as ��ĿID,A.���� as ��Ŀ����,A.���� as ��Ŀ����,A.���㵥λ,A.���ηѱ�," & _
                " D.�������� as ����,C.ID as ������ĿID,C.���� as ������Ŀ,C.���� as �������,C.�վݷ�Ŀ,B.�ּ� as ����,-1 as ִ�п�������" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D,Table(f_str2list([3])) E" & _
                " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=E.Column_Value " & _
                " And " & strDateCondition & " Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                strWherePriceGrade
        End If
    End If
    
    If strSQL = "" Then Exit Function
    
    '������,����,����˳������
    strSQL = "Select * From (" & strSQL & ") Order by ����,��Ŀ����,�������"
    
    On Error GoTo errH
    If strDate <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng��Ŀid, IIf(lng����ϸĿID = 0, gCurSendCard.str��׼��Ŀ, lng����ϸĿID), str������ĿID, strPriceGrade, CDate(strDate))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng��Ŀid, IIf(lng����ϸĿID = 0, gCurSendCard.str��׼��Ŀ, lng����ϸĿID), str������ĿID, strPriceGrade)
    End If
    
    If Not rsTmp.EOF Then
        '�ȴ�����¼��
        Set rsItems = New ADODB.Recordset
        rsItems.Fields.Append "����", adSmallInt '1-����,2-����,3-������,4-���￨��
        rsItems.Fields.Append "ִ�п���ID", adBigInt
        rsItems.Fields.Append "���", adVarChar, 1
        rsItems.Fields.Append "��ĿID", adBigInt
        rsItems.Fields.Append "��Ŀ����", adVarChar, 80
        rsItems.Fields.Append "���㵥λ", adVarChar, 20, adFldIsNullable
        rsItems.Fields.Append "����", adSingle
        rsItems.Fields.Append "������Ŀ��", adSmallInt, , adFldIsNullable
        rsItems.Fields.Append "���մ���ID", adBigInt, , adFldIsNullable
        rsItems.Fields.Append "���ձ���", adVarChar, 80
        
        rsItems.CursorLocation = adUseClient
        rsItems.LockType = adLockOptimistic
        rsItems.CursorType = adOpenStatic
        rsItems.Open
        
        Set rsIncomes = New ADODB.Recordset
        rsIncomes.Fields.Append "��ĿID", adBigInt
        rsIncomes.Fields.Append "������ĿID", adBigInt
        rsIncomes.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        rsIncomes.Fields.Append "����", adSingle
        rsIncomes.Fields.Append "Ӧ��", adCurrency
        rsIncomes.Fields.Append "ʵ��", adCurrency
        rsIncomes.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
        rsIncomes.CursorLocation = adUseClient
        rsIncomes.LockType = adLockOptimistic
        rsIncomes.CursorType = adOpenStatic
        rsIncomes.Open
        
        For i = 1 To rsTmp.RecordCount
            '�Һ���Ŀ����
            If lngԭ��ID <> rsTmp!��ĿID Then
                rsItems.AddNew
                rsItems!���� = rsTmp!����
                 '0-����ȷ����,1-�������ڿ���,2-�������ڲ���,3-���������ڿ���,4-ָ������
                If bytMode = 10 Then
                    If rsTmp!ִ�п������� = -1 Then
                        rsItems!ִ�п���ID = lng�Һſ���ID
                    Else
                        rsItems!ִ�п���ID = Get�շ�ִ�п���ID(rsTmp!��ĿID, rsTmp!ִ�п�������)
                    End If
                Else
                    If rsTmp!ִ�п������� = -1 Then
                        rsItems!ִ�п���ID = 0      '0-��ʾ�Һſ���
                    Else
                        rsItems!ִ�п���ID = Get�շ�ִ�п���ID(rsTmp!��ĿID, rsTmp!ִ�п�������)
                    End If
                End If
                
                rsItems!��� = rsTmp!���
                rsItems!��ĿID = rsTmp!��ĿID
                rsItems!��Ŀ���� = rsTmp!��Ŀ����
                rsItems!���㵥λ = rsTmp!���㵥λ
                rsItems!���� = Format(Nvl(rsTmp!����, 0), "0.000")
                rsItems.Update
            End If
            lngԭ��ID = rsTmp!��ĿID
            
            '������Ŀ����
            rsIncomes.AddNew
            rsIncomes!��ĿID = rsTmp!��ĿID
            rsIncomes!������ĿID = rsTmp!������ĿID
            rsIncomes!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            rsIncomes!���� = Format(Nvl(rsTmp!����, 0), gstrFeePrecisionFmt)
            rsIncomes!Ӧ�� = Format(rsItems!���� * rsIncomes!����, "0.00")
            If Nvl(rsTmp!���ηѱ�, 0) = 1 Then
                rsIncomes!ʵ�� = rsIncomes!Ӧ��
            Else
                rsIncomes!ʵ�� = Format(GetActualMoney(str�ѱ�, rsTmp!������ĿID, rsIncomes!Ӧ��, rsTmp!��ĿID), "0.00")
            End If
            rsIncomes.Update
            rsTmp.MoveNext
        Next
        ReadRegistPrice = rsItems.RecordCount
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set rsItems = Nothing
    Set rsIncomes = Nothing
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim strValue As String
    On Error Resume Next
    
    '�ϰ�ʱ��
    strValue = UCase(zlDatabase.GetPara(1, glngSys, , "08:00 AND 12:00"))
    gstr�ϰ�ʱ�� = Format(Trim(Split(strValue, "AND")(0)), "HH:mm")
    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '������ʾ��ʽ
    'gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"

    '�Һ�Ʊ�ݺ��볤��
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(IIf(gblnSharedInvoice, 0, 3)))
    'gbyt�ſ� = Val(Split(strValue, "|")(4))
    
    '�Һ���Ч����
    '���˺�:34717
    '��λ:ǰһλ���ܹҺ�;��һλ����Һ�
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gSysPara.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gSysPara.Sy_Reg.bytNoDayseMergency = Val(Mid(strValue, 2, 1))
    If gSysPara.Sy_Reg.bytNODaysGeneral = 0 Then gSysPara.Sy_Reg.bytNODaysGeneral = 1
    If gSysPara.Sy_Reg.bytNoDayseMergency = 0 Then gSysPara.Sy_Reg.bytNoDayseMergency = 1
    
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill�Һ� = (Mid(strValue, IIf(gblnSharedInvoice, 1, 4), 1) = "1")
    'gblnBill�ſ� = (Mid(strValue, 5, 1) = "1")
    
        
    'һ��ͨ������֤
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdblԤ��������鿨 = Val(Split(strValue, "|")(0))
    gbytԤ����˷��鿨 = Val(Split(strValue, "|")(1))
    gbln���ѿ��˷��鿨 = zlDatabase.GetPara(282, glngSys) = "1"
            
    'ˢ��Ҫ����������
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
        
    '�Һ������ԤԼ����
    gintԤԼ���� = zlDatabase.GetPara(66, glngSys, , 15)
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    gbln���֤Ψһ = Val(zlDatabase.GetPara("ͬһ���ֻ֤�ܶ�Ӧһ����������", glngSys)) = 1    '117954
    Call InitAddressLength
    InitSysPar = True
End Function

Public Sub InitLocPar(lngModul As Long)
'���ܣ���ʼ����������
    Dim strValue As String
    On Error Resume Next
                
    
    'b.���ݿ�洢�Ĺ���ȫ�ֲ���
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytRegistMode = Val(Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(0))
    If Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(zlDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
        
    
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    If lngModul = 1111 Then
        glngInterval = Val(zlDatabase.GetPara("�Զ�ˢ�¼��", glngSys, lngModul))
        gbln�Զ������ = zlDatabase.GetPara("�Զ������", glngSys, lngModul) = "1"
        gblnPrice = zlDatabase.GetPara("��Ϊ���۵�", glngSys, lngModul) = "1"
        gblnPrePayPriority = zlDatabase.GetPara("����ʹ��Ԥ����", glngSys, lngModul) = "1"
        
        'ȱʡֵ
        gstr���ʽ = zlDatabase.GetPara("ȱʡ���ʽ", glngSys, lngModul)
        gstr�ѱ� = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, lngModul)
        gstr�Ա� = zlDatabase.GetPara("ȱʡ�Ա�", glngSys, lngModul)
        gstr���㷽ʽ = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, lngModul)
        
        '��������Һſ���ID
        gstr�Һſ���ID = zlDatabase.GetPara("�Һſ���", glngSys, lngModul)
        
        
        gblnSeekName = zlDatabase.GetPara("����ģ������", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("������������", glngSys, lngModul))
        gbln�ɿ���� = zlDatabase.GetPara("�ɿ�ҺŽ���", glngSys, lngModul) = "1"
        gblnҽ�� = zlDatabase.GetPara("����ҽ��", glngSys, lngModul) = "1"
        gblnPrintFree = zlDatabase.GetPara("����ô�ӡ", glngSys, lngModul) = "1"
        gbytInvoice = Val(zlDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, lngModul, , 1))
        gByt��ӡ�������� = Val(zlDatabase.GetPara("���������ӡ��ʽ", glngSys, lngModul, , 1))
        gblnPrintCase = zlDatabase.GetPara("��ӡ������ǩ", glngSys, lngModul, "0") = "1"
        gbln������� = Val(zlDatabase.GetPara("�ƻ��Ű�Һ�Ĭ�Ͻ���", glngSys, lngModul, 0)) = 1
        
        
        gbln���� = zlDatabase.GetPara("��������", glngSys, lngModul) = "1"
        gbln�Ա� = zlDatabase.GetPara("�����Ա�", glngSys, lngModul) = "1"
        gbln���� = zlDatabase.GetPara("��������", glngSys, lngModul) = "1"
        gbln��ͥ��ַ = zlDatabase.GetPara("�����ͥ��ַ", glngSys, lngModul) = "1"
        gbln���ʽ = zlDatabase.GetPara("���븶�ʽ", glngSys, lngModul) = "1"
        gbln�ѱ� = zlDatabase.GetPara("����ѱ�", glngSys, lngModul) = "1"
        gbln���㷽ʽ = zlDatabase.GetPara("������㷽ʽ", glngSys, lngModul) = "1"
        gbln�绰 = zlDatabase.GetPara("������ϵ�绰", glngSys, lngModul) = "1"
        
        
        gblnAutoAddName = zlDatabase.GetPara("�Զ���������", glngSys, lngModul) = "1"
        gblnNewCardNoPop = zlDatabase.GetPara("������������", glngSys, lngModul) = "1"
        gbln���ѽ����� = zlDatabase.GetPara("��ȡ����", glngSys, lngModul) <> "1"
        gbln�˷��ش� = zlDatabase.GetPara("�˷��ش�", glngSys, lngModul) = "1"
        '����:35176
        gbyt���������Ϣ = Val(zlDatabase.GetPara("�˺����������Ϣ", glngSys, lngModul))
        
        '�շѺ͹ҺŹ���Ʊ��
        gblnSharedInvoice = zlDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121) = "1"
        '���ع��ùҺ�����ID
        If gblnSharedInvoice Then
            glng�Һ�ID = Val(zlDatabase.GetPara("�����շ�Ʊ������", glngSys, 1121, ""))
        Else
            glng�Һ�ID = Val(zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, lngModul, ""))
        End If
        If glng�Һ�ID > 0 Then
            If Not ExistBill(glng�Һ�ID, IIf(gblnSharedInvoice, 1, 4)) Then
                If gblnSharedInvoice Then
                    zlDatabase.SetPara "�����շ�Ʊ������", "0", glngSys, 1121
                Else
                    zlDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, lngModul
                End If
                glng�Һ�ID = 0
            End If
        End If
        
        gstr�ſ�ID = Val(zlDatabase.GetPara("���þ��￨����", glngSys, lngModul, ""))
        
        
        '�Ƿ�ʹ��LED����������
        gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
        '��ʼ����������Ϣ
        Call InitSendCardPreperty(lngModul)
    ElseIf lngModul = 1114 Then
        Call InitLocVisitPlanPar(1114)
    End If
End Sub

Public Sub InitLocVisitPlanPar(ByVal lngModul As Long)
    '��ʼ���ٴ������ģ�����
    With gVisitPlan_ModulePara
        .byt������ӡ��ʽ = Val(zlDatabase.GetPara("������ӡ��ʽ", glngSys, lngModul, "0"))
        .str��Դά��վ�� = zlDatabase.GetPara("δ����վ��ĺ�Դ��ά��վ��", glngSys, lngModul)
        .byt����ȽϷ�ʽ = Val(zlDatabase.GetPara("��������ȽϷ�ʽ", glngSys, lngModul))
    End With
End Sub


Public Sub InitSendCardPreperty(ByVal lngModule As Long, Optional lng�����ID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ˢ������
    '����:���˺�
    '����:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str���� As String, varData As Variant, i As Long
    Dim varTemp  As Variant, ty_Card As Ty_CardProperty
    If lng�����ID <> 0 Then
        lngCardTypeID = lng�����ID
    Else
        lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, lngModule, 0))
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�����:57326
    strSQL = "" & _
    "   Select Id, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����, nvl(�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� ,nvl(ȱʡ��־,0) as ȱʡ��־, " & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������," & _
    "           nvl(�Ƿ��ƿ�,0) as �Ƿ��ƿ�,nvl(�Ƿ񷢿�,0) as �Ƿ񷢿�,nvl(�Ƿ�д��,0) as �Ƿ�д��, " & _
    "           nvl(��������,0) as ��������, nvl(��������,'1000')  as ��������,nvl(��������,0) as �������� " & _
    "    From ҽ�ƿ���� A" & _
    "    Where nvl(�Ƿ�����,0)=1 And (ID=[1] " & IIf(lng�����ID = 0, "or nvl(ȱʡ��־,0)=1", "") & ")" & _
    "    Order by ����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ʼ��������", lngCardTypeID)
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        '85565,���ϴ�,2015/7/10:��������
        With ty_Card
            .lng�����ID = Val(Nvl(rsTemp!id))
            .str������ = Nvl(rsTemp!����)
            .lng���ų��� = Val(Nvl(rsTemp!���ų���))
            .lng���㷽ʽ = Trim(Nvl(rsTemp!���㷽ʽ))
            .bln���ƿ� = Val(Nvl(rsTemp!�Ƿ�����)) = 1
            .bln�ϸ���� = Val(Nvl(rsTemp!�Ƿ��ϸ����)) = 1
            .str�������� = Nvl(rsTemp!��������)
            .int���볤�� = Val(Nvl(rsTemp!���볤��))
            .int���볤������ = Val(Nvl(rsTemp!���볤������))
            .int������� = Val(Nvl(rsTemp!�������))
            .bln���￨ = .str������ = "���￨" And Val(Nvl(rsTemp!�Ƿ�̶�)) = 1
            .str��׼��Ŀ = Trim(Nvl(rsTemp!�ض���Ŀ))
            .blnȱʡ��־ = Val(Nvl(rsTemp!ȱʡ��־)) = 1
            '�����:56599
            .bln�Ƿ��ƿ� = Val(Nvl(rsTemp!�Ƿ��ƿ�)) = 1
            .bln�Ƿ񷢿� = Val(Nvl(rsTemp!�Ƿ񷢿�)) = 1
            .bln�Ƿ�д�� = Val(Nvl(rsTemp!�Ƿ�д��)) = 1
            '�����:57326
            .lng�������� = Val(Nvl(rsTemp!��������))
            .bln�ظ�ʹ�� = Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��)) = 1
            .str�������� = Nvl(rsTemp!��������, "1000")
            .byt�������� = Val(Nvl(rsTemp!��������))
            .blnOneCard = False
            .str������ = Nvl(rsTemp!����)
            If Trim(Nvl(rsTemp!�ض���Ŀ)) <> "" Then
                Set .rs���� = zlGetSpecialItemFee(Trim(Nvl(rsTemp!�ض���Ŀ)))
                If .bln���￨ Then .blnOneCard = GetOneCard.RecordCount > 0
            Else
                Set .rs���� = Nothing
            End If
            str���� = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, lngModule, "0")
            '����ID,�����ID|...
             .lng�������� = 0
            varData = Split(str����, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If Val(varTemp(1)) = .lng�����ID Then
                        .lng�������� = Val(varTemp(0)): Exit For
                    End If
                 End If
            Next
        End With
    End If
    gCurSendCard = ty_Card
End Sub

Public Function Check��������(lng����ID As Long, lng�����ID As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ����Ƿ����Ʋ��˵ķ�������
    '���:lng����ID - ����ID;lng�����ID  - ҽ�ƿ������ID
    '     blnShowMsg-�Ƿ񵯳���ʾ��
    '����:����
    '����:57326
    '����:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    strSQL = "Select ����, �������� " & _
            "   From ҽ�ƿ���� A, ����ҽ�ƿ���Ϣ B Where A.ID = B.�����ID And B.״̬=0 And B.����ID=[1] And B.�����ID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������", lng����ID, lng�����ID)
    If rsTemp.RecordCount = 0 Then Check�������� = True: Exit Function
    Select Case Val(Nvl(rsTemp!��������, 0))
        Case 0 '������
            Check�������� = True
        Case 1 'ͬһ������ֻ����һ�ſ�
            If blnShowMsg Then
                MsgBox "�ò����Ѿ�����" & Nvl(rsTemp!����) & ",�����ڽ��з�������!", vbInformation + vbOKOnly
            End If
            Check�������� = False
        Case 2 'ͬһ�������������ſ�,����Ҫ����
            If blnShowMsg Then
                Check�������� = MsgBox("�ò����Ѿ�����" & Nvl(rsTemp!����) & ",�Ƿ�Ҫ���з�������?", vbQuestion + vbYesNo) = vbYes
            Else
                Check�������� = True
            End If
    End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetNext�ű�() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(����) as ���� From �ҺŰ��� Where Length(����)=(Select Max(Length(����)) From �ҺŰ���)"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    
    If Not rsTmp.EOF Then GetNext�ű� = zlStr.Increase(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ҩ��(lng����ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ҩ��ID,����ҩ��,������Ӧ From ���˹���ҩ�� Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", lng����ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Get����ҩ�� = Get����ҩ�� & ";" & IIf(IsNull(rsTmp!����ҩ��ID), "", rsTmp!����ҩ��ID) & "|" & IIf(IsNull(rsTmp!����ҩ��), "", rsTmp!����ҩ��) & "|" & Nvl(rsTmp!������Ӧ)
            rsTmp.MoveNext
        Next
        Get����ҩ�� = Mid(Get����ҩ��, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetField(strSQL As String) As String
'���ܣ�����SQL������ݷ��ص�һ���ֶ�����
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    If Not rsTmp.EOF Then GetField = IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBill(ByVal frmParent As Object, ByVal bytFunc As Byte, _
        ByVal strNO As String, ByVal lng����ID As Long, ByVal intInsure As Integer, _
        ByVal blnVirtualPrint As Boolean, _
        Optional strUseType As String, Optional ByVal bln�ش� As Boolean, _
        Optional ByVal blnConfirmInvoice As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ǰ�տ��¼���´�ӡһ��Ʊ��
    '���:
    '   bytFunc:2-�˷Ѵ�ӡ,3-�ش�,4-����Ʊ��
    '   blnVirtualPrint-ҽ���ӿ��ڵ��ô�ӡ��HISֻ��Ʊ�Ų�ʵ�ʴ�ӡ
    '   blnConfirmInvoice:�Ƿ���Ҫȷ�Ϸ�Ʊ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 17:18:19
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    Dim strInvoice As String
    Dim blnValid As Boolean
    Dim lng����ID As Long, strBackInvoice As String
    Dim blnReprint As Boolean
    
    '����ϸ����Ʊ��ʹ��
    If gblnBill�Һ� Then
        If bln�ش� Then
            lng����ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), glng�Һ�ID, , strUseType)
            Select Case lng����ID
                Case -1
                    MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            If lng����ID <= 0 Then Exit Function
        End If
        If bytFunc = 3 Then
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
        End If
        blnReprint = bln�ش�
    End If
    
     'ȡ��һ��Ʊ�ݺ���
    If Not gblnBill�Һ� Then
        If bln�ش� = False And bytFunc = 2 Then Exit Function
        '�п����ǵ�һ��ʹ��
        Do
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            If gblnSharedInvoice Then
                strInvoice = zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
            Else
                strInvoice = zlDatabase.GetPara("��ǰ�Һ�Ʊ�ݺ�", glngSys, 1111)
            End If
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ���ҺŽ�Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                If blnConfirmInvoice Then
                    strInvoice = UCase(InputBox("��ȷ�ϹҺ�" & IIf(bytFunc = 4, "����", "�ش�") & "ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                End If
            End If
                
            '�û�ȡ������,�����ӡ
            If strInvoice = "" Then
                If MsgBox("��ȷ��������Һ�Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnValid = True
            Else
                '���������Ч��
                If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                    MsgBox "����ĹҺ�Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                    blnConfirmInvoice = True
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
    Else
        If blnReprint Then
            Do
                '����Ʊ�����ö�ȡ
                strInvoice = GetNextBill(lng����ID)
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ�ҺŽ�Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                ElseIf blnConfirmInvoice Then
                    strInvoice = UCase(InputBox("��ȷ�ϹҺ�" & IIf(bytFunc = 4, "����", "�ش�") & "ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function
                
                '���������Ч��
                If GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), 1, lng����ID, glng�Һ�ID, strInvoice, strUseType) = -3 Then
                    MsgBox "������ĹҺ�Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    blnConfirmInvoice = True
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        Else
            strInvoice = ""
        End If
    End If
    
    'ִ�����ݴ���
    Call frmPrint.ReportPrint(bytFunc, strNO, strBackInvoice, lng����ID, glng�Һ�ID, strInvoice, _
        zlDatabase.Currentdate, , , , blnVirtualPrint, bytFunc = 2, strUseType)

    RePrintBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub TaxInterface(ByVal byt���� As Byte, ByVal strPrintNO As String, ByVal strModiNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˰�ش�ӡ�ӿ�
    '���:byt����-1-������ӡ(���޸�);2-�ش�;3-�˷�
    '        strPrintNO-Ҫ��ӡ�ĵ��ݺţ����ʱ�ö��ŷָ�:'F0000001','F0000002',...
    '        strModiNos-�޸Ķ൥���е�һ��ʱ,ָ�ö��ŵ��ݵ�����NO���ö��ŷָ�:'F0000001','F0000002',...
    '����:���˺�
    '����:2013-03-27 14:24:03
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    'δ����˰��,ֱ�ӷ���
    If Not gblnTax Then Exit Sub
    If byt���� = 3 Then
        '�˷�
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If byt���� = 2 Then
        '�ش�
        MsgBox "����׼����֮��ȷ����ʼ��ӡ��", vbInformation, gstrSysName
        gstrTax = gobjTax.zlTaxOutReput(gcnOracle, strPrintNO, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If strModiNos <> "" Then
        gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strModiNos, "2")
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, strPrintNO, "2")
    If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Public Function CheckExecuted(strNO As String, blnEnableDel As Boolean) As Boolean
'���ܣ��ж�ָ���ĹҺŵ����Ƿ��Ѿ���ִ��,����ҽ��������ҽ�������Ϻ�,ȡ������,Ҳ��ʾִ�й���
'����:blnEnableDel-�Ƿ�����ֻ����ȡ����ҽ���Ĳ����˺�
'����:True ��ʾ�ѱ�ִ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    CheckExecuted = False
    If blnEnableDel Then strSQL = " And ҽ��״̬<>4"
    strSQL = _
        " Select count(ID) num From ���˹Һż�¼ Where NO=[1] And ִ��״̬>0 and ��¼����=1 and ��¼״̬ =1 " & _
        " Union All " & _
        " Select count(ID) num From ����ҽ����¼ Where �Һŵ�=[1] And (������Դ=1 or ������Դ=2)" & strSQL
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    Do While Not rsTmp.EOF
        If rsTmp!Num > 0 Then
            CheckExecuted = True
        End If
        rsTmp.MoveNext
    Loop
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInsure(strNO As String) As Integer
'���ܣ��жϹҺż�¼���Ƿ����ָ����ҽ�����㷽ʽ
'������strNO=�Һŵ��ݺ�
'���أ���������򷵻ص��ݵ�ʱ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select B.���� From ������ü�¼ A,���ս����¼ B" & _
       " Where A.��¼����=4 And A.���=1 And A.��¼״̬ IN(1,3) And A.NO=[1]" & _
       " And B.����=1 And A.����ID=B.��¼ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.EOF Then ExistInsure = Val(IIf(IsNull(rsTmp!����), 0, rsTmp!����))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ExistFee(strNO As String) As Boolean
'���ܣ��жϲ��˵ĹҺŵ������Ƿ��������Һŵ�,�����,�򲻼���Ƿ���������,
'      ���û��,�����Ƿ��չ���,����黮�۷���,���ʷ���,�Զ�����,���￨����
'������strNO=�Һŵ��ݺ�

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '�Һŵ�����֮���Ƿ��йҺŵ�(���ܹҺſ���)(����˺�,���˹Һż�¼�е����ݼ�¼״̬��Ϊ1)
    strSQL = "Select a.NO, a.����id, a.ִ�в���id, a.�Ǽ�ʱ��,b.ִ�в���id as �Һſ���id" & vbNewLine & _
            "From ���˹Һż�¼ a, ���˹Һż�¼ b" & vbNewLine & _
            "Where b.No = [1] And a.����id = b.����id and a.��¼����=1 and a.��¼״̬=1 and b.��¼����=1 and b.��¼״̬=1 And a.�Ǽ�ʱ�� >= Trunc(b.�Ǽ�ʱ��) and a.��¼����=1 and a.��¼״̬=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    
    If Not rsTmp.EOF Then
        '����Һŵ���ͬһ�����ж��ŹҺŵ�,�Ͳ�����Ƿ�����������,��Ϊ�޷����������ŹҺŵ���
        rsTmp.Filter = "ִ�в���id=" & rsTmp!�Һſ���id
        If rsTmp.RecordCount > 1 Then Exit Function
        
        '������ŹҺŵ��Ŀ����ڱ��ιҺź��Ƿ���ڷ���(δ�˷�)
        rsTmp.Filter = "NO='" & strNO & "'"
        strSQL = "Select NO" & vbNewLine & _
             "From ������ü�¼" & vbNewLine & _
             "Where ����id=[1] And ��������ID+0=[2] And �Ǽ�ʱ��+0>=[3]" & vbNewLine & _
             "      And ��¼����=1 And ��¼״̬>0 " & vbNewLine & _
             "Group by NO Having Sum(����*����)<>0"
             
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", Val(rsTmp!����ID), Val(rsTmp!ִ�в���id), CDate(rsTmp!�Ǽ�ʱ��))
        ExistFee = Not rsTmp.EOF
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPriceHaveFee(strNO As String, ByRef str����NO As String) As Boolean
'����:���ҺŲ����Ļ��۵��Ƿ��Ѿ��չ���
'����:δ�շѵĻ��۵�

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select NO,��¼״̬ From ������ü�¼ " & _
            " Where ��¼����=1 And ����ID=(Select ����ID From ���˹Һż�¼ Where NO=[1] And ��¼����=1 and ��¼״̬=1 and  Rownum<2 )" & _
            " And ��¼״̬ IN(0,1,3) And ���=1 And ժҪ Like [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, "%" & strNO & "%")
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!��¼״̬, 0) = 1 Then
            MsgBox "�ùҺŵ���Ӧ�ķ����Ѿ��������շѣ������˺š�", vbInformation, gstrSysName
            CheckPriceHaveFee = True
        ElseIf Nvl(rsTmp!��¼״̬, 0) = 0 Then
            str����NO = rsTmp!NO
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
'����:���ҽ�����Ƿ��Ѵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select 1 From ������Ϣ Where ҽ���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����,�����ҽ�����Ѵ���!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBill����ID(ByVal strNO As String, ByVal byt��¼���� As Byte, _
    Optional ByRef lng����ID As Long, Optional ByRef bln���ʷ��� As Boolean) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ݵĽ���ID
    '���:strNo-���ݺ�
    '       byt��¼����:4-�Һ�,5-���￨
    '����:lng����ID-���ز���ID
    '       bln���ʷ���-���ظõ����Ƿ���ʷ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 16:23:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    lng����ID = 0
    If byt��¼���� <> 5 Then
        strSQL = "Select ����ID,����ID,���ʷ��� From ������ü�¼" & _
           " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(1,3) And ���=1"
    Else
        strSQL = "Select ����ID,����ID,���ʷ��� From סԺ���ü�¼" & _
           " Where NO=[1] And ��¼����=[2] And ��¼״̬ IN(1,3) And ���=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, byt��¼����)
    If rsTmp.EOF Then Exit Function
    lng����ID = Val(Nvl(rsTmp!����ID))
    GetBill����ID = Val(Nvl(rsTmp!����ID))
    bln���ʷ��� = Val(Nvl(rsTmp!���ʷ���)) = 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ڶ��ν���
    '����:True-���ڶ��ν������� False-�����ڶ��ν�������
    '���:strNO-�Һŵĵ��ݺ�
    '����:������
    '����:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "" & _
    " Select 1" & vbNewLine & _
    " From ���ò����¼ A," & vbNewLine & _
    "     (Select Distinct ����id" & vbNewLine & _
    "       From ������ü�¼" & vbNewLine & _
    "       Where NO = [1] And ��¼���� = 4" & vbNewLine & _
    "       Union" & vbNewLine & _
    "       Select Distinct ����id From סԺ���ü�¼ Where NO = [1] And ��¼���� = 5) B" & vbNewLine & _
    " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", strNO)

    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
End Function

Public Function GetDoctor(Optional ByVal lngSectID As Long = 0, Optional strCodeAliasName As String = "����") As ADODB.Recordset
    '�õ�ָ�������µ�����ҽ��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        "Select c.��� " & IIf(strCodeAliasName = "", "", " as " & strCodeAliasName) & ",c.����,c.����,c.id From ��Ա����˵�� a, ������Ա b ,��Ա�� c" & vbCrLf & _
        "Where b.��Աid=c.id And b.��Աid=a.��Աid  And  a.��Ա����=[1] And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) " & vbCrLf & _
        " And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
        IIf(lngSectID = 0, "", "   And b.����id = [2]")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", "ҽ��", lngSectID)
    Set GetDoctor = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ�Ƹ��ʽ(byt��� As Byte) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����,���� From ҽ�Ƹ��ʽ Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", byt���)
    If Not rsTmp.EOF Then
        Getҽ�Ƹ��ʽ = rsTmp!���� & "-" & rsTmp!����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDefaultTime(ByVal str�ű� As String, vDate As Date) As String
'���ܣ����ݺű���ָ�����ڵ�ԤԼ�ó�ָ�����ڵ�ȱʡʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Decode(" & Weekday(vDate) & ",1,A.����,2,A.��һ,3,A.�ܶ�,4,A.����,5,A.����,6,A.����,7,A.����,NULL)"
    strSQL = "Select B.��ʼʱ�� From �ҺŰ��� A,ʱ��� B Where " & strSQL & "=B.ʱ���"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlRegEvent")
    If Not rsTmp.EOF Then
        GetDefaultTime = Format(rsTmp!��ʼʱ��, "HH:mm:ss")
    Else
        GetDefaultTime = "00:00:00"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist�����(str����� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ��������Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ������Ϣ Where �����=[1] And ����ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str�����, lng����ID)
    If rsTmp.RecordCount > 0 Then Exist����� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist�ֻ���(str�ֻ��� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ���ֻ����Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ������Ϣ Where �ֻ���=[1] And ����ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str�ֻ���, lng����ID)
    If rsTmp.RecordCount > 0 Then Exist�ֻ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistCardFee(ByVal strNO As String, ByRef lng����ID As Long, Optional ByRef strCardNo As String) As String
'���ܣ��ж�ָ���Һŵ��Ƿ�ͬʱ��ȡ�˾��￨��
'       ���ؾ��￨���õ��ݺ�,���￨���õĽ���ID
'      strCardNo - ����
    Dim rsTmp As ADODB.Recordset
    Dim rsҽ�ƿ���� As Recordset '�����:56599
    Dim str���� As String '�����:56599
    Dim strSQL As String
    
    On Error GoTo errH
    '�����:58536
    strSQL = "Select NO,����ID,ʵ��Ʊ�� as ���� From סԺ���ü�¼ Where ��¼����=5 And ��¼״̬=1 And (����ID,�Ǽ�ʱ��) = " & _
            " (Select ����ID,�Ǽ�ʱ�� From ������ü�¼ Where ��¼����=4 And NO=[1] And Rownum=1)"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO)
    If rsTmp.RecordCount > 0 Then
        ExistCardFee = rsTmp!NO
        lng����ID = Val(Nvl(rsTmp!����ID))
        strCardNo = Nvl(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check�Һ�ʱ����(strNO As String, str�Һ����� As String) As Boolean
    '����:�ж��˺�ʱ,������Ϣ�Ƿ��ǹҺ�ʱ������,�����,��Ҫ��ʾ�Ƿ���������
    '���ڹҺ�ʱ�����½��Ĳ��˵�����ʱ����Һŵ�ʱ�䲻һ��,�Լ����˿�������,�������������ٹҺŵ����,����,�����ò��˵Ǽ�ʱ����Һ�ʱ��ֱ���ж�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select ����id From ������Ϣ   " & _
    "   Where  ABS(To_Date([2])-�Ǽ�ʱ��)< (Select Decode(Max(nvl(����,0)),0,[3],[4])  From ���˹Һż�¼ Where NO=[1]  and ��¼����=1 and ��¼״̬=1) " & _
    "       And ����ID=(Select ����id From ���˹Һż�¼  Where ����id = (Select ����id From ���˹Һż�¼ Where No = [1] and ��¼����=1 and ��¼״̬=1) and ��¼״̬=1 and ��¼����=1 " & _
    "   Group By ����id " & _
    "   Having Count(Id) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, CDate(str�Һ�����), gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    Check�Һ�ʱ���� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str���֤�� As String) As String
'���ܣ���鲡���Ƿ����������Ϣ
'���أ����Ƽ�¼�Ĳ���ID��,��"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = _
        " Select ����ID,����,Nvl(���֤��,'δ�Ǽ�') ���֤��,�����,Nvl(��ͥ��ַ,'δ�Ǽ�') ��ַ,To_Char(�Ǽ�ʱ��,'YYYY-MM-DD') �Ǽ�ʱ�� " & _
        " From ������Ϣ Where ���֤��=[1]" & _
        " Order by ����ID Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str���֤��)
    
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!����ID & ",����:" & rsTmp!���� & ",�����:" & Nvl(rsTmp!�����, "��") & ",���֤��:" & rsTmp!���֤�� & ",��ַ:" & rsTmp!��ַ & ",�Ǽ�����:" & rsTmp!�Ǽ�ʱ��
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function Select����(ByVal frmMain As Form, ByVal lngMoudle As Long, ByVal rs���� As ADODB.Recordset, cbo���� As ComboBox, ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�����
    '���: frmMain-������
    '      rs����-���ƵĿ��ҵı��ؼ�,
    '      cbo����-����
    '      strKey-ѡ����ҵ�����
    '����:
    '����:
    '����:���˺�
    '����:2009-10-12 09:57:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSearch As String
    
    strSearch = "'" & gstrLike & strKey & "%'"
    If IsNumeric(strKey) Then   '�������ȫ����
        rs����.Filter = "���� like " & UCase(strSearch)
        rs����.Sort = "����"
    ElseIf zlCommFun.IsCharAlpha(strKey) Then  '�������ȫ��ĸ
        rs����.Filter = "���� like " & UCase(strSearch)
        rs����.Sort = "����"
    ElseIf zlCommFun.IsCharChinese(strKey) Then '�Ƿ��к���,'���к���,�϶���������
        rs����.Filter = "���� like " & strSearch
        rs����.Sort = "����"
    Else
        rs����.Filter = "���� like " & strSearch & " or ���� like " & strSearch & " or ���� like " & strSearch
        rs����.Sort = "����"
    End If
    If rs����.RecordCount = 0 Then
        rs����.Filter = 0: Exit Function
    End If
    If rs����.RecordCount = 1 Then
        zlControl.CboLocate cbo����, Val(Nvl(rs����!id)), True
        rs����.Filter = 0: Select���� = True: Exit Function
    End If
    '����ѡ����
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngMoudle, cbo����, rs����, True, "", "�������", rsReturn) Then
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                zlControl.CboLocate cbo����, Val(Nvl(rsReturn!id)), True
                DoEvents
                If cbo����.Enabled Then cbo����.SetFocus
                Select���� = True: Exit Function
            End If
        End If
    End If
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Աѡ��ѡ����
    '���:cboSel-ָ���Ĳ���ѡ�񲿼�
    '     rsPerson-ָ������Ա��Ϣ(ID,���,����,����)
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str����-��������(������,���в���Ա��)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String, strLike As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���� <> "" Then
        str���� = zlCommFun.SpellCode(str����)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strSearch Then lngID = Nvl(!id): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!���) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!id))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!���)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!id)
        
    '���˺�:ֱ�Ӷ�λ
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!id))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboSel
End Function


Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub
Public Function zlGetIDCardSex(ByVal strInput As String) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ���֤�ŵ��Ա�
    '���أ��Ա�
    '���ƣ����˺�
    '���ڣ�2010-07-15 10:31:08
    '˵����15λ���֤���룺��7��8λΪ�������(��λ��)����9��10λΪ�����·ݣ���11��12λ����������ڣ���15λ�����Ա�����Ϊ�У�ż��ΪŮ��
   '          18λ���֤���룺��7��8��9��10λΪ�������(��λ��)����11����12λΪ�����·ݣ���13��14λ����������ڣ���17λ�����Ա�����Ϊ�У�ż��ΪŮ��
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSex As String, i As Integer
    i = zlCommFun.ActualLen(strInput)
    If i <> 15 And i <> 18 Then Exit Function
    i = Val(Mid(strInput, IIf(i = 15, 15, 17), 1))
    If i Mod 2 = 0 Then
        zlGetIDCardSex = "Ů"
    Else
        zlGetIDCardSex = "��"
    End If
End Function
Public Function zlGetIDCardAge(ByVal strbirthday As Date, ByRef str��λ As String) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݳ������ڻ�ȡ����
    '���ƣ����˺�
    '���ڣ�2010-07-15 10:48:15
    '˵����Zl_Age_Calc
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngDiffDay As Long
    If IsDate(strbirthday) = False Then Exit Function
    lngDiffDay = Now - CDate(strbirthday)
    If lngDiffDay < 32 Then '����Ϊ��λ
        str��λ = "��": zlGetIDCardAge = lngDiffDay
    ElseIf lngDiffDay < 365 Then
        str��λ = "��": zlGetIDCardAge = Int(lngDiffDay / 30)
    Else
        str��λ = "��": zlGetIDCardAge = Int(lngDiffDay / 365)
    End If
End Function

Public Sub zlAutoCalcBackLists(ByVal lng����ID As Long)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��Զ����������
    '���ƣ����˺�
    '���ڣ�2010-07-15 16:32:10
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo Errhand:
    strSQL = "Zl_Regist_Autointoblacklist(" & lng����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, "����ԤԼ������"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetPatiInfo(lng����ID As Long) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ���Ĳ�����Ϣ
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-07-19 10:56:31
    '˵����
    '------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '��ҳID=0ʱ(����NULL)����ʾԤԼ��Ժ
    strSQL = _
        " Select A.����ID,Decode(B.����ID,NULL,NULL,Nvl(B.��ҳID,0)) as ��ҳID," & _
        "           A.����,A.סԺ��,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID(+) And A.����ID=[1]" & _
        " Order by Nvl(B.��ҳID,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID)
    
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlPatiMerge(ByVal lng���ϲ�����ID As Long, ByRef lng�ϲ�����ID As Long, Optional blnInput�ϲ�ԭ�� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������������Ϣ���кϲ�
    '���:blnInput�ϲ�ԭ��-�Ƿ�Ҫ������ϲ�ԭ��
    '���أ��ϲ��ɹ�,����true, ���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-19 10:53:12
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPatiS As ADODB.Recordset, rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str�ϲ�ԭ�� As String
    
    If lng�ϲ�����ID <= 0 Or lng���ϲ�����ID <= 0 Then
        Exit Function
    End If
    
    If lng�ϲ�����ID = lng���ϲ�����ID Then
        MsgBox "��ͬ���˲��ý��кϲ�������", vbInformation, gstrSysName
        Exit Function
    End If
        
    Set rsPatiS = GetPatiInfo(lng���ϲ�����ID)
    Set rsPatiO = GetPatiInfo(lng�ϲ�����ID)
    
    'A��B��һ��������ԤԼ��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Nvl(rsPatiS!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsNull(rsPatiO!��ҳID) And Nvl(rsPatiO!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'AB��ס��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Not IsNull(rsPatiO!��ҳID) Then
        '1.��סԺ����Ժ,������(�Ⱥ�סԺ����Ϊ����Ժ-��Ժ,��Ժ-��Ժ����������Ժ-��Ժ,��Ժ-��Ժ)
        '��Ϊ�����˺ϲ���,���򲻶��⴦���Զ���Ժ������Ժ
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!��Ժ���� <= rsPatiO!��Ժ���� Then
            If IsNull(rsPatiS!��Ժ����) Then
                MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!��Ժ����) Then
                MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.ʱ�佻����ʾ�Ƿ����
        Curdate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    MsgBox "���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                        vbCrLf & "���ཻ�棬���ܽ��кϲ���", _
                        vbInformation, gstrSysName
                        Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '�ϲ�ԭ��
    If blnInput�ϲ�ԭ�� Then
        str�ϲ�ԭ�� = InputBox("�ϲ��������ܳ���,������!" & vbCrLf & vbCrLf & "������ϲ�ԭ��:" & vbCrLf & vbCrLf, gstrSysName, "")
        If zlCommFun.ActualLen(str�ϲ�ԭ��) > 250 Then
            MsgBox "�ϲ�ԭ���ܶ���250���ַ�,�밴Ctrl+C�������������,����ִ��ʱ������:" & _
                vbCrLf & vbCrLf & str�ϲ�ԭ��, vbInformation, gstrSysName
            Exit Function
        ElseIf Trim(str�ϲ�ԭ��) = "" Then
            MsgBox "��������ϲ�ԭ����ܽ��кϲ�!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        str�ϲ�ԭ�� = "ԤԼ�ҺŽ���ʱ,�Զ��ϲ�!"
    End If
    Screen.MousePointer = 11
    DoEvents
    On Error GoTo errH
    strSQL = "zl_������Ϣ_MERGE(" & lng���ϲ�����ID & "," & lng�ϲ�����ID & ",'" & str�ϲ�ԭ�� & "','" & UserInfo.���� & "')"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "�Զ��ϲ�����")
    On Error GoTo 0
    Screen.MousePointer = 0
        
    '�ϲ���Ӧֻʣһ������
    strSQL = "Select ����ID From ������Ϣ Where ����ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�Զ��ϲ�����", lng���ϲ�����ID, lng�ϲ�����ID)
    
    lng�ϲ�����ID = Val(rsTmp!����ID)
    MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ " & lng�ϲ�����ID & "��", vbInformation, gstrSysName
    zlPatiMerge = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String, _
    Optional blnվ�� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnվ��-�Ƿ����վ������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single, strSQL As String
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    If strTable = "����" Then
        strSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 And Nvl(����,0) <3 "
    Else
        strSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    End If
    If strKey <> "" Then
        strSQL = strSQL & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  "
    End If
    strSQL = strSQL & strWhere & IIf(blnվ��, zl_��ȡվ������, "") & _
    "   order by ����"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.Hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If blnδ�ҵ����� Then
            If zlCommFun.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str����
                    End If
                End With
            Else
                If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgbox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!����)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            End If
        End With
    Else
        If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!����)
        objCtl.Tag = Nvl(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.Hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    Dim int���� As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, strSQL, strTittle
    
    int���� = rsTemp!length
    
    strSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, strSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str����)
    
    
    strSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure strSQL, strTittle
    str���� = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
    
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Set gobjSquare.objSquareCard = Nothing
     End If
     Set gobjSquare = Nothing
     If Err <> 0 Then Err.Clear: Err = 0
End Sub

Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, Optional strʹ����� As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '����:��ӡ��ʽ(���)
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim lngFormat As Long
    Dim lngFormat1 As Long
    
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    strShareTypeFormat = Trim(zlDatabase.GetPara("�շѷ�Ʊ��ʽ", glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = strʹ����� And lngFormat <> 0 Then
            zl_GetInvoicePrintFormat = lngFormat: Exit Function
        End If
    Next
    zl_GetInvoicePrintFormat = lngFormat1
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional strʹ����� As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '����:int��ӡ��ʽ-��ӡ��ʽ()
    '����:0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    strShareTypeFormat = Trim(zlDatabase.GetPara("�շѷ�Ʊ��ӡ��ʽ", glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ӡ��ʽ1|ʹ�����2,��ӡ��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = strʹ����� Then
            zl_GetInvoicePrintMode = intPrintMode: Exit Function
        End If
    Next
    zl_GetInvoicePrintMode = intPrintMode1
End Function

Public Function zl_GetԤԼ��ʽByID(lng�Һ�ID As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݹҺ�ID��ȡ����ԤԼ��ʽ
    '���:lng�Һ�ID-���˹Һ�ID
    '����:ԤԼ��ʽ
    '����:����
    '����:2012-07-03
    '�����:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strԤԼ��ʽ As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select ԤԼ��ʽ From ���˹Һż�¼ Where ��¼״̬=1 And ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ��ʽ", lng�Һ�ID)
    If rsTemp Is Nothing Then zl_GetԤԼ��ʽByID = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_GetԤԼ��ʽByID = "": Exit Function
    While rsTemp.EOF = False
        strԤԼ��ʽ = Nvl(rsTemp!ԤԼ��ʽ)
        rsTemp.MoveNext
    Wend
    zl_GetԤԼ��ʽByID = strԤԼ��ʽ
End Function

Public Function zl_GetԤԼ��ʽByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݹҺŵ��ݺŻ�ȡ����ԤԼ��ʽ
    '���:strNo-�Һŵ��ݺ�
    '����:ԤԼ��ʽ
    '����:����
    '����:2012-07-03
    '�����:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strԤԼ��ʽ As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select ԤԼ��ʽ From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ��ʽ", strNO)
    If rsTemp Is Nothing Then zl_GetԤԼ��ʽByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_GetԤԼ��ʽByNo = "": Exit Function
    While rsTemp.EOF = False
        strԤԼ��ʽ = Nvl(rsTemp!ԤԼ��ʽ)
        rsTemp.MoveNext
    Wend
    zl_GetԤԼ��ʽByNo = strԤԼ��ʽ
End Function

Public Function zl_Getҽ�ƿ�����(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '����:����ҽ������ID��ȡҽ������
    '���:lngTypeID-ҽ�ƿ�����ID
    '����:���Ͷ���
    '����:����
    '����:2012-07-06
    '�����:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSQL = "" & _
    "       Select ���볤��,������������,�Ƿ�ȱʡ���� " & _
    "       From ҽ�ƿ���� " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�ƿ����", lngTypeId)
    If rsTemp Is Nothing Then zl_Getҽ�ƿ����� = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Getҽ�ƿ����� = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = Nvl(rsTemp!���볤��, "0")
    arr(1) = Nvl(rsTemp!������������, "0")
    arr(2) = Nvl(rsTemp!�Ƿ�ȱʡ����, "0")
    zl_Getҽ�ƿ����� = arr
End Function
Public Function zlReadRegThreeBalance(ByVal strNO As String, _
    ByRef cllBillBalance As Collection, Optional ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������㽻����Ϣ
    '���:strNo-���ݺ�
    '����:��ȡ�ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-08-08 10:10:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngҽ�ƿ����ID As Long, byt���ѿ� As Byte
    Dim objCards As Cards, objTemp As Card
    
    Set cllBillBalance = Nothing
    On Error GoTo errHandle
    '����:51527: and Mod(B.��¼����,10)<>1"
    gstrSQL = _
        "Select b.����id, b.�����id, b.���㿨���, b.����, b.������ˮ��, b.����˵��, b.������λ, d.���ѿ�id" & vbNewLine & _
        "From ������ü�¼ A, ����Ԥ����¼ B, ���˿������¼ D" & vbNewLine & _
        "Where a.����id = b.����id And b.Id = d.����id(+) And a.No = [1]" & vbNewLine & _
        "      And a.��¼���� = 4 And a.��¼״̬ = 1 And Mod(b.��¼����, 10) <> 1" & vbNewLine & _
        "      And (Nvl(b.�����id, 0) <> 0 Or Nvl(b.���㿨���, 0) <> 0)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���㽻����Ϣ", strNO)
    If rsTemp.EOF Then Exit Function
    
    lngҽ�ƿ����ID = IIf(Val(Nvl(rsTemp!�����ID)) > 0, Val(Nvl(rsTemp!�����ID)), Val(Nvl(rsTemp!���㿨���)))
    byt���ѿ� = IIf(Val(Nvl(rsTemp!���㿨���)) <> 0, 1, 0)
    Set objCard = New Card
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID,���ѿ�ID
    Set cllBillBalance = New Collection
    cllBillBalance.Add Array(lngҽ�ƿ����ID, Trim(Nvl(rsTemp!����)), byt���ѿ�, _
        Trim(Nvl(rsTemp!������ˮ��)), Trim(Nvl(rsTemp!����˵��)), strNO, Val(Nvl(rsTemp!����ID)), Val(Nvl(rsTemp!���ѿ�ID))), strNO
    zlReadRegThreeBalance = True
    If gobjSquare.objSquareCard.zlGetCard(lngҽ�ƿ����ID, byt���ѿ� = 1, objCard) = False Then Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegThreeMoney(lng����ID As Long, lngCard����ID As Long, _
    ByVal cllBancel As Collection) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�˷ѵ��������׽��
    '����:���˺�
    '����:2011-08-08 14:43:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng�����ID As Long
    Dim strCardNo As String
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    If lng����ID = 0 And lngCard����ID = 0 Then Exit Function
    If cllBancel Is Nothing Then Exit Function
    lng�����ID = Val(cllBancel(1)(0))
    strCardNo = Trim(cllBancel(1)(1))
    strSQL = "Select sum(nvl(��Ԥ��,0)) as ���ʽ�� From ����Ԥ����¼ Where ����ID in ([1],[2]) and  (�����Id=[3] or ���㿨���=[3]) and mod(��¼����,10)<>1 "
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������׽��", lng����ID, lngCard����ID, lng�����ID, strCardNo)
    If rsTemp.EOF Then Exit Function
    zlGetRegThreeMoney = Val(Nvl(rsTemp!���ʽ��))
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function �Ƿ��Ѿ�ǩԼ(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ҫ�󶨵Ŀ����Ƿ��Ѿ�ǩԼ
    '���:�󶨿���
    '����:����
    '����:2012-08-31 11:32:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng���֤���ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    lng���֤���ID = Getҽ�ƿ����ID("�������֤")
    strSQL = "" & _
    "   Select Count(1) as �Ƿ�ǩԼ From ����ҽ�ƿ���Ϣ Where ����=[1] And �����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ���", strCardNo, lng���֤���ID)
    �Ƿ��Ѿ�ǩԼ = rsTemp!�Ƿ�ǩԼ > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Public Sub AddSQL�󶨿�(ByVal lng����ID As Long, �����ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�󶨿�����
    '���:lng����ID;strCard-�󶨿���;strPassWord-��������
    '����:lngCard����ID-���ѵĽ���ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str�䶯ԭ�� As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str�䶯ԭ�� = "���˹Һŷ���"
          'Zl_ҽ�ƿ��䶯_Insert
          strSQL = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSQL = strSQL & "" & 11 & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSQL = strSQL & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSQL = strSQL & "" & �����ID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSQL = strSQL & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSQL = strSQL & "NULL)"
     zlAddArray cllPro, strSQL
End Sub

Public Function Getҽ�ƿ����ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����ID
    '���:strTypeName ҽ�ƿ��������
    '����:ҽ�ƿ����ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select ID From ҽ�ƿ���� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ����", strTypeName)
    If rsTemp Is Nothing Then Getҽ�ƿ����ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Getҽ�ƿ����ID = 0: Exit Function
    Getҽ�ƿ����ID = rsTemp!id
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiByID(str���� As String, strValue As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������ȡ��ͬ�����µĲ�����Ϣ
    '���:str���ͣ���ѯ�������� strValue ����ֵ
    '����:������Ϣ����
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandl
    strSQL = "" & _
    "   Select ����ID,�����,סԺ��,���￨��,����֤��,�ѱ�,ҽ�Ƹ��ʽ,����,�Ա�,����,��������,�����ص�,���֤��,����֤��,���,ְҵ,����,����,����,����,ѧ��,����״��,��ͥ��ַ,��ͥ�绰,��ͥ��ַ�ʱ�,�໤��," & _
    "   ��ϵ������,��ϵ�˹�ϵ,��ϵ�˵�ַ,��ϵ�˵绰,���ڵ�ַ,���ڵ�ַ�ʱ�,Email,QQ,��ͬ��λID,������λ,��λ�绰,��λ�ʱ�,��λ������,��λ�ʺ�,������,��������,����ʱ��,����״̬,��������,סԺ����,��ǰ����ID,��ǰ����," & _
    "   ��Ժʱ��,��Ժʱ��,��Ժ,IC����,������,ҽ����,����,��ѯ����,�Ǽ�ʱ��,ͣ��ʱ��,����,��ϵ�����֤��,����ģʽ,��������,�ֻ��� " & _
    "   From ������Ϣ " & _
    "   Where " & str���� & "=[1]"
    
    Set GetPatiByID = zlDatabase.OpenSQLRecord(strSQL, "����Һ�", strValue)
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetǩԼ��������(str���֤ As String) As String
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:�������֤��ȡǩԼ���˵�����
'���:str���֤ �������֤��
'����:��������
'����:����
'����:2012-08-31 04:36:33
'�����:53408
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng���֤���ID As Long
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl

    lng���֤���ID = Getҽ�ƿ����ID("�������֤")
    strSQL = "" & _
           "   Select ���� FROM  ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.����ID=B.����ID And B.����=[1] And B.�����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Һ�", str���֤, lng���֤���ID)
    If rsTemp Is Nothing Then GetǩԼ�������� = "": Exit Function
    If rsTemp.RecordCount Then GetǩԼ�������� = "": Exit Function

    GetǩԼ�������� = rsTemp!����
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function


Public Function Bln�ѷ���(str���� As String, lng����� As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡָ�������Ƿ��Ѿ�����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:True :�Ѿ�����;False:δ����
'����:����
'����:2012-10-11 04:36:33
'�����:54390
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select ����ID From ����ҽ�ƿ���Ϣ Where ����=[1]  And �����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Һ�", str����, lng�����)
    Bln�ѷ��� = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!����ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetCardLastChangeType(ByVal str���� As String, ByVal lng����� As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡ�����ı䶯����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:0-δ�ҵ������Ϣ   1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
'����:��⸣
'����:2013-2-4 17:36:33
'�����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select �䶯���" & vbNewLine & _
           "    From (With ҽ�ƿ��䶯 As (Select ����id, ID, �䶯���, �䶯ʱ�� " & vbNewLine & _
           "                              From ����ҽ�ƿ��䶯 Bd" & vbNewLine & _
           "                              Where Bd.���� = [2] And �����id = [1] And ����id = [3])" & vbNewLine & _
           "           Select A.�䶯���" & vbNewLine & _
           "           From ҽ�ƿ��䶯 A, (Select Max(�䶯ʱ��) As �䶯ʱ�� From ҽ�ƿ��䶯 C) B" & vbNewLine & _
           "           Where A.�䶯ʱ�� = B.�䶯ʱ��) A"
    On Error GoTo Errhand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����䶯��Ϣ", lng�����, str����, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!�䶯���))
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Public Function zlGetRegAdvanceMoney(lng����ID As Long, lngCard����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�˷ѵ�Ԥ�����
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If lng����ID = 0 And lngCard����ID = 0 Then Exit Function
    strSQL = "Select sum(nvl(��Ԥ��,0)) as ���ʽ�� From ����Ԥ����¼ Where ����ID in ([1],[2]) and mod(��¼����,10)=1 "
    On Error GoTo Hd
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������׽��", lng����ID, lngCard����ID)
    If rsTemp.EOF Then Exit Function
    zlGetRegAdvanceMoney = Val(Nvl(rsTemp!���ʽ��))
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function zlCheckIsAllowBackSN(ByVal strNO As String, _
    ByVal bln���� As Boolean, Optional ByRef bln���� As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ������˺�(ֻ�����ʲ���)
    '���:strNO-�˺ŵ��ݺ�
    '       bln����-�Ƿ���ʷ���
    '����:bln����-�Ƿ���Ҫ��������ý���
    '����:�����˺ŷ���true,���򷵻�False
    '����:���˺�
    '����:2013-12-26 09:29:02
    '˵��:68991
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long
    On Error GoTo errHandle
    'ֻ��Լ��ʷ��ý��м��
    bln���� = False
    If bln���� = False Then zlCheckIsAllowBackSN = True: Exit Function
    
    strSQL = " " & _
    "   Select Max(ҽ��) As ҽ��, Max(����) As ���� " & _
    "   From (Select 1 As ҽ��, 0 As ���� " & _
    "          From ����ҽ����¼ " & _
    "          Where �Һŵ� =[1] " & _
    "          Union All " & _
    "          Select 0 As ҽ��, 1 As ���� " & _
    "          From ���˹Һż�¼ " & _
    "          Where NO = [1] And ��¼���� = 1 And ��¼״̬ In (1, 3) And ִ��״̬ In (1, 2))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����ʹҺŷ��õ����״̬", strNO)
    
    '1.����Һŵ��Ѿ�������ҵ�����ݣ���ҽ�����ݣ�����Ҳ�������˺�
   If Val(Nvl(rsTemp!ҽ��)) = 1 Then
        MsgBox "ע��:" & vbCrLf & _
                      "       �Һŵ�Ϊ" & strNO & "���Ѿ�������ҽ������,�������˺�!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
         
    '2.����Һŵ��Ѿ�����,�������˺�.
    If Val(Nvl(rsTemp!����)) = 1 Then
        MsgBox "ע��:" & vbCrLf & _
                      "       �Һŵ�Ϊ" & strNO & "���Ѿ��������ɽ���,�������˺�!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    '3.����Һŵ�����Ӧ�ļ��ʵ��Ѿ����ʣ��������˺�;
    
    strSQL = "" & _
    " Select Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0) As δ����, Max(����id) As ����id, Sum(ʵ�ս��) As ʵ�ս��,Max(����ID) as ����ID " & _
    "   From ������ü�¼ " & _
    "   Where NO = [1] And Mod(��¼����, 10) = 4 And nvl(���ʷ���,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����ʹҺŷ��õ����״̬", strNO)
    
    If Val(Nvl(rsTemp!δ����)) = 0 And Val(Nvl(rsTemp!ʵ�ս��)) <> 0 Then
        '֤���Ѿ�������
        MsgBox "ע��:" & vbCrLf & _
                      "       �Һŵ�Ϊ" & strNO & "���Ѿ�������,�������˺�!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!δ����)) = 0 And Val(Nvl(rsTemp!ʵ�ս��)) = 0 And Val(Nvl(rsTemp!����ID)) > 0 Then
        '��Ѻ�,ʵ�ս��δ��,��Ҳ���ܴ��ڽ��ʵ����
        MsgBox "ע��:" & vbCrLf & _
                      "       �Һŵ�Ϊ" & strNO & "���Ѿ�������,�������˺�!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
     '4.�����ǰ�Һŵ������һ����Ч�ĹҺŵ�(���ܴ��ڶ����,����׼���Һ���Ч����)�����˻����ڼ�������ʱ��Ҳ�������˺�
    lng����ID = Val(Nvl(rsTemp!����ID))

    strSQL = " " & _
    "   Select Count(*)  as ����,Max(��ǰ����) as ��ǰ���� " & _
    "   From ( Select  distinct NO,decode(NO,[2],1,0) as ��ǰ���� " & _
    "               From ������ü�¼ " & _
    "               Where ����id = [1] And ��¼���� = 4 And ��¼״̬ = 1 And " & _
    "                         (       (Nvl(�Ӱ��־, 0) = 1 And �Ǽ�ʱ��+0 >= Sysdate-" & gSysPara.Sy_Reg.bytNoDayseMergency & ")  " & _
    "                            Or (Nvl(�Ӱ��־, 0) = 0 And �Ǽ�ʱ��+0 >= Sysdate -" & gSysPara.Sy_Reg.bytNODaysGeneral & "))) "
   Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����ʹҺŷ��õ����״̬", lng����ID, strNO)
    If Val(Nvl(rsTemp!����)) = 1 Then
        If Val(Nvl(rsTemp!��ǰ����)) = 1 Then
            '���һ�ŵ���,��Ҫ����Ƿ��м�������
            strSQL = "" & _
            "   Select  sum(���) as ��� " & _
            "   From ����δ����� " & _
            "   Where ����id = [1] And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 " & _
            "   UNION ALL " & _
            "   Select -1*Sum(ʵ�ս�� ) From ������ü�¼ " & _
            "   Where No=[2] and ��¼����=4 and ��¼״̬=1 and nvl(���ʷ���,0)=1 "
            strSQL = "Select sum(���) as ��� From (" & strSQL & ")"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����ʹҺŷ��õ����״̬", lng����ID, strNO)
            
            If Val(Nvl(rsTemp!���)) <> 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "       �Һŵ�Ϊ" & strNO & "���Ѿ�������,�������˺�!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            '�˹Һ���Ч����ĵ���,�ݲ�����
            bln���� = True
        End If
    End If
    zlCheckIsAllowBackSN = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = zlDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ṹ����ַ�ؼ��е���Ϣ¼���Ƿ���ȷ
    '���:objCtl-�ṹ����ַ�ؼ���lngLen-���Ƴ���
    '����:True-������Ϣ�Ϸ�
    '����:���ϴ�
    '����:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "���ֻ������" & lngLen \ 2 & "������,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "ע��:" & vbCrLf & "   " & objCtl.Tag & "��" & objCtl.CheckNullValue & "��δ����,���顣", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Function ExcPlugInFun(ByVal bytFunc As Byte, ByVal lngRegID As Long, Optional strDoctor As String, Optional strRoom As String, _
                                Optional strNewArrange As String, Optional lngNewArrangeID As Long) As Boolean
    '����:�Һŷ�����ӿ�
    'bytFunc - 0-����;1-����;2-��ɾ���(13-�ָ�����);3-���Ϊ������;4-ǩ��(14-ȡ��ǩ��);5-����(15-ȡ������);6-���˴���
    If gblnPlugin = False Then ExcPlugInFun = True: Exit Function
    
    On Error Resume Next
    ExcPlugInFun = gobjPlugIn.PatiRegTriageCheck(glngSys, glngModul, bytFunc, lngRegID, strDoctor, strRoom, strNewArrange, lngNewArrangeID)
    If Err.Number <> 0 Then
        Call zlPlugInErrH(Err, "PatiRegTriageCheck")
        Err.Clear
        ExcPlugInFun = True
    End If
    On Error GoTo 0
End Function

Public Function Get��Ŀ���(ByVal lng��Ŀid As Long, ByVal strPriceGrade As String) As Double
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = "Select ��Ŀid, Sum(Nvl(�ּ�, 0)) As ���" & vbNewLine & _
            "From (Select /*+cardinality(D,10)*/" & vbNewLine & _
            "        b.�ּ�, d.Column_Value As ��Ŀid" & vbNewLine & _
            "       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list([1])) D" & vbNewLine & _
            "       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And" & vbNewLine & _
            "             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select /*+cardinality(E,10)*/" & vbNewLine & _
            "        b.�ּ� * d.��������, e.Column_Value As ��Ŀid" & vbNewLine & _
            "       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_Str2list([1])) E" & vbNewLine & _
            "       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = e.Column_Value And Sysdate Between b.ִ������ And" & vbNewLine & _
            "             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')))" & vbNewLine & _
            "Group By ��Ŀid"
    
    strSQL = "" & vbNewLine & _
            "Select b.�ּ�" & vbNewLine & _
            "       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C" & vbNewLine & _
            "       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = [1] And Sysdate Between b.ִ������ And" & vbNewLine & _
            "             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = strSQL & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select b.�ּ� * d.��������" & vbNewLine & _
            "       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D" & vbNewLine & _
            "       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = [1] And Sysdate Between b.ִ������ And" & vbNewLine & _
            "             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = "Select Sum(Nvl(�ּ�, 0)) As ��� From (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һ���Ŀ���", lng��Ŀid, strPriceGrade)
    If rsTemp.EOF Then
        Get��Ŀ��� = 0
    Else
        Get��Ŀ��� = Val(Nvl(rsTemp!���))
    End If
End Function

Public Function Get��Ŀ��Ϣ(ByVal str��Ŀids As String, ByVal strPriceGrade As String) As ADODB.Recordset
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = "Select /*+cardinality(D,10)*/" & vbNewLine & _
            "       b.�ּ�, d.Column_Value As ��Ŀid" & vbNewLine & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list([1])) D" & vbNewLine & _
            " Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And" & vbNewLine & _
            "       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select /*+cardinality(E,10)*/" & vbNewLine & _
            "        b.�ּ� * d.��������, e.Column_Value As ��Ŀid" & vbNewLine & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_Str2list([1])) E" & vbNewLine & _
            " Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = e.Column_Value And Sysdate Between b.ִ������ And" & vbNewLine & _
            "       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                    strWherePriceGrade
    strSQL = "Select ��Ŀid, Sum(Nvl(�ּ�, 0)) As ���" & vbNewLine & _
            " From (" & strSQL & ")" & vbNewLine & _
            " Group By ��Ŀid"
    Set Get��Ŀ��Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һ���Ŀ���", str��Ŀids, strPriceGrade)
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'���ܣ��������뷽ʽ��ʽ������
'������intBit=���С��λ��
'����ţ�94552
'˵����VB�Դ���Round�����м����뷨,��ʵ�ʲ�һ�¡���Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If gobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    gintPriceGradeStartType = gobjPublicExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then Exit Sub
    '��ȡվ��۸�ȼ�
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , gstrPriceGrade)
End Sub

Public Function ZlGetBillFormat(ByVal intFormat As Integer) As String
    '���ܣ���ȡƱ�ݸ�ʽ����
    '��Σ�
    '   intFormat - Ʊ�ݸ�ʽ���
    '���أ�Ʊ�ݸ�ʽ������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo ErrHandler
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1111"
    
    If intFormat = 0 Then '��ȱʡƱ�ݸ�ʽ��ʾ
        intFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    strSQL = _
        "Select b.˵��" & vbNewLine & _
        "From zlReports A, zlRPTFMTs B" & vbNewLine & _
        "Where a.Id = b.����id And a.��� = [1] And b.��� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ʽ˵��", strRptName, intFormat)
    If rsTmp.EOF Then Exit Function
    
    ZlGetBillFormat = Nvl(rsTmp!˵��)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CreateRegisterObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ù����ĹҺŶ���
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-04 10:06:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zlPublicExpense.clsRegist")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsRegist)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    If gobjRegist Is Nothing Then Exit Function
    'zlInitCommon(ByVal lngSys As Long,  ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    If gobjRegist.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsRegist)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    CreateRegisterObject = True
End Function

Public Function CheckBillRepeat(ByVal lng����ID As Long, ByVal bytƱ�� As Byte, ByVal strFactNO As String) As Boolean
'���ܣ���ʹ����Ʊ��֮ǰ������Ƿ��ظ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select ����" & vbNewLine & _
        "From Ʊ��ʹ����ϸ" & vbNewLine & _
        "Where ����ID = [1] And Ʊ��=[2] And ����=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID, bytƱ��, strFactNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiValiedCheckByPlugIn(ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByVal strPatiInforXML As String) As Boolean
    '������ҽӿ� PatiValiedCheck ��鲡����Ϣ
    '�����:102230,106686,138602
    '˵����
    '   1.û����Ҳ���ʱ����Ϊ���ͨ��
    '   2.��Ҳ�������PatiValiedCheck�ӿڣ�Ҳ��Ϊ���ͨ��
    '   3.����������ʶ���˳ɹ�����ã�δ���������ڱ�������ǰ����
    
    If CreatePlugInOK(lngModule) = False Then PatiValiedCheckByPlugIn = True: Exit Function
    If gobjPlugIn Is Nothing Then PatiValiedCheckByPlugIn = True: Exit Function
    
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    '���ܣ���鵱ǰ�����Ƿ���ָ�������ⲡ��
    '���أ�trueʱ�������������Falseʱ���������
    '������
    '      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '      lngType �������ͣ�1������Һţ�2��סԺ��Ժ��3�������շѣ�4��סԺ���ʡ�
    '      lngPatiID-����ID: �½����ģ�Ϊ0,�����뽨������ID
    '      lngPageID-��ҳID: �½����ģ�Ϊ0,�����뽨����ҳID(סԺ������ҳID) ����˵������ lngType=4 ʱ�Ŵ��� lngPageID����������0
    '      strPatiInforXML-������Ϣ:���δ�������˴��룬"�������Ա����䣬�������ڣ�ҽ���ţ����֤�ţ�ҽ������"���������� ��ʽ:2016-11-11 12:12:12
    '                      �̶���ʽ��<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH><YSXM></YSXM>
    '                   �������˴��룬"ҽ������"(106686),��ʽ��<YSXM></YSXM>
    '      strReserve=��������,������չʹ��
    If gobjPlugIn.PatiValiedCheck(glngSys, lngModule, 1, lng����ID, 0, strPatiInforXML) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                PatiValiedCheckByPlugIn = True: Exit Function
            End If
            Call zlPlugInErrH(Err, "PatiValiedCheck")
        End If
        Exit Function
    End If
    PatiValiedCheckByPlugIn = True
End Function

Public Sub InitAddressLength()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ��ͥ��ַ, ���ڵ�ַ, �����ص�, ��ϵ�˵�ַ From ������Ϣ Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ַ����")
    If Not rsTmp.EOF Then
        glngMax��ͥ��ַ = rsTmp.Fields("��ͥ��ַ").DefinedSize
        glngMax���ڵ�ַ = rsTmp.Fields("���ڵ�ַ").DefinedSize
        glngMax�����ص� = rsTmp.Fields("�����ص�").DefinedSize
        glngMax��ϵ�˵�ַ = rsTmp.Fields("��ϵ�˵�ַ").DefinedSize
    End If
    If glngMax��ͥ��ַ = 0 Then glngMax��ͥ��ַ = 100: If glngMax���ڵ�ַ = 0 Then glngMax���ڵ�ַ = 100
    If glngMax�����ص� = 0 Then glngMax�����ص� = 100: If glngMax��ϵ�˵�ַ = 0 Then glngMax��ϵ�˵�ַ = 100
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Get�������(ByVal lngSys As Long, ByVal lngModul As Long, ByVal strPrivs As String) As String
    Dim rsTmp As ADODB.Recordset
    Dim str������� As String, strTmp As String
    On Error GoTo errH
    str������� = zlDatabase.GetPara("�������", lngSys, lngModul)
    If InStr(strPrivs, "���п���") = 0 Then
        Set rsTmp = GetDepartments("'�ٴ�'", "1,3", InStr(strPrivs, "���п���") = 0)
        If str������� = "" Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & Nvl(rsTmp!id)
                rsTmp.MoveNext
            Loop
        Else
            Do While Not rsTmp.EOF
                If InStr("," & str������� & ",", "," & Nvl(rsTmp!id) & ",") > 0 Then
                    strTmp = strTmp & "," & Nvl(rsTmp!id)
                End If
                rsTmp.MoveNext
            Loop
        End If
        If strTmp <> "" Then
            str������� = Mid(strTmp, 2)
        Else
            str������� = "0"
        End If
    End If
    Get������� = str�������
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Get������� = "0"
End Function
