Attribute VB_Name = "mdlRegEventSelf"
Option Explicit

Public Function GetRoom(str�ű� As String) As String
'���ന2003��1��6�յ��������ú���
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    gstrSQL = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ҺŰ��� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ҺŰ���", str�ű�)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        'ָ������
        Dim lng�ű�ID As Long
        lng�ű�ID = Val(Nvl(rsTmp!ID))
        Set rsTmp = GetRs�Һ�����
        If rsTmp Is Nothing Then
            gstrSQL = "Select �ű�ID,�������� From �ҺŰ������� Where �ű�ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ָ������", lng�ű�ID)
        End If
        rsTmp.Filter = "�ű�ID=" & lng�ű�ID
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
        rsTmp.Filter = 0
        
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        gstrSQL = _
        " Select ��������,Sum(NUM) as NUM  " & _
        " From ( Select ��������,0 as NUM From �ҺŰ������� Where �ű�ID=[1]" & _
        "        Union ALL" & _
        "       Select ��ҩ����,Count(��ҩ����) as NUM From ������ü�¼" & _
        "       Where ��¼����=4 And ��¼״̬=1 And ���=1 And Nvl(ִ��״̬,0)=0" & _
        "               And �Ǽ�ʱ�� Between Trunc(Sysdate) And Sysdate And ���㵥λ=[2]" & _
        "               And ��ҩ���� IN(Select �������� From �ҺŰ������� Where �ű�ID=[1])" & _
        "       Group by ��ҩ����) " & _
        " Group by �������� " & _
        " Order by Num"
        
       Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������������", Val(Nvl(rsTmp!ID)), str�ű�)
       If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        gstrSQL = "Select �ű�ID,��������,��ǰ���� From �ҺŰ������� Where �ű�ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ƽ����������", Val(Nvl(rsTmp!ID)))
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

Public Function GetRegistPrice(ByVal lng��ĿID) As Variant
    '******************************************************************************************************************
    '���ܣ�����ָ���Һ����ͣ���ָ��ʱ��ļ۸��ά�����У����顣
    '   ��һ��Ϊ�۸񣬵ڶ��б�ʾ������ĿID����������д������Ŀ,������Ϊ���㵥λ,������Ϊ����,������Ϊ�շ�ϸĿID,������(�۸����),�ڰ���(��������)
    '������lng��ĿID=�Һ���ĿID(�շ�ϸĿID)
    '���أ�����
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim aryTmp(), i As Integer
    Dim int���� As Integer, int���� As Integer, lng������ĿID As Long
    On Error GoTo errH

    gstrSQL = "Select 1 as ����,A.���,A.ID as ��ĿID,A.���㵥λ,B.������ĿID,1 as ����,C.�վݷ�Ŀ,B.�ּ�" & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=[1] " & _
        " And ((To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS') Between To_Char(B.ִ������,'YYYY-MM-DD HH24:MI:SS') And To_Char(B.��ֹ����,'YYYY-MM-DD HH24:MI:SS')) or (To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS')>=To_Char(B.ִ������,'YYYY-MM-DD HH24:MI:SS') And (B.��ֹ���� is NULL Or B.��ֹ����=To_Date('3000-01-01','YYYY-MM-DD'))))"
    gstrSQL = gstrSQL & " Union ALL " & _
        "Select 2 as ����,A.���,A.ID as ��ĿID,A.���㵥λ,C.ID as ������ĿID,D.�������� as ����,C.�վݷ�Ŀ,B.�ּ�" & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շѴ�����Ŀ D" & _
        " Where B.�շ�ϸĿID=A.ID And B.������ĿID=C.ID And A.ID=D.����ID And D.����ID=[1]" & _
        "        And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        ""
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng��ĿID)
    If rs.EOF Then
        GetRegistPrice = Null
    Else
        ReDim aryTmp(rs.RecordCount - 1, 8)
        int���� = 0: lng������ĿID = 0
        For i = 1 To rs.RecordCount
            If lng��ĿID = Val(Nvl(rs!��ĿID)) Then
                If lng������ĿID <> Val(Nvl(rs!������ĿID)) Then
                    int���� = 1: int���� = i:
                     lng������ĿID = Val(Nvl(rs!������ĿID))
                End If
            Else
                int���� = 2
            End If
            
            aryTmp(i - 1, 0) = zlCommFun.Nvl(rs("�ּ�").Value, 0)
            aryTmp(i - 1, 1) = zlCommFun.Nvl(rs("������ĿID").Value, 0)
            aryTmp(i - 1, 2) = zlCommFun.Nvl(rs("�վݷ�Ŀ").Value)
            aryTmp(i - 1, 3) = zlCommFun.Nvl(rs("���㵥λ").Value)
            aryTmp(i - 1, 4) = zlCommFun.Nvl(rs("����").Value)
            aryTmp(i - 1, 5) = zlCommFun.Nvl(rs("��ĿID").Value)
            aryTmp(i - 1, 6) = IIf(int���� = 1 And i <> int����, int����, 0)
            aryTmp(i - 1, 7) = IIf(int���� = 2 And i <> int����, int����, 0)
            rs.MoveNext
        Next
        GetRegistPrice = aryTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetRegistPrice = Null
End Function

Public Function ActualMoney(ByVal str�ѱ� As String, ByVal lng����ID As Long, ByVal curӦ�� As Currency) As Currency
'���ܣ�����ָ���ķѱ��������Ŀ,����ָ������ʵ���տ���
'������
'   str�ѱ�   ���ѱ�
'   lng����ID  ��������ĿID
'   curӦ�գ�Ӧ�ս��ֵ
'���أ�ʵ��Ӧ�յĽ��
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    gstrSQL = "Select ʵ�ձ��� From �ѱ���ϸ " & _
        " Where �ѱ�=[1] And ������ĿID= [2] " & _
        " And ABS([3]) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ" & _
        " Order by Ӧ�ն�βֵ"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", str�ѱ�, CStr(lng����ID), CStr(curӦ��))
    
    If rsTmp.EOF Then
        ActualMoney = curӦ��
    Else
        ActualMoney = curӦ�� * rsTmp!ʵ�ձ��� / 100
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim rsPar As New ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    
    '������ʾ��ʽ
    gblnShowCard = Not (-Abs(Val(zlDatabase.GetPara(12, glngSys))))
    
    '�Һ�Ʊ�ݺ��볤��
    strTmp = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbyt�Һ� = Val(Split(strTmp, "|")(3))
    
    gstrSQL = "Select ���ų��� From ҽ�ƿ���� where ����='���￨' and nvl(�Ƿ�̶�,0)=1"
    gbytCardNOLen = 7
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���￨���ų���")
    If Not rsTemp.EOF Then
        gbytCardNOLen = Val(Nvl(rsTemp!���ų���))
    End If
    'Ʊ���ϸ����
    strTmp = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill�Һ� = (Mid(strTmp, 4, 1) = "1")
    
    '�ձ�ͳ��ʱ������
    gblnDailyTime = zlDatabase.GetPara(22, glngSys, , 0)
     
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Sub InitLocPar()
'���ܣ���ʼ�����ñ�������
    '��ѡ����ֵ
    '���ع���Ԥ��Ʊ������ID
    
    glng�Һ�ID = Val(zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, 1111, 0))

    If glng�Һ�ID > 0 Then
        If Not ExistBill(glng�Һ�ID, 4) Then
            
            Call zlDatabase.SetPara("���ùҺ�Ʊ������", 0, glngSys, 1111)

            glng�Һ�ID = 0
        End If
    End If
    
End Sub
Public Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'���ܣ��ж��Ƿ����ָ����Ʊ������
    Dim rsTmp As New ADODB.Recordset

    On Error GoTo errH

    gstrSQL = "Select ID From Ʊ�����ü�¼ Where ID= [1] And Ʊ��= [2] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lngID, bytKind)
    
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNext�ű�() As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    gstrSQL = "Select Max(����) as ���� From �ҺŰ��� Where Length(����)=(Select Max(Length(����)) From �ҺŰ���)"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf")
    
    If Not rsTmp.EOF Then GetNext�ű� = IncStr(IIf(IsNull(rsTmp!����), "", rsTmp!����))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, Optional ByVal strBill As String) As Long
'���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
'������bytKind=Ʊ��
'      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
'      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
'˵����
'    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
'    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
'    3.���ж�������ʱ,ȱʡ���ٵ�����,��������ԭ��ȡ
'���أ�
'      ������Ʊ������ID>0
'      0=ʧ��
'      -1:û������(�����δ����)��Ҳû�й���(δ����)
'      -2:���õĹ���������
'      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
     '������ʣ�������Ʊ�ݼ�
    gstrSQL = _
        " Select " & zlGetFeeFields("Ʊ�����ü�¼") & " From Ʊ�����ü�¼ Where Ʊ��=[1]" & _
        " And ʹ�÷�ʽ=1 And ʣ������>0 And ������=[2]" & _
        " Order by ʣ������,�Ǽ�ʱ��"
        
    Set rsSelf = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", bytKind, UserInfo.����)
    
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        CheckUsedBill = rsSelf!ID
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        gstrSQL = "Select " & zlGetFeeFields("Ʊ�����ü�¼") & " From Ʊ�����ü�¼ Where Ʊ��=[1] And ID= [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", bytKind, lng����ID)
        If rsTmp!ʹ�÷�ʽ = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                CheckUsedBill = rsSelf!ID
            Else
                'û������ȡ����
                If rsTmp!ʣ������ = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                CheckUsedBill = rsTmp!ID
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If rsTmp!ʣ������ > 0 Then
                '��ʣ��
                CheckUsedBill = rsTmp!ID
            Else
                '������ʣ�������
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                CheckUsedBill = rsSelf!ID
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
                CheckUsedBill = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!��ʼ����) And UCase(strBill) <= UCase(rsTmp!��ֹ����) And Len(strBill) = Len(rsTmp!��ʼ����)) Then
                CheckUsedBill = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & CheckUsedBill
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                CheckUsedBill = -3
                rsSelf.Filter = "ID<>" & CheckUsedBill
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)))) <> UCase(IIf(IsNull(rsSelf!ǰ׺�ı�), "", rsSelf!ǰ׺�ı�)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!��ʼ����) And UCase(strBill) <= UCase(rsSelf!��ֹ����) And Len(strBill) = Len(rsSelf!��ʼ����)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then CheckUsedBill = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function


Public Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵������ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
    Dim rsTmp As New ADODB.Recordset
    Dim strBill As String
    
    On Error GoTo errH
    
    gstrSQL = "Select ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ���� From Ʊ�����ü�¼ Where ʣ������>0 And ID=[1] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng����ID)
    
    If rsTmp.EOF Then Exit Function
    
    If IsNull(rsTmp!��ǰ����) Then
        strBill = UCase(rsTmp!��ʼ����)
    Else
        strBill = UCase(IncStr(rsTmp!��ǰ����))
    End If
    '��鷶Χ
    If Left(strBill, Len(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�))) <> UCase(IIf(IsNull(rsTmp!ǰ׺�ı�), "", rsTmp!ǰ׺�ı�)) Then
        Exit Function
    ElseIf Not (strBill >= UCase(rsTmp!��ʼ����) And strBill <= UCase(rsTmp!��ֹ����)) Then
        Exit Function
    End If
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte

    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function


'����Ϊ��ӵĺ���
Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, _
    Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
    Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
    Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean = True, _
    Optional Ratio As Single = 1) As Boolean
'���ܣ���ָ���豸�ϰ�ָ����ʽ��������ֻ�ͼ��
'������
'   Dev=����豸,ΪPrinter��PictureBox����
'   Data=�������,Ϊ����(x)���ַ���("xxx")��ͼ��(stdPicture)���ַ���������vbCrLf,��Data����Ϊ������ʱ,��ʾ�������
'   TW,TH=������޶���Χ,���������Χ���Զ�ȡ������С,Ϊ0ʱ��Ч
'   Border=�߿���,��������,"1111"��ʾȫ��
'   Align=���ֶ���,0=��,1=��,2=��,��ˮƽ���뼰��ֱ����
'   Warp=���������Ϊ�ַ���ʱ,��ʾ�Ƿ��Զ����С����Զ�����ʱ,�����ݲ������
'   Ratio=�������,������,���궼��Ӱ��,ȱʡΪ1(100%)
'˵����1.��ʹ�øú���֮ǰ,Ӧ��û�иı��豸����ͼ��ʼֵ
'      2.�����λ���λ���ڱ��������Χ�����Ͻ�
    Dim i As Long, Text As String, arrText() As String
    Dim LINE_W As Integer, blnW As Boolean, blnH As Boolean
    
    On Error GoTo errH
    
    DrawCell = True
    
    '��Χ�޶�
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '����
        Else
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF 'ʵ�ľ���(����)
        End If
    ElseIf TypeName(Data) = "String" Then
        '����
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.Name = "����"
            Font.Size = 9
        End If
        'ǧ��Ҫ��Set Dev.Font=Font,��֪Ϊ��,�õ���ByVal
        Dev.Font.Name = Font.Name
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        
        '�����ź���������������,�ж�ʱ��ԭʼ��СΪ׼
        If H >= Dev.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True '�߶��Ƿ���(�ӻس�����һ�и߶�)
        If W >= Dev.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0 '����Ƿ���(�ӻس���Ϊ������,�Ա����)
        
        '����
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        
        '�������
        Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        
        Dev.ForeColor = ForeColor
        '�������(�߿�֮���ٸ�һ��)
        '�����߶ȷ�Χ�����
        If blnH Then
            If blnW Then
                Select Case HAlign
                    Case 0
                        Dev.CurrentX = X + LINE_W
                    Case 1
                        Dev.CurrentX = X + (W - Dev.TextWidth(Data)) / 2
                    Case 2
                        Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Data)
                End Select
                Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Data)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Data)
                End Select
                Dev.Print Data
            Else
                If Not Warp Then
                    '���Զ�����ʱ�����ֲ����
                    For i = 1 To Len(Data)
                        If Dev.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next
                    Select Case HAlign
                        Case 0
                            Dev.CurrentX = X + LINE_W
                        Case 1
                            Dev.CurrentX = X + (W - Dev.TextWidth(Text)) / 2
                        Case 2
                            Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Text)
                    End Select
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + LINE_W
                        Case 1
                            Dev.CurrentY = Y + (H - Dev.TextHeight(Text)) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Text)
                    End Select
                    '�����ȡ����
                    Dev.Print Text
                Else
                    '������ֳɶ���(�ڿ�߷�Χ��)
                    ReDim arrText(0) '�ڴ�,��һ�в����ܳ���
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '���г������˳�,���߲��ݲ����
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '���г������˳�,���߲��ݲ����
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '�п���һ��һ���ַ���ȶ�����
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '�����ʼ����
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + LINE_W
                        Case 1
                            Dev.CurrentY = Y + (H - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '�������
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                            Case 0
                                Dev.CurrentX = X + LINE_W
                            Case 1
                                Dev.CurrentX = X + (W - Dev.TextWidth(arrText(i))) / 2
                            Case 2
                                Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(arrText(i))
                        End Select
                        Dev.Print arrText(i)
                    Next
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        
        'ͼ��(�߿�֮��)
        Dev.PaintPicture Data, X + 15, Y + 15, W - LINE_W, H - LINE_W
    End If
    
    If TypeName(Data) <> "Integer" Then
        '�����߿�
        If Mid(Border, 1, 1) Then Dev.Line (X, Y)-(X + W, Y), BorderColor
        If Mid(Border, 2, 1) Then Dev.Line (X, Y + H)-(X + W, Y + H), BorderColor
        If Mid(Border, 3, 1) Then Dev.Line (X, Y)-(X, Y + H), BorderColor
        If Mid(Border, 4, 1) Then Dev.Line (X + W, Y)-(X + W, Y + H), BorderColor
    End If
    Exit Function
errH:
    DrawCell = False
End Function

