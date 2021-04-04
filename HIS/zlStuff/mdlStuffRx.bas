Attribute VB_Name = "mdlStuffRx"
Option Explicit
'��������ṹ
Private Type TYPE_Para
    ���ĵ�λ As Integer
    �������� As String
    �շѵ��� As Integer
    bln����δ�շѵ��ݷ��� As Boolean
End Type

Private Enum mFindType
    ���ݺ� = 0
    ����� = 1
    ���� = 2
    ���֤ = 3
    IC�� = 4
    ҽ���� = 5
    סԺ�� = 6
End Enum

Private T_Para As TYPE_Para
Private mlngModule As Long

Private mstrOracleMoneyForamt As String

Public Sub GetPara(ByVal lngModule As Long)
'��ȡ���ز���
    On Error GoTo errHandle
    With T_Para
        .���ĵ�λ = Val(zlDataBase.GetPara("���ĵ�λ", glngSys, lngModule, "0"))
        .�շѵ��� = zlDataBase.GetPara("�շѴ�����ʾ��ʽ", glngSys, lngModule, "0")
        .�������� = zlDataBase.GetPara("��ѯҵ������", glngSys, lngModule, "0")
        .bln����δ�շѵ��ݷ��� = zlDataBase.GetPara("����δ�շѵ����ﻮ�۴�������", glngSys, lngModule, "0")
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetCheckPara(ByVal lng����ID As Long) As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:
    '����:
    '����:0-����飬1-��飬�������ѣ�2-�����ֹ����
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"

    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "GetCheckPara_��������", lng����ID)
    With rsTemp
        If Not .EOF Then
            GetCheckPara = Nvl(!�����, 0)
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_GetDept(ByVal strPrivs As String) As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� || '-' || a.���� As ���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       AND b.���� ='W' " & _
        "       AND a.id = c.����id " & _
        "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(strPrivs, "���в���") <> 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])") & _
        " Order by a.���� || '-' || a.����"
    
    Set Stuff_GetDept = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ŀⷿ_Stuff_GetDept", UserInfo.Id, gstrNodeNo)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_GetPrePeople(ByVal strKey As String) As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "" & _
        "   Select distinct a.��� as ����,A.���� As ����,����" & _
        "   From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
        "   Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id " & _
        "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND B.����id in (Select ����ID From ������Ա where ��Աid=[2] ) "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "    And  ((A.����) like [1] or  A.���  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    
    gstrSQL = "Select rownum as ID,a.* from (" & gstrSQL & ") A" & _
        "   ORDER BY ���� "
        
    Set Stuff_GetPrePeople = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ������Ա_Stuff_GetPrePeople", UserInfo.Id, gstrNodeNo)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxValied(ByVal strPrivs As String, ByVal intType As Integer, ByVal lng�ⷿID As Long, ByVal int���� As Integer, ByVal strNo As String, ByVal rsData As Recordset) As Boolean
'��鵱ǰ�����Ƿ���Խ��з��ϲ���
'����ֵΪBoolean���ͣ�true-���Է��ϣ�false-���ܷ���
'rsTemp���������뱾�η��ϵ�����
'����;strPrivs-Ȩ���ַ���
'strIDS-�շ�id��
'intType-ҵ�����ͣ�1-���ϣ�2-����
    Dim strTemp As String
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim int����� As Integer
    
    On Error GoTo errHandle
    If strNo = "" Then Exit Function
    
'1 ��鵥���Ƿ����
    gstrSQL = " Select 1 From ҩƷ�շ���¼" & _
             " Where No=[1] And (�ⷿID=[3] Or �ⷿID Is NULL) And ����=[2]"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxValied_��鵥���Ƿ����", strNo, int����, lng�ⷿID)
    
    If rsTemp.EOF Then
        MsgBox "�õ��ݲ����ڣ����鵥����Ϣ��"
        Stuff_RxValied = False
        Exit Function
    End If
    
'2 ��鵥���Ƿ��Ѿ���������Ӧ�Ĳ���
    If intType = 0 Then
        gstrSQL = " Select ����,���շ� From δ��ҩƷ��¼" & _
                 " Where No=[1] And (�ⷿID=[3] Or �ⷿID Is NULL) And ����=[2]"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxValied_��鵥���Ƿ����", strNo, int����, lng�ⷿID)
        
        If rsTemp.EOF Then
            MsgBox "�õ����Ѿ���ҩ", vbInformation, gstrSysName
            Stuff_RxValied = False
            Exit Function
        End If
        
        '3 ��û�й�ѡ����"�����δ�շѵ����ﻮ�۴�������"����鵱ǰ�����Ƿ��Ѿ��շ�
        If rsTemp!���� = 8 And rsTemp!���շ� = 0 And T_Para.bln����δ�շѵ��ݷ��� = False Then
            MsgBox "�õ��ݻ�δ�շѣ�����ִ�з��ϲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If


'4 ��鵥�ݵ����ĺ͵�ǰ���ϲ����Ƿ������˴洢�ⷿ
    
'5 ���ݲ���"�����"����鵱ǰ�Ŀ���Ƿ����㵱ǰ���ݵ���������
    Set rsTemp = GetMatStock(lng�ⷿID)
    int����� = GetCheckPara(lng�ⷿID)
    
    rsData.MoveFirst
    Do While Not rsData.EOF
        rsTemp.Filter = "�շ�ϸĿid=" & rsData!����ID
        
        If rsTemp.EOF Then MsgBox rsData!�������� & "�ò���δ���ô洢�ⷿ", vbInformation, gstrSysName
        
        If intType = 0 Then
            If int����� <> 0 Then
                If Not LocaleStockData(rsData!����, lng�ⷿID, rsData!����ID, rsData!����, rsData!���) Then
                    If int����� = 1 Then
                        If MsgBox("��ǰ��治���Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    ElseIf int����� = 2 Then
                        MsgBox "��ǰ��治���ֹ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName
                    End If
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    
'6 ��鴦���Ƿ��Ѿ�����,�����Ƿ��Ѿ���Ժ���ٶ�Ȩ�޽�����صļ��
    rsData.MoveFirst
    Call Stuff_Check��Ժ����(strPrivs, int����, strNo, rsData!��¼����, rsData!�����־)
    Call Stuff_Check���ʴ���(strPrivs, int����, strNo, rsData!��¼����, rsData!�����־)

    Stuff_RxValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function LocaleStockData(ByVal lngʵ������ As Long, _
    ByVal lng���ϲ���ID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, Optional ByRef lng��� As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ָ������ָ�����Ŀ���Ƿ����
    '���:rsStock-ָ�����Ŀ������(����Ϊ�ռ�¼),�����Զ���չ
    '     lng���ϲ���ID-���ϲ���id
    '     lng����id-����id
    '     lng����-����
    '
    '����:lng���-���ؿ������
    '����:�ɹ�,��ʾ�ҵ�,�����ʾδ�ҵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim dbl��� As Double
    LocaleStockData = False
    
    err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "" & _
    " Select nvl(F.�Ƿ���,0) ���,nvl(A.ʵ������,0) ����" & _
    " From �������� B,�շ���ĿĿ¼ F," & _
    "      (Select A.ҩƷid as ����ID,a.ʵ������ From ҩƷ��� A Where ����=1 And �ⷿID=[1] And ҩƷID=[2] And nvl(����,0)=[3]) A" & _
    " Where B.����ID=F.ID And B.����ID=A.����ID(+) And B.����ID=[2] "
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "LocaleStockData", lng���ϲ���ID, lng����ID, lng����)
    
    dbl��� = Val(Nvl(rsTemp!����))
    
    If lngʵ������ > dbl��� Then LocaleStockData = False

    LocaleStockData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMatStock(ByVal lng�ⷿID As Long) As Recordset
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1] "
    Set GetMatStock = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ�洢�ⷿ", lng�ⷿID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Stuff_RxRefReturnDetail(ByVal int���� As Integer, ByVal strNo As String, ByVal lng�ⷿID As Long, ByVal int��¼״̬ As Integer) As Recordset
'��ȡָ���ѷ��ϵ��ݵ���ϸ��Ϣ
'��������ָ������ϸ��¼
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim str�������� As String
    
    On Error GoTo errHandle
    
    Select Case T_Para.���ĵ�λ
    Case 0  'ɢװ��λ
         strFields = "X.���㵥λ ��λ,1 as ����ϵ��, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,"
    End Select
    
 
    '��ȡ�ѷ��ϻ����ϵĽ��
    strTable = " " & _
    "   Select A.ID, A.NO, A.����, A.���, A.ҩƷid, A.����id, A.����, A.����, A.Ч��, Nvl(A.����, 0) ����, " & _
    "          Nvl(A.����, 1) ����, A.ʵ������ ʵ������, (A.ʵ������ - B.�ѷ�����) ��������, B.�ѷ����� �ѷ�����, " & _
    "          A.��¼״̬, A.���ۼ�, A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����id, A.�ⷿid, " & _
    "          A.����, Decode(Nvl(A.������, ''), '', '', DECODE(mod(a.��¼״̬,3),2,'(��)','(��)') || A.������) ������, H.ҽ�����, " & _
    "          H.��� As �������,H.������ as ����ҽ��,H.����,H.����id,H.��¼����,H.�����־,H.��ʶ��,'' ����,1 �ɲ���" & _
    "   From ҩƷ�շ���¼ A, ������ü�¼ H,������Ϣ H1, " & _
    "        (Select A.NO, A.����, A.ҩƷid, A.���, Sum(Nvl(A.����, 1) * A.ʵ������) �ѷ����� " & _
    "          From ҩƷ�շ���¼ A " & _
    "          Where A.����� Is Not Null And A.�ⷿid + 0 = [3] And A.NO=[2] " & _
    "          Group By A.NO, A.����, A.ҩƷid, A.���) B " & _
    "   Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.���  " & _
    "         And A.����� Is Not Null And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0)  " & _
    "         And A.����id = H.ID And H.����ID=H1.����id(+) "
    

    '�嵥��ʾÿ�ʲ�������
    strTable = strTable & strWhere1
    If blnHistory Then
        strTable = AnalyseHistorySQL(strTable, "1 �ɲ���", "-99 �ɲ���")
    End If
    
    strTable1 = " Union All " & _
    "     Select A.ID, A.NO, A.����, A.���, A.ҩƷid, A.����id, A.����, A.����, A.Ч��, Nvl(A.����, 0), Nvl(A.����, 1) ����, " & _
    "            A.ʵ������ ʵ������, 0 ������, 0 �ѷ�����, A.��¼״̬, A.���ۼ�, A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, " & _
    "            A.�������, A.�Է�����id, A.�ⷿid, " & _
    "            A.����, " & _
    "            Decode(Nvl(A.������, ''), '', '',Decode(A.��¼״̬, 2,'(��)', '(��)' )|| A.������) ������,H.ҽ�����, " & _
    "          H.��� As �������,H.������ as ����ҽ��,H.����,H.����id,H.��¼����,H.�����־,H.��ʶ��,'' ����, Decode(A.��¼״̬, 1, 1,Mod(A.��¼״̬, 3) + 1) �ɲ��� " & _
    "     From ҩƷ�շ���¼ A, ������ü�¼ H ,������Ϣ H1" & _
    "     Where A.����id=H.id And H.����id=H1.����ID(+) and A.����� Is Not Null And Not (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0) And A.�ⷿid + 0 = [3] "
    
    If blnHistory Then
        '��ʷ���ݣ����ܲ���
        strTable1 = AnalyseHistorySQL(strTable1, "Decode(A.��¼״̬, 1, 1,Mod(A.��¼״̬, 3) + 1) �ɲ���", "-99 �ɲ���")
    End If
    
    strTable = strTable & vbCrLf & strTable1
    gstrSQL = " " & _
    "     Select /*+ Rule*/ Distinct S.ID, S.����, S.ҩƷid ����id, S.NO, S.���, S.����, P.���� ����, S.��¼����,S.�����־, S.��ʶ��, S.����id, S.����, " & _
    "                     S.����,M.�Ա�,M.����, '[' || X.���� || ']' || X.���� ��������, Nvl(D.���÷���, 0) ����, X.���, " & strFields & _
    "                     S.���� ��, S.ʵ������/" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & " ����, S.��������/" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & " ��������, S.�ѷ�����/" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & " ׼����, " & _
    "                     Decode(S.����, Null, '', S.����)  ����, " & _
    "                     Nvl(S.����, 0) ����, S.Ч��, S.���ۼ�*" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & " ����, S.���۽�� ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, " & _
    "                     To_Char(S.�������, 'YYYY-MM-DD HH24:MI:SS') ����ʱ��, S.�����, S.�������, �ɲ���, S.ҽ�����, " & _
    "                     I.���㵥λ, Nvl(S.����, Nvl(X.����, '')) ����, Nvl(M.�����, -1) �����, " & _
    "                     Nvl(S.ҽ�����, -1) ҽ��id, S.������, '' �ⷿ��λ, Z.���� As ������,S.��¼״̬,s.����ҽ�� " & _
    "     From (" & strTable & ") S, ���ű� P, �������� D, �շ���ĿĿ¼ X, " & _
    "          �շ���Ŀ���� A, ������ĿĿ¼ I, ����ҽ����¼ M, ������Ŀ���� Z " & _
    "     Where S.ҩƷid = D.����id And D.����id = X.ID And S.�Է�����id + 0 = P.ID And D.����id = I.ID And " & _
    "           S.ҽ����� = M.ID(+) And D.����id = Z.������Ŀid(+) And Z.����(+) = 2 And D.����id = A.�շ�ϸĿid(+) And " & _
    "           A.����(+) = 3 And  S.���� =[1] and S.NO=[2] And S.����� Is Not Null And S.��¼״̬=[4] "
    
     '����
     str���� = Replace(gstrSQL, "H.���˲���ID", "H.��������ID")
     gstrSQL = Replace(gstrSQL, "'' ����", "H.����")
     gstrSQL = str���� & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")

    gstrSQL = gstrSQL & " Order By NO, ����, �������"
    
    Set Stuff_RxRefReturnDetail = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxRefReturnDetail", _
        int����, _
        strNo, _
        lng�ⷿID, _
        int��¼״̬)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional strԭ�� As String = "", Optional str�ִ� As String = "") As String
    '������ʷ���ݵ�SQL���
    Dim strTemp As String
    strTemp = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
    strTemp = Replace(strTemp, "������ü�¼", "H������ü�¼")
    strTemp = Replace(strTemp, "סԺ���ü�¼", "HסԺ���ü�¼")
    If strԭ�� <> "" Then
        strTemp = Replace(strTemp, strԭ��, str�ִ�)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function

Public Function Stuff_RxRefSendDetail(ByVal int���� As Integer, ByVal strNo As String, ByVal lng�ⷿID As Long) As Recordset
'��ȡָ�������ϵ��ݵ���ϸ��Ϣ
'��������ָ������ϸ��¼
Dim lngRow As Long, strWhere As String, strFields As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�������� As String
    Dim strסԺ As String
    
    On Error GoTo errHandle

    If T_Para.���ĵ�λ = 0 Then
        strFields = "x.���㵥λ as ��λ,1 as ����ϵ��,"
    Else
        strFields = "d.��װ��λ as ��λ,d.����ϵ��,"
    End If
    
    gstrSQL = "" & _
        "      Select Distinct s.Id, s.ҩƷid AS ����ID, Nvl(n.���շ�, 0) ���շ�, p.���� ����, s.��ҩ�� AS ������ ,S.����ID, c.������ ����ҽ��, " & _
        "          c.����Ա���� �����, s.����, Nvl(s.����, 0) ����, s.No, s.���, nvl(c.����id,0) as ����ID, '' ����, c.����,m.�Ա�,m.����, " & _
        "          c.��ʶ��, c.����Ա����, '[' || x.���� || ']' || x.���� ��������, s.���� ��, (s.ʵ������/" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & ") ����, " & _
        "          Nvl(d.���÷���, 0) ����, x.���, c.�Ǽ�ʱ��," & strFields & _
        "          s.���ۼ�*" & IIf(T_Para.���ĵ�λ = 0, 1, "d.����ϵ��") & " ����, s.���۽�� ���, s.����, s.Ƶ��, s.�÷�, s.ժҪ ˵��, " & _
        "          Decode(s.����, Null, '', s.����) || Decode(s.����, Null, '', 0, '', '(' || s.���� || ')') ����, " & _
        "          Nvl(s.����, 0) ����, c.ҽ�����, i.���㵥λ, Nvl(s.����, Nvl(x.����, '')) ����, " & _
        "          Nvl(m.�����, -1) �����, Nvl(c.ҽ�����, -1) ҽ��id, '' �ⷿ��λ,x.�Ƿ���, m.���id, " & _
        "          s.�Է�����id As ����id, c.��� �������, C.��¼����,C.�����־,0 �������, z.���� As ������ " & _
        "       From δ��ҩƷ��¼ n,ҩƷ�շ���¼ s, ������ü�¼ c,������Ϣ c1, ����ҽ����¼ m,   " & _
        "          ���ű� p, �������� d, �շ���ĿĿ¼ x, �շ���Ŀ���� e,������ĿĿ¼ i, ������Ŀ���� z " & _
        "       Where n.���� = s.���� And  n.No = s.No AND nvl(n.�ⷿid,[3])+0=nvl(s.�ⷿid,[3])  " & _
        "             And s.����id = c.Id AND s.�Է�����id + 0 = p.Id  " & _
        "             And s.ҩƷid = d.����id And S.ҩƷid = x.Id  " & _
        "             And s.ҩƷid = e.�շ�ϸĿid(+)  And e.����(+) = 3 " & _
        "             And Nvl(Ltrim(Rtrim(s.ժҪ)), 'NOT�ܷ�') <> '�ܷ�'  AND s.����� Is Null And Nvl(s.��ҩ��ʽ, 0) <> -1 " & _
        "             And Mod(s.��¼״̬, 3) = 1 And s.����=[1] " & _
        "             AND d.����ID=i.id  and C.����ID=c1.����ID(+) " & _
        "             AND D.����id = z.������Ŀid(+) And z.����(+) = 2    " & _
        "             AND c.ҽ����� = m.Id(+)  And Nvl(c.����״̬,0)<>1 " & _
        "             And Nvl(n.�ⷿid, [3]) + 0 = [3] and S.����=[1] And S.no=[2] " & _
        "             "
    
    '�ų���δ��ҩƷ�����ʼ�¼
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ���˷������� X " & _
        " Where X.������� = 0 And X.״̬ = 0 And X.�շ�ϸĿid = S.ҩƷid And X.����id = S.����id) "
    
    '����
    str���� = Replace(gstrSQL, "C.���˲���ID", "C.��������id")
    gstrSQL = Replace(gstrSQL, "'' ����", "c.����")
    strסԺ = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    strסԺ = Replace(strסԺ, "And Nvl(c.����״̬,0)<>1", "")
    gstrSQL = str���� & " Union All " & strסԺ

    
    gstrSQL = gstrSQL & "  Order By No, �������"
    
    Set Stuff_RxRefSendDetail = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxRefSendDetail", _
        int����, _
        strNo, _
        lng�ⷿID)
        
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_RxRefReturnNO(ByVal lng�ⷿID As Long, ByVal strBeginTime As String, ByVal strEndTime As String, ByVal intType As Integer, ByVal strContent As String, ByVal bln��ʾ�������� As Boolean, ByVal bln����ģʽ As Boolean, ByVal int������� As Integer) As Recordset
    '��ȡ�ѷ��ϵĵ���
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim str�������� As String
    Dim str���� As String
    Dim i As Integer
    Dim strסԺ As String
    Dim strGroup As String
    Dim strSQL As String

    On Error GoTo errHandle

    Select Case T_Para.���ĵ�λ
    Case 0  'ɢװ��λ
         strFields = "X.���㵥λ ��λ,1 as ����ϵ��, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,"
    End Select


    strWhere1 = ""
    If bln����ģʽ Then
        Select Case intType
            Case mFindType.���ݺ�
                strWhere1 = "  AND A.NO =[4]  "
            Case mFindType.IC��, mFindType.���֤
                strWhere1 = "  AND H.����iD=[4]  "
            Case mFindType.סԺ��
                strWhere1 = "  AND H.��ʶ��=[4] and H.�����־=2 "
            Case mFindType.����
                strWhere1 = "  AND H.���� like [4] "
            Case mFindType.�����
                strWhere1 = "  AND H.��ʶ��=[4] and H.�����־=1 "
            Case mFindType.ҽ����
                strWhere1 = "  AND H1.���￨��=[4]  "
        End Select
    End If
    
    If T_Para.�������� = "" Then
        str���� = " A.���� in (24,25,26)"
    Else
        For i = 0 To UBound(Split(T_Para.��������, ","))
            If str���� = "" Then
               str���� = "(A.����=" & Split(T_Para.��������, ",")(i)
            Else
               str���� = str���� & " or A.����=" & Split(T_Para.��������, ",")(i)
            End If
        Next
        str���� = str���� & ")"
    End If
    
    If bln��ʾ�������� = False Then
        gstrSQL = " SELECT DISTINCT '' As ��ɫ, A.��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 24, '�շ�', 25, '����') ����," & lng�ⷿID & " �ⷿid,A.��¼״̬," & _
                 "      A.����,1 ���շ�,A.����� ��ҩ��,A.NO,H.����,sum(A.���۽��) AS ���," & _
                 "      TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,1 �ɲ���,' ' ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼���� " & _
                 " FROM " & _
                 "      (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                 "          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                 "          A.���ۼ�,B.���۽�� ���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.������, A.�������� " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.Ч��,A.����,A.ʵ������,A.��¼״̬,A.��ҩ����,A.���ۼ�,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.������, Nvl(A.ע��֤��, 0) As �������� " & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                 "          AND A.�ⷿID+0=[1] And A.������� Between [2] And [3] and " & str���� & _
                 "          ) A," & _
                 "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����,SUM(A.���۽��) ���۽��" & _
                 "          FROM ҩƷ�շ���¼ A" & _
                 "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL and " & str���� & _
                 "          AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & _
                 "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
        gstrSQL = gstrSQL & _
                 "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0" & _
                 "     ) A,������ü�¼ H,������Ϣ B" & _
                 " WHERE A.�ⷿID+0=[1] and H.����id=B.����id(+)  " & _
                  strWhere1 & _
                 " AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0) AND A.����� IS NOT NULL AND A.����ID=H.ID AND A.ʵ������<>0 "
    Else
        '�嵥��ʾÿ�ʲ�������
         gstrSQL = " SELECT DISTINCT '' As ��ɫ, A.��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 24, '�շ�', 25, '����') ����," & lng�ⷿID & " �ⷿid,A.��¼״̬,A.����,1 ���շ�,A.����� ��ҩ��," & _
                  "      A.NO,H.����,sum(A.���۽��) AS ���,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,A.�ɲ���," & _
                  "      DECODE(A.��¼״̬,1,'��1�η���',DECODE(MOD(A.��¼״̬,3),0,'��1�η���',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η���',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'������')) ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼����,Zl_Get�շ����(A.����,A.NO,[1]) As �շ���� " & _
                  " FROM " & _
                  "      (SELECT * FROM" & _
                  "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                  "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                  "              A.���ۼ� ,A.���۽�� ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ���, A.������, A.�������� " & _
                  "          FROM" & _
                  "              (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.���,A.����ID,A.����,A.����,A.Ч��,A.����,A.ʵ������,A.��¼״̬,A.��ҩ����,A.���ۼ�,A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID, A.������, Nvl(A.ע��֤��, 0) As �������� " & _
                  "              FROM ҩƷ�շ���¼ A" & _
                  "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                  "              AND A.�ⷿID+0=[1] And A.������� Between [2] And [3] and " & str���� & _
                  "              ) A," & _
                  "              (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                  "              FROM ҩƷ�շ���¼ A" & _
                  "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL and " & str���� & _
                  "              AND A.�ⷿID+0=[1] And A.������� Between [2] And [3]  " & _
                  "              GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
         gstrSQL = gstrSQL & _
                  "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.���)" & _
                  "          UNION" & _
                  "          SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                  "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬,A.��ҩ����," & _
                  "          A.���ۼ� , A.���۽�� ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                  "          DECODE(��¼״̬,1,1,DECODE(MOD(��¼״̬,3),0,1,MOD(��¼״̬,3)+1)) �ɲ���, A.������, Nvl(A.ע��֤��, 0) As �������� " & _
                  "          FROM ҩƷ�շ���¼ A" & _
                  "          WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0) And A.������� Between [2] And [3] and " & str����
         gstrSQL = gstrSQL & _
                  "     ) A,������ü�¼ H,������Ϣ B" & _
                  " WHERE A.�ⷿID+0=[1] and H.����id=B.����id(+) " & _
                  strWhere1 & _
                  " AND A.����� IS NOT NULL AND A.����ID=H.ID "
    End If
    
    If bln��ʾ�������� = False Then
        strGroup = " GROUP BY A.��������,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 24, '�շ�', 25, '����'),A.����,1,A.�����,A.NO,H.����,A.��¼״̬," & _
            " TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼���� "
    Else
        strGroup = " GROUP BY A.��������,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 24, '�շ�', 25, '����') ,A.����,1,A.�����,A.��¼״̬," & _
            " A.NO,H.����,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),A.�ɲ���," & _
            " DECODE(A.��¼״̬,1,'��1�η���',DECODE(MOD(A.��¼״̬,3),0,'��1�η���',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η���',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'������')),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,B.ҽ����,B.סԺ��,H.�����־, H.��¼���� "
    End If
    
    '�������סԺ
    If int������� = 1 Then
        '���ﻮ�ۼ��������
        gstrSQL = gstrSQL & strGroup
    Else
        If int������� = 0 Then
            '���ＰסԺ���е���
            str���� = gstrSQL
            strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
            
            str���� = str���� & strGroup
            strסԺ = strסԺ & strGroup
            
            gstrSQL = str���� & " Union All " & strסԺ
        Else
            'סԺ����
            strסԺ = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            strסԺ = strסԺ & strGroup
            gstrSQL = strסԺ
        End If
    End If
     
    'order by
    gstrSQL = gstrSQL & " order by ����,����,NO "
    
    '�жϴӿ�ʼ���ں��Ƿ����ת���Ĵ�������
    blnHistory = sys.IsMovedByDate(strBeginTime)
    
    '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ����
    If blnHistory Then
        strSQL = gstrSQL
        strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
        gstrSQL = gstrSQL & " UNION ALL " & strSQL
    End If

    Set Stuff_RxRefReturnNO = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ���ϵ���-Stuff_RxRefReturnNO", _
        lng�ⷿID, _
        CDate(strBeginTime), CDate(strEndTime), _
        strContent)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_RxRefSendNO(ByVal lng�ⷿID As Long, ByVal strBeginTime As String, ByVal strEndTime As String, ByVal intType As Integer, ByVal strContent As String, ByVal bln��ҩ���� As Boolean, ByVal bln����ģʽ As Boolean, ByVal int������� As Integer) As Recordset
'��ȡ�����ϵĵ���
    Dim strSQL As String
    Dim str���� As String
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim strWhere As String
    Dim str���� As String
    Dim strסԺ As String
    
    On Error GoTo errHandle
    
    
    strWhere = ""
    If bln����ģʽ Then
        Select Case intType
            Case mFindType.���ݺ�
                strWhere = "  AND A.NO =[4]  "
            Case mFindType.IC��, mFindType.���֤
                strWhere = "  AND d.����iD=[4]  "
            Case mFindType.סԺ��
                strWhere = "  AND d.��ʶ��=[4] and d.�����־=2 "
            Case mFindType.����
                strWhere = "  AND d.���� like [4] "
            Case mFindType.�����
                strWhere = "  AND d.��ʶ��=[4] and d.�����־=1 "
        End Select
    End If
    
    If T_Para.�������� = "" Then
        str���� = " a.���� in (24,25,26)"
    Else
        For i = 0 To UBound(Split(T_Para.��������, ","))
            If str���� = "" Then
               str���� = "(a.����=" & Split(T_Para.��������, ",")(i)
            Else
               str���� = str���� & " or a.����=" & Split(T_Para.��������, ",")(i)
            End If
        Next
        str���� = str���� & ")"
    End If
    
    
    gstrSQL = "Select /*+ Rule*/" & vbNewLine & _
        " ����, ���շ�, No, ����, To_Char(Sum(Round(���۽��, 2)), '999999990.00') As ���, ����, �ɲ���, ˵��, ���￨��, �����, ���֤��, Ic����, ����id, ҽ����, סԺ��," & vbNewLine & _
        " Sum(Round(ʵ�ս��, 2)) ʵ�ս��, �����־, ��¼����,��¼״̬," & lng�ⷿID & " �ⷿid" & vbNewLine & _
        "From ("

    strSQL = "Select a.����, a.���շ�, a.No, a.����, c.���۽��, a.����, a.�ɲ���, a.˵��, a.���￨��, a.�����, a.���֤��, a.Ic����, a.����id, a.ҽ����, a.סԺ��," & vbNewLine & _
        "              d.ʵ�ս��, Nvl(a.��������, Nvl(c.ע��֤��, 0)) As ��������, d.�����־, d.��¼����,c.��¼״̬, d.�շ����" & vbNewLine & _
        "" & vbNewLine & _
        "       From (Select Distinct b.���￨��, b.�����, b.���֤��, b.Ic����, b.ҽ����, b.סԺ��, a.���ȼ�, a.��ҩ����, a.��������," & vbNewLine & _
        "                              Decode(Nvl(a.���շ�, 0), 1, '', '(δ)') || Decode(a.����, 8, '�շ�', 9, '����') ����, a.����, a.���շ�, a.No," & vbNewLine & _
        "                              a.����, To_Char(a.��������, 'yyyy-MM-dd hh24:mi:ss') ����, 1 �ɲ���, ' ' ˵��, b.����id, a.��������, a.�Է�����id" & vbNewLine & _
        "              From δ��ҩƷ��¼ a, ������Ϣ b" & vbNewLine & _
        "              Where 1 = 1 And (a.�ⷿid =[1] Or a.�ⷿid Is Null) And" & vbNewLine & _
        "                    a.�������� Between [2] And [3]" & vbNewLine & _
        "                     And a.����id = b.����id(+) And " & str���� & _
        "                     )a, ҩƷ�շ���¼ c, ������ü�¼ d, ���ű� b" & vbNewLine & _
        "       Where c.����id = d.Id And Nvl(c.��ҩ��ʽ, -999) <> -1 And a.���� = c.���� And a.No = c.No And c.����� Is Null And d.ִ��״̬ <> 9 And" & vbNewLine & _
        "             (c.�ⷿid = [1] Or c.�ⷿid Is Null) And a.�Է�����id = b.Id And " & IIf(bln��ҩ����, "Mod(c.��¼״̬, 3) = 1", "c.��¼״̬=1") & strWhere
        
    If int������� = 0 Then
        '����
        str���� = Replace(strSQL, "C.���˲���ID", "C.��������id")
        strSQL = Replace(strSQL, "'' ����", "c.����")
        strסԺ = Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
        strסԺ = Replace(strסԺ, "And Nvl(c.����״̬,0)<>1", "")
        strSQL = str���� & " Union All " & strסԺ
    ElseIf int������� = 3 Then
        'סԺ���ʵ�
        strSQL = Replace(strSQL, "'' ����", "c.����")
        strSQL = Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
        strSQL = Replace(strSQL, "And Nvl(c.����״̬,0)<>1", "")
    End If
    
    gstrSQL = gstrSQL & strSQL
    
    gstrSQL = gstrSQL & ") a" & vbNewLine & _
        "Group By a.����, a .���շ�, a.No, a.����, a.����, a.�ɲ���, a.˵��, a.���￨��, a.�����, a.���֤��, a.Ic����, a.����id, a.ҽ����, a.סԺ��, a.��������," & vbNewLine & _
        "         a.�����־, a.��¼����,a.��¼״̬" & vbNewLine & _
        "Order By a.����, a.No"

    Set Stuff_RxRefSendNO = zlDataBase.OpenSQLRecord(gstrSQL, "RefreshSendList", lng�ⷿID, CDate(strBeginTime), CDate(strEndTime), strContent)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxWork(ByVal intType As Integer, ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal lng�ⷿID As Long, ByVal int���� As Integer, ByVal strNo As String, ByVal str��ҩ���� As String) As Boolean
'�Ե��ݽ��з����ϲ���
'��������ֵ��true-�����ɹ���false-����ʧ��
'intType:0-���ϣ�1-����
    '������֤
    If Stuff_RxValied(strPrivs, intType, lng�ⷿID, int����, strNo, rsTemp) = False Then
        Stuff_RxWork = False
        Exit Function
    End If
    
    If intType = 0 Then
        '���ϴ���
        Stuff_RxWork = Stuff_RxSend(strPrivs, rsTemp, lng�ⷿID, str��ҩ����)
    Else
        '���ϴ���
        Stuff_RxWork = Stuff_RxReturn(strPrivs, rsTemp, str��ҩ����)
    End If
End Function


Public Function Stuff_RxReturn(ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal str��ҩ���� As String) As Boolean
'��ָ�����ѷ��ϵ��ݽ������ϲ���
'��������ֵ��true-�����ɹ���false-����ʧ��
'������strPrivs-Ȩ���ַ���
'rsTemp-���ϲ��������ݼ�
    Dim strDate As String
    Dim arrSQL As Variant
    Dim i As Integer
    Dim blnTrans As Boolean
    Dim dbl��ҩ���� As Double
    
    On Error GoTo errHandle
    
    arrSQL = Array()
    strDate = sys.Currentdate
    With rsTemp
        If Not rsTemp Is Nothing Then
            If .EOF Then
                Stuff_RxReturn = False
                Exit Function
            End If
            
            Do While Not .EOF
                dbl��ҩ���� = Val(Mid(Mid(str��ҩ����, InStr(1, str��ҩ����, "," & !Id & ",") + 2 + Len(!Id)), 1, InStr(1, Mid(str��ҩ����, InStr(1, str��ҩ����, "," & !Id & ",") + 2 + Len(!Id)), "|") - 1))
                'Zl_�����շ���¼_��������
                gstrSQL = "Zl_�����շ���¼_��������("
                '    �շ�id_In   In ҩƷ�շ���¼.ID%Type,
                gstrSQL = gstrSQL & "" & Nvl(!Id) & ","
                '    �����_In   In ҩƷ�շ���¼.�����%Type,
                gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                '    �������_In In ҩƷ�շ���¼.�������%Type,
                gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                '    ����_In     In ҩƷ���.�ϴ�����%Type := Null,
                gstrSQL = gstrSQL & "'" & Nvl(!����) & "',"
                '    Ч��_In     In ҩƷ���.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(IsNull(!Ч��), "NULL", IIf(Nvl(!Ч��) = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                '    ����_In     In ҩƷ���.�ϴβ���%Type := Null,
                gstrSQL = gstrSQL & "'" & Nvl(!����) & "',"
                '    ��������_In In ҩƷ�շ���¼.ʵ������%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��ҩ���� & ","
                '    �Զ�����_In Integer := 0,
                gstrSQL = gstrSQL & "0,"
                '    ������_In   In ҩƷ�շ���¼.������%Type := Null
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "')"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                .MoveNext
            Loop
            
            gcnOracle.BeginTrans
            blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDataBase.ExecuteProcedure(CStr(arrSQL(i)), "���ݷ���_Stuff_RxReturn")
            Next
            gcnOracle.CommitTrans
            blnTrans = False
        End If
    End With
    Stuff_RxReturn = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxSend(ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal lng�ⷿID As Long, ByVal str������ As String) As Boolean
'��ָ����δ���ϵ��ݽ��з��Ų���
'��������ֵ��true-�����ɹ���false-����ʧ��
'������strPrivs-Ȩ���ַ���
'rsTemp-���ϲ��������ݼ�
    Dim strDate As String
    Dim strID���� As String
    
    On Error GoTo errHandle
    
    strDate = sys.Currentdate
    With rsTemp
        If Not rsTemp Is Nothing Then
            If .EOF Then
                Stuff_RxSend = False
                Exit Function
            End If
            
            Do While Not .EOF
                strID���� = IIf(strID���� = "", "", strID���� & "|") & !Id & "," & !����
                .MoveNext
            Loop
            'Zl_ҩƷ�շ���¼_��������
            gstrSQL = "Zl_ҩƷ�շ���¼_��������("
            '    �շ�id_In     In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
            gstrSQL = gstrSQL & "'" & strID���� & "',"
            '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
            gstrSQL = gstrSQL & "" & lng�ⷿID & ","
            '    �����_In     In ҩƷ�շ���¼.�����%Type,
            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
            '    �������_In   In ҩƷ�շ���¼.�������%Type,
            gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
            '    ���Ϸ�ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3, --1-��������;2-��������;3-���ŷ���;-1 ֹͣ����
            gstrSQL = gstrSQL & "1,"
            '    ������_In     In ҩƷ�շ���¼.������%Type := Null,
            gstrSQL = gstrSQL & "'',"
            '    ���ϱ�ʶ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
            gstrSQL = gstrSQL & "Null,"
            '    ������_In     In ҩƷ�շ���¼.��ҩ��%Type := Null
            gstrSQL = gstrSQL & "'" & str������ & "',"
            '    ����Ա����
            gstrSQL = gstrSQL & "'" & UserInfo.��� & "')"
            
            Call zlDataBase.ExecuteProcedure(gstrSQL, "���ݷ���_Stuff_RxSend")
        End If
    End With
    Stuff_RxSend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function LoadPerson(ByVal lng��Աid As Long) As Recordset
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select distinct a.id,a.��� as ����,A.���� As ����,����" & _
        "   From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
        "   Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id " & _
        "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND B.����id in (Select ����ID From ������Ա where ��Աid=[1] ) " & _
        "   ORDER BY ���� "
    Set LoadPerson = zlDataBase.OpenSQLRecord(gstrSQL, "����������-LoadPerson", lng��Աid)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlfuncCard_GetPatiID(ByVal lngCardID As Long, ByVal strCardNo As String) As Long
    'һ��ͨ���ܣ�ͨ������ȡ����ID
    Dim lng����id As Long
    
    On Error GoTo errHandle
    If Not gobjSquareCard Is Nothing Then
        'ͨ����ID�Ϳ��Ų��Ҳ���ID
        gobjSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng����id
        
        If lng����id > 0 Then
            zlfuncCard_GetPatiID = lng����id
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlfuncCard_SetText(ByVal objTxt As TextBox, ByVal strCardProperty As String)
    '�������������
    '���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
    objTxt.Text = ""
    objTxt.Tag = ""
    objTxt.MaxLength = 0
    
    objTxt.Tag = strCardProperty
    objTxt.MaxLength = Val(Split(strCardProperty, "|")(gCardFormat.���ų���))
    objTxt.PasswordChar = IIf(Trim(Split(strCardProperty, "|")(gCardFormat.��������)) <> "", "*", "")
End Sub

Public Function Stuff_Check��Ժ����(ByVal strPrivs As String, ByVal lng���� As Long, ByVal strNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer, Optional ByVal lng����id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ժ�����Ƿ�������,��Ҫ����Ȩ�޿���(���û��Ȩ�ޡ����˳�Ժ���˴����������������ϲ���)
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------

    '����˵���������ǰ������סԺ���ˣ�
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng���� = 24 Then
        Stuff_Check��Ժ���� = True
        Exit Function
    End If
    
    '���δ���벡��ID�����Զ���ȡ
    If lng����id = 0 Then
        gstrSQL = "Select ����ID From ������ü�¼ Where ID = (Select ����ID From ҩƷ�շ���¼ Where ����=[1] And NO=[2] And Rownum<2)"
        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����, strNo)
        lng����id = rsTemp!����ID
    End If
    
    'ȡ��������
    gstrSQL = "Select ���� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��������", lng����id)
    If rsTemp.EOF Then
        MsgBox "�ڴ���[" & strNo & "]�У����˲��棬������ֹ��", vbInformation, gstrSysName
        Exit Function
    End If
    str���� = rsTemp!����
    
    '�����ǰ������סԺ���ˣ����û��Ȩ�ޡ����˳�Ժ���˴���������������ҩ����
    If zlStr.IsHavePrivs(strPrivs, "���˳�Ժ���˴���") = False Then
        '��鲡����Ԥ��Ժ���Ժ
        gstrSQL = " Select 1 From ������ҳ A,������Ϣ B" & _
                  " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ID=[1] " & _
                  " And (Nvl(A.״̬,0)=3 Or A.��Ժ���� Is Not NULL)"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѳ�Ժ", lng����id)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "�ڴ���[" & strNo & "]�У����ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵Ĵ������з��ϡ����ϵ�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Stuff_Check��Ժ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Stuff_Check���ʴ���(ByVal strPrivs As String, ByVal lng���� As Long, ByVal strNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��鴦���Ƿ��Ѿ�������,���ʵĴ������ܷ����ϲ���
    '���:  lng����    ����ǰ��������
    '       strNO      ����ǰ���ݺ�
    '       lng����ID  �����Զಡ�˵���Ч
    '       str��ţ���ص������,��,����
    '����:
    '����:���ݺϷ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng���� = 24 Then
        Stuff_Check���ʴ��� = True
        Exit Function
    End If
    
    '���û��Ȩ�ޡ����˽��ʴ����������ô����Ƿ��ѽ��ʣ��ѽ��ʴ������������ϲ���
    If zlStr.IsHavePrivs(strPrivs, "���˽��ʴ���") = 0 Then
    
        gstrSQL = "Select Nvl(Sum(Nvl(���ʽ��,0)),0) AS ���ʽ��   " & _
                 "  From ������ü�¼   " & _
                 "  Where Mod(��¼����,10) = 2 and NO = [1]"
        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        gstrSQL = gstrSQL & " Order By ���ʽ�� Desc"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѽ���", strNo)
        If Nvl(rsTemp!���ʽ��, 0) <> 0 Then
            MsgBox "�ô���[" & strNo & "]�ѽ��ʣ���û�ж��ѽ��ʴ������з��ϡ����ϵ�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Stuff_Check���ʴ��� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



