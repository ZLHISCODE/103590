Attribute VB_Name = "mdlSQL"
Option Explicit

'######################################################################################################################

Public Enum SQL

    ���˻�����Ϣ
    ���������嵥
    
    �������ѡ��
    ������Ϲ���

    ��������ѡ��
    �����������
    
    ������϶���
    
    ����ʽѡ��
    ����ʽ����
    
    ִ�з���ѡ��
    
    ����������¼
    �ȴ��������
    ���������¼
    
    ������Ŀѡ��
    ������Ŀ����
    
    ������Ŀѡ��
    ������Ŀ����
    �շ�ִ�п���
    
    ����ҩƷѡ��
    ����ҩƷ����
    
    ҩƷ��Ŀѡ��
    ҩƷ��Ŀ����
    
    ������Ŀѡ��
    ������Ŀ����
    
    ��Ա��Ϣѡ��
    ��Ա��Ϣ����
    
    ������Ϣѡ��
    ������Ϣ����
    
    ��Ա����ѡ��
    ��Ա���Ź���
    
    �����������
    ������ϼ�¼
    
    ����������
    �����鷽ʽ
    
    ������ҩ�ο�
    �������ϲο�
    �������Ʋο�
    
    ������ҩѡ��
    ��������ѡ��
    ��������ѡ��
    ����ִ�п���
    
    �ٴ����ż�¼
    ��Լ��λѡ��
    ��Լ��λ����
    ����ҽ����Ա
    ��Ա����ѡ��
End Enum

'######################################################################################################################

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String) As String
    '******************************************************************************************************************
    '����:  ���в���SQL���
    '����:  strMenu             Ҫ������SQL����
    '       strParam            ������,��ʽ:"����ֵ1'����ֵ2"
    '����:  SQL���
    '******************************************************************************************************************
    
    Dim strSQL As String
    Dim varParam As Variant
    Dim strTmp As String
    Dim rs As New ADODB.Recordset

    Dim lng���ͺ� As Long
    Dim str�ѱ� As String
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���˻�����Ϣ
        
            strSQL = "SELECT A.ID," & _
                     "A.����," & _
                     "D.�����,D.������,D.���￨��,D.IC����," & _
                     "D.����," & _
                     "D.�Ա�," & _
                     "D.����," & _
                     "D.����״��," & _
                     "A.���ʱ��," & _
                     "C.��첡��id,C.����ʱ��,C.�������,A.�������,D.��ϵ�˵绰,D.������λ," & _
                     "B.���� AS �������� " & _
                "FROM ���ǼǼ�¼ A,��Լ��λ B,�����Ա���� C,������Ϣ D " & _
                "WHERE A.ID=C.�Ǽ�id AND A.��Լ��λID=B.ID(+) AND D.����id=C.����id AND C.ID=[1] "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������嵥
        
            If strParam = "����" Then
            
                strSQL = "SELECT A.����,A.����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='����' ORDER BY A.����||'-'||A.����"
            
            Else
                strSQL = "SELECT A.����,A.����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='����' " & _
                            "AND A.ID IN (SELECT ����id FROM ������Ա WHERE ��Աid=[1])  ORDER BY A.����||'-'||A.����"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�������ѡ��
            
            strSQL = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "���� " & _
                "FROM ������Ϸ��� " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "B.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.���� " & _
                "FROM �������Ŀ¼ A,����������� B " & _
                "WHERE A.ID=B.���ID"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������Ϲ���

            If Val(varParam(0)) = 1 Then
                '��ȫ���֣����������
                
                strSQL = "SELECT A.����,A.����,A.ID " & _
                            "FROM �������Ŀ¼ A " & _
                            "Where A.��� = 1 " & _
                                  "AND A.���� LIKE [1]"
                
            ElseIf Val(varParam(0)) = 2 Then
                '��ȫ��ĸ�����������
                
                strSQL = "SELECT A.����,A.����,A.ID " & _
                            "FROM �������Ŀ¼ A " & _
                            "Where A.��� = 1 " & _
                                  "And A.id IN (SELECT B.���id FROM ������ϱ��� B WHERE ���� LIKE [2])"
                
            Else
                
                strSQL = "SELECT A.����,A.����,A.ID " & _
                            "FROM �������Ŀ¼ A " & _
                            "Where A.��� = 1 " & _
                                  "AND ((���� LIKE [1] OR ���� LIKE [2]) " & _
                                  "OR A.id IN (SELECT B.���id FROM ������ϱ��� B WHERE (���� LIKE [2] OR ���� LIKE [2])))"
                
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��������ѡ��
            
            strSQL = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "NULL AS ����," & _
                        "����," & _
                        "NULL AS ����,Null As ���� " & _
                "FROM ����������� " & _
                "WHERE ���='D' " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.����,a.���� " & _
                "FROM ��������Ŀ¼ A " & _
                "WHERE ���=[1] " & _
                    "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (0,1,2) "
                            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����������
        
            strSQL = "SELECT   ����," & _
                               "����," & _
                               "����," & _
                               "����," & _
                               "ID " & _
                        "FROM ��������Ŀ¼ " & _
                        "WHERE ���=[3] " & _
                            "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (0,1,2) "
                            
            If Val(varParam(0)) = 1 Then
                '��ȫ���֣����������
                
                strSQL = strSQL & " And ���� LIKE [1] "
                
            ElseIf Val(varParam(0)) = 2 Then
                '��ȫ��ĸ�����������
                
                strSQL = strSQL & " And ���� LIKE [2] "
                
            Else
                
                strSQL = strSQL & "AND (���� LIKE [1] OR ���� LIKE [2] OR ���� LIKE [2])"
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������϶���
        
            strSQL = "SELECT A.����ID,A.���ID,B.���� AS ��������,C.���� AS ������� " & _
                "FROM ������϶��� A,��������Ŀ¼ B,�������Ŀ¼ C " & _
                "WHERE A.����ID=B.ID AND A.���ID=C.ID AND (A.����ID=[1] OR A.���ID=[2])"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.����������¼
            
            strTmp = ""
            
            If Trim(Split(strParam, ";")(2)) <> "" Then strTmp = strTmp & " AND b.���� LIKE [2] "
            If Trim(Split(strParam, ";")(3)) <> "" Then strTmp = strTmp & " AND b.סԺ�� = [3] "
            If Trim(Split(strParam, ";")(4)) <> "" Then strTmp = strTmp & " AND b.��ǰ���� = [4] "
            If Trim(Split(strParam, ";")(5)) <> "" Then strTmp = strTmp & " AND b.����� = [5] "
            If Val(Trim(Split(strParam, ";")(7))) > 0 Then strTmp = strTmp & " AND a.������ĿID = [6] "
            
            strSQL = "Select  e.ID,Decode(e.����״̬,1,'���',2,'����',3,'����',4,'���','����') As ͼ��,a.Id As ҽ��id," & vbNewLine & _
                        "       Decode(a.������־,1,'����','') As ������־," & vbNewLine & _
                        "       DECODE(a.������Դ,1,'����',2,'סԺ','����') AS ������Դ," & vbNewLine & _
                        "       Decode(a.������Ŀid,Null,a.ҽ������,f.����) As ҽ������," & vbNewLine & _
                        "       a.����ʱ��," & vbNewLine & _
                        "       b.����," & vbNewLine & _
                        "       b.�����," & vbNewLine & _
                        "       b.סԺ��,b.��ǰ���� As ����," & vbNewLine & _
                        "       c.���� As ���˿���," & vbNewLine & _
                        "       d.���� As ��������," & vbNewLine & _
                        "       a.����ҽ�� As ������," & vbNewLine & _
                        "       a.ҽ��״̬," & vbNewLine & _
                        "       a.����id," & vbNewLine & _
                        "       a.��ҳid," & vbNewLine & _
                        "       a.������Ŀid," & vbNewLine & _
                        "       e.����״̬,g.���ͺ�,g.ִ��״̬,a.�Һŵ�,0 As ״̬,b.��Ժʱ�� As ��Ժ����,b.��ǰ����id,b.��ǰ����id,b.IC����,b.���֤�� "
            strSQL = strSQL & _
                        "From ����ҽ����¼ a,����ҽ������ g, " & vbNewLine & _
                        "     ������Ϣ b," & vbNewLine & _
                        "     ���ű� c," & vbNewLine & _
                        "     ���ű� d," & vbNewLine & _
                        "     ����������¼ e,������ĿĿ¼ f " & vbNewLine & _
                        "Where Nvl(a.�������,'F')='F' " & vbNewLine & _
                        "      And a.���id Is Null" & vbNewLine & _
                        "      And a.ҽ��״̬<>4 " & strTmp & vbNewLine & _
                        "      And a.ִ�п���id+0=[1]" & vbNewLine & _
                        "      And b.����id=a.����id" & vbNewLine & _
                        "      And c.Id=a.���˿���id" & vbNewLine & _
                        "      And d.Id=a.��������id And f.Id(+)=a.������Ŀid " & vbNewLine & _
                        "      And a.Id=e.ҽ��id  And e.����״̬=3 And a.ID=g.ҽ��id(+) And (e.������=[7] Or [7] Is Null) "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������¼

            strTmp = ""
            
            If Val(Trim(Split(strParam, ";")(10))) = 1 Then
                If Split(strParam, ";")(0) <> "" Then strTmp = " AND e.��������ʱ�� BETWEEN [3] AND [4] "
            Else
                If Split(strParam, ";")(0) <> "" Then strTmp = " AND a.��ʼִ��ʱ�� BETWEEN [3] AND [4] "
            End If
            If Trim(Split(strParam, ";")(2)) <> "" Then strTmp = strTmp & " AND b.���� LIKE [5] "
            If Trim(Split(strParam, ";")(3)) <> "" Then strTmp = strTmp & " AND b.סԺ�� = [6] "
            If Trim(Split(strParam, ";")(4)) <> "" Then strTmp = strTmp & " AND b.��ǰ���� = [7] "
            If Trim(Split(strParam, ";")(5)) <> "" Then strTmp = strTmp & " AND b.����� = [8] "
            If Val(Trim(Split(strParam, ";")(7))) > 0 Then strTmp = strTmp & " AND a.������ĿID = [9] "
                        
            
            strSQL = "Select  /*+rule*/ e.ID,Decode(e.����״̬,1,'���',2,'����',3,'����',4,'���','����') As ͼ��,a.Id As ҽ��id," & vbNewLine & _
                        "       Decode(a.������־,1,'����','') As ������־," & vbNewLine & _
                        "       DECODE(a.������Դ,1,'����',2,'סԺ','����') AS ������Դ," & vbNewLine & _
                        "       Decode(a.������Ŀid,Null,a.ҽ������,f.����) As ҽ������," & vbNewLine & _
                        "       a.����ʱ��," & vbNewLine & _
                        "       b.����," & vbNewLine & _
                        "       b.�����," & vbNewLine & _
                        "       b.סԺ��,b.��ǰ���� As ����," & vbNewLine & _
                        "       c.���� As ���˿���," & vbNewLine & _
                        "       d.���� As ��������," & vbNewLine & _
                        "       a.����ҽ�� As ������," & vbNewLine & _
                        "       a.ҽ��״̬," & vbNewLine & _
                        "       a.����id," & vbNewLine & _
                        "       a.��ҳid," & vbNewLine & _
                        "       a.������Ŀid," & vbNewLine & _
                        "       e.����״̬,g.���ͺ�,g.ִ��״̬,a.�Һŵ�,0 As ״̬,b.��Ժʱ�� As ��Ժ����,b.��ǰ����id,b.��ǰ����id,b.IC����,b.���֤�� "
            strSQL = strSQL & _
                        "From ����ҽ����¼ a,����ҽ������ g, " & vbNewLine & _
                        "     ������Ϣ b," & vbNewLine & _
                        "     ���ű� c," & vbNewLine & _
                        "     ���ű� d," & vbNewLine & _
                        "     ����������¼ e,������ĿĿ¼ f " & vbNewLine & _
                        "Where Nvl(a.�������,'F')='F' " & vbNewLine & _
                        "      And a.���id Is Null" & vbNewLine & _
                        "      And a.ҽ��״̬<>4 " & strTmp & vbNewLine & _
                        "      And a.ִ�п���id+0=[1]" & vbNewLine & _
                        "      And b.����id=a.����id" & vbNewLine & _
                        "      And c.Id=a.���˿���id" & vbNewLine & _
                        "      And d.Id=a.��������id And f.Id(+)=a.������Ŀid " & vbNewLine & _
                        "      And a.Id=e.ҽ��id(+)  And e.����״̬=[2] And a.ID=g.ҽ��id "

    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀѡ��
    
        strSQL = "SELECT DISTINCT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "����," & _
                        "NULL AS ��λ " & _
                "FROM ���Ʒ���Ŀ¼ " & _
                "START WITH ID IN (SELECT DISTINCT ����id FROM ������ĿĿ¼ WHERE ��� = 'F' AND ������� IN (1, 2, 3) AND (����ʱ�� = TO_DATE('30000101', 'YYYYMMDD') OR ����ʱ�� IS NULL)) CONNECT BY PRIOR �ϼ�ID=ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.���㵥λ AS ��λ " & _
                "FROM ������ĿĿ¼ A "
        strSQL = strSQL & _
                "WHERE (����ʱ�� = TO_DATE('30000101', 'YYYYMMDD') OR ����ʱ�� IS NULL) " & _
                    "AND ������� IN (1, 2, 3) " & _
                    "AND ��� = 'F'"
        strSQL = "SELECT * FROM (" & strSQL & ") ORDER BY ����"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀ����
        
        Select Case Val(varParam(0))
        Case 1
            '��ȫ���֣����������
            strSQL = "Select a.ID,a.����,a.���� " & _
                        "From ������ĿĿ¼ a " & _
                        "Where a.��� = 'F' " & _
                            "And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null) " & _
                            "And a.���� Like [1]"
                    
        Case 2
            '��ȫ��ĸ�����������
            strSQL = "Select Distinct a.ID,a.����,a.���� " & _
                        "From    ������ĿĿ¼ a," & _
                                "������Ŀ���� b " & _
                        "Where   a.��� = 'F' " & _
                            "And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null) " & _
                            "And a.ID=b.������Ŀid " & _
                            "And ����_In Is Not Null " & _
                            "And b.���� Like [2] "
        Case Else
            strSQL = "Select Distinct a.ID,a.����,a.����  " & _
                        "From    ������ĿĿ¼ a, " & _
                                "������Ŀ���� b " & _
                        "Where   a.��� = 'F' " & _
                            "And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null) " & _
                            "And a.ID=b.������Ŀid " & _
                            "And (a.���� Like [1] Or a.���� Like [2] Or b.���� Like [2] Or b.���� Like [2])"
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀѡ��
        
'        strSQL = "Select *" & vbNewLine & _
'                    "From (Select ID,�ϼ�ID,0 As ĩ��,����,���� ,'' As ��λ,'' As ���,'' ������,'' As ����,'' As ���" & vbNewLine & _
'                    "     From �շѷ���Ŀ¼" & vbNewLine & _
'                    "     Start With �ϼ�ID Is Null Connect by Prior ID = �ϼ�ID" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -1 As ID,Null+0 As �ϼ�ID,0 As ĩ��,'-1' As ����,'����ҩ' As ���� ,'' As ��λ,'' As ���,'' ������,'' As ����,'' As ��� From Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -2 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-2' As ����,'�г�ҩ' As ���� ,'' As ��λ,'' As ���,'' ������,'' As ����,'' As ��� from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -3 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-3' As ����,'�в�ҩ' As ���� ,'' As ��λ,'' As ���,'' ������,'' As ����,'' As ��� from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -7 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-7' As ����,'��������' As ���� ,'' As ��λ,'' As ���,'' ������,'' As ����,'' As ��� from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select a.ID,Decode(a.���,'5',-1,'6',-2,'7',-3,'4',-7,a.����id) As �ϼ�ID,1 As ĩ��, a.����,a.����,a.���㵥λ As ��λ,b.���� As ���,a.��� As ������,Trim(To_Char(c.����,'9999999999999.00')) As ����,a.���" & vbNewLine & _
'                    "     From  �շ���ĿĿ¼ a," & vbNewLine & _
'                    "          �շ���Ŀ��� b," & vbNewLine & _
'                    "          (Select �շ�ϸĿid,Sum(�ּ�) As ���� From �շѼ�Ŀ Where ִ������<=Sysdate And (��ֹ���� Is Null Or ��ֹ����>Sysdate) Group by �շ�ϸĿid) c" & vbNewLine & _
'                    "     Where c.�շ�ϸĿid(+)=a.ID" & vbNewLine & _
'                    "            And Nvl(a.�Ƿ���,0)=0" & vbNewLine & _
'                    "            And a.���=b.����" & vbNewLine & _
'                    "            And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null)) a" & vbNewLine & _
'                    "Order By a.ĩ��, a.����"

        strSQL = "Select *" & vbNewLine & _
                    "From (Select ID,�ϼ�ID,0 As ĩ��,����,���� ,'' As ��λ,'' As ���,'' ������,'' As ���" & vbNewLine & _
                    "     From �շѷ���Ŀ¼" & vbNewLine & _
                    "     Start With �ϼ�ID Is Null Connect by Prior ID = �ϼ�ID" & vbNewLine & _
                    "     Union All" & vbNewLine & _
                    "     Select a.ID,a.����id As �ϼ�ID,1 As ĩ��, a.����,a.����,a.���㵥λ As ��λ,b.���� As ���,a.��� As ������,a.���" & vbNewLine & _
                    "     From  �շ���ĿĿ¼ a," & vbNewLine & _
                    "          �շ���Ŀ��� b " & vbNewLine & _
                    "     Where Nvl(a.�Ƿ���,0)=0" & vbNewLine & _
                    "            And a.���=b.���� And a.��� Not In ('5','6','7','4') " & vbNewLine & _
                    "            And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null)) a" & vbNewLine & _
                    "Order By a.ĩ��, a.����"
                    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀ����
        'And a.��� Not In ('5','6','7','4')
        strSQL = "Select a.ID,a.����,a.����,a.���㵥λ As ��λ,b.���� As ���,a.���" & vbNewLine & _
                    "From   �շ���ĿĿ¼ a," & vbNewLine & _
                    "   �շ���Ŀ��� b " & vbNewLine & _
                    "Where  Nvl(a.�Ƿ���,0)=0" & vbNewLine & _
                    "   And a.���=b.����  And a.��� Not In ('5','6','7','4') " & vbNewLine & _
                    "   And (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� Is Null) "
                    
        Select Case Val(varParam(0))
        Case 1                  '��ȫ���֣����������
            strSQL = strSQL & " And a.���� Like [1] "
        Case 2                  '��ȫ��ĸ�����������
            strSQL = strSQL & " And Exists (Select 1 From �շ���Ŀ���� bb Where (bb.���� Like [2] Or bb.���� Like [2]) And a.ID=bb.�շ�ϸĿid) "
        Case Else               '������ĸ���
            strSQL = strSQL & " And (a.���� Like [1] Or a.���� Like [2] Or Exists (Select 1 From �շ���Ŀ���� bb Where (bb.���� Like [2] Or bb.���� Like [2]) And a.ID=bb.�շ�ϸĿid)) "
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����ҩƷѡ��
        
        strSQL = "SELECT ���,DECODE(ҩƷID,0,NULL,ҩƷID) AS ҩƷID,ID,���,DECODE(�ϼ�ID,0,NULL,�ϼ�ID) AS �ϼ�ID,ĩ��,����,����,��λ,����,��� " & _
                     "FROM (SELECT DECODE(����, '5', 'A0', '6', 'B0', 'C0') AS ���, " & _
                                  "0 AS ҩƷID, " & _
                                  "DECODE(����, '5', -1, '6', -2, -3) AS ID, " & _
                                  "0 AS �ϼ�ID, " & _
                                  "0 AS ĩ��, " & _
                                  "����, " & _
                                  "����, NULL AS ��λ," & _
                                  "NULL AS ���," & _
                                  "NULL AS ����,'' As ��� " & _
                             "From ������Ŀ��� WHERE ���� IN ('5','6','7') " & _
                           "Union All " & _
                             "SELECT DISTINCT DECODE(����, 1, 'A', 2, 'B', 'C') || ���� AS ���, " & _
                                    "0 AS ҩƷID, " & _
                                    "ID, " & _
                                    "DECODE(�ϼ�ID,NULL,DECODE(����, 1, -1, 2, -2, -3),�ϼ�ID) AS �ϼ�ID, " & _
                                    "0 AS ĩ��, " & _
                                    "����, " & _
                                    "����, NULL AS ��λ," & _
                                    "NULL AS ���," & _
                                    "NULL AS ����,'' As ��� " & _
                               "From ���Ʒ���Ŀ¼  where DECODE(����,1,'5',2,'6','7') IN ('5','6','7') "
        strSQL = strSQL & _
                            "Start With ID IN (SELECT Y.����id FROM ������ĿĿ¼ Y,ҩƷ���� X WHERE X.ҩ��id=Y.ID AND X.�������='����ҩ') " & _
                             "Connect by Prior �ϼ�ID = ID " & _
                             "Union All " & _
                                "SELECT 'D1' AS ���, " & _
                                      "B.ҩƷID, " & _
                                      "B.ҩ��ID AS ID, " & _
                                      "C.����ID AS �ϼ�ID," & _
                                      "1 as ĩ��, " & _
                                      "D.����, " & _
                                      "D.����, " & _
                                      "d.���㵥λ AS ��λ, " & _
                                      "D.���," & _
                                      "A.ҩƷ���� AS ����,D.��� " & _
                                 "FROM ҩƷ���� A,ҩƷ��� B,������ĿĿ¼ C,�շ���ĿĿ¼ D " & _
                                "WHERE A.ҩ��id=B.ҩ��id " & _
                                        "AND C.ID=A.ҩ��id " & _
                                        "AND D.ID=B.ҩƷid " & _
                                        "AND C.��� IN ('5','6','7') " & _
                                        "AND A.�������='����ҩ'" & _
                                        "AND (D.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.����ʱ�� is NULL) " & _
                           ") " & _
                    "ORDER BY ĩ��,���"
                        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����ҩƷ����
        
        strSQL = "SELECT DECODE(C.���, '5', '����ҩ', '6', '�г�ҩ', '�в�ҩ') AS ����," & _
                   "D.����," & _
                   "D.����," & _
                   "D.���," & _
                   "A.ҩƷ���� As ����," & _
                   "d.���㵥λ As ��λ," & _
                   "B.ҩƷID," & _
                   "B.ҩƷID As ID," & _
                   "A.ҩ��ID,d.��� " & _
             "FROM ҩƷ���� A, ҩƷ��� B, ������ĿĿ¼ C, �շ���ĿĿ¼ D " & _
             "WHERE A.ҩ��ID = B.ҩ��ID " & _
                   "AND C.ID = A.ҩ��ID " & _
                   "AND D.ID = B.ҩƷID " & _
                   "AND A.������� = '����ҩ' " & _
                   "AND C.��� IN ('5','6','7') " & _
                   "AND (D.����ʱ�� IS NULL OR D.����ʱ�� = TO_DATE('3000-01-01', 'yyyy-MM-dd')) "
                       
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
            
            strSQL = strSQL & _
                       "AND D.���� LIKE [1] "
            
        Case 2                          '��ȫ��ĸ�����������
        
            strSQL = strSQL & _
                       "AND Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE [2]) And B.ҩƷid=bb.�շ�ϸĿID) "
            
        Case Else
        
            strSQL = strSQL & _
                       "AND (D.���� LIKE [1] OR D.���� LIKE [2] OR Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� LIKE [2] OR bb.���� LIKE [2]) And B.ҩƷid=bb.�շ�ϸĿID )) "
            
        End Select
        strSQL = strSQL & " ORDER BY D.����,D.����"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.ҩƷ��Ŀѡ��
        
        strSQL = "SELECT ���,DECODE(ҩƷID,0,NULL,ҩƷID) AS ҩƷID,ID,���,DECODE(�ϼ�ID,0,NULL,�ϼ�ID) AS �ϼ�ID,ĩ��,����,����,��λ,����,��� " & _
                     "FROM (SELECT DECODE(����, '5', 'A1', '6', 'A2', 'A3') AS ���, " & _
                                  "0 AS ҩƷID, " & _
                                  "DECODE(����, '5', -1, '6', -2, -3) AS ID, " & _
                                  "0 AS �ϼ�ID, " & _
                                  "0 AS ĩ��, " & _
                                  "����, " & _
                                  "����, NULL AS ��λ," & _
                                  "NULL AS ���," & _
                                  "NULL AS ����,'' As ��� " & _
                             "From ������Ŀ��� WHERE ���� IN ('5','6','7') " & _
                           "Union All " & _
                             "SELECT DECODE(����, 1, 'A1', 2, 'A2', 'A3') || TO_CHAR(ROWNUM,'0000000000') AS ���, " & _
                                    "0 AS ҩƷID, " & _
                                    "ID, " & _
                                    "DECODE(�ϼ�ID,NULL,DECODE(����, 1, -1, 2, -2, -3),�ϼ�ID) AS �ϼ�ID, " & _
                                    "0 AS ĩ��, " & _
                                    "����, " & _
                                    "����, NULL AS ��λ," & _
                                    "NULL AS ���," & _
                                    "NULL AS ����,'' As ��� " & _
                               "From ���Ʒ���Ŀ¼  Where DECODE(����,1,'5',2,'6','7') IN ('5','6','7') "
        strSQL = strSQL & _
                            "Start With �ϼ�ID is NULL " & _
                             "Connect by Prior ID = �ϼ�ID " & _
                             "Union All " & _
                                "SELECT 'B1' AS ���, " & _
                                      "B.ҩƷID, " & _
                                      "B.ҩ��ID AS ID, " & _
                                      "C.����ID AS �ϼ�ID," & _
                                      "1 as ĩ��, " & _
                                      "D.����, " & _
                                      "D.����, " & _
                                      "d.���㵥λ AS ��λ, " & _
                                      "D.���," & _
                                      "A.ҩƷ���� AS ����,d.��� " & _
                                 "FROM ҩƷ���� A,ҩƷ��� B,������ĿĿ¼ C,�շ���ĿĿ¼ D " & _
                                "WHERE A.ҩ��id=B.ҩ��id " & _
                                        "AND C.ID=A.ҩ��id " & _
                                        "AND D.ID=B.ҩƷid " & _
                                        "AND C.��� IN ('5','6','7') " & _
                                        "AND (D.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.����ʱ�� is NULL) " & _
                           ") " & _
                    "ORDER BY ���"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.ҩƷ��Ŀ����
    
        strSQL = "SELECT Decode(C.���, '5', '����ҩ', '6', '�г�ҩ', '�в�ҩ') AS ����," & _
                       "D.����," & _
                       "D.����," & _
                       "D.���," & _
                       "A.ҩƷ���� AS ����," & _
                       "d.���㵥λ AS ��λ," & _
                       "B.ҩƷID," & _
                       "B.ҩƷID AS ID," & _
                       "A.ҩ��ID,d.��� " & _
                 "FROM ҩƷ���� A, ҩƷ��� B, ������ĿĿ¼ C, �շ���ĿĿ¼ D " & _
                 "WHERE A.ҩ��ID = B.ҩ��ID " & _
                       "AND C.ID = A.ҩ��ID " & _
                       "AND D.ID = B.ҩƷID " & _
                       "AND C.��� IN ('5','6','7') " & _
                       "AND (D.����ʱ�� IS NULL OR D.����ʱ�� = TO_DATE('3000-01-01', 'yyyy-MM-dd')) "
                       
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
            strSQL = strSQL & _
                       "AND D.���� LIKE [1] "
        Case 2                          '��ȫ��ĸ�����������
            strSQL = strSQL & _
                       "AND Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE [2]) And B.ҩƷid=bb.�շ�ϸĿID) "
        Case Else
            strSQL = strSQL & _
                       "AND (D.���� LIKE [1] OR D.���� LIKE [2] OR Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE [2]) And B.ҩƷid=bb.�շ�ϸĿID)) "
            
        End Select
        strSQL = strSQL & " ORDER BY D.����,D.����"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀѡ��
    
        strSQL = "Select ID," & _
                    "�ϼ�ID," & _
                    "0 as ĩ��," & _
                    "����," & _
                    "����," & _
                    "'' as ���,'' As ����," & _
                    "'' as ��λ," & _
                    "0 AS �Ƿ���,0 As ����ּ�,0 As ����ּ�,'' As ���� " & _
              "From ���Ʒ���Ŀ¼ " & _
              "where ����=7 " & _
              "Start With �ϼ�ID is NULL " & _
                "Connect by Prior ID = �ϼ�ID " & _
                "Union All "
        strSQL = strSQL & _
                  "Select A.����ID AS ID, " & _
                     "C.����id AS �ϼ�ID, " & _
                     "1 as ĩ��, " & _
                     "B.����, " & _
                     "B.����, " & _
                     "B.���,B.����, " & _
                     "B.���㵥λ as ��λ, " & _
                     "B.�Ƿ���,D.ԭ�� as ����ּ�,D.�ּ� as ����ּ�,DECODE(B.�Ƿ���,1,TRIM(TO_CHAR(D.ԭ��,'999999990.99'))||'��'||TRIM(TO_CHAR(D.�ּ�,'999999990.99')),TRIM(TO_CHAR(D.�ּ�,'999999990.99'))) as ���� " & _
                "FROM �������� A,�շ���ĿĿ¼ B,������ĿĿ¼ C,�շѼ�Ŀ d  " & _
               "Where A.����id=B.ID AND (B.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or B.����ʱ�� is NULL) " & _
                    "AND C.ID=A.����id And d.ִ������<=SYSDATE AND (d.��ֹ����>=SYSDATE OR d.��ֹ���� IS NULL)"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ŀ����
            
        strSQL = "SELECT A.�Ƿ���,B.ԭ�� as ����ּ�,B.�ּ� as ����ּ�,C.���� AS ���,A.����,A.����,A.���,A.����,A.���㵥λ,DECODE(A.�Ƿ���,1,TRIM(TO_CHAR(B.ԭ��,'999999990.99'))||'��'||TRIM(TO_CHAR(B.�ּ�,'999999990.99')),TRIM(TO_CHAR(B.�ּ�,'999999990.99'))) as ����,A.ID  " & _
                    "FROM �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C " & _
                    "WHERE C.����=A.��� " & _
                            "AND A.ID=B.�շ�ϸĿID " & _
                            "AND A.���='4' " & _
                            "AND B.ִ������<=SYSDATE AND (B.��ֹ����>=SYSDATE OR B.��ֹ���� IS NULL) " & _
                            "AND (A.����ʱ�� IS NULL OR A.����ʱ��=TO_DATE('3000-01-01','yyyy-MM-dd'))"
                                
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
        
            strSQL = strSQL & " AND A.���� Like [1] "
            
        Case 2                          '��ȫ��ĸ�����������
            
            strSQL = strSQL & " AND Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE [2]) And a.ID=bb.�շ�ϸĿID) "
        Case Else

            strSQL = strSQL & " AND (A.���� Like [1] or A.���� Like [2] Or Exists (SELECT 1 FROM �շ���Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE [2]) And a.ID=bb.�շ�ϸĿID))"

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����ʽѡ��
        
        strSQL = "SELECT ����,����,���㵥λ AS ��λ,�������� AS ��������,ID FROM ������ĿĿ¼ a WHERE ���='G' "
    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����ʽ����

        strSQL = "SELECT ����,����,���㵥λ AS ��λ,�������� AS ��������,ID FROM ������ĿĿ¼ a WHERE ���='G' "
                                                
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
        
            strSQL = strSQL & " AND A.���� Like [1] "
            
        Case 2                          '��ȫ��ĸ�����������
            
            strSQL = strSQL & " AND Exists (SELECT 1 FROM ������Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE  [2]) And a.ID=bb.������Ŀid) "

        Case Else

            strSQL = strSQL & " AND (A.���� Like [1] or A.���� Like [2] Or Exists (SELECT 1 FROM ������Ŀ���� bb WHERE (bb.���� Like [2] Or bb.���� LIKE  [2]) And a.ID=bb.������Ŀid))"

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Ա��Ϣѡ��
                
        strSQL = "SELECT   A.���," & _
                           "A.����," & _
                           "A.����," & _
                           "C.���� AS ����," & _
                           "A.ID " & _
                    "FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D " & _
                    "WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID AND D.ȱʡ=1 " & _
                        "AND B.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
        
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Ա��Ϣ����
    
        strSQL = "SELECT   A.���," & _
                           "A.����," & _
                           "A.����," & _
                           "C.���� AS ����," & _
                           "A.ID " & _
                    "FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D " & _
                    "WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID AND D.ȱʡ=1 " & _
                        "AND B.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
        
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
        
            strSQL = strSQL & " AND A.��� Like [3] "
            
        Case 2                          '��ȫ��ĸ�����������
            
            strSQL = strSQL & " AND A.���� LIKE  [4]) "

        Case Else

            strSQL = strSQL & " AND (A.��� LIKE [3] OR A.���� LIKE [4] OR A.���� LIKE [4]) "

        End Select
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ϣѡ��
                
        
        strSQL = "SELECT a.����,a.����,a.����,a.ID FROM ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������=[1]"
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������Ϣ����
    
        strSQL = "SELECT a.����,a.����,a.����,a.ID FROM ���ű� a,��������˵�� b Where a.ID=b.����id And b.��������=[1]"
        
        
        Select Case Val(varParam(0))
        Case 1                          '��ȫ���֣����������
        
            strSQL = strSQL & " AND A.���� Like [2] "
            
        Case 2                          '��ȫ��ĸ�����������
            
            strSQL = strSQL & " AND A.���� LIKE  [3]) "

        Case Else

            strSQL = strSQL & " AND (A.���� LIKE [2] OR A.���� LIKE [3] OR A.���� LIKE [3]) "

        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Ա����ѡ��
            
        strSQL = "SELECT   A.���," & _
                           "A.����," & _
                           "A.����," & _
                           "C.���� As ����," & _
                           "Decode(e.��Աid,Null,'����',Decode(e.����״̬,2,'Ԥ��',3,'����')) As ״̬," & _
                           "A.ID " & _
                    "FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D, " & _
                                        "(SELECT AA.��Աid,bb.����״̬ " & _
                                         "FROM ����������Ա AA," & _
                                                "����������¼ BB, " & _
                                                "����������¼ DD " & _
                                        "WHERE AA.��¼ID = BB.ID " & _
                                                "AND BB.ҽ��id <> [3] " & _
                                                "AND BB.����״̬ In (2,3) " & _
                                                "AND DD.ҽ��id = [3] " & _
                                                "AND NOT (DD.������ʼʱ�� > BB.��������ʱ�� OR DD.��������ʱ�� < BB.������ʼʱ��)) e " & _
                    "WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID AND D.ȱʡ=1 " & _
                        "AND A.ID=e.��Աid(+) And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                        "AND B.��Ա����=[1] And ((b.��Ա����='��ʿ' And c.ID=[2]) Or b.��Ա����<>'��ʿ') " & _
                        "ORDER BY Decode(c.ID,[2],1,2)"
                        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Ա���Ź���
        
        strSQL = "SELECT   A.���," & _
                           "A.����," & _
                           "A.����," & _
                           "C.���� AS ����," & _
                           "Decode(e.��Աid,Null,'����',Decode(e.����״̬,2,'Ԥ��',3,'����')) As ״̬," & _
                           "A.ID " & _
                    "FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D, " & _
                                        "(SELECT AA.��Աid,bb.����״̬ " & _
                                         "FROM ����������Ա AA," & _
                                                "����������¼ BB, " & _
                                                "����������¼ DD " & _
                                        "WHERE AA.��¼ID = BB.ID " & _
                                                "AND BB.ҽ��id <> [3] " & _
                                                "AND BB.����״̬ In (2,3) " & _
                                                "AND DD.ҽ��id = [3] " & _
                                                "AND NOT (DD.������ʼʱ�� > BB.��������ʱ�� OR DD.��������ʱ�� < BB.������ʼʱ��)) e " & _
                    "WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID AND D.ȱʡ=1 AND A.ID=e.��Աid(+) And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                        "AND b.��Ա����=[1] And ((b.��Ա����='��ʿ' And c.ID=[2]) Or b.��Ա����<>'��ʿ')  "
                        
'        Select Case Val(varParam(0))
'        Case 1                          '��ȫ���֣����������
'
'            strSQL = strSQL & " AND A.��� Like [4] "
'
'        Case 2                          '��ȫ��ĸ�����������
'
'            strSQL = strSQL & " AND A.���� LIKE  [5]) "
'
'        Case Else

            strSQL = strSQL & " AND (A.��� LIKE [4] OR A.���� LIKE [5] OR A.���� LIKE [5]) "

'        End Select
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�����������
    
        strSQL = "SELECT DECODE(A.������ĿID,null,'2-����','1-����') AS ���뷽ʽ," & _
                        "A.��������," & _
                        "A.ȱʡ," & _
                        "DECODE(A.������Ŀid,Null,A.��������ID,A.������Ŀid) As ID " & _
                    "FROM ����������� A " & _
                    "WHERE A.��¼id=[1] " & _
                            "AND A.����=[2] "
                            
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������ϼ�¼
        
        strSQL = "Select 1 AS ID,���id," & _
                         "����id," & _
                         "c.���� As ��ϱ���," & _
                         "d.���� As ��������," & _
                         "������� " & _
                    "From ������ϼ�¼ a,����������¼ b,�������Ŀ¼ c,��������Ŀ¼ d " & _
                   "where a.ҽ��id = b.ҽ��id and ������� = [2] And b.ID=[1] And c.ID(+)=a.���id And d.ID(+)=a.����id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�շ�ִ�п���
    
        Select Case varParam(0)
        Case "5", "6", "7", "4"
            strSQL = "Select a.ID,a.����,a.����" & vbNewLine & _
                        "From  ���ű� a" & vbNewLine & _
                        "Where (a.����ʱ�� IS NULL OR a.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
                        "    And a.ID In (   Select  b.����id" & vbNewLine & _
                        "            From    ��������˵�� b" & vbNewLine & _
                        "            Where   b.������� in (1,2,3)" & vbNewLine & _
                        "                And b.��������=Decode('" & varParam(0) & "','5','��ҩ��','6','��ҩ��','7','��ҩ��','4','���ϲ���')) " & _
                        " Order By Decode(a.ID,[1],0,1) "

        Case Else
            strSQL = "Select * From (" & vbNewLine & _
                        "Select  a.ID,a.����,a.����,1 As ĩ��,b.OrderCol" & vbNewLine & _
                        "From    ���ű� a," & vbNewLine & _
                        "    (" & vbNewLine & _
                        "    Select a.ID,1 As OrderCol From ���ű� a,�շ���ĿĿ¼ X Where X.ID=[2] And X.ִ�п���=1 And A.ID=[3]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,2 As OrderCol From ���ű� a,�������Ҷ�Ӧ B,�շ���ĿĿ¼ X Where X.ID=[2] And X.ִ�п���=2 And A.ID=B.����id And B.����ID=[3]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,3 As OrderCol From ���ű� a,�շ���ĿĿ¼ X Where X.ID=[2] And X.ִ�п���=3 And A.ID=[4]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,4 As OrderCol From ���ű� a,�շ�ִ�п��� B,�շ���ĿĿ¼ X Where X.ID=[2] And X.ִ�п���=4 And A.ID=B.ִ�п���id And B.������Դ=1 And B.�շ�ϸĿid=X.ID" & vbNewLine & _
                        "    ) b," & vbNewLine & _
                        "    ��������˵�� c" & vbNewLine & _
                        "Where a.ID = c.����ID" & vbNewLine & _
                        "    And c.������� In (1,2,3)" & vbNewLine & _
                        "    And a.ID=b.ID(+)" & vbNewLine & _
                        "Order By Decode(a.ID,[1],0,b.OrderCol)" & vbNewLine & _
                        ") "

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����������
        
        strSQL = "Select Nvl(Sum(A.��������),0) As ��� From ҩƷ��� A Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate)) And A.����=1 And A.ҩƷID=[1] And A.�ⷿID=[2]"
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�����鷽ʽ
        
        If varParam(0) = "4" Then
            strSQL = "Select ��鷽ʽ From ���ϳ����� Where �ⷿID=[1]"
        Else
            strSQL = "Select ��鷽ʽ From ҩƷ������ Where �ⷿID=[1]"
        End If
    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������ҩ�ο�
        
        strSQL = "SELECT D.ID,d.���,D.���,B.���㵥λ As ��λ,A.���� As ��ҩ����,D.���� AS ҩƷ����,A.����,A.���� As ׼������ " & _
                "FROM ������ҩ�ο� A,������ĿĿ¼ B,ҩƷ��� C,�շ���ĿĿ¼ D " & _
                "WHERE A.ҩ��id=C.ҩƷid AND B.ID=C.ҩ��id AND D.ID=C.ҩƷid AND A.����ID=[1]"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�������ϲο�

        strSQL = "SELECT B.���㵥λ As ��λ,B.ID,B.����,A.����,A.���� As ׼������,B.��� " & _
                "FROM �������ϲο� A,�շ���ĿĿ¼ B " & _
                "WHERE A.����id=B.ID AND A.����ID=[1]"

                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�������Ʋο�
    
        strSQL = "SELECT D.ID,b.���� As ���,d.���㵥λ,D.����,A.����,d.��� As ������ " & _
                "FROM �������Ѳο� A,�շ���ĿĿ¼ D,�շ���Ŀ��� B " & _
                "WHERE A.ϸĿid=D.ID And B.����=d.��� And A.����ID=[1]"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.������ҩѡ��
        
        
        strSQL = "Select    Null AS ����," & _
                           "ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� " & _
                     "WHERE ID IN (Select ����ID From ������������ A, ����������¼ B, ����ҽ����¼ C WHERE A.������ĿID = C.������Ŀid AND B.ҽ��id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                            
        strSQL = strSQL & _
                "Select Null AS ����," & _
                           "ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� b " & _
                     "WHERE Not Exists (Select 1 From ������������ a Where a.����id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                 "SELECT A.����, " & _
                 "      D.ID, " & _
                 "      d.����, " & _
                 "      d.����, " & _
                 "      d.���㵥λ AS ��λ, " & _
                 "      D.���, " & _
                 "      TRIM(TO_CHAR(a.����, '9999999990.99')) AS ����, " & _
                 "      A.����ID AS �ϼ�id, " & _
                 "      1 AS ĩ��  " & _
                 " FROM ������ҩ�ο� A, ������ĿĿ¼ B,ҩƷ��� C,�շ���ĿĿ¼ D  " & _
                 " WHERE A.ҩ��id = C.ҩƷid AND C.ҩ��id=B.id AND D.ID=C.ҩƷID"
                 
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��������ѡ��
        
        strSQL = "Select   ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� " & _
                     "WHERE ID IN (Select ����ID From ������������ A, ����������¼ B, ����ҽ����¼ C WHERE A.������ĿID = C.������Ŀid AND B.ҽ��id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                            
        strSQL = strSQL & _
                "Select    ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� b " & _
                     "WHERE Not Exists (Select 1 From ������������ a Where a.����id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                      "SELECT ROWNUM AS ID," & _
                             "B.����," & _
                             "B.����," & _
                             "B.���,B.���㵥λ AS ��λ," & _
                             "TRIM(TO_CHAR(A.����, '9999999990.99')) AS ����," & _
                             "����ID AS �ϼ�id," & _
                             "1 AS ĩ�� " & _
                        "FROM �������ϲο� A, �շ���ĿĿ¼ B " & _
                       "WHERE A.����id = B.id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��������ѡ��
    
        strSQL = "Select   ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� " & _
                     "WHERE ID IN (Select ����ID From ������������ A, ����������¼ B, ����ҽ����¼ C WHERE A.������ĿID = C.������Ŀid AND B.ҽ��id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                "Select    ID," & _
                           "����," & _
                           "����," & _
                           "Null AS ��λ,NULL AS ���," & _
                           "Null AS ����," & _
                           "0 As �ϼ�id," & _
                           "0 As ĩ�� " & _
                      "From ���������ο� b " & _
                     "WHERE Not Exists (Select 1 From ������������ a Where a.����id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                      "SELECT ROWNUM AS ID," & _
                             "B.����," & _
                             "B.����," & _
                             "B.���,B.���㵥λ AS ��λ," & _
                             "TRIM(TO_CHAR(A.����, '9999999990.99')) AS ����," & _
                             "����ID AS �ϼ�id," & _
                             "1 AS ĩ�� " & _
                        "FROM �������Ѳο� A, �շ���ĿĿ¼ B " & _
                       "WHERE A.ϸĿID = B.id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.�ٴ����ż�¼
        
        If varParam(0) = "����" Then
        
            strSQL = "SELECT A.����||'-'||A.���� As ����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='�ٴ�' ORDER BY A.����||'-'||A.����"
        
        Else
            strSQL = "SELECT A.����||'-'||A.���� As ����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='�ٴ�' " & _
                        "AND A.ID IN (SELECT ����id FROM ������Ա WHERE ��Աid=" & UserInfo.ID & ")  ORDER BY A.����||'-'||A.����"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Լ��λѡ��

        strSQL = "SELECT -1 AS ID,NULL+0 AS �ϼ�id,'0' AS ����,'����' AS ����,'' as ����,'' as ��ַ,0 AS ĩ��,'' AS ��ϵ��,'' AS �绰,'' AS �����ʼ�,'' AS ��������,'' AS �ʺ�,'' AS ��ַ,'' AS ˵�� from dual " & _
                    "Union All " & _
                    "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,��ַ,0 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ  where ĩ��<>1 " & _
                    "Start With �ϼ�id is null connect by prior ID=�ϼ�id " & _
                    "Union All " & _
                    "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,��ַ,1 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ  where ĩ��=1"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Լ��λ����
        
            strSQL = "select ID,����,����,����,��ַ,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ  where ĩ��=1 " & _
                " AND (���� Like [1] or ���� Like [1] OR ���� Like [1])"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.����ҽ����Ա
    
        strSQL = _
            "Select Distinct A.����,A.ID,B.����ID,A.���,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
            " And C.��Ա���� IN('ҽ��') And B.����ID=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
        strSQL = strSQL & " Order by ����,��Ա���� Desc"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.��Ա����ѡ��
        
            strSQL = "SELECT C.����id AS ID,C.��ǰ����id," & _
                    "C.����," & _
                    "C.�����," & _
                    "C.����," & _
                    "C.�Ա�," & _
                    "C.��������," & _
                    "C.���֤��," & _
                    "C.����״��,c.ְҵ,c.סԺ��,c.����,c.����,c.ҽ�Ƹ��ʽ,c.�ѱ�, " & _
                    "C.��ͬ��λid,c.������λ,c.��ϵ������,c.��ϵ�˵绰,c.��ͥ��ַ,c.��ͥ�绰,c.�����ʱ�,c.��λ�绰,c.��λ�ʱ�,b.��ҳid " & _
                "FROM ������Ϣ C,������ҳ b " & _
                "WHERE c.����id=b.����id(+) And b.��Ժ���� Is Null " & IIf(strParam = "'", "", strParam)

            
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function


