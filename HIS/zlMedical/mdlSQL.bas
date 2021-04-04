Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL

    ���˻�����Ϣ
    ��λ��Աѡ��
    ��첿���嵥
    
'    ��������ѡ��
'    �����������
    
    �շ���Ŀѡ��
    �շ���Ŀ����
    �����Ŀѡ��
    �����Ŀ����ѡ��
    �����Ŀ�嵥
    �����Ա����
    �����Ա����_����
    ������ͷ���
    ������ͷ���ѡ��
    �������ѡ��
    ������͹���ѡ��
    ���������Ŀ
    ������ͼƼ�
    ����������Ŀ
    �����Ϸ���
    �������ѡ��
    ������Ŀѡ��
    ������Ŀ����ѡ��
    ��Ա�����Ŀ
    ��Աԭʼ��Ŀ
    ���������Ŀ
    �������ѡ��
    ���������Ŀ����ѡ��
    ���������Ŀѡ��
    ����δ����ϸ
    
    ���˷��øſ�
    ������øſ�
    
    �����Ŀ�۱�
    ���ԤԼ����
    ���Ǽǵ���
    ����ִ�п���
    �շ�ִ�п���
    ҩƷִ�п���
    ��Ա����
    ��������Ա
    ��������Ա1
    ��Ա����ѡ��
    �������ͳ��
End Enum

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String, Optional ByVal blnMoveOuted As Boolean = False) As String
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
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��Ա����
            strSQL = "select A.����id AS ID,A.����id,A.����,A.���֤�� AS ���֤,A.�Ա�,A.����,TO_CHAR(A.��������,'yyyy-mm-dd') AS ��������,A.����״��," & _
                    "A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.��ϵ������,A.��ϵ�˵绰,A.��ϵ�˵�ַ,A.������λ " & _
                    "from ������Ϣ A " & _
                    "WHERE A.����ID=[1]"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��Ա����ѡ��
            
            strSQL = "SELECT C.����id AS ID," & _
                    "C.����," & _
                    "C.�����," & _
                    "C.������," & _
                    "C.����," & _
                    "C.�Ա�," & _
                    "TO_CHAR(C.��������,'yyyy-mm-dd') AS ��������," & _
                    "C.���֤��," & _
                    "C.����״��, " & _
                    "C.��ͬ��λid " & _
                "FROM ������Ϣ C " & _
                "WHERE 1=1 " & strParam
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��λ��Աѡ��
        
            strSQL = "Select * From (SELECT 1 As ѡ��,C.����id AS ID," & _
                    "C.����," & _
                    "C.�����," & _
                    "C.������," & _
                    "Decode(c.��������,Null,Decode(c.����,Null,0,Decode(Trim(Substr(c.����,Length(c.����),1)),'��',Zl_To_Number(Substr(c.����,1,Length(c.����)-1)),'��',Zl_To_Number(Substr(c.����,1,Length(c.����)-1))/12,'��',Zl_To_Number(Substr(c.����,1,Length(c.����)-1))/365,Zl_To_Number(c.����))),Trunc(Months_between(Sysdate,c.��������)/12)) As ����," & _
                    "C.�Ա�," & _
                    "TO_CHAR(C.��������,'yyyy-mm-dd') AS ��������," & _
                    "C.���֤��,����,����,ѧ��,ְҵ,���,��ϵ������,��ϵ�˵绰,��ϵ�˵�ַ,������λ,IC����,���￨��," & _
                    "C.����״��, " & _
                    "C.��ͬ��λid " & _
                "FROM ������Ϣ c Where ��ͬ��λid In (Select ID From ��Լ��λ Start With ID=[1] Connect by Prior ID=�ϼ�id)) " & _
                "WHERE (Instr(�Ա�,[2])>0 Or [2] Is Null) And ���� Between [3] And [4] "

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���˻�����Ϣ
        
            strSQL = "SELECT A.ID," & _
                     "A.����," & _
                     "D.�����,D.������,D.���￨��,D.IC����," & _
                     "D.����," & _
                     "D.�Ա�," & _
                     "D.����," & _
                     "D.����״��," & _
                     "C.���ʱ��," & _
                     "C.��첡��id,C.����ʱ��,C.�������,A.�������,D.��ϵ�˵绰,D.������λ," & _
                     "B.���� AS �������� " & _
                "FROM ���ǼǼ�¼ A,��Լ��λ B,�����Ա���� C,������Ϣ D " & _
                "WHERE A.ID=C.�Ǽ�id AND A.��Լ��λID=B.ID(+) AND D.����id=C.����id AND C.ID=[1] "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��첿���嵥
        
            If strParam = "����" Then
            
                strSQL = "SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='���' ORDER BY A.����||'-'||A.����"
            
            Else
                strSQL = "SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='���' " & _
                            "AND A.ID IN (SELECT ����id FROM ������Ա WHERE ��Աid=[1])  ORDER BY A.����||'-'||A.����"
            End If
        
'        Case SQL.��������ѡ��
'
'            strSQL = "Select * " & _
'                     "from (Select 0 As ѡ��,ID,�ϼ�ID,0 as ĩ��,���,'' As ����,���� ,'' as ����,'' AS ���� " & _
'                             "From ����������� " & _
'                            "Where ��� = 'D' " & _
'                            "Start With �ϼ�id Is Null Connect by Prior ID = �ϼ�ID " & _
'                           "Union All " & _
'                             "Select 0 As ѡ��,A.ID,A.����id AS �ϼ�ID,1 as ĩ��,0 As ���, A.����,A.����,A.����,A.���� " & _
'                               "FROM ��������Ŀ¼ A " & _
'                              "Where A.���='D' " & _
'                           ") A Order by A.ĩ��,A.��� "
'        Case SQL.�����������
'
'            varParam(0) = "'%" & UCase(varParam(0)) & "%'"
'            strSQL = "SELECT A.ID,A.����,A.����,A.����,A.���� " & _
'                        "FROM ��������Ŀ¼ A " & _
'                        "WHERE A.��� ='D' "
'
'            strSQL = strSQL & " AND (UPPER(A.����) Like " & varParam(0) & " OR A.���� Like " & varParam(0) & " OR A.���� Like " & varParam(0) & ")"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�շ���Ŀѡ��
    
            strSQL = "select * " & _
                     "from (Select ID,�ϼ�ID,0 as ĩ��,����,���� ,'' as ��λ,'' AS ���,'' As ����,'' As ��� " & _
                             "From �շѷ���Ŀ¼ " & _
                            "Start With �ϼ�ID Is Null " & _
                           "Connect by Prior ID = �ϼ�ID "
            
            strSQL = strSQL & " Union All Select -1 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-1' As ����,'����ҩ' As ���� ,'' as ��λ,'' AS ���,'' As ����,'' As ��� from dual "
            strSQL = strSQL & " Union All Select -2 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-2' As ����,'�г�ҩ' As ���� ,'' as ��λ,'' AS ���,'' As ����,'' As ��� from dual "
            strSQL = strSQL & " Union All Select -3 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-3' As ����,'�в�ҩ' As ���� ,'' as ��λ,'' AS ���,'' As ����,'' As ��� from dual "
            strSQL = strSQL & " Union All Select -7 As ID,Null+0 As �ϼ�ID,0 as ĩ��,'-7' As ����,'��������' As ���� ,'' as ��λ,'' AS ���,'' As ����,'' As ��� from dual "
             
            strSQL = strSQL & _
                           "Union All " & _
                             "Select A.ID,Decode(A.���,'5',-1,'6',-2,'7',-3,'4',-7,A.����id) AS �ϼ�ID,1 as ĩ��, A.����,A.����,A.���㵥λ AS ��λ,A.���,Trim(To_Char(c.����,'9999999999999.00000')) As ����,a.��� " & _
                               "FROM �շ���ĿĿ¼ A,�շ���Ŀ��� B,(select �շ�ϸĿid,sum(�ּ�) AS ���� from �շѼ�Ŀ where ִ������<=SYSDATE and (��ֹ���� IS NULL OR ��ֹ����>SYSDATE) group by �շ�ϸĿid) C " & _
                              "Where C.�շ�ϸĿid(+)=A.ID AND Nvl(a.�Ƿ���,0)=0 And A.���=b.���� AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) " & _
                           ") A " & _
                    "ORDER BY A.ĩ��, A.����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�շ���Ŀ����
                        
            If CheckStrType(varParam(0), 1) And Left(ParamInfo.�շ�������Ŀƥ��, 1) = 1 Then
                '��ȫ���֣����������
                
                strSQL = "SELECT A.ID,A.����,A.����,A.���㵥λ AS ��λ,A.���,Trim(To_Char(c.����,'9999999999999.00000')) As ����,a.��� " & _
                            "FROM �շ���ĿĿ¼ a,�շ���Ŀ��� b,(select �շ�ϸĿid,sum(�ּ�) AS ���� from �շѼ�Ŀ where ִ������<=SYSDATE and (��ֹ���� IS NULL OR ��ֹ����>SYSDATE) group by �շ�ϸĿid) c " & _
                            "WHERE c.�շ�ϸĿid(+)=a.ID and Nvl(a.�Ƿ���,0)=0 And a.���=b.���� AND (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� is NULL) "
                        
                strSQL = strSQL & " AND a.���� Like [1]"
                
            ElseIf CheckStrType(varParam(0), 2) And Left(ParamInfo.�շ�������Ŀƥ��, 2) = 1 Then
                '��ȫ��ĸ�����������

                strSQL = "SELECT Distinct A.ID,A.����,A.����,A.���㵥λ AS ��λ,A.���,Trim(To_Char(c.����,'9999999999999.00000')) As ����,a.��� " & _
                            "FROM �շ���ĿĿ¼ a,�շ���Ŀ��� b,�շ���Ŀ���� d,(select �շ�ϸĿid,sum(�ּ�) AS ���� from �շѼ�Ŀ where ִ������<=SYSDATE and (��ֹ���� IS NULL OR ��ֹ����>SYSDATE) group by �շ�ϸĿid) c " & _
                            "WHERE c.�շ�ϸĿid(+)=a.ID and Nvl(a.�Ƿ���,0)=0 And a.���=b.���� AND (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� is NULL) "
                        
                strSQL = strSQL & " AND a.ID=d.�շ�ϸĿID AND [1] Is Not Null And d.���� Like [2]"
                
            Else
                strSQL = "SELECT Distinct A.ID,A.����,A.����,A.���㵥λ AS ��λ,A.���,Trim(To_Char(c.����,'9999999999999.00000')) As ����,a.��� " & _
                            "FROM �շ���ĿĿ¼ a,�շ���Ŀ��� b,�շ���Ŀ���� d,(select �շ�ϸĿid,sum(�ּ�) AS ���� from �շѼ�Ŀ where ִ������<=SYSDATE and (��ֹ���� IS NULL OR ��ֹ����>SYSDATE) group by �շ�ϸĿid) c " & _
                            "WHERE c.�շ�ϸĿid(+)=a.ID and Nvl(a.�Ƿ���,0)=0 And a.���=b.���� AND (a.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or a.����ʱ�� is NULL) "
                        
                strSQL = strSQL & " AND A.ID=d.�շ�ϸĿID AND (a.���� Like [1] OR a.���� Like [2] Or d.���� Like [2] Or d.���� Like [2])"
            End If

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ŀѡ��
            
            strSQL = "select * " & _
                     "from (Select DISTINCT 0 As ѡ��,ID,�ϼ�ID,0 as ĩ��,����,���� ,'' as ��λ,'' AS ���,'' As �걾��λ, " & _
                                           "DECODE(�ϼ�ID, Null, ID * POWER(10, 20), �ϼ�ID * POWER(10, 20) + ID) As ���� " & _
                             "From ���Ʒ���Ŀ¼ " & _
                            "Where ���� = 5 " & _
                            "Start With ID IN (SELECT DISTINCT ����id FROM ������ĿĿ¼ WHERE (����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or ����ʱ�� is NULL) AND ��� IN ('C','D')) " & _
                           "Connect by Prior �ϼ�ID = ID " & _
                           "Union All " & _
                             "Select 0 As ѡ��,A.ID,A.����id AS �ϼ�ID,1 as ĩ��, A.����,A.����,A.���㵥λ AS ��λ,DECODE(A.���,'C','����','���') AS ���,a.�걾��λ, " & _
                                    "1 AS ���� " & _
                               "FROM ������ĿĿ¼ A " & _
                              "Where A.�����Ա� In (0,[1],[2]) And A.��� IN ('C','D') AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) " & _
                           ") A " & _
                    "ORDER BY A.ĩ��, A.����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ŀ����ѡ��
            
            If CheckStrType(varParam(0), 1) And Left(ParamInfo.�շ�������Ŀƥ��, 1) = 1 Then
                '��ȫ���֣����������
                    
                strSQL = "SELECT A.ID,A.����,A.����,A.���㵥λ AS ��λ,DECODE(A.���,'C','����','���') AS ���,a.�걾��λ " & _
                        "FROM ������ĿĿ¼ A " & _
                        "WHERE A.�����Ա� In (0,[5],[6]) And A.��� IN ([1],[2]) AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) "
                strSQL = strSQL & " AND A.���� Like [3]"
                
            ElseIf CheckStrType(varParam(0), 2) And Left(ParamInfo.�շ�������Ŀƥ��, 2) = 1 Then
                '��ȫ��ĸ�����������
                
                strSQL = "SELECT Distinct A.ID,A.����,A.����,A.���㵥λ AS ��λ,DECODE(A.���,'C','����','���') AS ���,a.�걾��λ " & _
                        "FROM ������ĿĿ¼ A,������Ŀ���� B " & _
                        "WHERE A.�����Ա� In (0,[5],[6]) And A.��� IN ([1],[2]) AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) "
                strSQL = strSQL & " AND A.ID=B.������Ŀid AND [3] Is Not Null And b.���� Like [4]"
                
            Else
            
                strSQL = "SELECT Distinct A.ID,A.����,A.����,A.���㵥λ AS ��λ,DECODE(A.���,'C','����','���') AS ���,a.�걾��λ " & _
                        "FROM ������ĿĿ¼ A,������Ŀ���� B " & _
                        "WHERE A.�����Ա� In (0,[5],[6]) And A.��� IN ([1],[2]) AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) "
                strSQL = strSQL & " AND A.ID=B.������Ŀid AND (A.���� Like [3] OR A.���� Like [4] Or B.���� Like [4] Or B.���� Like [4])"
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.����δ����ϸ
            
            Dim strSub As String
            Dim strCond As String
            Dim blnZero As Boolean
            
            blnZero = (Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������ý��н���", 1)) = 1)
    
            strSQL = "Select Nvl(B.����, 'δ֪') as ����, " & _
                            "A.ʱ��, " & _
                            "A.NO as ���ݺ�, " & _
                            "Nvl(E.����, C.����) as ��Ŀ, " & _
                            "A.�վݷ�Ŀ as ��Ŀ, " & _
                            "A.ID, " & _
                            "A.���, " & _
                            "A.��¼����, " & _
                            "A.��¼״̬, " & _
                            "A.ִ��״̬, " & _
                            "A.��ҳID, " & _
                            "A.��������ID, " & _
                            "A.�Ǽ�ʱ��, " & _
                            "Nvl(A.δ����, 0) δ����, " & _
                            "Nvl(A.δ����, 0) ���ʽ��, " & _
                            "Nvl(A.��������, C.��������) As ���� " & _
                    "From ( "
                    
            strSQL = strSQL & _
                            "SELECT A.ID, " & _
                                     "A.NO, " & _
                                     "A.���, " & _
                                     "A.��¼����, " & _
                                     "A.��¼״̬, " & _
                                     "A.ִ��״̬, " & _
                                     "A.��ҳID, " & _
                                     "A.��������ID, " & _
                                     "To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') as ʱ��, " & _
                                     "A.�Ǽ�ʱ��, " & _
                                     "A.�շ�ϸĿID, " & _
                                     "A.������ĿID, " & _
                                     "A.�վݷ�Ŀ, " & _
                                     "Nvl(A.ʵ�ս��, 0) as δ����, " & _
                                     "�������� "
            
            strCond = " And A.ҽ����� IN (SELECT A.ID FROM ����ҽ����¼ A,���ǼǼ�¼ B WHERE A.�Һŵ� = B.���� and A.������Դ = 4 AND B.��Լ��λID=[1]) "
            
            If blnZero Then
                strSQL = strSQL & _
                                    "From ���˷��ü�¼ A " & _
                                    "Where A.��¼״̬ <> 0 And A.���ʷ��� = 1  And A.����id Is Null " & strCond
            Else
                strSub = _
                        "Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս�� " & _
                        "From ���˷��ü�¼ A " & _
                        "Where A.��¼״̬<>0  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0 And A.����id Is Null " & strCond & _
                        "Group by A.NO,A.���,A.��¼���� " & _
                        "Having Nvl(Sum(A.ʵ�ս��),0)<>0 "
                
                strSQL = strSQL & _
                            "From ���˷��ü�¼ A," & _
                                "(" & strSub & ") B " & _
                            "Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� " & _
                            "And A.��¼״̬<>0 And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0 And A.����id Is Null "
            End If
                                                
            strSQL = strSQL & " Union All " & _
                              "SELECT 0 as ID, " & _
                                     "A.NO, " & _
                                     "A.���, " & _
                                     "Mod(A.��¼����, 10) as ��¼����, " & _
                                     "A.��¼״̬, " & _
                                     "A.ִ��״̬, " & _
                                     "A.��ҳID, " & _
                                     "A.��������ID, " & _
                                     "To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') as ʱ��, " & _
                                     "A.�Ǽ�ʱ��, " & _
                                     "A.�շ�ϸĿID, " & _
                                     "A.������ĿID, " & _
                                     "A.�վݷ�Ŀ, " & _
                                     "Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) as δ����, " & _
                                     "A.�������� " & _
                                "FROM ���˷��ü�¼ A " & _
                                "Where A.����id Is Not Null And A.��¼״̬ <> 0 And A.���ʷ��� = 1 And Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0) " & strCond & "  " & _
                                "Having Sum (Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0 " & _
                                "Group by A.NO, A.���, Mod(A.��¼����, 10), A.��¼״̬, A.ִ��״̬ , A.��ҳID,A.��������ID,To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������"
            
            strSQL = strSQL & ") A," & _
                          "���ű� B," & _
                          "�շ���ĿĿ¼ C," & _
                          "������Ŀ D," & _
                          "�շ���Ŀ���� E " & _
                    "Where A.��������ID = B.ID(+) And A.�շ�ϸĿID = C.ID And A.������ĿID = D.ID And A.�շ�ϸĿID = E.�շ�ϸĿID(+) And E.����(+) = 1 And E.����(+) = 1 " & _
                    "Order by A.ʱ�� Desc, A.NO Desc, A.��¼����, A.���"
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ŀ�嵥
        
            strSQL = "SELECT A.ID," & _
                          "DECODE(A.���, 'C', '����', 'D', '���') AS ���," & _
                          "A.����," & _
                          "B.�����۸�,"
                          
            strSQL = strSQL & _
                          "D.���� as ִ�п���," & _
                          "C.���� as �ɼ���ʽ, " & _
                          "B.�������, " & _
                          "B.�ɼ���ʽid, " & _
                          "DECODE(B.����;��,1,'����','�շ�') AS ����, " & _
                          "B.ִ�п���id, " & _
                          "B.��鲿λ, " & _
                          "B.��鲿λid, " & _
                          "B.����걾, " & _
                          "B.���۸�, " & _
                          "B.������� AS ��� " & _
                     "FROM ������ĿĿ¼ A, �����Ŀ�嵥 B,������ĿĿ¼ C,���ű� D " & _
                    "WHERE B.����id IS NULL AND B.ִ�п���id=D.ID(+) AND B.�ɼ���ʽid=C.ID(+) and  A.ID = B.������ĿID AND B.�Ǽ�id=[1] and B.�������=[2]"
        
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ա����
                    
            strSQL = "Select A.������,A.IC����,TO_CHAR(B.���ʱ��,'yyyy-mm-dd') aS ���ʱ��,B.����,A.����id AS ID,A.����id,A.����,A.�����,A.���֤�� AS ���֤,A.�Ա�,A.����,TO_CHAR(A.��������,'yyyy-mm-dd') AS ��������,A.����״��,C.������� AS ���," & _
                    "A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.��ϵ������,A.��ϵ�˵绰,B.�����ʼ�,A.��ϵ�˵�ַ,A.������λ,a.���￨��,B.�Ǽ�ʱ�� " & _
                    "from ������Ϣ A,�����Ա���� B,(SELECT * FROM ������ WHERE �Ǽ�id=[1]) C  " & _
                    "WHERE A.����ID=B.����ID AND B.�������=C.�������(+) AND B.�Ǽ�id=[1] Order By C.�������,A.����� "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ա����_����
                    
            strSQL = "select A.������,A.IC����,A.����id AS ID,A.����id,A.����,A.�����,A.���֤�� AS ���֤,A.�Ա�,A.����,TO_CHAR(A.��������,'yyyy-mm-dd') AS ��������,A.����״��,C.������� AS ���," & _
                    "A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.��ϵ������,A.��ϵ�˵绰,B.�����ʼ�,A.��ϵ�˵�ַ,A.������λ,a.���￨��,B.�Ǽ�ʱ�� " & _
                    "from ������Ϣ A,�����Ա���� B,(SELECT * FROM ������ WHERE �Ǽ�id=[1]) C  " & _
                    "WHERE A.����ID=B.����ID AND B.�������=C.�������(+) AND B.�Ǽ�id=[1] AND B.����id=[2]"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������ͷ���
        
            strSQL = "SELECT 0 AS ID,NULL+0 AS �ϼ�id,'���з���' AS ����,1 AS ͼ��,1 AS ��ͼ�� FROM dual union all " & _
                    "SELECT ��� AS ID,DECODE(�ϼ����,NULL,0,�ϼ����) AS �ϼ�ID,'['||����||']'||���� AS ����,1 AS ͼ��,1 AS ��ͼ�� " & _
                        "FROM ������� " & _
                        "WHERE NVL(ĩ��,0)=0 " & _
                        "START WITH �ϼ���� IS NULL " & _
                        "CONNECT BY PRIOR ���=�ϼ���� "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������ͷ���ѡ��
                        
            strSQL = "SELECT 0 As ѡ��,-1 AS ID,NULL+0 AS �ϼ�id,'���з���' AS ����,'' AS ����,'' AS ����,'' AS ˵��,1 AS ͼ��,1 AS ��ͼ��,0 AS ĩ�� FROM dual union all " & _
                    "SELECT 0 As ѡ��,��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID,����,����,����,˵��,1 AS ͼ��,1 AS ��ͼ��,ĩ�� " & _
                        "FROM ������� " & _
                        "WHERE ���÷�Χ IN (0,[1]) " & _
                        "START WITH �ϼ���� IS NULL " & _
                        "CONNECT BY PRIOR ���=�ϼ���� "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ϸ���
        
            strSQL = "SELECT 0 AS ID,NULL+0 AS �ϼ�id,'���з���' AS ����,'class' AS ͼ��,'class' AS ��ͼ��,1 AS ����,'0' as ���� FROM dual union all " & _
                    "SELECT ��� AS ID,DECODE(�ϼ����,NULL,0,�ϼ����) AS �ϼ�ID,'['||����||']'||���� AS ����,'class' AS ͼ��,'class' AS ��ͼ��,2 AS ����,���� " & _
                        "FROM �����Ͻ��� " & _
                        "WHERE NVL(ĩ��,0)=0 " & _
                        "START WITH �ϼ���� IS NULL " & _
                        "CONNECT BY PRIOR ���=�ϼ����  order by ����,���� "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�������ѡ��
        
            strSQL = "SELECT -1 AS ID,NULL+0 AS �ϼ�id,'���з���' AS ����,'' AS ����,'' AS ����,'' AS ˵��,1 AS ͼ��,1 AS ��ͼ��,0 AS ĩ�� FROM dual"
            
            strSQL = strSQL & " UNION ALL " & _
                    "SELECT ��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID,'['||����||']'||���� AS ����,����,����,˵��,1 AS ͼ��,1 AS ��ͼ��,ĩ�� " & _
                        "FROM ������� " & _
                        "WHERE NVL(ĩ��,0)=0 " & _
                        "START WITH �ϼ���� IS NULL " & _
                        "CONNECT BY PRIOR ���=�ϼ���� "
                        
            strSQL = strSQL & " UNION ALL " & _
                    "SELECT ��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID,����,����,����,˵��,1 AS ͼ��,1 AS ��ͼ��,ĩ�� " & _
                        "FROM ������� " & _
                        "WHERE NVL(ĩ��,0)=1 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������͹���ѡ��
            
            
            strSQL = "SELECT ��� AS ID,����,����,����,˵�� " & _
                        "FROM ������� " & _
                        "WHERE ĩ��=1 AND (���� LIKE [1] OR ���� LIKE [2] OR ���� LIKE [2])"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�������ѡ��
            
            '����:  1.
            '       2.
            
            strSQL = "select ID,����,����,����,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ " & _
                " Where (���� Like [1] or ���� Like [1] OR ���� Like [1])"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�������ѡ��
            
            '����:  1.frmSchedualEdit\cmd_Click
            '       2.
            
            strSQL = "SELECT -1 AS ID,NULL+0 AS �ϼ�id,'0' AS ����,'����' AS ����,'' as ����,0 AS ĩ��,'' AS ��ϵ��,'' AS �绰,'' AS �����ʼ�,'' AS ��������,'' AS �ʺ�,'' AS ��ַ,'' AS ˵�� from dual " & _
                        "Union All " & _
                        "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,0 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ   " & _
                        "Start With �ϼ�id is null connect by prior ID=�ϼ�id " & _
                        "Union All " & _
                        "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,1 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ "
                                                
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������Ŀѡ��
        
            '����:  1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "(select -1 AS ID,0 AS �ϼ�id,'0' AS ����,'������Ŀ' AS ����,'' AS �ٴ�����,'' AS ��ֵ��,0 AS ĩ��,0 AS ����,0 AS ����,'' AS ��λ from dual UNION ALL " & _
                        "Select DISTINCT ID," & _
                                        "DECODE(�ϼ�ID,NULL,-1,�ϼ�ID) AS �ϼ�ID," & _
                                        "����," & _
                                        "����," & _
                                        "'' as �ٴ�����," & _
                                        "'' as ��ֵ��," & _
                                        "0 as ĩ��," & _
                                        "DECODE(�ϼ�ID,Null,ID * POWER(10, 20),�ϼ�ID * POWER(10, 20) + ID) As ����,0 AS ����,'' AS ��λ " & _
                                  "From ������������ " & _
                                 "Start With ID IN " & _
                                               "( " & _
                                               "SELECT ����id from ����������Ŀ A " & _
                                               "where A.ID IN (SELECT A.������id " & _
                                                              "FROM ���������� A, ����Ԫ��Ŀ¼ B " & _
                                                              "WHERE A.Ԫ��id = B.ID AND B.���� = 2 AND B.���� LIKE '%1') " & _
                                               "Union " & _
                                               "SELECT ����id from ����������Ŀ A " & _
                                               "where A.ID IN (SELECT DISTINCT ������Ŀid from ���鱨����Ŀ A) " & _
                                               ") " & _
                                "Connect by Prior �ϼ�ID = ID) "

            strSQL = strSQL & _
                        "Union All " & _
                        "(SELECT ID, ����id AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,1 AS ĩ��,1 AS ����,����,��λ " & _
                          "from ����������Ŀ A " & _
                         "where A.ID IN " & _
                               "(SELECT A.������id FROM ���������� A, ����Ԫ��Ŀ¼ B WHERE A.Ԫ��id = B.ID AND B.���� = 2 AND B.���� LIKE '%1') " & _
                        "Union " & _
                        "SELECT ID, DECODE(����id,NULL,-1,����id) AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,1 AS ĩ��,1 AS ����,����,��λ " & _
                          "from ����������Ŀ A " & _
                         "where A.ID IN " & _
                               "(SELECT DISTINCT ������Ŀid from ���鱨����Ŀ A)) " & _
                        ") A ORDER BY A.ĩ��,A.����"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������Ŀ����ѡ��
            
            '������:1.
            '       2.
            
'            varParam(0) = "'%" & UCase(varParam(0)) & "%'"
            
            strSQL = "SELECT * FROM (" & _
                        "SELECT ID, ����id AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,Ӣ����,����,��λ " & _
                          "from ����������Ŀ A " & _
                         "where A.ID IN " & _
                               "(SELECT A.������id FROM ���������� A, ����Ԫ��Ŀ¼ B WHERE A.Ԫ��id = B.ID AND B.���� = 2 AND B.���� LIKE '%1') " & _
                        "Union " & _
                        "SELECT ID, DECODE(����id,NULL,-1,����id) AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,Ӣ����,����,��λ " & _
                          "from ����������Ŀ A " & _
                         "where A.ID IN " & _
                               "(SELECT DISTINCT ������Ŀid from ���鱨����Ŀ A) " & _
                        ") A WHERE A.���� LIKE [1] OR A.���� LIKE [2] OR A.Ӣ���� LIKE [2] Or zlSpellCode(A.����) Like [2]  ORDER BY A.����"


        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������Ŀѡ��
            
            '������:1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "(select -1 AS ID,0 AS �ϼ�id,'0' AS ����,'������Ŀ' AS ����,'' AS �ٴ�����,'' AS ��ֵ��,0 AS ĩ��,0 AS ����,0 AS ����,'0' As ���� from dual UNION ALL " & _
                        "Select DISTINCT ID," & _
                                        "DECODE(�ϼ�ID,NULL,-1,�ϼ�ID) AS �ϼ�ID," & _
                                        "����," & _
                                        "����," & _
                                        "'' as �ٴ�����," & _
                                        "'' as ��ֵ��," & _
                                        "0 as ĩ��," & _
                                        "DECODE(�ϼ�ID,Null,ID * POWER(10, 20),�ϼ�ID * POWER(10, 20) + ID) As ����,0 AS ����,���� " & _
                                  "From ������������ " & _
                                "Connect by Prior �ϼ�ID = ID) "

            strSQL = strSQL & _
                        "Union All " & _
                        "(SELECT ID, ����id AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,1 AS ĩ��,1 AS ����,����,'z' As ���� " & _
                          "from ����������Ŀ A ) " & _
                        ") A ORDER BY A.����,A.ĩ��,A.����"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������Ŀ����ѡ��
        
            '������:1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "SELECT ID, ����id AS �ϼ�id, ����, ������ AS ����, �ٴ�����, ��ֵ��,Ӣ����,���� " & _
                          "from ����������Ŀ A " & _
                         "where ����id IS NOT NULL " & _
                        ") A WHERE A.���� LIKE [1] OR A.���� LIKE [2] OR A.Ӣ���� LIKE [2] ORDER BY A.����"
                                 
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��Ա�����Ŀ
            
            '������:1.frmMedicalStation\EditData

            strSQL = "SELECT A.ID,B.ID AS �嵥id,F.�����嵥id," & _
                          "DECODE(A.���, 'C', '����', 'D', '���') AS ���," & _
                          "A.����," & _
                          "Decode(B.����id,NULL,'','����') AS ����," & _
                          "D.���� as ִ�п���," & _
                          "G.���� as �ɼ�����," & _
                          "B.�����۸�,"
                          
            strSQL = strSQL & _
                          "C.���� as �ɼ���ʽ, " & _
                          "E.�������, " & _
                          "B.�ɼ���ʽid, " & _
                          "B.�ɼ�����id, " & _
                          "B.ִ�п���id, " & _
                          "B.��鲿λ, " & _
                          "B.�������, " & _
                          "B.���۸�,Decode(b.�����۸�,0,0,Null,0,10*B.���۸�/B.�����۸�) As �ۿ�," & _
                          "DECODE(B.����;��,1,'����','�շ�') AS ���㷽ʽ, " & _
                          "B.��鲿λid, " & _
                          "Decode(B.����id,NULL,'1','0') AS ����, " & _
                          "B.����걾 " & _
                     "FROM ������ĿĿ¼ A, �����Ŀ�嵥 B,������ĿĿ¼ C,���ű� D,�����Ա���� E,�����Ŀҽ�� F,���ű� G " & _
                    "WHERE B.ִ�п���id=D.ID(+) AND B.�ɼ���ʽid=C.ID(+) AND B.�ɼ�����id=G.ID(+) And  A.ID = B.������ĿID AND E.�Ǽ�id=[1] AND E.�Ǽ�id=B.�Ǽ�id AND E.����id=F.����id AND F.�嵥id=B.ID AND F.����id=[2] "
            
            strSQL = strSQL & " Order By ���� Desc,A.����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��Աԭʼ��Ŀ
            strSQL = "SELECT A.ID,B.ID AS �嵥id," & _
                          "DECODE(A.���, 'C', '����', 'D', '���') AS ���," & _
                          "A.����," & _
                          "Decode(B.����id,NULL,'','����') AS ����," & _
                          "D.���� as ִ�п���," & _
                          "G.���� as �ɼ�����," & _
                          "B.�����۸�,"
                          
            strSQL = strSQL & _
                          "C.���� as �ɼ���ʽ, " & _
                          "E.�������, " & _
                          "B.�ɼ���ʽid, " & _
                          "B.�ɼ�����id, " & _
                          "B.ִ�п���id, " & _
                          "B.��鲿λ, " & _
                          "B.�������, " & _
                          "B.���۸�,Decode(b.�����۸�,0,0,Null,0,10*B.���۸�/B.�����۸�) As �ۿ�," & _
                          "DECODE(B.����;��,1,'����','�շ�') AS ���㷽ʽ, " & _
                          "B.��鲿λid, " & _
                          "Decode(B.����id,NULL,'1','0') AS ����, " & _
                          "B.����걾,0 As ѡ�� " & _
                     "FROM ������ĿĿ¼ A, �����Ŀ�嵥 B,������ĿĿ¼ C,���ű� D,�����Ա���� E,�����Ŀҽ�� F,���ű� G " & _
                    "WHERE B.ִ�п���id=D.ID(+) AND B.�ɼ���ʽid=C.ID(+) AND B.�ɼ�����id=G.ID(+) And  A.ID = B.������ĿID AND E.�Ǽ�id=[1] AND E.�Ǽ�id=B.�Ǽ�id AND E.����id=F.����id AND F.�嵥id=B.ID AND F.����id=[2] and F.�����嵥id Is Null "
            
            strSQL = strSQL & " Order By ���� Desc,A.����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.����������Ŀ
        
            '���������Ա����id

            strSQL = "SELECT A.ID," & _
                          "A.���� AS ��Ŀ,Decode(f.�����嵥id,0,0,Null,0,255) As ǰ��ɫ," & _
                          "Decode(B.����id,NULL,'','����') AS ����," & _
                          "D.���� as ִ�п���," & _
                          "(SELECT DECODE(�����ļ�id,NULL,'','����') FROM ���Ƶ���Ӧ�� WHERE Ӧ�ó���=4 AND ������Ŀid=A.ID) AS ״̬ " & _
                     "FROM ������ĿĿ¼ A, �����Ŀ�嵥 B,���ű� D,�����Ա���� E,�����Ŀҽ�� F " & _
                    "WHERE B.ִ�п���id=D.ID(+)  And  A.ID = B.������ĿID AND E.ID=[1] AND E.�Ǽ�id=B.�Ǽ�id AND E.����id=F.����id AND F.�嵥id=B.ID"
            
            strSQL = strSQL & " Order By D.����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������Ŀ
        
            '������:1.frmMedicalStation\EditData
            '       2.frmSchedual\
                            
            strSQL = _
               "SELECT A.ID,B.ID As �嵥id,0 As �����嵥id, " & _
                  "DECODE(A.���, 'C', '����', 'D', '���') AS ���," & _
                  "A.����," & _
                  "D.���� as ִ�п���," & _
                  "E.���� as �ɼ�����," & _
                  "C.���� as �ɼ���ʽ," & _
                  "B.����걾," & _
                  "B.��鲿λ," & _
                  "B.�ɼ���ʽid," & _
                  "B.�ɼ�����id," & _
                  "B.��鲿λid," & _
                  "B.ִ�п���id," & _
                  "B.�������," & _
                  "B.�������," & _
                  "B.���۸�,'1' As ����," & _
                  "DECODE(B.����;��,1,'����','�շ�') AS ���㷽ʽ," & _
                  "B.�����۸�,Decode(B.�����۸�,0,0,Null,0,10*B.���۸�/B.�����۸�) As �ۿ� "

            strSQL = strSQL & _
                "FROM ������ĿĿ¼ A, " & _
                      "�����Ŀ�嵥 B, " & _
                      "������ĿĿ¼ C, " & _
                      "���ű� D, " & _
                      "���ű� E " & _
                "Where B.������� Is Not Null " & _
                      "AND B.ִ�п���id=D.ID(+) " & _
                      "AND B.�ɼ�����id=E.ID(+) " & _
                      "AND B.�ɼ���ʽid=C.ID(+) " & _
                      "AND A.ID = B.������ĿID " & _
                      "AND B.�Ǽ�id=[1] " & _
                "ORDER BY B.������� "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���������Ŀ
            
            strSQL = "SELECT x.ID, " & _
                          "DECODE(x.���, 'C', '����', 'D', '���') AS ��Ŀ���, " & _
                          "x.���� As ��Ŀ����, " & _
                          "x.���㵥λ, " & _
                          "y.����걾, " & _
                          "y.��鲿λ, " & _
                          "t.���� As �ɼ���ʽ, " & _
                          "z.�����۸�,z.���۸�, " & _
                          "Decode(z.�����۸�,Null,0,0,0,10*z.���۸�/z.�����۸�) As �ۿ�," & _
                          "y.�ɼ���ʽid, " & _
                          "y.��鲿λid,'' As �Ʒ���ϸ " & _
                     "FROM ������ĿĿ¼ x, " & _
                          "�������Ŀ¼ y, " & _
                          "(Select a.������Ŀid,Sum(b.�ּ�*a.����) As �����۸�,Sum(b.�ּ�*a.����*Nvl(a.�ۿ�,1)) As ���۸� " & _
                           "From ������ͼƼ� a, " & _
                                "�շѼ�Ŀ b " & _
                           "Where a.��� = [1] " & _
                                 "and b.�շ�ϸĿid=a.�շ�ϸĿid " & _
                                 "and b.ִ������<=SYSDATE and (b.��ֹ���� IS NULL OR b.��ֹ����>SYSDATE) " & _
                           "Group by a.������Ŀid " & _
                          ") z, " & _
                          "������ĿĿ¼ t " & _
                    "Where x.ID = y.������ĿID And y.��� = [1] and x.id=z.������Ŀid(+) " & _
                          "and t.id(+)=y.�ɼ���ʽid"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������ͼƼ�
            
            strSQL = "Select c.����,c.���㵥λ,b.�ּ�,a.����,b.������Ŀid,b.�ּ�*a.���� As ���,c.id " & _
                        "from ������ͼƼ� a,�շѼ�Ŀ b,�շ���ĿĿ¼ c " & _
                        "Where a.�շ�ϸĿid = c.ID " & _
                              "and b.�շ�ϸĿid=a.�շ�ϸĿid " & _
                              "and b.ִ������<=SYSDATE and (b.��ֹ���� IS NULL OR b.��ֹ����>SYSDATE) and a.���=[1]"
       
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���˷��øſ�

            strSQL = "SELECT D.Ӧ�ս��,D.ʵ�ս��,D.���ʽ��,D.���ʷ��� " & _
                     "FROM ���˷��ü�¼ D, " & _
                          "(SELECT C.ID " & _
                             "FROM �����Ա���� A, ���ǼǼ�¼ B, ����ҽ����¼ C " & _
                            "WHERE A.�Ǽ�ID = B.ID AND A.����ID = C.����ID AND C.������Դ = 4 AND " & _
                                  "B.���� = C.�Һŵ� AND A.ID = [1]) E " & _
                    "WHERE D.ҽ����� = E.ID"
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.������øſ�

            strSQL = "SELECT D.Ӧ�ս��,D.ʵ�ս��,D.���ʽ��,D.���ʷ��� " & _
                     "FROM ���˷��ü�¼ D, " & _
                          "(SELECT C.ID " & _
                             "FROM ���ǼǼ�¼ B, ����ҽ����¼ C " & _
                            "WHERE C.������Դ = 4 AND " & _
                                  "B.���� = C.�Һŵ� AND B.ID = [1]) E " & _
                    "WHERE D.ҽ����� = E.ID"
                    
'            strSQL = "SELECT NVL(SUM(D.ʵ�ս��), 0) AS ʵ�ս��," & _
'                        "NVL(SUM(DECODE(D.���ʷ���,1,D.ʵ�ս��,0)),0) AS ���ʽ��," & _
'                        "NVL(SUM(DECODE(D.���ʷ���,1,0,D.ʵ�ս��)),0) AS �շѽ��, " & _
'                        "NVL(SUM(DECODE(D.���ʷ���,1,NVL(D.ʵ�ս��,0) - NVL(D.���ʽ��,0),0)),0) AS δ����, " & _
'                        "NVL(SUM(DECODE(D.���ʷ���,1,0,NVL(D.ʵ�ս��,0) -  NVL(D.���ʽ��,0))),0) AS δ�ս��, " & _
'                        "NVL(SUM(NVL(D.ʵ�ս��,0) - NVL(D.���ʽ��,0)), 0) AS δ����ϼ� " & _
'                     "FROM ���˷��ü�¼ D, " & _
'                          "(SELECT C.ID " & _
'                             "FROM �����Ա���� A, ���ǼǼ�¼ B, ����ҽ����¼ C " & _
'                            "WHERE A.�Ǽ�ID = B.ID AND A.����ID = C.����ID AND C.������Դ = 4 AND " & _
'                                  "C.ҽ��״̬ <> 4 AND B.���� = C.�Һŵ� AND A.ID = [1]) E " & _
'                    "WHERE D.��¼״̬ IN (0, 1) AND D.ҽ����� = E.ID"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�����Ŀ�۱�
            
            If varParam(2) = "" Then
                
                strSQL = "Select y.����,y.���㵥λ,z.�շ�����,x.�ּ�,y.id,Decode(x.������Ŀid,[2],2,1) As �Ƽ�����,y.��� " & _
                            "From ( " & _
                              "Select a.������Ŀid,a.�շ���Ŀid,Sum(c.�ּ�) As �ּ� " & _
                              "From �շѼ�Ŀ c, " & _
                                   "�����շѹ�ϵ a, " & _
                                   "������ĿĿ¼ b " & _
                              "Where a.�շ���Ŀid = c.�շ�ϸĿid " & _
                                    "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                    "AND b.ID=a.������Ŀid " & _
                                    "AND NVL(b.�Ƽ�����,0)=0 " & _
                                    "and a.������Ŀid In ([1],[2]) " & _
                              "Group by a.������Ŀid,a.�շ���Ŀid " & _
                            ") x, " & _
                            "�շ���ĿĿ¼ y, " & _
                            "�����շѹ�ϵ z " & _
                            "Where x.�շ���Ŀid = y.ID " & _
                                  "and z.�շ���Ŀid=x.�շ���Ŀid " & _
                                  "and z.������Ŀid=x.������Ŀid"
                                  
            Else
            
                strTmp = Val(varParam(0)) & "," & Val(varParam(1)) & "," & varParam(2)
                If Right(strTmp, 1) = "," Then strTmp = strTmp & "0"
            
                strSQL = "Select y.����,y.���㵥λ,z.�շ�����,x.�ּ�,y.id,Decode(x.������Ŀid," & Val(varParam(1)) & ",2,1) As �Ƽ�����,y.��� " & _
                            "From ( " & _
                              "Select a.������Ŀid,a.�շ���Ŀid,Sum(c.�ּ�) As �ּ� " & _
                              "From �շѼ�Ŀ c, " & _
                                   "�����շѹ�ϵ a, " & _
                                   "������ĿĿ¼ b " & _
                              "Where a.�շ���Ŀid = c.�շ�ϸĿid " & _
                                    "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                    "AND b.ID=a.������Ŀid " & _
                                    "AND NVL(b.�Ƽ�����,0)=0 " & _
                                    "and a.������Ŀid In (" & strTmp & ") " & _
                              "Group by a.������Ŀid,a.�շ���Ŀid " & _
                            ") x, " & _
                            "�շ���ĿĿ¼ y, " & _
                            "�����շѹ�ϵ z " & _
                            "Where x.�շ���Ŀid = y.ID " & _
                                  "and z.�շ���Ŀid=x.�շ���Ŀid " & _
                                  "and z.������Ŀid=x.������Ŀid"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���ԤԼ����

            
            strSQL = "SELECT A.ID," & _
                             "A.����," & _
                             "A.��ϵ�� AS ԤԼ��," & _
                             "A.��ϵ�绰," & _
                             "DECODE(A.�Ƿ�����,1,'','����') AS ����," & _
                             "DECODE(A.���״̬,1,'�¿�',2,'ȷ��',3,'ȡ��',4,'��ʼ',5,'���') AS ״̬," & _
                             "(SELECT COUNT(1) FROM �����Ա���� WHERE ����id>0 AND �Ǽ�id=A.ID) AS ����," & _
                             "A.���״̬,A.��Լ��λid," & _
                             "A.��ϵ��ַ, "
            
            strSQL = strSQL & _
                            "(SELECT NVL(SUM(�����۸�),0) FROM �����Ŀ�嵥 X,�����Ա���� T WHERE X.�Ǽ�ID = T.�Ǽ�ID AND X.�������=T.������� AND T.����id>0 AND T.�Ǽ�id=A.ID) AS Ӧ�ս��,"
                            
            strSQL = strSQL & _
                            "(SELECT NVL(SUM(���۸�),0) FROM �����Ŀ�嵥 X,�����Ա���� T WHERE X.�Ǽ�ID = T.�Ǽ�ID AND X.�������=T.������� AND T.����id>0 AND T.�Ǽ�id=A.ID) AS ���۸�,"
            
            strSQL = strSQL & _
                             "A.�����ۿ�," & _
                             "A.�������," & _
                             "TO_CHAR(A.���ʱ��,'yyyy-MM-dd') AS ԤԼʱ��," & _
                             "to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd HH:mm') AS �Ǽ�ʱ��," & _
                             "B.���� AS ����," & _
                             "A.����˵�� " & _
                        "FROM ���ǼǼ�¼ A,��Լ��λ B " & _
                        "WHERE A.��Լ��λID=B.ID(+) AND A.��첿��id=[1] " & strParam
            
            strSQL = "SELECT ID,���� AS No,ԤԼ��,����,����˵��,��ϵ�绰,����,״̬,����,���״̬,��Լ��λid,��ϵ��ַ,DECODE(�����ۿ�,1,NULL,10*�����ۿ�) AS �ۿ�,Ӧ�ս��,���۸� AS ʵ�ս��,�������,ԤԼʱ��,�Ǽ�ʱ�� FROM (" & strSQL & ") ORDER BY ���� DESC"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�������ͳ��
            
            If Val(varParam(0)) = 2 Then
                strSQL = "Select Count(1) From �����Ա���� A,���ǼǼ�¼ B Where A.�Ǽ�id=B.ID AND A.����id>0 And A.��챨��=[4] AND B.���״̬=[1] AND B.���ʱ�� BETWEEN [2] AND [3]"
            Else
                strSQL = "Select Count(1) From �����Ա���� A,���ǼǼ�¼ B Where A.�Ǽ�id=B.ID And a.����id>0 And a.��챨��=[4] AND a.���״̬=[1] AND b.���ʱ�� BETWEEN [2] AND [3]"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��������Ա1
        
            strSQL = "SELECT B.ID," & _
                            "1 AS ����," & _
                            "B.������� AS ����2," & _
                            "C.��Լ��λid AS �ϼ�id," & _
                            "A.����id,A.�����,a.������,a.���￨��,b.�����, " & _
                            "C.���� AS ��쵥��," & _
                            "B.����," & _
                            "A.����," & _
                            "A.�Ա�," & _
                            "A.����," & _
                            "A.����״��,B.�Ǽ�id," & _
                            "Decode(C.��Լ��λid,NULL,98,99) AS ��־," & _
                            "C.���� AS ���ݺ�," & _
                            "'����' AS ����," & _
                            "DECODE(B.���״̬,1,'ȷ��',4,'��ʼ',5,'���') AS ״̬ " & _
                        "FROM ������Ϣ A,�����Ա���� B,���ǼǼ�¼ C " & _
                        "WHERE B.��챨��=[4]  AND C.���״̬=[3] AND A.����ID=B.����ID AND C.ID=B.�Ǽ�id " & _
                            "AND B.�������=[2]  "
            If Val(varParam(0)) > 0 Then
                strSQL = strSQL & " AND C.ID=[1] "
            Else
                strSQL = strSQL & "AND Nvl(C.�Ƿ�����,0)=0 AND C.���ʱ�� BETWEEN [5] AND [6]"
            End If
            
            If blnMoveOuted Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "�����Ա����", "H�����Ա����")
                strTmp = Replace(strTmp, "���ǼǼ�¼", "H���ǼǼ�¼")
                strSQL = "Select * From (" & strSQL & " Union All " & strTmp & ") a "
            End If
            
            strSQL = strSQL & " Order By a.�����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.��������Ա
            
            strSQL = _
                "Select A.ID, 1 As ����, A.������� As ����2, A.��Լ��λid As �ϼ�id, A.���� As ��쵥��, a.�����,A.����, A.����id, B.�����," & vbNewLine & _
                "       B.������, B.���￨��, B.����, B.�Ա�, B.����, B.����״��, A.�Ǽ�id, Decode(A.��Լ��λid, Null, 98, 99) As ��־," & vbNewLine & _
                "       A.���� As ���ݺ�, Decode(A.����, A.����, '����', '����') As ����," & vbNewLine & _
                "       Decode(A.���״̬, 1, 'ȷ��', 4, '��ʼ', 5, '���') As ״̬" & vbNewLine & _
                "From (Select A.�Ǽ�id, A.����id, A.����,a.�����, A.ID, A.�������, A.��Լ��λid, A.����, A.���״̬," & vbNewLine & _
                "              Sum(Decode(B.�����ļ�id, Null, 0, 1)) As ����, Sum(Decode(A.����id, Null, 0, 1)) As ����" & vbNewLine & _
                "       From (Select A.�Ǽ�id, A.����id, A.����,a.�����, A.ID, A.�������, A.��Լ��λid, A.����, A.���״̬, D.�������," & vbNewLine & _
                "                     D.������Ŀid, E.����id, D.���id" & vbNewLine & _
                "              From (Select A.�Ǽ�id, A.����id, B.����,a.�����, A.ID, A.�������, B.��Լ��λid, A.����, A.���״̬" & vbNewLine & _
                "                     From �����Ա���� A, ���ǼǼ�¼ B" & vbNewLine & _
                "                     Where A.�Ǽ�id = B.ID And A.��챨�� = [4] And B.���״̬ = [3] And A.������� = [2] "
        
            If Val(varParam(1)) = 1 Then
                '��������Ա��
                                        
                If Val(varParam(0)) > 0 Then
                    strSQL = strSQL & _
                            "                          And b.ID=[1] And a.���ʱ�� Between [5] And [6] "
                Else
                
                    strSQL = strSQL & _
                        "                          And Nvl(B.�Ƿ�����, 0) = 0 And a.���ʱ�� Between [5] And [6] "
                        
                End If
                
            Else
                If Val(varParam(0)) > 0 Then
    
                    strSQL = strSQL & _
                        "                          And b.ID=[1] "
                Else
                
                    strSQL = strSQL & _
                        "                          And Nvl(B.�Ƿ�����, 0) = 0 And B.���ʱ�� Between [5] And [6] "
                End If
            End If
            
        
                
            strSQL = strSQL & _
                ") A, ����ҽ����¼ D," & vbNewLine & _
                "                   ����ҽ������ E" & vbNewLine & _
                "              Where D.�Һŵ�(+) = A.���� And D.����id(+) = A.����id And D.�������(+) <> 'E' And D.������Դ(+) = 4 And" & vbNewLine & _
                "                    D.ҽ��״̬(+) <> 4 And D.ID = E.ҽ��id(+)) A, ���Ƶ���Ӧ�� B" & vbNewLine & _
                "       Where ((A.������� = 'D' And A.���id Is Null) Or A.������� = 'C' Or A.������� Is Null) And" & vbNewLine & _
                "             A.������Ŀid = B.������Ŀid(+) And B.Ӧ�ó���(+) = 4" & vbNewLine & _
                "       Group By A.�Ǽ�id, A.����id, A.����,a.�����, A.ID, A.�������, A.��Լ��λid, A.����, A.���״̬) A, ������Ϣ B" & vbNewLine & _
                "Where A.����id = B.����id "

            If blnMoveOuted Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "�����Ա����", "H�����Ա����")
                strTmp = Replace(strTmp, "���ǼǼ�¼", "H���ǼǼ�¼")
                strTmp = Replace(strTmp, "����ҽ����¼", "H����ҽ����¼")
                strTmp = Replace(strTmp, "����ҽ������", "H����ҽ������")
                strSQL = "Select * From (" & strSQL & " Union All " & strTmp & ") b "
            End If
            strSQL = strSQL & " Order By b.�����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.���Ǽǵ���
            '0-���˷�����;1-����������;2-���������;98-�������ܼ���Ա��;99-�����ܼ���Ա
            '������쵥ͷ
            strSQL = "SELECT -1 AS ID," & _
                            "-1 AS ����1," & _
                            "'' AS ����2," & _
                            "0 AS �ϼ�id," & _
                            "'' AS ״̬," & _
                            "'' AS ����," & _
                            "NULL+0 AS ����id," & _
                            "'' AS ��쵥��," & _
                            "0 AS �����," & _
                            "'' AS ������,'' As ���￨��,'' As �����," & _
                            "'<����>' AS ����," & _
                            "'' AS �Ա�," & _
                            "'' AS ����," & _
                            "'' AS ����״��," & _
                            "NULL+0 AS �Ǽ�id," & _
                            "0 AS ��־,Null+0 As ����," & _
                            "'' AS ���ݺ� " & _
                        "FROM DUAL "

            '��������
            strSQL = strSQL & " UNION ALL " & _
                        "SELECT DISTINCT A.ID," & _
                                    "1 AS ����1," & _
                                    "'' AS ����2," & _
                                    "0 AS �ϼ�id," & _
                                    "'' AS ״̬," & _
                                    "'' AS ����," & _
                                    "0 AS ����id," & _
                                    "B.���� AS ��쵥��," & _
                                    "NULL+0 AS �����," & _
                                    "'' AS ������,'' As ���￨��,'' As �����," & _
                                    "A.����||'('||B.����||')' AS ����," & _
                                    "'' AS �Ա�," & _
                                    "'' AS ����," & _
                                    "'' AS ����״��," & _
                                    "B.ID AS �Ǽ�id, " & _
                                    "1 AS ��־,Null+0 As ����," & _
                                    "B.���� AS ���ݺ� " & _
                    "FROM   ��Լ��λ A," & _
                            "���ǼǼ�¼ B " & _
                    "WHERE  A.ID=B.��Լ��λid " & _
                            "AND B.���״̬=[4] " & _
                            "AND B.��첿��id+0=[1] "
                            
            If Val(varParam(0)) = 1 Then
                '������ʱ���������쵥��
                strSQL = strSQL & "AND B.ID In (Select �Ǽ�id From �����Ա���� Where ��챨��=1 And ���ʱ�� BETWEEN [2] AND [3]) "
            Else
                strSQL = strSQL & "AND B.���ʱ�� BETWEEN [2] AND [3] "
            End If
            
            If blnMoveOuted Then
            
                strSQL = strSQL & " UNION ALL " & _
                            "SELECT DISTINCT A.ID," & _
                                        "1 AS ����1," & _
                                        "'' AS ����2," & _
                                        "0 AS �ϼ�id," & _
                                        "'' AS ״̬," & _
                                        "'' AS ����," & _
                                        "0 AS ����id," & _
                                        "B.���� AS ��쵥��," & _
                                        "NULL+0 AS �����," & _
                                        "'' AS ������,'' As ���￨��,'' As �����," & _
                                        "A.����||'('||B.����||')' AS ����," & _
                                        "'' AS �Ա�," & _
                                        "'' AS ����," & _
                                        "'' AS ����״��," & _
                                        "B.ID AS �Ǽ�id, " & _
                                        "1 AS ��־,Null+0 As ����," & _
                                        "B.���� AS ���ݺ� " & _
                        "FROM   ��Լ��λ A," & _
                                "H���ǼǼ�¼ B " & _
                        "WHERE  A.ID=B.��Լ��λid " & _
                                "AND B.���״̬=[4] " & _
                                "AND B.��첿��id+0=[1] "
                                
                If Val(varParam(0)) = 1 Then
                    '������ʱ���������쵥��
                    strSQL = strSQL & "AND B.ID In (Select �Ǽ�id From �����Ա���� Where ��챨��=1 And ���ʱ�� BETWEEN [2] AND [3]) "
                Else
                    strSQL = strSQL & "AND B.���ʱ�� BETWEEN [2] AND [3] "
                End If
            
            End If
            
            '�������
            strSQL = strSQL & " UNION ALL " & _
                        "SELECT DISTINCT A.ID," & _
                                    "1 AS ����1," & _
                                    "C.������� AS ����2," & _
                                    "A.ID AS �ϼ�id," & _
                                    "'' AS ״̬," & _
                                    "'' AS ����," & _
                                    "NULL+0 AS ����id," & _
                                    "B.���� AS ��쵥��," & _
                                    "0 AS �����,'' As ���￨��,'' As �����," & _
                                    "'' AS ������," & _
                                    "C.������� AS ����," & _
                                    "'' AS �Ա�," & _
                                    "'' AS ����," & _
                                    "'' AS ����״��," & _
                                    "B.ID AS �Ǽ�id, " & _
                                    "2 AS ��־,Null+0 as ����," & _
                                    "B.���� AS ���ݺ� " & _
                    "FROM   ��Լ��λ A," & _
                            "���ǼǼ�¼ B,������ C " & _
                    "WHERE  C.�Ǽ�id=B.ID AND A.ID=B.��Լ��λid " & _
                            "AND B.���״̬=[4] " & _
                            "AND B.��첿��id+0=[1] "
                            
            If Val(varParam(0)) = 1 Then
                '������ʱ���������쵥��
                strSQL = strSQL & "AND B.ID In (Select �Ǽ�id From �����Ա���� Where ��챨��=1 And ���ʱ�� BETWEEN [2] AND [3]) "
            Else
                strSQL = strSQL & "AND B.���ʱ�� BETWEEN [2] AND [3] "
            End If
                
            If blnMoveOuted Then
                strSQL = strSQL & " UNION ALL " & _
                            "SELECT DISTINCT A.ID," & _
                                        "1 AS ����1," & _
                                        "C.������� AS ����2," & _
                                        "A.ID AS �ϼ�id," & _
                                        "'' AS ״̬," & _
                                        "'' AS ����," & _
                                        "NULL+0 AS ����id," & _
                                        "B.���� AS ��쵥��," & _
                                        "0 AS �����,'' As ���￨��,'' As �����," & _
                                        "'' AS ������," & _
                                        "C.������� AS ����," & _
                                        "'' AS �Ա�," & _
                                        "'' AS ����," & _
                                        "'' AS ����״��," & _
                                        "B.ID AS �Ǽ�id, " & _
                                        "2 AS ��־,Null+0 as ����," & _
                                        "B.���� AS ���ݺ� " & _
                        "FROM   ��Լ��λ A," & _
                                "H���ǼǼ�¼ B,H������ C " & _
                        "WHERE  C.�Ǽ�id=B.ID AND A.ID=B.��Լ��λid " & _
                                "AND B.���״̬=[4] " & _
                                "AND B.��첿��id+0=[1] "
                                
                If Val(varParam(0)) = 1 Then
                    '������ʱ���������쵥��
                    strSQL = strSQL & "AND B.ID In (Select �Ǽ�id From �����Ա���� Where ��챨��=1 And ���ʱ�� BETWEEN [2] AND [3]) "
                Else
                    strSQL = strSQL & "AND B.���ʱ�� BETWEEN [2] AND [3] "
                End If
            
            End If
            
            strSQL = "SELECT * FROM (" & strSQL & ") A ORDER BY ����1,��쵥��,�ϼ�id,����2,�����"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.����ִ�п���
                                                                                        
            '����:������Ŀid'���˿���id'��������id'��������
            
            strSQL = _
                "SELECT A.ID FROM ���ű� A,������ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,��λ״����¼ B,������ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=2 AND A.ID=B.����id AND B.����ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,������ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,����ִ�п��� B,������ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=4 AND A.ID=B.ִ�п���id AND B.������Դ=1 AND B.������Ŀid=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,����ִ�п��� B,������ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=4 AND " & _
                            "A.ID=B.ִ�п���id AND B.������Դ IS NULL AND (B.��������id IS NULL OR B.��������id=[3]) AND B.������Ŀid=X.ID "
                            
            If Val(varParam(0)) = 0 Then
            
                strSQL = _
                    "SELECT 1 As ĩ��,A.����,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4])"
                    
            Else
                strSQL = _
                    "SELECT 1 As ĩ��,A.����,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4]) Union All " & _
                    "SELECT Distinct 1 As ĩ��,A.����,A.����,A.����,A.ID FROM ���ű� A,��������˵�� B WHERE A.ID=B.����ID And B.������� In (1,3) And A.ID Not IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4])"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.�շ�ִ�п���
        
            '����:������Ŀid'���˿���id'��������id'��������
            
            strSQL = _
                "SELECT A.ID FROM ���ű� A,�շ���ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,��λ״����¼ B,�շ���ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=2 AND A.ID=B.����id AND B.����ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,�շ���ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,�շ�ִ�п��� B,�շ���ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=4 AND A.ID=B.ִ�п���id AND B.������Դ=1 AND B.�շ�ϸĿid=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM ���ű� A,�շ�ִ�п��� B,�շ���ĿĿ¼ X WHERE X.ID=[1] AND X.ִ�п���=4 AND " & _
                            "A.ID=B.ִ�п���id AND B.������Դ IS NULL AND (B.��������id IS NULL OR B.��������id=[3]) AND B.�շ�ϸĿid=X.ID "
            
            If Val(varParam(0)) = 0 Then
                strSQL = _
                    "SELECT 1 As ĩ��,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4])"
                    
            Else
                strSQL = _
                    "SELECT 1 As ĩ��,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4]) Union All " & _
                    "SELECT Distinct 1 As ĩ��,A.����,A.����,A.ID FROM ���ű� A,��������˵�� B WHERE A.ID=B.����ID And B.������� In (1,3) And A.ID Not IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4]) "
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.ҩƷִ�п���
            
            strSQL = "SELECT Distinct 1 As ĩ��,A.����,A.����,A.ID " & _
                    "from ���ű� A,��������˵�� B " & _
                    "where (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                    "and A.ID=B.����ID and B.������� in (1,3) " & _
                    "and B.��������=Decode([1],'5','��ҩ��','6','��ҩ��','7','��ҩ��','4','���ϲ���')"
                              
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




