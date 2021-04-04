Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL
    
    ��Ա��������
    �ֿ���Ŀ���
    �ֿ���Ŀ����
    �ֿ���Ŀ���
    �ܼ챨�潨��
    ������Ͻ��
    
    ҩƷִ�п���
    ����ִ�п���
    �շ�ִ�п���
    �����Ŀ�۱�
End Enum

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '����:  ���в���SQL���
    '����:  strMenu             Ҫ������SQL����
    '       strParam            ������,��ʽ:"����ֵ1'����ֵ2"
    '����:  SQL���
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim varParam As Variant
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        Case SQL.��Ա��������
            
'            strSQL = "Select * From �����Ա����_�ɱ� A,���ǼǼ�¼_�ɱ� B,������Ϣ_�ɱ� C,������Ϣ D,�����Ա���� E  " & _
'                        "WHERE D.����id=A.����id And C.����id=A.����id And A.�������=B.������� And E.�Ǽ�ID=B.�Ǽ�ID And E.����ID=A.����ID " & _
'                                "AND B.�������='" & varParam(0) & "' And A.����id=" & Val(varParam(1))
                                
            strSQL = "Select * From �����Ա����_�ɱ� A,���ǼǼ�¼_�ɱ� B,������Ϣ D,�����Ա���� E,������_�ɱ� c  " & _
                        "WHERE D.����id=A.����id And A.�������=B.������� And E.�Ǽ�ID=B.�Ǽ�ID And E.����ID=A.����ID " & _
                                "AND c.�Ǽ�id=b.�Ǽ�id and c.�������=e.������� and B.�������=[1] And A.����id=[2]"
                                
        Case SQL.�ֿ���Ŀ���
            
            strSQL = _
            "SELECT ������id,�����Ŀid,ִ�в���id,���,��־,��λ,�ο�,��ϱ���,��Ͽ��� " & _
            "FROM ( " & _
              "SELECT " & _
                     "R.ִ�в���ID, " & _
                     "R.���, " & _
                     "R.��λ,R.��ϱ���,R.��Ͽ���, " & _
                     "DECODE(SIGN(INSTR(R.��־�ο�,'''')),1,SUBSTR(R.��־�ο�,1,INSTR(R.��־�ο�,'''')-1),'') AS ��־, " & _
                     "DECODE(SIGN(INSTR(R.��־�ο�,'''')),1,SUBSTR(R.��־�ο�,INSTR(R.��־�ο�,'''')+1,1000),'') AS �ο�, " & _
                     "�����Ŀid, " & _
                     "������id " & _
              "FROM ( " & _
                    "Select " & _
                           "A.ִ�в���ID, " & _
                           "A.�����Ŀid,A.��ϱ���,A.��Ͽ���," & _
                           "A.ID, " & _
                           "X.�������, " & _
                           "Y.�����, " & _
                           "DECODE(SIGN(INSTR(Y.���,'''')),1,SUBSTR(Y.���,1,INSTR(Y.���,'''')-1),Y.���) AS ���, " & _
                           "Y.��λ, " & _
                           "Y.������id, " & _
                           "DECODE(SIGN(INSTR(Y.���,'''')),1,SUBSTR(Y.���,INSTR(Y.���,'''')+1,1000),'') AS ��־�ο� " & _
                    "From "
                    
            strSQL = strSQL & _
                         "( " & _
                         "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A6.�ɱ����� As ��ϱ���,A6.��Ͽ��� " & _
                         "from �����Ŀҽ�� A1, " & _
                               "����ҽ����¼ A2, " & _
                               "����ҽ������ A3, " & _
                               "���˲�����¼ A4, " & _
                               "�����Ŀ�嵥 A5,������ĿĿ¼_�ɱ� A6 " & _
                         "Where A1.����id =[2] " & _
                                " AND A5.�Ǽ�id=[1] " & _
                                " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                                "AND A3.ҽ��ID=A2.ID " & _
                                "AND A4.ID=A3.����ID AND A6.������Ŀid=A2.������Ŀid " & _
                                "AND A5.ID=A1.�嵥ID  AND A2.������� In ('C','D') " & _
                         ") A, " & _
                         "���˲������� X, " & _
                         "( " & _
                         "select " & _
                                 "A.����ID, " & _
                                 "A.�ؼ��� AS �����, " & _
                                 "A.�������� AS ���, " & _
                                 "B.��λ, " & _
                                 "A.������id " & _
                          "From "
                          
            strSQL = strSQL & _
                            "���˲��������� A, " & _
                            "����������Ŀ B " & _
                          "Where A.������id = B.ID And ������id > 0 " & _
                          ") Y " & _
                    "Where x.������¼id = A.ID And X.ID = Y.����ID " & _
                    ") R " & _
                ") A"
            
            strSQL = "Select W.*,T.�ɱ����� As ��Ŀ����,T.��Ŀ��֧,T.��Ŀ����,T.�ɱ����� As ��Ŀ���� From (" & strSQL & ") W,����������Ŀ_�ɱ� T,�����Ա����_�ɱ� K " & _
                        "WHERE T.������Ŀid=W.������id " & _
                                "AND K.����id=[2] And K.�������=[3]"
        
        Case SQL.�ֿ���Ŀ����
                
            strSQL = _
                "Select " & _
                       "Distinct y.��������,0 As ����id,Y.��Ͻ���, A.ִ�в���ID, A.�����Ŀid,A.��д��,A.�������� " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A4.��д��,A4.�������� " & _
                     "from �����Ŀҽ�� A1, " & _
                           "����ҽ����¼ A2, " & _
                           "����ҽ������ A3, " & _
                           "���˲�����¼ A4, " & _
                           "�����Ŀ�嵥 A5 " & _
                     "Where A1.����id =[2] " & _
                            " AND A5.�Ǽ�id=[1] " & _
                            " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                            "AND A3.ҽ��ID=A2.ID " & _
                            "AND A4.ID=A3.����ID " & _
                            "AND A5.ID=A1.�嵥ID AND A2.������� In ('C','D') " & _
                     ") A, " & _
                     "���˲������� X, " & _
                     "�����Ա���� Y " & _
                "Where x.������¼id = A.ID And x.ID = y.����ID And y.�������� Is Not Null "
            
            strSQL = "Select Distinct W.*,T.�ɱ����� As ��ϱ���,T.��Ͽ���,T.�ɱ����� As ������� From (" & strSQL & ") W,������ĿĿ¼_�ɱ� T,�����Ա����_�ɱ� K " & _
                        "WHERE T.������Ŀid=W.�����Ŀid " & _
                                "AND K.����id=[2] And K.�������=[3] Order By T.��Ͽ���,T.�ɱ�����"
                                
        Case SQL.�ֿ���Ŀ���
            
            strSQL = _
                "Select " & _
                       "Distinct 0 As ����id,Y.��Ͻ���,Y.����id,A.ִ�в���ID,A.�����Ŀid,A.��д��,A.�������� " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A4.��д��,A4.�������� " & _
                     "from �����Ŀҽ�� A1, " & _
                           "����ҽ����¼ A2, " & _
                           "����ҽ������ A3, " & _
                           "���˲�����¼ A4, " & _
                           "�����Ŀ�嵥 A5 " & _
                     "Where A1.����id =[2] " & _
                            " AND A5.�Ǽ�id=[1] " & _
                            " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                            "AND A3.ҽ��ID=A2.ID " & _
                            "AND A4.ID=A3.����ID " & _
                            "AND A5.ID=A1.�嵥ID  AND A2.������� In ('C','D') " & _
                     ") A, " & _
                     "���˲������� X, " & _
                     "�����Ա���� Y " & _
                "Where x.������¼id = A.ID And x.ID = y.����ID And Y.����id Is Not Null"
            
            strSQL = "Select Distinct W.*,T.�ɱ����� As ��ϱ���,T.��Ͽ���,T.�ɱ����� As �������,X.�ɱ����� As ��ϱ���,X.�ɱ����� As �������,X.�������� From (" & strSQL & ") W,������ĿĿ¼_�ɱ� T,�����Ա����_�ɱ� K,�����Ͻ���_�ɱ� X " & _
                        "WHERE T.������Ŀid=W.�����Ŀid " & _
                                "AND X.����id=W.����id AND X.�������� Is Not Null " & _
                                "AND K.����id=[2] And K.�������=[3] Order By T.��Ͽ���,T.�ɱ�����"
        Case SQL.������Ͻ��
            
            strSQL = _
                "SELECT  Distinct 0 As ����id,X1.��Ͻ���,X1.����id,Y.����id,Y.��д��,Y.�������� " & _
                "FROM �����Ա���� A, " & _
                     "���˲������� X, " & _
                     "���˲�����¼ Y, " & _
                      "�����Ա���� X1,����Ԫ��Ŀ¼ X2 " & _
                "Where X.������¼id = A.��첡��ID " & _
                      "AND X.ID=X1.����id " & _
                      "AND Y.ID=X.������¼id " & _
                      "AND A.����ID=[2] " & _
                      " AND A.�Ǽ�ID=[1] " & _
                      " AND X.Ԫ������=4 and X.Ԫ�ر���=X2.���� AND Upper(X2.����)='ZL9CISCORE.USRMEDICALSUM'"
            
            strSQL = "Select Distinct W.*,X.�ɱ����� As ��ϱ���,X.�ɱ����� As �������,X.�������� From (" & strSQL & ") W,�����Ա����_�ɱ� K,�����Ͻ���_�ɱ� X " & _
                        "WHERE X.����id=W.����id AND X.�������� Is Not Null " & _
                                "AND K.����id=[2] And K.�������=[3] "
                                
        Case SQL.�ܼ챨�潨��
        
            strSQL = _
                "SELECT  DECODE(SIGN(INSTR(���,'�������飺')),1,SUBSTR(���,8,INSTR(���,'�������飺')-11),���) AS ����ͷ," & _
                        "DECODE(SIGN(INSTR(���,'�������飺')),1,SUBSTR(���,INSTR(���,'�������飺')+7,4000),���) AS ����ָ��," & _
                        "��д��, " & _
                        "TO_CHAR(��д����,'yyyy-mm-dd') AS ��д���� " & _
                "FROM ( " & _
                "select " & _
                       "X.�������, " & _
                       "X1.�����, " & _
                       "X1.���, " & _
                       "Y.��д��, " & _
                       "y.��д���� " & _
                "From " & _
                     "�����Ա���� A, " & _
                     "���˲������� X, " & _
                     "���˲�����¼ Y,����Ԫ��Ŀ¼ X2, " & _
                      "(select ����id,0 AS �����,'' AS ��Ŀ,���� AS ��� from ���˲����ı��� ) X1 " & _
                "Where x.������¼id = A.��첡��ID " & _
                      "AND X.ID=X1.����id " & _
                      "AND Y.ID=X.������¼id " & _
                      "AND A.����ID=[2] " & _
                      " AND A.�Ǽ�ID=[1] " & _
                      " AND X.Ԫ������=4 and X.Ԫ�ر���=X2.���� AND upper(X2.����)='ZL9CISCORE.USRMEDICALSUM' " & _
                ") ORDER BY  �������,�����"

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
            
            strSQL = _
                "SELECT 1 As ĩ��,A.����,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4])"
        
    Case SQL.ҩƷִ�п���
            
            strSQL = "SELECT Distinct 1 As ĩ��,A.����,A.����,A.ID " & _
                    "from ���ű� A,��������˵�� B " & _
                    "where (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                    "and A.ID=B.����ID and B.������� in (2,3) " & _
                    "and B.��������=Decode([1],'5','��ҩ��','6','��ҩ��','7','��ҩ��','4','���ϲ���')"
                    
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
            
            strSQL = _
                "SELECT 1 As ĩ��,A.����,A.����,A.ID FROM ���ű� A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.����) Like [4] OR UPPER(A.����) Like [4] OR A.���� Like [4])"
                                    
                                    
    Case SQL.�����Ŀ�۱�
            
            strTmp = Val(varParam(0)) & "," & varParam(2)
            If Right(strTmp, 1) = "," Then strTmp = strTmp & "0"
            
            strSQL = "Select y.����,y.���㵥λ,z.�շ�����,x.�ּ�,y.id,1 As �Ƽ�����,y.��� " & _
                        "From ( " & _
                          "Select a.������Ŀid,a.�շ���Ŀid,Sum(c.�ּ�) As �ּ� " & _
                          "From �շѼ�Ŀ c, " & _
                               "�����շѹ�ϵ a, " & _
                               "������ĿĿ¼ b " & _
                          "Where a.�շ���Ŀid = c.�շ�ϸĿid " & _
                                "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                "AND b.ID=a.������Ŀid " & _
                                "AND NVL(b.�Ƽ�����,0)=0 " & _
                                "and a.������Ŀid IN (" & strTmp & ") " & _
                          "Group by a.������Ŀid,a.�շ���Ŀid " & _
                        ") x, " & _
                        "�շ���ĿĿ¼ y, " & _
                        "�����շѹ�ϵ z " & _
                        "Where x.�շ���Ŀid = y.ID " & _
                              "and z.�շ���Ŀid=x.�շ���Ŀid " & _
                              "and z.������Ŀid=x.������Ŀid"
                                          
            strSQL = strSQL & " Union All Select y.����,y.���㵥λ,z.�շ�����,x.�ּ�,y.id,2 As �Ƽ�����,y.��� " & _
                        "From ( " & _
                          "Select a.������Ŀid,a.�շ���Ŀid,Sum(c.�ּ�) As �ּ� " & _
                          "From �շѼ�Ŀ c, " & _
                               "�����շѹ�ϵ a, " & _
                               "������ĿĿ¼ b " & _
                          "Where a.�շ���Ŀid = c.�շ�ϸĿid " & _
                                "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                "AND b.ID=a.������Ŀid " & _
                                "AND NVL(b.�Ƽ�����,0)=0 " & _
                                "and a.������Ŀid=" & Val(varParam(1)) & " " & _
                          "Group by a.������Ŀid,a.�շ���Ŀid " & _
                        ") x, " & _
                        "�շ���ĿĿ¼ y, " & _
                        "�����շѹ�ϵ z " & _
                        "Where x.�շ���Ŀid = y.ID " & _
                              "and z.�շ���Ŀid=x.�շ���Ŀid " & _
                              "and z.������Ŀid=x.������Ŀid"
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




