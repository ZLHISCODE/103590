Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL
    
    ��Ա��������
    �ֿ���Ŀ���
    �ֿ���Ŀ����
    �ֿ���Ŀ���
    �ܼ챨�潨��
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
            
            strSQL = "Select * From �����Ա����_�ɱ� A,���ǼǼ�¼_�ɱ� B,������Ϣ_�ɱ� C,������Ϣ D,�����Ա���� E  " & _
                        "WHERE D.����id=A.����id And C.����id=A.����id And A.�������=B.������� And E.�Ǽ�ID=B.�Ǽ�ID And E.����ID=A.����ID " & _
                                "AND B.�������='" & varParam(0) & "' And A.����id=" & Val(varParam(1))
                                
        Case SQL.�ֿ���Ŀ���
            
            strSQL = _
            "SELECT ������id,�����Ŀid,ִ�в���id,���,��־,��λ,�ο� " & _
            "FROM ( " & _
              "SELECT " & _
                     "R.ִ�в���ID, " & _
                     "R.���, " & _
                     "R.��λ, " & _
                     "DECODE(SIGN(INSTR(R.��־�ο�,'''')),1,SUBSTR(R.��־�ο�,1,INSTR(R.��־�ο�,'''')-1),'') AS ��־, " & _
                     "DECODE(SIGN(INSTR(R.��־�ο�,'''')),1,SUBSTR(R.��־�ο�,INSTR(R.��־�ο�,'''')+1,1000),'') AS �ο�, " & _
                     "�����Ŀid, " & _
                     "������id " & _
              "FROM ( " & _
                    "Select " & _
                           "A.ִ�в���ID, " & _
                           "A.�����Ŀid, " & _
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
                         "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid " & _
                         "from �����Ŀҽ�� A1, " & _
                               "����ҽ����¼ A2, " & _
                               "����ҽ������ A3, " & _
                               "���˲�����¼ A4, " & _
                               "�����Ŀ�嵥 A5 " & _
                         "Where A1.����id = " & Val(varParam(1)) & _
                                " AND A5.�Ǽ�id=" & Val(varParam(0)) & _
                                " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                                "AND A3.ҽ��ID=A2.ID " & _
                                "AND A4.ID=A3.����ID " & _
                                "AND A5.ID=A1.�嵥ID " & _
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
            
            strSQL = "Select W.*,P.��ϱ���,P.��Ŀ����,P.��Ŀ��֧,P.��Ŀ����,P.��Ͽ���,T.�ɱ����� As ��Ŀ���� From (" & strSQL & ") W,����������Ŀ_�ɱ� T,�����Ŀ�嵥_�ɱ� P,�����Ա����_�ɱ� K " & _
                        "WHERE T.������Ŀid=W.������id " & _
                                "And T.�ɱ�����=P.��Ŀ���� " & _
                                "AND K.�������=P.������� " & _
                                "AND K.�ײͱ���=P.�ײͱ��� " & _
                                "AND K.�ײ����=P.�ײ���� " & _
                                "AND K.����id=" & Val(varParam(1)) & " And K.�������='" & varParam(2) & "'"
        
        Case SQL.�ֿ���Ŀ����
                
            strSQL = _
                "Select " & _
                       "Distinct y.��������,Y.����id,Y.��Ͻ���, A.ִ�в���ID, A.�����Ŀid,A.��д��,A.�������� " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A4.��д��,A4.�������� " & _
                     "from �����Ŀҽ�� A1, " & _
                           "����ҽ����¼ A2, " & _
                           "����ҽ������ A3, " & _
                           "���˲�����¼ A4, " & _
                           "�����Ŀ�嵥 A5 " & _
                     "Where A1.����id = " & Val(varParam(1)) & _
                            " AND A5.�Ǽ�id=" & Val(varParam(0)) & _
                            " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                            "AND A3.ҽ��ID=A2.ID " & _
                            "AND A4.ID=A3.����ID " & _
                            "AND A5.ID=A1.�嵥ID " & _
                     ") A, " & _
                     "���˲������� X, " & _
                     "�����Ա���� Y " & _
                "Where x.������¼id = A.ID And x.ID = y.����ID And y.�������� Is Not Null "
            
            strSQL = "Select Distinct W.*,P.��ϱ���,P.��Ͽ���,T.�ɱ����� As ������� From (" & strSQL & ") W,������ĿĿ¼_�ɱ� T,�����Ŀ�嵥_�ɱ� P,�����Ա����_�ɱ� K " & _
                        "WHERE T.������Ŀid=W.�����Ŀid " & _
                                "And T.�ɱ�����=P.��ϱ��� " & _
                                "AND K.�������=P.������� " & _
                                "AND K.�ײͱ���=P.�ײͱ��� " & _
                                "AND K.�ײ����=P.�ײ���� " & _
                                "AND K.����id=" & Val(varParam(1)) & " And K.�������='" & varParam(2) & "' Order By P.��Ͽ���,P.��ϱ���"
                                
        Case SQL.�ֿ���Ŀ���
            
            strSQL = _
                "Select " & _
                       "Distinct Y.����id,Y.��Ͻ���,Y.����id,A.ִ�в���ID,A.�����Ŀid,A.��д��,A.�������� " & _
                "From " & _
                     "( " & _
                     "Select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A4.��д��,A4.�������� " & _
                     "from �����Ŀҽ�� A1, " & _
                           "����ҽ����¼ A2, " & _
                           "����ҽ������ A3, " & _
                           "���˲�����¼ A4, " & _
                           "�����Ŀ�嵥 A5 " & _
                     "Where A1.����id = " & Val(varParam(1)) & _
                            " AND A5.�Ǽ�id=" & Val(varParam(0)) & _
                            " AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id) " & _
                            "AND A3.ҽ��ID=A2.ID " & _
                            "AND A4.ID=A3.����ID " & _
                            "AND A5.ID=A1.�嵥ID " & _
                     ") A, " & _
                     "���˲������� X, " & _
                     "�����Ա���� Y " & _
                "Where x.������¼id = A.ID And x.ID = y.����ID And Y.����id Is Not Null And Y.����id Is Not Null"
            
            strSQL = "Select Distinct W.*,P.��ϱ���,P.��Ŀ��֧,P.��Ͽ���,T.�ɱ����� As �������,X.�ɱ����� As ��ϱ���,X.�ɱ����� As �������,L.���� As �������� From (" & strSQL & ") W,������ĿĿ¼_�ɱ� T,�����Ŀ�嵥_�ɱ� P,�����Ա����_�ɱ� K,�����Ͻ���_�ɱ� X,��������Ŀ¼ L " & _
                        "WHERE T.������Ŀid=W.�����Ŀid " & _
                                "And T.�ɱ�����=P.��ϱ��� " & _
                                "AND K.�������=P.������� " & _
                                "AND K.�ײͱ���=P.�ײͱ��� " & _
                                "AND K.�ײ����=P.�ײ���� " & _
                                "AND X.����id=W.����id " & _
                                "AND L.ID=W.����id " & _
                                "AND K.����id=" & Val(varParam(1)) & " And K.�������='" & varParam(2) & "' Order By P.��Ͽ���,P.��ϱ���"
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
                 "���˲�����¼ Y, " & _
                  "(select ����id,0 AS �����,'' AS ��Ŀ,���� AS ��� from ���˲����ı��� ) X1 " & _
            "Where x.������¼id = A.��첡��ID " & _
                  "AND X.ID=X1.����id " & _
                  "AND Y.ID=X.������¼id " & _
                  "AND A.����ID=" & Val(varParam(1)) & _
                  " AND A.�Ǽ�ID=" & Val(varParam(0)) & _
                  " AND X.Ԫ������=4 and X.Ԫ�ر���='000055' " & _
            ") ORDER BY  �������,�����"

    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




