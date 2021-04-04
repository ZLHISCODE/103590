Attribute VB_Name = "mdlPublic"
Option Explicit
Public gcnOracle As New ADODB.Connection            '�������ݿ�����
Public gstrSql As String                            '����SQL�ִ�
Public mclsZip As New cZip
Public mclsUnzip As New cUnzip
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public gobjComLib As Object                         '������������
Private gstrSysName As String

Private gstrDbUser As String                 '��ǰ���ݿ��û�
Private glngUserId As Long                   '��ǰ�û�id
Private gstrUserCode As String               '��ǰ�û�����
Private gstrUserName As String               '��ǰ�û�����
Private gstrUserAbbr As String               '��ǰ�û�����

Private glngDeptId As Long                   '��ǰ�û�����id
Private gstrDeptCode As String               '��ǰ�û����ű���
Private gstrDeptName As String               '��ǰ�û���������
Private gstrPrivs As String                  'Ȩ��
Private gstr�ӿ�Ȩ��  As String              '�ӿڱ������Ȩ

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
        '------------------------------------------------
        '���ܣ� ��ָ�������ݿ�
        '������
        '   strServerName�������ַ���
        '   strUserName���û���
        '   strUserPwd������
        '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
        '------------------------------------------------
        Dim strSQL As String
        Dim strError As String
    
        On Error Resume Next
100     Err = 0
102     DoEvents
104     With gcnOracle
106         If .State = adStateOpen Then .Close
108         .Provider = "MSDataShape"
110         .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, TranPasswd(strUserPwd)
112         If Err <> 0 Then
                '���������Ϣ
114             strError = Err.Description
116             If InStr(strError, "�Զ�������") > 0 Then
118                 WriteLog "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��"
120             ElseIf InStr(strError, "ORA-12154") > 0 Then
122                 WriteLog "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������"
124             ElseIf InStr(strError, "ORA-12541") > 0 Then
126                 WriteLog "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������"
128             ElseIf InStr(strError, "ORA-01033") > 0 Then
130                 WriteLog "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�"
132             ElseIf InStr(strError, "ORA-01034") > 0 Then
134                 WriteLog "ORACLE�����ã������������ݿ�ʵ���Ƿ�������"
136             ElseIf InStr(strError, "ORA-02391") > 0 Then
138                 WriteLog "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��"
140             ElseIf InStr(strError, "ORA-01017") > 0 Then
142                 WriteLog "�����û�������������ָ�������޷���¼��"
144             ElseIf InStr(strError, "ORA-28000") > 0 Then
146                 WriteLog "�����û��Ѿ������ã��޷���¼��"
                Else
148                 WriteLog strError
                End If
            
150             OraDataOpen = False
                Exit Function
            End If
        End With
    
152     Err = 0
        On Error GoTo errHand
    
154     gstrDbUser = UCase(strUserName)
    
    
156     Call gobjComLib.InitCommon(gcnOracle)
158     Call GetUserInfo
160     If CheckRegInfo = True Then
162         gstrPrivs = gobjComLib.GetPrivFunc(100, 1208)
164         OraDataOpen = True
        Else
166         gcnOracle.Close
168         Set gcnOracle = Nothing
        End If
        Exit Function
    
errHand:
170     WriteLog "OraOpen " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
172     OraDataOpen = False
174     Err = 0
End Function

Private Function CheckRegInfo() As Boolean

        Dim objFile As Scripting.TextStream, strLine As String, str���� As String, date���� As Date
        Dim strUnti As String, strCode As String, strKey As String
        Dim strPrivs As String '����Ȩ����
        On Error GoTo hErr
100     strKey = "�¶�"
    
102     If gobjComLib.RegCheck = False Then
104         WriteLog "δ��ͨHISϵͳע����ؼ�飬���ȼ��ZLLIS�ܷ��������У�"
            Exit Function
        End If
106     strUnti = Trim(gobjComLib.zlRegInfo("��λ����", , -1))
108     If gobjFSO.FileExists(App.Path & "\RegFile.ini") Then
110         Set objFile = gobjFSO.OpenTextFile(App.Path & "\RegFile.ini")
        
112         Do Until objFile.AtEndOfLine
114             strLine = objFile.ReadLine
116             If strLine Like "��Ȩ��ֹ����=*" Then
118                 str���� = Trim(Split(strLine, "=")(1))
120             ElseIf strLine Like "��Ȩ��=*" Then
122                 strCode = Trim(Split(strLine, "=")(1))
124             ElseIf strLine Like "��Ȩ����=*" Then
126                 strPrivs = Trim(Split(strLine, "=")(1))
                End If
            Loop
    
128         If IsDate(str����) Then
130             date���� = gobjComLib.zlDatabase.Currentdate
132             If date���� <= CDate(str����) Then
134                 If strCode <> Md5_String_Calc(strUnti & "|" & str���� & "|" & strKey & strPrivs) Then
                    
136                     WriteLog "��Ȩ�벻��ȷ��" & vbNewLine & _
                               "��λ��" & strUnti & vbNewLine & _
                               "ע���룺" & strCode & vbNewLine & _
                               "���ڣ�" & str���� & vbNewLine & _
                               "Ȩ�ޣ�" & strPrivs & vbNewLine & _
                               "ע���ļ���" & App.Path & "\RegFile.ini"
                    
                    Else
138                     CheckRegInfo = True
140                     If strPrivs <> "" Then gstr�ӿ�Ȩ�� = strPrivs
                    End If
                Else
142                 WriteLog "�ѳ����������ޣ�"
                End If
            Else
144             WriteLog "�������ڴ���"
            End If
        Else
146         WriteLog "��������Ŀ¼ȱ����Ȩ�ļ���RegFile.ini����"
        End If
        Exit Function
hErr:
148     WriteLog "CheckReg " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function GetApplication(strPatientID As String) As String
        '=========================================================================================
        '����:                              �õ��������뵥�ļ�¼��
        '����
        'strPatientID                       ����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�
        '=========================================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim lngPatientID As Long
        Dim strData As String, blnBacode As Boolean
    
        'û�в�ѯ����ʱ�˳�
100     If strPatientID = "" Then Exit Function
102     blnBacode = False
        On Error GoTo errH
    
104     Select Case Mid(strPatientID, 1, 1)
            Case "-"
106             gstrSql = "select ����ID,����,�Ա�,����,�����,סԺ��,���￨��,���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,������,IC����,ҽ����,���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and ����id = [1]"
108         Case "+"
110             gstrSql = "select ����ID,����,�Ա�,����,�����,סԺ��,���￨��,���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,������,IC����,ҽ����,���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and a.סԺ�� = [1] "
112         Case "*"
114             gstrSql = "select ����ID,����,�Ա�,����,�����,סԺ��,���￨��,���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,������,IC����,ҽ����,���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and a.����� = [1] "
116         Case "."
118             gstrSql = "select ����ID,����,�Ա�,����,�����,סԺ��,���￨��,���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,������,IC����,ҽ����,���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and a.�Һŵ� = [2] "
120         Case "/"
122             gstrSql = "Select Distinct b.����ID,b.����,b.�Ա�,b.����,b.�����,b.סԺ��,b.���￨��,b.���֤��,c.���� as ��ǰ���ұ���,c.���� as ��ǰ��������,b.������,b.IC����,b.ҽ����,b.����  " & vbNewLine & _
                        "From ���˷��ü�¼ A, ������Ϣ B , ���ű� C " & vbNewLine & _
                        "Where A.����id = B.����id And A.NO = [2] And A.����id Is Not Null And A.�����־ = 1 and b.��ǰ����id = c.id(+) " & vbNewLine & _
                        "Order By ����id Desc"
124         Case "\" '������
126             gstrSql = "select a.����ID,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.���￨��,a.���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,a.������,a.IC����,a.ҽ����,a.���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and a.������ = [2] "
128         Case Else
130             If Len(strPatientID) >= 12 Then
132                 blnBacode = True
134                 gstrSql = "Select Distinct c.����ID,c.����,c.�Ա�,c.����,c.�����,c.סԺ��,c.���￨��,c.���֤��,d.���� as ��ǰ���ұ���,d.���� as ��ǰ��������,c.������,c.IC����,c.ҽ����,c.���� " & vbNewLine & _
                                " From ����ҽ����¼ A, ����ҽ������ B , ������Ϣ C,���ű� d Where A.ID = B.ҽ��id and " & vbNewLine & _
                                " a.����ID = C.����ID and c.��ǰ����ID = d.id(+) And B.�������� = [2] "
                Else
136                 gstrSql = "Select a.����ID,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.���￨��,a.���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,a.������,a.IC����,a.ҽ����,a.���� " & vbNewLine & _
                                "From ������Ϣ a,���ű� b " & vbNewLine & _
                                "Where a.��ǰ����ID = b.id(+) and  ���￨�� = [2] "
                End If
        End Select
    
138     If InStr(",-,+,*,.,/,\,", "," & Mid(strPatientID, 1, 1) & ",") > 0 Then
140         strPatientID = Mid(strPatientID, 2)
        End If
        
142     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "��ȡ���嵥", Val(strPatientID), CStr(strPatientID))
    
        'û���ҵ����ʱ�˳�
144     If rsTmp.EOF = True Then Exit Function
    
        '���ҵ��������ʱ�˳�(���ض�����˵�ID����Ϣ)
146     If rsTmp.RecordCount > 1 Then
148         Do Until rsTmp.EOF
150             strData = strData & "|" & rsTmp("����ID") & "^" & rsTmp("����") & "^" & rsTmp("�Ա�") & "^" & rsTmp("����") & _
                          "^" & rsTmp("�����") & "^" & rsTmp("סԺ��") & "^" & rsTmp("���￨��") & "^" & rsTmp("���֤��") & _
                          "^" & rsTmp("��ǰ���ұ���") & "^" & rsTmp("��ǰ��������") & "^" & rsTmp("������") & "^" & rsTmp("����")
152             rsTmp.MoveNext
            Loop
154         If strData <> "" Then
156             GetApplication = Mid(strData, 2)
            End If
            Exit Function
        End If
    
158     lngPatientID = "" & rsTmp("����ID")
    
        '��ȡ���뵥
    '    gstrSql = "Select A.*, To_Char(B.����ʱ��, 'YYYY-MM-DD HH24:MI') As ����ʱ��, B.ҽ������ As ������Ŀ, B.�걾��λ As �걾����, B.��������id, B.����ҽ��" & vbNewLine & _
                    "From (Select Decode(Sum(Decode(Z.��¼״̬, 1, 1, 0)), 0, 0, 1) As ѡ��, A.���id As ID," & vbNewLine & _
                    "              C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)') As ����, C.�����, C.סԺ��, D.���� As �������, A.����ҽ�� As ������, F.������," & vbNewLine & _
                    "              F.����ʱ��, 'Item' As ͼ��, Decode(Sum(Decode(Z.��¼״̬, 1, 1, 0)), 0, '  ', '��') As �շ�, Nvl(A.������־, 0) As ����, H.��������," & vbNewLine & _
                    "              Max(Decode(I.��Ŀ���, 2, 2, 1)) As ��Ŀ���, Max(F.������) As ������, Max(F.����ʱ��) As ����ʱ��" & vbNewLine & _
                    "       From ����ҽ����¼ A, ������Ϣ C, ���ű� D, ����ҽ������ F, ���鱨����Ŀ G, ������ĿĿ¼ H, ������Ŀ I, ���˷��ü�¼ Z" & vbNewLine & _
                    "       Where A.������� = 'C' And A.����id = C.����id And A.��������id = D.ID And A.���id Is Not Null And A.ҽ��״̬ = 8 And A.ID = F.ҽ��id And" & vbNewLine & _
                    "             A.������Ŀid = G.������Ŀid And G.ϸ��id Is Null And G.������Ŀid = I.������Ŀid And A.������Ŀid = H.ID And F.ִ��״̬ = 0 And" & vbNewLine & _
                    "             A.����id = [1] And F.NO = Z.NO(+) And F.��¼���� = Z.��¼����(+) And F.ҽ��id = Z.ҽ�����(+) + 0" & vbNewLine & _
                    "       Group By A.���id, C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)'), C.�����, C.סԺ��, D.����, A.����ҽ��, 'Item', Nvl(A.������־, 0)," & vbNewLine & _
                    "                H.��������, F.������, F.����ʱ��" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sum(Decode(Z.��¼״̬, 1, 1, 0)), 0, 0, 1) As ѡ��, A.���id As ID," & vbNewLine & _
                    "              C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)') As ����, C.�����, C.סԺ��, D.���� As �������, A.����ҽ�� As ������, F.������," & vbNewLine & _
                    "              F.����ʱ��, 'Item' As ͼ��, Decode(Sum(Decode(Z.��¼״̬, 1, 1, 0)), 0, '  ', '��') As �շ�, Nvl(A.������־, 0) As ����, H.��������," & vbNewLine & _
                    "              Max(Decode(I.��Ŀ���, 2, 2, 1)) As ��Ŀ���, Max(F.������) As ������, Max(F.����ʱ��) As ����ʱ��" & vbNewLine & _
                    "       From ����ҽ����¼ A, ������Ϣ C, ���ű� D, ����ҽ������ F, ���鱨����Ŀ G, ������ĿĿ¼ H, ������Ŀ I, ���˷��ü�¼ Z, ����걾��¼ J, ������Ŀ�ֲ� K" & vbNewLine & _
                    "       Where A.������� = 'C' And A.����id = C.����id And A.��������id = D.ID And A.���id Is Not Null And A.ҽ��״̬ = 8 And A.ID = F.ҽ��id And" & vbNewLine & _
                    "             A.������Ŀid = G.������Ŀid And G.ϸ��id Is Null And G.������Ŀid = I.������Ŀid And A.������Ŀid = H.ID And F.ִ��״̬ = 3 And" & vbNewLine & _
                    "             A.����id = [1] And F.NO = Z.NO(+) And F.��¼���� = Z.��¼����(+) And F.ҽ��id = Z.ҽ�����(+) + 0 And A.���id = K.ҽ��id(+) And" & vbNewLine & _
                    "             J.ID = K.�걾id And J.ID = 0" & vbNewLine & _
                    "       Group By A.���id, C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)'), C.�����, C.סԺ��, D.����, A.����ҽ��, 'Item', Nvl(A.������־, 0)," & vbNewLine & _
                    "                H.��������, F.������, F.����ʱ��) A, ����ҽ����¼ B" & vbNewLine & _
                "Where A.ID = B.ID"
160     gstrSql = "Select Distinct A.���id As ID, D.����, D.�Ա�, D.����, A.������Դ, D.�����, D.סԺ��, E.���� As ������ұ���, E.���� As �����������, A.����ҽ��, A.����ʱ��,D.����,C.��������,C.�����ӡ " & vbNewLine & _
                    "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C, ������Ϣ D, ���ű� E" & vbNewLine & _
                    "Where A.������Ŀid = B.ID And B.��� = 'C' And A.ID = C.ҽ��id And A.���id Is Not Null And C.ִ��״̬ = 0 And A.����id = [1] And" & vbNewLine & _
                    "      A.����id = D.����id And A.��������id = E.ID And A.ҽ��״̬ = 8"
162     If blnBacode Then gstrSql = gstrSql & " And C.�������� = [2] "
    
164     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "��ȡ���뵥", lngPatientID, strPatientID)
166     Do Until rsTmp.EOF
168         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("����") & "^" & rsTmp("�Ա�") & "^" & rsTmp("����") & "^" & rsTmp("������Դ") & _
                      "^" & rsTmp("�����") & "^" & rsTmp("סԺ��") & "^" & rsTmp("������ұ���") & "^" & rsTmp("�����������") & _
                      "^" & rsTmp("����ҽ��") & "^" & rsTmp("����ʱ��") & "^" & rsTmp("����") & "^" & rsTmp("��������") & "^" & rsTmp("�����ӡ")
170         rsTmp.MoveNext
        Loop
172     If strData <> "" Then
174         GetApplication = Mid(strData, 2)
        End If
    
        Exit Function
errH:
176     WriteLog "GetApplication " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '���ܣ� �ر����ݿ�
    '������
    '���أ� �ر����ݿ⣬����True��ʧ�ܣ�����False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

Public Function InsertReport(lngID As Long, strReportPath As String, ErrInfo As String, Optional lngDeviceID As Long, Optional strSampleNo As String, Optional strItems As String) As Boolean
        '===================================================================
        '����                               ���뱨�浽HIS
        '����
        'lngID                              ҽ��ID
        'strReportPath                      ����·��
        '===================================================================
        Dim rsTmp As ADODB.Recordset
        Dim aStrSQL() As String                     '����SQL�ִ�
        Dim intLoop  As Integer
        Dim strZipFile As String                    'ѹ������ļ�
        Dim strUnZipFile As String                  '��ѹ����ļ�
        Dim strPath As String                       '��ʱ�ļ�·��
    
        On Error GoTo errH
    
100     If Dir(strReportPath) = "" Then Exit Function
102     strPath = IIf(Len(App.Path) <= 3, App.Path & "TMP.RTF", App.Path & "\TMP.RTF")
    
104     If gobjFSO.FileExists(strPath) = True Then gobjFSO.DeleteFile strPath
    
106     Call gobjFSO.CopyFile(strReportPath, strPath)
    
108     If gobjFSO.FileExists(strPath) = False Then Exit Function
    
110     gstrSql = "Zl_���鱨�浥_Insert(" & lngID & ",0)"
112     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "���뱨��"
    
114     gstrSql = "Select Nvl(A.����id, 0) As �ļ�id From ����ҽ������ A Where A.ҽ��id = [1] "
116     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "���뱨��", lngID)
118     If rsTmp.EOF = True Then Exit Function
    
    
120     strZipFile = zlFileZip(strPath)
    
122     strUnZipFile = zlFileUnzip(strZipFile)
    
    
124     If zlLisBlobSql(rsTmp("�ļ�ID"), strZipFile, aStrSQL) = False Then Exit Function
    
126     For intLoop = 0 To UBound(aStrSQL)
128         gobjComLib.zlDatabase.ExecuteProcedure Replace(aStrSQL(intLoop), "Call", ""), "���뱨��"
    '        Debug.Print aStrSQL(intLoop)
        Next
130     gobjFSO.DeleteFile strZipFile
132     gobjFSO.DeleteFile strPath
134     InsertReport = True
        Exit Function
errH:
136     ErrInfo = CStr(Erl()) & "," & Err.Description
138     WriteLog "InsertReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description

End Function

Private Function zlLisBlobSql(ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    '���ɱ��汨���ļ�
    'KeyWord �ļ�ID
    'strFile �ļ�·��
    'arySql ���ɵ�SQL����ڴ�������
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 500
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
'            If lngCount = 0 Then strText = "100;" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_���Ӳ�����ʽ_Insert(" & KeyWord & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strZipPath & "TMP.RTF") Then gobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function

Public Function DeleteReport(lngID As Long) As Boolean
        '===================================================================
        '����                               ɾ������
        '����
        'lngID                              ҽ��ID
        '===================================================================
        On Error GoTo errH
100     gstrSql = "Zl_���鱨�浥_Insert(" & lngID & ",1)"
102     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "ɾ������"
104     DeleteReport = True
        Exit Function
errH:
106     WriteLog "DeleteReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description

End Function
    
Public Function GetClinicItem(lngAdivce As Long) As String
        '===================================================================
        '����                               ȡ��Ҫ����������Ŀ����
        '����
        'lngAdivce                          ҽ��ID
        '����                               �ִ���ʽ:������ĿID^������Ŀ����^������Ŀ����^ִ�п��ұ���^ִ�п�������^����^���^�Ƿ��շ�
        '===================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim strData As String, str������Դ As String
        On Error GoTo errH
    
    '    gstrSql = "Select a.������Ŀid as ID, b.���� as ������Ŀ����, b.���� as ������Ŀ����, c.���� as ִ�п��ұ���, C.���� As ִ�п�������,E.ʵ�ս��,E.��׼����,E.��¼״̬,'0' as �Ƿ�ɼ�" & vbNewLine & _
    '            "From ���˷��ü�¼ E,����ҽ������ D,����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
    '            "Where D.��¼����=E.��¼����(+) And D.No=E.No(+) And D.��¼���=E.���(+) And A.�������='C' And a.ID=D.ҽ��Id And A.������Ŀid = B.ID And A.ִ�п���id = C.ID And A.���id = [1] " & _
    '            "Union all " & _
    '            "Select a.������Ŀid as ID, b.���� as ������Ŀ����, b.���� as ������Ŀ����, c.���� as ִ�п��ұ���, C.���� As ִ�п�������,E.ʵ�ս��,E.��׼����,E.��¼״̬,'1' as �Ƿ�ɼ�" & vbNewLine & _
    '            "From ���˷��ü�¼ E,����ҽ������ D,����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
    '            "Where D.��¼����=E.��¼����(+) And D.No=E.No(+) And D.��¼���=E.���(+) And A.�������='E' And a.ID=D.ҽ��Id And A.������Ŀid = B.ID And A.ִ�п���id = C.ID And A.id = [1] "
100     gstrSql = "Select ������Դ From ����ҽ����¼ Where ID=[1]"
102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "�õ�������Դ", lngAdivce)
104     Do Until rsTmp.EOF
106         str������Դ = Trim("" & rsTmp!������Դ)
108         rsTmp.MoveNext
        Loop
110     If str������Դ = "4" Then
            '��첡��
112         gstrSql = "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
                    "       Sum(E.��׼����) As ��׼����, E.��¼״̬, '0' As �Ƿ�ɼ�" & vbNewLine & _
                    "From ���˷��ü�¼ E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "Where D.��¼���� = E.��¼����(+) And D.No = E.No(+) And D.ҽ��id = E.ҽ�����(+) And E.��¼״̬(+) <> 2 And A.������� = 'C' And" & vbNewLine & _
                    "      A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.���id = [1]" & vbNewLine & _
                    "Group By A.������Ŀid, B.����, B.����, C.����, C.����, E.��¼״̬" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
                    "       Sum(E.��׼����) As ��׼����, Decode(E.��¼״̬, 1, 1, 3, 1, 0) As �Ʒ�״̬, '1' As �Ƿ�ɼ�" & vbNewLine & _
                    "From ���˷��ü�¼ E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "Where D.��¼���� = E.��¼����(+) And D.No = E.No(+) And D.ҽ��id = E.ҽ�����(+) And E.��¼״̬(+) <> 2 And A.������� = 'E' And" & vbNewLine & _
                    "      A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.Id = [1]" & vbNewLine & _
                    "Group By A.������Ŀid, B.����, B.����, C.����, C.����, E.��¼״̬"


        Else
114         gstrSql = "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.���� * E.����) As ʵ�ս��," & vbNewLine & _
                "       Sum(E.����) As ��׼����, D.�Ʒ�״̬, '0' As �Ƿ�ɼ�, F.��¼״̬" & vbNewLine & _
                "From ���˷��ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                "Where D.ҽ��id = F.ҽ�����(+) And D.No = F.No(+) And D.��¼���� = F.��¼����(+) And D.��¼��� = F.���(+) And F.��¼״̬(+) <> 2 And" & vbNewLine & _
                "      A.Id = E.ҽ��id And A.������� = 'C' And A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.���id = [1]" & vbNewLine & _
                "Group By A.������Ŀid, B.����, B.����, C.����, C.����, D.�Ʒ�״̬, F.��¼״̬" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.���� * E.����) As ʵ�ս��," & vbNewLine & _
                "       Sum(E.����) As ��׼����, D.�Ʒ�״̬, '1' As �Ƿ�ɼ�, F.��¼״̬" & vbNewLine & _
                "From ���˷��ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                "Where D.ҽ��id = F.ҽ�����(+) And D.No = F.No(+) And D.��¼���� = F.��¼����(+) And D.��¼��� = F.���(+) And F.��¼״̬(+) <> 2 And" & vbNewLine & _
                "      A.Id = E.ҽ��id And A.������� = 'E' And A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.Id = [1]" & vbNewLine & _
                "Group By A.������Ŀid, B.����, B.����, C.����, C.����, D.�Ʒ�״̬, F.��¼״̬"

        End If
116     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "�õ�������Ŀ", lngAdivce)
    
118     Do Until rsTmp.EOF
120         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("������Ŀ����") & "^" & rsTmp("������Ŀ����") & "^" & rsTmp("ִ�п��ұ���") & _
                        "^" & rsTmp("ִ�п�������") & "^" & rsTmp("��׼����") & "^" & rsTmp("ʵ�ս��") & "^" & rsTmp("��¼״̬") & "^" & rsTmp("�Ƿ�ɼ�")
122         rsTmp.MoveNext
        Loop

124     If strData <> "" Then
126         GetClinicItem = Mid(strData, 2)
        End If
    
        Exit Function
errH:
128     WriteLog "GetClinicItem " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetItemList(lngClinicID As Long) As String
        '===================================================================
        '����                               ȡ��������Ŀ��ָ����ϸ
        '����
        'lngClinicID                        ������ĿID
        '����
        '===================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim strData As String
        On Error GoTo errH
    
100     gstrSql = "Select B.����, B.������, B.Ӣ���� " & vbNewLine & _
                " From ���鱨����Ŀ A, ����������Ŀ B " & vbNewLine & _
                " Where A.������Ŀid = B.ID And a.������ĿID = [1] "

102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "ȡ��ָ����ϸ", lngClinicID)
104     Do Until rsTmp.EOF
106         strData = strData & "|" & rsTmp("����") & "^" & rsTmp("������") & "^" & rsTmp("Ӣ����")
108         rsTmp.MoveNext
        Loop
    
110     If strData <> "" Then
112         GetItemList = Mid(strData, 2)
        End If
    
        Exit Function
errH:
114     WriteLog "GetItemList " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function SetRegister(lngAdivce As Long, intTag As Integer) As Boolean
        '=====================================================================
        '����                               �걾���ջ�ȡ������
        '����
        'lngAdivce                          ҽ��ID
        'intTag                             1=���� 0=ȡ������ 11-��LIS�к��գ�10-��LIS��ȡ������
        '=====================================================================
        On Error GoTo errH
100     gstrSql = "Zl_����ҽ�����_Edit(" & lngAdivce & "," & intTag & ")"
102     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "���ջ�ȡ������"
104     SetRegister = True

        Exit Function
errH:
106     WriteLog "SetRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetAllItem(Optional strFindItem As String) As String()
        '=====================================================================
        '����                               ȡ�����е�������Ŀ���������
        '����
        'strItem                            ��ѡ�����ұ����������ͬ��������Ŀ��Ŀ
        '����                               ���ҵ���������Ŀ����
        '=====================================================================
        Dim astrItem() As String
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String
        Dim strItem As String
        Dim intLoop As Integer
    
100     ReDim Preserve astrItem(0)
102     gstrSql = "select ID,����,����,�����Ŀ from ������ĿĿ¼  where ��� = 'C' "
104     If strFindItem <> "" Then
106         gstrSql = gstrSql & " And (���� = [1] or ���� like '%[1]%') "
        End If
    
108     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "��ȡ������Ŀ", CStr(strFindItem))
    
110     Do Until rsTmp.EOF
112         strItem = strItem & ";" & rsTmp("����") & "," & rsTmp("����") & "," & rsTmp("�����Ŀ")
114         intLoop = intLoop + 1
116         If intLoop >= 200 Then
118             If astrItem(0) <> "" Then
120                 ReDim Preserve astrItem(UBound(astrItem) + 1)
                End If
122             astrItem(UBound(astrItem)) = Mid(strItem, 2)
124             strItem = ""
126             intLoop = 0
            End If
128         rsTmp.MoveNext
        Loop
130     If intLoop <> 0 Then
132         If astrItem(0) <> "" Then
134             ReDim Preserve astrItem(UBound(astrItem) + 1)
            End If
136         astrItem(UBound(astrItem)) = Mid(strItem, 2)
        End If
    
138     GetAllItem = astrItem
        On Error GoTo errH
        Exit Function
errH:
140     WriteLog "GetAllItem " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function UpdateTestResults(ByVal lngID As Long, ByVal strTestName As String, ByVal strTestTime As String, ByVal strTestResults As String) As String
        '===================================================================
        '����                               ���ؼ����������ϵͳ
        '����
        'lngID                              ҽ��ID
        'strTestName                        ������
        'strTestTime                        ����ʱ�䣬��ʽ 2009-01-01 10:30:01
        'strTestResults                     ҽ��ID��Ӧ�ļ����������ԶԶ��ٸ�����ָ��һ������ϸ��ʽ���£�
        '
        '                                     ������Ŀid;������1;��λ1;�����1��;�����־1|������Ŀid;������2;��λ2;����ο�2;�����־2......
        '
        '                                     ���У������־�� ��ƫ��,ƫ��,�쳣,�մ�����ѡ��һ�����ء�
        '����: �գ���ʾ���³ɹ����ǿգ���ʾ������Ϣ��
        '===================================================================
        Dim strSQL As String
        Dim rsTmp As ADODB.Recordset, i As Integer
        Dim varItem As Variant, strItem As String, str���ָ�� As String, str������Ŀid As String, strErrInfo As String
        Dim strEditSQL() As String
        On Error GoTo errH
    
100     str���ָ�� = ""
102     strErrInfo = ""
104     ReDim strEditSQL(0) As String
    
106     If Not strTestTime Like "####-##-## ##:##:##" Or IsDate(CDate(strTestTime)) = False Then
108         strErrInfo = strErrInfo & "0|�������ڸ�ʽ����ȷ���밴yyyy-MM-dd HH24:MI:SS�ĸ�ʽ������" & vbNewLine
110         UpdateTestResults = strErrInfo
            Exit Function
        End If
            
112     strSQL = "Select /*+Rule */" & vbNewLine & _
                " a.����id, a.�嵥id, a.����id, c.�����, c.���ʱ��, c.���ָ��id, d.������Ŀid, a.�ɼ�ҽ��id, f.����" & vbNewLine & _
                "From ������ĿĿ¼ f, ���ָ��Ŀ¼ d, ���鱨����Ŀ e, ��������� c, ��������� a" & vbNewLine & _
                "Where a.�ɼ�ҽ��id = [1] And a.����id = c.����id And a.����id = c.����id And a.�嵥id = c.�嵥id And" & vbNewLine & _
                "           c.���ָ��id = d.Id And f.�����Ŀ = 0 And d.������Ŀid = e.������Ŀid And e.������Ŀid = f.Id"
            
114     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ָ��", lngID)
116     Do Until rsTmp.EOF
118         str���ָ�� = str���ָ�� & "," & rsTmp!����
120         rsTmp.MoveNext
        Loop
    
122     If str���ָ�� <> "" Then
124         varItem = Split(strTestResults, "|")
126         For i = LBound(varItem) To UBound(varItem)
128             strItem = varItem(i)
130             If InStr(strItem, ";") > 0 Then
132                 If UBound(Split(strItem, ";")) >= 4 Then
134                     If InStr(str���ָ�� & ",", "," & Trim(Split(strItem, ";")(0)) & ",") <= 0 Then
136                         strErrInfo = strErrInfo & "0|����: " & Split(strItem, ";")(0) & "δ�ҵ���Ӧ���룬����!" & vbNewLine
138                     ElseIf InStr(strItem, "'") > 0 Then
140                         strErrInfo = strErrInfo & "0|��" & i & "�������,�����Ų����ڽӿ��г��֣��������" & vbNewLine
142                     ElseIf InStr(strItem, """") > 0 Then
144                         strErrInfo = strErrInfo & "0|��" & i & "�������,˫���Ų����ڽӿ��г��֣��������" & vbNewLine
                        Else
146                         rsTmp.MoveFirst
148                         Do Until rsTmp.EOF
        '                        ����id_In     In ���������.����id%Type,
        '                        ����id_In     In ���������.����id%Type,
        '                        �嵥id_In     In ���������.�嵥id%Type,
        '                        ���ָ��id_In In ���������.���ָ��id%Type,
        '                        ������_In     In ���������.�����%Type,
        '                        ����ʱ��_In   In ���������.���ʱ��%Type,
        '                        ���_In       In ���������.���%Type,
        '                        ��λ_In       In ���������.��λ%Type,
        '                        �ο�_In       In ���������.�ο�%Type,
        '                        ����_In       In ���������.����%Type
150                             If Trim("" & rsTmp!����) = Trim(Split(strItem, ";")(0)) And Trim(Split(strItem, ";")(0)) <> "" Then
152                                 If Trim("" & rsTmp!�����) = "" Then '������һ��
154                                     If strEditSQL(UBound(strEditSQL)) <> "" Then ReDim Preserve strEditSQL(UBound(strEditSQL) + 1)
156                                     strEditSQL(UBound(strEditSQL)) = "Zl_���ָ��_Externaledit(" & rsTmp!����id & "," & rsTmp!����id & "," & rsTmp!�嵥id & "," & rsTmp!���ָ��id & ",'" & strTestName & "',to_date('" & strTestTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                                                                         "'" & Split(strItem, ";")(1) & "','" & Split(strItem, ";")(2) & "','" & Split(strItem, ";")(3) & "','" & Split(strItem, ";")(4) & "')"
                                    Else
158                                      strErrInfo = strErrInfo & "1|��Ŀ" & Val(Split(strItem, ";")(0)) & "�Ѿ��н��" & vbNewLine
                                    End If
                                    Exit Do
                                End If
160                             rsTmp.MoveNext
                            Loop
                        End If
                    Else
162                     strErrInfo = strErrInfo & "0|��" & i & "�������,ȱ����Ŀ�����飡" & vbNewLine
                    End If
                Else
164                 strErrInfo = strErrInfo & "0|��" & i & "�������,��ʽ����ȷ�����飡" & vbNewLine
                End If
            Next
        Else
166         strErrInfo = strErrInfo & "0|δ�ҵ�ҽ��id=" & lngID & "������¼!" & vbNewLine
        End If
    
168     For i = LBound(strEditSQL) To UBound(strEditSQL)
170         If strEditSQL(i) <> "" Then gobjComLib.zlDatabase.ExecuteProcedure strEditSQL(i), "�������ָ��"
        Next
172     UpdateTestResults = strErrInfo
    
        Exit Function
errH:
174     UpdateTestResults = strErrInfo & "0|���ִ���" & CStr(Erl()) & "," & Err.Description
176     WriteLog "UpdateTestResults " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function ZipFile(strPath As String) As String
    ZipFile = zlFileZip(strPath)
End Function

Public Function UnZipFile(strPath As String) As String
    UnZipFile = zlFileUnzip(strPath)
End Function

Public Function zlLISRegister(ByVal lngDeviceID As Long, ByVal lngID As Long, ByVal strSampleNo As String, ByRef strErrInfo As String) As Boolean
        '���ں��ձ걾
        Dim strSQL As String, rsTmp As ADODB.Recordset, rs As New ADODB.Recordset
        Dim lngKey As Long, strItemRecords As String
        Dim lngDeptID As Long '��ǰ��������
        Dim rsItem As New ADODB.Recordset
        Dim strItem As String                           '������Ŀ
        Dim str���� As String, str�Ա� As String, str���� As String
        Dim dtSampleDate As Date, dStart As Date, dEnd As Date
    
        On Error GoTo errH
100     If InStr(gstr�ӿ�Ȩ��, "ZLLIS�걾����") <= 0 Then
102         strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
            Exit Function
        End If
        '������������
104     strSQL = "Select ʹ��С��id From �������� Where ID = [1]"
106     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��������걾", lngDeviceID)
108     lngDeptID = 0
110     If Not rsTmp.EOF Then
112         lngDeptID = Val("" & rsTmp("ʹ��С��id"))
        End If
114     If lngDeptID <= 0 Then
            '�˳�-������ʾ
116         strErrInfo = "��������δָ����Ӧ�ļ���С�飡"
            Exit Function
        End If
118     dtSampleDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
120     strSQL = "Select ID, ����, �Ա�, ����, NO, ��Ŀid, ���, ��־, ����ο�, ����, ����ʱ��, ������, Rownum As �������, ������Ŀid," & vbNewLine & _
                "       ����,�걾��λ,��������ID,����ҽ��,��ʶ��,��ǰ����,���˿��� " & vbNewLine & _
                "From (Select A.���id As ID, C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)') As ����, A.�Ա�, A.����, F.NO," & vbNewLine & _
                "              I.������Ŀid As ��Ŀid, Decode(I.�������, 3, Nvl(I.Ĭ��ֵ, '-'), 2, I.Ĭ��ֵ, '') As ���, '' As ��־," & vbNewLine & _
                "              Trim(Replace(Replace(' ' || Zlgetreference(I.������Ŀid, A.�걾��λ, Decode(A.�Ա�, '��', 1, 'Ů', 2, 0)," & vbNewLine & _
                "                                                          C.��������, Y.����id, A.����), ' .', '0.'), '��.', '��0.')) As ����ο�," & vbNewLine & _
                "              Nvl(A.������־, 0) As ����, F.����ʱ��, F.������, G.�������, A.������Ŀid, M.����, " & vbNewLine & _
                "              a.�걾��λ,��������ID,����ҽ��,decode(a.������Դ,2, decode(nvl(c.סԺ��,''),'',c.�����,c.סԺ��),c.�����) as ��ʶ��,c.��ǰ����,l.���� as ���˿��� " & vbNewLine & _
                "       From ����ҽ����¼ A, ������Ϣ C, ����ҽ������ F, ���鱨����Ŀ G, ������Ŀ I, ����������Ŀ Y, ������ĿĿ¼ M ,���ű� L " & vbNewLine & _
                "       Where A.������� = 'C' And A.����id = C.����id And A.���id Is Not Null And A.ҽ��״̬ = 8 And A.ID = F.ҽ��id And" & vbNewLine & _
                "             A.������Ŀid = G.������Ŀid And G.ϸ��id Is Null And G.������Ŀid = Y.��Ŀid(+) And" & vbNewLine & _
                "             G.������Ŀid = I.������Ŀid And A.������Ŀid = M.ID(+) And a.���˿���ID = l.ID" & vbNewLine & _
                "             and (Y.����id + 0 = [1] Or (Y.����id Is Null And F.ִ�в���id = [2])) And nvl(F.ִ��״̬,0) = 0  And A.���ID = [3]" & vbNewLine & _
                "       Order By M.����, G.�������)"

122     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, lngDeptID, lngID)
124     If rsTmp.EOF Then
126         strErrInfo = "û���ҵ��������룡"
            Exit Function
        End If


        
128     If Val(strSampleNo) <= 0 Then
130         strErrInfo = "�걾�Ŵ�����ֻ֧�ִ���������֣�"
            Exit Function
        Else
132         strSampleNo = Val(strSampleNo)
        End If
134     dStart = CDate(Format(dtSampleDate, "yyyy-MM-dd 00:00:00"))
136     dEnd = CDate(Format(dtSampleDate, "yyyy-MM-dd 23:59:59"))
138     strSQL = "Select ������,����ʱ�� from ����걾��¼ where ����ID=[1] and �걾���=[2] And ����ʱ�� Between [3] and [4]"
140     Set rsItem = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, strSampleNo, dStart, dEnd)
142     If Not rsItem.EOF Then
144         strErrInfo = strSampleNo & "�ű걾�Ѵ��ڣ�" & vbNewLine & "�����ˣ�" & rsItem!������ & " ����ʱ��:" & Format(rsItem!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            Exit Function
        End If
    
146     strSQL = "Select B.����id, B.��ҳid, B.���, B.Ӥ������, B.Ӥ���Ա�" & vbNewLine & _
                        "From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                        "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.Ӥ�� = B.��� And A.���id = [1] And Rownum = 1"
148     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngID)
150     If rs.EOF = False Then
152         str���� = Trim("" & rs("Ӥ������"))
154         str�Ա� = Trim("" & rs("Ӥ���Ա�"))
156         str���� = "Ӥ��"
        Else
158         str���� = Trim("" & rsTmp("����"))
160         str�Ա� = Trim("" & rsTmp("�Ա�"))
162         str���� = Trim("" & rsTmp("����"))
        End If
    
        '����������Ŀ
164     strSQL = "select distinct ҽ������ from ����ҽ����¼ a , ����ҽ������ b, ���鱨����Ŀ c , ����������Ŀ d " & vbNewLine & _
                  "  where a.id = b.ҽ��ID and a.������ĿID = c.������ĿID and " & vbNewLine & _
                  "  c.������ĿID = d.��ĿID(+) and a.���id=[1] "
166     Set rsItem = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lngID)
168     Do Until rsItem.EOF
170         strItem = strItem & " " & Trim("" & rsItem("ҽ������"))
172         rsItem.MoveNext
        Loop
174     strItem = Trim(strItem) & "(" & Trim("" & rsTmp("�걾��λ")) & ")"
        
        '�����걾��¼
        '------------10.25
176     lngKey = gobjComLib.zlDatabase.GetNextId("����걾��¼")
     
178     strSQL = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
            rsTmp("ID") & ",'" & _
            strSampleNo & "'," & _
            IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
            IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
            lngDeviceID & "," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
            "1,'" & _
            gstrUserName & "'," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0,0,0," & _
            rsTmp("����") & ",NULL,'" & _
            str���� & "','" & str�Ա� & "','" & str���� & "','" & Trim("" & rsTmp("No")) & "','" & _
            Trim("" & rsTmp("�걾��λ")) & "'," & Trim("" & rsTmp("��������ID")) & ",'" & Trim("" & rsTmp("����ҽ��")) & "','" & _
            Trim("" & rsTmp("��ʶ��")) & "','" & Trim("" & rsTmp("��ǰ����")) & "','" & Trim("" & rsTmp("���˿���")) & "','" & _
            strItem & "')"
    
        '---------- 10.26 ��SQL
    
    '    gstrSql = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
    '        rsTmp("ID") & ",'" & rsTmp("ID") & "',0,'" & _
    '        strSampleNo & "'," & _
    '        IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
    '        IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
    '        lngDeviceID & "," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
    '        "'" & _
    '        gobjComLib.zlDatabase.GetUserInfo.Fields("����").value & "'," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0," & _
    '        intType & ",NULL,'" & _
    '        str���� & "','" & str�Ա� & "','" & str���� & "','" & Trim("" & rsTmp("No")) & "','" & _
    '        Trim("" & rsTmp("�걾��λ")) & "'," & Trim("" & rsTmp("��������ID")) & ",'" & Trim("" & rsTmp("����ҽ��")) & "'," & _
    '        Trim("" & rsTmp("��ʶ��")) & ",'" & Trim("" & rsTmp("��ǰ����")) & "','" & Trim("" & rsTmp("���˿���")) & "','" & _
    '        strItem & "',Null,Null,Null,'" & gstrUserCode & "','" & gstrUserName & "')"
    
        '-------------------------------------------------------------------------------------
    
180     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "��������걾"
                                                                
        '��дָ��
182     strItemRecords = ""
184     Do While Not rsTmp.EOF
186         strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("��ĿID") & "^" & _
                Trim("" & rsTmp("���")) & "^" & Val("" & rsTmp("��־")) & "^" & Trim("" & rsTmp("����ο�")) & "^" & _
                Trim("" & rsTmp("������ĿID")) & "^" & Trim("" & rsTmp("�������"))
            
188         rsTmp.MoveNext
        Loop
    
190     If Len(strItemRecords) > 0 Then
192         strItemRecords = Mid(strItemRecords, 2)
            
194         strSQL = "Zl_������ͨ���_Write(" & lngKey & "," & _
                lngDeviceID & ",'" & strItemRecords & "',0,0)"
196         gobjComLib.zlDatabase.ExecuteProcedure strSQL, "��������걾"
        End If
    
198     zlLISRegister = True
        Exit Function
errH:
        'Resume
200     strErrInfo = CStr(Erl()) & "," & Err.Description
202     WriteLog "zlLISRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnRegister(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
        'ȡ����ZLLIS���Ѻ��յı걾
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     If InStr(gstr�ӿ�Ȩ��, "ZLLISȡ������") <= 0 Then
102         strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
            Exit Function
        End If
        '�Ƿ��ȡ�����յĲ����ڴ洢�����У����Դ˴��������
104     strSQL = "Zl_����걾��¼_ȡ������(" & lngID & ")"
106     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "ȡ������"
108     zlLisUnRegister = True
        Exit Function
errH:
110     strErrInfo = CStr(Erl()) & "," & Err.Description
112     WriteLog "zlLisUnRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function ZLLisInsterReport(ByVal lngID As Long, strItems As String, ByRef strErrInfo As String) As Boolean
        Dim str�걾 As String, lng����ID As Long, str�Ա� As String, str����  As String
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim str��Ŀ As String, varItem As Variant
        On Error GoTo errH
100     If InStr(gstr�ӿ�Ȩ��, "ZLLIS�걾���") <= 0 Then
102         strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
            Exit Function
        End If
104     If InStr(strItems, "'") > 0 Then
106         strErrInfo = "��������������ţ�"
            Exit Function
108     ElseIf InStr(strItems, """") > 0 Then
110         strErrInfo = "���������˫���ţ�"
            Exit Function
112     ElseIf InStr(strItems, "^") < 0 Then
114         strErrInfo = "�����ٴ���һ�������"
            Exit Function
        End If
    
116     strSQL = "Select b.Id, b.�����,b.�Ա�, b.����id, b.�걾����, to_char(b.��������,'YYYY-MM-DD HH24:MI:SS') as ��������, b.΢����걾" & vbNewLine & _
                "From ����ҽ����¼ A, ����걾��¼ B" & vbNewLine & _
                "Where a.Id = b.ҽ��id(+) And a.Id = [1]"

118     Set rsSample = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lngID)
    
120     If rsSample.EOF Then
122         strErrInfo = "δ�ҵ���Ӧҽ����"
            Exit Function
        End If
    
124     If Trim("" & rsSample!�����) <> "" Then
126         strErrInfo = "����˱걾�������޸ģ�"
            Exit Function
        End If
    
128     If InStr(1, gstrPrivs, "��˱걾") <= 0 Then
130         strErrInfo = "��û��Ȩ�޽������,�����µ�½���������Ա�������!"
            Exit Function
        End If
    
        '11210 Ȩ�ޡ�δ�շ���ˡ�������˵�������ʱ��δ��Ч��
132     If InStr(gstrPrivs, "δ�շ����") <= 0 Then
134         strErrInfo = CheckChargeState(lngID, False)
136         If strErrInfo <> "" Then Exit Function
        End If
    
        '21137 �ѹ鵵���治�����
138     gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
140     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "��˼��", lngID)
142     If rsTmp.EOF = False Then
144         strErrInfo = "���˱���סԺ�Ĳ������ύ��飬���ܽ�����ˣ�"
            Exit Function
        End If
    
        '���סԺ�����Ƿ��Ժ���л��۵�
146     strErrInfo = CheckExesState(lngID)
148     If strErrInfo <> "" Then Exit Function

        '������תΪ��ĿID
        Dim i As Integer, strCode As String, strValue As String
150     varItem = Split(strItems, "|")
152     str��Ŀ = ""
154     strErrInfo = ""
156     For i = LBound(varItem) To UBound(varItem)
158         If InStr(varItem(i), "^") > 0 Then
160             strCode = Trim(Split(varItem(i), "^")(0))
162             strValue = Split(varItem(i), "^")(1)
            
164             gstrSql = "Select A.������ĿID,B.����, B.������, B.Ӣ���� " & vbNewLine & _
                    " From ���鱨����Ŀ A, ����������Ŀ B " & vbNewLine & _
                    " Where A.������Ŀid = B.ID And B.���� = [1] "
    
166             Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "ȡ��ָ��ID", strCode)
168             If rsTmp.EOF Then
170                 strErrInfo = strErrInfo & vbNewLine & strCode & " δ�ҵ���Ӧ��Ŀ!"
                Else
172                 str��Ŀ = str��Ŀ & "|" & rsTmp!������ĿID & "^" & strValue
                End If
            
            End If
        Next
174     If strErrInfo <> "" Then
            Exit Function
176     ElseIf str��Ŀ = "" Then
178         strErrInfo = "û��Ҫ���µ����ݣ�"
            Exit Function
        End If
180     str��Ŀ = Mid(str��Ŀ, 2)
        '����
182     str�Ա� = Trim("" & rsSample!�Ա�)
184     If str�Ա� = "��" Then
186         str�Ա� = "1"
188     ElseIf str�Ա� = "Ů" Then
190         str�Ա� = "2"
        Else
192         str�Ա� = "9"
        End If
194     strSQL = "ZL_������ͨ���_BATCHUPDATE(" & rsSample!ID & "," & _
                        rsSample!����ID & ",'" & Trim("" & rsSample!�걾����) & "'," & str�Ա� & "," & _
                        IIf(Trim("" & rsSample!��������) = "", "Null", "To_Date('" & Trim("" & rsSample!��������) & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                        str��Ŀ & "'," & rsSample!΢����걾 & ")"
196     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "��д���"

        '���
198     strSQL = "ZL_����걾��¼_�������(" & rsSample!ID & ",'" & gstrUserName & "')"
200     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "��˱���"
202     ZLLisInsterReport = True
        Exit Function
errH:
204     ZLLisInsterReport = False
206     strErrInfo = CStr(Erl()) & "," & Err.Description
208     WriteLog "ZLLisInsterReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnAudit(ByVal lngID As Long, strErrInfo As String) As Boolean
        'ȡ�����
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim d���ʱ�� As Date, dCurr As Date
        On Error GoTo errH
100     If InStr(gstr�ӿ�Ȩ��, "ZLLISȡ�����") <= 0 Then
102         strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
            Exit Function
        End If
    
104     strSQL = "Select a.ID,a.��ӡ����, a.���ʱ�� From ����걾��¼ A Where ҽ��ID=[1]"
106     Set rsSample = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ȡ����˼��", lngID)
    
108     If rsSample.EOF Then
110         strErrInfo = "δ�ҵ���Ӧ�����¼��"
            Exit Function
        End If
112     If IsNull(rsSample!���ʱ��) Then
114         strErrInfo = "�걾δ��ˣ�����ȡ����ˣ�"
            Exit Function
        End If
    
116     If InStr(";" & gstrPrivs & ";", ";���ȡ��;") <= 0 Then
118         d���ʱ�� = rsTmp!���ʱ��
120         dCurr = gobjComLib.zlDatabase.Currentdate
122         If DateDiff("h", d���ʱ��, dCurr) > 24 Then
124             strErrInfo = "ֻ��ȡ��24Сʱ�ڵ���˱��浥������ϵ�ϼ���ʦȡ�����!"
                Exit Function
            End If
        End If
        '21434
126     If InStr(";" & gstrPrivs & ";", ";�����Ѵ�ӡ�ɻع�;") <= 0 Then
128         If Val("" & rsSample!��ӡ����) > 0 Then
130             strErrInfo = "ֻ��ȡ��δ��ӡ����˱��浥������ϵ�ϼ���ʦȡ�����!"
                Exit Function
            End If
        End If
        '21137 �ѹ鵵���治��ȡ��
132     gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ҽ��ID=[1] " & vbNewLine & _
                " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
134     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "ȡ�����", lngID)
136     If rsTmp.EOF = False Then
138         strErrInfo = "���˱���סԺ�Ĳ������ύ��飬����ȡ����ˣ�"
            Exit Function
        End If
    
140     strSQL = "ZL_����걾��¼_���ȡ��(" & rsSample!ID & ")"
142     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "ȡ�����"
144     zlLisUnAudit = True
        Exit Function
errH:
146     strErrInfo = CStr(Erl()) & " " & Err.Description
148     WriteLog "zlLisUnAudit " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetAllDevice(ByRef strErrInfo As String) As String
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     strSQL = "Select ID,����,���� From ��������"
102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ȡ��������")
104     GetAllDevice = ""
106     Do Until rsTmp.EOF
108         GetAllDevice = GetAllDevice & "|" & rsTmp!ID & "^" & rsTmp!���� & "^" & rsTmp!����
110         rsTmp.MoveNext
        Loop
112     If GetAllDevice <> "" Then GetAllDevice = Mid(GetAllDevice, 2)
114     If GetAllDevice = "" Then
116         strErrInfo = "û�г�ʼ��������"
        End If
        Exit Function
errH:
118     strErrInfo = CStr(Erl()) & "," & " " & Err.Description
120     WriteLog "GetAllDevice " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Private Sub GetUserInfo()
    '����:�õ��û�����Ϣ

        Dim rsTemp As New ADODB.Recordset
        On Error GoTo errHand
100     glngUserId = 0
102     gstrUserCode = ""
104     gstrUserName = ""
106     gstrUserAbbr = ""
108     glngDeptId = 0
110     gstrDeptCode = ""
112     gstrDeptName = ""
    
114     Set rsTemp = gobjComLib.zlDatabase.GetUserInfo
    
116     Do Until rsTemp.EOF
118         glngUserId = Val("" & rsTemp.Fields("ID").value)               '��ǰ�û�id
120         gstrUserCode = "" & rsTemp.Fields("���").value            '��ǰ�û�����
122         gstrUserName = "" & rsTemp.Fields("����").value            '��ǰ�û�����
124         gstrUserAbbr = "" & rsTemp.Fields("����").value          '��ǰ�û�����
126         glngDeptId = Val("" & rsTemp.Fields("����id").value)            '��ǰ�û�����id
128         gstrDeptCode = "" & rsTemp.Fields("������").value        '��ǰ�û�
130         gstrDeptName = "" & rsTemp.Fields("������").value        '��ǰ�û�
    
132         rsTemp.MoveNext
        Loop
        Exit Sub
errHand:
134     WriteLog "GetUser " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
136     Err = 0
End Sub


Private Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As String
        '�����շ�״̬
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        Dim strSQLbak As String
        Dim intPatientType As Integer               '������Դ
        On Error GoTo errH
    
100     CheckChargeState = "����δ�շѣ����ܽ�����ˣ�"
    
102     strSQL = "select ������Դ from ����걾��¼ where id = [1]"
104     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��������", lngKey)
106     If rs.EOF = True Then Exit Function
108     intPatientType = rs("������Դ")
    
110     If blnOrder Then
112         strSQL = _
                "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                      "from ���˷��ü�¼ A, " & _
                      "( " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id))  " & _
                           "Union " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id)) " & _
                      ") B " & _
                    "Where A.NO = B.NO "
    '        If intPatientType <> 2 Then
    '            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    '        End If
        Else
114         strSQL = _
                "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                      "from ���˷��ü�¼ A, " & _
                      "( " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                           "Union " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                      ") B " & _
                    "Where A.NO = B.NO and a.��¼���� = b.��¼���� "
    '        If intPatientType <> 2 Then
    '            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    '        End If
        End If
    
116     strSQL = strSQL & " Order by ��¼״̬ "
    
118     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

120     If rs.BOF Then Exit Function
122     If rs("��¼״̬").value = 0 Then Exit Function
    
124     CheckChargeState = ""
        Exit Function
errH:
126     CheckChargeState = CStr(Erl()) & "," & Err.Description
128     WriteLog "CheckChargeState " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Private Function CheckExesState(lngKey As Long) As String
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '����:      ���סԺ���˳�Ժ���Ƿ��л��۵���Ҫ�������
        '����       �걾ID
        '����       �л��۵�δ��� = Fasle û���� = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim rsTmp As New ADODB.Recordset
        On Error GoTo errH
100     CheckExesState = ""
    
        '81��ϵͳ����Чʱ�����
        '    ִ�к��Զ���˻��۵�
        '    ָ����Ӧ��ҽ��ִ�к�(����ҩƷ��ҩ)��0-������1-�Զ���˻��۵�Ϊ���ʵ���
102     If gobjComLib.zlDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
        '��ǰ�����Ƿ��ѳ�Ժ��Ԥ��Ժ
104     gstrSql = "select d.no" & vbNewLine & _
                "from (select distinct d.ҽ��id" & vbNewLine & _
                "       from ����걾��¼ a, ������Ϣ b, ������ҳ c, ������Ŀ�ֲ� d" & vbNewLine & _
                "       where a.����id = b.����id and a.����id = c.����id and a.��ҳid = c.��ҳid and" & vbNewLine & _
                "             a.id = [1] and a.������Դ = 2 and (b.��Ժʱ�� is not null or c.״̬ = 3) and" & vbNewLine & _
                "             a.id = d.�걾id) a, ����ҽ����¼ b, ����ҽ������ c, ���˷��ü�¼ d" & vbNewLine & _
                "where a.ҽ��id in (b.���id, b.id) and b.id = c.ҽ��id and c.��¼���� = d.��¼���� and" & vbNewLine & _
                "      c.no = d.no and d.��¼���� = 2 and d.��¼״̬ = 0 "
106     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "���鼼ʦ����վ-����״̬���", lngKey)
    
108     If rsTmp.EOF Then
110         CheckExesState = ""
        Else
112         CheckExesState = "��ǰסԺ���˻��л��۵�δ��ˣ����ѳ�Ժ��Ԥ��Ժ��"
        End If
        Exit Function
errH:
114     CheckExesState = CStr(Erl()) & "," & Err.Description
116     WriteLog "CheckExesState " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function


Private Sub WriteLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��

    strFileName = App.Path & "\zlLisInterface_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not gobjFSO.FileExists(strFileName) Then Call gobjFSO.CreateTextFile(strFileName)
    Set objStream = gobjFSO.OpenTextFile(strFileName, ForAppending)
    
    objStream.WriteLine (strDate & ":" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub
