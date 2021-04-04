Attribute VB_Name = "mdlPublic"
Option Explicit
Public gcnOracle As ADODB.Connection             '�������ݿ�����
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
Private mstrVirtualHis  As String           '�ӿڱ��������ģ����Ȩ
Private mstrVirtualPeis As String           '���Ȩ�޼��
Private mstrR() As String                    '���沿������
Const CONS_DEBUG = 0                         '���Բ�������

Private mblnOldVer As Boolean               '�Ƿ��ϵ���Ȩ��ʽ
Private mblnVerifyTotal As Boolean          '���۵�ת���ʵ�ʱ���Ƿ���Ƿ�����

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
        '------------------------------------------------
        '���ܣ� ��ָ�������ݿ�
        '������
        '   strServerName�������ַ���
        '   strUserName���û���
        '   strUserPwd������
        '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
        '------------------------------------------------
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strError As String
        Dim objRegister As Object
        On Error Resume Next
100     Err = 0
102     DoEvents
        On Error GoTo errH
        Set objRegister = CreateObject("zlRegister.clsRegister")
        If Not objRegister Is Nothing Then
            If Not objRegister.LoginValidate(strServerName, strUserName, strUserPwd, strError) Then
                If strError <> "" And strError <> "���뷵�ش�����Ϣ" Then
                    MsgBox strError, vbInformation, "������Ϣ"
                    OraDataOpen = False
                    Set objRegister = Nothing
                End If
                Exit Function
            End If
        Else
errH:
104         If gcnOracle Is Nothing Then Set gcnOracle = New ADODB.Connection
106         With gcnOracle
                
108             If .State = adStateOpen Then .Close
110             .Provider = "MSDataShape"
112             .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, TranPasswd(strUserPwd)
114             If Err <> 0 Then
                    '���������Ϣ
116                 strError = Err.Description
118                 If InStr(strError, "�Զ�������") > 0 Then
120                     WriteLog "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��"
122                 ElseIf InStr(strError, "ORA-12154") > 0 Then
124                     WriteLog "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������"
126                 ElseIf InStr(strError, "ORA-12541") > 0 Then
128                     WriteLog "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������"
130                 ElseIf InStr(strError, "ORA-01033") > 0 Then
132                     WriteLog "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�"
134                 ElseIf InStr(strError, "ORA-01034") > 0 Then
136                     WriteLog "ORACLE�����ã������������ݿ�ʵ���Ƿ�������"
138                 ElseIf InStr(strError, "ORA-02391") > 0 Then
140                     WriteLog "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��"
142                 ElseIf InStr(strError, "ORA-01017") > 0 Then
144                     WriteLog "�����û�������������ָ�������޷���¼��"
146                 ElseIf InStr(strError, "ORA-28000") > 0 Then
148                     WriteLog "�����û��Ѿ������ã��޷���¼��"
                    Else
150                     WriteLog CStr(Erl()) & "," & strError
                    End If
                
152                 OraDataOpen = False
                    Exit Function
                End If
            End With
        End If
    
154     Err = 0
        On Error GoTo errHand
    
156     gstrDbUser = UCase(strUserName)
    
    
158     Call gobjComLib.InitCommon(gcnOracle)
160     Call gobjComLib.SetDbUser(gstrDbUser)
        
        '2012-05-23 ��ȡ���ز�������
        mblnOldVer = True
        gstrSql = "Select �汾�� From zlSystems Where ���=100"
        Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ�汾��")
        Do Until rsTmp.EOF
            WriteLog "�汾��" & rsTmp!�汾��
            
            If Trim$("" & rsTmp!�汾��) >= "10.30.10" Then mblnOldVer = False
            WriteLog "�汾��" & rsTmp!�汾�� & "," & IIf(mblnOldVer, "True", "False")
            rsTmp.MoveNext
        Loop
162     Call GetUserInfo
        
164     If mblnOldVer Then
166         If CheckRegInfo = True Then
168             gstrPrivs = gobjComLib.GetPrivFunc(100, 1215)
170             OraDataOpen = True
            Else
172             gcnOracle.Close
174             Set gcnOracle = Nothing
            End If
        Else
176         mstrVirtualHis = gobjComLib.GetPrivFunc(100, 1215)    '��ȡ����ģ��Ȩ��
178         If Right(mstrVirtualHis, 1) <> ";" Then mstrVirtualHis = mstrVirtualHis & ";"
180         If Left(mstrVirtualHis, 1) <> ";" Then mstrVirtualHis = ";" & mstrVirtualHis

182         If InStr(mstrVirtualHis, ";����;") <= 0 Then
184             WriteLog "OraOpen,Ȩ�޲��㣬���ڹ������н���1215 ������LIS�ӿڡ�Ȩ������" & gstrDbUser
186             gcnOracle.Close
188             Set gcnOracle = Nothing
            Else
190             OraDataOpen = True
            End If

192         mstrVirtualPeis = gobjComLib.GetPrivFunc(2100, 2138)    '��ȡ�������ģ��Ȩ��
194         If Right(mstrVirtualPeis, 1) <> ";" Then mstrVirtualPeis = mstrVirtualPeis & ";"
196         If Left(mstrVirtualPeis, 1) <> ";" Then mstrVirtualPeis = ";" & mstrVirtualPeis
    
198         gstrPrivs = gobjComLib.GetPrivFunc(100, 1215)
200         If Right(gstrPrivs, 1) <> ";" Then gstrPrivs = gstrPrivs & ";"
202         If Left(gstrPrivs, 1) <> ";" Then gstrPrivs = ";" & gstrPrivs
        End If
         
        Exit Function
    
errHand:
204     WriteLog "OraOpen " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
206     OraDataOpen = False
208     Err = 0
End Function

Private Function CheckRegInfo() As Boolean

        Dim objFile As Scripting.TextStream, strLine As String, str���� As String, date���� As Date
        Dim strUnti As String, strCode As String, strKey As String
        Dim strPrivs As String '����Ȩ����
        On Error GoTo hErr
    
100     strKey = "�¶�"
102     If CONS_DEBUG = 0 Then
104         If gobjComLib.RegCheck = False Then
106             WriteLog "δ��ͨHISϵͳע����ؼ�飬���ȼ��ZLLIS�ܷ��������У�"
                Exit Function
            End If
108         strUnti = gobjComLib.zlRegInfo("��λ����", , -1)
        End If

110     If gobjFSO.FileExists(App.Path & "\RegFile.ini") Then
112         Set objFile = gobjFSO.OpenTextFile(App.Path & "\RegFile.ini")
        
114         Do Until objFile.AtEndOfLine
116             strLine = objFile.ReadLine
118             If strLine Like "��Ȩ��ֹ����=*" Then
120                 str���� = Trim(Split(strLine, "=")(1))
122             ElseIf strLine Like "��Ȩ��=*" Then
124                 strCode = Trim(Split(strLine, "=")(1))
126             ElseIf strLine Like "��Ȩ����=*" Then
128                 strPrivs = Trim(Split(strLine, "=")(1))
                ElseIf strLine Like "���ʼ�����" Then
                    If Trim(Split(strLine, "=")(1)) = "1" Then
                        mblnVerifyTotal = True
                    Else
                        mblnVerifyTotal = False
                    End If
                End If
            Loop
            
            '���Բ���������һ���̶�����
130         If CONS_DEBUG = 1 Then str���� = "2011-12-01"
            
132         If IsDate(str����) Then
134             date���� = gobjComLib.zlDataBase.Currentdate
136             If date���� <= CDate(str����) Then
138                 If CONS_DEBUG = 1 Then
140                     CheckRegInfo = True
142                     If strPrivs <> "" Then mstrVirtualHis = strPrivs
                    Else
144                     If strCode <> Md5_String_Calc(strUnti & "|" & str���� & "|" & strKey & strPrivs) Then
                        
146                         WriteLog "��Ȩ�벻��ȷ��" & vbNewLine & _
                                   strUnti & vbNewLine & _
                                   strCode & vbNewLine & _
                                   strPrivs & vbNewLine & _
                                   str����
    
                        Else
148                         CheckRegInfo = True
150                         If strPrivs <> "" Then mstrVirtualHis = strPrivs
                        End If
                    End If
                Else
152                 WriteLog IIf(CONS_DEBUG = 1, "���Բ���-", "��ʽ����-") & "�ѳ����������ޣ�" & str����
                End If
            Else
154             WriteLog "�������ڴ���"
            End If
        Else
156         WriteLog "��������Ŀ¼(" & App.Path & ")ȱ����Ȩ�ļ���RegFile.ini����"
        End If
        Exit Function
hErr:
158     WriteLog "CheckReg " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
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
        Dim rsTmp As New ADODB.Recordset, rsProvisional As Recordset
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
                        "From ������ü�¼ A, ������Ϣ B , ���ű� C " & vbNewLine & _
                        "Where A.����id = B.����id And A.NO = [2] And A.����id Is Not Null And A.�����־ = 1 and b.��ǰ����id = c.id(+) " & vbNewLine & _
                        "Order By ����id Desc"
124         Case "\" '������
126             gstrSql = "select a.����ID,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.���￨��,a.���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,a.������,a.IC����,a.ҽ����,a.���� " & _
                         ",��ǰ���� from ������Ϣ a , ���ű� b where a.��ǰ����ID = b.ID(+) and a.������ = [2] "
128         Case "!" '���￨
130             gstrSql = "Select a.����ID,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.���￨��,a.���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,a.������,a.IC����,a.ҽ����,a.���� " & vbNewLine & _
                            "From ������Ϣ a,���ű� b " & vbNewLine & _
                            "Where a.��ǰ����ID = b.id(+) and  ���￨�� = [2] "
132         Case "=" '���֤��
134             gstrSql = "Select a.����ID,a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.���￨��,a.���֤��,b.���� as ��ǰ���ұ���,b.���� as ��ǰ��������,a.������,a.IC����,a.ҽ����,a.���� " & vbNewLine & _
                        "From ������Ϣ a,���ű� b " & vbNewLine & _
                        "Where a.��ǰ����ID = b.id(+) and  ���֤�� = [2] "
136         Case Else
140             blnBacode = True
142             gstrSql = "Select Distinct c.����ID,c.����,c.�Ա�,c.����,c.�����,c.סԺ��,c.���￨��,c.���֤��,d.���� as ��ǰ���ұ���,d.���� as ��ǰ��������,c.������,c.IC����,c.ҽ����,c.���� " & vbNewLine & _
                        " From ����ҽ����¼ A, ����ҽ������ B , ������Ϣ C,���ű� d Where A.ID = B.ҽ��id and " & vbNewLine & _
                        " a.����ID = C.����ID and c.��ǰ����ID = d.id(+) And B.�������� = [2] "
        End Select
    
144     If InStr(",-,+,*,.,/,\,!,=,", "," & Mid(strPatientID, 1, 1) & ",") > 0 Then
146         strPatientID = Mid(strPatientID, 2)
        End If
        
148     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ���嵥", Val(strPatientID), CStr(strPatientID))
    
        'û���ҵ����ʱ�˳�
150     If rsTmp.EOF = True Then Exit Function
    
        '���ҵ��������ʱ�˳�(���ض�����˵�ID����Ϣ)
152     If rsTmp.RecordCount > 1 Then
154         Do Until rsTmp.EOF
156             strData = strData & "|" & rsTmp("����ID") & "^" & rsTmp("����") & "^" & rsTmp("�Ա�") & "^" & rsTmp("����") & _
                          "^" & rsTmp("�����") & "^" & rsTmp("סԺ��") & "^" & rsTmp("���￨��") & "^" & rsTmp("���֤��") & _
                          "^" & rsTmp("��ǰ���ұ���") & "^" & rsTmp("��ǰ��������") & "^" & rsTmp("������") & "^" & rsTmp("����")
158             rsTmp.MoveNext
            Loop
160         If strData <> "" Then
162             GetApplication = Mid(strData, 2)
            End If
            Exit Function
        End If
    
164     lngPatientID = "" & rsTmp("����ID")
    
        '��ȡ���뵥
166     gstrSql = "Select Distinct A.���id As ID, D.����, D.�Ա�, D.����, A.������Դ, D.�����, D.סԺ��, E.���� As ������ұ���, E.���� As �����������, F.��� As ҽ�����,A.����ҽ��, A.����ʱ��,D.����,C.��������,C.�����ӡ,D.��ǰ����,G.���� As ���˿��ұ���, G.���� As ���˿�������,A.Ӥ�� " & vbNewLine & _
                    "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C, ������Ϣ D, ���ű� E, ��Ա�� F, ���ű� G" & vbNewLine & _
                    "Where A.������Ŀid = B.ID And B.��� = 'C' And A.ID = C.ҽ��id And A.���id Is Not Null And C.ִ��״̬ = 0 And A.����id = [1] And" & vbNewLine & _
                    "      A.����id = D.����id And A.��������id = E.ID And A.ҽ��״̬ = 8 And A.����ҽ��=F.����(+) And A.���˿���id=G.ID"
168     If blnBacode Then gstrSql = gstrSql & " And C.�������� = [2] "
    
170     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ���뵥", lngPatientID, strPatientID)
172     Do Until rsTmp.EOF
174         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("����") & "^" & rsTmp("�Ա�") & "^" & rsTmp("����") & "^" & rsTmp("������Դ") & _
                      "^" & rsTmp("�����") & "^" & rsTmp("סԺ��") & "^" & rsTmp("������ұ���") & "^" & rsTmp("�����������") & _
                      "^" & rsTmp("����ҽ��") & "^" & rsTmp("����ʱ��") & "^" & rsTmp("����") & "^" & rsTmp("��������") & "^" & rsTmp("�����ӡ") & _
                      "^" & rsTmp!ҽ����� & "^" & rsTmp!��ǰ���� & "^" & rsTmp!���˿��ұ��� & "^" & rsTmp!���˿������� & "^" & rsTmp!Ӥ��
                      
176         gstrSql = "Select ���� From ����ҽ������ Where  ��Ŀ Like '%���' And  ҽ��id=[1]"
178         Set rsProvisional = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ���", CLng(Val("" & rsTmp("ID"))))
180         If rsProvisional.EOF Then
182             strData = strData & "^"
            Else
184             strData = strData & "^" & Trim("" & rsProvisional!����)
            End If
186         rsTmp.MoveNext
        Loop
188     If strData <> "" Then
190         GetApplication = Mid(strData, 2)
        End If
    
        Exit Function
errH:
192     WriteLog CStr(Erl()) & "�г��ִ���" & Err.Number & " " & Err.Description
End Function

Public Function GetDeptPatiList(ByVal strDeptNo As String, ByRef strReturn As String, ByVal lngType As Long, ByVal strStartDate As String, ByVal strEndDate As String, ByRef ErrInfo As String) As Boolean
        Dim rsTmp As ADODB.Recordset
        Dim strData As String
        Dim lngCount As Long, lngI As Long
        Dim strR() As String
        
        On Error GoTo errH
        
100     If lngType <= 0 Then ReDim mstrR(0) As String
102     If strDeptNo = "" Then
104         ErrInfo = "���ұ��벻��Ϊ��"
            Exit Function
        End If
106     If Not (IsDate(strStartDate)) Or Not (strStartDate Like "####-##-##") Then
108         ErrInfo = "��ʼ���ڸ�ʽ��ΪYYYY-MM-DD"
            Exit Function
        End If
        
110     If Not (IsDate(strEndDate)) Or Not (strEndDate Like "####-##-##") Then
112         ErrInfo = "�������ڸ�ʽ��ΪYYYY-MM-DD"
            Exit Function
        End If
114     strEndDate = strEndDate & " 23:59:59"
        
        
116     If lngType <= 0 Then
118         gstrSql = "Select Distinct A.���id As ID, D.����, D.�Ա�, D.����, A.������Դ, D.�����, D.סԺ��, E.���� As ������ұ���, E.���� As �����������, F.��� As ҽ�����,A.����ҽ��, A.����ʱ��,D.����,C.��������,C.�����ӡ,G.��Ժ���� as ���� " & vbNewLine & _
                        "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C, ������Ϣ D, ���ű� E,��Ա�� F,������ҳ G" & vbNewLine & _
                        "Where A.������Ŀid = B.ID And B.��� = 'C' And A.ID = C.ҽ��id And A.���id Is Not Null And C.ִ��״̬ = 0 " & vbNewLine & _
                        "      And A.������Դ=2 And A.����id = D.����id And A.��������id = E.ID And A.ҽ��״̬ = 8 And A.����ҽ��=F.����(+) And A.����ID=G.����ID and A.��ҳID=G.��ҳID " & _
                        "      And d.��Ժ = 1 And E.����  = [1] And A.����ʱ�� between [2] And [3]  Order By A.����ʱ��"
120         Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ���뵥", strDeptNo, CDate(strStartDate), CDate(strEndDate))
122         Do Until rsTmp.EOF
124             strData = rsTmp("ID") & "^" & rsTmp("����") & "^" & rsTmp("�Ա�") & "^" & rsTmp("����") & "^" & rsTmp("������Դ") & _
                          "^" & rsTmp("�����") & "^" & rsTmp("סԺ��") & "^" & rsTmp("������ұ���") & "^" & rsTmp("�����������") & _
                          "^" & rsTmp("����ҽ��") & "^" & rsTmp("����ʱ��") & "^" & rsTmp("����") & "^" & rsTmp("��������") & "^" & rsTmp("�����ӡ") & _
                          "^" & rsTmp!ҽ����� & "^" & rsTmp!����
126             If mstrR(UBound(mstrR)) <> "" Then ReDim Preserve mstrR(UBound(mstrR) + 1)
128             mstrR(UBound(mstrR)) = strData
130             rsTmp.MoveNext
            Loop
        End If
        
132     ReDim strR(0) As String
134     lngCount = 0
136     For lngI = LBound(mstrR) To UBound(mstrR)
138         If lngCount < 100 Then
140             If mstrR(lngI) <> "" Then strReturn = strReturn & "|" & mstrR(lngI)
            Else
142             If strR(UBound(strR)) <> "" Then ReDim Preserve strR(UBound(strR) + 1)
144             strR(UBound(strR)) = mstrR(lngI)
            End If
146         lngCount = lngCount + 1
        Next
148     mstrR = strR
 
150     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
152     GetDeptPatiList = True
        Exit Function
errH:
154     ErrInfo = CStr(Erl()) & "�г��ִ���," & Err.Description & "(" & Err.Number & ")"
156     WriteLog CStr(Erl()) & "�г��ִ���," & Err.Number & " " & Err.Description
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
112     gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "���뱨��"
    
114     gstrSql = "Select Nvl(A.����id, 0) As �ļ�id From ����ҽ������ A Where A.ҽ��id = [1] "
116     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "���뱨��", lngID)
118     If rsTmp.EOF = True Then Exit Function

120     strZipFile = zlFileZip(strPath)
        
        
'122     strUnZipFile = zlFileUnzip(strZipFile)
    
    
124     If zlLisBlobSql(rsTmp("�ļ�ID"), strZipFile, aStrSQL) = False Then Exit Function
    
126     For intLoop = 0 To UBound(aStrSQL)
128         gobjComLib.zlDataBase.ExecuteProcedure Replace(aStrSQL(intLoop), "Call", ""), "���뱨��"
    '        Debug.Print aStrSQL(intLoop)
        Next
'130     gobjFSO.DeleteFile strZipFile
'132     gobjFSO.DeleteFile strPath
134     InsertReport = True
        Exit Function
errH:
136     ErrInfo = CStr(Erl()) & "," & Err.Number & " " & Err.Description
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
100     Err = 0: On Error Resume Next
102     lngLBound = LBound(arySql): lngUBound = UBound(arySql)
104     If Err <> 0 Then lngLBound = 0: lngUBound = -1
106     Err = 0: On Error GoTo 0
    
108     lngFileNum = FreeFile
110     Open strFile For Binary Access Read As lngFileNum
112     lngFileSize = LOF(lngFileNum)
    
114     Err = 0: On Error GoTo errHand
116     conChunkSize = 500
118     lngModSize = lngFileSize Mod conChunkSize
120     lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
122     ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
124     For lngCount = 0 To lngBlocks
126         If lngCount = lngFileSize \ conChunkSize Then
128             lngCurSize = lngModSize
            Else
130             lngCurSize = conChunkSize
            End If
        
132         ReDim aryChunk(lngCurSize - 1) As Byte
134         ReDim aryHex(lngCurSize - 1) As String
136         Get lngFileNum, , aryChunk()
138         For lngBound = LBound(aryChunk) To UBound(aryChunk)
140             aryHex(lngBound) = Hex(aryChunk(lngBound))
142             If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
            Next
144         strText = Join(aryHex, "")
146         If strText <> "" Then
    '            If lngCount = 0 Then strText = "100;" & strText
148             arySql(lngUBound + lngCount + 1) = "Zl_���Ӳ�����ʽ_Insert(" & KeyWord & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
            End If
        Next
150     Close lngFileNum
152     zlLisBlobSql = True
        Exit Function

errHand:
154     Close lngFileNum
156     zlLisBlobSql = False
158     WriteLog "zlLisBlobSql " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
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
102     gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "ɾ������"
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
102     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "�õ�������Դ", lngAdivce)
104     Do Until rsTmp.EOF
106         str������Դ = Trim("" & rsTmp!������Դ)
108         rsTmp.MoveNext
        Loop
110     If str������Դ = "4" Then
            '��첡��
112         gstrSql = "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
                    "       Sum(E.��׼����) As ��׼����, Decode(E.��¼״̬,1,Decode(E.ִ��״̬,9,0,E.��¼״̬),E.��¼״̬) as ��¼״̬, '0' As �Ƿ�ɼ�" & vbNewLine & _
                    "From ������ü�¼ E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "Where D.��¼���� = E.��¼����(+) And D.No = E.No(+) And D.ҽ��id = E.ҽ�����(+) And E.��¼״̬(+) <> 2 And A.������� = 'C' And" & vbNewLine & _
                    "      A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.���id = [1]" & vbNewLine & _
                    "Group By A.������Ŀid, B.����, B.����, C.����, C.����, E.��¼״̬" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select A.������Ŀid As ID, B.���� As ������Ŀ����, B.���� As ������Ŀ����, C.���� As ִ�п��ұ���, C.���� As ִ�п�������, Sum(E.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
                    "       Sum(E.��׼����) As ��׼����, Decode(E.��¼״̬,1,Decode(E.ִ��״̬,9,0,E.��¼״̬),E.��¼״̬) As ��¼״̬, '1' As �Ƿ�ɼ�" & vbNewLine & _
                    "From ������ü�¼ E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "Where D.��¼���� = E.��¼����(+) And D.No = E.No(+) And D.ҽ��id = E.ҽ�����(+) And E.��¼״̬(+) <> 2 And A.������� = 'E' And" & vbNewLine & _
                    "      A.Id = D.ҽ��id And A.������Ŀid = B.Id And A.ִ�п���id = C.Id And A.Id = [1]" & vbNewLine & _
                    "Group By A.������Ŀid, B.����, B.����, C.����, C.����, E.��¼״̬"


114     ElseIf str������Դ = 2 Then
116         gstrSql = "Select a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.��׼����) As ��׼����, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬," & vbNewLine & _
                    "       a.������־, a.�걾��λ" & vbNewLine & _
                    "From (Select Distinct a.������Ŀid As ID, b.���� As ������Ŀ����, b.���� As ������Ŀ����, c.���� As ִ�п��ұ���, c.���� As ִ�п�������,e.�շ�ϸĿid, e.���� * e.���� As ʵ�ս��," & vbNewLine & _
                    "                       e.���� As ��׼����, d.�Ʒ�״̬, '0' As �Ƿ�ɼ�, f.��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "       From סԺ���ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "       Where d.ҽ��id = f.ҽ�����(+) And d.No = f.No(+) And d.��¼���� = f.��¼����(+) And f.��¼״̬(+) <> 2 And a.Id = e.ҽ��id And" & vbNewLine & _
                    "             a.������� = 'C' And a.Id = d.ҽ��id And a.������Ŀid = b.Id And a.ִ�п���id = c.Id And a.���id = [1]) A" & vbNewLine & _
                    "Group By a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.��׼����) As ��׼����, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬," & vbNewLine & _
                    "       a.������־, a.�걾��λ" & vbNewLine & _
                    "From (Select a.������Ŀid As ID, b.���� As ������Ŀ����, b.���� As ������Ŀ����, c.���� As ִ�п��ұ���, c.���� As ִ�п�������,e.�շ�ϸĿid, e.���� * e.���� As ʵ�ս��," & vbNewLine & _
                    "              e.���� As ��׼����, d.�Ʒ�״̬, '1' As �Ƿ�ɼ�, f.��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "       From סԺ���ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "       Where d.ҽ��id = f.ҽ�����(+) And d.No = f.No(+) And d.��¼���� = f.��¼����(+) And f.��¼״̬(+) <> 2 And a.Id = e.ҽ��id And" & vbNewLine & _
                    "             a.������� = 'E' And a.Id = d.ҽ��id And a.������Ŀid = b.Id And a.ִ�п���id = c.Id And a.Id = [1]) A" & vbNewLine & _
                    "Group By a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬, a.������־, a.�걾��λ"
        Else
118         gstrSql = "Select a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.��׼����) As ��׼����, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬," & vbNewLine & _
                    "       a.������־, a.�걾��λ" & vbNewLine & _
                    "From (Select Distinct a.������Ŀid As ID, b.���� As ������Ŀ����, b.���� As ������Ŀ����, c.���� As ִ�п��ұ���, c.���� As ִ�п�������,e.�շ�ϸĿid, e.���� * e.���� As ʵ�ս��," & vbNewLine & _
                    "                       e.���� As ��׼����, d.�Ʒ�״̬, '0' As �Ƿ�ɼ�, Decode(f.��¼״̬,1,Decode(f.ִ��״̬,9,0,f.��¼״̬),f.��¼״̬) As ��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "       From ������ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "       Where d.ҽ��id = f.ҽ�����(+) And d.No = f.No(+) And d.��¼���� = f.��¼����(+) And f.��¼״̬(+) <> 2 And a.Id = e.ҽ��id And" & vbNewLine & _
                    "             a.������� = 'C' And a.Id = d.ҽ��id And a.������Ŀid = b.Id And a.ִ�п���id = c.Id And a.���id = [1]) A" & vbNewLine & _
                    "Group By a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.��׼����) As ��׼����, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬," & vbNewLine & _
                    "       a.������־, a.�걾��λ" & vbNewLine & _
                    "From (Select a.������Ŀid As ID, b.���� As ������Ŀ����, b.���� As ������Ŀ����, c.���� As ִ�п��ұ���, c.���� As ִ�п�������, e.�շ�ϸĿid, e.���� * e.���� As ʵ�ս��," & vbNewLine & _
                    "              e.���� As ��׼����, d.�Ʒ�״̬, '1' As �Ƿ�ɼ�,  Decode(f.��¼״̬,1,Decode(f.ִ��״̬,9,0,f.��¼״̬),f.��¼״̬) As ��¼״̬, a.������־, a.�걾��λ" & vbNewLine & _
                    "       From ������ü�¼ F, ����ҽ���Ƽ� E, ����ҽ������ D, ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C" & vbNewLine & _
                    "       Where d.ҽ��id = f.ҽ�����(+) And d.No = f.No(+) And d.��¼���� = f.��¼����(+) And f.��¼״̬(+) <> 2 And a.Id = e.ҽ��id And" & vbNewLine & _
                    "             a.������� = 'E' And a.Id = d.ҽ��id And a.������Ŀid = b.Id And a.ִ�п���id = c.Id And a.Id = [1]) A" & vbNewLine & _
                    "Group By a.Id, a.������Ŀ����, a.������Ŀ����, a.ִ�п��ұ���, a.ִ�п�������, a.�Ʒ�״̬, a.�Ƿ�ɼ�, a.��¼״̬, a.������־, a.�걾��λ"

        End If
120     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "�õ�������Ŀ", lngAdivce)
    
122     Do Until rsTmp.EOF
124         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("������Ŀ����") & "^" & rsTmp("������Ŀ����") & "^" & rsTmp("ִ�п��ұ���") & _
                        "^" & rsTmp("ִ�п�������") & "^" & rsTmp("��׼����") & "^" & rsTmp("ʵ�ս��") & "^" & rsTmp("��¼״̬") & "^" & rsTmp("�Ƿ�ɼ�") & _
                        "^" & IIf(Val("" & rsTmp("������־")) = 1, "1", "0") & "^" & IIf(Trim("" & rsTmp("�걾��λ")) = "", "ѪҺ", Trim("" & rsTmp("�걾��λ")))
126         rsTmp.MoveNext
        Loop

128     If strData <> "" Then
130         GetClinicItem = Mid(strData, 2)
        End If
    
        Exit Function
errH:
132     WriteLog "GetClinicItem " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
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

102     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ��ָ����ϸ", lngClinicID)
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
102     gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "���ջ�ȡ������"
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
        On Error GoTo errH
    
102     gstrSql = "select ID,����,����,�����Ŀ from ������ĿĿ¼  where ��� = 'C' "
104     If strFindItem <> "" Then
106         gstrSql = gstrSql & " And (���� = [1] or ���� like '%[1]%') "
        End If
    
108     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ������Ŀ", CStr(strFindItem))
    
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
        If InStr(mstrVirtualPeis, ";����;") <= 0 Then
            UpdateTestResults = "0|ȱ�����ϵͳ��ģ��Ȩ�ޣ����ڹ������н�2138������LIS�ӿ�ģ���Ȩ������" & gstrDbUser
            Exit Function
        End If
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
            
114     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "��ȡ����ָ��", lngID)
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
170         If strEditSQL(i) <> "" Then gobjComLib.zlDataBase.ExecuteProcedure strEditSQL(i), "�������ָ��"
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
        WriteLog "��Ȩ�汾" & mblnOldVer & ",Ȩ��" & mstrVirtualHis
100     If mblnOldVer Then
            
102         If InStr(mstrVirtualHis, ";ZLLIS�걾����;") <= 0 Then
104             strErrInfo = "δ����걾����Ȩ�ޣ����ܵ��ã�"
                Exit Function
            End If
        Else
             
106         If InStr(gstrPrivs, ";���ձ걾;") <= 0 Then
108             strErrInfo = "δ������鼼ʦ����վģ��ġ����ձ걾��Ȩ�ޣ����ܵ��ã�"
                Exit Function
            End If
        End If
        '������������
110     strSQL = "Select ʹ��С��id From �������� Where ID = [1]"
112     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "��������걾", lngDeviceID)
114     lngDeptID = 0
116     If Not rsTmp.EOF Then
118         lngDeptID = Val("" & rsTmp("ʹ��С��id"))
        End If
120     If lngDeptID <= 0 Then
            '�˳�-������ʾ
122         strErrInfo = "��������δָ����Ӧ�ļ���С�飡"
            Exit Function
        End If
124     dtSampleDate = CDate(Format(gobjComLib.zlDataBase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    
126     strSQL = "Select ID, ����, �Ա�, ����, NO, ��Ŀid, ���, ��־, ����ο�, ����, ����ʱ��, ������, Rownum As �������, ������Ŀid," & vbNewLine & _
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

128     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, lngDeptID, lngID)
130     If rsTmp.EOF Then
132         strErrInfo = "û���ҵ��������룡"
            Exit Function
        End If
    
134     If Val(strSampleNo) <= 0 Then
136         strErrInfo = "�걾�Ŵ�����ֻ֧�ִ���������֣�"
            Exit Function
        Else
138         strSampleNo = Val(strSampleNo)
        End If
140     dStart = CDate(Format(dtSampleDate, "yyyy-MM-dd 00:00:00"))
142     dEnd = CDate(Format(dtSampleDate, "yyyy-MM-dd 23:59:59"))
144     strSQL = "Select ������,����ʱ�� from ����걾��¼ where ����ʱ�� Between [3] and [4] and ����ID=[1] and �걾���=[2] "
146     Set rsItem = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, strSampleNo, dStart, dEnd)
148     If Not rsItem.EOF Then
150         strErrInfo = strSampleNo & "�ű걾�Ѵ��ڣ�" & vbNewLine & "�����ˣ�" & rsItem!������ & " ����ʱ��:" & Format(rsItem!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            Exit Function
        End If
152     gstrSql = "Select B.����id, B.��ҳid, B.���, B.Ӥ������, B.Ӥ���Ա�" & vbNewLine & _
                        "From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                        "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.Ӥ�� = B.��� And A.���id = [1] And Rownum = 1"
154     Set rs = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "zlLISRegister", lngID)
156     If rs.EOF = False Then
158         str���� = Trim("" & rs("Ӥ������"))
160         str�Ա� = Trim("" & rs("Ӥ���Ա�"))
162         str���� = "Ӥ��"
        Else
164         str���� = Trim("" & rsTmp("����"))
166         str�Ա� = Trim("" & rsTmp("�Ա�"))
168         str���� = Trim("" & rsTmp("����"))
        End If
    
        '����������Ŀ
170     gstrSql = "select distinct ҽ������ from ����ҽ����¼ a , ����ҽ������ b, ���鱨����Ŀ c , ����������Ŀ d " & vbNewLine & _
                  "  where a.id = b.ҽ��ID and a.������ĿID = c.������ĿID and " & vbNewLine & _
                  "  c.������ĿID = d.��ĿID(+) and a.���id=[1] "
172     Set rsItem = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��ȡ��������", lngID)
174     Do Until rsItem.EOF
176         strItem = strItem & " " & Trim("" & rsItem("ҽ������"))
178         rsItem.MoveNext
        Loop
180     strItem = Trim(strItem) & "(" & Trim("" & rsTmp("�걾��λ")) & ")"
        
        '�����걾��¼
        '------------10.25
182     lngKey = gobjComLib.zlDataBase.GetNextId("����걾��¼")
    '    gstrSql = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
    '        rsTmp("ID") & ",'" & _
    '        strSampleNo & "'," & _
    '        IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
    '        IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
    '        lngDeviceID & "," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
    '        "1,'" & _
    '        gobjComLib.zlDatabase.GetUserInfo.Fields("����").value & "'," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0,0,0," & _
    '        rsTmp("����") & ",NULL,'" & _
    '        str���� & "','" & str�Ա� & "','" & str���� & "','" & Trim("" & rsTmp("No")) & "','" & _
    '        Trim("" & rsTmp("�걾��λ")) & "'," & Trim("" & rsTmp("��������ID")) & ",'" & Trim("" & rsTmp("����ҽ��")) & "'," & _
    '        Trim("" & rsTmp("��ʶ��")) & ",'" & Trim("" & rsTmp("��ǰ����")) & "','" & Trim("" & rsTmp("���˿���")) & "','" & _
    '        strItem & "')"
    
        '---------- 10.26 ��SQL
    
184     gstrSql = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
            rsTmp("ID") & ",'" & rsTmp("ID") & "',0,'" & _
            strSampleNo & "'," & _
            IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
            IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
            lngDeviceID & "," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
            "'" & _
            gstrUserName & "'," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0," & _
            rsTmp("����") & ",NULL,'" & _
            str���� & "','" & str�Ա� & "','" & str���� & "','" & Trim("" & rsTmp("No")) & "','" & _
            Trim("" & rsTmp("�걾��λ")) & "'," & Trim("" & rsTmp("��������ID")) & ",'" & Trim("" & rsTmp("����ҽ��")) & "','" & _
            Trim("" & rsTmp("��ʶ��")) & "','" & Trim("" & rsTmp("��ǰ����")) & "','" & Trim("" & rsTmp("���˿���")) & "','" & _
            strItem & "',Null,Null,Null,'" & gstrUserCode & "','" & gstrUserName & "')"
    
        '-------------------------------------------------------------------------------------
    
186     gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "��������걾"
                                                                
        '��дָ��
188     strItemRecords = ""
190     Do While Not rsTmp.EOF
192         strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("��ĿID") & "^" & _
                Trim("" & rsTmp("���")) & "^" & Val("" & rsTmp("��־")) & "^" & Trim("" & rsTmp("����ο�")) & "^" & _
                Trim("" & rsTmp("������ĿID")) & "^" & Trim("" & rsTmp("�������"))
            
194         rsTmp.MoveNext
        Loop
    
196     If Len(strItemRecords) > 0 Then
198         strItemRecords = Mid(strItemRecords, 2)
            
200         gstrSql = "Zl_������ͨ���_Write(" & lngKey & "," & _
                lngDeviceID & ",'" & strItemRecords & "',0,0)"
202         gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "���ɱ걾"
        End If
    
204     zlLISRegister = True
        Exit Function
errH:
206     strErrInfo = CStr(Erl()) & "," & Err.Number & " " & Err.Description
208     WriteLog "zlLISRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnRegister(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
        'ȡ����ZLLIS���Ѻ��յı걾
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     If mblnOldVer Then
102         If InStr(mstrVirtualHis, ";ZLLISȡ������;") <= 0 Then
104             strErrInfo = "δ����ȡ������Ȩ�ޣ����ܵ��ã�"
                Exit Function
            End If
        Else
106         If InStr(gstrPrivs, ";���ճ���;") <= 0 Then
108             strErrInfo = "δ������鼼ʦ����վģ��ġ����ճ�����Ȩ�ޣ����ܵ��ã�"
                Exit Function
            End If
        End If
        '�Ƿ��ȡ�����յĲ����ڴ洢�����У����Դ˴��������
110     strSQL = "Zl_����걾��¼_ȡ������(" & lngID & ")"
112     gobjComLib.zlDataBase.ExecuteProcedure strSQL, "ȡ������"
114     zlLisUnRegister = True
        Exit Function
errH:
116     strErrInfo = CStr(Erl()) & "," & Err.Description
118     WriteLog "zlLisUnRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function ZLLisInsterReport(ByVal lngID As Long, strItems As String, ByRef strErrInfo As String) As Boolean
        Dim str�걾 As String, lng����ID As Long, str�Ա� As String, str����  As String
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim str��Ŀ As String, varItem As Variant
        On Error GoTo errH
        
        If mblnOldVer Then
100         If InStr(mstrVirtualHis, "ZLLIS�걾���") <= 0 Then
102             strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
                Exit Function
            End If
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

118     Set rsSample = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "��ȡ��������", lngID)
    
120     If rsSample.EOF Then
122         strErrInfo = "δ�ҵ���Ӧҽ����"
            Exit Function
        End If
    
124     If Trim("" & rsSample!�����) <> "" Then
126         strErrInfo = "����˱걾�������޸ģ�"
            Exit Function
        End If
    
128     If InStr(1, gstrPrivs, ";��˱걾;") <= 0 Then
130         strErrInfo = "û�м��鼼ʦ����վ�����Ȩ��!"
            Exit Function
        End If
    
        '11210 Ȩ�ޡ�δ�շ���ˡ�������˵�������ʱ��δ��Ч��
132     If InStr(gstrPrivs, ";δ�շ����;") <= 0 Then
134         strErrInfo = CheckChargeState(lngID, False)
136         If strErrInfo <> "" Then Exit Function
        End If
    
        '21137 �ѹ鵵���治�����
138     gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
140     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "��˼��", lngID)
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
    
166             Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ��ָ��ID", strCode)
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
194     strSQL = "ZL_������ͨ���_BATCHUPDATE(" & rsSample!id & "," & _
                        rsSample!����ID & ",'" & Trim("" & rsSample!�걾����) & "'," & str�Ա� & "," & _
                        IIf(Trim("" & rsSample!��������) = "", "Null", "To_Date('" & Trim("" & rsSample!��������) & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                        str��Ŀ & "'," & rsSample!΢����걾 & ")"
196     gobjComLib.zlDataBase.ExecuteProcedure strSQL, "��д���"

        '���
198     strSQL = "ZL_����걾��¼_�������(" & rsSample!id & ",'" & gstrUserName & "')"
200     gobjComLib.zlDataBase.ExecuteProcedure strSQL, "��˱���"
202     ZLLisInsterReport = True
        Exit Function
errH:
204     strErrInfo = CStr(Erl()) & "," & Err.Description
206     WriteLog "ZLLisInsterReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnAudit(ByVal lngID As Long, strErrInfo As String) As Boolean
        'ȡ�����
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim d���ʱ�� As Date, dCurr As Date
        On Error GoTo errH
        If mblnOldVer Then
100         If InStr(mstrVirtualHis, "ZLLISȡ�����") <= 0 Then
102             strErrInfo = "�˽ӿ�δ��Ȩ�����ܵ��ã�"
                Exit Function
            End If
        End If
        
104     strSQL = "Select a.ID,a.��ӡ����, a.���ʱ�� From ����걾��¼ A Where ҽ��ID=[1]"
106     Set rsSample = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "ȡ����˼��", lngID)
    
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
120         dCurr = gobjComLib.zlDataBase.Currentdate
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
134     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ�����", lngID)
136     If rsTmp.EOF = False Then
138         strErrInfo = "���˱���סԺ�Ĳ������ύ��飬����ȡ����ˣ�"
            Exit Function
        End If
    
140     strSQL = "ZL_����걾��¼_���ȡ��(" & rsSample!id & ")"
142     gobjComLib.zlDataBase.ExecuteProcedure strSQL, "ȡ�����"
144     zlLisUnAudit = True
        Exit Function
errH:
146     strErrInfo = CStr(Erl()) & "," & Err.Description
148     WriteLog "zlLisUnAudit " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetAllDevice(ByRef strErrInfo As String) As String
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     strSQL = "Select ID,����,���� From ��������"
102     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "ȡ��������")
104     GetAllDevice = ""
106     Do Until rsTmp.EOF
108         GetAllDevice = GetAllDevice & "|" & rsTmp!id & "^" & rsTmp!���� & "^" & rsTmp!����
110         rsTmp.MoveNext
        Loop
112     If GetAllDevice <> "" Then GetAllDevice = Mid(GetAllDevice, 2)
114     If GetAllDevice = "" Then
116         strErrInfo = "û�г�ʼ��������"
        End If
        Exit Function
errH:
118     strErrInfo = CStr(Erl()) & "," & Err.Description
120     WriteLog "GetAllDevice " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function CriticalvalueNotice(ByVal lngID As Long, ByVal strNoticeTitle As String, ByVal strNotice As String) As Long
    'Σ��ֵ֪ͨ
    Dim rsTmp As ADODB.Recordset
    Dim lngPatiID As Long, lngPageID As Long, lngNoticeID As Long
    Dim strSaveTitle As String, strSaveNotice As String
    gstrSql = "Select ����id,��ҳid,������Դ From ����ҽ����¼ Where ID=[1]"
    Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "�õ�������Դ", lngID)
    
    Do Until rsTmp.EOF
        'str��Դ = Trim("" & rsTmp!������Դ)
        lngPatiID = Val("" & rsTmp!����id)
        lngPageID = Val("" & rsTmp!��ҳid)
        rsTmp.MoveNext
    Loop
    
    If lngPatiID <= 0 Then
        WriteLog "CriticalvalueNotice :ID" & lngID & "�޶�Ӧ��ҽ����¼��"
        Exit Function
    End If
    
    strSaveTitle = Replace(strNoticeTitle, "'", "")
    strSaveTitle = Replace(strSaveTitle, """", "")
    
    strSaveNotice = Replace(strNotice, "'", "")
    strSaveNotice = Replace(strSaveNotice, """", "")
    
    If strSaveTitle <> "" And strSaveNotice <> "" Then
        lngNoticeID = gobjComLib.zlDataBase.GetNextId("�ٴ�ҵ������")
        
        gstrSql = "Zl_�ٴ�ҵ������_Edit(1," & lngNoticeID & "," & lngPatiID & "," & IIf(lngPageID <> 0, lngPageID, "Null") & ",3,301," & _
                  lngID & ",'" & strSaveTitle & "','" & strSaveNotice & "','" & gstrUserName & "(" & gstrUserCode & ")')"
        gobjComLib.zlDataBase.ExecuteProcedure gstrSql, "����Σ��ֵ����"
        CriticalvalueNotice = lngNoticeID
        
    Else
        WriteLog "CriticalvalueNotice : ��������ݲ���Ϊ�գ�"
    End If
     
    Exit Function
errH:
     WriteLog "CriticalvalueNotice :" & CStr(Erl()) & "��," & Err.Description
End Function

Public Function Incomeverify(ByVal lngAdivce As Long, ByRef strErrInfo) As Boolean
    'סԺ���۵����
    Dim str��Դ As String, lngPatiID As Long, lngPageID As Long, curTotal As Currency, curOver As Currency
    Dim rsTmp As ADODB.Recordset, rsFee As ADODB.Recordset, curVerifyTotal As Currency
    Dim strExtSQL() As String, i As Integer, blnTrans As Boolean, strNos As String
    On Error GoTo errH
100 Incomeverify = False
102 gstrSql = "Select ����id,��ҳid,������Դ From ����ҽ����¼ Where ID=[1]"
104 Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "�õ�������Դ", lngAdivce)
106 Do Until rsTmp.EOF
108     str��Դ = Trim("" & rsTmp!������Դ)
110     lngPatiID = Val("" & rsTmp!����id)
112     lngPageID = Val("" & rsTmp!��ҳid)
114     rsTmp.MoveNext
    Loop
        
116 If str��Դ <> "2" Then
118     strErrInfo = "����סԺ���ˣ����ܽ��к���������"
        Exit Function
    End If
120 gstrSql = "select ��Ժ���� from ������ҳ where ����id=[1] and ��ҳid=[2]"
122 Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ��Ժ����", lngPatiID, lngPageID)
124 Do Until rsTmp.EOF
126     If Trim$("" & rsTmp!��Ժ����) <> "" Then
128         strErrInfo = "סԺ�����ѳ�Ժ�����ܽ��к���������"
            Exit Function
        End If
130     rsTmp.MoveNext
    Loop
    

    
132 curTotal = 0: curVerifyTotal = 0
134 strNos = ""
136 ReDim strExtSQL(0) As String
138 gstrSql = "select distinct a.��¼����,a.NO from ����ҽ������ a,����ҽ����¼ b where (a.ҽ��id=b.id or a.ҽ��id=[1] ) and b.���id=[1]"
140 Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ���ݺ�", lngAdivce)
142 Do Until rsTmp.EOF
144     gstrSql = "select sum(Nvl(a.ʵ�ս��,0)) as ���  From סԺ���ü�¼ a where a.��¼����=[2] and a.No=[1] and a.��¼״̬=0 "
146     Set rsFee = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ���ݷ���", CStr(Trim("" & rsTmp!no)), Val("" & rsTmp!��¼����))
148     If Val("" & rsFee!���) <> 0 Then
150         curTotal = curTotal + Val("" & rsFee!���)
152         If strExtSQL(UBound(strExtSQL)) <> "" Then ReDim Preserve strExtSQL(UBound(strExtSQL) + 1)
154         strExtSQL(UBound(strExtSQL)) = "Zl_סԺ���ʼ�¼_Verify('" & CStr(Trim("" & rsTmp!no)) & "','" & gstrUserCode & "','" & gstrUserName & "',''," & lngPatiID & ")"
156         strNos = strNos & "," & rsTmp!no
    '        Else
    '            gstrSql = "select sum(Nvl(a.ʵ�ս��,0)) as ���  From סԺ���ü�¼ a where a.��¼����=[2] and a.No=[1] and a.��¼״̬<>0 And a.��¼״̬<>9 "
    '            Set rsFee = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ���ݷ���", CStr(Trim("" & rsTmp!no)), Val("" & rsTmp!��¼����))
    '            If Val("" & rsFee!���) <> 0 Then curVerifyTotal = curVerifyTotal + Val("" & rsFee!���)
        End If
158     rsTmp.MoveNext
    Loop
160 If curTotal <= 0 Then
162     strErrInfo = "û����Ҫ���ʵĻ��۵���"
        Exit Function
    End If
    
    If Not mblnOldVer Then
        If InStr(mstrVirtualHis, ";���ʼ�����;") > 0 Then
            mblnVerifyTotal = True
        Else
            mblnVerifyTotal = False
        End If
    End If
    
    If mblnVerifyTotal Then '2012-05-14 ����ҽԺ Ҫ��ȥ������� ,2012-05-23 ��Ϊ��ȡ�����ļ��е�����
164     curOver = 0
        '�������   ����0-�ڳ���1-��ĩ������1=���2=סԺ
166     gstrSql = "select Nvl(Ԥ�����,0)-Nvl(�������,0) as ��� from ������� where ����id=[1] and ����=1 and ����=2 "
168     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "ȡ�������", lngPatiID)
170     Do Until rsTmp.EOF
172         curOver = Val("" & rsTmp!���)
174         rsTmp.MoveNext
        Loop
    
176     If curOver < curTotal Then
178         strErrInfo = "סԺ�������㣬���ܽ��к���������"
            Exit Function
        End If
    End If
180 blnTrans = False
182 gcnOracle.BeginTrans
184 blnTrans = True
186 For i = LBound(strExtSQL) To UBound(strExtSQL)
188     If strExtSQL(i) <> "" Then gobjComLib.zlDataBase.ExecuteProcedure strExtSQL(i), "���۵�ת���ʵ�"
    Next
190 gcnOracle.CommitTrans
192 If strNos <> "" Then strErrInfo = Mid(strNos, 2)
194 Incomeverify = True
    Exit Function
errH:
196 If blnTrans Then gcnOracle.RollbackTrans
198  strErrInfo = CStr(Erl()) & "��," & Err.Description
200  WriteLog "Incomeverify :" & strErrInfo
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
    
114     Set rsTemp = gobjComLib.zlDataBase.GetUserInfo
    
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
134     WriteLog "GetUserInfo " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
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
104     Set rs = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "��������", lngKey)
106     If rs.EOF = True Then Exit Function
108     intPatientType = rs("������Դ")
    
110     If blnOrder Then
112         strSQL = _
                "select DeCode(NVL(A.��¼״̬,0),1,Decode(A.ִ��״̬,9,0,A.��¼״̬),A.��¼״̬) As ��¼״̬ " & _
                      "from סԺ���ü�¼ A, " & _
                      "( " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id))  " & _
                           "Union " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id)) " & _
                      ") B " & _
                    "Where A.NO = B.NO "
114         If intPatientType <> 2 Then
116             strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            End If
        Else
118         strSQL = _
                "select DeCode(NVL(A.��¼״̬,0),1,Decode(A.ִ��״̬,9,0,A.��¼״̬),A.��¼״̬) As ��¼״̬ " & _
                      "from סԺ���ü�¼ A, " & _
                      "( " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                           "Union " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                      ") B " & _
                    "Where A.NO = B.NO and a.��¼���� = b.��¼���� "
120         If intPatientType <> 2 Then
122             strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            End If
        End If
    
124     strSQL = strSQL & " Order by ��¼״̬ "
    
126     Set rs = gobjComLib.zlDataBase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

128     If rs.BOF Then Exit Function
130     If rs("��¼״̬").value = 0 Then Exit Function
    
132     CheckChargeState = ""
        Exit Function
errH:
134     CheckChargeState = CStr(Erl()) & "," & Err.Description
136     WriteLog "CheckChargeState " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
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
102     If gobjComLib.zlDataBase.GetPara(81, 100) <> 1 Then Exit Function
        
        '��ǰ�����Ƿ��ѳ�Ժ��Ԥ��Ժ
104     gstrSql = "select d.no" & vbNewLine & _
                "from (select distinct d.ҽ��id" & vbNewLine & _
                "       from ����걾��¼ a, ������Ϣ b, ������ҳ c, ������Ŀ�ֲ� d" & vbNewLine & _
                "       where a.����id = b.����id and a.����id = c.����id and a.��ҳid = c.��ҳid and" & vbNewLine & _
                "             a.id = [1] and a.������Դ = 2 and (b.��Ժʱ�� is not null or c.״̬ = 3) and" & vbNewLine & _
                "             a.id = d.�걾id) a, ����ҽ����¼ b, ����ҽ������ c, סԺ���ü�¼ d" & vbNewLine & _
                "where a.ҽ��id in (b.���id, b.id) and b.id = c.ҽ��id and c.��¼���� = d.��¼���� and" & vbNewLine & _
                "      c.no = d.no and d.��¼���� = 2 and d.��¼״̬ = 0 "
106     Set rsTmp = gobjComLib.zlDataBase.OpenSQLRecord(gstrSql, "���鼼ʦ����վ-����״̬���", lngKey)
    
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
