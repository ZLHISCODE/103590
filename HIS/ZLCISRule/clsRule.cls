VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitCISRule(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngProgram As Long) As Boolean
          Dim strFilePath As String

1         On Error GoTo ErrH

2         gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
3         gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
4         glngSys = lngSys
5         Set gcnOracle = cnMain
6         Call GetUserInfo
          '�������û����־�ļ��½���־�ļ�
7         strFilePath = App.Path & "\" & App.EXEName & ".ini"
8         If Dir(strFilePath) = "" Then
              '[DEBUG]
              'OutPutDebugInfo = 0 ;���������Ϣ��zlCISRule.log�ļ���0-�����,1-������нӿڲ�����Ϣ
9             WriteIni "DEBUG", "OutPutDebugInfo", "0 ;���������Ϣ��" & App.EXEName & ".log�ļ�,0-�����,1-���", strFilePath
               
10        End If
          
11        Call InitObjLis(lngProgram)
          
12        gblnLog = (Val(ReadIni("Debug", "OutPutDebugInfo", strFilePath)) = 1)
          
13        InitCISRule = True

14        Exit Function

ErrH:
15        MsgBox "��zlCISRule.clsRule.InitCISRule�ĵ�" & Erl() & "�г�����" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function SubmitMain(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
    ByVal objMain As Object, ByVal bytScene As Byte, Optional colDiag As Collection) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : SubmitMain
      ' Author    : YWJ
      ' Date      : 2019-09-20 17:16:31
      ' Parameter :
      '             lngPatiID -����ID
      '             lngVisitId -��ҳID
      '             strRegNO -�Һŵ���
      '             objMain-����\סԺҽ���´ﴫ��ҽ����¼��;����\סԺҽ������վ�´����뵥 �������弯��
      '             bytScene-���� 1-����\סԺҽ���´�;
      '                           21:�������,22-��������,23-��Ѫ����,24-��������;25-��������;
      '                           3-סԺҽ����ҳ���\���ﲡ����Ϣ¼�����
      '                           4-�걾�ɼ�
      '             colDiag-����༭�´ﴫ��
      ' Purpose   : �ύ������Ϣ
      ' Return    : T-��ֹ����;F-��������
      '---------------------------------------------------------------------------------------
          Dim strJsonIn As String
          Dim strJsonOut As String
          
          Dim strPati As String
          Dim strBusinss As String
          
          Dim strMainInfo As String

          Dim lngRegId As Long    '�Һ�ID
          Dim bytType As Byte
              
          Dim blnRet As Boolean
          
          Dim objJson As Object
          
          Dim colList As Collection

          '--��־
1         On Error GoTo ErrH
2         WriteLog "������SubmitMain ��ʼ" & vbNewLine & _
                   "��Σ�����ID=" & lngPatiID & ",��ҳID=" & lngVisitId & ",�Һŵ�=" & strRegNo & _
                   ",��������=" & TypeName(objMain) & ",����=" & bytScene & ",���=" & (Not colDiag Is Nothing) & vbNewLine
3         bytType = bytScene
4         If InStr(",21,22,23,24,25,", "," & bytScene & ",") Then
5           bytType = 1
6         ElseIf bytScene = 3 Then
7           bytType = 2 '���������
8         End If
          '��ȡ������Ϣ
9         strPati = GetPatiInfo(lngPatiID, lngVisitId, strRegNo, bytType, lngRegId)

          '��ȡҽ����Ϣ
10        If TypeName(objMain) = "Recordset" Then
11            strMainInfo = GetMainInfo(lngPatiID, lngVisitId, strRegNo, lngRegId, objMain, colDiag)
12        ElseIf TypeName(objMain) = "Collection" Then
13            strMainInfo = GetMainJson(objMain)
14        End If
          
15        strJsonIn = "{""�ӿ�json_in"":""{\""cdss_in\"":{" & strPati
16        If strMainInfo <> "" Then
17            strJsonIn = strJsonIn & "," & strMainInfo & "}}""}"
18        Else
19            strJsonIn = strJsonIn & "}}""}"
20        End If
          
          '����YWJ
          'strJsonIn = TestJson(4)
          
21        If SubmitMainInfo(strJsonIn, strJsonOut) Then
              'strJsonOut = TestJson(5)
22            If strJsonOut = "" Then Exit Function
23            Set objJson = mdlJSON.parse(strJsonOut) 'Dictionary
24            If objJson Is Nothing Then
25                MsgBox "���ύ������Ϣ������ֵ����ʧ�ܣ�" & vbCrLf & _
                         " ԭ��:" & mdlJSON.GetParserErrors(), vbExclamation, gstrSysName
26                Exit Function
27            End If
              
28            If TypeName(objJson) = "Dictionary" Then
29                If objJson.Exists("inquiry") Then
                      '��������
30                    Set colList = objJson.Item("inquiry") 'Collection
31                    strJsonOut = ""
32                    If colList.Count > 0 Then
33                        If objJson.Exists("businss") Then
34                            strBusinss = objJson.Item("businss")
35                        End If
36                        If frmInquiryInfo.ShowMe(colList, strJsonOut) Then
37                            strJsonIn = "{""����json_in"":""{\""businss\"":\""" & strBusinss & "\"""
38                            If strJsonOut <> "" Then
39                                strJsonIn = strJsonIn & "," & strJsonOut & "}""}"
40                            Else
41                                strJsonIn = strJsonIn & "}""}"
42                            End If
      '                        Call SubmitMainByAdviceID("2187620,2187621")
                              
43                            If SubmitInquiriyInfo(strJsonIn, strJsonOut) Then
44                                If strJsonOut = "" Then Exit Function
45                                Set objJson = mdlJSON.parse(strJsonOut) 'Dictionary
46                                If objJson Is Nothing Then
47                                    MsgBox "���ύ������Ϣ������ֵ����ʧ�ܣ�" & vbCrLf & _
                                             " ԭ��:" & mdlJSON.GetParserErrors(), vbExclamation, gstrSysName
48                                    Exit Function
49                                End If
50                            End If
51                        End If
52                    End If
53                End If
                  '���ڹ����Ԥʱ������ʾ
54                If objJson.Exists("messages") Then
55                    Set colList = objJson.Item("messages")
56                    If colList.Count > 0 Then
57                        If HandleMessage(colList) Then blnRet = True
58                    End If
59                End If
60            Else
61                MsgBox "���ύ������Ϣ������ֵ��ʽ�����޷�����������", vbExclamation, gstrSysName
62                Exit Function
63            End If
64        End If
65        SubmitMain = blnRet
66        WriteLog "������SubmitMain ����" & vbNewLine
67        Exit Function
ErrH:
68        MsgBox "��zlCISRule.clsRule.SubmitMain�ĵ�" & Erl() & "�г�����" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function
 
Public Function SubmitMainByAdviceID(ByVal strAdviceIds As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : SubmitMainByAdviceID
      ' Author    : YWJ
      ' Date      : 2019-09-20 17:16:31
      ' Parameter :
      '             strAdviceIds-ҽ��IDs
      ' Purpose   : ���ڱ걾�ɼ��ύ������Ϣ
      ' Return    : T-��ֹ����;F-��������
      '---------------------------------------------------------------------------------------

          Dim strSQL As String
          Dim rsAdvice As ADODB.Recordset
          
1         On Error GoTo ErrH
2         WriteLog "������SubmitMainByAdviceID ��ʼ" & vbNewLine & _
                   "��Σ�" & strAdviceIds & vbNewLine
3         strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                   "    a.ID, Nvl(a.���id, 0) As ���id, a.����id, a.��ҳid, a.�Һŵ�, a.���, a.�������, a.ҽ����Ч, a.������Ŀid, a.�걾��λ, a.��鷽��, a.ҽ������,a.����ҽ��, a.�ܸ�����, a.��������, a.ִ��Ƶ��" & vbCrLf & _
                   "   From ����ҽ����¼ A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                   "   Where a.id = b.Column_Value" & vbCrLf & _
                   "   Order By a.���"
          
4         Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "SubmitMainByAdviceID", strAdviceIds)
          
5         If rsAdvice.EOF Then Exit Function
          
6         SubmitMainByAdviceID = SubmitMain(CLng(rsAdvice!����ID), Val(rsAdvice!��ҳID & ""), rsAdvice!�Һŵ� & "", rsAdvice, 4)
7         WriteLog "������SubmitMainByAdviceID ����" & vbNewLine
                   
8         Exit Function

ErrH:
9         MsgBox "��zlCISRule.clsRule.SubmitMainByAdviceID�ĵ�" & Erl() & "�г�����" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function
