Attribute VB_Name = "mdlTAT"
Option Explicit

Private mstrItem As String

Public Function getTATTime(ByVal intType As Integer, ByVal strItem As String, _
                           ByVal strDept As String, ByVal strGroup As String, _
                           ByVal strMachine As String, ByVal strSex As String, _
                           intMsg As Integer, Optional strShowBef As String, _
                           Optional lngTATTime As Long, Optional strUser As String) As String
      '����       ���TAT�Ƿ�ʱ

      '���
      'intType            '1=�ͼ�,2=ǩ��,3=����,4=���
      'strItem            '������ĿID  ��ĿID1,��Ŀ����1,�ϸ�ʱ��ڵ�1,����1;��ĿID2,��Ŀ����2,�ϸ�ʱ��ڵ�2,����2;
      'strDept            '�������ID
      'strGroup           '����С��ID
      'strMachine         '��������ID
      'strSex             '�����Ա�
      'strUser            '����Ա

      '����
      'intMsg             '����       1=ֻ��ʾ,2=��ʾ������
      'strShowBef         '��ʾ��Ϣ

      'GetTatTime=true��ʾ��ʱ,=false��ʾδ��ʱ

          Dim strSQL As String
          Dim Dtime As Date

          Dim var_tmp As Variant
          Dim var_ItemID As Variant
          Dim rsTAT As ADODB.Recordset
          Dim rsTATMX As ADODB.Recordset
          Dim rsOldItems As ADODB.Recordset
          Dim rsNewItems As ADODB.Recordset
          Dim strOldItemID As String
          Dim strOldItemCode As String
          Dim strNewItemID As String
          Dim strMsgShow As String
          Dim strOldMid As String
          Dim strNewMid As String
          Dim var_MidOld As Variant
          Dim var_MidNew As Variant
          Dim blnFind As Boolean
          Dim lngTATTimeBefor As Long
          Dim i As Integer, J As Integer

1         On Error GoTo getTATTime_Error

2         mstrItem = strItem
3         lngTATTime = 0

4         Select Case intType
          Case 1
5             strMsgShow = "δ�����걾�����ͼ�"
6         Case 2
7             strMsgShow = "δ�ͼ�걾���ܵǼ�"
8         Case 3
9             strMsgShow = "δ�ǼǱ걾���ܺ���"
10        Case 4
11            strMsgShow = "δ���ձ걾�������"
12        End Select
13        var_tmp = Split(strItem, ";")

          '��Ŀ����
14        If intType = 1 Or intType = 2 Then
              '�ͼ��ǩ��ʱ��Ҫ����Ŀ���ж���
15            strOldItemID = ""
16            For i = LBound(var_tmp) To UBound(var_tmp)
17                strOldItemID = strOldItemID & "," & Split(var_tmp(i), ",")(0)
18            Next
19            If strOldItemID <> "" Then strOldItemID = Mid(strOldItemID, 2)
              '�����ϰ���ĿID��ѯ�ϰ���Ŀ����
20            strSQL = "Select /*+cardinality(b,10)*/ A.id,A.���� From ������ĿĿ¼ A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.Id = b.Column_Value"
21            Set rsOldItems = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ���Ŀ����", strOldItemID)
22            strOldItemCode = ""
23            strOldMid = ""
24            Do While rsOldItems.EOF = False
25                strOldItemCode = strOldItemCode & "," & rsOldItems("����") & ""
26                strOldMid = strOldMid & ";" & rsOldItems("ID") & "," & rsOldItems("����") & ""
27                rsOldItems.MoveNext
28            Loop
29            If strOldItemCode <> "" Then strOldItemCode = Mid(strOldItemCode, 2)
30            If strOldMid <> "" Then strOldMid = Mid(strOldMid, 2)
              '�����ϰ���Ŀ�����ѯ�°���ĿID
31            If gUserInfo.NodeNo <> "-" Then
32                strSQL = "Select /*+cardinality(b,10)*/ A.ID,A.���Ʊ��� From ���������Ŀ A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.���Ʊ��� = b.Column_Value and (a.վ��=[2] or a.վ�� is null)"
33            Else
34                strSQL = "Select /*+cardinality(b,10)*/ A.ID,A.���Ʊ��� From ���������Ŀ A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.���Ʊ��� = b.Column_Value"
35            End If
36            Set rsNewItems = ComOpenSQL(Sel_Lis_DB, strSQL, "�°���ĿID", strOldItemCode, gUserInfo.NodeNo)
37            strNewMid = ""
38            Do While rsNewItems.EOF = False
39                strNewItemID = strNewItemID & "," & rsNewItems("ID")
40                strNewMid = strNewMid & ";" & rsNewItems("ID") & "," & rsNewItems("���Ʊ���") & ""
41                rsNewItems.MoveNext
42            Loop
43            If strNewItemID <> "" Then
44                strNewItemID = Mid(strNewItemID, 2)
45                If strNewMid <> "" Then strNewMid = Mid(strNewMid, 2)

46                var_ItemID = Split(strNewItemID, ",")
47                var_MidOld = Split(strOldMid, ";")
48                var_MidNew = Split(strNewMid, ";")
49                For J = LBound(var_MidOld) To UBound(var_MidOld)
50                    For i = LBound(var_MidNew) To UBound(var_MidNew)
51                        If Split(var_MidNew(i), ",")(1) = Split(var_MidOld(J), ",")(1) Then
52                            strItem = Replace(strItem, Split(var_MidOld(J), ",")(0) & ",", Split(var_MidNew(i), ",")(0) & ",")
53                        End If
54                    Next
55                Next

56                If strItem <> "" Then
57                    var_tmp = Split(strItem, ";")
58                End If
59            End If
60        End If


61        For i = LBound(var_tmp) To UBound(var_tmp)
62            blnFind = False
              '���ݴ�����Ŀ��ѯTATʱ��
63            strSQL = "Select Distinct a.Id, a.�ͼ���ʱ, a.ǩ����ʱ, a.������ʱ," & _
                     " a.�����ʱ, a.Ӧ�ÿ���,a.Ӧ��С��, a.Ӧ������, a.�Ա�," & _
                     " a.����, a.����, a.��ʾ��Ϣ From ����tatʱ�� A, ����tatʱ����ϸ B" & _
                     " Where a.Id = b.Tatʱ��id And a.�Ƿ���Ч = 1 and b.������Ŀid = [1] and a.����=[2]"
64            Select Case intType
              Case 1
65                strSQL = strSQL & " and a.�ͼ���ʱ is  not null"
66            Case 2
67                strSQL = strSQL & " and a.ǩ����ʱ is  not null"
68            Case 3
69                strSQL = strSQL & " and a.������ʱ is  not null"
70            Case 4
71                strSQL = strSQL & " and a.�����ʱ is  not null"
72            End Select
73            Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "����TATʱ��", Split(var_tmp(i), ",")(0), Split(var_tmp(i), ",")(3))
              '���ݲ�ѯ������TATʱ��ID����ѯ��ص�TATʱ����ϸ��
74            Do While rsTAT.EOF = False
75                blnFind = True
76                If Not IsNull(rsTAT("ID")) Then
77                    strSQL = "Select f_List2str(Cast(Collect(To_Char(�������)) As t_Strlist)) �������," & _
                             " f_List2str(Cast(Collect(To_Char(����С��id)) As t_Strlist)) ����С��id," & _
                             " f_List2str(Cast(Collect(To_Char(��������id)) As t_Strlist)) ��������id" & _
                             " From ����tatʱ����ϸ Where Tatʱ��id =[1]"
78                    Set rsTATMX = ComOpenSQL(Sel_Lis_DB, strSQL, "����TATʱ����ϸ", Val(rsTAT("ID")))
                      '��¼����ʾ������ʾ����ֹ
                      '                intMsg = rsTAT("����")
79                    strShowBef = rsTAT("��ʾ��Ϣ") & ""

                      '����
80                    If Split(var_tmp(i), ",")(3) <> 1 And Val(rsTAT("����") & "") = 1 Then
81                        blnFind = False
82                    End If
                      '����
83                    If Val(rsTAT("Ӧ�ÿ���") & "") = 2 Then
84                        If InStr("," & rsTATMX("�������") & ",", IIf(strDept = "", ",strDept,", "," & strDept & ",")) <= 0 Then
85                            blnFind = False
86                        End If
87                    End If
                      '�Ա�
88                    If rsTAT("�Ա�") & "" <> strSex And rsTAT("�Ա�") & "" <> "����" And Not IsNull(rsTAT("�Ա�")) Then
89                        blnFind = False
90                    End If

91                    If intType = 3 Or intType = 4 Then
                          'С��
92                        If Val(rsTAT("Ӧ��С��") & "") = 2 Then
93                            If rsTATMX("����С��id") & "" <> "" And InStr("," & rsTATMX("����С��id") & ",", IIf(strGroup = "", ",strGroup,", "," & strGroup & ",")) <= 0 Then
94                                blnFind = False
95                            End If
96                        End If
                          '����
97                        If Val(rsTAT("Ӧ������") & "") = 2 Then
98                            If InStr("," & rsTATMX("��������ID") & ",", IIf(strMachine = "", ",strMachine,", "," & strMachine & ",")) <= 0 Then
99                                blnFind = False
100                           End If
101                       End If
102                   End If

103                   If Split(var_tmp(i), ",")(2) = "" Then
104                       Dtime = CDate(Format("2000/01/01 01:01:01", "yyyy/mm/dd hh:mm:ss"))
105                   Else
106                       Dtime = CDate(Split(var_tmp(i), ",")(2))
107                   End If
108                   If blnFind = True Then Exit Do
109               End If
110               rsTAT.MoveNext
111           Loop

112           If blnFind = True Then

113               Call GetMsgItems(Dtime, IIf(intType = 1, Val(rsTAT("�ͼ���ʱ") & ""), _
                                              IIf(intType = 2, Val(rsTAT("ǩ����ʱ") & ""), _
                                                  IIf(intType = 3, Val(rsTAT("������ʱ") & ""), _
                                                      IIf(intType = 4, Val(rsTAT("�����ʱ") & ""), 0)))), _
                                                      Split(var_tmp(i), ",")(1), strShowBef, CInt(rsTAT("����")), intType, strUser)

114           End If
              '����ʱ��Ҫ�������TAT��ʱ
115           If intType = 3 Then
                  'д��ҽ��TAT���׶ε�����ʱ��
116               Call setTATAllTime(Val(Split(var_tmp(i), ",")(0)), Val(Split(var_tmp(i), ",")(5)), Val(Split(var_tmp(i), ",")(4)), strSex, Split(var_tmp(i), ",")(3), strDept, strGroup, strMachine)
                  
                  '����ʱ��ȡ����ʱ
117               lngTATTimeBefor = GetTATAllTimefun(Val(Split(var_tmp(i), ",")(4)), Val(Split(var_tmp(i), ",")(5)))
118               If lngTATTime = 0 And lngTATTimeBefor <> 0 Then
119                   lngTATTime = lngTATTimeBefor
120               ElseIf lngTATTime > lngTATTimeBefor Then
121                   lngTATTime = lngTATTimeBefor
122               End If
123           End If
124       Next

125       getTATTime = mstrItem


126       Exit Function
getTATTime_Error:
127       Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "ִ��(getTATTime)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
128       Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/12/11
'��    ��:��ȡTAT����ʱ
'��    ��:
'           lngApplyID      �����������.����ID
'           lngAdviceID     �����������.ҽ��ID
'��    ��:
'           lngTATTime      TAT����ʱ
'��    ��:
'---------------------------------------------------------------------------------------
Public Function GetTATAllTimefun(ByVal lngAdviceID As Long, ByVal lngApplyID As Long) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strTATBegin As String
          Dim intTATBegin As Integer
          
          '��ȡ����
1         On Error GoTo GetTATAllTimefun_Error
2         If lngAdviceID = 0 Or lngApplyID = 0 Then Exit Function
          
3         strTATBegin = ComGetPara(Sel_Lis_DB, "TAT����ʱ��ʼ��", gSysInfo.SysNo, gSysInfo.ModlNo)
4         If strTATBegin = "" Then
5             intTATBegin = 3
6         Else
7             intTATBegin = Val(strTATBegin)
8         End If
          
          '��ȡ����ʱ
9         Select Case intTATBegin
              Case 0
10                strSQL = "select �ͼ���ʱ+ǩ����ʱ+������ʱ+�����ʱ TATʣ��ʱ�� from ����������� where ����ID=[1] and ҽ��ID=[2]"
11            Case 1
12                strSQL = "select ǩ����ʱ+������ʱ+�����ʱ TATʣ��ʱ�� from ����������� where ����ID=[1] and ҽ��ID=[2]"
13            Case 2
14                strSQL = "select ������ʱ+�����ʱ TATʣ��ʱ�� from ����������� where ����ID=[1] and ҽ��ID=[2]"
15            Case 3
16                strSQL = "select �����ʱ TATʣ��ʱ�� from ����������� where ����ID=[1] and ҽ��ID=[2]"
17        End Select
18        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngApplyID, lngAdviceID)
19        If rsTmp.RecordCount > 0 Then
20            GetTATAllTimefun = Val(rsTmp("TATʣ��ʱ��") & "")
21        End If


22        Exit Function
GetTATAllTimefun_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "ִ��(GetTATAllTimefun)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear
End Function

Private Function GetMsgItems(ByVal Dtime As Date, ByVal lngTime As Long, _
                            ByVal strItem As String, ByVal strShowBef As String, _
                            ByVal intMsg As Integer, ByVal intType As Integer, ByVal strUser As String) As String
          '����û�г�ʱ����Ŀ
          'Dtime          ��һ��ʱ��ڵ�
          'lngTime        tat��ʱ
          'strItem        ��Ŀ
          'strShowBef     ��ʾ��Ϣ
          '��ʾ����        0=ֻд��־,1=��ʾ��д��־,2=д��־����ֹ
          'intType        ��Դ 1=�ͼ�,2=ǩ��,3=����,4=���
          
          '�����ַ�����ʽ
              '���ID,�������,��һ��ʱ��ڵ�,�Ƿ���,ҽ��ID,���ID,����,��ʱʱ��,��ʾ��Ϣ,��ʾ����
          Dim dCurrentdate As Date
          Dim var_tmp As Variant
          Dim var_tmp1 As Variant
          Dim strFrom As String
          Dim strSQL As String
          Dim i As Integer
              
1         On Error GoTo GetMsgItems_Error

2         dCurrentdate = Currentdate
          
3         strShowBef = Replace(strShowBef, ",", "")
          
4         If DateDiff("n", Dtime, dCurrentdate) > lngTime Then
          
              'д����������־
5             Select Case intType
                  Case 1
6                     strFrom = "TAT�ͼ쳬ʱ"
7                 Case 2
8                     strFrom = "TATǩ�ճ�ʱ"
9                 Case 3
10                    strFrom = "TAT���ճ�ʱ"
11                Case 4
12                    strFrom = "TAT��˳�ʱ"
13            End Select
14            strSQL = "Zl_���������־_Insert(19,6,'" & strUser & "',null,'" & strFrom & "','ҽ��ID" & Split(mstrItem, ",")(4) & "|" & Replace(Replace(strShowBef, "[��Ŀ]", strItem), "[��ʱ]", DateDiff("n", Dtime, dCurrentdate) - lngTime) & "����')"
15            ComExecuteProc Sel_Lis_DB, strSQL, "���������־"
              
16            If strItem <> "" Then
17                var_tmp = Split(mstrItem, ";")
18                For i = LBound(var_tmp) To UBound(var_tmp)
19                    var_tmp1 = Split(var_tmp(i), ",")
20                    If InStr(var_tmp(i), strItem) > 0 Then
21                        If UBound(var_tmp1) > 8 Then
22                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp1(0) & "," & var_tmp1(1) & "," & var_tmp1(2) & "," & _
                                          var_tmp1(3) & "," & var_tmp1(4) & "," & var_tmp1(5) & "," & var_tmp1(6) & "," & (DateDiff("n", Dtime, dCurrentdate) - lngTime) & "," & strShowBef & "," & intMsg)
23                        Else
24                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp(i) & "," & DateDiff("n", Dtime, dCurrentdate) - lngTime & "," & strShowBef & "," & intMsg)
25                        End If
26                    Else
27                        If UBound(var_tmp1) < 9 Then
28                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp(i) & ",0" & "," & strShowBef & "," & intMsg)
29                        End If
30                    End If
31                Next
              
32            End If
33        End If


34        Exit Function
GetMsgItems_Error:
35        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "ִ��(GetMsgItems)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
36        Err.Clear
          
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/12/11
'��    ��:д��ҽ��TAT���׶ε�����ʱ��
'��    ��:
'           lngItemID           �����ĿID
'           lngApplyID          �����������.����ID
'           lngAdivceID         �����������.ҽ��ID
'           strSex              �Ա�
'           strJiZhen           �Ƿ��� 0=��1=��
'           strDept             �������
'           strGroup            ����С��
'           strMachine          ��������
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function setTATAllTime(ByVal lngItemid As Long, ByVal lngApplyID As Long, ByVal lngAdivceID As Long, ByVal strSex As String, ByVal strJiZhen As String, _
                               ByVal strDept As String, ByVal strGroup As String, ByVal strMachine As String)
          Dim strSQL As String
          Dim rsTAT As ADODB.Recordset
          Dim rsTATMX As ADODB.Recordset
          Dim lngSJXS As Long             '�ͼ���ʱ
          Dim lngQSXS As Long             'ǩ����ʱ
          Dim lngHSXS As Long             '������ʱ
          Dim lngSHXS As Long             '�����ʱ
          
          Dim blnTATTime As Boolean


1         On Error GoTo setTATAllTime_Error

2         If lngItemid = 0 Then Exit Function
3         If lngApplyID = 0 Or lngAdivceID = 0 Then Exit Function
          
'          '���ҽ���Ƿ��Ѿ�д��TAT��ʱ�䣬�����д�룬�����ظ�д��
'4         strSQL = "Select Count(*) ����" & vbCrLf & _
'                   " From ����������� " & vbCrLf & _
'                   " Where ����id = [1] And ҽ��id = [2] And �ͼ���ʱ Is not Null And ǩ����ʱ Is not Null And ������ʱ Is not Null And �����ʱ Is not Null"
'5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngApplyID, lngAdivceID)
'6         If rsTmp.RecordCount > 0 Then
'7             If Val(rsTmp("����") & "") > 0 Then Exit Function
'8         Else
'9             Exit Function
'10        End If
          
          '��ѯ��Ŀ��ص�TAT����
11        strSQL = "Select a.id, Decode(Sign(Length(a.�ͼ���ʱ)), 1, '�ͼ���ʱ',Decode(Sign(Length(a.ǩ����ʱ)), 1, 'ǩ����ʱ'," & _
                   " Decode(Sign(Length(a.������ʱ)), 1, '������ʱ', Decode(Sign(Length(a.�����ʱ)), 1, '�����ʱ', Null)))) ʱ������," & _
                   " Decode(Sign(Length(a.�ͼ���ʱ)), 1, a.�ͼ���ʱ,Decode(Sign(Length(a.ǩ����ʱ)), 1, a.ǩ����ʱ," & _
                   " Decode(Sign(Length(a.������ʱ)), 1, a.������ʱ, Decode(Sign(Length(a.�����ʱ)), 1, a.�����ʱ, Null)))) ʱ��," & _
                   " A.Ӧ�ÿ��� , A.Ӧ��С��, A.Ӧ������,a.����,a.�Ա� From ����tatʱ�� A, ����tatʱ����ϸ B Where a.Id = b.Tatʱ��id And a.�Ƿ���Ч = 1" & _
                   " And b.������Ŀid = [1] And (a.�Ա� = '����' or a.�Ա� is null Or a.�Ա� = [2]) and a.����=[3]"

12        Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "����tatʱ��", lngItemid, strSex, strJiZhen)
          
          '���ͨ���Ա𡢼���û�в�ѯ������ֻͨ��IDȥ��ѯ
13        If rsTAT.RecordCount < 1 Then
14            strSQL = "Select a.id, Decode(Sign(Length(a.�ͼ���ʱ)), 1, '�ͼ���ʱ',Decode(Sign(Length(a.ǩ����ʱ)), 1, 'ǩ����ʱ'," & _
                   " Decode(Sign(Length(a.������ʱ)), 1, '������ʱ', Decode(Sign(Length(a.�����ʱ)), 1, '�����ʱ', Null)))) ʱ������," & _
                   " Decode(Sign(Length(a.�ͼ���ʱ)), 1, a.�ͼ���ʱ,Decode(Sign(Length(a.ǩ����ʱ)), 1, a.ǩ����ʱ," & _
                   " Decode(Sign(Length(a.������ʱ)), 1, a.������ʱ, Decode(Sign(Length(a.�����ʱ)), 1, a.�����ʱ, Null)))) ʱ��," & _
                   " A.Ӧ�ÿ��� , A.Ӧ��С��, A.Ӧ������,a.����,a.�Ա� From ����tatʱ�� A, ����tatʱ����ϸ B Where a.Id = b.Tatʱ��id And a.�Ƿ���Ч = 1" & _
                   " And b.������Ŀid = [1]"
15             Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "����tatʱ��", lngItemid, strSex)
16        End If
               
          'ɸѡ��¼����ȡ��ȷ��TAT��ʱ
17        Do While rsTAT.EOF = False
18            strSQL = "Select f_List2str(Cast(Collect(To_Char(�������)) As t_Strlist)) �������," & _
                      " f_List2str(Cast(Collect(To_Char(����С��id)) As t_Strlist)) ����С��id," & _
                      " f_List2str(Cast(Collect(To_Char(��������id)) As t_Strlist)) ��������id" & _
                      " From ����tatʱ����ϸ Where Tatʱ��id =[1]"
19            Set rsTATMX = ComOpenSQL(Sel_Lis_DB, strSQL, "����TATʱ����ϸ", Val(rsTAT("ID")))
20            blnTATTime = True
              '����
21            If strJiZhen <> 1 And Val(rsTAT("����") & "") = 1 Then
22                blnTATTime = False
23            End If
              '����
24            If Val(rsTAT("Ӧ�ÿ���") & "") = 2 Then
25                If InStr("," & rsTATMX("�������") & ",", IIf(strDept = "", ",strDept,", "," & strDept & ",")) <= 0 Then
26                    blnTATTime = False
27                End If
28            End If
              '�Ա�
29            If rsTAT("�Ա�") & "" <> strSex And rsTAT("�Ա�") & "" <> "����" And Not IsNull(rsTAT("�Ա�")) Then
30                blnTATTime = False
31            End If
              
32            If rsTAT("ʱ������") & "" = "������ʱ" Or rsTAT("ʱ������") & "" = "�����ʱ" Then
                  'С��
33                If Val(rsTAT("Ӧ��С��") & "") = 2 Then
34                    If rsTATMX("����С��id") & "" <> "" And InStr("," & rsTATMX("����С��id") & ",", IIf(strGroup = "", ",strGroup,", "," & strGroup & ",")) <= 0 Then
35                        blnTATTime = False
36                    End If
37                End If
                  '����
38                If Val(rsTAT("Ӧ������") & "") = 2 Then
39                    If InStr("," & rsTATMX("��������ID") & ",", IIf(strMachine = "", ",strMachine,", "," & strMachine & ",")) <= 0 Then
40                        blnTATTime = False
41                    End If
42                End If
43            End If
              
              'д�����ݿ�
44            If blnTATTime = True Then
45                If rsTAT("ʱ������") & "" = "�ͼ���ʱ" Then
46                    lngSJXS = Val(rsTAT("ʱ��") & "")
47                ElseIf rsTAT("ʱ������") & "" = "ǩ����ʱ" Then
48                    lngQSXS = Val(rsTAT("ʱ��") & "")
49                ElseIf rsTAT("ʱ������") & "" = "������ʱ" Then
50                    lngHSXS = Val(rsTAT("ʱ��") & "")
51                ElseIf rsTAT("ʱ������") & "" = "�����ʱ" Then
52                    lngSHXS = Val(rsTAT("ʱ��") & "")
53                End If
54            End If
55            rsTAT.MoveNext
56        Loop
          
57        If lngSJXS <> 0 Or lngQSXS <> 0 Or lngHSXS <> 0 Or lngSHXS <> 0 Then
58            strSQL = "Zl_�����������_Tat��ʱ(" & lngApplyID & "," & lngAdivceID & "," & lngSJXS & "," & lngQSXS & "," & lngHSXS & "," & lngSHXS & ")"
59            Call ComExecuteProc(Sel_Lis_DB, strSQL, "TAT��ʱ")
60        End If
          
61        Exit Function
setTATAllTime_Error:
62        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "ִ��(setTATAllTime)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
63        Err.Clear

End Function

