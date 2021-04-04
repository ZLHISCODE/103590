Attribute VB_Name = "mdlCISPath"
Option Explicit
Public gobjFile As New FileSystemObject     '�ļ���������
Public gfrmMain As Object                   '����̨����
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gMainPrivs As String                 '���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Public gstrPrivs As String
Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String            'OEM��Ʒ����
Public glngSys As Long
Public glngModul As Long

Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjEmr As Object                    '�°����ܵ��Ӳ�������
Public gcolIcons As Collection              '��������ٴ�·��ͼ�꼯
Public gobjPlugIn As Object                 '��ҹ��ܶ���
Public gobjLIS As Object                    'LIS��������
Public gbln˫��� As Boolean
Public glngHwnd As Long                     '��������
Public gblnGetPath As Boolean

'ϵͳ����
Public gstrLike As String   '�����˫��ƥ�䣬��Ϊ%
Public gint���� As Integer  '����ƥ�䷽ʽ��0-ƴ��,1-���

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p�ٴ�·������ = 1078
    pסԺ���ʲ��� = 1150
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    P�ٴ�·��Ӧ�� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p���Ӳ������� = 1259
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    P����·��Ӧ�� = 1248
    P����·������ = 1083
    P����·������ = 1272
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.����ID = NVL(rsTmp!����ID, 0)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.���� = Get��Ա����
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.�û���
End Function

Public Sub InitSysPar()
'���ܣ���ʼ��ϵͳ����
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
End Sub

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    If str���� <> "" Then
        strSql = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", str����)
    Else
        strSql = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceDefineText(ByVal strҽ��IDs As String, Optional rsAdvice As ADODB.Recordset) As String
'���ܣ���ȡ·����Ŀ��Ӧ��ҽ����������������
'������rsAdvice=�ڴ��¼������������򲻴����ݿ��ȡ
    Dim rsTmp As ADODB.Recordset
    Dim strFilter As String, lngPre���ID As Long
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    
    If Not rsAdvice Is Nothing Then
        '���ɶ�̬SQL
        For i = 0 To UBound(Split(strҽ��IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(strҽ��IDs, ",")(i)
        Next
        With rsAdvice
            strSql = ""
            .Filter = Mid(strFilter, 5)
            Do While Not .EOF
                strSql = strSql & " Union ALL Select "
                For i = 0 To .Fields.count - 1
                    If Not IsNull(.Fields(i).Value) Then
                        If Rec.IsType(.Fields(i).Type, adVarChar) Then
                            strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                        Else
                            strSql = strSql & .Fields(i).Value 'û��������
                        End If
                    Else
                        strSql = strSql & "Null"
                    End If
                    strSql = strSql & " As " & .Fields(i).Name & ","
                Next
                strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
                .MoveNext
            Loop
            .Filter = ""
            strSql = "(" & Mid(strSql, 12) & ")"
        End With
    Else
        strSql = "·��ҽ������"
    End If
    
    strSql = "Select /*+ Rule*/ A.ID,A.���ID,Decode(A.��Ч,1,'����','����') as ��Ч,B.���," & _
        " Nvl(A.ҽ������,B.����)||Decode(C.���,NULL,NULL,'('||C.���||')') as ����," & _
        " Decode(A.��������,NULL,NULL,A.��������||'""'||B.���㵥λ||'""') as ����," & _
        " Decode(A.�ܸ�����,NULL,NULL,Decode(Instr('56',Nvl(B.���,'*')),0,A.�ܸ�����||'""'||B.���㵥λ||'""',A.�ܸ�����/D.סԺ��װ||'""'||D.סԺ��λ||'""')) as ����," & _
        " A.ִ��Ƶ��,A.ҽ������" & _
        " From " & strSql & " A,������ĿĿ¼ B,�շ���ĿĿ¼ C,ҩƷ��� D" & _
        " Where Nvl(A.������ĿID,0)=B.ID(+) And Nvl(A.�շ�ϸĿID,0)=C.ID(+) And Nvl(A.�շ�ϸĿID,0)=D.ҩƷID(+)" & _
        " And A.ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetAdviceDefineText", strҽ��IDs)
    
    strSql = ""
    Do While Not rsTmp.EOF
        If (InStr(",5,6,C,*,", NVL(rsTmp!���, "*")) > 0 Or IsNull(rsTmp!���id)) _
            And Not (NVL(rsTmp!���) = "E" And rsTmp!ID = lngPre���ID) Then
            strSql = strSql & vbCrLf & "����" & rsTmp!��Ч & "��" & rsTmp!���� & _
                IIf(Not IsNull(rsTmp!����), "��ÿ��" & rsTmp!����, "") & _
                IIf(Not IsNull(rsTmp!����), "����" & rsTmp!����, "") & _
                IIf(Not IsNull(rsTmp!ִ��Ƶ��), "��" & rsTmp!ִ��Ƶ��, "") & _
                IIf(Not IsNull(rsTmp!ҽ������), "��" & rsTmp!ҽ������, "")
        End If
        
        lngPre���ID = NVL(rsTmp!���id, 0)
        rsTmp.MoveNext
    Loop
    
    GetAdviceDefineText = Mid(strSql, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEPRDefineText(Optional ByVal str����IDs As String, Optional ByVal lng��ĿID As Long) As String
'���ܣ���ȡ·����Ŀ��Ӧ�Ĳ�����������������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    If lng��ĿID <> 0 Then '�°���Ӳ������ϰ�ͬʱ
        strSql = "Select Nvl(a.����, b.����) as ���� From �ٴ�·������ A, �����ļ��б� B Where a.��Ŀid = [2] And a.�ļ�id = b.Id(+)" & vbNewLine & _
                "order by a.���"
    ElseIf str����IDs <> "" And lng��ĿID = 0 Then '�ϰ�
        strSql = "Select /*+ Rule*/ ���� From �����ļ��б�" & _
            " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " Order by ���"
    Else '�°�
        strSql = "select ���� from �ٴ�·������ t where t.��Ŀid=[2] and t.�ļ�id is null and t.ԭ��id IN (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))) order by ���"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetEPRDefineText", str����IDs, lng��ĿID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "��" & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    GetEPRDefineText = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Checkҽ����Ŀ(ByVal lngִ��ID As Long) As Boolean
'���ܣ����ָ����ִ����Ŀ�Ƿ�����ҽ����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From ����·��ҽ�� Where ·��ִ��ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Checkҽ����Ŀ", lngִ��ID)
    
    Checkҽ����Ŀ = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Check������Ŀ(ByVal lngִ��ID As Long) As Boolean
'���ܣ����ָ����ִ����Ŀ�Ƿ����ڲ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From ���Ӳ�����¼ Where ·��ִ��ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Check������Ŀ", lngִ��ID)
    
    Check������Ŀ = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSameDayOfPhase(ByVal lngPhase As Long, ByVal lngDay As Long) As Boolean
'���ܣ���鵱���Ƿ������õ����������׶�(��ǰ�׶μ���֧����)
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    '�����ǰ�Ƿ�֧�׶Σ���ȡ�丸ID
    strSql = "Select ��ID From �ٴ�·���׶� Where ID = [1] And ��ID is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶�", lngPhase)
    If rsTmp.RecordCount > 0 Then lngPhase = rsTmp!��ID
    
    strSql = "Select 1" & vbNewLine & _
            "From �ٴ�·���׶� A, �ٴ�·���׶� B" & vbNewLine & _
            "Where a.Id = [1] And a.·��id = b.·��id And a.�汾�� = b.�汾��  And nvl(a.��֧id,0)=nvl(b.��֧ID,0) And b.��� > a.���" & vbNewLine & _
            "      And [2] Between b.��ʼ���� And Nvl(b.��������, b.��ʼ����) And b.��ID is Null And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶�", lngPhase, lngDay)
    If rsTmp.RecordCount > 0 Then CheckSameDayOfPhase = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInPath(t_pati As TYPE_Pati, ByVal lng����·��Id As Long, Optional ByRef lngȷ������ As Long) As Date
'���ܣ���ȡ���˵Ľ���·���Ŀ�ʼʱ��
'����������lngȷ������=��ǰ��ѡ�׶εĵ�����
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select a.��ʼʱ��,b.ȷ������  From �����ٴ�·�� a,�ٴ�·��Ŀ¼ b Where a.Id =[1] And a.·��id = b.id"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�뾶ʱ��", lng����·��Id)
    If IsNull(rsTmp!��ʼʱ��) Then
        GetPatiInPath = zlDatabase.Currentdate
        'GetPatiInPath = GetPatiInDate(t_pati)
    Else
        GetPatiInPath = rsTmp!��ʼʱ��
    End If
    
    lngȷ������ = Val("" & rsTmp!ȷ������)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInDate(t_pati As TYPE_Pati, Optional lng��Ժ���� As Long) As Date
'���ܣ���ȡ���˵���Ժ�����\ת��ʱ��
'���أ�lng��Ժ��������Ժ���������
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Max(��ʼʱ��) ��ʼʱ��,To_number(Trunc(Sysdate)-Trunc(Max(��ʼʱ��)))+1 as ��Ժ����" & vbNewLine & _
            "From (Select ��Ժ���� As ��ʼʱ��" & vbNewLine & _
            "       From ������ҳ" & vbNewLine & _
            "       Where ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select ��ʼʱ�� From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� In (2, 3) And ����id = [3])"
           
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ʱ��", t_pati.����ID, t_pati.��ҳID, t_pati.����ID)
    GetPatiInDate = CDate(rsTmp!��ʼʱ��)
    lng��Ժ���� = rsTmp!��Ժ����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select NVL(B.����,a.����) ����,NVL(B.�Ա�,a.�Ա�) �Ա� ,NVL(B.����,a.����) ���� , To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') ��������, b.��ǰ����," & vbNewLine & _
            "       d.���� As ��������,a.����� ,b.סԺ��, b.��Ժ����, b.��Ժ����,e.���� As ����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� C, �ٴ��������� D, ���ű� E" & vbNewLine & _
            "Where a.����id = b.����id And b.����id = [1] And b.��ҳid = [2] And e.Id = b.��ǰ����id And" & vbNewLine & _
            "      b.����id = c.����id(+) And b.��ҳid = c.��ҳid(+) And c.��Ϣ��(+) = '��������' And c.��Ϣֵ = d.����(+)"
            
    On Error GoTo errH
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", lng����ID, lng��ҳID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetAdvice(strIDs As String) As ADODB.Recordset
'���ܣ���ȡ·����Ŀ��Ӧ��ҽ����¼��
    Dim strSql As String
 
    strSql = "Select /*+ rule*/ a.·����ĿID,a.ҽ������ID,b.��Ч,Nvl(b.���ID,b.ID) ���ID,b.������ĿID" & vbNewLine & _
            "From �ٴ�·��ҽ�� A,·��ҽ������ B,(Select Column_Value As ID From Table(f_Num2list([1]))) C" & vbNewLine & _
            "Where a.ҽ������id=b.id And a.·����Ŀid = c.Id" & vbNewLine & _
            "Order by b.���"
    On Error GoTo errH
    Set GetAdvice = zlDatabase.OpenSQLRecord(strSql, "��ȡҽ����¼", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFile(strIDs As String, Optional ByVal int���� As Integer = 2) As ADODB.Recordset
'���ܣ���ȡ·����Ŀ��Ӧ�Ĳ����ļ���¼��
'int����=1���2-סԺ
    Dim strSql As String

    strSql = "select A.��ĿID as ·����ĿID,A.�ļ�ID,A.ԭ��ID,B.����  from " & IIf(int���� = 1, "����·������", "�ٴ�·������") & " a,�����ļ��б� b  where a.��Ŀid in (Select Column_Value From Table(f_Num2list([1]))) and a.�ļ�ID=b.id(+)"
  
    On Error GoTo errH
    Set GetFile = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����ļ�", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDelPathItem(ByVal lngִ��ID As Long, ByVal int���� As Integer) As Boolean
'���ܣ����ָ����ҽ����·����Ŀִ�м�¼�Ƿ����ɾ������������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strIDs As String
    Dim i As Long

    '���ǵ������ɵĳ������������ɺ��Զ�ֹͣ�������Ƿ��ͣ�
    '�ǵ������ɵĳ�������У�Ե�δ���ϣ�������ȡ��(��ֹͣ��Ҳ������)��δУ�Եģ�ȡ��ʱ�Զ�ɾ����Ӧ��ҽ����
    strSql = "Select 1" & vbNewLine & _
             "From ����·��ҽ�� A, ����·��ҽ�� B" & vbNewLine & _
             "Where a.·��ִ��id = [1] And a.����ҽ��id = b.����ҽ��id And b.·��ִ��id <> a.·��ִ��id  And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
    If rsTmp.RecordCount = 0 Then '��������
        strSql = "Select 1 From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And c.ҽ��״̬ <> 4 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ������У�Ե�δ���ϵ�ҽ������������ҽ������ִ�д˲�����", vbInformation, gstrSysName
            Exit Function
        End If

        If int���� = 1 Then
            '�����Ѿ�����˵�ҽ�����������޸�ɾ����
            strSql = "Select 1 From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.����ҽ�� Like '%/%' And rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "����Ŀ��Ӧ��ҽ���Ѿ���ҽ����ˣ�����ִ�д˲�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else '�ǵ�������
        'ǰ��У�Ժ�δͣ�õĳ�������·������Ŀ����ʽ��·������չʾ�����Ҫɾ����,��ȷ�����೤����ֹͣ�����ϲ���ɾ��
        strSql = "Select c.ҽ������" & vbNewLine & _
                 "From ����·��ҽ�� B, ����ҽ����¼ C,����·��ִ�� D " & vbNewLine & _
                 "Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And d.id=b.·��ִ��Id And d.��ĿID is null " & _
                 " And c.ҽ��״̬ Not In (4,8,9) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            strIDs = ""
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & vbNewLine & rsTmp!ҽ������
                rsTmp.MoveNext
            Next
            MsgBox "��·������Ŀ������У�Ե�δ���ϻ�ֹͣ�ĳ���ҽ����" & strIDs & vbNewLine & "�������ϻ�ֹͣ��ҽ������ִ��ȡ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDelPathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPathIcon(ByVal lngͼ��ID As Long) As StdPicture
'���ܣ���ȡָ��ID���ٴ�·��ͼ��
'˵������һ�ζ�ȡʱ���ü��Ͻ��л���
    Dim rsTmp As ADODB.Recordset
    Dim objIcon As StdPicture
    Dim strFile As String, strSql As String
    Dim blnExist As Boolean
    
    blnExist = True
    If gcolIcons Is Nothing Then
        Set gcolIcons = New Collection
        blnExist = False
    End If
    If blnExist Then
        On Error Resume Next
        Set GetPathIcon = gcolIcons("_" & lngͼ��ID)
        If Err.Number <> 0 Then
            Err.Clear: blnExist = False
        End If
    End If
    
    On Error GoTo errH
    
    If Not blnExist Then
        Screen.MousePointer = 11
        
        strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\zlTemplate.bmp"
                
        strSql = "Select ID From �ٴ�·��ͼ�� Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPathIcon", lngͼ��ID)
        If Not rsTmp.EOF Then
            If sys.ReadLob(glngSys, 11, lngͼ��ID, strFile) <> "" Then
                gcolIcons.Add LoadPicture(strFile), "_" & lngͼ��ID
                gobjFile.DeleteFile strFile
            End If
        End If
        
        Screen.MousePointer = 0
    End If
    
    Set GetPathIcon = gcolIcons("_" & lngͼ��ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNextCode(ByVal str���� As String, Optional ByVal intMode As Integer = 0) As String
'���ܣ���ȡָ��������ٴ�·����ȱʡ�±���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intMax As Integer
    Dim strTable As String
    
    On Error GoTo errH
    
    If intMode = 1 Then
        strTable = "����·��Ŀ¼"
    Else
        strTable = "�ٴ�·��Ŀ¼"
    End If
    'ȡ��󳤶ȣ���Max��001��0001��
    If str���� = "" Then
        strSql = "Select Max(Length(����)) As ���� From " & strTable & " Where ���� Is Null"
    Else
        strSql = "Select Max(Length(����)) As ���� From " & strTable & " Where ����=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextCode", str����)
    If rsTmp.EOF Then
        GetNextCode = "01": Exit Function
    ElseIf IsNull(rsTmp!����) Then
        GetNextCode = "01": Exit Function
    Else
        intMax = rsTmp!����
    End If
    
    '����󳤶ȱ���
    If str���� = "" Then
        strSql = "Select Max(����) As ���� From " & strTable & " Where ���� Is Null And Length(����)=[2]"
    Else
        strSql = "Select Max(����) As ���� From " & strTable & " Where ����=[1] And Length(����)=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextCode", str����, intMax)
    GetNextCode = zlCommFun.IncStr(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetNextPhase(ByVal lng�׶�ID As Long, ByVal lng��ǰ�׶η�֧ID As Long) As Long
'���ܣ���ȡָ���׶εĺ����׶�ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ��ID From �ٴ�·���׶� Where id = [1] And ��ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��һ�׶�", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then lng�׶�ID = Val(rsTmp!��ID)
    
    strSql = "Select b.ID From �ٴ�·���׶� a,�ٴ�·���׶� b " & _
            "Where a.·��ID= b.·��ID And a.�汾��= b.�汾�� And b.���>a.��� And NVL(b.��֧ID,0)=[2] And a.ID = [1] And b.��ID Is Null And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��һ�׶�", lng�׶�ID, lng��ǰ�׶η�֧ID)
    
    If rsTmp.RecordCount > 0 Then GetNextPhase = Val(rsTmp!ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMustDay(ByVal lng����·��Id As Long, ByVal lng��ǰ���� As Long, Optional ByVal blnIsNotMinus As Boolean, _
        Optional ByVal lng�ϲ�·����¼ID As Long) As Long
'���ܣ���ȡ����·��ִ�������ϵĵ�ǰ����(=��ǰʵ������-�����ӳٵ�����+��ǰ����(�п���һ����ǰ����))
'������blnIsNotMinus=�Ƿ񲻼�ȥ�ӳ�ʱ�䣨����ʱ��ǰ������
'      lng�׶�ID��lng����=����Ǻϲ�·������ӵ������֮������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng�ӳ����� As Long
    Dim lng��ǰ���� As Long
    Dim i As Long
    Dim lng�׶�ʵ������ As Long
    Dim lng�׶ο�ʼ���� As Long
    Dim byt��ǰ���� As Byte
    
    On Error GoTo errH
    If lng�ϲ�·����¼ID <> 0 Then
        strSql = "Select Max(Decode(g.ʱ�����, 1, 1, 2, 2, 0)) As �׶��Ƿ���ǰ, c.��ʼ����, Nvl(c.��������, c.��ʼ����) as ��������, Sum(Decode(g.ʱ�����, -1, 1, 0)) As �׶��Ӻ�����," & vbNewLine & _
                "       Count(*) As �׶�ʵ������" & vbNewLine & _
                "From ���˺ϲ�·������ A, �ٴ�·����֧ B, �ٴ�·���׶� C, �ٴ�·���׶� D, �ٴ�·���׶� E, �ٴ�·���׶� F,����·������ G" & vbNewLine & _
                "Where a.�ϲ�·���׶�id = c.Id And c.��id = d.Id(+) And b.Id(+) = c.��֧id And b.ǰһ�׶�id = e.Id(+) And e.��id = f.Id(+) And a.·����¼id=g.·����¼id And a.�׶�id=g.�׶�id and a.����=g.����" & vbNewLine & _
                "      And a.�ϲ�·����¼id = [1]" & vbNewLine & _
                "Group By c.��ʼ����, Nvl(c.��������, c.��ʼ����), a.�׶�id, c.��֧id, d.���, c.���, f.���, e.���" & vbNewLine & _
                "Order By Decode(c.��֧id, Null, Nvl(d.���, c.���), Nvl(d.���, c.���) + Nvl(f.���, e.���))"
   
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDay", lng�ϲ�·����¼ID)
    Else
       strSql = "Select Max(Decode(a.ʱ�����, 1, 1, 2, 2, 0)) As �׶��Ƿ���ǰ, c.��ʼ����, Nvl(c.��������, c.��ʼ����) as ��������, Sum(Decode(a.ʱ�����, -1, 1, 0)) As �׶��Ӻ�����," & vbNewLine & _
                "       Count(1) As �׶�ʵ������" & vbNewLine & _
                "From ����·������ A, �ٴ�·����֧ B, �ٴ�·���׶� C, �ٴ�·���׶� D, �ٴ�·���׶� E, �ٴ�·���׶� F" & vbNewLine & _
                "Where a.�׶�id = c.Id And c.��id = d.Id(+) And b.Id(+) = c.��֧id And b.ǰһ�׶�id = e.Id(+) And e.��id = f.Id(+) And" & vbNewLine & _
                "      a.·����¼id = [1] " & vbNewLine & _
                "Group By c.��ʼ����, Nvl(c.��������, c.��ʼ����), a.�׶�id, c.��֧id, d.���, c.���, f.���, e.���" & vbNewLine & _
                "Order By Decode(c.��֧id, Null, Nvl(d.���, c.���), Nvl(d.���, c.���) + Nvl(f.���, e.���))"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDay", lng����·��Id)
    End If
    For i = 0 To rsTmp.RecordCount - 1
        '�ӳ�����
        lng�ӳ����� = lng�ӳ����� + Val(rsTmp!�׶��Ӻ����� & "")
        '��ǰ����
        If Val(rsTmp!�׶��Ƿ���ǰ & "") = 1 Or Val(rsTmp!�׶��Ƿ���ǰ & "") = 2 Then
            '�ϲ�·������ʼ�׶ο�������ǰ����Ҫ·����ǰ�������ٵ���ϲ�·����,��һ���׶ξ���Ҫ������ǰ����(�����Ѿ������˺ϲ�·������ӵ�һ���׶ο�ʼ����������ж���ʱ��Ч����������Ҫ�ϲ�·���Ӻ��濪ʼ����ɿ���)
'            If i = 0 And rsTmp!��ʼ���� & "" <> "1" Then
'                lng��ǰ���� = Val(rsTmp!��ʼ���� & "") - 1
'                rsTmp.MoveNext
'            Else
                '���һ���׶�����ǰ�����1�죬��Ϊ����֪�������ѡ��һ���׶�
                If i = rsTmp.RecordCount - 1 Or rsTmp!��ʼ���� & "" = rsTmp!�������� & "" Then
                    If Val(rsTmp!�׶��Ƿ���ǰ & "") = 1 Then
                        lng��ǰ���� = lng��ǰ���� + 1
                    ElseIf Val(rsTmp!�׶��Ƿ���ǰ & "") = 2 Then
                        '��һ�׶���ǰ������,��ʱ����Ҫ����һ�׶���ǰ�����족�ٶ����һ��
                    End If
                    rsTmp.MoveNext
                Else
                    '�ȼ�¼�½׶�ʵ�������Ϳ�ʼ����
                    lng�׶ο�ʼ���� = Val(rsTmp!��ʼ���� & "")
                    lng�׶�ʵ������ = Val(rsTmp!�׶�ʵ������ & "")
                    byt��ǰ���� = Val(rsTmp!�׶��Ƿ���ǰ & "")
                    rsTmp.MoveNext
                    lng��ǰ���� = lng��ǰ���� + (Val(rsTmp!��ʼ���� & "") - lng�׶ο�ʼ���� - lng�׶�ʵ������ + IIf(byt��ǰ���� = 2, 0, 1))
                End If
'            End If
        Else
            rsTmp.MoveNext
        End If
    Next
    
    GetMustDay = lng��ǰ���� - IIf(blnIsNotMinus, 0, lng�ӳ�����) + lng��ǰ����
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPhaseNO(ByVal lng�׶�ID As Long) As Long
'���ܣ���ȡָ���׶ε����(����ý׶��Ƿ�֧����ȡ���׶ε����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ��ID From �ٴ�·���׶� Where id = [1] And ��ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then lng�׶�ID = Val(rsTmp!��ID)
    
    strSql = "Select ��� From �ٴ�·���׶� Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then GetPhaseNO = Val(rsTmp!���)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastPhaseNO(ByVal lng����·��Id As Long, ByVal lng·��ID As Long)
'���ܣ���ȡ����ָ��·�����һ���׶ε����
Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Max(Nvl(c.���, b.���)) ���" & vbNewLine & _
            "From ����·��ִ�� A, �ٴ�·���׶� B, �ٴ�·���׶� C" & vbNewLine & _
            "Where a.·����¼id = [1] And a.�׶�id = b.Id And b.·��id = [2] And b.��id = c.Id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng����·��Id, lng·��ID)
    
    GetLastPhaseNO = Val("" & rsTmp!���)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExportPathToXML(ByVal lng·��ID As Long, ByVal int�汾�� As Integer, ByVal strFile As String) As Boolean
'���ܣ������ٴ�·����XML�ļ�
'������strFile=����·�����ļ���
'˵������������·����Ϣ��ָ���汾����Ϣ
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    Dim xPI As IXMLDOMProcessingInstruction
    
    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsItemAdvice As ADODB.Recordset
    Dim rsItemEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Set xPath = New DOMDocument
    
    'ע��
    xPath.appendChild xPath.createComment(gstrSysName & "  ����Ա:" & UserInfo.���� & ",����:" & UserInfo.������ & ",ʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
    
    '�����
    Set xRoot = xPath.createElement("ClinicalPathways")
    Set xPath.documentElement = xRoot
    Call xRoot.setAttribute("ID", lng·��ID)
    Call xRoot.setAttribute("Version", int�汾��)

    '�ٴ�·����Ϣ
    strSql = "Select A.����,A.����,A.����,A.ͨ��,A.���°汾,A.��������," & _
        " A.���ò���,A.�����Ա�,A.��������,A.˵��,B.��׼סԺ��,B.��׼����," & _
        " B.�汾˵��,B.������,B.����ʱ��,B.�����,B.���ʱ��,B.ͣ����,B.ͣ��ʱ��,A.ȷ������,A.����·������,A.����" & _
        " From �ٴ�·��Ŀ¼ A,�ٴ�·���汾 B Where A.ID=B.·��ID And A.ID=[1] And B.�汾��=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    Set xNode = CreateNode(1, xRoot, "PathInfo", NODE_ELEMENT, "")
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "ͨ��", , NVL(rsTmp!ͨ��)
        CreateNode 2, xNode, "���°汾", , NVL(rsTmp!���°汾)
        CreateNode 2, xNode, "��������", , NVL(rsTmp!��������)
        CreateNode 2, xNode, "���ò���", , NVL(rsTmp!���ò���)
        CreateNode 2, xNode, "�����Ա�", , NVL(rsTmp!�����Ա�)
        CreateNode 2, xNode, "��������", , NVL(rsTmp!��������)
        CreateNode 2, xNode, "˵��", , NVL(rsTmp!˵��)
        CreateNode 2, xNode, "��׼סԺ��", , NVL(rsTmp!��׼סԺ��)
        CreateNode 2, xNode, "��׼����", , NVL(rsTmp!��׼����)
        CreateNode 2, xNode, "�汾˵��", , NVL(rsTmp!�汾˵��)
        CreateNode 2, xNode, "������", , NVL(rsTmp!������)
        CreateNode 2, xNode, "����ʱ��", , Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "�����", , NVL(rsTmp!�����)
        CreateNode 2, xNode, "���ʱ��", , Format(NVL(rsTmp!���ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "ͣ����", , NVL(rsTmp!ͣ����)
        CreateNode 2, xNode, "ͣ��ʱ��", , Format(NVL(rsTmp!ͣ��ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "ȷ������", , NVL(rsTmp!ȷ������)
        CreateNode 2, xNode, "����·������", , NVL(rsTmp!����·������)
        CreateNode 2, xNode, "����", , NVL(rsTmp!����, 0)
    
    '�ٴ�·������
    strSql = "Select B.ID,B.����,B.���� From �ٴ�·������ A,���ű� B Where A.·��ID=[1] And A.����ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDepts", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDept", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "����ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "����", , rsTmp!����
                CreateNode 3, xSubNode1, "����", , rsTmp!����
            rsTmp.MoveNext
        Loop
    End If
    
    '�ٴ�·������
    strSql = "Select A.����ID,B.���� as ������,B.���� as ������," & _
        " A.���ID,C.���� as �����,C.���� as �����, a.���� as ����" & _
        " From �ٴ�·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
        " Where Nvl(A.����ID,0)=B.ID(+) And Nvl(A.���ID,0)=C.ID(+) And A.·��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDiseases", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDisease", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "����ID", , NVL(rsTmp!����id)
                CreateNode 3, xSubNode1, "������", , NVL(rsTmp!������)
                CreateNode 3, xSubNode1, "������", , NVL(rsTmp!������)
                CreateNode 3, xSubNode1, "���ID", , NVL(rsTmp!���id)
                CreateNode 3, xSubNode1, "�����", , NVL(rsTmp!�����)
                CreateNode 3, xSubNode1, "�����", , NVL(rsTmp!�����)
                CreateNode 3, xSubNode1, "����", , NVL(rsTmp!����)
            rsTmp.MoveNext
        Loop
    End If
    
    '�ٴ�·����֧
    strSql = "Select ID,·��ID,�汾��,����,˵��,ǰһ�׶�ID,��׼סԺ��,��׼����,������,����ʱ�� From �ٴ�·����֧ where ·��ID=[1] AND �汾��=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathBranchs", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathBranch", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "·��ID", , rsTmp!·��ID
                CreateNode 3, xSubNode1, "�汾��", , rsTmp!�汾��
                CreateNode 3, xSubNode1, "����", , NVL(rsTmp!����)
                CreateNode 3, xSubNode1, "˵��", , NVL(rsTmp!˵��)
                CreateNode 3, xSubNode1, "ǰһ�׶�ID", , NVL(rsTmp!ǰһ�׶�ID)
                CreateNode 3, xSubNode1, "��׼סԺ��", , NVL(rsTmp!��׼סԺ��)
                CreateNode 3, xSubNode1, "��׼����", , NVL(rsTmp!��׼����)
                CreateNode 3, xSubNode1, "������", , NVL(rsTmp!������)
                CreateNode 3, xSubNode1, "����ʱ��", , Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            rsTmp.MoveNext
        Loop
    End If
    
    '��������
    strSql = "Select B.��������,B.�׶�ID,A.ID,A.����ָ��,A.ָ������,A.ָ����,b.��֧ID" & _
        " From ·������ָ�� A,�ٴ�·������ B" & _
        " Where A.����ID=B.ID And B.·��ID=[1] And �汾��=[2]" & _
        " Order by B.��������,B.�׶�ID,A.���"
    Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select B.��������,B.�׶�ID,A.ָ��ID,A.��ĿID,A.��ϵʽ,A.����ֵ,A.�������,b.��֧ID" & _
        " From ·���������� A,�ٴ�·������ B" & _
        " Where A.����ID=B.ID And B.·��ID=[1] And �汾��=[2]" & _
        " Order by B.��������,B.�׶�ID"
    Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    rsEvalMark.Filter = "��������=1"
    rsEvalCond.Filter = "��������=1"
    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
        Set xNode = CreateNode(1, xRoot, "ImportEval", NODE_ELEMENT, "")
            If Not rsEvalMark.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Marks", NODE_ELEMENT, "")
                Do While Not rsEvalMark.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Mark", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ID", , rsEvalMark!ID
                        CreateNode 4, xSubNode2, "����ָ��", , rsEvalMark!����ָ��
                        CreateNode 4, xSubNode2, "ָ������", , rsEvalMark!ָ������
                        CreateNode 4, xSubNode2, "ָ����", , rsEvalMark!ָ����
                    rsEvalMark.MoveNext
                Loop
            End If
            If Not rsEvalCond.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Conditions", NODE_ELEMENT, "")
                Do While Not rsEvalCond.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ָ��ID", , rsEvalCond!ָ��ID
                        CreateNode 4, xSubNode2, "��ϵʽ", , rsEvalCond!��ϵʽ
                        CreateNode 4, xSubNode2, "����ֵ", , rsEvalCond!����ֵ
                        CreateNode 4, xSubNode2, "�������", , rsEvalCond!�������
                    rsEvalCond.MoveNext
                Loop
            End If
    End If
    
    '·��ҽ������
    strSql = "Select Distinct A.ID,A.���ID,A.���,A.��Ч,A.������ĿID,D.���� as ���Ʊ���,D.���� as ��������," & _
        " A.�շ�ϸĿID,E.���� as �շѱ���,E.���� as �շ�����,A.ҽ������,A.��������,A.�ܸ�����," & _
        " A.�걾��λ,A.��鷽��,A.ҽ������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ," & _
        " A.ִ������,A.ִ�п���ID,F.���� as ִ�п�����,F.���� as ִ�п�����,A.ʱ�䷽��,A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�䷽ID,A.�����ĿID" & _
        " From ·��ҽ������ A,�ٴ�·��ҽ�� B,�ٴ�·����Ŀ C,������ĿĿ¼ D,�շ���ĿĿ¼ E,���ű� F" & _
        " Where A.ID=B.ҽ������ID And B.·����ĿID=C.ID And C.·��ID=[1] And C.�汾��=[2]" & _
        " And Nvl(A.������ĿID,0)=D.ID(+) And Nvl(A.�շ�ϸĿID,0)=E.ID(+) And Nvl(A.ִ�п���ID,0)=F.ID(+)" & _
        " Order by A.���,A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathAdvices", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathAdvice", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "���ID", , NVL(rsTmp!���id)
                CreateNode 3, xSubNode1, "���", , rsTmp!���
                CreateNode 3, xSubNode1, "��Ч", , rsTmp!��Ч
                CreateNode 3, xSubNode1, "������ĿID", , NVL(rsTmp!������ĿID)
                CreateNode 3, xSubNode1, "���Ʊ���", , NVL(rsTmp!���Ʊ���)
                CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
                CreateNode 3, xSubNode1, "�շ�ϸĿID", , NVL(rsTmp!�շ�ϸĿID)
                CreateNode 3, xSubNode1, "�շѱ���", , NVL(rsTmp!�շѱ���)
                CreateNode 3, xSubNode1, "�շ�����", , NVL(rsTmp!�շ�����)
                CreateNode 3, xSubNode1, "ҽ������", , NVL(rsTmp!ҽ������)
                CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
                CreateNode 3, xSubNode1, "�ܸ�����", , NVL(rsTmp!�ܸ�����)
                CreateNode 3, xSubNode1, "�걾��λ", , NVL(rsTmp!�걾��λ)
                CreateNode 3, xSubNode1, "��鷽��", , NVL(rsTmp!��鷽��)
                CreateNode 3, xSubNode1, "ҽ������", , NVL(rsTmp!ҽ������)
                CreateNode 3, xSubNode1, "ִ��Ƶ��", , NVL(rsTmp!ִ��Ƶ��)
                CreateNode 3, xSubNode1, "Ƶ�ʴ���", , NVL(rsTmp!Ƶ�ʴ���)
                CreateNode 3, xSubNode1, "Ƶ�ʼ��", , NVL(rsTmp!Ƶ�ʼ��)
                CreateNode 3, xSubNode1, "�����λ", , NVL(rsTmp!�����λ)
                CreateNode 3, xSubNode1, "ִ������", , NVL(rsTmp!ִ������)
                CreateNode 3, xSubNode1, "ִ�п���ID", , NVL(rsTmp!ִ�п���ID)
                CreateNode 3, xSubNode1, "ִ�п�����", , NVL(rsTmp!ִ�п�����)
                CreateNode 3, xSubNode1, "ִ�п�����", , NVL(rsTmp!ִ�п�����)
                CreateNode 3, xSubNode1, "ʱ�䷽��", , NVL(rsTmp!ʱ�䷽��)
                CreateNode 3, xSubNode1, "�Ƿ�ȱʡ", , NVL(rsTmp!�Ƿ�ȱʡ, 0)
                CreateNode 3, xSubNode1, "�Ƿ�ѡ", , NVL(rsTmp!�Ƿ�ѡ, 0)
                CreateNode 3, xSubNode1, "�䷽ID", , NVL(rsTmp!�䷽ID)
                CreateNode 3, xSubNode1, "�����ĿID", , NVL(rsTmp!�����ĿID)
            rsTmp.MoveNext
        Loop
    End If
    
    '�ٴ�·������
    strSql = "Select ����,��֧ID From �ٴ�·������ Where ·��ID=[1] And �汾��=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    Set xNode = CreateNode(1, xRoot, "PathCategorys", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        Set xSubNode1 = CreateNode(2, xNode, "PathCategory", NODE_ELEMENT, NVL(rsTmp!����))
        CreateNode 2, xSubNode1, "����", NODE_ELEMENT, NVL(rsTmp!����)
        CreateNode 2, xSubNode1, "��֧ID", NODE_ELEMENT, NVL(rsTmp!��֧ID)
        rsTmp.MoveNext
    Loop
    
    '�ٴ�·���׶�/��Ŀ
    strSql = "Select ID,Nvl(��ID,0) as ��ID,���,����,��ʼ����,��������,��־,����,˵��,��֧ID" & _
        " From �ٴ�·���׶� Where ·��ID=[1] And �汾��=[2] Order by Nvl(��ID,0) Desc,���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select ID,�׶�ID,����,��Ŀ���,��Ŀ����,ִ�з�ʽ,ִ����,������,��Ŀ���,ͼ��ID,����Ҫ��,��֧ID" & _
        " From �ٴ�·����Ŀ Where ·��ID=[1] And �汾��=[2] Order by �׶�ID,����,��Ŀ���"
    Set rsItem = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select A.·����ĿID,A.ҽ������ID From �ٴ�·��ҽ�� A,�ٴ�·����Ŀ B" & _
        " Where A.·����ĿID=B.ID And B.·��ID=[1] And �汾��=[2]"
    Set rsItemAdvice = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select A.��ĿID,A.�ļ�ID,C.���,C.���� From �ٴ�·������ A,�ٴ�·����Ŀ B,�����ļ��б� C" & _
        " Where A.��ĿID=B.ID And A.�ļ�ID=C.ID And B.·��ID=[1] And �汾��=[2]"
    Set rsItemEPR = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng·��ID, int�汾��)
    
    Set rsClone = rsTmp.Clone: rsTmp.Filter = "��ID=0"
    
    Set xNode = CreateNode(1, xRoot, "PathTimeSteps", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        'ȱʡ��֧
        Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
            CreateNode 3, xSubNode1, "ID", , rsTmp!ID
            CreateNode 3, xSubNode1, "��ID", , ""
            CreateNode 3, xSubNode1, "���", , rsTmp!���
            CreateNode 3, xSubNode1, "����", , rsTmp!����
            CreateNode 3, xSubNode1, "��ʼ����", , NVL(rsTmp!��ʼ����)
            CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
            CreateNode 3, xSubNode1, "��־", , NVL(rsTmp!��־)
            CreateNode 3, xSubNode1, "˵��", , NVL(rsTmp!˵��)
            CreateNode 3, xSubNode1, "����", , NVL(rsTmp!����)
            CreateNode 3, xSubNode1, "��֧ID", , NVL(rsTmp!��֧ID)
            
            '�׶ε���Ŀ
            rsItem.Filter = "�׶�ID=" & rsTmp!ID
            Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
            Do While Not rsItem.EOF
                Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                    CreateNode 5, xSubNode3, "ID", , rsItem!ID
                    CreateNode 5, xSubNode3, "����", , rsItem!����
                    CreateNode 5, xSubNode3, "��Ŀ���", , rsItem!��Ŀ���
                    CreateNode 5, xSubNode3, "��Ŀ����", , rsItem!��Ŀ����
                    CreateNode 5, xSubNode3, "ִ�з�ʽ", , NVL(rsItem!ִ�з�ʽ)
                    CreateNode 5, xSubNode3, "ִ����", , NVL(rsItem!ִ����)
                    CreateNode 5, xSubNode3, "������", , NVL(rsItem!������, 1)
                    CreateNode 5, xSubNode3, "��Ŀ���", , NVL(rsItem!��Ŀ���)
                    CreateNode 5, xSubNode3, "ͼ��ID", , NVL(rsItem!ͼ��ID)
                    CreateNode 5, xSubNode3, "����Ҫ��", , NVL(rsItem!����Ҫ��, 0)
                    CreateNode 5, xSubNode3, "��֧ID", , NVL(rsItem!��֧ID)
                    
                    '��Ŀ��Ӧ��ҽ��
                    rsItemAdvice.Filter = "·����ĿID=" & rsItem!ID
                    If Not rsItemAdvice.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                        Do While Not rsItemAdvice.EOF
                            CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!ҽ������ID
                            rsItemAdvice.MoveNext
                        Loop
                    End If
                    '��Ŀ��Ӧ�Ĳ���
                    rsItemEPR.Filter = "��ĿID=" & rsItem!ID
                    If Not rsItemEPR.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                        Do While Not rsItemEPR.EOF
                            Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                CreateNode 7, xSubNode5, "�ļ�ID", , rsItemEPR!�ļ�ID
                                CreateNode 7, xSubNode5, "�ļ����", , rsItemEPR!���
                                CreateNode 7, xSubNode5, "�ļ�����", , rsItemEPR!����
                            rsItemEPR.MoveNext
                        Loop
                    End If
                    
                rsItem.MoveNext
            Loop
        
            '�׶ε�����
            rsEvalMark.Filter = "��������=2 And �׶�ID=" & rsTmp!ID
            rsEvalCond.Filter = "��������=2 And �׶�ID=" & rsTmp!ID
            If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                    If Not rsEvalMark.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                        Do While Not rsEvalMark.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                CreateNode 6, xSubNode4, "����ָ��", , rsEvalMark!����ָ��
                                CreateNode 6, xSubNode4, "ָ������", , rsEvalMark!ָ������
                                CreateNode 6, xSubNode4, "ָ����", , rsEvalMark!ָ����
                                CreateNode 6, xSubNode4, "��֧ID", , NVL(rsEvalMark!��֧ID)
                            rsEvalMark.MoveNext
                        Loop
                    End If
                    If Not rsEvalCond.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                        Do While Not rsEvalCond.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ָ��ID", , NVL(rsEvalCond!ָ��ID)
                                CreateNode 6, xSubNode4, "��ĿID", , NVL(rsEvalCond!��ĿID)
                                CreateNode 6, xSubNode4, "��ϵʽ", , rsEvalCond!��ϵʽ
                                CreateNode 6, xSubNode4, "����ֵ", , rsEvalCond!����ֵ
                                CreateNode 6, xSubNode4, "�������", , rsEvalCond!�������
                                CreateNode 6, xSubNode4, "��֧ID", , NVL(rsEvalCond!��֧ID)
                            rsEvalCond.MoveNext
                        Loop
                    End If
            End If
        
        '��ѡ��֧
        rsClone.Filter = "��ID=" & rsTmp!ID
        If Not rsClone.EOF Then
            Do While Not rsClone.EOF
                Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
                    CreateNode 3, xSubNode1, "ID", , rsClone!ID
                    CreateNode 3, xSubNode1, "��ID", , rsClone!��ID
                    CreateNode 3, xSubNode1, "���", , rsClone!���
                    CreateNode 3, xSubNode1, "����", , rsClone!����
                    CreateNode 3, xSubNode1, "��ʼ����", , NVL(rsClone!��ʼ����)
                    CreateNode 3, xSubNode1, "��������", , NVL(rsClone!��������)
                    CreateNode 3, xSubNode1, "��־", , NVL(rsClone!��־)
                    CreateNode 3, xSubNode1, "˵��", , NVL(rsClone!˵��)
                    CreateNode 3, xSubNode1, "��֧ID", , NVL(rsClone!��֧ID)
                
                    '�׶ε���Ŀ
                    rsItem.Filter = "�׶�ID=" & rsClone!ID
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
                    Do While Not rsItem.EOF
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                            CreateNode 5, xSubNode3, "ID", , rsItem!ID
                            CreateNode 5, xSubNode3, "����", , rsItem!����
                            CreateNode 5, xSubNode3, "��Ŀ���", , rsItem!��Ŀ���
                            CreateNode 5, xSubNode3, "��Ŀ����", , rsItem!��Ŀ����
                            CreateNode 5, xSubNode3, "ִ�з�ʽ", , NVL(rsItem!ִ�з�ʽ)
                            CreateNode 5, xSubNode3, "ִ����", , NVL(rsItem!ִ����)
                            CreateNode 5, xSubNode3, "������", , NVL(rsItem!������, 1)
                            CreateNode 5, xSubNode3, "��Ŀ���", , NVL(rsItem!��Ŀ���)
                            CreateNode 5, xSubNode3, "ͼ��ID", , NVL(rsItem!ͼ��ID)
                            CreateNode 5, xSubNode3, "��֧ID", , NVL(rsItem!��֧ID)
                            
                            '��Ŀ��Ӧ��ҽ��
                            rsItemAdvice.Filter = "·����ĿID=" & rsItem!ID
                            If Not rsItemAdvice.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                                Do While Not rsItemAdvice.EOF
                                    CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!ҽ������ID
                                    rsItemAdvice.MoveNext
                                Loop
                            End If
                            '��Ŀ��Ӧ�Ĳ���
                            rsItemEPR.Filter = "��ĿID=" & rsItem!ID
                            If Not rsItemEPR.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                                Do While Not rsItemEPR.EOF
                                    Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                        CreateNode 7, xSubNode5, "�ļ�ID", , rsItemEPR!�ļ�ID
                                        CreateNode 7, xSubNode5, "�ļ����", , rsItemEPR!���
                                        CreateNode 7, xSubNode5, "�ļ�����", , rsItemEPR!����
                                    rsItemEPR.MoveNext
                                Loop
                            End If
                            
                        rsItem.MoveNext
                    Loop
                    
                    '�׶ε�����
                    rsEvalMark.Filter = "��������=2 And �׶�ID=" & rsClone!ID
                    rsEvalCond.Filter = "��������=2 And �׶�ID=" & rsClone!ID
                    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                        Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                            If Not rsEvalMark.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                                Do While Not rsEvalMark.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                        CreateNode 6, xSubNode4, "����ָ��", , rsEvalMark!����ָ��
                                        CreateNode 6, xSubNode4, "ָ������", , rsEvalMark!ָ������
                                        CreateNode 6, xSubNode4, "ָ����", , rsEvalMark!ָ����
                                        CreateNode 6, xSubNode4, "��֧ID", , NVL(rsEvalMark!��֧ID)
                                    rsEvalMark.MoveNext
                                Loop
                            End If
                            If Not rsEvalCond.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                                Do While Not rsEvalCond.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ָ��ID", , NVL(rsEvalCond!ָ��ID)
                                        CreateNode 6, xSubNode4, "��ĿID", , NVL(rsEvalCond!��ĿID)
                                        CreateNode 6, xSubNode4, "��ϵʽ", , rsEvalCond!��ϵʽ
                                        CreateNode 6, xSubNode4, "����ֵ", , rsEvalCond!����ֵ
                                        CreateNode 6, xSubNode4, "�������", , rsEvalCond!�������
                                        CreateNode 6, xSubNode4, "��֧ID", , NVL(rsEvalCond!��֧ID)
                                    rsEvalCond.MoveNext
                                Loop
                            End If
                    End If
                
                rsClone.MoveNext
            Loop
        End If
        
        rsTmp.MoveNext
    Loop
    
    'XML��Ϣ
    Set xPI = xPath.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call xPath.insertBefore(xPI, xPath.childNodes(0))
    
    '������ļ�
    xPath.Save strFile
    Set xPath = Nothing
    
    ExportPathToXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, ByVal SubNodeName As String, Optional ByVal DefaultValue As String) As String
    Dim NodeTMP As IXMLDOMNode
    
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
End Function

Public Function ImportPathFromXML(ByVal strFile As String, _
    Optional ByVal lng·��ID As Long, Optional ByVal int�汾�� As Integer, _
    Optional ByVal intLimit As Integer, Optional ByRef blnLimit As Boolean) As Boolean
'���ܣ�����ָ�����ٴ�·��XML�ļ�
'������lng·��ID,int�汾��=���ָ������ֻ����汾��ز�����Ϣ�����û��ָ��������ݸ���XML�е���Ϣ����·������������ȫ����
'      intLimit=�������Ƶ����·������,Ϊ0��ʾ������
'      blnLimit=�Ƿ���������·�����������Ƶ���ʧ��
    Dim rsTmp As ADODB.Recordset
    Dim rsIcon As ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim arrSQL As Variant, strSql As String
    Dim colItemID As Collection
    Dim colStepID As Collection
    Dim colMarkID As Collection
    Dim colAdviceID As Collection
    Dim colAdviceOriginalID As Collection
    Dim colBranchID As Collection
    Dim colPreID As Collection
    
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    
    Dim str���� As String, lng�׶�ID As Long
    Dim strValue As String, strTemp1 As String
    Dim strTemp2 As String, strTemp3 As String
    Dim blnDo As Boolean, blnTran As Boolean
    Dim strValueTurn As String
    Dim strValueTurn1 As String
    Dim i As Long, k As Long, n As Long, m As Long
    Dim strTmp As String
    Dim strPreStep As String
    Dim strtemp4 As String
    Dim lngType As Long
    Dim strImportRef As String
    Dim lng������ As Long '��¼ͬһ·����Ŀҽ���ĵ���״̬0��ȫ��δ���룬1��ȫ�����룬2�����ֵ���
    Dim lngCount As Long, str��IDs As String, arrID As Variant, lng��ID As Long, strFilter As String
    Dim lng��ĿID As Long
    
    On Error GoTo errH
    
    rsAdvice.Fields.Append "ID", adBigInt
    rsAdvice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "����ο�", adVarChar, 200, adFldIsNullable
    rsAdvice.Fields.Append "��ĿID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "����״̬", adInteger
    
    rsAdvice.CursorLocation = adUseClient
    rsAdvice.LockType = adLockOptimistic
    rsAdvice.CursorType = adOpenStatic
    rsAdvice.Open
    
    blnLimit = False
    
    Set xPath = New DOMDocument
    xPath.Load strFile
    
    '����������κ�Ԫ�أ����˳�
    If xPath.documentElement Is Nothing Then
        Set xPath = Nothing
        Screen.MousePointer = 0
        Exit Function
    End If
    
    arrSQL = Array()
    
    '��ȡXML����
    Set xRoot = xPath.selectSingleNode("ClinicalPathways")
    Set xNode = xRoot.selectSingleNode("PathInfo")
    If lng·��ID = 0 Then
        '��ȡӦ�ÿ��ҵ����
        strTemp1 = ""
        If Val(GetNodeValue(xNode, "ͨ��")) = 2 Then
            Set xSubNode1 = xRoot.selectSingleNode("PathDepts")
            If Not xSubNode1 Is Nothing Then
                strSql = "Select A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� C" & _
                    " Where A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
                    " Order by A.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML")
                
                For Each xSubNode2 In xSubNode1.childNodes
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode2, "����") & "' And ����='" & GetNodeValue(xSubNode2, "����") & "'"
                    If Not rsTmp.EOF Then strTemp1 = strTemp1 & "," & rsTmp!ID
                Next
            
                strTemp1 = Mid(strTemp1, 2)
            End If
        End If
        
        '��ȡӦ�ü��������
        strValue = ""
        Set xSubNode1 = xRoot.selectSingleNode("PathDiseases")
        If Not xSubNode1 Is Nothing Then
            strTemp2 = "": strTemp3 = ""
            For Each xSubNode2 In xSubNode1.childNodes
                If Val(GetNodeValue(xSubNode2, "����")) = 0 Then
                    If Val(GetNodeValue(xSubNode2, "����ID")) <> 0 Then
                        strSql = "Select ID From ��������Ŀ¼ Where ����=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "������"), GetNodeValue(xSubNode2, "������"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & "," & rsTmp!ID
                    ElseIf Val(GetNodeValue(xSubNode2, "���ID")) <> 0 Then
                        strSql = "Select ID From �������Ŀ¼ Where ����=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "�����"), GetNodeValue(xSubNode2, "�����"))
                        If Not rsTmp.EOF Then strTemp3 = strTemp3 & "," & rsTmp!ID
                    End If
                Else
                    If Val(GetNodeValue(xSubNode2, "����ID")) <> 0 Then
                        strSql = "Select ID From ��������Ŀ¼ Where ����=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "������"), GetNodeValue(xSubNode2, "������"))
                        If Not rsTmp.EOF Then strValueTurn = strValueTurn & "," & rsTmp!ID
                    ElseIf Val(GetNodeValue(xSubNode2, "���ID")) <> 0 Then
                        strSql = "Select ID From �������Ŀ¼ Where ����=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "�����"), GetNodeValue(xSubNode2, "�����"))
                        If Not rsTmp.EOF Then strValueTurn1 = strValueTurn1 & "," & rsTmp!ID
                    End If
                End If
            Next
            If strTemp2 <> "" Or strTemp3 <> "" Then
                strValue = Mid(strTemp2, 2) & ";" & Mid(strTemp3, 2)
            End If
            If strValueTurn <> "" Or strValueTurn1 <> "" Then
                strValueTurn = Mid(strValueTurn, 2) & ";" & Mid(strValueTurn1, 2)
            End If
        End If
        
        '�����ٴ�·����Ϣ
        strSql = "Select ID,����,���°汾 From �ٴ�·��Ŀ¼ Where ����=[1] And ����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xNode, "����"), GetNodeValue(xNode, "����"))
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If Not rsTmp.EOF Then
            '�����汾���߸��ǰ汾
            lng·��ID = rsTmp!ID
            int�汾�� = NVL(rsTmp!���°汾, 0) + 1 '���ܸ���δ��˰汾
            str���� = rsTmp!����
            arrSQL(UBound(arrSQL)) = "zl_�ٴ�·��Ŀ¼_Update(" & _
                lng·��ID & ",'" & GetNodeValue(xNode, "����") & "','" & str���� & "'," & _
                "'" & GetNodeValue(xNode, "����") & "','" & GetNodeValue(xNode, "˵��") & "'," & _
                "'" & GetNodeValue(xNode, "��������") & "','" & GetNodeValue(xNode, "���ò���") & "'," & _
                Val(GetNodeValue(xNode, "�����Ա�")) & ",'" & GetNodeValue(xNode, "��������") & "'," & _
                Val(GetNodeValue(xNode, "ͨ��")) & ",'" & strTemp1 & "','" & strValue & "'," & Val(GetNodeValue(xNode, "ȷ������")) & ",'" _
                & strValueTurn & "'," & Val(GetNodeValue(xNode, "����·������")) & "," & Val(GetNodeValue(xNode, "����")) & ")"
        
        Else
            '�����Ȩ����
            If intLimit > 0 Then
                strSql = "Select Nvl(Count(*),0) as ���� From �ٴ�·��Ŀ¼"
                Set rsTmp = New ADODB.Recordset
                Call zlDatabase.OpenRecordset(rsTmp, strSql, "ImportPathFromXML")
                If rsTmp!���� >= intLimit Then
                    blnLimit = True
                    Set xPath = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            
            '����·��
            lng·��ID = zlDatabase.GetNextId("�ٴ�·��Ŀ¼")
            int�汾�� = 1
            str���� = GetNextCode(GetNodeValue(xNode, "����"))
            arrSQL(UBound(arrSQL)) = "zl_�ٴ�·��Ŀ¼_Insert(" & _
                "'" & GetNodeValue(xNode, "����") & "','" & str���� & "'," & _
                "'" & GetNodeValue(xNode, "����") & "','" & GetNodeValue(xNode, "˵��") & "'," & _
                "'" & GetNodeValue(xNode, "��������") & "','" & GetNodeValue(xNode, "���ò���") & "'," & _
                Val(GetNodeValue(xNode, "�����Ա�")) & ",'" & GetNodeValue(xNode, "��������") & "'," & _
                Val(GetNodeValue(xNode, "ͨ��")) & ",'" & strTemp1 & "','" & strValue & "'," & lng·��ID & "," & Val(GetNodeValue(xNode, "ȷ������")) & ",'" & _
                strValueTurn & "'," & Val(GetNodeValue(xNode, "����·������")) & "," & Val(GetNodeValue(xNode, "����")) & ")"
        End If
    Else
        strSql = "Select ���� From �ٴ�·��Ŀ¼ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", lng·��ID)
        If rsTmp.RecordCount > 0 Then lngType = Val(rsTmp!���� & "")
    End If
    
    'ɾ���汾��ص����ݣ����²���
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Delete(" & lng·��ID & "," & int�汾�� & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Update(" & lng·��ID & "," & int�汾�� & "," & _
        "'" & GetNodeValue(xNode, "��׼סԺ��") & "','" & GetNodeValue(xNode, "��׼����") & "'," & _
        "'" & GetNodeValue(xNode, "�汾˵��") & "')"
    
    '��������
    Set xNode = xRoot.selectSingleNode("ImportEval")
    If Not xNode Is Nothing Then
        Set xSubNode1 = xNode.selectSingleNode("Marks")
        If Not xSubNode1 Is Nothing Then
            k = 1
            Set colItemID = New Collection
            For Each xSubNode2 In xSubNode1.childNodes
                strValue = zlDatabase.GetNextId("·������ָ��")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode2, "ID")
                            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_·������ָ��_Insert(" & lng·��ID & "," & int�汾�� & ",NULL,1," & _
                    strValue & "," & k & ",'" & GetNodeValue(xSubNode2, "����ָ��") & "'," & _
                    Val(GetNodeValue(xSubNode2, "ָ������")) & ",'" & GetNodeValue(xSubNode2, "ָ����") & "')"
                
                k = k + 1
            Next
        End If
        Set xSubNode1 = xNode.selectSingleNode("Conditions")
        If Not xSubNode1 Is Nothing Then
            For Each xSubNode2 In xSubNode1.childNodes
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_·����������_Insert(" & lng·��ID & "," & int�汾�� & ",NULL,1," & _
                    colItemID("_" & GetNodeValue(xSubNode2, "ָ��ID")) & ",NULL,'" & GetNodeValue(xSubNode2, "��ϵʽ") & "'," & _
                    "'" & GetNodeValue(xSubNode2, "����ֵ") & "','" & GetNodeValue(xSubNode2, "�������") & "')"
            Next
        End If
    End If
    
    '�ٴ�·����֧
    Set xNode = xRoot.selectSingleNode("PathBranchs")
    If Not xNode Is Nothing Then
        Set colBranchID = New Collection
        Set colPreID = New Collection
        For Each xSubNode1 In xNode.childNodes
                strValue = zlDatabase.GetNextId("�ٴ�·����֧")
                colBranchID.Add strValue, "_" & GetNodeValue(xSubNode1, "ID")
                If GetNodeValue(xSubNode1, "ǰһ�׶�ID") <> "" Then
                    strPreStep = strPreStep & "," & GetNodeValue(xSubNode1, "ǰһ�׶�ID")
                    On Error Resume Next
                    If colPreID("_" & GetNodeValue(xSubNode1, "ǰһ�׶�ID")) = "" Then
                        Err.Clear
                        colPreID.Add strValue, "_" & GetNodeValue(xSubNode1, "ǰһ�׶�ID")
                    Else
                        strTmp = colPreID("_" & GetNodeValue(xSubNode1, "ǰһ�׶�ID"))
                        colPreID.Remove "_" & GetNodeValue(xSubNode1, "ǰһ�׶�ID")
                        colPreID.Add strTmp & "," & strValue, "_" & GetNodeValue(xSubNode1, "ǰһ�׶�ID")
                    End If
                    On Error GoTo 0
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                'ǰһ�׶�ID����Null����Ϊǰ��ɾ���˵�ǰ�汾�Ľ׶Σ���û�еõ��µ�ǰһ�׶�ID������ɵĻ����Ҳ�������
                arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·����֧_Update(" & strValue & "," & lng·��ID & "," & int�汾�� & ",'" & GetNodeValue(xSubNode1, "����") & "',Null,'" & _
                    GetNodeValue(xSubNode1, "��׼סԺ��") & "','" & GetNodeValue(xSubNode1, "��׼����") & "','" & GetNodeValue(xSubNode1, "˵��") & "')"
        Next
        strPreStep = Mid(strPreStep, 2)

    End If
    
    '·��ҽ������
    Set xNode = xRoot.selectSingleNode("PathAdvices")
    If Not xNode Is Nothing Then
        Set colAdviceID = New Collection
        Set colAdviceOriginalID = New Collection
        For Each xSubNode1 In xNode.childNodes
            strValue = zlDatabase.GetNextId("·��ҽ������")
            strTemp1 = GetNodeValue(xSubNode1, "ID")
            colAdviceID.Add strValue, "_" & strTemp1
            colAdviceOriginalID.Add strTemp1, "_" & strValue
        Next
        k = 1
        For Each xSubNode1 In xNode.childNodes
            blnDo = True: strTemp1 = "": strTemp2 = "": strTemp3 = ""
                
            '��֤������ĿID
            If Val(GetNodeValue(xSubNode1, "������ĿID")) <> 0 Then
                strSql = "Select ����,ID From ������ĿĿ¼ Where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "��������"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "���Ʊ���") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp1 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp1 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '��֤�շ�ϸĿID
            If blnDo And Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0 Then
                strSql = "Select ����,ID From �շ���ĿĿ¼ Where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "�շ�����"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "�շѱ���") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp2 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp2 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '��ȡ����ο�
            strImportRef = IIf(Val(GetNodeValue(xSubNode1, "������ĿID")) <> 0, Trim(GetNodeValue(xSubNode1, "��������")) & _
                IIf(Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0, "(" & Trim(GetNodeValue(xSubNode1, "�շ�����")) & ")", ""), "" & _
                IIf(Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0, Trim(GetNodeValue(xSubNode1, "�շ�����")), ""))
            '����·��ҽ���ĵ���״��������ʱ��¼��
            rsAdvice.AddNew
            rsAdvice!ID = Val(GetNodeValue(xSubNode1, "ID"))
            rsAdvice!���id = Val(GetNodeValue(xSubNode1, "���ID"))
            rsAdvice!����ο� = strImportRef
            rsAdvice!����״̬ = IIf(blnDo, 1, 0)
            rsAdvice.Update
            
            If blnDo Then
                '��ִ֤�п���ID
                If Val(GetNodeValue(xSubNode1, "ִ�п���ID")) <> 0 Then
                    strSql = "Select ����,ID From ���ű� Where ����=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "ִ�п�����"))
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "ִ�п�����") & "'"
                        If rsTmp.RecordCount > 0 Then
                            strTemp3 = rsTmp!ID
                        Else
                            rsTmp.Filter = ""
                            strTemp3 = rsTmp!ID
                        End If
                    End If
                End If
                
                strValue = GetNodeValue(xSubNode1, "���ID")
                If strValue <> "" Then strValue = colAdviceID("_" & strValue)
                                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_·��ҽ������_Insert(" & _
                    colAdviceID("_" & GetNodeValue(xSubNode1, "ID")) & "," & ZVal(strValue) & "," & _
                    k & "," & Val(GetNodeValue(xSubNode1, "��Ч")) & "," & ZVal(strTemp1) & "," & _
                    "'" & GetNodeValue(xSubNode1, "ҽ������") & "'," & ZVal(GetNodeValue(xSubNode1, "��������")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "�ܸ�����")) & "," & ZVal(strTemp2) & "," & _
                    "'" & GetNodeValue(xSubNode1, "�걾��λ") & "','" & GetNodeValue(xSubNode1, "��鷽��") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "ִ��Ƶ��") & "'," & ZVal(GetNodeValue(xSubNode1, "Ƶ�ʴ���")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "Ƶ�ʼ��")) & ",'" & GetNodeValue(xSubNode1, "�����λ") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "ҽ������") & "'," & Val(GetNodeValue(xSubNode1, "ִ������")) & "," & _
                    ZVal(strTemp3) & ",'" & GetNodeValue(xSubNode1, "ʱ�䷽��") & "',Null,Null," & GetNodeValue(xSubNode1, "�Ƿ�ȱʡ", 0) & "," & _
                    GetNodeValue(xSubNode1, "�Ƿ�ѡ", 0) & ",Null," & ZVal(GetNodeValue(xSubNode1, "�䷽ID", 0)) & "," & ZVal(GetNodeValue(xSubNode1, "�����ĿID", 0)) & ")"
                k = k + 1
            Else
                '��������IDΪ��ҽ���ģ�����Щҽ����Ӧ����
                strValue = GetNodeValue(xSubNode1, "ID")
                For n = 0 To UBound(arrSQL)
                    If arrSQL(n) <> "" Then
                        If Split(arrSQL(n), ",")(1) = colAdviceID("_" & strValue) Then
                            '������ҽ��������
                            strTemp1 = Split(Split(arrSQL(n), ",")(0), "(")(1)
                            colAdviceID.Remove "_" & colAdviceOriginalID("_" & strTemp1)
                            colAdviceID.Add "0", "_" & colAdviceOriginalID("_" & strTemp1)
                            arrSQL(n) = ""
                        End If
                    End If
                Next
                '������ҽ��������
                colAdviceID.Remove "_" & strValue
                colAdviceID.Add "0", "_" & strValue
            End If
        Next
    End If
    
    '�ٴ�·������
    Set xNode = xRoot.selectSingleNode("PathCategorys")
    k = 1
    For Each xSubNode1 In xNode.childNodes
        strTmp = GetNodeValue(xSubNode1, "��֧ID")
        If strTmp = "" Then
            strTmp = "Null"
        Else
            strTmp = colBranchID("_" & strTmp)
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·������_Insert(" & lng·��ID & "," & int�汾�� & "," & k & ",'" & IIf(GetNodeValue(xSubNode1, "����") = "", xSubNode1.Text, GetNodeValue(xSubNode1, "����")) & "',Null," & strTmp & ")"
        k = k + 1
    Next
    
    '�ٴ�·���׶�
    Set xNode = xRoot.selectSingleNode("PathTimeSteps")
    k = 1
    Set colStepID = New Collection
    For Each xSubNode1 In xNode.childNodes
        lng�׶�ID = zlDatabase.GetNextId("�ٴ�·���׶�")
        colStepID.Add lng�׶�ID, "_" & GetNodeValue(xSubNode1, "ID")
        
        strTemp1 = GetNodeValue(xSubNode1, "��ID")
        If strTemp1 <> "" Then strTemp1 = colStepID("_" & strTemp1)
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        strTmp = GetNodeValue(xSubNode1, "��֧ID")
        If strTmp = "" Then
            strTmp = "Null"
        Else
            strTmp = colBranchID("_" & strTmp)
        End If
        If strPreStep <> "" Then
            If InStr("," & strPreStep & ",", "," & GetNodeValue(xSubNode1, "ID") & ",") > 0 Then
                strtemp4 = colPreID("_" & GetNodeValue(xSubNode1, "ID"))
            End If
        End If
        arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���׶�_Insert(" & _
            lng�׶�ID & "," & lng·��ID & "," & int�汾�� & "," & ZVal(strTemp1) & "," & _
            IIf(strTemp1 = "", k, GetNodeValue(xSubNode1, "���")) & ",'" & GetNodeValue(xSubNode1, "����") & "'," & _
            ZVal(GetNodeValue(xSubNode1, "��ʼ����")) & "," & ZVal(GetNodeValue(xSubNode1, "��������")) & "," & _
            "'" & GetNodeValue(xSubNode1, "��־") & "','" & GetNodeValue(xSubNode1, "˵��") & "'," & _
            "'" & GetNodeValue(xSubNode1, "����") & "'," & strTmp & _
            ",'" & strtemp4 & "')"
        If strTemp1 = "" Then k = k + 1
        strtemp4 = ""
        
        '�׶��е�·����Ŀ
        Set xSubNode2 = xSubNode1.selectSingleNode("Items")
        If Not xSubNode2 Is Nothing Then
            Set colItemID = New Collection
            For Each xSubNode3 In xSubNode2.childNodes
                strTemp1 = "": strTemp2 = ""
                '��Ŀ����ҽ��
                lng��ĿID = Val(GetNodeValue(xSubNode3, "ID"))
                Set xSubNode4 = xSubNode3.selectSingleNode("Advices")
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '����ʱ�ṹ��¼��������ҽ������Ŀ�Ĺ���
                        rsAdvice.Filter = "ID=" & Val(xSubNode5.Text)
                        If rsAdvice.RecordCount <> 0 Then
                            Call rsAdvice.Update("��ĿID", lng��ĿID)
                        End If
                        rsAdvice.Filter = ""
                        
                        If Val(colAdviceID("_" & xSubNode5.Text)) <> 0 Then
                            strTemp1 = strTemp1 & "," & colAdviceID("_" & xSubNode5.Text)
                        End If
                    Next
                    strTemp1 = Mid(strTemp1, 2)
                End If
                
                '��Ŀ��������
                Set xSubNode4 = xSubNode3.selectSingleNode("EPRFiles")
                i = 1
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '��֤�����ļ�ID
                        strSql = "Select ID From �����ļ��б� Where ���=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode5, "�ļ����"), GetNodeValue(xSubNode5, "�ļ�����"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & ";" & rsTmp!ID & ",," & GetNodeValue(xSubNode5, "�ļ�����") & "," & i + 1
                    Next
                    strTemp2 = Mid(strTemp2, 2)
                End If
                
                'ͼ�����֤��ֻ֧�ֹ���ͼ��
                strTemp3 = GetNodeValue(xSubNode3, "ͼ��ID")
                If strTemp3 <> "" Then
                    If rsIcon Is Nothing Then
                        strSql = "Select ID,Nvl(����,0) as ���� From �ٴ�·��ͼ��"
                        Set rsIcon = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML")
                    End If
                    rsIcon.Filter = "ID=" & strTemp3 & " And ����=1"
                    If rsIcon.EOF Then strTemp3 = ""
                End If
                
                strValue = zlDatabase.GetNextId("�ٴ�·����Ŀ")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode3, "ID")
                
                strTmp = GetNodeValue(xSubNode3, "��֧ID")
                If strTmp = "" Then
                    strTmp = "Null"
                Else
                    strTmp = colBranchID("_" & strTmp)
                End If
                
                rsAdvice.Filter = "��ĿID=" & lng��ĿID
                
                lngCount = rsAdvice.RecordCount
                strImportRef = ""
                lng������ = 1
                str��IDs = ""
                
                rsAdvice.Filter = rsAdvice.Filter & " And ����״̬=0"
                '��ȡ����״̬
                If rsAdvice.RecordCount <> 0 Then
                    lng������ = IIf(rsAdvice.RecordCount = lngCount, 0, 2)
                    '��ȡδ����ɹ�ҽ������ID
                    For n = 1 To rsAdvice.RecordCount
                        lng��ID = rsAdvice!���id
                        If lng��ID = 0 Then lng��ID = rsAdvice!ID
                        If InStr(str��IDs & ",", "," & lng��ID & ",") = 0 Then
                            str��IDs = str��IDs & "," & lng��ID
                        End If
                        rsAdvice.MoveNext
                    Next
                End If
                If Len(str��IDs) > 0 Then str��IDs = Mid(str��IDs, 2)

                arrID = Split(str��IDs, ",")
                '��ȡ����ο�
                For m = LBound(arrID) To UBound(arrID)
                    '����δ�����ͬһ��ҽ��
                    strFilter = "(��ĿID = " & lng��ĿID & " AND ���ID = " & Val(arrID(m)) & ") OR (��ĿID = " & lng��ĿID & " AND ID=" & Val(arrID(m)) & ")"
                    rsAdvice.Filter = strFilter
                    rsAdvice.Sort = "���ID,ID"
                    If rsAdvice.RecordCount <> 0 Then
                        For n = 1 To rsAdvice.RecordCount
                            If n = 1 And strImportRef = "" Then
                                strImportRef = rsAdvice!����ο�
                            ElseIf n = 1 And strImportRef <> "" Then
                                strImportRef = strImportRef & Chr(10) & Chr(13) & rsAdvice!����ο� '�Ѿ���������ҽ���Ѿ�������strImportRef
                            Else
                                strImportRef = strImportRef & ";" & rsAdvice!����ο�
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                Next
   
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·����Ŀ_Insert(" & _
                    strValue & "," & lng·��ID & "," & int�汾�� & "," & lng�׶�ID & "," & _
                    "'" & GetNodeValue(xSubNode3, "����") & "'," & GetNodeValue(xSubNode3, "��Ŀ���") & "," & _
                    "'" & GetNodeValue(xSubNode3, "��Ŀ����") & "'," & Val(GetNodeValue(xSubNode3, "ִ�з�ʽ")) & "," & _
                    ZVal(GetNodeValue(xSubNode3, "ִ����")) & ",'" & GetNodeValue(xSubNode3, "��Ŀ���") & "'," & _
                    ZVal(strTemp3) & ",'" & strTemp1 & "','" & strTemp2 & "'," & GetNodeValue(xSubNode3, "����Ҫ��", 0) & _
                    "," & strTmp & ",'" & Trim(strImportRef) & "'," & IIf(Trim(strImportRef) = "" And lng������ = 1, "Null", lng������) & _
                    "," & ZVal(GetNodeValue(xSubNode3, "������")) & ")"
            Next
        End If
        
        '�׶�����-����Ǻϲ�·�����򲻵���׶�����
        If lngType = 0 Then
        Set xSubNode2 = xSubNode1.selectSingleNode("StepEval")
            If Not xSubNode2 Is Nothing Then
                '����ָ��
                Set xSubNode3 = xSubNode2.selectSingleNode("Marks")
                If Not xSubNode3 Is Nothing Then
                    i = 1
                    Set colMarkID = New Collection
                    For Each xSubNode4 In xSubNode3.childNodes
                        strValue = zlDatabase.GetNextId("·������ָ��")
                        colMarkID.Add strValue, "_" & GetNodeValue(xSubNode4, "ID")
                        
                        strTmp = GetNodeValue(xSubNode4, "��֧ID")
                        If strTmp = "" Then
                            strTmp = "Null"
                        Else
                            strTmp = colBranchID("_" & strTmp)
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_·������ָ��_Insert(" & _
                            lng·��ID & "," & int�汾�� & "," & lng�׶�ID & ",2," & _
                            strValue & "," & i & ",'" & GetNodeValue(xSubNode4, "����ָ��") & "'," & _
                            Val(GetNodeValue(xSubNode4, "ָ������")) & ",'" & GetNodeValue(xSubNode4, "ָ����") & _
                            "," & strTmp & "')"
                        i = i + 1
                    Next
                End If
                'ָ������
                Set xSubNode3 = xSubNode2.selectSingleNode("Conditions")
                If Not xSubNode3 Is Nothing Then
                    For Each xSubNode4 In xSubNode3.childNodes
                        strTemp1 = GetNodeValue(xSubNode4, "ָ��ID")
                        If strTemp1 <> "" Then strTemp1 = colMarkID("_" & strTemp1)
                        strTemp2 = GetNodeValue(xSubNode4, "��ĿID")
                        If strTemp2 <> "" Then strTemp2 = colItemID("_" & strTemp2)
                        
                        strTmp = GetNodeValue(xSubNode4, "��֧ID")
                        If strTmp = "" Then
                            strTmp = "Null"
                        Else
                            strTmp = colBranchID("_" & strTmp)
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_·����������_Insert(" & _
                            lng·��ID & "," & int�汾�� & "," & lng�׶�ID & ",2," & _
                            ZVal(strTemp1) & "," & ZVal(strTemp2) & ",'" & GetNodeValue(xSubNode4, "��ϵʽ") & "'," & _
                            "'" & GetNodeValue(xSubNode4, "����ֵ") & "'," & Val(GetNodeValue(xSubNode4, "�������")) & _
                            "," & strTmp & ")"
                    Next
                End If
            End If
        End If
    Next
    
    'ִ���ύ����
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ImportPathFromXML"
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Set xPath = Nothing
    ImportPathFromXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

'-------------------------------------------------------------
'������ش������
'-------------------------------------------------------------
Public Function ReadRTFData(ByVal lng����ID As Long, edtEditor As Editor) As Boolean
'���ܣ���ȡ�����ļ���RTF���ݵ�editor�ؼ���
    Dim strZipFile As String, strTempFile As String
        
    On Error GoTo errH
    strZipFile = ReadLobForPath(glngSys, 5, lng����ID)
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
    
     'ɾ����ʱ�ļ�
    Kill strTempFile
    Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Public Function ReadLobForPath(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'���ܣ���ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'������
'lngSys:ϵͳ���
'Action:�������ͣ����������ǲ����ĸ���
'---ϵͳ100,Zl_Lob_Append
'0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
'5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��)��8-���Ӳ�������;9-�����ص����
'10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
'17-ҩƷ���(ʹ��˵��);18-ҩƷ���(ͼƬ);23-��Ӧ��ͼƬ
'---ϵͳ2400,Zl24_Lob_Append
'���鳣��ͼ��,��Action
'---ϵͳ2100,Zl21_Lob_Append
'1-�������͵���;2-���������(��ͼƬֻ�ж�ȡ��û�б���);3-����걨��¼;4-���������Ա,5-���������
'---ϵͳ2600,Zl26_Lob_Append
'14-����ؼ�Ŀ¼,15-������ԴĿ¼
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile:�û�ָ����ŵ��ļ�������ָ��ʱ���Զ�ȡ��ʱ�ļ���
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0������¼,1��ȡת���󱸱��¼
'���أ�������ݵ��ļ�����ʧ���򷵻��㳤��""
    Const conChunkSize As Integer = 10240
    
    Dim rsLob As ADODB.Recordset
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim strSql As String
    Dim objFile As New FileSystemObject
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSql = "Select Zl_Lob_ReadForPath([1],[2],[3],[4],[5]) as Ƭ�� From Dual"
        Case 24
            strSql = "Select Zl24_Lob_Read([2],[3]) as Ƭ�� From Dual"
        Case 21
            strSql = "Select Zl21_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
        Case 26
            strSql = "Select Zl26_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
    End Select
    If strSql = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        Do
            Set rsLob = zlDatabase.OpenSQLRecord(strSql, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).Value) Then Exit Do
            strText = rsLob.Fields(0).Value
            
            ReDim aryChunk(Len(strText) / 2 - 1) As Byte
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLob = zlDatabase.OpenSQLRecord(strSql, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).Value) Then Exit Do
            strText = rsLob.Fields(0).Value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLobForPath = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        If lngCount = 0 Then
            Kill strFile: ReadLobForPath = ""
        Else
            ReadLobForPath = strFile
        End If
    End If
    Err.Clear
End Function

Public Function SaveRTFData(ByVal lng����ID As Long, ByVal lng����ID As Long, lng��ҳID As Long, lngBaby As Long, edtEditor As Editor, Optional ByVal intType As Integer) As Boolean
'���ܣ����没�˲�����ʽRTF����
'������
    Dim strZipFile As String, strTempFile As String, i As Long
        
    'Ҫ�����ݸ���
    Call ElementsUpdate(lng����ID, lng����ID, lng��ҳID, lngBaby, edtEditor, intType)
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    'ѹ���ļ�
    strZipFile = zlFileZip(strTempFile)
    '�����ʽ
    sys.SaveLob glngSys, 5, lng����ID, strZipFile
    
    'ɾ����ʱ�ļ�
    Kill strTempFile
    Kill strZipFile

    SaveRTFData = True
    Exit Function
errH:
    SaveRTFData = False
End Function

Private Function ElementsUpdate(ByVal lng����ID As Long, ByVal lng����ID As Long, lng��ҳID As Long, lngBaby As Long, edtEditor As Editor, Optional ByVal intType As Integer) As Boolean
'���ܣ�����Editor�ؼ��е��滻Ҫ�����ݣ��Ա㱣��ΪRTF�ļ�
'    intType=1 ����
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSql = "Select ������,ID From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And ��ֹ��=0 and �������� =0 And �滻�� =1 order by ������ "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˲���", lng����ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(NVL(rsTmp("������"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_�������༭, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.count
        If ThisElements(i).�滻�� = 1 Then
            ThisElements(i).�����ı� = GetReplaceEleValue(ThisElements(i).Ҫ������, lng����ID, lng��ҳID, IIf(intType = 1, cprPF_����, cprPF_סԺ), 0, lngBaby)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).�滻�� = 1 And ThisElements(i).�Զ�ת�ı� Then
            EleToString edtEditor, ThisElements(i)     '�Զ�ת��Ϊ���ı�����ʱ��ɾ����Ҫ�أ�
        End If
    Next
    Set ThisElements = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    
    bBeteenKeys = FindNextKey(edtThis, 0, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str���� As String
        str���� = Ele.�����ı�
        lngLen = Len(str����)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str����
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
    End If
End Sub

Private Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lngҽ��ID As Long, _
    ByVal lngBaby As Long) As String

    Dim rsTmp As ADODB.Recordset, strSql As String
    
    strSql = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��ID, lngBaby)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    Dim clsZip As zlRichEPR.cZip
    
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    Set clsZip = New zlRichEPR.cZip
    
    With clsZip
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
    Set clsZip = Nothing
End Function

Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim objFSO As New Scripting.FileSystemObject    'FSO����
    Dim clsUnZip As zlRichEPR.cUnzip
    
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If objFSO.FileExists(strZipPath & "TMP.RTF") Then objFSO.DeleteFile strZipPath & "TMP.RTF"
    
    Set clsUnZip = New zlRichEPR.cUnzip
    With clsUnZip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
    Set clsUnZip = Nothing
End Function

Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Public Function Get����ID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lngType As Long, Optional ByVal lng����ID As Long, Optional ByRef bln��ҽ As Boolean = False) As ADODB.Recordset
'������ lngType=1  ȡ���˳���Ҫ���֮�����ϣ�����Ժ�ǵ�һ��ϻ�������Ϸǵ�һ��ϣ���סԺ��ϵ�������ϡ�����֢��ϣ�
'             =0 Ĭ�� �����ǰһ����������ȡ��Ժ�������ҽ��Ժ����ҽ�������;�������ҽ��ʱ, ���ȼ�����ҽ��Ժ����Ժ����ҽ�������
'             =2���ǵڶ��ε�����Ч����Ҫ·����ʱ��ȡ������Ժ����Ժ���
'             =3������ȡ��Ժ�������ҽ��Ժ����ҽ�������;�������ҽ��ʱ, ���ȼ�����ҽ��Ժ����Ժ����ҽ������������Ҫ��ϣ�ͬʱ����϶�Ӧ����Ҫ·����
'˵��:���ų�����¼������
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    gblnGetPath = Val(zlDatabase.GetPara(54, glngSys, 1261)) = 1
    If lngType = 0 Then
        bln��ҽ = sys.DeptHaveProperty(lng����ID, "��ҽ��")
        If bln��ҽ Then
            strSql = "Select ����id, ���id, �������, �������, ��¼��Դ" & vbNewLine & _
                    "From ������ϼ�¼" & vbNewLine & _
                    "Where ��¼��Դ In (1, 2, 3) And" & IIf(gblnGetPath, " ������� In (2,12)", " ������� In (1, 2, 11, 12)") & " And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And" & vbNewLine & _
                    "      Nvl(�Ƿ�����, 0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                    "Order By Decode(�������, 12, 1, 2, 2, 11, 3, 1, 4), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
        Else
            strSql = "Select ����id, ���id, �������,�������,��¼��Դ" & vbNewLine & _
                "From ������ϼ�¼" & vbNewLine & _
                "Where ��¼��Դ In (1, 2, 3) And" & IIf(gblnGetPath, " ������� In (2,12)", " ������� In (1, 2, 11, 12)") & " And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And" & vbNewLine & _
                "       Nvl(�Ƿ�����,0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                "Order By Sign(�������-10),������� Desc, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
        End If
    ElseIf lngType = 1 Then
        strSql = "Select a.����id, a.���id, a.�������, a.�������, a.��¼��Դ" & vbNewLine & _
            "From ������ϼ�¼ A" & vbNewLine & _
            "Where a.��¼��Դ In (1, 2, 3) And a.ȡ��ʱ�� Is Null And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
            "(" & IIf(gblnGetPath, " ������� In (2, 3, 12,13)", " a.������� In (1, 2, 3, 11, 12,13)") & " And a.��ϴ��� <> 1 Or" & vbNewLine & _
            "      a.������� = 10) And Nvl(a.�Ƿ�����, 0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
            "Order By Sign(a.������� - 10), a.������� Desc, Decode(a.��¼��Դ, 1, 4, a.��¼��Դ) Desc"
    ElseIf lngType = 2 Then
        strSql = "Select ����id, ���id, �������,�������,��¼��Դ" & vbNewLine & _
            "From ������ϼ�¼" & vbNewLine & _
            "Where ��¼��Դ In (1, 2, 3) And ������� In ( 2,3,12,13) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And Nvl(�Ƿ�����,0) = 0 " & vbNewLine & _
            "       And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
            "Order By �������, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
    Else
        bln��ҽ = sys.DeptHaveProperty(lng����ID, "��ҽ��")
        If bln��ҽ Then
            strSql = "Select Distinct a.Id, k.����id, k.���id, k.�������, K.�������, K.��¼��Դ,k.���� " & vbNewLine & _
                "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B, �ٴ�·���汾 C," & vbNewLine & _
                "     (Select Rownum As ����, ����id, ���id, �������, �������, ��¼��Դ " & vbNewLine & _
                "       From ������ϼ�¼" & vbNewLine & _
                "       Where ��¼��Դ In (1, 2, 3) And" & IIf(gblnGetPath, " ������� In (2,12)", " ������� In (1, 2, 11, 12)") & " And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� <> 1 And" & vbNewLine & _
                "             Nvl(�Ƿ�����, 0) = 0 And Not (Nvl(����id, 0) = 0 And Nvl(���id, 0) = 0)" & vbNewLine & _
                "       Order By Decode(�������, 12, 1, 2, 2, 11, 3, 1, 4), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc, ��ϴ���) K" & vbNewLine & _
                "Where a.Id = b.·��id And a.Id = b.·��id And a.Id = c.·��id And a.���°汾 = c.�汾�� And a.���� = 0 And b.���� = 0 And" & vbNewLine & _
                "      (b.����id = k.����id Or b.���id = k.���id) And" & vbNewLine & _
                "      (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))" & vbNewLine & _
                "Order By k.����"
        Else
            strSql = "Select Distinct a.Id, k.����id, k.���id, k.�������,K.�������, K.��¼��Դ,k.���� " & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B, �ٴ�·���汾 C," & vbNewLine & _
            "     (Select Rownum As ����, ����id, ���id, �������, �������, ��¼��Դ " & vbNewLine & _
            "       From ������ϼ�¼" & vbNewLine & _
            "       Where ��¼��Դ In (1, 2, 3) And" & IIf(gblnGetPath, " ������� In (2,12)", " ������� In (1, 2, 11, 12)") & " And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� <> 1 And" & vbNewLine & _
            "             Nvl(�Ƿ�����, 0) = 0 And Not (Nvl(����id, 0) = 0 And Nvl(���id, 0) = 0)" & vbNewLine & _
            "       Order By Sign(������� - 10), ������� Desc, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc, ��ϴ���) K" & vbNewLine & _
            "Where a.Id = b.·��id And a.Id = b.·��id And a.Id = c.·��id And a.���°汾 = c.�汾�� And a.���� = 0 And b.���� = 0 And" & vbNewLine & _
            "      (b.����id = k.����id Or b.���id = k.���id) And" & vbNewLine & _
            "      (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))" & vbNewLine & _
            "Order By k.����"
        End If
    End If
    '��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
    '�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
    '�ж����ϵ�����£�������ϴ���ֻȡ��һ����Ҫ���
    '���������������ȣ���Ҫ��Ϊ��֧��������ϡ�
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lng����ID, lng��ҳID, lng����ID)
    Set Get����ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPathTable(ByVal lng����ID As Long, ByVal lng���ID As Long, ByVal lng����ID As Long, ByVal lngPathID As Long, Optional ByVal str����IDs As String, Optional ByVal lng����·��Id As Long, _
            Optional ByVal str���IDs As String, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long) As ADODB.Recordset
'������ str����IDs<>"" �ϲ�·������ʱ���ڶ�����ʱ
'       lng����ID<>0 �ڶ��ε�����Ч·����ʱ�򣬸���������Ժ����Ժ��Ͻ��е��룬�ſ��Ѿ��������·����=0 ����ϲ�·��
    Dim strSql As String
    
    If str����IDs = "" And str���IDs = "" Then
        '�����Distinct����Ϊ�����id�ͼ���id���˰󶨶�Ӧ�����ԣ�����������ظ�ֵ
        strSql = "Select Distinct a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������" & vbNewLine & _
                "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
                "Where a.Id = b.·��id And (b.����id = [1] Or b.���id = [2]) And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                "And a.Id = c.·��id And a.����=0 And b.����=" & IIf(lngPathID = 0, "0", "1") & " And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))" & _
                 IIf(lngPathID = 0, "", " And a.id<>[4]")
    Else
        If lng����ID = 0 Then
            '�����Distinct����Ϊ�����id�ͼ���id���˰󶨶�Ӧ�����ԣ�����������ظ�ֵ���ſ��Ѿ������˵ĺϲ�·��
            strSql = "Select Distinct a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������,b.����ID,b.���ID" & vbNewLine & _
                    "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
                    "Where a.Id = b.·��id And (instr(',' || [5] || ',',',' || b.����ID || ',')>0 and [5] is not null Or instr(',' || [7] || ',',',' || b.���ID || ',')>0 and [7] is not null)  And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                    "And a.Id = c.·��id And a.����=1 And b.����=0 And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))" & _
                    " And Not Exists(Select 1 From ���˺ϲ�·�� D Where a.id=d.·��ID  and d.��Ҫ·����¼ID=[6])"
        Else
            strSql = "Select Distinct a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������" & vbNewLine & _
                    "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
                    "Where a.Id = b.·��id And (instr(',' || [5] || ',',',' || b.����ID || ',')>0 and [5] is not null Or instr(',' || [7] || ',',',' || b.���ID || ',')>0 and [7] is not null)  And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                    "And a.Id = c.·��id And a.����=0 And b.����=" & IIf(lngPathID = 0, "0", "1") & " And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))" & _
                    " And Not Exists(Select 1 From �����ٴ�·�� D Where a.ID=d.·��ID And d.����ID=[8] And D.��ҳID=[9])" & _
                    IIf(lngPathID = 0, "", " And a.id<>[4]")
        End If
    End If
    On Error GoTo errH
    Set GetPathTable = zlDatabase.OpenSQLRecord(strSql, "��ȡ·��Ŀ¼", lng����ID, lng���ID, lng����ID, lngPathID, str����IDs, lng����·��Id, str���IDs, lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathOutLog() As Boolean
'���ܣ�����Ƿ���ڲ��˳����Ǽ���Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From ·������ṹ Where ����ID = 2 And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ·������ṹ")
    CheckPathOutLog = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathOutDiag(ByVal lng·��ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ������д�˳�Ժ��ϣ����жϳ�Ժ����Ƿ�͵���·���������ͬ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim str����IDs As String, str���IDs As String, lngDiagType As Long '������ͣ�1/2  ��ҽ  ��11.12   ��ҽ
 
    strSql = "Select b.����ID,b.���ID,a.������� From �����ٴ�·�� A,�ٴ�·������ B Where A.·��ID=B.·��ID And A.ID = [1] And NVL(b.����,0)=0"
    
    CheckPathOutDiag = True
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�жϳ�Ժ���", lng·��ID)
    If rsTmp.RecordCount > 0 Then
        lngDiagType = Val(rsTmp!������� & "")
        Do While Not rsTmp.EOF
            If Val(rsTmp!����id & "") <> 0 Then
                str����IDs = str����IDs & "," & Val(rsTmp!����id & "")
            End If
            If Val(rsTmp!���id & "") <> 0 Then
                str���IDs = str���IDs & "," & Val(rsTmp!���id & "")
            End If
            rsTmp.MoveNext
        Loop
        str����IDs = Mid(str����IDs, 2)
        str���IDs = Mid(str���IDs, 2)
        
        strSql = "Select ����ID,���ID From ������ϼ�¼ Where ��ϴ���=1 And NVL(�������,1) = 1 and ��¼��Դ=3 And ����ID=[1] And ��ҳID=[2]"
        If lngDiagType = 1 Or lngDiagType = 2 Then
            strSql = strSql & " and �������=3"
        Else
            strSql = strSql & " and �������=13"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�жϳ�Ժ���", lng����ID, lng��ҳID)
        '���û���Ժ��ϣ��򲻼��
        If rsTmp.RecordCount > 0 Then
            If InStr("," & str����IDs & ",", "," & Val(rsTmp!����id & "") & ",") = 0 And InStr("," & str���IDs & ",", "," & Val(rsTmp!���id & "") & ",") = 0 Then
                CheckPathOutDiag = False
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˲���״̬(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional lngӤ������ID As Long, Optional lngӤ������ID As Long) As Long
'���ܣ���ȡ���˲����ύ״̬
'      0-δ�ύ;1-�ȴ����(�ύ);2-�ܾ����;3-�������;4-��鷴��;5-���鵵;6-�������;13-���ڳ��;14-��鷴��;16-�������
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ����״̬,Ӥ������ID,Ӥ������ID From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����ύ״̬", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        Get���˲���״̬ = Val("" & rsTmp!����״̬)
        lngӤ������ID = Val(rsTmp!Ӥ������ID & "")
        lngӤ������ID = Val(rsTmp!Ӥ������ID & "")
    Else
        lngӤ������ID = 0
        lngӤ������ID = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiPathOutLog(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ�����Ƿ���ڲ��˳�����¼
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From ���˳�����¼ Where ����ID = [1] And ��ҳID = [2] And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˳�����¼", lng����ID, lng��ҳID)
    CheckPatiPathOutLog = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get�׶η���(Optional ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long) As String
'���ܣ���ȡ����ʹ�ù��Ľ׶εķ��ֻ࣬�з�֧·�����з��࣬���ʹ���˸÷��࣬��������·���ڼ�ֻ��ѡ��÷��࣬����ֻ������һ������
'������lng�׶�ID=ָ���ò���ʱ����ȡָ���׶εķ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If lng�׶�ID <> 0 Then
        strSql = "Select ���� From �ٴ�·���׶� Where id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶η���", lng�׶�ID)
    Else
        strSql = "Select a.����" & vbNewLine & _
                "From �ٴ�·���׶� A, (Select Distinct �׶�id From ����·��ִ�� Where ·����¼id = [1]) B" & vbNewLine & _
                "Where a.Id = b.�׶�id And a.���� Is Not Null And rownum<2"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶η���", lng·����¼ID)
    End If
    If rsTmp.RecordCount > 0 Then Get�׶η��� = "" & rsTmp!����
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDiagnose(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int��Դ As Integer) As ADODB.Recordset
'���ܣ���ȡ���˵���Ҫ��ϻ��Ҫ���
'������lng����ID=�Һ�ID����ҳID
'      int��Դ=1-����,2-סԺ
'���أ������¼��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ����ID,b.����,��¼��Դ,Mod(�������,10) as ���� From ������ϼ�¼ a ,��������Ŀ¼ b" & _
        " Where ����ID=[1] And ��ҳID=[2] And NVL(A.�������,1) = 1 And ������� IN(" & IIf(int��Դ = 1, "1,11", "1,2,3,11,12,13") & ") and a.����ID=b.ID and ����ID is not null and nvl(�Ƿ�����,0)=0" & _
        " Order by ��¼��Դ,�������,��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPatiDiagnose", lng����ID, lng����ID)
    
    '�Ȱ���Դ����˳�����
    rsTmp.Filter = "��¼��Դ=3" '��ҳ����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=2" '��Ժ�Ǽ�
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=1" '����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=4" '������¼��
    
    'סԺ�ٰ���������˳�����
    If Not rsTmp.EOF And int��Դ = 2 Then
        strSql = rsTmp.Filter
        rsTmp.Filter = strSql & " And ����=3"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And ����=2"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And ����=1"
    End If
    
    Set GetPatiDiagnose = zl9ComLib.zlDatabase.CopyNewRec(rsTmp)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathSend(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����ò��˵���סԺ�Ƿ����ɹ���Ŀ
'���أ�true=���ɹ���false=δ���ɹ�
    Dim strSql As String, rsPati As Recordset
    
    strSql = "Select Max(״̬) as ״̬ From �����ٴ�·�� Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckPathSend", lng����ID, lng��ҳID)
    If rsPati.RecordCount > 0 Then
        If Val(rsPati!״̬ & "") <> 0 Then CheckPathSend = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AddOutPathItem(ByVal strAdviceIDs As String, ByVal lngMode As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
   Optional ByVal bytType As Byte, Optional ByRef colSQL As Collection)
'����:��ҽ��û��ֹͣ�ĳ���������Ϊ·������Ŀ
'������strAdviceIds - id,id...lngMode=1:����ҽ��ID�а�ֹͣ��δֹͣ�ĳ���ҽ����
'                             lngMode=2�����л��˵�ҽ��ID(����ֹͣ����ҽ��,�������� ����|��ʱҽ��)
'       lngMode��1-·������ ��
'                2-ҽ����ʿվ��ҽ��״̬Ϊ�����ϻ�ֹͣ��ҽ���ڻ���ʱ����
'       bytType =��lngMode=2�� =4 ��������;=8 ����ֹͣ
'       colSQL =���ؿ�ִ��SQL

    Dim strSql As String, strStopIds As String, strPathOut As String
    Dim rsTmp       As ADODB.Recordset
    Dim AddDate     As Date
    Dim strDate, strAddDate As String
    Dim i           As Long, j As Long
    Dim str����ԭ��     As String
    Dim blnTrans    As Boolean
    Dim lng����·��Id, lng�׶�ID, lng���� As Long
    Dim varTemp As Variant
    Dim strItemType As String
    Set colSQL = New Collection
    On Error GoTo errH
    'a.���ݴ��˵�strAdviceIds��ѯû��ֹͣ�ĳ���
    If lngMode = 1 Then
        strSql = "Select /*+ rule*/" & vbNewLine & _
                 " Column_Value As ����ҽ��id" & vbNewLine & _
                 "From Table(f_Num2list([1])) A, ����ҽ����¼ B" & vbNewLine & _
                 "Where a.Column_Value = b.Id And b.ͣ��ʱ�� Is Null"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strAdviceIDs)

        '��ȡδֹͣ�ĳ���ID
        For i = 1 To rsTmp.RecordCount
            strStopIds = strStopIds & "," & rsTmp!����ҽ��id
            rsTmp.MoveNext
        Next
        strStopIds = Mid(strStopIds, 2)
    ElseIf lngMode = 2 Then    '���ϻ�ֹͣ��ҽ������ʱ������ҽ�� ��ID
        '��Ҫֹͣ��ҽ�����ų�����·�����Ѿ����ɹ���
        strSql = "Select f_List2str(Cast(Collect(t.����ҽ��id || '') As t_Strlist)) As ����ҽ��ids" & vbNewLine & _
                 "From (Select a.Column_Value As ����ҽ��id" & vbNewLine & _
                 "       From Table(f_Num2list([1])) A" & vbNewLine & _
                 "       Minus" & vbNewLine & _
                 "       Select d.����ҽ��id" & vbNewLine & _
                 "       From Table(f_Num2list([1])) C, ����·��ҽ�� D, ����·��ִ�� E, �����ٴ�·�� F" & vbNewLine & _
                 "       Where d.����ҽ��id = c.Column_Value And d.·��ִ��id = e.Id And f.Id = e.·����¼id And f.��ǰ�׶�id = e.�׶�id And f.��ǰ���� = e.����) T"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strAdviceIDs)
        If rsTmp.RecordCount = 1 Then strStopIds = rsTmp!����ҽ��ids & ""
    End If
    
    If strStopIds <> "" Then
        '���˵�·���ڵ�ҽ��ID(1-����·������Ŀ�����Զ�����Ϊ·����)����·����ҽ��ID
        Call CheckStopAdvice(lng����ID, lng��ҳID, strStopIds, colSQL)
        If strAdviceIDs = "" Then Exit Sub
        '��������ı���ԭ�����
        strSql = "Select ���� From ���쳣��ԭ�� Where (����='����' Or ����='����')  And ����=1 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem")
        If rsTmp.RecordCount > 0 Then
            str����ԭ�� = rsTmp!���� & ""
        End If

        '��ȡ��ǰ·����·����¼ID,��ǰ�׶�Id,��ǰ���ڣ�����
        strSql = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, b.���� " & vbNewLine & _
                 "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, Max(b.Id) ִ��id" & vbNewLine & _
                 "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
                 "       Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.·����¼id And b.�׶�id = a.��ǰ�׶�id And b.���� = a.��ǰ����" & vbNewLine & _
                 "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����) A, ����·��ִ�� B" & vbNewLine & _
                 "Where a.ִ��id = b.Id"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", lng����ID, lng��ҳID)

        If rsTmp.RecordCount = 1 Then
            lng����·��Id = Val(rsTmp!·����¼id)
            lng�׶�ID = Val(rsTmp!��ǰ�׶�ID)
            strDate = "To_Date('" & Format(rsTmp!����, "yyyy-MM-dd") & "','YYYY-MM-DD')"
            lng���� = Val(rsTmp!��ǰ����)
        Else
            Exit Sub
        End If

        strSql = "Select a.Id, a.����, Decode(a.��Ŀid, Null, a.��Ŀ����, c.��Ŀ����) As ��Ŀ����, Decode(a.��Ŀid, Null, a.ִ����, c.ִ����) As ִ����," & vbNewLine & _
                 "         Decode(a.��Ŀid, Null, a.��Ŀ���, c.��Ŀ���) As ��Ŀ���, Decode(a.��Ŀid, Null, a.ͼ��id, c.ͼ��id) As ͼ��id," & vbNewLine & _
                 "         f_List2str(Cast(Collect(b.����ҽ��id || '') As t_Strlist)) As ����ҽ��ids" & vbNewLine & _
                 "  From (Select " & vbNewLine & _
                 "          Row_Number() Over(Partition By d.����ҽ��id Order By d.·��ִ��id Desc) As Top, d.·��ִ��id, d.����ҽ��id" & vbNewLine & _
                 "         From ����·��ҽ�� D, Table(f_Num2list([1])) E" & vbNewLine & _
                 "         Where d.����ҽ��id = e.Column_Value" & vbNewLine & _
                 "         Group By d.·��ִ��id, d.����ҽ��id) B, ����·��ִ�� A, �ٴ�·����Ŀ C" & vbNewLine & _
                 "  Where b.Top = 1 And b.·��ִ��id = a.Id And a.��Ŀid = c.Id(+)" & vbNewLine & _
                 "  Group By a.Id, a.����, Decode(a.��Ŀid, Null, a.��Ŀ����, c.��Ŀ����), Decode(a.��Ŀid, Null, a.ִ����, c.ִ����)," & vbNewLine & _
                 "           Decode(a.��Ŀid, Null, a.��Ŀ���, c.��Ŀ���), Decode(a.��Ŀid, Null, a.ͼ��id, c.ͼ��id)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strStopIds)

        AddDate = zlDatabase.Currentdate
        strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        '��ǰһ�׶λ�ǰһ�죬δֹͣ��ҽ����Ϊ·������Ŀ����ӵ�·����
        For i = 1 To rsTmp.RecordCount
            strSql = "Zl_����·������_Insert(1," & lng����ID & "," & lng��ҳID & ",Null,0," & _
                                     lng����·��Id & "," & lng�׶�ID & "," & strDate & "," & lng���� & ",'" & NVL(rsTmp!����) & "',Null" & ",'" & rsTmp!����ҽ��ids & "',Null,Null" & _
                                     ",'" & UserInfo.���� & "'," & strAddDate & ",'" & CStr(NVL(rsTmp!��Ŀ����)) & "'" & _
                                     "," & Val(NVL(rsTmp!ִ����, 1)) & ",'" & CStr(NVL(rsTmp!��Ŀ���)) & "'," & NVL(rsTmp!ͼ��ID, "Null") & ",'δͣ�õĳ���','" & str����ԭ�� & "' ,0)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
            AddDate = AddDate + 1 / 24 / 60 / 60
            '��ȡδ���ӳɹ���ҽ������·������Ŀ
            varTemp = Split(rsTmp!����ҽ��ids, ",")
            For j = LBound(varTemp) To UBound(varTemp)
                strStopIds = Replace("," & strStopIds & ",", "," & varTemp(j) & ",", ",")
                If Left(strStopIds, 1) = "," Then strStopIds = Mid(strStopIds, 2)
                If Right(strStopIds, 1) = "," Then strStopIds = Mid(strStopIds, 1, Len(strStopIds) - 1)
            Next
            rsTmp.MoveNext
        Next
        If strStopIds <> "" Then
            Call GetPatiPathInfo(lng����ID, lng��ҳID, strItemType)
            strSql = "Zl_����·������_Insert(1," & lng����ID & "," & lng��ҳID & ",Null,0," & _
                         lng����·��Id & "," & lng�׶�ID & "," & strDate & "," & lng���� & ",'" & strItemType & "',Null" & ",'" & strStopIds & "',Null,Null" & _
                         ",'" & UserInfo.���� & "'," & strAddDate & ",'·������Ŀ'" & _
                         ",1,'�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��',NULL,NULL,'" & str����ԭ�� & "' ,1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
        '�ύ����,��������
        If lngMode = 1 Then
            gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.count
                Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "AddOutPathItem")
            Next
            gcnOracle.CommitTrans: blnTrans = False
        End If
    End If
    Exit Sub
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CheckStopAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef strUnStopIDs As String, Optional ByRef colSQL As Collection)
'����:���˵��ܹ�ƥ��Ϊ·����ҽ����ҽ��ID
'����:
'   strUnStopIDs-ҽ��IDS
'����:
'   strUnStopIDs-δֹͣ��ҽ��ID��һ��ҽ��������ID������Ҫ��ӵ�·������Ŀ
'   colSQL-���ؿ�ִ��SQL

    Dim rsUnStop As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset

    Dim strSql As String
    Dim i As Long, j As Long
    Dim k As Long

    Dim lng����·��Id  As Long
    Dim lng�׶�ID As Long
    Dim lng���� As Long
    Dim lngPos As Long, lngPathPos As Long
    Dim strDate As String
    Dim strTag As String
    Dim str���ID As String
    Dim AddDate As Date
    
    Dim blnTrans As Boolean
    Dim strҽ��ID As String
    Dim blnƥ����Ч As Boolean
    
    
    On Error GoTo errH
    
    Set colSQL = New Collection
    
    blnƥ����Ч = CBool(zlDatabase.GetPara("ƥ��ʱ��Ч��ͬ��·������Ŀ", glngSys, P�ٴ�·��Ӧ��, "0"))
    '��ȡ��ǰ·����·����¼ID,��ǰ�׶�Id,��ǰ���ڣ�����
    strSql = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, b.���� " & vbNewLine & _
             "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, Max(b.Id) ִ��id" & vbNewLine & _
             "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
             "       Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.·����¼id And b.�׶�id = a.��ǰ�׶�id And b.���� = a.��ǰ����" & vbNewLine & _
             "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����) A, ����·��ִ�� B" & vbNewLine & _
             "Where a.ִ��id = b.Id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng����ID, lng��ҳID)

    If rsTmp.RecordCount = 1 Then
        lng����·��Id = Val(rsTmp!·����¼id)
        lng�׶�ID = Val(rsTmp!��ǰ�׶�ID)
        strDate = "To_Date('" & Format(rsTmp!����, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        lng���� = Val(rsTmp!��ǰ����)
    Else
        Exit Sub
    End If

    strSql = "select a.ID, a.���ID, b.���, a.������ĿID, b.��������,a.ҽ����Ч " & vbNewLine & _
            "  from ����ҽ����¼ a, ������ĿĿ¼ b" & vbNewLine & _
            " where a.������ĿID = b.id" & vbNewLine & _
            "   and a.id in (Select Column_Value As ����ҽ��id" & vbNewLine & _
            "                  From Table(f_Num2list([1]))) Order by a.���"


    Set rsUnStop = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", strUnStopIDs)

    strSql = "Select c.ID, c.���ID,c.������Ŀid,a.id as ·����ĿID,a.���� as ·����Ŀ����,c.��Ч " & vbNewLine & _
            "From �ٴ�·����Ŀ a, �ٴ�·��ҽ�� b, ·��ҽ������ c" & vbNewLine & _
            "where a.id = b.·����Ŀid" & vbNewLine & _
            "   and b.ҽ������id = c.id" & vbNewLine & _
            "   and a.�׶�id = [1]"
            
    Set rsPath = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng�׶�ID)

    strTag = ""
    Set rsPathAdvice = Nothing
    For i = 1 To rsUnStop.RecordCount
        lngPos = rsUnStop.AbsolutePosition
        If Val(rsUnStop!���id & "") = 0 And Not (rsUnStop!��� & "" = "E" And rsUnStop!�������� & "" = "2") Or InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
            '��һ����ҩ������һ��ʱ��ֻ�������õ�ǰ�У���Ϊ·������Ŀ���ܺ�·������Ŀһ����ҩ
            If InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
                'ҩƷ����ҩ����ƥ�� 65982
                rsUnStop.Filter = "ID=" & rsUnStop!ID
                str���ID = Val(rsUnStop!���id & "")
            Else
                rsUnStop.Filter = "ID=" & rsUnStop!ID & " Or ���ID=" & rsUnStop!ID
                str���ID = Val(rsUnStop!ID & "")
            End If
            'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,9-��Ѫ�ɼ�,���鲻���ɼ���ʽ��������������������������鲻����λ����
            If Not (rsUnStop!��� & "" = "E" And InStr(",2,3,4,6,8,9,", "," & rsUnStop!�������� & ",") > 0) _
                And Not (InStr(",G,F,D,", "," & rsUnStop!��� & ",") > 0 And Val(rsUnStop!���id & "") <> 0) Then
                If blnƥ����Ч Then
                    rsPath.Filter = "������Ŀid=" & NVL(rsUnStop!������ĿID, 0) & " And ��Ч = " & rsUnStop!ҽ����Ч
                Else
                    rsPath.Filter = "������Ŀid=" & NVL(rsUnStop!������ĿID, 0)
                    If rsUnStop!ҽ����Ч = 0 Then  '���Ȱ���Чƥ��·��
                        rsPath.Sort = "��Ч ASC"
                    Else
                        rsPath.Sort = "��Ч DESC"
                    End If
                End If
                If rsPath.RecordCount > 0 Then
                    '·������Ŀ
                    If InStr("," & strTag & ",", "," & str���ID & ",") = 0 Then
                        rsUnStop.Filter = "���ID=" & str���ID & " OR ID =" & str���ID
                        If InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
                            strTag = strTag & "," & rsUnStop!���id
                        Else
                            strTag = strTag & "," & rsUnStop!ID
                        End If
                        
                        If rsPathAdvice Is Nothing Then Set rsPathAdvice = MakePathAdivceRS
                        For k = 1 To rsUnStop.RecordCount
                            rsPathAdvice.Filter = "·����ĿID = " & rsPath!·����ĿID
                            If rsPathAdvice.RecordCount = 0 Then
                                rsPathAdvice.AddNew
                                rsPathAdvice!·����ĿID = rsPath!·����ĿID & ""
                                rsPathAdvice!·����Ŀ���� = rsPath!·����Ŀ���� & ""
                                rsPathAdvice!ҽ��IDs = rsUnStop!ID & ""
                            Else
                                rsPathAdvice!ҽ��IDs = rsPathAdvice!ҽ��IDs & "," & rsUnStop!ID
                            End If
                            rsPathAdvice.Update
                            '��δֹͣ�ĳ������Ƴ�
                            strUnStopIDs = Replace("," & strUnStopIDs & ",", "," & rsUnStop!ID & ",", ",")
                            If Left(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 2)
                            If Right(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 1, Len(strUnStopIDs) - 1)
                            rsUnStop.MoveNext
                        Next
                    End If
                End If
            End If
        End If
        rsUnStop.Filter = ""
        rsUnStop.AbsolutePosition = lngPos
        rsUnStop.MoveNext
    Next
    
    If rsPathAdvice Is Nothing Then Exit Sub
    rsPathAdvice.Filter = ""
    AddDate = zlDatabase.Currentdate
    For j = 1 To rsPathAdvice.RecordCount
        strSql = "Zl_����·������_Insert(1," & lng����ID & "," & lng��ҳID & ",NULL,0," & lng����·��Id & "," & lng�׶�ID & _
            "," & strDate & "," & lng���� & ",'" & rsPathAdvice!·����Ŀ���� & "'," & rsPathAdvice!·����ĿID & ",'" & rsPathAdvice!ҽ��IDs & "',Null,Null" & _
            ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
            
        colSQL.Add strSql, "C" & colSQL.count + 1
        '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
        AddDate = AddDate + 1 / 24 / 60 / 60
        rsPathAdvice.MoveNext
    Next
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetPatiPathInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByRef str·����Ŀ���� As String = "-1") As ADODB.Recordset
'���ܣ���ȡ·�����˵�ǰ·����Ϣ
'���أ�str����=��ǰ�������һ��·����Ŀ�����ķ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str·����Ŀ���� <> "-1"
    str·����Ŀ���� = ""
    strSql = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, a.·��ID, a.�汾��, a.��ʼ����, b.����, b.����" & vbNewLine & _
            "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, a.·��ID, a.�汾��, Max(b.Id) ִ��id, Min(c.����) As ��ʼ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B, ����·��ִ�� C" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.Id = c.·����¼id And b.�׶�id + 0 = a.��ǰ�׶�id And b.���� = a.��ǰ���� And a.״̬ = 1 And" & vbNewLine & _
            "             a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
            "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����, a.·��ID, a.�汾��) A, ����·��ִ�� B" & vbNewLine & _
            "Where a.ִ��id = b.Id"

    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSql, "���˵�ǰ·����Ϣ", lng����ID, lng��ҳID)
    If rsRet.RecordCount > 0 And blnDo Then
        str·����Ŀ���� = "" & rsRet!����
        
        '�������������ҽ������Ŀ����ȡҽ������Ŀ�ķ���
        strSql = "Select ����" & vbNewLine & _
                "From ����·��ִ��" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From ����·��ִ�� A" & vbNewLine & _
                "            Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And Exists (Select 1 From ����·��ҽ�� B Where a.Id = b.·��ִ��id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˵�ǰ·����Ϣ", Val(rsRet!·����¼id), Val(rsRet!��ǰ�׶�ID), CDate(rsRet!����))
        If rsTmp.RecordCount > 0 Then
            str·����Ŀ���� = "" & rsTmp!����
        End If
    End If
    Set GetPatiPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod, int����)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Function FuncGetEMRInfo(ByVal strInfo As String) As ADODB.Recordset
'����:���޸ĺ�ĵ��Ӳ�����Ϣ�����ɼ�¼�����ء�
'����: �������飺��ʽ���ļ�ID,ԭ��ID,�ļ�����,���;�ļ�ID,ԭ��ID,�ļ�����,���...
'˵�����޸ĵ��Ӳ�����,����Ƶ�·����Ŀ���޷���ʾ�޸�����,��Ϊ�°���Ӳ������ݻ�û�в��뵽��׼����
    Dim rsEMR As ADODB.Recordset
    Dim i As Long
    Dim arrtmp As Variant
    Dim arrTmpSub As Variant
    
    Set rsEMR = New ADODB.Recordset
    rsEMR.Fields.Append "�ļ�ID", adBigInt
    rsEMR.Fields.Append "ԭ��ID", adVarChar, 32
    rsEMR.Fields.Append "����", adVarChar, 100
    rsEMR.Fields.Append "���", adVarChar, 10
    rsEMR.Fields.Append "�汾", adVarChar, 2
    
    rsEMR.CursorLocation = adUseClient
    rsEMR.LockType = adLockOptimistic
    rsEMR.CursorType = adOpenStatic
    rsEMR.Open
    arrtmp = Split(strInfo, ";")
    For i = LBound(arrtmp) To UBound(arrtmp)
        arrTmpSub = Split(arrtmp(i), ",")
        rsEMR.AddNew
        rsEMR!�ļ�ID = Val(arrTmpSub(0))
        rsEMR!ԭ��ID = arrTmpSub(1)
        rsEMR!���� = arrTmpSub(2)
        rsEMR!��� = arrTmpSub(3)
        If Val(arrTmpSub(0)) = 0 Then
            rsEMR!�汾 = 2
        Else
            rsEMR!�汾 = 1
        End If
        rsEMR.Update
    Next
    If rsEMR.RecordCount <> 0 Then rsEMR.MoveFirst
    Set FuncGetEMRInfo = rsEMR
End Function

Public Sub ZLHIS_CIS_001(ByRef objMip As zl9ComLib.clsMipModule, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByVal lng����ID As Long, ByVal lng����ID As Long)
'���ܣ�����ҽ�����´���Ϣ  ZLHIS_CIS_001��·��ҽ�����ɺ�·����ҽ�����
    Dim strXML As String
    Dim strSql As String, strTmp As String
    Dim str��� As String, str����������� As String
    Dim bln��Ϣƽ̨ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long
    
    On Error GoTo errH
    
    If Not objMip Is Nothing Then
        If objMip.IsConnect Then bln��Ϣƽ̨ = True
    End If
    
    If bln��Ϣƽ̨ Then
        strSql = "select ���� from ���ű� where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng����ID)
        str����������� = rsTmp!���� & ""
    End If
    
    '����У�Ե�ҽ��
    strSql = "select id,������־ from ����ҽ����¼ a where A.ҽ��״̬=1 and a.����id=[1] and a.��ҳid=[2]" & _
        " And Nvl(A.���״̬,0) Not in(1,3,4,5) And Exists ( Select M.���� From ��Ա�� M,ִҵ��� N" & _
        " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
        " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')) And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        lngҽ��ID = Val(rsTmp!ID & "")
        rsTmp.Filter = "������־=1"
        If Not rsTmp.EOF Then
            lngҽ��ID = Val(rsTmp!ID & "")
        End If
    End If
    
    If lngҽ��ID = 0 Then Exit Sub
    
    'ȡһ��ҽ�����ɣ�����ҽ������
    strSql = "select a.id as ҽ��id,a.������Դ,decode(a.������־,1,1,0) as ������־,a.ҽ����Ч,a.�������,b.��������,a.����ҽ��," & vbNewLine & _
        " to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,a.��������id,c.����" & vbNewLine & _
        " from ����ҽ����¼ a,������ĿĿ¼ b,���ű� c where a.id=[1] and a.������Ŀid=b.id(+) and a.��������id=c.id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lngҽ��ID)
    
    str��� = rsTmp!������� & ""
    If rsTmp!������� & "" = "E" Then
        If rsTmp!�������� & "" = "2" Then
            str��� = "5"
        ElseIf rsTmp!�������� & "" = "4" Then
            str��� = "7"
        ElseIf rsTmp!�������� & "" = "6" Then
            str��� = "C"
        End If
    End If
    
    strXML = "": strTmp = ""
    
    strXML = strXML & "<patient_info>"
    strXML = strXML & "   <patient_id>" & lng����ID & "</patient_id>"
    strXML = strXML & "</patient_info>"
    strXML = strXML & "<patient_clinic>"
    strXML = strXML & "   <patient_source>" & rsTmp!������Դ & "</patient_source>"
    strXML = strXML & "   <clinic_id>" & lng��ҳID & "</clinic_id>"
    strXML = strXML & "   <clinic_area_id>" & lng����ID & "</clinic_area_id>"
    strXML = strXML & "   <clinic_dept_id>" & lng����ID & "</clinic_dept_id>"
    strXML = strXML & "   <clinic_dept_title>" & str����������� & "</clinic_dept_title>"
    strXML = strXML & "</patient_clinic>"
    strXML = strXML & "<new_order>"
    strXML = strXML & "   <order_id>" & rsTmp!ҽ��id & "</order_id>"
    strXML = strXML & "   <order_urgency>" & rsTmp!������־ & "</order_urgency>"
    strXML = strXML & "   <order_expiry>" & rsTmp!ҽ����Ч & "</order_expiry>"
    strXML = strXML & "   <order_kind>" & str��� & "</order_kind>"
    strXML = strXML & "   <operation_kind>" & rsTmp!�������� & "</operation_kind>"
    strXML = strXML & "   <create_doctor>" & rsTmp!����ҽ�� & "</create_doctor>"
    strXML = strXML & "   <create_time>" & rsTmp!����ʱ�� & "</create_time>"
    strXML = strXML & "   <create_dept_id>" & rsTmp!��������id & "</create_dept_id>"
    strXML = strXML & "   <create_dept_title>" & rsTmp!���� & "</create_dept_title>"
    strXML = strXML & "</new_order>"
    
    If bln��Ϣƽ̨ Then Call objMip.CommitMessage("ZLHIS_CIS_001", strXML, strTmp)
    
    Call zlDatabase.SendMsg("ZLHIS_CIS_001", IIf(strTmp = "", strXML, strTmp))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function MakePathItems() As ADODB.Recordset
    Set MakePathItems = New ADODB.Recordset
    
    MakePathItems.Fields.Append "ID", adBigInt
    MakePathItems.Fields.Append "Ӥ��", adBigInt
    MakePathItems.Fields.Append "���id", adBigInt
    MakePathItems.Fields.Append "������ĿID", adBigInt
    MakePathItems.Fields.Append "���", adVarChar, 10
    MakePathItems.Fields.Append "��������", adVarChar, 20
    MakePathItems.Fields.Append "·����ĿID", adBigInt
    MakePathItems.Fields.Append "·����Ŀ����", adVarChar, 50, adFldIsNullable
    MakePathItems.Fields.Append "��Ч", adSmallInt
     
    MakePathItems.CursorLocation = adUseClient
    MakePathItems.LockType = adLockOptimistic
    MakePathItems.CursorType = adOpenStatic
    MakePathItems.Open
End Function

Private Function MakePathAdivceRS() As ADODB.Recordset
    Set MakePathAdivceRS = New ADODB.Recordset
    
    MakePathAdivceRS.Fields.Append "�к�", adBigInt
    MakePathAdivceRS.Fields.Append "·����ĿID", adBigInt
    MakePathAdivceRS.Fields.Append "ԭҽ��ID", adBigInt
    
    MakePathAdivceRS.Fields.Append "·����Ŀ����", adVarChar, 50, adFldIsNullable
    MakePathAdivceRS.Fields.Append "ҽ��IDS", adLongVarWChar, 4000, adFldIsNullable
    MakePathAdivceRS.CursorLocation = adUseClient
    MakePathAdivceRS.LockType = adLockOptimistic
    MakePathAdivceRS.CursorType = adOpenStatic
    MakePathAdivceRS.Open
End Function

Private Function MakePathRichEPR() As ADODB.Recordset
    Set MakePathRichEPR = New ADODB.Recordset
    
    MakePathRichEPR.Fields.Append "ID", adBigInt
    MakePathRichEPR.Fields.Append "�ļ�ID", adBigInt
    MakePathRichEPR.Fields.Append "·����ĿID", adBigInt
    MakePathRichEPR.Fields.Append "·����Ŀ����", adVarChar, 50, adFldIsNullable
    
    MakePathRichEPR.CursorLocation = adUseClient
    MakePathRichEPR.LockType = adLockOptimistic
    MakePathRichEPR.CursorType = adOpenStatic
    MakePathRichEPR.Open
End Function

Private Function Get֤��IDs(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ�ò���֤��IDs�����ŷָ�
    Dim strSql As String, rsTmp As Recordset
    Dim str֤��IDs As String
    
    strSql = "Select ֤��ID From ������ϼ�¼ Where ����id = [1] And ��ҳid = [2] And ֤��id Is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get֤��IDs", lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        str֤��IDs = str֤��IDs & "," & rsTmp!֤��id
        rsTmp.MoveNext
    Loop
    Get֤��IDs = Mid(str֤��IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckPathInItem(ByVal str������ĿIDs As String, ByRef str���� As String, udtPati As TYPE_Pati, Optional ByRef lngDay As Long, Optional ByVal lng�׶�ID As Long, _
        Optional ByRef rsStepAdvice As Recordset, Optional ByVal bln��ҩ�䷽ As Boolean, Optional ByVal byt��Ч As Byte) As Long
'���ܣ�����ٴ�·�����ˣ���ǰ�����ҽ����һ��������Ŀ���Ƿ��ǵ�ǰ�׶ε�·������Ŀ������ǣ��򷵻���ĿID
'      �����ҽ�ִ��һ�ε���Ŀ������ʱ�ض��Ѳ���������Ӿ͵���·������Ŀ��
'������
'      str������ĿIDs= 'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
'      udtPati-������Ϣ
'      lngDay-��ǰƥ����ǵڼ����ҽ��
'      rsStepAdvice-�������е�ҽ��
'      rsStepAdvice,���ǰһ�׶κ͵�ǰ�Ľ׶���ͬ����Ϊǰһ�׶κ͵�ǰ�׶�ҽ���ļ��ϣ�����Ϊ��ǰ�׶�ҽ��
'      bln��ҩ�䷽=��ҩ�䷽�����������ݲ������õ������޸ĵ���ҩ��������
'      byt��Ч-��ͬ��Ч��ͬһ������Ŀ���ֱ���ͬһ�׶εĲ�ͬ·����Ŀʱ�ϸ�����Чƥ��
'���أ�·����ĿID�ͷ�������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim k As Long
    Dim strTmp As String
    Dim str֤��IDs As String
    Dim lng�Ķ���ҩζ�� As Long, dbl��ҩζ�� As Double
    Dim lng��ҩζ�� As Long
    Dim i As Long
    Dim arrtmp As Variant
    Dim blnTmp As Boolean
    Dim blnƥ����Ч As Boolean
    
    str���� = ""
    If str������ĿIDs = "0" Then
    '����¼���ҽ���̶�����·������Ŀ
        CheckPathInItem = 0
    Else
        '��ҩ;��������Ϊ��ȡ��ͬ���õ�ԭ��ʵ��ʹ��ʱ���Ƕ���ĸ�ҩ;��������ֻ�ж�ҩƷ��ͬ����
        'blnƥ����Ч ���øò�����Ч��һ�µ�����·������Ŀ
        blnƥ����Ч = CBool(zlDatabase.GetPara("ƥ��ʱ��Ч��ͬ��·������Ŀ", glngSys, P�ٴ�·��Ӧ��, "0"))
        lng��ҩζ�� = Val(zlDatabase.GetPara("��ҩ�䷽�����޸ĵ���ҩζ������", glngSys, P�ٴ�·��Ӧ��, "30"))
        
        If Not bln��ҩ�䷽ Then
            strSql = "Select ����, ·����Ŀid,������Ŀids,ִ�з�ʽ,��Ч " & vbNewLine & _
                    "From (Select ����, ·����Ŀid, ��id, f_List2str(Cast(Collect(To_Char(������Ŀid)) As t_Strlist)) As ������Ŀids,ִ�з�ʽ,��Ч " & vbNewLine & _
                    "       From (Select c.·����Ŀid, b.����, d.������Ŀid, Nvl(d.���id, d.Id) ��id, d.���,b.ִ�з�ʽ,d.��Ч" & vbNewLine & _
                    "              From  �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
                    "              Where b.�׶�id = [2] And b.Id = c.·����Ŀid And c.ҽ������id = d.Id" & vbNewLine & _
                    "                    And Not Exists(Select 1 From ������ĿĿ¼ E Where D.������ĿID = E.ID And E.��� = 'E' And  E.�������� In('2','3','4','6'))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From ������ĿĿ¼ E Where D.������ĿID = E.ID And E.��� In('G','F','D') And D.���ID<>0 )" & vbNewLine & _
                    "              Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
                    "       Group By ����, ·����Ŀid, ��id,ִ�з�ʽ,��Ч)" & vbNewLine & _
                    IIf(InStr(str������ĿIDs, ",") > 0, "Where instr(������Ŀids,',')>0 ", "Where (������Ŀids = [1] or instr(','||������Ŀids||',',','||[1]||',')>0)") & IIf(blnƥ����Ч, " And ��Ч =[3]", "")
        
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckPathInItem", str������ĿIDs, lng�׶�ID, byt��Ч)
            
            '����ж��·����Ŀ����ֻȡ��һ��
            If rsTmp.RecordCount > 0 Then
                '�����ֻ��ִ��һ�εģ����ж�֮ǰ�Ƿ��Ѿ�ƥ��
                If rsTmp!ִ�з�ʽ & "" = "4" Then
                    Do While Not rsTmp.EOF
                        '�ų��Ѿ�ƥ���
                        strTmp = rsTmp!������Ŀids & ""
                        rsStepAdvice.Filter = "·����ĿID=" & rsTmp!·����ĿID
                        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.MoveFirst
                        For k = 0 To rsStepAdvice.RecordCount - 1
                            'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
                            If Not (rsStepAdvice!��� & "" = "E" And InStr(",2,3,4,6,", "," & rsStepAdvice!�������� & ",") > 0) _
                                And Not (InStr(",G,F,D,", "," & rsStepAdvice!��� & ",") > 0 And Val(rsStepAdvice!���id & "") <> 0) _
                                And Val(rsStepAdvice!������ĿID & "") <> 0 Then
                                strTmp = Replace("," & strTmp & ",", "," & rsStepAdvice!������ĿID & ",", ",")
                                strTmp = Mid(strTmp, 2)
                                If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            End If
                            rsStepAdvice.MoveNext
                        Next
                        If InStr(str������ĿIDs, ",") > 0 Then
                            arrtmp = Split(str������ĿIDs, ",")
                            blnTmp = True
                            For i = 0 To UBound(arrtmp)
                                If InStr("," & strTmp & ",", "," & arrtmp(i) & ",") = 0 Then
                                    blnTmp = False
                                    Exit For
                                End If
                            Next
                            If blnTmp Then
                                CheckPathInItem = rsTmp!·����ĿID
                                str���� = rsTmp!����
                                Exit Function
                            End If
                        Else
                            If strTmp = str������ĿIDs Or InStr("," & strTmp & ",", "," & str������ĿIDs & ",") > 0 And strTmp <> "" Then
                                CheckPathInItem = rsTmp!·����ĿID
                                str���� = rsTmp!����
                                If Not blnƥ����Ч Then
                                    '���ڶ������¸�����Ч����ƥ��,����ƥ��������ĿID����Ч����ͬ�ģ�Ϊ�˱�֤��ͬһ�׶�,ͬһҩƷ,��ͬ��Ŀ,��Ч��һ�µ������,������ƥ�䣩
                                    For i = 1 To rsTmp.RecordCount
                                        If rsTmp!��Ч & "" = byt��Ч & "" Then
                                            CheckPathInItem = rsTmp!·����ĿID
                                            str���� = rsTmp!����
                                            Exit For
                                        End If
                                        rsTmp.MoveNext
                                    Next
                                End If
                                Exit Do
                            End If
                        End If
                        '����Ѿ�ƥ�䣬���������һ��ƥ�����Ŀ
                        rsTmp.MoveNext
                    Loop
                Else
                    If InStr(str������ĿIDs, ",") > 0 Then
                        '�����Ŀ�ж�ʱ���������Ŀ������˳�����������һ����·�������ôһ�����·�����
                        arrtmp = Split(str������ĿIDs, ",")
                        Do While Not rsTmp.EOF
                            blnTmp = True
                            For i = 0 To UBound(arrtmp)
                                If InStr("," & rsTmp!������Ŀids & ",", "," & arrtmp(i) & ",") = 0 Then
                                    blnTmp = False
                                    Exit For
                                End If
                            Next
                            If blnTmp Then
                                CheckPathInItem = rsTmp!·����ĿID
                                str���� = rsTmp!����
                                Exit Function
                            End If
                            
                            rsTmp.MoveNext
                        Loop
                    Else
                        CheckPathInItem = rsTmp!·����ĿID
                        str���� = rsTmp!����
                        If Not blnƥ����Ч Then
                            '���ڶ������¸�����Ч����ƥ��,����ƥ��������ĿID����Ч����ͬ�ģ�Ϊ�˱�֤��ͬһ�׶�,ͬһҩƷ,��ͬ��Ŀ,��Ч��һ�µ������,������ƥ�䣩
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!��Ч & "" = byt��Ч & "" Then
                                    CheckPathInItem = rsTmp!·����ĿID
                                    str���� = rsTmp!����
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            End If
        Else
            'ƥ��ʱ�ſ������ϵ�֤��
            str֤��IDs = Get֤��IDs(udtPati.����ID, udtPati.��ҳID)
            strSql = "Select ����, ·����Ŀid,������Ŀids,ִ�з�ʽ" & vbNewLine & _
                    "From (Select ����, ·����Ŀid, ��id, f_List2str(Cast(Collect(To_Char(������Ŀid)) As t_Strlist)) As ������Ŀids,ִ�з�ʽ" & vbNewLine & _
                    "       From (Select c.·����Ŀid, b.����, d.������Ŀid, Nvl(d.���id, d.Id) ��id, d.���,b.ִ�з�ʽ" & vbNewLine & _
                    "              From  �ٴ�·����Ŀ B, �ٴ�·��ҽ�� C, ·��ҽ������ D" & vbNewLine & _
                    "              Where b.�׶�id = [1] And b.Id = c.·����Ŀid And c.ҽ������id = d.Id" & vbNewLine & _
                    "              And Exists(Select 1 From ������ĿĿ¼ E Where d.������Ŀid = e.Id And e.��� = '7') " & vbNewLine & _
                    IIf(str֤��IDs <> "", " And (Instr(',' || [2] || ',', ',' || d.�����Ŀid || ',') > 0 Or d.�����Ŀid Is Null)", "") & vbNewLine & _
                    "              Order By b.����, b.��Ŀ���, d.���)" & vbNewLine & _
                    "       Group By ����, ·����Ŀid, ��id,ִ�з�ʽ)" & vbNewLine
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckPathInItem", lng�׶�ID, str֤��IDs)
            Do While Not rsTmp.EOF
                If rsTmp!������Ŀids & "" <> "" Then
                    '����Ķ�����ҩ
                    dbl��ҩζ�� = (UBound(Split(rsTmp!������Ŀids & "", ",")) + 1) * lng��ҩζ�� / 100
                    lng�Ķ���ҩζ�� = 0
                    '���ң��䷽�����ҩ
                    For i = 0 To UBound(Split(str������ĿIDs, ","))
                        If InStr("," & rsTmp!������Ŀids & ",", "," & Split(str������ĿIDs, ",")(i) & ",") = 0 Then
                            lng�Ķ���ҩζ�� = lng�Ķ���ҩζ�� + 1
                        End If
                    Next
                    '�����䷽�е���ҩ������ȱ�ٵ�
                    If rsTmp!������Ŀids & "" <> "" Then
                        For i = 0 To UBound(Split(rsTmp!������Ŀids & "", ","))
                            If InStr("," & str������ĿIDs & ",", "," & Split(rsTmp!������Ŀids & "", ",")(i) & ",") = 0 Then
                                lng�Ķ���ҩζ�� = lng�Ķ���ҩζ�� + 1
                            End If
                        Next
                    End If
                    '���������ķ�Χ֮�ڣ���ƥ��ɹ����������ƥ��
                    If lng�Ķ���ҩζ�� <= dbl��ҩζ�� Then
                        CheckPathInItem = rsTmp!·����ĿID
                        str���� = rsTmp!����
                        Exit Do
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetPathRows(ByRef rsAdvice As Recordset, ByRef udtPati As TYPE_Pati, Optional ByVal lngDay As Long, Optional ByRef lng�׶�ID As Long, _
            Optional ByRef rsStepAdvice As Recordset) As Boolean
'���ܣ������ٴ�·���м�����е���Ϣ����ʾ
'������lngDay-��ǰƥ����ǵڼ����ҽ��
'      rsStepAdvice,���ǰһ�׶κ͵�ǰ�Ľ׶���ͬ����Ϊǰһ�׶κ͵�ǰ�׶�ҽ���ļ��ϣ�����Ϊ��ǰ�׶�ҽ��
'���أ�lng�׶�ID����ǰƥ��Ľ׶�ID
    Dim k As Long, lngBegin As Long, lngEnd As Long, lngRow As Long
    Dim str������ĿIDs As String, lng·����ĿID As Long, str���� As String
    Dim blnOut As Boolean
    Dim lngRecord As Long
    Dim strSql As String, rsTmp As Recordset
    Dim bln��ҩ�䷽ As Boolean
    Dim byt��Ч As Byte
    
    Do While Not rsAdvice.EOF
        'һ��ҽ������һ��,ҩƷ����ҩ����ƥ�� 65982
        If Val(rsAdvice!���id & "") = 0 And Not (rsAdvice!��� & "" = "E" And rsAdvice!�������� & "" = "2") Or InStr(",5,6,", "," & rsAdvice!��� & ",") > 0 Then
            '��һ����ҩ������һ��ʱ��ֻ�������õ�ǰ�У���Ϊ·������Ŀ���ܺ�·������Ŀһ����ҩ
            lngRecord = rsAdvice.AbsolutePosition
            If InStr(",5,6,", "," & rsAdvice!��� & ",") > 0 Then
                'ҩƷ����ҩ����ƥ�� 65982
                rsAdvice.Filter = "ID=" & rsAdvice!ID
            Else
                rsAdvice.Filter = "ID=" & rsAdvice!ID & " Or ���ID=" & rsAdvice!ID
            End If
            '����¼���ҽ��������·����������
            str������ĿIDs = ""
            bln��ҩ�䷽ = False
            If Val(rsAdvice!Ӥ�� & "") = 0 Then
                For k = 0 To rsAdvice.RecordCount - 1
                    'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
                    If Not (rsAdvice!��� & "" = "E" And InStr(",2,3,4,6,", "," & rsAdvice!�������� & ",") > 0) _
                        And Not (InStr(",G,F,D,", "," & rsAdvice!��� & ",") > 0 And Val(rsAdvice!���id & "") <> 0) _
                        And Val(rsAdvice!������ĿID & "") <> 0 Then
                        str������ĿIDs = str������ĿIDs & "," & rsAdvice!������ĿID
                        If rsAdvice!��� & "" = "7" Then bln��ҩ�䷽ = True
                        byt��Ч = Val(rsAdvice!��Ч & "")
                    End If
                    rsAdvice.MoveNext
                Next
                str������ĿIDs = Mid(str������ĿIDs, 2)
                If str������ĿIDs <> "" Then
                    lng·����ĿID = CheckPathInItem(str������ĿIDs, str����, udtPati, lngDay, lng�׶�ID, rsStepAdvice, bln��ҩ�䷽, byt��Ч)
                    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
                    For k = 0 To rsAdvice.RecordCount - 1
                        If lng·����ĿID = 0 Then
                            blnOut = True
                        Else
                            rsAdvice!·����ĿID = lng·����ĿID
                            rsAdvice!·����Ŀ���� = str����
                            rsStepAdvice.Filter = "ID=" & rsAdvice!ID & " And ·����ĿID = 0"
                            rsStepAdvice!·����ĿID = lng·����ĿID
                            rsStepAdvice!·����Ŀ���� = str����
                            rsStepAdvice.Update
                        End If
                        rsAdvice.Update
                        rsAdvice.MoveNext
                    Next
                    rsAdvice.Filter = 0
                    '���ü�¼����λ��
                    rsAdvice.AbsolutePosition = lngRecord
                    If InStr(",5,6,", "," & rsAdvice!��� & ",") > 0 Then
                        '��ҩ�ĸ�ҩ��ʽͬ���޸�
                        rsAdvice.Filter = "ID=" & rsAdvice!���id
                        rsAdvice!·����ĿID = lng·����ĿID
                        rsAdvice!·����Ŀ���� = str����
                        rsStepAdvice.Filter = "ID=" & rsAdvice!ID & " And ·����ĿID = 0"
                        If rsStepAdvice.RecordCount > 0 Then
                            rsStepAdvice!·����ĿID = lng·����ĿID
                            rsStepAdvice!·����Ŀ���� = str����
                            rsStepAdvice.Update
                        End If
                    End If
                End If
            End If
            rsAdvice.Filter = 0
            '���ü�¼����λ��
            rsAdvice.AbsolutePosition = lngRecord
        End If
        
        rsAdvice.MoveNext
    Loop
    
    If blnOut Then
        If InStr(GetInsidePrivs(P�ٴ�·��Ӧ��), ";·������Ŀ;") = 0 Then
            MsgBox "��û�����·������Ŀ��Ȩ�ޣ������Զ�����֮ǰ��·����Ŀ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    SetPathRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetPathRichEPR(ByRef rsRichEPR As Recordset, ByRef rsStepRichEPR As Recordset, ByVal lng�׶�ID As Long) As Boolean
'���ܣ��Զ�ƥ���Ѿ����ɲ���
'������rsStepRichEPR=һ���׶εĲ�����rsRichEPR=��ǰ�׶εĲ���
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim k As Long, strTmp As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    Do While Not rsRichEPR.EOF
        blnDo = False
        strSql = "Select ����, ·����Ŀid,�ļ�ids,ִ�з�ʽ " & _
                " From (Select ����, ·����Ŀid, f_List2str(Cast(Collect(To_Char(�ļ�id)) As t_Strlist)) As �ļ�ids, ִ�з�ʽ" & vbNewLine & _
                " From (Select b.Id As ·����Ŀid, a.�ļ�id, b.����, b.ִ�з�ʽ" & vbNewLine & _
                "       From �ٴ�·������ A, �ٴ�·����Ŀ B" & vbNewLine & _
                "       Where a.��Ŀid = b.Id And �׶�id = [2])" & vbNewLine & _
                " Group By ����, ·����Ŀid, ִ�з�ʽ)" & _
                " Where �ļ�ids = [1] or instr(','||�ļ�ids||',',','||[1]||',')>0"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�Զ�ƥ�䲡��", Val(rsRichEPR!�ļ�ID & ""), lng�׶�ID)
        Do While Not rsTmp.EOF
            If rsTmp!ִ�з�ʽ & "" = "4" Then
                '�ų��Ѿ�ƥ���
                strTmp = rsTmp!�ļ�ids & ""
                rsStepRichEPR.Filter = "·����ĿID=" & rsTmp!·����ĿID
                If rsStepRichEPR.RecordCount > 0 Then rsStepRichEPR.MoveFirst
                For k = 0 To rsStepRichEPR.RecordCount - 1
                    strTmp = Replace("," & strTmp & ",", "," & rsStepRichEPR!�ļ�ID & ",", ",")
                    strTmp = Mid(strTmp, 2)
                    If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                    rsStepRichEPR.MoveNext
                Next
                If strTmp = rsRichEPR!�ļ�ID & "" Or InStr("," & strTmp & ",", "," & rsRichEPR!�ļ�ID & ",") > 0 And strTmp <> "" Then
                    blnDo = True
                    Exit Do
                End If
            Else
                blnDo = True
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
        If blnDo Then
            rsRichEPR!·����Ŀ���� = rsTmp!���� & ""
            rsRichEPR!·����ĿID = Val(rsTmp!·����ĿID & "")
            rsStepRichEPR.Filter = "ID=" & rsRichEPR!ID & " And ·����ĿID=0"
            If rsStepRichEPR.RecordCount > 0 Then
                rsStepRichEPR!·����Ŀ���� = rsTmp!���� & ""
                rsStepRichEPR!·����ĿID = Val(rsTmp!·����ĿID & "")
                rsStepRichEPR.Update
            End If
            rsRichEPR.Update
        End If
        rsRichEPR.MoveNext
    Loop
    SetPathRichEPR = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFirstType(ByRef udtPP As TYPE_PATH_Pati, Optional ByVal bytFunc As Byte = 0)
'���ܣ����һ����Ŀ��û�У�������ݿ���ȡ��һ������
'����:bytFunc=0  -ȱʡȡ��һ������,1-����ƥ�京��ҽ�����ؼ��ֵķ���,ƥ�䲻�ϲ�ȡȱʡ�ĵ�һ������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strFirstName As String
    Dim strType As String
    
    On Error GoTo errH
    If bytFunc = 0 Then
        strSql = "Select ���� from �ٴ�·������ where ·��ID=[1] and �汾��=[2] and NVL(��֧ID,0)=0 And ���=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ȡ��һ������", udtPP.·��ID, udtPP.�汾��)
        If rsTmp.RecordCount > 0 Then GetFirstType = rsTmp!���� & ""
    ElseIf bytFunc = 1 Then
        strSql = "Select ����,��� from �ٴ�·������ where ·��ID=[1] and �汾��=[2] and NVL(��֧ID,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ȡ·������", udtPP.·��ID, udtPP.�汾��)
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!��� & "") = 1 Then strFirstName = rsTmp!���� & ""
            If InStr(rsTmp!���� & "", "ҽ��") > 0 Then
                strType = rsTmp!���� & ""
                Exit For
            End If
            rsTmp.MoveNext
        Next
        GetFirstType = IIf(strType = "", strFirstName, strType)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePathItem(ByVal dateCur As Date, ByVal DateInPath As Date, udtPati As TYPE_Pati, udtPP As TYPE_PATH_Pati, _
    ByVal lng·����¼ID As Long, ByRef colSQL As Collection) As Boolean
'����:��������ҽ����������·����Ŀ
'����:
'���:
    
'����
'   colSQL: ���ؿ�ִ�е�SQL
'����ֵ:
'   T-���óɹ�;F����ʧ��
'
    Dim strSql As String
    Dim strAdivcePathOut As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rsStepAdvice As Recordset    '���ǰһ���׶κ͵�ǰ�׶�һ�������ҽ������һ����¼���У������ж�ִ�з�ʽ=4�ģ����ظ�ƥ��
    Dim rsStepRichEPR As Recordset
    Dim rsRichEPR As Recordset
    Dim lngǰһ�׶�ID As Long
    Dim lng�׶�ID As Long
    Dim str����ԭ�� As String      '�������������ԭ��
    Dim strFirstType As String
    Dim strAdviceType As String
    Dim i As Long, j As Long
    Dim AddDate As Date

    AddDate = dateCur
    For i = 1 To Int(dateCur) - Int(DateInPath)
        strSql = "Select a.Id,a.Ӥ��, a.���id,a.������ĿID, b.���, b.��������,a.ҽ����Ч as ��Ч " & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                "Where a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And NVL(A.Ӥ��,0)=0" & vbNewLine & _
                "     And (a.��ʼִ��ʱ�� Between [3] And [4] And a.ҽ����Ч = 1 Or a.ҽ����Ч = 0 And [3] Between Trunc(a.��ʼִ��ʱ��) And Trunc(Nvl(a.ִ����ֹʱ��, [4]))) And a.ҽ��״̬ Not In(-1,4)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ƥ��·����Ŀ", udtPati.����ID, udtPati.��ҳID, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1 / 24 / 60 / 60)
        Set rsAdvice = MakePathItems
        Do While Not rsTmp.EOF
            rsAdvice.AddNew
            rsAdvice!ID = rsTmp!ID
            rsAdvice!Ӥ�� = Val(rsTmp!Ӥ�� & "")
            rsAdvice!���id = Val(rsTmp!���id & "")
            rsAdvice!������ĿID = Val(rsTmp!������ĿID & "")
            rsAdvice!��� = rsTmp!��� & ""
            rsAdvice!�������� = rsTmp!�������� & ""
            rsAdvice!��Ч = Val(rsTmp!��Ч & "")
            rsTmp.MoveNext
        Loop
        If rsAdvice.RecordCount > 0 Then rsAdvice.Update
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        'ƥ����Ŀ
        lngǰһ�׶�ID = lng�׶�ID
        lng�׶�ID = 0
        strAdivcePathOut = ""
        '��ȡƥ��׶Σ�Ĭ�������С��
        strSql = "Select ID" & vbNewLine & _
                "      From (Select ID" & vbNewLine & _
                "              From �ٴ�·���׶�" & vbNewLine & _
                "            Where ·��id = [1] And �汾�� = [2] And [3] Between ��ʼ���� And Nvl(��������, ��ʼ����) And ��֧id Is Null And" & vbNewLine & _
                "                  ��id Is Null" & vbNewLine & _
                "            Order By ���)" & vbNewLine & _
                "      Where Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CreatePathItem", udtPP.·��ID, udtPP.�汾��, i)
        If Not rsTmp.EOF Then
            lng�׶�ID = Val(rsTmp!ID & "")
        Else
            Exit For
        End If
        
        '����׶���ͬ�����¼�½׶ε�ҽ��
        If lng�׶�ID <> lngǰһ�׶�ID Then
            Set rsStepAdvice = MakePathItems
        End If
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        rsStepAdvice.Filter = 0
        Do While Not rsAdvice.EOF
            rsStepAdvice.AddNew
            rsStepAdvice!ID = rsAdvice!ID
            rsStepAdvice!Ӥ�� = Val(rsAdvice!Ӥ�� & "")
            rsStepAdvice!���id = Val(rsAdvice!���id & "")
            rsStepAdvice!������ĿID = Val(rsAdvice!������ĿID & "")
            rsStepAdvice!��� = rsAdvice!��� & ""
            rsStepAdvice!�������� = rsAdvice!�������� & ""
            rsStepAdvice!·����ĿID = Val(rsAdvice!·����ĿID & "")
            rsStepAdvice!·����Ŀ���� = rsAdvice!·����Ŀ���� & ""
            rsAdvice.MoveNext
        Loop
        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.Update
        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.MoveFirst
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        If Not SetPathRows(rsAdvice, udtPati, i, lng�׶�ID, rsStepAdvice) Then Exit Function
        
        Set rsPathAdvice = MakePathAdivceRS
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        Do While Not rsAdvice.EOF
            '·������Ŀ(Ӥ��ҽ��������¼��ҽ��������·������Ŀ,Ҳ����Ϊ·��������Ŀ)
            If Val(rsAdvice!·����ĿID & "") = 0 Then
                If Val(rsAdvice!Ӥ�� & "") = 0 And Val(rsAdvice!������ĿID & "") <> 0 Then
                    strAdivcePathOut = strAdivcePathOut & "," & rsAdvice!ID
                End If
            Else
                rsPathAdvice.Filter = "·����ĿID = " & rsAdvice!·����ĿID
                If rsPathAdvice.RecordCount = 0 Then
                    rsPathAdvice.AddNew
                    rsPathAdvice!·����ĿID = rsAdvice!·����ĿID & ""
                    rsPathAdvice!·����Ŀ���� = rsAdvice!·����Ŀ���� & ""
                    rsPathAdvice!ҽ��IDs = rsAdvice!ID & ""
                Else
                    rsPathAdvice!ҽ��IDs = rsPathAdvice!ҽ��IDs & "," & rsAdvice!ID
                End If
                rsPathAdvice.Update
            End If
            rsAdvice.MoveNext
        Loop
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        '�ٴ�·������(Ҫ������ʱ���ύ֮����Ϊ�����Լ��)
        rsPathAdvice.Filter = ""
        For j = 1 To rsPathAdvice.RecordCount
            strSql = "Zl_����·������_Insert(1," & udtPati.����ID & "," & udtPati.��ҳID & ",Null,0," & _
                lng·����¼ID & "," & lng�׶�ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                ",'" & rsPathAdvice!·����Ŀ���� & "'," & rsPathAdvice!·����ĿID & ",'" & rsPathAdvice!ҽ��IDs & "',Null,Null" & _
                ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
                rsPathAdvice.MoveNext
            colSQL.Add strSql, "C" & colSQL.count + 1
            '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
            AddDate = AddDate + 1 / 24 / 60 / 60
        Next
        '·������Ŀ
        If strAdivcePathOut <> "" Then
            '��������ı���ԭ�����
            If str����ԭ�� = "" Then
                strSql = "Select ���� From ���쳣��ԭ�� Where (����='����' Or ����='����')  And ����=1 And rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CreatePathItem")
                If rsTmp.RecordCount > 0 Then
                    str����ԭ�� = rsTmp!���� & ""
                End If
            End If
            strAdivcePathOut = Mid(strAdivcePathOut, 2)
            If strAdviceType = "" Then strAdviceType = GetFirstType(udtPP, 1)
            strSql = "Zl_����·������_Insert(1," & udtPati.����ID & "," & udtPati.��ҳID & ",Null,0," & _
                lng·����¼ID & "," & lng�׶�ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                ",'" & strAdviceType & "',Null" & ",'" & strAdivcePathOut & "',Null,Null" & _
                ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'·������Ŀ'" & _
                ",1,Null,Null,Null,'" & str����ԭ�� & "' ,1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
            AddDate = AddDate + 1 / 24 / 60 / 60
        End If
        
        'ƥ�䲡��
        strSql = "select ID,�ļ�ID from ���Ӳ�����¼ Where ����id = [1] And ��ҳid = [2] And �ļ�ID is not Null And ����ʱ�� Between [3] And [4] And NVL(Ӥ��,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ƥ��·����Ŀ", udtPati.����ID, udtPati.��ҳID, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1 / 24 / 60 / 60)
        Set rsRichEPR = MakePathRichEPR
        If lng�׶�ID <> lngǰһ�׶�ID Then
            Set rsStepRichEPR = MakePathRichEPR
        End If
        Do While Not rsTmp.EOF
            rsRichEPR.AddNew
            rsRichEPR!ID = rsTmp!ID
            rsRichEPR!�ļ�ID = rsTmp!�ļ�ID
            rsStepRichEPR.AddNew
            rsStepRichEPR!ID = rsTmp!ID
            rsStepRichEPR!�ļ�ID = rsTmp!�ļ�ID
            rsTmp.MoveNext
        Loop
        If rsRichEPR.RecordCount > 0 Then rsRichEPR.MoveFirst
        If Not SetPathRichEPR(rsRichEPR, rsStepRichEPR, lng�׶�ID) Then Exit Function
        If rsRichEPR.RecordCount > 0 Then
            rsRichEPR.Filter = "·����ĿID<>0"
            If rsRichEPR.RecordCount > 0 Then rsRichEPR.MoveFirst
            For j = 1 To rsRichEPR.RecordCount
                strSql = "Zl_����·������_Insert(1," & udtPati.����ID & "," & udtPati.��ҳID & ",Null,0," & _
                    lng·����¼ID & "," & lng�׶�ID & _
                    ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                    ",'" & rsRichEPR!·����Ŀ���� & "'," & rsRichEPR!·����ĿID & ",Null,Null,Null" & _
                    ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1," & rsRichEPR!ID & ")"
                rsRichEPR.MoveNext
                colSQL.Add strSql, "C" & colSQL.count + 1
                '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
                AddDate = AddDate + 1 / 24 / 60 / 60
            Next
        End If
        
        '���һ����Ŀ��û�У�������һ��������Ŀ
        If strAdivcePathOut = "" And rsPathAdvice.RecordCount = 0 And rsRichEPR.RecordCount = 0 Then
            If strFirstType = "" Then strFirstType = GetFirstType(udtPP, 0)
            strSql = "Zl_����·������_Insert(1," & udtPati.����ID & "," & udtPati.��ҳID & ",NULL," & udtPati.����ID & "," & _
                    lng·����¼ID & "," & lng�׶�ID & _
                     ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                    ",'" & strFirstType & "',Null" & ",Null,Null,Null" & _
                    ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'δ�����κ���Ŀ'" & _
                    ",Null,'�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��',Null,Null,'',1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
            AddDate = AddDate + 1 / 24 / 60 / 60
        End If
        '����
        strSql = "Zl_����·������_Insert(1," & lng·����¼ID & "," & lng�׶�ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & ",'" & _
                UserInfo.���� & "'," & IIf(strAdivcePathOut = "", "0", "1") & ",'','" & UserInfo.���� & "','" & UserInfo.���� & "','',0,Null,Null" & ",Null,1" & ")"
                
        colSQL.Add strSql, "C" & colSQL.count + 1
    Next
    
    CreatePathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
