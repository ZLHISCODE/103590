Attribute VB_Name = "mdlPatient"
Option Explicit 'Ҫ���������
'=======ϵͳ������ر���============
Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4   '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5 '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6   'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7 'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8     'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9 '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10      '�����˻���δ�����ʱ��Ժ
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
End Enum

Public gobjPublicPatient As Object                 '������Ϣ�ӿڶ���
Public gclsInsure As New clsInsure
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gobjXWHIS As Object     'RIS�ӿڲ���zl9XWInterface.clsHISInner
Public gblnXW As Boolean      'ϵͳ������������ҽѧӰ����Ϣϵͳרҵ��ӿڡ�
Public gblnPatiByID As Boolean   'ͬһ���ֻ֤�ܶ�Ӧһ����������

Public gobjPlugIn As Object   '�������

Public gstrUnitName As String
Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String '��ǰ�û���
Public gstrPatiTypeColor As String '������ɫ��  ����,��ɫֵ|����,��ɫֵ----
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum ENUM_INSIDE_PROGRAM
    P��Լ��λ���� = 1100
    P������Ϣ���� = 1101
    P���￨���Ź��� = 1102
    PԤ������� = 1103
    PԤ��������ձ� = 1104
    P��Լ��λ���� = 1105
    P���˷������� = 1106
End Enum

'�ṹ����ַ���� 1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ
Public Enum Enum_IX_ADDRESS
    E_IX_�����ص� = 1
    E_IX_���� = 2
    E_IX_��סַ = 3
    E_IX_���ڵ�ַ = 4
    E_IX_��ϵ�˵�ַ = 5
End Enum

Public gint����ʣ��Ʊ������ As Integer      '�շ�ʱ,Ʊ����ʣ��X�ź�ʼ�����շ�Ա:-1��������
Public gobjSquare As SquareCard  '�����㲿��
'�ṹ����ַ
Public gbln���ýṹ����ַ As Boolean
Public gbln��ʾ���� As Boolean

Public Function InitPatiType() As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    gstrPatiTypeColor = ""
    gstrSQL = "select ����,��ɫ from ��������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������")
    Do Until rsTemp.EOF
        gstrPatiTypeColor = gstrPatiTypeColor & rsTemp!���� & "," & NVL(rsTemp!��ɫ, 0) & "|"
        rsTemp.MoveNext
    Loop
    If Len(gstrPatiTypeColor) > 0 Then
        gstrPatiTypeColor = Mid(gstrPatiTypeColor, 1, Len(gstrPatiTypeColor) - 1)
    Else
        gstrPatiTypeColor = "��ͨ����,0|ҽ������,255"
    End If
    InitPatiType = True
    Exit Function
errH:
    gstrPatiTypeColor = "��ͨ����,0|ҽ������,255"
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetPatiColor(ByVal strPatiType) As Long
Dim arrType As Variant, i As Integer
    arrType = Split(gstrPatiTypeColor, "|")
    For i = LBound(arrType) To UBound(arrType)
        If Split(arrType(i), ",")(0) = strPatiType Then
            GetPatiColor = Split(arrType(i), ",")(1)
            Exit Function
        End If
    Next
End Function

Public Function GetUnitID(bytFlag As Byte, lngID As Long) As Long
'���ܣ������շ��ض���Ŀ��ִ�п���
'������bytFlag=ִ�п��ұ�־,lngID=�շ�ϸĿID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '����ȷ����
            GetUnitID = UserInfo.����ID 'ȡ����Ա���ڿ���
        Case 4 'ָ������
            strSQL = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                GetUnitID = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
            Else
                GetUnitID = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
            End If
        Case 1, 2, 3 '���˿���,����Ա����
            GetUnitID = UserInfo.����ID '��ȡ����Ա����
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNOFromCard(strCardNo As String) As String
'���ܣ��ɾ��￨�Ż�ȡ���￨���ü�¼���ݺ�
'˵��������õ����Ѿ����ϣ����˲������п��ţ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select NO " & _
    " From סԺ���ü�¼ A,������Ϣ B" & _
    " Where A.����ID=B.����ID And A.ʵ��Ʊ��=B.���￨��" & _
    "       And A.��¼����=5 And A.��¼״̬=1 And A.���=1 And B.���￨��=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strCardNo)

    
    If Not rsTmp.EOF Then GetNOFromCard = rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarINFO(lngID As Long) As ADODB.Recordset
'���ܣ���ȡ��ָ��������Ϣ���ƵĲ�����Ϣ
'������lngID=����ID
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    'by lesfeng 2010-03-08 �����Ż�
    strTemp = " ����ID,�����,סԺ��,����,�Ա�,����,�ѱ�,��������,�����ص�,���֤��,���,����,����,ѧ��,ְҵ," & _
              " ����״��,��ͥ��ַ,��ͥ�绰,������λ,��λ�绰,סԺ����,��ǰ���� ��Ժʱ��,��Ժʱ��,��ǰ����ID,��ǰ����ID "
    strSQL = _
        "Select A.����ID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,A.�ѱ�," & _
        " To_Char(A.��������,'YYYY-MM-DD') as ��������,A.�����ص�,A.���֤��," & _
        " A.���,A.����,A.����,A.ѧ��,A.ְҵ,A.����״��,A.��ͥ��ַ,A.��ͥ�绰," & _
        " A.������λ,A.��λ�绰,A.סԺ����,C.���� as ����,D.���� as ����," & _
        " A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
        " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ�� From " & _
        " (Select " & strTemp & " From ������Ϣ Where ����ID<>[1]) A, " & _
        " (Select " & strTemp & " From ������Ϣ Where ����ID =[1]) B,���ű� C,���ű� D " & _
        " Where A.��ǰ����ID=C.ID(+) And A.��ǰ����ID=D.ID(+)" & _
        " And Nvl(A.����,'X')=Nvl(B.����,'X') And Nvl(A.����,'X')=Nvl(B.����,'X')" & _
        " And Nvl(A.��������,Sysdate)=Nvl(B.��������,Sysdate)" & _
        " And Nvl(A.���֤��,'X')=Nvl(B.���֤��,'X') And Nvl(A.�Ա�,'X')=Nvl(B.�Ա�,'X')" & _
        " And A.����=B.���� Order by A.����ID Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    Set SimilarINFO = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str���� As String, str���� As String, dat�������� As Date, str�Ա� As String, str���� As String, str���֤�� As String, ByRef rsRet As ADODB.Recordset) As String
'���ܣ���鲡���Ƿ����������Ϣ
'���أ����Ƽ�¼�Ĳ���ID��,��"234,235,236"
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    'by lesfeng 2010-03-08 �����Ż� TO_DATE('" & Format(dat��������, "YYYY-MM-DD") & "','YYYY-MM-DD')
    strSQL = _
        " Select Rownum+1 ID,����ID,�����,סԺ��,Nvl(���֤��,'δ�Ǽ�') ���֤��,Nvl(��ͥ��ַ,'δ�Ǽ�') ��ַ,To_Char(�Ǽ�ʱ��,'YYYY-MM-DD') �Ǽ�ʱ�� " & _
        " From ������Ϣ Where (����=[1] And ����=[2] And �Ա�=[3] And ����=[4]" & _
        " And ��������=[6]) Or ���֤��=[5] " & _
        " Order by ����ID Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str����, str����, str�Ա�, str����, str���֤��, CDate(Format(dat��������, "YYYY-MM-DD")))
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!����ID & ",�����:" & NVL(rsTmp!�����, "��") & ",סԺ��:" & NVL(rsTmp!סԺ��, "��") & ",���֤��:" & rsTmp!���֤�� & ",��ַ:" & rsTmp!��ַ & ",�Ǽ�����:" & rsTmp!�Ǽ�ʱ��
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    rsTmp.Filter = ""
    Set rsRet = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'�������û��ʹ��
Public Function GetUnionName(lngID As Long) As String
'���ܣ���ȡ��ͬ��λ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ��Լ��λ Where ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount <> 0 Then GetUnionName = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMax��ҳID(lng����ID As Long) As Long
'���ܣ���ȡ���˵���󲡰���ҳID
'���أ�
'     >0:�ɹ�
'      0:ʧ��
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ��ҳID From ������ҳ Where ����ID=[1] Order by ��ҳID Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng����ID)
    
    If rsTmp.RecordCount = 0 Then
        GetMax��ҳID = 1
    ElseIf IsNull(rsTmp!��ҳID) Then
        GetMax��ҳID = 1
    Else
        GetMax��ҳID = rsTmp!��ҳID + 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function һ�����ٴ���Ժ(lng����ID, dat��Ժʱ�� As Date) As Boolean
'���ܣ��жϲ����Ƿ���һ�����ٴ���Ժ
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �ϴγ�Ժʱ�� From ������Ϣ Where ����ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng����ID)
    
    If rsTmp.RecordCount = 0 Then
        Exit Function
    ElseIf IsNull(rsTmp!�ϴγ�Ժʱ��) Then
        Exit Function
    ElseIf Abs(CDate(Format(rsTmp!�ϴγ�Ժʱ��, "yyyy-MM-dd")) - CDate(Format(dat��Ժʱ��, "yyyy-MM-dd"))) <= 7 Then
        һ�����ٴ���Ժ = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isLookRoom(lngID As Long) As ADODB.Recordset
'���ܣ��жϿ����Ƿ�۲���
'���أ���=����
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    'by lesfeng 2010-03-08 �����Ż� Select *
    strSQL = "Select A.�������,A.ID,A.�ϼ�ID,A.����,A.����,A.����,A.λ��,A.ĩ��,A.����ʱ��,A.����ʱ��,A.վ��" & _
             "  From ���ű� A,��������˵�� B Where B.����ID=A.ID And B.�������=1 And B.��������='����' And A.ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount <> 0 Then
        Set isLookRoom = rsTmp
    Else
        Set isLookRoom = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set isLookRoom = New ADODB.Recordset
End Function

Public Function NextBedNo(lngUnitID As Long) As Long
'���ܣ���ȡָ����������һ��λ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = _
        "Select A.���� From ��λ״����¼ A " & _
        "Where A.����ID=[1] And Not Exists(Select ���� From ��λ״����¼ B Where B.����=A.����+1 And B.����ID=A.����ID) " & _
        "Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngUnitID)
    
    If rsTmp.RecordCount <> 0 Then
        NextBedNo = rsTmp!���� + 1
    Else
        NextBedNo = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isRepeat(lngUnitID As Long, strBeds As String) As String
'���ܣ��ж���ָ�������ڵ�һϵ�д����Ƿ��Ѿ�����
'������lngUnitID=����ID,strBeds=�����ַ���,��"12,13,15..."
'���أ���=��������,����"12,13..."��Щ�����ظ�
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ��λ״����¼ Where ����ID=[1] And instr(','||[2]||',',','||����||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngUnitID, strBeds)
    
    If rsTmp.RecordCount <> 0 Then
        For i = 1 To rsTmp.RecordCount
            isRepeat = isRepeat & rsTmp!���� & ","
            rsTmp.MoveNext
        Next
        isRepeat = Left(isRepeat, Len(isRepeat) - 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiNO(strסԺ�� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ��סԺ���Ƿ��Ѿ����������ݿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 1 From ������Ϣ Where סԺ��=[1] And ����ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strסԺ��, lng����ID)
    
    If rsTmp.RecordCount > 0 Then ExistInPatiNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistClinicNO(str����� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ��������Ƿ��Ѿ����������ݿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 �����Ż� Select *
    strSQL = "Select ����ID,����� From ������Ϣ Where �����=[1] And ����ID<>[2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str�����, lng����ID)
    
    If rsTmp.RecordCount > 0 Then ExistClinicNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ExistInPatiID(lngID As Long) As Boolean
'���ܣ��ж�ָ������ID�Ƿ��Ѿ����������ݿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 �����Ż� Select *
    strSQL = "Select ����ID From ������Ϣ Where ����ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount > 0 Then ExistInPatiID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastInfo(lngID As Long) As String
'���ܣ���ȡ�������һ��Ԥ���λ��Ϣ
'���أ�"�ɿλ|��λ������|��λ�ʺ�"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��ģ���¼����=1
    strSQL = "Select �ɿλ,��λ������,��λ�ʺ� From ����Ԥ����¼ " & _
            " Where (�ɿλ is Not NULL Or ��λ������ is Not NUll Or ��λ�ʺ� is Not NULL) And ��¼����=1 And ����ID=[1] Order by �տ�ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If Not rsTmp.EOF Then
        GetLastInfo = IIf(IsNull(rsTmp!�ɿλ), "", rsTmp!�ɿλ) & "|" & IIf(IsNull(rsTmp!��λ������), "", rsTmp!��λ������) & "|" & IIf(IsNull(rsTmp!��λ�ʺ�), "", rsTmp!��λ�ʺ�)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitSysPar()
'���ܣ���ʼ��ϵͳ����
    Dim strValue As String
    On Error Resume Next
        
    gbln������ = zlDatabase.GetPara(3, glngSys) = "1"
                    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '������ʾ��ʽ
    gblnShowCard = Not ISPassShowCard 'zldatabase.GetPara(12, glngSys) = "0"
    
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    'gbyt�ſ� = Val(Split(strValue, "|")(0))
    gbytԤ�� = Val(Split(strValue, "|")(1))
    gbytCardNOLen = Val(Split(strValue, "|")(4))
    'If gbyt�ſ� = 0 Then gbyt�ſ� = 7
    If gbytԤ�� = 0 Then gbytԤ�� = 7
    If gbytCardNOLen = 0 Then gbytCardNOLen = 7
    
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBillԤ�� = Mid(strValue, 2, 1) = "1"
    'gblnBill�ſ� = Mid(strValue, 5, 1) = "1"
    
    '���￨�������ĸǰ׺
    gstrCardMask = UCase(zlDatabase.GetPara(27, glngSys))
    
    'һ��ͨ������֤
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gbytԤ��������鿨 = Val(Split(strValue, "|")(0))
    gbln���ѿ��˷��鿨 = zlDatabase.GetPara(282, glngSys) = "1"
    
    '��Ժ�Ǽ�ʱˢ����������
    gblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 5, 1) = "1"
    
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    
    gblnXW = Val(zlDatabase.GetPara(255, glngSys)) = 1
    'ͬһ���ֻ֤�ܶ�Ӧһ����������
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1
End Sub

Public Function ISPassShowCard() As Boolean
'���ܣ��Ƿ�������ʾ���￨��
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnPassShowCard As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select �������� From ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ����")
    If Not rsTemp.EOF Then
        blnPassShowCard = NVL(rsTemp!��������) <> ""
    End If
    
    ISPassShowCard = blnPassShowCard
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SaveIDCard(bytStyle As Byte, strNO As String, lng����ID As Long, lng��ҳID As Long, _
        lng���˲���ID As Long, lng���˿���ID As Long, str��ʶ�� As String, str�ѱ� As String, _
        strԭ���� As String, str���� As String, str�Ա� As String, str���� As String, str���� As String, str���� As String, _
        curӦ�ս�� As Currency, curʵ�ս�� As Currency, str���㷽ʽ As String, Dat����ʱ�� As Date, lng����ID As Long, rsMoney As ADODB.Recordset, ByVal strICCard As String) As String
'���ܣ�����һ�����￨���ü�¼SQL���
'������bytStyle=0-����,1-����,2-����
'      cur���=���￨���
'      str���㷽ʽ=���Ϊ��,��ʾ����,�����ֽ�
'      rsMoney:�������￨�շ���Ϣ�ļ�¼��
'      strԭ����=������ʱ��
'      lng����ID=��ǰ���õľ��￨����ID
'      strICCard=IC����,ͨ����IC����ʽ����ʱ,ͬʱ��д������Ϣ��IC���ֶ�
    Dim lngUnitID As Long
    Dim strSQL As String
    
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
    Select Case rsMoney!���ұ�־
        Case 4 'ָ������
            lngUnitID = GetUnitID(rsMoney!���ұ�־, rsMoney!�շ�ϸĿID)
        Case 1, 2 '���˿���
            If lng���˿���ID <> 0 Then
                lngUnitID = lng���˿���ID
            Else
                lngUnitID = UserInfo.����ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.����ID
    End Select
    
    '���ù���"zl_���￨��¼_Insert"
    strSQL = "zl_���￨��¼_INSERT(" & bytStyle & ",'" & strNO & "'," & lng����ID & "," & lng��ҳID & "," & _
        str��ʶ�� & ",'" & str�ѱ� & "','" & UCase(strԭ����) & "','" & str���� & "','" & str���� & "','" & str���� & _
        "','" & str�Ա� & "','" & str���� & "'," & lng���˲���ID & "," & lng���˿���ID & "," & rsMoney!�շ�ϸĿID & _
        ",'" & rsMoney!�շ���� & "','" & IIf(IsNull(rsMoney!���㵥λ), "", rsMoney!���㵥λ) & "'," & _
        rsMoney!������ĿID & ",'" & rsMoney!�վݷ�Ŀ & "'," & curӦ�ս�� & "," & lngUnitID & "," & UserInfo.����ID & _
        ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(OverTime(Dat����ʱ��), "1", "0") & _
        ",To_Date('" & Format(Dat����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & str���㷽ʽ & "'," & IIf(lng����ID = 0, "NULL", lng����ID) & ",'" & strICCard & "'," & curӦ�ս�� & "," & curʵ�ս�� & ")"
    
    SaveIDCard = strSQL
End Function

Public Sub InitLocPar(ByVal lngModul As Long)
'���ܣ���ʼ��ģ�����
    Dim strValue As String
    On Error Resume Next

    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = 0, "%", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
    
    If lngModul = P������Ϣ���� Or lngModul = P���￨���Ź��� Or lngModul = PԤ������� Then
        gstr�ſ�ID = zlDatabase.GetPara("���þ��￨����", glngSys, lngModul, "")
'        glngԤ��ID = zldatabase.GetPara("����Ԥ��Ʊ������", glngSys, lngModul, 0)
        'LED��������
        gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
        gblnLedWelcome = Val(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, lngModul, 1)) <> 0
        gbln���� = zlDatabase.GetPara("���Ѽ���", glngSys, lngModul) = "1"
    End If
    
    If lngModul = P������Ϣ���� Then
        gblnMustCard = zlDatabase.GetPara("����ͬʱ���뷢��", glngSys, lngModul) = "1"
        gbln���ýṹ����ַ = Val(zlDatabase.GetPara("���˵�ַ�ṹ��¼��", glngSys)) <> 0
        gbln��ʾ���� = Val(zlDatabase.GetPara("�����ַ�ṹ��¼��", glngSys)) <> 0
    ElseIf lngModul = P���￨���Ź��� Then
    
    ElseIf lngModul = PԤ������� Then
        gblnShowHave = zlDatabase.GetPara("���������Ľɿ", glngSys, lngModul) = "1"
        gblnAllowOut = zlDatabase.GetPara("�����Ժ���˽�סԺԤ��", glngSys, lngModul) = "1"
        gblnBanIn = zlDatabase.GetPara("��ֹ��Ժ���˽�����Ԥ��", glngSys, lngModul) = "1"
        gbln�ɿ���� = zlDatabase.GetPara("������Ľɿ����", glngSys, lngModul) = "1"
        strValue = Trim(zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, lngModul, "0|10"))
        gbln��վ����ʾ = zlDatabase.GetPara("Ԥ�����վ����ʾ", glngSys, lngModul) = "1"
        '37372
        If Val(Split(strValue & "|", "|")(0)) = 0 Then
            gint����ʣ��Ʊ������ = -1
        Else
            gint����ʣ��Ʊ������ = Val(Split(strValue & "|", "|")(1))     '����:26948
        End If
    End If
End Sub

Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'���ܣ���ȡ�����б��ѡ��ĵ���
'������
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select ���� as ID,����,����,���� From ����" & _
                 " Where (���� Like [1] Or upper(����) Like '" & gstrLike & "'||[1]||'%' Or ���� Like '" & gstrLike & "'||[1]||'%') And  NVL(����,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select ���� as ID,����,����,���� From ���� Where  NVL(����,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnAuditedFee(lng����ID As Long, Optional ByVal bln���� As Boolean = True, Optional ByVal bytPrepayType As Byte = 0) As Currency
'���ܣ�bln����=true:��ȡ����δ��˵Ļ��۵����ʷ���
'      �����ȡ����δ�ɻ��۽��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If bytPrepayType = 0 Then
        strSQL = _
        " Select Sum(Nvl(���,0)) as ��� " & _
        " From (Select Sum(Nvl(ʵ�ս��,0)) as ��� " & _
        "       From ������ü�¼" & _
        "       Where ���ʷ���=[2] And ��¼״̬=0 And ����ID=[1]" & _
        "       Union ALL " & _
        "       Select Sum(Nvl(ʵ�ս��,0)) as ��� " & _
        "       From סԺ���ü�¼" & _
        "       Where ���ʷ���=[2] And ��¼״̬=0 And ����ID=[1] ) "
    ElseIf bytPrepayType = 1 Then
        strSQL = _
        " Select Sum(Nvl(ʵ�ս��,0)) as ��� " & _
        "       From ������ü�¼" & _
        "       Where ���ʷ���=[2] And ��¼״̬=0 And ����ID=[1]"
    Else
        strSQL = _
        " Select Sum(Nvl(ʵ�ս��,0)) as ��� " & _
        "       From סԺ���ü�¼" & _
        "       Where ���ʷ���=[2] And ��¼״̬=0 And ����ID=[1]"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, IIf(bln����, 1, 0))
    If Not rsTmp.EOF Then
        GetUnAuditedFee = Val("" & rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strDict = "����" Then
        strSQL = "Select ����,����,0 as ȱʡ From " & strDict & " Order by ����"
    Else
        strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPatient")
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInsure(strNO As String) As Integer
'���ܣ��ж��շѼ�¼���Ƿ����ָ����ҽ�����㷽ʽ
'������strNO=�շѵ��ݺ�
'���أ��������,�򷵻ظ�ҽ�����㷽ʽ���㵱ʱ�ı�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.���� From ����Ԥ����¼ A,���ս����¼ B" & _
        " Where A.��¼����=1 And A.��¼״̬=1 And A.NO=[1]" & _
        " And A.ID=B.��¼ID And B.����=3"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then
        ExistInsure = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistFeeInsurePatient(lng����ID As Long) As Boolean
'���ܣ��ж�ҽ�������Ƿ����δ�����
'���أ�
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSQL = "Select Nvl(sum(B.�������),0) ������� From ������Ϣ A,������� B Where A.����ID=B.����ID And Nvl(A.����,0)<>0 And A.����ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng����ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!������� <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveSpare(strNO As String) As Double
'���ܣ�����Ԥ�����ݺ��жϲ����Ƿ���Ԥ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(B.Ԥ�����,0) as ���" & _
        " From ����Ԥ����¼ A,������� B" & _
        " Where A.��¼����=1 And A.��¼״̬ IN(1,3) and nvl(A.Ԥ�����,2)=b.����(+)" & _
        " And A.NO=[1] And A.����ID=B.����ID"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then HaveSpare = Val(NVL(rsTmp!���))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBalance(strNO As String) As Double
'���ܣ�����Ԥ�����ݺ��жϸõ����Ƿ񱻽���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Sum(Nvl(��Ԥ��,0)) as ��Ԥ��" & _
        " From ����Ԥ����¼ Where NO=[1] And ��¼���� IN(1,11)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then HaveBalance = Val(NVL(rsTmp!��Ԥ��))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function StrToNum(ByVal strNumber As String) As Double
    '����:���ַ���ת��������
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function

Public Function zlIsExistsSquareCard(ByRef lngԤ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����õ����Ƿ�Ϊ�����㵥��
    '���:strNo-Ԥ�����ĵ��ݺ�
    '����:����,�򷵻�true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 17:02:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    strSQL = "Select 1 From ���˿������¼ A Where a.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ԥ�����Ƿ����ˢ����¼", lngԤ��ID)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub
Public Function zl_Getҽ�ƿ�����(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '����:����ҽ������ID��ȡҽ������
    '���:lngTypeID-ҽ�ƿ�����ID
    '����:���Ͷ���
    '����:����
    '����:2012-07-06
    '�����:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSQL = "" & _
    "       Select ���볤��,������������,�Ƿ�ȱʡ���� " & _
    "       From ҽ�ƿ���� " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�ƿ����", lngTypeId)
    If rsTemp Is Nothing Then zl_Getҽ�ƿ����� = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Getҽ�ƿ����� = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = NVL(rsTemp!���볤��, "0")
    arr(1) = NVL(rsTemp!������������, "0")
    arr(2) = NVL(rsTemp!�Ƿ�ȱʡ����, "0")
    zl_Getҽ�ƿ����� = arr
End Function

Public Function �Ƿ��Ѿ�ǩԼ(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ҫ�󶨵Ŀ����Ƿ��Ѿ�ǩԼ
    '���:�󶨿���
    '����:����
    '����:2012-08-31 11:32:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng���֤���ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    
    lng���֤���ID = Getҽ�ƿ����ID("�������֤")
    strSQL = "" & _
    "   Select Count(1) as �Ƿ�ǩԼ From ����ҽ�ƿ���Ϣ Where ����=[1] And �����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ���", strCardNo, lng���֤���ID)
    �Ƿ��Ѿ�ǩԼ = rsTemp!�Ƿ�ǩԼ > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function


Public Sub AddSQL�󶨿�(ByVal lng����ID As Long, �����ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�󶨿�����
    '���:lng����ID;strCard-�󶨿���;strPassWord-��������
    '����:lngCard����ID-���ѵĽ���ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str�䶯ԭ�� As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str�䶯ԭ�� = "���˹Һŷ���"
          'Zl_ҽ�ƿ��䶯_Insert
          strSQL = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSQL = strSQL & "" & 11 & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSQL = strSQL & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSQL = strSQL & "" & �����ID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSQL = strSQL & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSQL = strSQL & "NULL)"
     zlAddArray cllPro, strSQL
End Sub

Public Function Getҽ�ƿ����ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����ID
    '���:strTypeName ҽ�ƿ��������
    '����:ҽ�ƿ����ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select ID From ҽ�ƿ���� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ����", strTypeName)
    If rsTemp Is Nothing Then Getҽ�ƿ����ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Getҽ�ƿ����ID = 0: Exit Function
    Getҽ�ƿ����ID = rsTemp!ID
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zl��ǰ�û����֤�Ƿ��(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�û����֤�Ƿ��ѱ���
    '���:lng����ID
    '����:True �Ѱ� false δ��
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng���֤���ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    lng���֤���ID = Getҽ�ƿ����ID("�������֤")
    strSQL = "" & _
    " Select count(1) as �Ƿ�� From ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.���֤�� =B.���� And A.����ID=B.����ID And A.����ID=[1] And B.�����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ���", lng����ID, lng���֤���ID)
    zl��ǰ�û����֤�Ƿ�� = rsTemp!�Ƿ�� > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub
Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = zlDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "", Optional ByVal datCalc As Date) As String
    '����:����Ϸ��Լ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        If datCalc = CDate(0) Then
            strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Else
            strSQL = "select Zl_Age_Check([1],[2],[3]) From dual"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay), datCalc)
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = NVL(rsTemp.Fields(0).Value)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePublicPatient() As Boolean
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set gobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        If gobjPublicPatient Is Nothing Then
            MsgBox "����������Ϣ��������(zlPublicPatient.clsPublicPatient)ʧ��!", vbInformation, gstrSysName
        Else
            Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPublicPatient Is Nothing Then CreatePublicPatient = True
End Function

Public Sub LoadStructAddressDef(ByRef strAddress() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������е�ȱʡ��ַ
    '���:PatiAddress-�ṹ����ַ�ؼ�
    '����:
    '����:��ΰ��
    '����:2016/1/7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select ����,����,level From ���� " & _
            " Start With ȱʡ��־=1 " & _
            " Connect by Prior �ϼ�����=���� " & _
            " Order by level Desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȱʡ����")
    If rsTmp.RecordCount = 0 Then Exit Sub
    Do While Not rsTmp.EOF
        strAddress(Val(NVL(rsTmp!����))) = NVL(rsTmp!����)
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReadStructAddress(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef PatiAddress As Object)
'����:��ȡ�ṹ����ַ
    Dim i As Long
    Dim rsStruct As ADODB.Recordset
    Dim rsAddress As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select a.ʡ, a.��, a.��, a.����, a.����, a.��ַ��� From ���˵�ַ��Ϣ A Where a.����id = [1] And NVL(a.��ҳid,0) = [2]"
    Set rsStruct = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˽ṹ����ַ", lng����ID, lng��ҳID)
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        rsStruct.Filter = "��ַ���=" & i
        If rsStruct.RecordCount > 0 Then
            Call PatiAddress(i).LoadStructAdress(rsStruct!ʡ & "", rsStruct!�� & "", rsStruct!�� & "", rsStruct!���� & "", rsStruct!���� & "")
        Else
            If rsAddress Is Nothing Then
                'ͬһ������ֻ��ȡһ��
                If lng��ҳID <> 0 Then
                    strSQL = "Select c.�����ص�, c.����, Nvl(b.��ͥ��ַ, c.��ͥ��ַ) As ��סַ, Nvl(b.���ڵ�ַ, c.���ڵ�ַ) As ���ڵ�ַ, Nvl(b.��ϵ�˵�ַ, c.��ϵ�˵�ַ) As ��ϵ�˵�ַ" & vbNewLine & _
                        "From ������ҳ B, ������Ϣ C" & vbNewLine & _
                        "Where b.����id = c.����id And b.����id = [1] And b.��ҳid = [2] "
                Else
                    strSQL = "Select c.�����ص�, c.����, c.��ͥ��ַ As ��סַ,  c.���ڵ�ַ As ���ڵ�ַ,c.��ϵ�˵�ַ As ��ϵ�˵�ַ " & vbNewLine & _
                        "From ������Ϣ C" & vbNewLine & _
                        "Where c.����id = [1] "
                End If
                Set rsAddress = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˽ṹ����ַ", lng����ID, lng��ҳID)
            End If
            If rsAddress.RecordCount > 0 Then
                If NVL(rsAddress.Fields(PatiAddress(i).Tag).Value, "") <> "" Then
                    PatiAddress(i).Value = NVL(rsAddress.Fields(PatiAddress(i).Tag).Value, "")    '�������ýṹ����ַ֮ǰ������
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub CreateStructAddressSQL(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef varSQL As Variant, ByRef PatiAddress As Object, Optional ByVal bytFunc As Byte = 0)
'����:�����ṹ����ַSQL
'����:
'PatiAddress-�ṹ����ַ�ؼ�������
'varSQL-���ص�SQL���鼯��\�����Ǽ��϶���
'bytFunc ��ѡ����:=1 ���ؼ�ֵΪ��ʱ,����ɾ��
    Dim i As Long
    Dim strSQL As String
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        If PatiAddress(i).Value <> "" Then
            '����\�޸�
            strSQL = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & "," & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & "," & i & ",'" & PatiAddress(i).valueʡ & "','" & PatiAddress(i).value�� & "','" & PatiAddress(i).value���� & "','" & PatiAddress(i).value���� & "','" & PatiAddress(i).value��ϸ��ַ & "','" & PatiAddress(i).Code & "')"
            If IsArray(varSQL) Then
                ReDim Preserve varSQL(UBound(varSQL) + 1)
                varSQL(UBound(varSQL)) = strSQL
            ElseIf UCase(TypeName(varSQL)) = UCase("Collection") Then
                varSQL.Add strSQL, "K" & (varSQL.Count + 1)
            End If
        Else
            'ɾ��
            If bytFunc = 1 Then
                strSQL = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & "," & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & "," & i & ")"
                If IsArray(varSQL) Then
                    ReDim Preserve varSQL(UBound(varSQL) + 1)
                    varSQL(UBound(varSQL)) = strSQL
                ElseIf UCase(TypeName(varSQL)) = UCase("Collection") Then
                    varSQL.Add strSQL, "K" & (varSQL.Count + 1)
                End If
            End If
        End If
    Next
End Sub

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ���(zl9XWInterface.clsHISInner) �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject(, "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String, Optional ByRef strErr As String = "0")
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    Dim strMsg As String
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        strMsg = "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description
        If strErr = "0" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            strErr = strMsg
        End If
    End If
End Sub
