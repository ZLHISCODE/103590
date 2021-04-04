Attribute VB_Name = "mdlShiftBase"
Option Explicit

Public gobjPublicAdvice As Object           '�ٴ���������
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetShiftType(ByVal bytType As Byte, ByVal strDeptID As String) As ADODB.Recordset
'��ͬ���ҿ�������ͬ���Ƶ�ֵ����
    
    On Error GoTo errH
    Select Case bytType
        Case 1 '��ȡ���е�ֵ������Ϣ
            gstrSQL = "Select b.���� ����, a.ֵ���� �������, To_Char(a.��ʼʱ��, 'hh24:mi') ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') ����ʱ��" & vbNewLine & _
                    "From ҽ��ֵ���� a, ���ű� b Where a.����id = b.Id And b.id In(Select * From Table(f_Str2list([1]))) Order By a.��ʼʱ��"
        Case 2 'ֻ��ȡ���е�ֵ��������
            gstrSQL = "Select Distinct a.ֵ���� ������� From ҽ��ֵ���� a, ���ű� b Where a.����id = b.Id " & vbNewLine & _
                "And b.id In(Select * From Table(f_Str2list([1])))"
    End Select
    Set GetShiftType = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ϣ", strDeptID)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "��ȡ�����Ϣ"
End Function

Public Function GetDeptName(ByVal strDeptID As String) As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "Select ���� ||'-' || ���� as ����, Id,���� From ���ű� Where Id In (Select * From Table(f_Str2list([1]))) Order by ����"
    Set GetDeptName = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ϣ", strDeptID)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "��ȡ��������"
End Function

Public Function GetPatientType() As ADODB.Recordset

    On Error GoTo errH
    gstrSQL = "Select ���, ����, ˳��,��ȡSQL From ҽ�����Ӱಡ������ Where �Ƿ�ͣ�� = 0 And ��ȡsql Is Not Null Order By ˳��"
    Set GetPatientType = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ϣ")
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "��ȡ����������Ϣ"
End Function

Public Function GetTimeRangePati(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal strDeptID As String) As ADODB.Recordset
'��ȡһ��ʱ�䷶Χ�ڵĻ�����Ϣ
    Dim rsTemp As ADODB.Recordset
        
    Set rsTemp = GetPatiType
    gstrSQL = ""
    Do While Not rsTemp.EOF
        gstrSQL = IIf(gstrSQL = "", "", gstrSQL & vbNewLine & "Union All ") & rsTemp!��ȡSQL
        rsTemp.MoveNext
    Loop
    gstrSQL = UCase(gstrSQL)
    gstrSQL = Replace(gstrSQL, "[��ʼʱ��]", zlStr.To_Date(dtBegin))
    gstrSQL = Replace(gstrSQL, "[����ʱ��]", zlStr.To_Date(dtEnd))
    gstrSQL = Replace(gstrSQL, "[����ID]", "[1]")
    Set GetTimeRangePati = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", strDeptID)
End Function

Public Function GetPatiType() As ADODB.Recordset
'��ȡ����������Ϣ

    On Error GoTo errH
    gstrSQL = "Select ���,����,˳��, ��ȡsql From ҽ�����Ӱಡ������ Where ��ȡsql Is Not Null"
    gstrSQL = gstrSQL & vbNewLine & "Union All Select '��Ժ','��Ժ����',98,'Select Distinct ''��Ժ'' ����, a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.��Ժ���� ����, a.סԺ�� ��ʶ��, a.��Ժ���� ��Ժʱ��, a.��Ժ��ʽ, a.��Ժ����id" & vbNewLine & _
        "From ������ҳ a" & vbNewLine & _
        "Where a.��Ժ���� >  [��ʼʱ��] And" & vbNewLine & _
        "      a.��Ժ���� <=  [����ʱ��] And a.�������� In(0,2) And a.��Ժ����id In (Select /*+cardinality(a,10)*/* From Table(f_Str2list([����ID] )) a) ' From Dual"
    gstrSQL = gstrSQL & " Order By ˳��"
    Set GetPatiType = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϢSQL")
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, "��ȡ����������Ϣ"
End Function

Public Sub ReadSignSource(ByVal lng��¼ID As Long, strSource As String)
'���ܣ���ȡ���Ӱ�������ڵ���ǩ��/��֤��Դ������
'������
'lng��¼ID ���Ӱ�ļ�¼id��dtBegin��¼�Ŀ�ʼʱ��
'���أ�ǩ��/��֤ǩ����Դ�����ɹ���
'      strSource=ǩ��/��֤ǩ���Ľ��Ӱ�Դ��
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String
    
    On Error GoTo errH
    
    gstrSQL = "Select a.��¼Id, a.����id, a.����ҽ��, a.������, a.���࿪ʼʱ��, a.�������ʱ��, a.�Ӱ�ҽ��, a.�Ӱ���, a.�Ӱ࿪ʼʱ��, a.�Ӱ����ʱ��, a.��¼��, b.����id, b.���, b.��������," & vbNewLine & _
        "       b.����id, b.��ҳid, b.����, b.�Ա�, b.����, b.����, b.��ʶ��, b.��Ժʱ��, b.��Ժ��ʽ, b.��������" & vbNewLine & _
        "From ҽ�����Ӱ��¼ a, ҽ�����Ӱ����� b" & vbNewLine & _
        "Where a.��¼Id = b.��¼id And a.��¼id = [1]" & vbNewLine & _
        "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlShiftBase", lng��¼ID)

    strField = "��¼ID,����ID,����ҽ��,������,���࿪ʼʱ��,�������ʱ��,�Ӱ�ҽ��,�Ӱ���,�Ӱ࿪ʼʱ��,�Ӱ����ʱ��,��¼��," & _
        "����ID,���,��������,����ID,��ҳID,����,�Ա�,����,����,��ʶ��,��Ժʱ��,��Ժ��ʽ,��������"
    arrField = Split(strField, ",")
        
    '����ҽ��ǩ��Դ��
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If IsDate(rsTmp.Fields(arrField(i)).Value) Then
                strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
            Else
                strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value & ""
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetUserInfo(ByVal strName As String) As ADODB.Recordset
'�������������û���Ϣ
    
    On Error GoTo errH
    gstrSQL = "Select b.��Աid, b.�û���, a.���, a.���� ,c.����id From ��Ա�� a, �ϻ���Ա�� b, ������Ա c Where a.Id = b.��Աid And a.Id = c.��Աid And a.���� = [1]"
    
    Set GetUserInfo = zlDatabase.OpenSQLRecord(gstrSQL, "mdlShiftBase", strName)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "��ȡ�û���Ϣ"
End Function

Public Function GetCA(ByVal strName As String) As Boolean
'�����û����ж��Ƿ����õ���ǩ����ͨ��������ȡ����ID
    Dim rsTemp As ADODB.Recordset
    Dim lngCA As Long
    
    On Error GoTo errH
    Set rsTemp = GetUserInfo(strName)
    If rsTemp.RecordCount = 0 Then Exit Function
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) ����ǩ�� From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ǩ�����ò���", 1, rsTemp!����ID)
    If rsTemp.RecordCount > 0 Then
        lngCA = Val(NVL(rsTemp!����ǩ��, 0))
    Else
        lngCA = 0
    End If
    GetCA = IIf(lngCA = 1, True, False)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "�Ƿ����õ���ǩ��Ϣ"
End Function

Public Function Get����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    '��ȡ��������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim lngTmp As Long
    
    '���ߣ���ȡ�ϰ没��
    strSQL = "Select a.id From ���Ӳ�����¼ A, ���Ӳ�����ʽ B" & vbNewLine & _
            "Where a.Id = b.�ļ�id and a.����id=[1] and a.��ҳid=[2] And (a.�������� like '%��Ժ��¼' or a.�������� like '%��Ժ����' or a.�������� like '%���Ժ��¼' or a.�������� like '%��Ժ������¼')" & vbNewLine & _
            "And b.�ı����� Is Not Null order by ���ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lng����ID, lng��ҳID)
    
    Do While Not rsTmp.EOF
        strTmp = Sys.ReadLobV2("���Ӳ�����ʽ", "�ı�����", "�ļ�id=[1]", "", Val(rsTmp!id & ""))
        If strTmp <> "" Then
            strTmp = Replace(Replace(strTmp, Chr(10), ""), Chr(13), "")
            lngTmp = InStr(strTmp, "�����ߡ�")
            If lngTmp > 0 Then
                lngTmp = lngTmp + 4
                strTmp = Mid(strTmp, lngTmp, InStr(lngTmp, strTmp, "��") - lngTmp)
                strTmp = Replace(strTmp, "��  �ߣ�", "")
                strTmp = Replace(strTmp, "��  ��", "")
                If strTmp <> "" Then
                    Exit Do
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '�ϰ�ûȡ����ȡ�°没��
    If strTmp = "" Then
        strTmp = GetItemAppendByEmr("��������", lng����ID, lng��ҳID)
    End If
    
    
    '�ж����߲��ܴ���50���ַ�
    If strTmp <> "" And zlCommFun.ActualLen(strTmp) > 50 Then
        strTmp = Mid(strTmp, 1, 25)
    End If
    
    Get���� = strTmp
End Function

Public Function GetItemAppendByEmr(ByVal str������ As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵȡ��Ӳ����л�ȡ����ֵ
    Dim strText As String
    Dim intType As Integer
    Dim lng����ID As Long
    
    On Error Resume Next
    
    If gobjEmr Is Nothing Then Exit Function
    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing: Exit Function
 
    If Not gobjEmr Is Nothing Then
        strText = gobjEmr.GetOrderInspectInfoEx(2, lng����ID, lng��ҳID, str������)
        If Err.Number <> 0 Then
            strText = gobjEmr.GetOrderInspectInfo(lng����ID, str������)
        End If
    End If
    
    Err.Clear
    GetItemAppendByEmr = strText
End Function

Public Function GetAdviceDiag(ByVal lngҽ��ID As Long, Optional ByRef str��� As String) As String
'���ܣ����ҽ����Ӧ�������Ϣ
'������str���=������ϵ���������ַ���
'���أ�������ϵ�ID�����ŷָ�
    Dim rsTmp As Recordset, strSQL As String
    Dim strReturn As String
    
    strSQL = "Select  A.ID,a.�������" & vbNewLine & _
            "From ������ϼ�¼ A, �������ҽ�� B" & vbNewLine & _
            "Where b.���id=a.id And  b.ҽ��ID=[1]" & vbNewLine & _
            "Order By b.rowID"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��������", lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str��� = str��� & "," & rsTmp!�������
            strReturn = strReturn & "," & rsTmp!id
            rsTmp.MoveNext
        Loop
        str��� = Mid(str���, 2)
        strReturn = Mid(strReturn, 2)
    End If
    GetAdviceDiag = strReturn
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetALL���(intType As Integer, ByVal lngPatiID As Long, ByVal lngPageID As Long, dtEnd As Date) As String
    '��ȡ�����������
    ' intType����������  '1,'����',2,'����',3,'һ������',4,'����',5,'��ǰ',6,'����',7,'��Ѫ',8,'Σ',9,'����',10,'Σ/��',11,'�ؼ�',12,'����'
    Dim rsTmp As ADODB.Recordset, rsTmp1 As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim str��� As String

    If intType <> 5 Then
        'ȡ��Ҫ���
        strSQL = "Select a.����id, a.�������, a.�������" & vbNewLine & _
                "From ������ϼ�¼ A" & vbNewLine & _
                "Where a.������� In (1, 2, 3, 11, 12, 13) And Nvl(a.�������, 1) = 1 And a.��ϴ��� = 1 And" & vbNewLine & _
                "      a.����id=[1] and a.��ҳid=[2]  And a.ȡ��ʱ�� Is Null" & vbNewLine & _
                "Order By a.����id Asc, a.��¼��Դ Desc, a.������� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatiID, lngPageID)
        Do While Not rsTmp.EOF
            If InStr(strTmp, rsTmp!�������) = 0 Then
                strTmp = strTmp & "," & rsTmp!�������
            End If
            rsTmp.MoveNext
        Loop
        strTmp = Mid(strTmp, 2)
        
        'ȡ������ȡ���±�������
        If strTmp = "" Then
            strSQL = "Select a.��¼��Դ, a.��ϴ���, a.�������, a.����id, a.���id, a.�������, a.��¼����, a.��¼�� From ������ϼ�¼ A Where a.����id = [1] And a.��ҳid =[2] And Nvl(a.�������, 1) = 1  And A.ȡ��ʱ�� is Null And" & vbNewLine & _
                    "a.��¼���� = (Select Max(a.��¼����) From ������ϼ�¼ A Where a.����id = [1] And a.��ҳid =[2] And A.ȡ��ʱ�� is Null And Nvl(a.�������, 1) = 1)"
    
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatiID, lngPageID)
            
            Do While Not rsTmp.EOF
                If InStr(strTmp, rsTmp!�������) = 0 Then
                    strTmp = strTmp & "," & rsTmp!�������
                End If
                rsTmp.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
        End If

        str��� = strTmp
    Else
        strSQL = "Select a.Id, a.���id, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.ҽ����Ч, a.ҽ��״̬, a.��ʼִ��ʱ��, a.�������, b.��������,b.���� as ��Ŀ����, a.У�Ի�ʿ, a.У��ʱ��,a.�걾��λ From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2] And a.������Ŀid = b.Id(+) And Nvl(a.ҽ��״̬, 0) Not In (-1,1,2, 4) And a.�������='F' and a.����ʱ��>[3] and a.У��ʱ�� is not null" & vbNewLine & _
                "Order By a.Id, a.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatiID, lngPageID, CDate(dtEnd))
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                '�Ӹ����л�ȡ��ǰ���������������Ը���Ϊ׼
                strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
                Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", Val(rsTmp!id & ""))
                If Not rsTmp1.EOF Then
                    str��� = rsTmp1!���� & ""
                Else
                    '��ȡ��ǰ���
                    Call GetAdviceDiag(Val(rsTmp!id & ""), strTmp)
                    str��� = strTmp
                End If
                
            End If
        End If
    End If
    GetALL��� = str���
End Function

Public Function GetNextId(strTable As String, Optional strFild As String) As Long
    '------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
    '������
    '   strTable��������;strFild�ֶ������������Ʋ�һ����ID�������¼ID
    '���أ�
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "������ü�¼" Or strtab = "סԺ���ü�¼" Then strtab = "���˷��ü�¼"
    If strFild <> "" Then
        strSQL = "Select " & strtab & "_" & strFild & ".Nextval From Dual"
    Else
        strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
    GetNextId = rsTmp.Fields(0).Value
'    Exit Function
'errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function InitObjPublicAdvice() As Boolean
'���ܣ���ʼ�ٴ���������
    If gobjPublicAdvice Is Nothing Then
        On Error Resume Next
        Set gobjPublicAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
        If Not gobjPublicAdvice Is Nothing Then
            Call gobjPublicAdvice.InitCommon(gcnOracle, glngSys, , , , , , gobjEmr)
        End If
        Err.Clear: On Error GoTo 0
    End If
    InitObjPublicAdvice = Not gobjPublicAdvice Is Nothing
End Function


