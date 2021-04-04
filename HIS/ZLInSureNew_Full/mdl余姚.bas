Attribute VB_Name = "mdl��Ҧ"
Option Explicit
    
'Public Declare Function f_Init Lib "dhpDLL.DLL" () As Integer
'Public Declare Function f_Close Lib "dhpDLL.DLL" () As Integer
'Public Declare Function f_Apply Lib "dhpDLL.DLL" (ByVal lngTradeTypeID As Long, _
'    ByVal dblTradeID As Double, ByVal strData As String, ByRef strMessage As String) As Integer
'cd 50223
'VVVV
Private Declare Function f_UserBargaingInit Lib "BargaingApply.DLL" (ByVal strData As String, ByVal strMessage As String) As Integer
Private Declare Function f_UserBargaingClose Lib "BargaingApply.DLL" (ByVal strData As String, ByVal strMessage As String) As Integer
Private Declare Function f_UserBargaingApply Lib "BargaingApply.DLL" (ByVal lngTradeTypeID As Integer, ByVal dblTradeID As Double, ByVal strData As String, ByVal strMessage As String) As Integer

Public gstrOutput��Ҧ As String, gstrInput��Ҧ As String, gcn��Ҧ As New ADODB.Connection, gstrIC���� As String
Private mstrBillNo As String, mcur��ҽ�� As Currency, mstr��ˮ�� As String
Private mblnInit As Boolean

Public Function makeBillNO(lng����ID As Long) As String
    Dim datCurr As Date
    datCurr = zlDatabase.Currentdate
    makeBillNO = toHex(CDbl(Format(datCurr, "yyyyMMddHHmmss") & lng����ID), 36)
End Function

Public Function makeICInfo(lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '����IC����
    gstrSQL = "Select A.����,B.����,B.�Ա�,A.��λ���� From �����ʻ� A,������Ϣ B Where A.����ID=[1]" & _
        " And A.����=[2] And A.����ID=B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ҧ)
    If rsTemp.EOF Then
        MsgBox "û���ҵ��ò��˵������Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    makeICInfo = Space(18 - LenB(StrConv(rsTemp(0), vbFromUnicode))) & rsTemp(0) & _
                 String(18, "0") & _
                 Space(20 - LenB(StrConv(rsTemp(1), vbFromUnicode))) & rsTemp(1) & _
                 Space(2 - LenB(StrConv(rsTemp(2), vbFromUnicode))) & rsTemp(2) & _
                 String(56, "0") & _
                 Space(10 - LenB(StrConv(rsTemp(3), vbFromUnicode))) & rsTemp(3) & _
                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
'    makeICInfo = Right(Space(18) & rsTemp(0), 18) & _
'                 String(18, "0") & _
'                 RightB(Space(20) & rsTemp(1), 20) & _
'                 RightB(Space(2) & rsTemp(2), 2) & _
'                 String(56, "0") & _
'                 Right(Space(10) & rsTemp(3), 10) & _
'                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
               'cd 050223  Right(Space(20) & rsTemp(1), 20) & _
               'cd 050223  Right(Space(2) & rsTemp(2), 2) & _

End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = IIf(dblMod <> 0, dblMod, "") & strTemp
    End If
    toHex = strTemp
End Function

Public Function CheckReturn_��Ҧ() As Boolean
    If glngReturn < 0 Then
        If Split(gstrOutput��Ҧ, "$$")(1) = "" Then
            MsgBox "����ҽ������ʱ��������", vbInformation, gstrSysName
        Else
            MsgBox "ҽ�������������´���" & vbCrLf & "    " & Split(gstrOutput��Ҧ, "$$")(1), vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CheckReturn_��Ҧ = True
End Function

Public Function ���뽻����ˮ_��Ҧ(str�������� As String) As String
    Dim strTemp As String
    ���뽻����ˮ_��Ҧ = ""
    strTemp = "$$" & str�������� & "$$"
    gstrOutput��Ҧ = Space(4000) 'cd 050223
    glngReturn = f_UserBargaingApply(23, 0, strTemp, gstrOutput��Ҧ)
    If CheckReturn_��Ҧ() = False Then Exit Function
    ���뽻����ˮ_��Ҧ = CDbl(Split(gstrOutput��Ҧ, "$$")(2))
End Function

Public Function openConn��Ҧ() As Boolean
    Dim rsTemp As New ADODB.Recordset, strServer As String, strUser As String, strPass As String, _
        strTemp As String, strDatabase As String
    On Error GoTo ErrH
    If gcn��Ҧ.State <> adStateOpen Then
        '���ȶ���������������
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ)
        Do Until rsTemp.EOF
            strTemp = Nvl(rsTemp("����ֵ"), "")
            Select Case rsTemp("������")
                Case "��Ҧ������"
                    strServer = strTemp
                Case "��Ҧ�û���"
                    strUser = strTemp
                Case "��Ҧ�û�����"
                    strPass = strTemp
                Case "��Ҧ���ݿ�"
                    strDatabase = strTemp
            End Select
            rsTemp.MoveNext
        Loop
    
        On Error Resume Next
        gcn��Ҧ.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & strDatabase & ";Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn��Ҧ.CursorLocation = adUseClient
        gcn��Ҧ.Open
        If Err.Number <> 0 Then
            MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
            openConn��Ҧ = False
            Exit Function
        End If
        On Error GoTo ErrH
    End If
    openConn��Ҧ = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    openConn��Ҧ = False
End Function

Public Function ҽ����ʼ��_��Ҧ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    If openConn��Ҧ() = False Then
        ҽ����ʼ��_��Ҧ = False
        Exit Function
    End If
    
    If mblnInit = False Then
        gstrInput��Ҧ = "$$$$"
        gstrOutput��Ҧ = Space(4000)
        
        glngReturn = f_UserBargaingInit(gstrInput��Ҧ, gstrOutput��Ҧ)
        
        ҽ����ʼ��_��Ҧ = CheckReturn_��Ҧ()
        If ҽ����ʼ��_��Ҧ Then mblnInit = True
    Else
        ҽ����ʼ��_��Ҧ = True
    End If
'cd 050223
'    gstrInput��Ҧ = "$$$$": gstrOutput��Ҧ = "$$$$$$"
''    glngReturn = f_UserBargaingInit(gstrInput��Ҧ, gstrOutput��Ҧ)
'    ҽ����ʼ��_��Ҧ = CheckReturn_��Ҧ()
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    ҽ����ʼ��_��Ҧ = False
End Function

Public Function ҽ����ֹ_��Ҧ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    On Error GoTo errHandle
    Set gcn��Ҧ = Nothing
    'cd 050223 VVVVVV
    gstrInput��Ҧ = "$$$$"
    gstrOutput��Ҧ = Space(4000)

    glngReturn = f_UserBargaingClose(gstrInput��Ҧ, gstrOutput��Ҧ)
    '^^^^^^^^^^^^^^^
'    glngReturn = f_Close()
    ҽ����ֹ_��Ҧ = CheckReturn_��Ҧ()
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ҽ����ֹ_��Ҧ = False
End Function

Public Function ҽ������_��Ҧ() As Boolean
    ҽ������_��Ҧ = frmSet��Ҧ.��������()
End Function

Public Function �����������_��Ҧ(rs������ϸ As Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strҽ���� As String, lng����ID As Long, datCurr As Date, rsTemp As New ADODB.Recordset, str��Ŀ���� As String, _
        str����ID As String, str���� As String, strSQL As String, strTemp As String, iLoop As Long, str������ϸ��ˮ�� As String, _
        strҽ������ As String, str��ϸ���� As String, str��Ŀ���� As String, str��� As String, str�Ը����� As String
    WriteInfo vbCrLf & "����Ԥ����"
    On Error GoTo errHandle
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷�����ϸ�����ܽ���ҽ������", vbInformation, gstrSysName
        Exit Function
    End If
    rs������ϸ.MoveFirst
    lng����ID = rs������ϸ!����ID
    datCurr = zlDatabase.Currentdate
    
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ҧ

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        str����ID = rsTemp!����
        str���� = rsTemp!����
    Else
        �����������_��Ҧ = False
        Exit Function
    End If
    
    mstrBillNo = makeBillNO(lng����ID)
    gstrSQL = "Select * From �����ʻ� Where ����=" & TYPE_��Ҧ & " And ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, gstrSysName)
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    
    If rsTemp.EOF Then
        MsgBox "û���ҵ��ò��˵�ҽ����Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    strҽ���� = rsTemp!����
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(29)
    If mstr��ˮ�� = "" Then Exit Function
    'д������
    strSQL = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr��ˮ�� & "','" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & _
        "',1," & Trim(gstrҽԺ����) & ",'" & strҽ���� & "','" & lng����ID & "',Null,Null,'" & rs������ϸ!������ & _
        "','" & str���� & "','" & str����ID & "',Null,'" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & _
        "','" & UserInfo.���� & "')"
    WriteInfo "дǰ�û���������:" & strSQL
    gcn��Ҧ.Execute strSQL
    mcur��ҽ�� = 0
    iLoop = 1
    strSQL = "select max(SerialNum) as SerialNum From "
    strSQL = strSQL & "(Select SerialNum From hi_upClinicPrescription  "
    strSQL = strSQL & " union all"
    strSQL = strSQL & " select SerialNum From hi_ClinicPrescription ) A"
    Set rsTemp = gcn��Ҧ.Execute(strSQL)
    If rsTemp.EOF Then
        str������ϸ��ˮ�� = 0
    Else
        str������ϸ��ˮ�� = Nvl(rsTemp(0), 0)
    End If
    str������ϸ��ˮ�� = AddNum(str������ϸ��ˮ��)
    
    While Not rs������ϸ.EOF
        'ȡ�շ���ϸ
        gstrSQL = "Select ����,����,���,nvl(���,'') as ��� From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs������ϸ!�շ�ϸĿID))
        str��ϸ���� = rsTemp!����: str��Ŀ���� = rsTemp!����
        str��� = Left(Left(rsTemp!��� & " |", InStr(rsTemp!��� & " |", "|") - 1), InStr(rsTemp!��� & " |", " ") - 1)
        '�ж���Ŀ����
        str��Ŀ���� = IIf(rsTemp!��� = "5" Or rsTemp!��� = "6", "ҩƷ", IIf(rsTemp!��� = "7", "��ҩ", "����"))
        ''''cd 2005 0301
        '������ҩƷ�У�������Ϊ���ƣ����Բ���ֱ�Ӱ�rstemp!������ж�
        gstrSQL = "select * from ҩƷĿ¼ where ҩƷID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs������ϸ!�շ�ϸĿID))
        
        If Not rsTemp.EOF Then '�����ҩƷĿ¼���д���Ŀ���ٸ���ҩƷĿ¼�ĸ�ע���ж��Ƿ�Ϊ����
           If Nvl(rsTemp!��ʶ��, "ҩƷ") = "����" Then str��Ŀ���� = "����"
        End If
        
        '^^^^^^^^^^^^^^^

        
        '�ӱ���֧����Ŀ�в����Ƿ��и�ҽ����Ŀ
        gstrSQL = "Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ Where �Ƿ�ҽ��=1 And ����=[1] And �շ�ϸĿID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, CLng(rs������ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then      'û����Ŀ����
           
            If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "1"
            Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "0"
            End If
        Else            '�и���Ŀʱ����
            str��ϸ���� = rsTemp!��Ŀ����
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str��ϸ���� & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str��ϸ���� & "'"
            
            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
            
            If rsTemp.EOF Then      '���ҽ������Ŀ¼��δ�ҵ�����Ŀ
                If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "1"
                Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "0"
                End If
            Else        '���ҽ������Ŀ¼���и�ҩƷ
                strҽ������ = IIf(rsTemp!zfbl = 0, "����", IIf(rsTemp!zfbl = 1, "����", "����"))
                str�Ը����� = rsTemp!zfbl
            End If
        End If
        strSQL = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
            "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & str������ϸ��ˮ�� & "," & Trim(gstrҽԺ����) & ",'" & mstr��ˮ�� & _
            "','" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','" & mstr��ˮ�� & "'," & IIf(str��Ŀ���� = "����", 2, 1) & ",'" & _
            str��ϸ���� & "','" & str��Ŀ���� & "','" & str��� & "'," & Format(rs������ϸ!ʵ�ս�� / rs������ϸ!����, "#.###") & "," & _
            rs������ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
        WriteInfo "������ϸ(д������ϸ):" & strSQL
        gcn��Ҧ.Execute strSQL
        str������ϸ��ˮ�� = AddNum(str������ϸ��ˮ��)
        
        'gstrSQL = "ZL_����֧����Ŀ_Modify(" & rs������ϸ!�շ�ϸĿID & "," & TYPE_��Ҧ & ",NULL,'" & str��ϸ���� & "','" & _
            str��Ŀ���� & "','" & strҽ������ & "',1)"
'        WriteInfo "�޸ı���֧����Ŀ:" & gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        

'
        If str�Ը����� = 1 Then mcur��ҽ�� = mcur��ҽ�� + rs������ϸ!ʵ�ս��
        rs������ϸ.MoveNext
    Wend
    WriteInfo " "
    gstrInput��Ҧ = "$$" & mcur��ҽ�� & "~1~" & mstr��ˮ�� & "~" & gstrIC���� & "~0000$$"
    gstrOutput��Ҧ = Space(4000)
    WriteInfo "Ԥ�������:f_UserBargaingApply(29, " & mstr��ˮ�� & ", """ & Replace(gstrInput��Ҧ, String(1053, "0"), "") & """, "" "")"
    glngReturn = f_UserBargaingApply(29, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "Ԥ���㷵��:" & gstrOutput��Ҧ
    �����������_��Ҧ = CheckReturn_��Ҧ()
    
    WriteInfo "���Ԥ����"
    �����������_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    �����������_��Ҧ = False
End Function

Public Function �������_��Ҧ(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rsTemp As New ADODB.Recordset, datCurr As Date, cur���� As Currency, rs��¼ As New ADODB.Recordset
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, _
        cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, strTemp As String, _
        str��Ŀ���� As String, cur�Ը����� As Currency, str��ϸ���� As String
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From ������ü�¼ Where �����־=1 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp!����ID
    While Not rsTemp.EOF
        cur���� = cur���� + rsTemp!ʵ�ս��
        'cd 050301 ͬʱ���½���ͳ����
        
        str��Ŀ���� = IIf(rsTemp!�շ���� = "5" Or rsTemp!�շ���� = "6", "ҩƷ", IIf(rsTemp!�շ���� = "7", "��ҩ", "����"))
        ''''cd 2005 0301
        '������ҩƷ�У�������Ϊ���ƣ����Բ���ֱ�Ӱ�rstemp!������ж�
        gstrSQL = "select * from ҩƷĿ¼ where ҩƷID=[1]"
        Set rs��¼ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!�շ�ϸĿID))
        
        If Not rs��¼.EOF Then '�����ҩƷĿ¼���д���Ŀ���ٸ���ҩƷĿ¼�ĸ�ע���ж��Ƿ�Ϊ����
           If Nvl(rs��¼!��ʶ��, "ҩƷ") = "����" Then str��Ŀ���� = "����"
        End If

        
        '�ӱ���֧����Ŀ�в����Ƿ��и�ҽ����Ŀ
        gstrSQL = "Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ Where �Ƿ�ҽ��=1 And ����=[1] And �շ�ϸĿID=[2]"
        Set rs��¼ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, CLng(rsTemp!�շ�ϸĿID))
        If rs��¼.EOF Then      'û����Ŀ����
           
            If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                cur�Ը����� = "1"
            Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                cur�Ը����� = "0"
            End If
        Else            '�и���Ŀʱ����
            str��ϸ���� = rs��¼!��Ŀ����
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str��ϸ���� & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str��ϸ���� & "'"
            
            Set rs��¼ = gcn��Ҧ.Execute(gstrSQL)
            
            If rs��¼.EOF Then      '���ҽ������Ŀ¼��δ�ҵ�����Ŀ
                If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                    cur�Ը����� = "1"
                Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                    cur�Ը����� = "0"
                End If
            Else        '���ҽ������Ŀ¼���и�ҩƷ
                'strҽ������ = IIf(rs��¼!zfbl = 0, "����", IIf(rs��¼!zfbl = 1, "����", "����"))
                cur�Ը����� = rs��¼!zfbl
            End If
        End If
                '�ڷ��ü�¼�м�¼����ͳ����
                '��Ŀ�����б�����Ŀ���ͣ�ҩƷ�����ƣ�,ժҪ�б����Ը�����,�ɸ��ݱ����õ����࣬����
        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rsTemp!ID & ","
        gstrSQL = gstrSQL & rsTemp!ʵ�ս�� - rsTemp!ʵ�ս�� * cur�Ը����� & ",NULL,1,'" & str��Ŀ���� & "',NULL,'" & cur�Ը����� & "')"
        WriteInfo "�޸���ϸ��Ŀ�Ľ���ͳ����:" & gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '^^^^^^^^^^
        rsTemp.MoveNext
    Wend
    
    ''cd 050225 Ԥ����ʱδ��д���ң��ڴ˲���VVVVV
    gstrSQL = "select ����,���� from ���ű� where ID=(Select ��������ID From ������ü�¼ Where �����־=1 and ����ID=[1] And Rownum < 2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If Not rsTemp.EOF Then
        gstrSQL = "Update HI_ClinicRx Set Department='" & rsTemp!���� & "',DepartmentID='" & rsTemp!���� & "' Where BillID=" & mstr��ˮ��
        WriteInfo "���¿������ƣ�" & gstrSQL
        gcn��Ҧ.Execute gstrSQL
    End If
    
    
    ''^^^^^^^^^^^^^^^^^^
'    gstrOutput��Ҧ = Space(4000)
'
'    gstrInput��Ҧ = "$$1~" & cur���� & "~1~" & mstr��ˮ�� & "~" & gstrIC���� & "$$"
'    WriteInfo vbCrLf & "�������:f_UserBargaingApply(30, " & mstr��ˮ�� & ", """ & Replace(gstrInput��Ҧ, String(1053, "0"), "") & """, "" "")"
'    glngReturn = f_UserBargaingApply(30, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
'    WriteInfo "���㷵��:" & gstrOutput��Ҧ
'    �������_��Ҧ = CheckReturn_��Ҧ()
'    If �������_��Ҧ = False Then
'        Exit Function
'    End If
'    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
'    cur���� = CCur(Split(strTemp, "~")(0))
'
'    gcn��Ҧ.Execute "Delete hi_ClinicRx Where BillID='" & mstr��ˮ�� & "'"
'    gcn��Ҧ.Execute "Delete hi_ClinicPrescription Where BillID='" & mstr��ˮ�� & "'"
    
    Call Get�ʻ���Ϣ(TYPE_��Ҧ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ҧ & "," & _
            lng����ID & "," & Year(datCurr) & ",0,0,0,0," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur���� & "," & curȫ�Ը� & ",0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & mstr��ˮ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------
    �������_��Ҧ = True
    WriteInfo "�������"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_��Ҧ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, str������ As String, rs��¼ As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, lng����ID As Long, strTemp As String
    Dim datCurr As Date, strSQL As String
    Dim cur�˷ѷ��� As Currency
    Dim cur�˷��Է� As Currency
    cur�˷ѷ��� = 0
    cur�˷��Է� = 0
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrIC���� = makeICInfo(lng����ID)
    If gstrIC���� = "" Then Exit Function
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("����ID")
    
    'ȡԭ���ݽ�����ˮ��
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵĽ�����ˮ�Ŷ�ʧ���������ϡ�"
        Exit Function
    End If
    strTemp = rsTemp!�������ý��
    str������ = rsTemp!��ע
'    strSql = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
'        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr��ˮ�� & "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "',1," & Trim(gstrҽԺ����) & ",'" & strҽ���� & "','" & lng����id & "',Null,Null,'" & rs������ϸ!������ & _
'        "','" & str���� & "','" & str����ID & "',Null,'" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "','" & UserInfo.���� & "')"
'    strSql = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
'        "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & iLoop + lng��ˮ & "," & Trim(gstrҽԺ����) & ",'" & mstrBillNo & _
'        "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','" & mstr��ˮ�� & "'," & IIf(str��Ŀ���� = "����", 2, 1) & ",'" & _
'        str��ϸ���� & "','" & str��Ŀ���� & "','" & str��� & "'," & Format(rs������ϸ!ʵ�ս�� / rs������ϸ!����, "#.###") & "," & _
'        rs������ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
'    strSql = "Select * From hi_ClinicRx Where BillID='" & str������ & "'"
    strSQL = "Select * From hi_upClinicRx Where BillID='" & str������ & "'"
    Set rsTemp = gcn��Ҧ.Execute(strSQL)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "ǰ�û���δ�ҵ��õ������ݣ����ϴ������ݲ����˷�"
        ����������_��Ҧ = False
        Exit Function
    End If
    
'    strSql = "Select * From hi_ClinicPrescription Where RecipeSerial='" & str������ & "'"
'    Set rsTemp = gcn��Ҧ.Execute(strSql)
'    If rsTemp.EOF Then
'        Err.Raise 9000,gstrSysName, "ǰ�û���δ�ҵ��õ������ݣ����ϴ������ݲ����˷�"
'        ����������_��Ҧ = False
'        Exit Function
'    End If
'    gcn��Ҧ.Execute "Delete hi_ClinicRx Where BillID='" & str������ & "'"
'    gcn��Ҧ.Execute "Delete hi_ClinicPrescription Where RecipeSerial='" & str������ & "'"
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(31)

    '���ýӿ�������
    gstrInput��Ҧ = "$$" & str������ & "~" & gstrIC���� & "$$"
    gstrOutput��Ҧ = Space(4000)
    WriteInfo "�˷ѵ��ã�f_UserBargaingApply(31, CDbl(" & mstr��ˮ�� & "), " & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    glngReturn = f_UserBargaingApply(31, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "�˷ѷ��أ�" & gstrOutput��Ҧ
    ����������_��Ҧ = CheckReturn_��Ҧ()
    If ����������_��Ҧ = False Then
        Exit Function
    End If
    
    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
    cur�˷ѷ��� = CCur(Split(strTemp, "~")(0))
    cur�˷��Է� = CCur(Split(strTemp, "~")(1))
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ҧ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ҧ & "," & lng����ID & "," & _
        Year(datCurr) & ",0,0,0,0," & intסԺ�����ۼ� & ",0,0,0," & cur�˷ѷ��� & "," & cur�˷��Է� & ",0,0," & _
        "0,0,0,0,NULL,NULL,NULL,'" & mstr��ˮ�� & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ����������_��Ҧ = True
    WriteInfo "�˷����"

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "������Ϣ��" & Err.Description & vbCrLf & "�ӿڷ��أ�" & gstrOutput��Ҧ
End Function

Public Function ��Ժ�Ǽ�_��Ҧ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date, strTemp As String
    Dim lng����ID As Long
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
'    If rsTmp.BOF Then ��Ժ�Ǽ�_��Ҧ = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ҧ
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_��Ҧ = False
        Exit Function
    End If
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.����,D.���� As ���ұ��� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(32)
'    gstrIC���� = makeICInfo(lng����ID)
'                                             '????? סԺ�ţ�����ˮ�Ŵ����в���
'    gstrInput��Ҧ = "$$" & gstrIC���� & "~" & mstr��ˮ�� & "~" & _
'        Format(Nvl(rsTemp(0), datCurr), "YYYY.M.D") & "~" & Nvl(rsTemp(4), "ҽ��") & "~" & strInNote & "~" & _
'        str���ֱ��� & "~" & Nvl(rsTemp!סԺ����, " ") & "~" & Nvl(rsTemp!���ұ���, "0") & "~" & Nvl(rsTemp!��Ժ����, "1") & "$$"
'    gstrOutput��Ҧ = Space(4000)
'
'    WriteInfo "��Ժ�Ǽǣ�f_UserBargaingApply(32, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
'
'    glngReturn = f_UserBargaingApply(32, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
'    WriteInfo "���أ�" & gstrOutput��Ҧ
'    ��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
'    If ��Ժ�Ǽ�_��Ҧ = False Then
'        Exit Function
'    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'˳���'," & mstr��ˮ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ҧ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��Ҧ = False
End Function

'VVVVVV
Public Function ��Ժ�Ǽǳ���_��Ҧ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim strסԺ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'ȡ���˵�סԺ��
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = rsTemp!˳��� ' Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val(rsTemp!˳���)

    '��ڲ��� (Data)
    '�����壺 סԺ�Ǽǽ��׺�+~+Ҫע����סԺ��+~+IC������Ϣ

    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(40)
    gstrIC���� = makeICInfo(lng����ID)
    
    '���ýӿ�
    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & strסԺ�� & "~" & gstrIC���� & "$$"
    gstrOutput��Ҧ = Space(4000)
    
    WriteInfo "��Ժ�Ǽǣ�f_UserBargaingApply(40, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    glngReturn = f_UserBargaingApply(40, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "���أ�" & gstrOutput��Ҧ
    
    ��Ժ�Ǽǳ���_��Ҧ = CheckReturn_��Ҧ()
    If ��Ժ�Ǽǳ���_��Ҧ = False Then
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ҧ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")

    ��Ժ�Ǽǳ���_��Ҧ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'^^^^^^

Public Function ���ʴ���_��Ҧ(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long, iLoop As Long, strSQL As String, lng��ˮ As String, _
        rs��ϸ As New ADODB.Recordset, strTemp As String, strסԺ�� As String, str��ϸ���� As String, str��Ŀ���� As String, _
        str��� As String, str��Ŀ���� As String, strҽ������ As String, str�Ը����� As String
    On Error GoTo errHandle
    'ȡ���������ҳID
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = rsTemp!˳��� ' Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val(rsTemp!˳���)
    
    'ȡ���˷��ü�¼
    If str���ݺ� <> "" Then
        gstrSQL = "Select * From סԺ���ü�¼ Where ʵ�ս��<>0 And ʵ�ս�� Is Not Null And ��¼״̬<>0" & _
            "And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ��¼����=[1] and NO=[2] order by ��ҳID,���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", int����, str���ݺ�)
    Else
        gstrSQL = "Select * From סԺ���ü�¼ Where ʵ�ս��<>0 And ʵ�ս�� Is Not Null And ��¼״̬<>0 " & _
            " And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����id=[1] And ��ҳid=[2] order by ��ҳID,���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID, lng��ҳID)
    End If
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(33)
    iLoop = 1
    strSQL = "select max(SerialNum) as SerialNum From "
    strSQL = strSQL & "(Select SerialNum From hi_upInpatientPrescription  "
    strSQL = strSQL & " union all"
    strSQL = strSQL & " select SerialNum From hi_InpatientPrescription ) A"
    
    Set rsTemp = gcn��Ҧ.Execute(strSQL)
    If rsTemp.EOF Then
        lng��ˮ = 0
    Else
        lng��ˮ = Nvl(rsTemp(0), 0)
    End If
    While Not rs��ϸ.EOF
        gstrSQL = "Select ����,����,���,nvl(���,'') as ��� From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        str��ϸ���� = rsTemp!����: str��Ŀ���� = rsTemp!����
        str��� = Left(Left(rsTemp!��� & " |", InStr(rsTemp!��� & " |", "|") - 1), InStr(rsTemp!��� & " |", " ") - 1)
        '�ж���Ŀ����
        str��Ŀ���� = IIf(rsTemp!��� = "5" Or rsTemp!��� = "6", "ҩƷ", IIf(rsTemp!��� = "7", "��ҩ", "����"))
        ''''cd 2005 0301
        '������ҩƷ�У�������Ϊ���ƣ����Բ���ֱ�Ӱ�rstemp!������ж�
        gstrSQL = "select * from ҩƷĿ¼ where ҩƷID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        
        If Not rsTemp.EOF Then '�����ҩƷĿ¼���д���Ŀ���ٸ���ҩƷĿ¼�ĸ�ע���ж��Ƿ�Ϊ����
           If Nvl(rsTemp!��ʶ��, "ҩƷ") = "����" Then str��Ŀ���� = "����"
        End If
        
        '^^^^^^^^^^^^^^^
        '�ӱ���֧����Ŀ�в����Ƿ��и�ҽ����Ŀ
        gstrSQL = "Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ Where �Ƿ�ҽ��=1 And ����=[1] And �շ�ϸĿID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then      'û����Ŀ����
            If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "1"
            Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                strҽ������ = "����": str�Ը����� = "0"
            End If
        Else            '�и���Ŀʱ����
            str��ϸ���� = rsTemp!��Ŀ����
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str��ϸ���� & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str��ϸ���� & "'"
            
            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
            
            If rsTemp.EOF Then      '���ҽ������Ŀ¼��δ�ҵ�����Ŀ
                If str��Ŀ���� = "ҩƷ" Then    '����ΪҩƷʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "1"
                Else        '��Ŀ����Ϊ���ƻ���ҩʱ��ҽ������Ϊ�����ࡱ
                    strҽ������ = "����": str�Ը����� = "0"
                End If
            Else        '���ҽ������Ŀ¼���и�ҩƷ
                strҽ������ = IIf(rsTemp!zfbl = 0, "����", IIf(rsTemp!zfbl = 1, "����", "����"))
                str�Ը����� = rsTemp!zfbl
            End If
        End If
''''''''''        strSQL = "Insert Into hi_InpatientPrescription (SerialNum,InpatientID,HospitalID,FeeType,RecipeSerial,DateDiagnose," & _
''''''''''            "Class,ItemID,ItemName,Specification,Price,Dosage,ScaleSelf,Operator) Values (" & lng��ˮ & ",'" & _
''''''''''            strסԺ�� & "'," & Trim(gstrҽԺ����) & ",1,Null,'" & Format(rs��ϸ!����ʱ��, "yyyy-MM-dd HH:mm:ss") & _
''''''''''            "'," & IIf(str��Ŀ���� = "����", 1, 2) & ",'" & str��ϸ���� & _
''''''''''            "','" & str��Ŀ���� & "','" & str��� & "'," & Format(Nvl(rs��ϸ!ʵ�ս��, 0) / (rs��ϸ!���� * rs��ϸ!����), _
''''''''''            "0.000") & "," & rs��ϸ!���� * rs��ϸ!���� & "," & str�Ը����� & ",'" & UserInfo.���� & "')"
''''''''''
''''''''''        WriteInfo "������ϸ(д������ϸ):" & strSQL
''''''''''        gcn��Ҧ.Execute strSQL
''''''''''        lng��ˮ = AddNum(lng��ˮ)
''''''''''
''''''''''        'gstrSQL = "ZL_����֧����Ŀ_Modify(" & rs��ϸ!�շ�ϸĿID & "," & TYPE_��Ҧ & ",NULL,'" & str��ϸ���� & "','" & _
''''''''''        '    str��Ŀ���� & "','" & strҽ������ & "',1)"
'''''''''''        WriteInfo "�޸ı���֧����Ŀ:" & gstrSQL
''''''''''        'Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
''''''''''
''''''''''        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
''''''''''        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
                '�ڷ��ü�¼�м�¼����ͳ����
                '��Ŀ�����б�����Ŀ���ͣ�ҩƷ�����ƣ�,ժҪ�б����Ը�����,�ɸ��ݱ����õ����࣬����
        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs��ϸ!ID & ","
        gstrSQL = gstrSQL & rs��ϸ!ʵ�ս�� - rs��ϸ!ʵ�ս�� * str�Ը����� & ",NULL,1,'" & str��Ŀ���� & "',NULL,'" & str�Ը����� & "')"
        WriteInfo "�޸���ϸ��Ŀ�Ľ���ͳ����:" & gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '^^^^^^^^^^
        rs��ϸ.MoveNext
    Wend
    '���ýӿ�
''''''''''    gstrIC���� = makeICInfo(lng����ID)
''''''''''    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & gstrIC���� & "~0000$$"
''''''''''    gstrOutput��Ҧ = Space(4000)
''''''''''
''''''''''    WriteInfo "���ʣ�f_UserBargaingApply(33, " & mstr��ˮ�� & ", " & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
''''''''''
''''''''''    glngReturn = f_UserBargaingApply(33, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
''''''''''    WriteInfo "���أ�" & gstrOutput��Ҧ
''''''''''
''''''''''    ���ʴ���_��Ҧ = CheckReturn_��Ҧ()
''''''''''    If ���ʴ���_��Ҧ = False Then
''''''''''        Exit Function
''''''''''    End If

    ���ʴ���_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ���ʴ���_��Ҧ = False
End Function

Public Function סԺ�������_��Ҧ(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim datCurr As Date, strסԺ�� As String
    
    On Error GoTo errHandle
    Set rs��ϸ = rs������ϸ.Clone
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    '��Ҫ���ϴ�������ϸ
    If ���ʴ���_��Ҧ("", 0, "", lng����ID) = False Then Exit Function
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = rsTemp!˳��� 'Format(Val(rsTemp!˳���), "0" & String(16, "#"))
    
    '�����ҽ�����
    mcur��ҽ�� = 0
    While Not rs��ϸ.EOF
'        gstrSQL = "Select B.���,A.��Ŀ����,A.��Ŀ����,Nvl(B.���,'') As ��� From ����֧����Ŀ A,�շ�ϸĿ B " & _
'            "Where A.�շ�ϸĿID=B.ID And A.�Ƿ�ҽ��=1 And A.����=" & TYPE_��Ҧ & " And A.�շ�ϸĿID=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If Not rsTemp.EOF Then
'            '�ж�ҽ��ǰ�û����Ƿ��и���Ŀ
'               gstrSQL = "Select MedicineID,zfbl From hi_Medicine  Where MedicineID='" & rsTemp(1) & "'"
'               gstrSQL = gstrSQL & " union all Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & rsTemp(1) & "'"
'            Set rsTemp = gcn��Ҧ.Execute(gstrSQL)
'            If rsTemp.EOF Then mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!���
'        Else
'            mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!���
'        End If
        '
        If Nvl(rs��ϸ!ժҪ, 0) = 1 Then mcur��ҽ�� = mcur��ҽ�� + rs��ϸ!���
        rs��ϸ.MoveNext
    Wend
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(34)
    '���ýӿ�
    '�ýӿ���ȫ�Էѣ�������������壬Ϊ���ٳ�����ܣ��ݲ����ýӿ�
    gstrIC���� = makeICInfo(lng����ID)
    gstrInput��Ҧ = "$$" & strסԺ�� & "~" & mcur��ҽ�� & "~" & gstrIC���� & "~0000$$"
    gstrOutput��Ҧ = Space(4000)
    WriteInfo "������㣺f_UserBargaingApply(34, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    glngReturn = f_UserBargaingApply(34, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "���أ�" & gstrOutput��Ҧ
    'if CheckReturn_��Ҧ then
    
    סԺ�������_��Ҧ = "ҽ������;0;0"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ�������_��Ҧ = ""
End Function

Public Function ��Ժ�Ǽ�_��Ҧ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, bln����ó�Ժ As Boolean, strסԺ�� As String, _
        strInNote As String, str���ֱ��� As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(sum(ʵ�ս��),0) as ��� from סԺ���ü�¼ where nvl(���ӱ�־,0)<>9 and ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˳�Ժ", lng����ID, lng��ҳID)
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (rsTemp("���") = 0)
    End If
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = rsTemp!˳��� ' Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Val()
    
    If bln����ó�Ժ = True Then
        '������Ժ�Ǽǳ���
        mstr��ˮ�� = ���뽻����ˮ_��Ҧ(40)
        gstrIC���� = makeICInfo(lng����ID)
        
        '���ýӿ�
        gstrInput��Ҧ = "$$" & strסԺ�� & "~" & strסԺ�� & "~" & gstrIC���� & "$$"
        gstrOutput��Ҧ = Space(4000)
        WriteInfo "��Ժ������f_UserBargaingApply(40, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
        glngReturn = f_UserBargaingApply(40, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
        WriteInfo "���أ�" & gstrOutput��Ҧ

        ��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
        Exit Function
    End If
    
    '��ȡ��Ժ���
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True)
    
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ҧ
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ȷ�Ｒ��")
    If rsTemp.State = 1 Then
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_��Ҧ = False
        Exit Function
    End If
    '��ȡסԺҽʦ
    gstrSQL = "select סԺҽʦ from ������ҳ Where ��ҳID = [2] And ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "����ȡ�ò��˵���Ժ�Ǽ���Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(35)
    gstrIC���� = makeICInfo(lng����ID)
    
    '���ýӿ�
    'gstrInput��Ҧ = "$$" & strסԺ�� & "~" & Nvl(rsTemp(0), " ") & "~" & strInNote & "~" & _
        str���ֱ��� & "~" & Format(datCurr, "yyyy-MM-dd") & "$$"
    'gstrOutput��Ҧ = Space(4000)
    'WriteInfo "��Ժ�Ǽǣ�f_UserBargaingApply(35, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    'glngReturn = f_UserBargaingApply(35, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    'WriteInfo "���أ�" & gstrOutput��Ҧ
    '��Ժ�Ǽ�_��Ҧ = CheckReturn_��Ҧ()
    'If ��Ժ�Ǽ�_��Ҧ = False Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ҧ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽ�_��Ҧ = True
  
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_��Ҧ = False
End Function

Public Function סԺ����_��Ҧ(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String, strTemp As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�������ý�� As Currency, curȫ�Ը���� As Currency
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From סԺ���ü�¼ Where ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rs��ϸ.EOF Then
        Err.Raise 9000, gstrSysName, "û�з�����ϸ�����ܽ��г�Ժ����"
        Exit Function
    End If
    lng����ID = rs��ϸ!����ID
    
    gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ҧ)
    str������ = rsTemp!˳��� ' Format(Val(rsTemp!˳���), "0" & String(16, "#")) ' Nvl(rsTemp!˳���)
    datCurr = zlDatabase.Currentdate
   
    '�����ҽ����Ŀ���
    mcur��ҽ�� = 0
    cur�������ý�� = 0
    curȫ�Ը���� = 0
    While Not rs��ϸ.EOF
        If Nvl(rs��ϸ!ժҪ, 0) = 1 Then mcur��ҽ�� = mcur��ҽ�� + Nvl(rs��ϸ!ʵ�ս��, 0)
        cur�������ý�� = cur�������ý�� + Nvl(rs��ϸ!ʵ�ս��, 0)
        rs��ϸ.MoveNext
    Wend
    'mstr��ˮ�� = ���뽻����ˮ_��Ҧ(36)
    'gstrIC���� = makeICInfo(lng����ID)
    'gstrInput��Ҧ = "$$1~" & str������ & "~" & mcur��ҽ�� & "~" & gstrIC���� & "$$"
    'gstrOutput��Ҧ = Space(4000)
    
    'WriteInfo "סԺ���㣺f_UserBargaingApply(36, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    'glngReturn = f_UserBargaingApply(36, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    'WriteInfo "���أ�" & gstrOutput��Ҧ
    
    'סԺ����_��Ҧ = CheckReturn_��Ҧ()
    'If סԺ����_��Ҧ = False Then Exit Function
    'strTemp = Split(gstrOutput��Ҧ, "$$")(2)
    '��Ҧ���ؽ�����ԣ�������ʱ����
    '    cur�������ý�� = CCur(Split(strTemp, "~")(0))
    
    'curȫ�Ը���� = CCur(Split(strTemp, "~")(1))
    '��Ҧ����ֻ�����ܶ��ȫ�Ը����ֽ��,��
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ҧ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ҧ & "," & _
            lng����ID & "," & Year(datCurr) & ",0," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,NULL,NULL," & cur�������ý�� & _
            "," & mcur��ҽ�� & ",0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & _
            str������ & "~" & mstr��ˮ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    סԺ����_��Ҧ = True
    '---------------------------------------------------------------------------------------------

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_��Ҧ(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, lng����ID As Long, str��ˮ�� As String, str������ As String, _
        lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, _
        curͳ�ﱨ���ۼ� As Currency, intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, curȫ�Ը���� As Currency
    Dim strTemp As String
    Dim datCurr As Date
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From סԺ���ü�¼ Where ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û�ҵ����˵ķ�����ϸ��¼�������˷�"
        Exit Function
    End If
    lng����ID = rsTemp("����ID")
    Do Until rsTemp.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�

    gstrSQL = "select distinct A.ID as ����ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              "  where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵľ����Ŷ�ʧ���������ϡ�"
        Exit Function
    End If
    str������ = Split(rsTemp!��ע, "~")(0)
    str��ˮ�� = Split(rsTemp!��ע, "~")(1)
    
    '���ýӿ�������
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(37)
    gstrIC���� = makeICInfo(lng����ID)

    '���ýӿ�
    gstrInput��Ҧ = "$$" & str��ˮ�� & "~" & gstrIC���� & "$$"
    gstrOutput��Ҧ = Space(4000)
    WriteInfo "סԺ���㣺f_UserBargaingApply(37, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    glngReturn = f_UserBargaingApply(37, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "���أ�" & gstrOutput��Ҧ
    
    סԺ�������_��Ҧ = CheckReturn_��Ҧ()
    If סԺ�������_��Ҧ = False Then Exit Function
    strTemp = Split(gstrOutput��Ҧ, "$$")(2)
    'curƱ���ܽ�� = CCur(Split(strTemp, "~")(0))
    curȫ�Ը���� = CCur(Split(strTemp, "~")(1))

'^^^^^^^^^^^^^^
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'˳���'," & str��ˮ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ҧ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ҧ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & "," & curȫ�Ը���� & ",0,0,0,0,0,0," & _
        "NULL,NULL,NULL,'" & str������ & "~" & str��ˮ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ�������_��Ҧ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��ݱ�ʶ_��Ҧ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify��Ҧ
    Dim strPatiInfo As String, cur��� As Currency, str������ As String
    Dim arr, datCurr As Date, str����� As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    'MODIFIED BY ZYB ����ҽ���ӿڿ���
    strPatiInfo = frmIDentified.GetPatient(bytType, lng����ID)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_��Ҧ)
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        ��ݱ�ʶ_��Ҧ = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    ��ݱ�ʶ_��Ҧ = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��Ҧ = ""
End Function

Public Function �������_��Ҧ(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    '��Ҧ������ȡ�����ʻ����
    �������_��Ҧ = 0
End Function

Public Function ת��ת��_��Ҧ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date, strTemp As String
    Dim lng����ID As Long
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    If rsTmp.BOF Then ת��ת��_��Ҧ = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ҧ
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ת��ת��_��Ҧ = False
        Exit Function
    End If
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.����,D.���� As ���ұ���,C.˳��� As סԺ��ˮ from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [1] And A.����ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng��ҳID, lng����ID)
    
    mstr��ˮ�� = ���뽻����ˮ_��Ҧ(38)
    gstrIC���� = makeICInfo(lng����ID)
    
    gstrInput��Ҧ = "$$" & rsTemp!סԺ��ˮ & "~" & Format(datCurr, "yyyy-MM-dd") & "~" & _
        rsTemp(3) & "~" & Nvl(rsTemp(4), " ") & "~" & strInNote & "~" & _
        str���ֱ��� & "~" & Nvl(rsTemp!סԺ����, " ") & "~" & Nvl(rsTemp!���ұ���, "0") & "$$"
    gstrOutput��Ҧ = Space(4000)
    
    WriteInfo "סԺ���㣺f_UserBargaingApply(38, " & mstr��ˮ�� & "," & Replace(gstrInput��Ҧ, String(1053, "0"), "") & ", gstrOutput��Ҧ)"
    glngReturn = f_UserBargaingApply(38, mstr��ˮ��, gstrInput��Ҧ, gstrOutput��Ҧ)
    WriteInfo "���أ�" & gstrOutput��Ҧ
    
    ת��ת��_��Ҧ = CheckReturn_��Ҧ()
    If ת��ת��_��Ҧ = False Then
        Exit Function
    End If
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ҧ & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ת��ת��_��Ҧ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ת��ת��_��Ҧ = False
End Function

Private Function AddNum(ByVal strSequence As String) As String
    Dim intAdd As Integer
    Dim intDO As Integer, intLen As Integer
    '��ɼӷ�����
    intLen = Len(strSequence)
    intAdd = Val(Right(strSequence, 1))
    intAdd = intAdd + 1
'    If intAdd > 9 Then
        '��λ
        For intDO = intLen To 1 Step -1
            intAdd = Val(Mid(strSequence, intDO, 1))
            intAdd = intAdd + 1
            If intAdd <= 9 Then
                '��ǰλ����ǰ��λ�����������λȫ��Ϊ��
                AddNum = Mid(strSequence, 1, intDO - 1) & CStr(intAdd) & String(intLen - intDO, "0")
                Exit For
            End If
        Next
        '������λ�����λ���������λ1������λΪ��
        If intDO = 0 Then AddNum = "1" & String(intLen, "0")
'    Else
'        AddNum = Mid(strSequence, 1, intLen - 1) & CStr(intAdd)
'    End If
End Function
