Attribute VB_Name = "mdl�Թ�"
Option Explicit
Private Declare Function zg_GetErrToStr Lib "DataExchange" Alias "GetErrToStr" (ByVal lngErr As Long) As String
Private Declare Function zg_ReadICCardInfo Lib "DataExchange" Alias "ReadICCardInfo" (ByVal strPass As String) As String
Private Declare Function zg_ChangePassword Lib "DataExchange" Alias "DllChangePassWord" (ByVal strPass As String, ByVal strNewPass As String) As String
Private Declare Function zg_ClinicCharge Lib "DataExchange" Alias "ClinicCharge" _
    (ByVal strPass As String, ByVal strClinicBalanceNO As String, ByVal blnIsPrepare As Long) As String
Private Declare Function zg_InHosReg Lib "DataExchange" Alias "InHosReg" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String) As Long
Private Declare Function zg_UnInHosReg Lib "DataExchange" Alias "UnInHosReg" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String) As Long
Private Declare Function zg_PreInHosBalance Lib "DataExchange" Alias "PreInHosBalance" _
    (ByVal strPass As String, ByVal strInHosRegisterNO As String, ByVal lngBalanceType As Long, _
    Optional ByVal strCheckupKind As String = "3", _
    Optional ByVal strSickKindCode As String = "0", _
    Optional ByVal intAccount As Long = 0) As String
Private Declare Function zg_InHosBalance Lib "DataExchange" Alias "InHosBalance" _
    (ByVal strPass As String, ByVal strInHosBalanceNO As String, ByVal lngBalanceType As Long, _
    Optional ByVal strCheckupKind As String = "3", _
    Optional ByVal strSickKindCode As String = "0", _
    Optional ByVal intAccount As Long = 0) As String
Private Declare Function zg_UnInHosBalance Lib "DataExchange" Alias "UnInHosBalance" _
    (ByVal strPass As String, ByVal strInHosBalanceNO As String, ByVal strUnInHosBalanceNo As String) As String

'Public gobj�Թ� As New clsZGYB              '������
Public mblnInit As Boolean
Public gcn�Թ� As New ADODB.Connection

Public Enum ҵ������_�Թ�
    ����
    �޸�����
    ����Ԥ��
    �������
    ��Ժ�Ǽ�
    ������Ժ�Ǽ�
    סԺԤ����
    סԺ����
    סԺ�������
End Enum

Private Type ������Ϣ_�Թ�
    ����ID As Long
    �ܽ�� As Currency
    ȫ�Է� As Currency
    �����Ը� As Currency
    ����ͳ�� As Currency
    �������� As Currency
    ʵ������ As Currency
    ͳ��֧�� As Currency
    ���ⶥ���� As Currency
    �����ʻ� As Currency
    �ֽ�֧�� As Currency
End Type
Private cur_������Ϣ As ������Ϣ_�Թ�
Private gstr����� As String    '���浥�ݺ�

Public Function ҽ����ʼ��_�Թ�() As Boolean
    If mblnInit Then
        ҽ����ʼ��_�Թ� = True
        Exit Function
    End If
    ҽ����ʼ��_�Թ� = ���ҽ��������_�Թ�
End Function

Private Function ���ҽ��������_�Թ�() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn����.State = adStateOpen Then
        ���ҽ��������_�Թ� = True
        Exit Function
    End If
    
    '��������ҽ��������������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Ĵ��Թ�)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(gcn�Թ�, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷�����ҽ��ǰ�û���", vbInformation, gstrSysName
        Exit Function
    End If
    
    ���ҽ��������_�Թ� = True
End Function

Public Function ���ýӿ�_�Թ�(ByVal intType As Integer, ByVal StrInput As String, strOutput As String) As Boolean
    Dim lngErr As Long
    Dim arrPara             '�����ֽ���Σ���ĳ��������Ҫ������ʱ
    
    Select Case intType
    Case ҵ������_�Թ�.����
        Call WriteBusinessLOG("ReadICCardInfo", "����ǰ���", "")
'        strOutPut = gobj�Թ�.ReadICCardInfo(strInput)
        strOutput = zg_ReadICCardInfo(StrInput)
        Call WriteBusinessLOG("ReadICCardInfo", "", strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case ҵ������_�Թ�.�޸�����
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("DllChangePassWord", "����ǰ���", "")
'        strOutPut = gobj�Թ�.ChangePass(arrPara(0), arrPara(1))
        strOutput = zg_ChangePassword(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("DllChangePassWord", arrPara(0) & "," & arrPara(1), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case ҵ������_�Թ�.����Ԥ��, ҵ������_�Թ�.�������
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("ClinicCharge", "����ǰ���", "")
'        strOutPut = gobj�Թ�.ClinicCharge(arrPara(0), arrPara(1), (Val(arrPara(2)) = 1))
        strOutput = zg_ClinicCharge(arrPara(0), arrPara(1), Val(arrPara(2)))
        Call WriteBusinessLOG("ClinicCharge", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case ҵ������_�Թ�.��Ժ�Ǽ�
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("InHosReg", "����ǰ���", "")
'        lngErr = gobj�Թ�.InHosReg(arrPara(0), arrPara(1))
        lngErr = zg_InHosReg(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("InHosReg", arrPara(0) & "," & arrPara(1), lngErr)
    Case ҵ������_�Թ�.������Ժ�Ǽ�
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("UnInHosReg", "����ǰ���", "")
'        lngErr = gobj�Թ�.UnInHosReg(arrPara(0), arrPara(1))
        lngErr = zg_UnInHosReg(arrPara(0), arrPara(1))
        Call WriteBusinessLOG("UnInHosReg", arrPara(0) & "," & arrPara(1), lngErr)
    Case ҵ������_�Թ�.סԺԤ����
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("PreInHosBalance", "����ǰ���", "")
'        strOutPut = gobj�Թ�.PreInHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        strOutput = zg_PreInHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        Call WriteBusinessLOG("PreInHosBalance", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)) & "," & arrPara(3) & "," & arrPara(4) & "," & CLng(arrPara(5)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case ҵ������_�Թ�.סԺ����
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("InHosBalance", "����ǰ���", "")
'        strOutPut = gobj�Թ�.InHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        strOutput = zg_InHosBalance(arrPara(0), arrPara(1), Val(arrPara(2)), arrPara(3), arrPara(4), CLng(arrPara(5)))
        Call WriteBusinessLOG("InHosBalance", arrPara(0) & "," & arrPara(1) & "," & Val(arrPara(2)) & "," & arrPara(3) & "," & arrPara(4) & "," & CLng(arrPara(5)), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    Case ҵ������_�Թ�.סԺ�������
        arrPara = Split(StrInput, "|")
        Call WriteBusinessLOG("UnInHosBalance", "����ǰ���", "")
'        strOutPut = gobj�Թ�.UnInHosBalance(arrPara(0), arrPara(1), arrPara(2))
        strOutput = zg_UnInHosBalance(arrPara(0), arrPara(1), arrPara(2))
        Call WriteBusinessLOG("UnInHosBalance", arrPara(0) & "," & arrPara(1) & "," & arrPara(2), strOutput)
        lngErr = Val(Mid(strOutput, 1, 4))
    End Select
    
    '�ж��Ƿ�������
    If lngErr <> 0 Then
'        MsgBox "����ҽ���ӿڷ��ش�����ϸ��Ϣ���£�" & vbCrLf & _
'            gobj�Թ�.GetErrToStr(lngErr), vbInformation, gstrSysName
        MsgBox "����ҽ���ӿڷ��ش�����ϸ��Ϣ���£�" & vbCrLf & _
            zg_GetErrToStr(lngErr), vbInformation, gstrSysName
        strOutput = ""
        Exit Function
    End If
    
    If Not (intType = ҵ������_�Թ�.��Ժ�Ǽ� Or intType = ҵ������_�Թ�.������Ժ�Ǽ�) Then
        If strOutput = "" Then
            MsgBox "�ӿڷ��ص����ݲ���ȷ�����ش�Ϊ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, strOutput, "|*") <> 0 Then
        strOutput = Mid(strOutput, 7)
        strOutput = Mid(strOutput, 1, InStr(1, strOutput, "|*") - 1)
    End If
    ���ýӿ�_�Թ� = True
End Function

Public Function ��ݱ�ʶ_�Թ�(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ� �ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strReturn As String
    On Error GoTo errHandle
    
    strReturn = frmIdentify�Թ�.GetPatient(bytType, lng����ID, (bytType <> 2), True)
    If strReturn = "" Then Exit Function
    
    ��ݱ�ʶ_�Թ� = strReturn
    gstr����� = ""          'ÿ��Ԥ����ʱ�����Ӳ��ű������л�ȡ����ֵ����Ϊ�����������Ĵ�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_�Թ�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strReturn As String, str�������� As String
    Dim arrBalance
    Dim strInsert As String, strValue As String
    Dim str�������� As String, strҽ������ As String, str���ֱ��� As String, str�������� As String, str�������� As String
    Dim strҽ����Ŀ���� As String, strҽԺ��Ŀ���� As String
    Dim rsTemp As New ADODB.Recordset
    
    Const int�ܽ�� As Integer = 0
    Const int�ʻ�֧�� As Integer = 1
    Const int�ֽ�֧�� As Integer = 2
    Const intͳ��֧�� As Integer = 3
    On Error GoTo errHand
    
    cur_������Ϣ.����ID = rs��ϸ!����ID
    str�������� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    gstr����� = zlDatabase.GetNextID("���ű�")
    
    '���²��뱾��Ԥ����Ĵ�����ϸ
    '��������(�����վݺţ����ۺţ������վݺ�-�ⲿ���������ң�ҽ�����������ֱ��룬�������ƣ��������ʹ��룬�����־����������)
    strInsert = " Insert Into ClinicBill " & _
                " (ClinicBillNO,ClinicBalanceNO,InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode,RedBillFlag,OccurDate) " & _
                " Values ("
    
    '��ȡ������������
    gstrSQL = "Select ���� From ���ű� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������", CLng(rs��ϸ!��������ID))
    str�������� = Nvl(rsTemp!����)
    strҽ������ = Nvl(rs��ϸ!������, "HIS")
    
    'todo:�����Ƿ���Ҫѡ����?�˴�������ȡ������Ϣ
'    gstrSQL = ""
'    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
'    str���ֱ��� = "001"
'    str�������� = "��ͨ��"
'    str�������� = "1"
    
    On Error Resume Next
    strValue = "'" & gstr����� & "','" & gstr����� & "','" & gstr����� & "','" & str�������� & "','" & strҽ������ & "'," & _
        Val(str���ֱ���) & ",'" & str�������� & "','" & str�������� & "',1,to_Date('" & str�������� & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSQL = strInsert & strValue & ")"
    gcn�Թ�.Execute gstrSQL
    On Error GoTo errHand
    
    With rs��ϸ
        '������ϸ��(�����շ���ϸ��ˮ�ţ������վݺţ��շ���Ŀ���룬ҽԺ�շ���Ŀ���ƣ����ۣ����������)
        strInsert = " Insert Into ClinicBillDetail" & _
                    " (ClinicBillDetailNO,ClinicBillNO,ItemNO,HosItemName,Price,Quantity,Amount)" & _
                    " Values ("
        Do While Not .EOF
            '��ȡ��Ŀ��ҽ������
            gstrSQL = " Select A.��Ŀ����,B.���� From ����֧����Ŀ A,�շ�ϸĿ B" & _
                      " Where A.�շ�ϸĿID=B.ID And A.����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��ҽ������", TYPE_�Ĵ��Թ�, CLng(!�շ�ϸĿID))
            strҽ����Ŀ���� = Nvl(rsTemp!��Ŀ����)
            strҽԺ��Ŀ���� = Nvl(rsTemp!����)
            
            '������ϸ��
            strValue = "'" & .AbsolutePosition & "','" & gstr����� & "','" & strҽ����Ŀ���� & "','" & strҽԺ��Ŀ���� & "'," & _
                       Val(Format(!ʵ�ս�� / !����, "#####0.0000")) & "," & !���� & "," & Nvl(!ʵ�ս��, 0)
            gstrSQL = strInsert & strValue & ")"
            gcn�Թ�.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    '���������������ӿڣ�����ֵ��ʽ�����������ܽ��|�����ʻ�֧��|�ֽ�֧��|ͳ��֧��|�����ʻ����|����|ҽ���ʺ�|����
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.����Ԥ��, GetPass(cur_������Ϣ.����ID) & "|" & gstr����� & "|" & 1, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = Val(arrBalance(int�ܽ��))
    cur_������Ϣ.�����ʻ� = Val(arrBalance(int�ʻ�֧��))
    cur_������Ϣ.�ֽ�֧�� = Val(arrBalance(int�ֽ�֧��))
    cur_������Ϣ.ͳ��֧�� = Val(arrBalance(intͳ��֧��))
    
    '���ؽ��㴮
    If cur_������Ϣ.�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "|�����ʻ�;" & cur_������Ϣ.�����ʻ� & ";0"
    If cur_������Ϣ.ͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & cur_������Ϣ.ͳ��֧�� & ";0"
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    �����������_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ����Ǽ�ȡ��_�Թ�() As Boolean
    On Error GoTo errHand
    
    ����Ǽ�ȡ��_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_�Թ�(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim strReturn As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    
    Const int�ܽ�� As Integer = 0
    Const int�ʻ�֧�� As Integer = 1
    Const int�ֽ�֧�� As Integer = 2
    Const intͳ��֧�� As Integer = 3
    On Error GoTo errHand
    '--ҽ�����Ĳ����ķ�Ʊ��--
'    '��ȡ���ν��㷢Ʊ�ţ����Ĳ������ķ�Ʊ�ţ�
'    gstrSQL = "Select ʵ��Ʊ�� AS ��Ʊ�� From ���˷��ü�¼ Where ����ID=" & lng����ID & " And Rownum<2"
'    Call OpenRecordset(rsTemp, "��ȡ���ν��㷢Ʊ��")
'    strBalanceNO = Nvl(rsTemp!��Ʊ��)
'    If strBalanceNO = "" Then
'        MsgBox "û�з�Ʊ�Ų��ܽ��н��㣬��ָ����Ʊ�ţ�", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '��Ԥ����ʱ��֪�������ţ��˴����´�����������ϸ��Ĵ��������վݺţ����Ĳ������ķ�Ʊ�ţ�����ȡ�����Դ˴�Ҳ��ʹ���ˣ�
'    gstrSQL = "Update ClinicBill Set ClinicBillNO='" & strBalanceNO & "',ClinicBalanceNO='" & strBalanceNO & "',InvoiceNO='" & strBalanceNO & "' Where ClinicBalanceNO='" & gstr����� & "'"
'    gcn�Թ�.Execute gstrSQL
'    gstrSQL = "Update ClinicBillDetail Set ClinicBillNO='" & strBalanceNO & "' Where ClinicBillNO='" & gstr����� & "'"
'    gcn�Թ�.Execute gstrSQL
    '------------------------
    
    '�����������ӿ�
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.�������, GetPass(cur_������Ϣ.����ID) & "|" & gstr����� & "|" & 0, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = Val(arrBalance(int�ܽ��))
    cur_������Ϣ.�����ʻ� = Val(arrBalance(int�ʻ�֧��))
    cur_������Ϣ.�ֽ�֧�� = Val(arrBalance(int�ֽ�֧��))
    cur_������Ϣ.ͳ��֧�� = Val(arrBalance(intͳ��֧��))
   
    Call Get�ʻ���Ϣ(TYPE_�Ĵ��Թ�, cur_������Ϣ.����ID, Year(zlDatabase.Currentdate()), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & cur_������Ϣ.����ID & "," & TYPE_�Ĵ��Թ� & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    'g��������.�����Ը�����б���������ﲡ�˾������ͣ�������ⲡ�������ͨ����������¼�ı�ע������ǲ��ֵ�����
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ĵ��Թ� & "," & cur_������Ϣ.����ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ",0,0,0," & cur_������Ϣ.�ܽ�� & ",0,0," & _
        cur_������Ϣ.����ͳ�� & "," & cur_������Ϣ.ͳ��֧�� & ",0,0," & cur�����ʻ� & ",'" & gstr����� & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    �������_�Թ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_�Թ�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim lng����ID As Long
    Dim strReturn As String, strBalanceNO As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    
    Const int�ܽ�� As Integer = 0
    Const int�ʻ�֧�� As Integer = 1
    Const int�ֽ�֧�� As Integer = 2
    Const intͳ��֧�� As Integer = 3
    On Error GoTo errHand
    
    '��ȡ���γ���ID
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Թ�ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")
    
    'ȡ���ϴν�����վݺ�
    gstrSQL = "Select ֧��˳��� From ���ս����¼ Where ����=[1] And ����=1 And ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ϴν�����վݺ�", TYPE_�Ĵ��Թ�, lng����ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "û���ҵ�ԭʼ���շѼ�¼���޷�������������"
        Exit Function
    End If
    strBalanceNO = Nvl(rsTemp!֧��˳���)
    If strBalanceNO = "" Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "��Ч�������վݺţ��޷�������������"
        Exit Function
    End If
    
    '��������������Ϳ����ˣ���ƱΪ-1
    On Error Resume Next
    gstrSQL = " Insert Into ClinicBill " & _
                " (ClinicBillNO,ClinicBalanceNO,InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode,RedBillFlag,OccurDate,StrikedBillNO) " & _
                " Select 'HCMZ" & strBalanceNO & "','HCMZ" & strBalanceNO & "',InvoiceNO,DepartmentName,DoctorName,SickSerialNO,SickName,SickKindCode," & _
                " -1,to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),'" & strBalanceNO & "'" & _
                " From ClinicBill Where ClinicBillNO='" & strBalanceNO & "'"
    gcn�Թ�.Execute gstrSQL
    On Error GoTo errHand
    
    '����������㽻���������������
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.�������, GetPass(lng����ID) & "|" & "HCMZ" & strBalanceNO & "|" & 0, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = -1 * Val(arrBalance(int�ܽ��))
    cur_������Ϣ.�����ʻ� = -1 * Val(arrBalance(int�ʻ�֧��))
    cur_������Ϣ.�ֽ�֧�� = -1 * Val(arrBalance(int�ֽ�֧��))
    cur_������Ϣ.ͳ��֧�� = -1 * Val(arrBalance(intͳ��֧��))
   
    Call Get�ʻ���Ϣ(TYPE_�Ĵ��Թ�, cur_������Ϣ.����ID, Year(zlDatabase.Currentdate()), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & cur_������Ϣ.����ID & "," & TYPE_�Ĵ��Թ� & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & _
        cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ĵ��Թ� & "," & cur_������Ϣ.����ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ",0,0,0," & cur_������Ϣ.�ܽ�� & ",0,0," & _
        cur_������Ϣ.����ͳ�� & "," & cur_������Ϣ.ͳ��֧�� & ",0,0," & cur_������Ϣ.�����ʻ� & ",'" & strBalanceNO & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    ����������_�Թ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �������_�Թ�(ByVal lng����ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Nvl(�ʻ����,0) AS ��� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ʻ����", TYPE_�Ĵ��Թ�, lng����ID)
    �������_�Թ� = rsTemp!���
End Function

Public Function ��Ժ�Ǽ�_�Թ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim strסԺ�� As String, strReturn As String
    Dim str���� As String, strҽ�� As String, str��Ժ���� As String
    Dim rsTemp As New ADODB.Recordset
    
    '���±����Ͳ����й�
    Dim lng���� As Long, str���� As String
    Dim rsSelected As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    On Error GoTo errHand
    
    'סԺҪѡ���֣���ȷ��һЩ�����շ���Ŀ
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS ����,A.SickName AS ����,A.SickSpell AS ���� " & _
            " From SickDefine A Where 1=2"
    Call OpenRecordset_OtherBase(rsSelected, "��ȡ��ѡ��Ĳ���", gstrSQL, gcn�Թ�)
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS ����,A.SickName AS ����,A.SickSpell AS ���� " & _
            " From SickDefine A Where 1=1"
    Set rs���� = New ADODB.Recordset
    Call OpenRecordset_OtherBase(rs����, "�����֤", gstrSQL, gcn�Թ�)
    
    If rs����.RecordCount > 0 Then
VirusSelect:
        If frm�ಡ��ѡ��_�Թ�.ShowSelect(rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�", rsSelected, False, gcn�Թ�) = True Then
            lng���� = 0
            str���� = ""
            With rs����
                If .RecordCount <> 0 Then .MoveFirst
                lng���� = rs����("ID")
                Do While Not .EOF
                    str���� = str���� & "|" & rs����!ID
                    .MoveNext
                Loop
                If str���� <> "" Then str���� = Mid(str����, 2)
            End With
        Else
            MsgBox "����Ҫѡ���֣�", vbInformation, gstrSysName
            GoTo VirusSelect
        End If
    End If
    
    strסԺ�� = Left(lng����ID & "_" & lng��ҳID, 16) & "_" & Mid(CStr(Get����(lng����ID)), 1, 3)
    '��ȡ���˵���Ժ����,ҽ��,��Ժ����
    gstrSQL = "Select B.���� As ����,A.����ҽʦ As ҽ��,A.��Ժ���� " & _
             " From ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID=B.Id And A.����ID=[1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵���Ժ���ҡ�ҽ������Ժ����", lng����ID, lng��ҳID)
    str���� = Nvl(rsTemp!����)
    strҽ�� = Nvl(rsTemp!ҽ��)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
    
    '������Ժ��¼(��Ժ�ǼǺ�(ϵͳ),��Ժ�ǼǺ�(�ⲿ),סԺ����,ҽ��,��Ժ����)
    On Error Resume Next
    gstrSQL = " Insert Into InHosRegister(INHOSREGISTERNO,INHOSNO,DEPARTMENTNAME,DOCTORNAME,INHOSDATE) " & _
              " Values ('" & strסԺ�� & "','" & strסԺ�� & "','" & str���� & "','" & strҽ�� & "'," & _
              " to_Date('" & str��Ժ���� & "','yyyy-MM-dd hh24:mi:ss'))"
    gcn�Թ�.Execute gstrSQL
    On Error GoTo errHand
    
    '���벡������
    Call InsertDisease("RegHosSick", strסԺ��, str����)
    
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.��Ժ�Ǽ�, GetPass(lng����ID) & "|" & strסԺ��, strReturn) Then Exit Function
    
    '�ı䲡�˵�ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    '��¼���˵���ҳID��Ҳ����˳���
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ",'˳���','''" & strסԺ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    
    ��Ժ�Ǽ�_�Թ� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�Թ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim strסԺ�� As String
    Dim strReturn As String
    On Error GoTo errHand
    strסԺ�� = GetסԺ��(lng����ID)
    
    On Error Resume Next
    '��д��Ժ�ǼǼ�¼
    gstrSQL = " Insert Into UnInHosRegister(InHosRegisterNo,UnRegisterReason) " & _
              " Values ('" & strסԺ�� & "','������Ժ')"
    gcn�Թ�.Execute gstrSQL
    On Error GoTo errHand
    
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.������Ժ�Ǽ�, GetPass(lng����ID) & "|" & strסԺ��, strReturn) Then Exit Function
    
    '�ı䲡�˵�ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    ��Ժ�Ǽǳ���_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�Թ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim bln���� As Boolean
    Dim strסԺ�� As String, strReturn As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '�жϸò����Ƿ�������û�н�����Ĳ��˷���Ϊ�㣬˵����Ҫ���þ���Ǽǳ���
        bln���� = False
        gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(����ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�õ��þ���Ǽǳ���", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            bln���� = True
        End If
        
        If Not bln���� Then
            '�޷ѳ�Ժ�Գ�����Ժ��ʽ����
            strסԺ�� = GetסԺ��(lng����ID)
            
            On Error Resume Next
            '��д��Ժ�ǼǼ�¼
            gstrSQL = " Insert Into UnInHosRegister(InHosRegisterNo,UnRegisterReason) " & _
                      " Values ('" & strסԺ�� & "','������Ժ')"
            gcn�Թ�.Execute gstrSQL
            On Error GoTo errHand
            
            If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.������Ժ�Ǽ�, GetPass(lng����ID) & "|" & strסԺ��, strReturn) Then Exit Function
        End If
    End If
    
    '�ı䲡�˵�ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    ��Ժ�Ǽ�_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�Թ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '�ı䲡�˵�ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    ��Ժ�Ǽǳ���_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �����ϴ�_�Թ�(ByVal str���ݺ� As String, ByVal int���� As Integer, ByVal int״̬ As Integer, Optional ByVal lng����ID As Long = 0, Optional ByVal bln���� As Boolean = False) As Boolean
    '���lng����ID��Ϊ�㣬������ϴ��ò��˵Ĵ�����ϸ
    '�Թ�ҽ������ֱ��¼�븺��
    'todo:����ǰ�û���ṹδ���ǵ��ಡ�˵������
    Dim strNO As String
    Dim strסԺ�� As String
    Dim rsTmp   As ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim rsHead As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not bln���� Then
        '����Ƿ����δ�������Ŀ������
        gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
        If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
            gstrSQL = " Select 1 " & _
                  " From סԺ���ü�¼ A,(Select * From ����֧����Ŀ Where ����=" & TYPE_�Ĵ��Թ� & ") B,�����ʻ� C,������ҳ D,������Ϣ E,�շ�ϸĿ F,���ű� G" & _
                  " Where A.NO='" & str���ݺ� & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
                  IIf(lng����ID = 0, "", " And A.����ID=[2]") & _
                  " And E.����ID=D.����ID And E.��ҳID=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
                  " And C.����ID=A.����ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And A.�շ�ϸĿID=F.ID" & _
                  " And C.����=[1] And Nvl(A.�Ƿ��ϴ�,0)=0" & _
                  " And B.��Ŀ���� Is NULL And Rownum<2"
        Else
            gstrSQL = " Select 1 " & _
                  " From סԺ���ü�¼ A,(Select * From ����֧����Ŀ Where ����=" & TYPE_�Ĵ��Թ� & ") B,�����ʻ� C,������ҳ D,������Ϣ E,�շ�ϸĿ F,���ű� G" & _
                  " Where A.NO='" & str���ݺ� & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
                  IIf(lng����ID = 0, "", " And A.����ID=[2]") & _
                  " And E.����ID=D.����ID And E.סԺ����=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
                  " And C.����ID=A.����ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And A.�շ�ϸĿID=F.ID" & _
                  " And C.����=[1] And Nvl(A.�Ƿ��ϴ�,0)=0" & _
                  " And B.��Ŀ���� Is NULL And Rownum<2"
        End If
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δ�������Ŀ������", TYPE_�Ĵ��Թ�, lng����ID)
        If rsCheck.RecordCount <> 0 Then
            MsgBox "�ô����д���δ�������Ŀ�����飡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�򿪴�������
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then

        gstrSQL = " Select A.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,SUM(A.ʵ�ս��) AS ���,G.���� AS ��������,A.������" & _
              " From סԺ���ü�¼ A,�����ʻ� C,������ҳ D,������Ϣ E,���ű� G" & _
              " Where A.NO='" & str���ݺ� & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
              IIf(lng����ID = 0, "", " And A.����ID=[2]") & _
              " And E.����ID=D.����ID And E.��ҳID=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
              " And C.����ID=A.����ID And C.����=[1] And Nvl(A.�Ƿ��ϴ�,0)=0" & _
              " Group by A.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,G.����,A.������"
    Else
        gstrSQL = " Select A.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,SUM(A.ʵ�ս��) AS ���,G.���� AS ��������,A.������" & _
              " From סԺ���ü�¼ A,�����ʻ� C,������ҳ D,������Ϣ E,���ű� G" & _
              " Where A.NO='" & str���ݺ� & "' And A.��¼����=" & int���� & " And A.��¼״̬=" & int״̬ & _
              IIf(lng����ID = 0, "", " And A.����ID=[2]") & _
              " And E.����ID=D.����ID And E.סԺ����=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
              " And C.����ID=A.����ID And C.����=[1] And Nvl(A.�Ƿ��ϴ�,0)=0" & _
              " Group by A.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,G.����,A.������"
    End If
    Set rsHead = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ҫ�ϴ��Ĵ�������", TYPE_�Ĵ��Թ�, lng����ID)
    
    With rsHead
        Do While Not .EOF
            '��������(���ʵ���,��Ժ�ǼǺ�,��������,ҽ��,��������,���,��Ʊ��־)
            'Insert Into InHosBill
            '(InHosBillNo,InHosRegisterNO,DepartmentName,DoctorName,SickName,Amount,RedBillFlag)
            'Values
            '()
            strNO = !NO & !��¼���� & !��¼״̬
            strסԺ�� = GetסԺ��(!����ID)
            On Error Resume Next
            gstrSQL = " Insert Into InHosBill" & _
                    " (InHosBillNo,InHosRegisterNO,DepartmentName,DoctorName,SickName,Amount,RedBillFlag)" & _
                    " Values" & _
                    "('" & strNO & "','" & strסԺ�� & "','" & Nvl(!��������) & "','" & Nvl(!������) & "',''," & Format(!���, "#####0.00;-#####0.00;0.00") & ",1)"
            gcn�Թ�.Execute gstrSQL
            
            On Error GoTo errHand
            '�򿪴�����ϸ��
            gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
            If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
                gstrSQL = " Select A.NO,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,B.��Ŀ����,A.����*A.���� AS ����,A.ʵ�ս��,F.���� AS ��Ŀ����,F.���㵥λ," & _
                      " G.���� AS ��������,A.������" & _
                      " From סԺ���ü�¼ A,����֧����Ŀ B,�����ʻ� C,������ҳ D,������Ϣ E,�շ�ϸĿ F,���ű� G" & _
                      " Where A.NO=[1] And A.��¼����=[2] And A.��¼״̬=[3] And A.����ID=[4]" & _
                      " And E.����ID=D.����ID And E.��ҳID=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
                      " And C.����ID=A.����ID And A.�շ�ϸĿID=B.�շ�ϸĿID And A.�շ�ϸĿID=F.ID" & _
                      " And C.����=B.���� And C.����=[5] And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.ʵ�ս��,0)<>0" & _
                      " Order by A.NO,A.��¼����,A.��¼״̬,A.���"
                      
            Else
                gstrSQL = " Select A.NO,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,B.��Ŀ����,A.����*A.���� AS ����,A.ʵ�ս��,F.���� AS ��Ŀ����,F.���㵥λ," & _
                      " G.���� AS ��������,A.������" & _
                      " From סԺ���ü�¼ A,����֧����Ŀ B,�����ʻ� C,������ҳ D,������Ϣ E,�շ�ϸĿ F,���ű� G" & _
                      " Where A.NO=[1] And A.��¼����=[2] And A.��¼״̬=[3] And A.����ID=[4]" & _
                      " And E.����ID=D.����ID And E.סԺ����=D.��ҳID And A.����ID=E.����ID And A.��������ID=G.ID(+)" & _
                      " And C.����ID=A.����ID And A.�շ�ϸĿID=B.�շ�ϸĿID And A.�շ�ϸĿID=F.ID" & _
                      " And C.����=B.���� And C.����=[5] And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.ʵ�ս��,0)<>0" & _
                      " Order by A.NO,A.��¼����,A.��¼״̬,A.���"
            End If
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ҫ�ϴ��Ĵ�����ϸ", CLng(!NO), CLng(!��¼����), CLng(!��¼״̬), CLng(!����ID), TYPE_�Ĵ��Թ�)
            
            With rsDetail
                Do While Not .EOF
                    '������ϸ(���ʵ���,��ϸ���,ҽ������,ҽԺ��Ŀ����,����,����,���,��λ)
                    'Insert Into InHosBillDetail
                    '(InHosBillNO,InHosBillDetailNO,ItemNO,HosItemName,Price,Quantity,Amount,Spec)
                    'Values
                    '()
                    
                    On Error Resume Next
                    gstrSQL = " Insert Into InHosBillDetail" & _
                              " (InHosRegisterNO,InHosBillNO,InHosBillDetailNO,ItemNO,HosItemName,Price,Quantity,Amount,Spec)" & _
                              " Values" & _
                              "('" & strסԺ�� & "','" & strNO & "','" & !��� & "','" & !��Ŀ���� & "','" & ToVarchar(!��Ŀ����, 100) & "'," & Format(!ʵ�ս�� / !����, "#####0.00000;-#####0.00000;0.00") & "," & _
                              Format(!����, "#####0.00000;-#####0.00000;0.00") & "," & Format(!ʵ�ս��, "#####0.00000;-#####0.00000;0.00") & ",'" & Nvl(!���㵥λ) & "')"
                    gcn�Թ�.Execute gstrSQL
                    On Error GoTo errHand
                    
                    gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    .MoveNext
                Loop
            End With
            
            .MoveNext
        Loop
    End With
    
    �����ϴ�_�Թ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    �����ϴ�_�Թ� = True
End Function

Public Function סԺ�������_�Թ�(ByVal rsExse As ADODB.Recordset, ByVal lng����ID As Long) As String
    Dim int�������� As Integer
    Dim strסԺ�� As String
    Dim strReturn As String
    Dim intRecur As Integer         '�Ƿ�ԭ�˸���
    Dim intCureKindCode As Integer  '���������Ժ��ʽ
    Dim bln�����ʻ� As Boolean
    Dim arrBalance
    
    Dim cur�ܷ��� As Currency
    Dim rsUpload As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Const int�ܽ�� As Integer = 0
    Const intȫ�Է� As Integer = 1
    Const int�����Ը� As Integer = 2
    Const int����ͳ�� As Integer = 3
    Const int���� As Integer = 4
    Const int�������� As Integer = 5
    Const intͳ��֧�� As Integer = 6
    Const intͳ���Ը� As Integer = 7
    Const int���ⶥ As Integer = 8
    Const int�ʻ�֧�� As Integer = 9
    Const int�ֽ�֧�� As Integer = 10
    On Error GoTo errHand
    
    '������ϸ�ϴ�(����ȡ����δ�ϴ��Ĵ���)
    gstrSQL = " Select NO,��¼����,��¼״̬,count(*) Records From סԺ���ü�¼ A,������Ϣ B " & _
              " Where A.����ID=[1] And A.����ID=B.����ID And A.��ҳID=B.סԺ���� " & _
              " And Nvl(��¼״̬,0)<>0 And Nvl(ʵ�ս��,0)<>0 And Nvl(�Ƿ��ϴ�,0)=0" & _
              " Having Count(*)>0" & _
              " Group by NO,��¼����,��¼״̬"
    Set rsUpload = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", lng����ID)
    With rsUpload
        Do While Not .EOF
            If Not �����ϴ�_�Թ�(!NO, !��¼����, !��¼״̬, 0, True) Then Exit Function
            .MoveNext
        Loop
    End With
    
    'ȡ���η����ܶ�
    gstrSQL = " Select Nvl(A.���,0) AS �����ܶ� From ����δ����� A,������Ϣ B" & _
              " Where A.����ID = [1] And A.����ID=B.����ID And A.��ҳID=B.סԺ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���η����ܶ�", lng����ID)
    If rsTemp.RecordCount <> 0 Then
        cur�ܷ��� = rsTemp!�����ܶ�
    Else
        cur�ܷ��� = 0
    End If
    
    bln�����ʻ� = (MsgBox("�Ƿ�ʹ�ø����ʻ�֧����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    
    '��������ѳ�Ժ�����ǳ�Ժ���㣻����Ϊ��;����(�������͡�0���㣬1�н�)
    gstrSQL = "Select A.����ID,A.��ҳID,��Ժ����,��Ժ��ʽ From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ����", lng����ID)
    strסԺ�� = GetסԺ��(rsTemp!����ID)
    If IsNull(rsTemp!��Ժ����) Then
        int�������� = 1
        intCureKindCode = 2
    Else
        int�������� = 0
        Select Case rsTemp!��Ժ��ʽ
        Case "����"
            intCureKindCode = 3
        Case "��ת"
            intCureKindCode = 1
        Case Else
            intCureKindCode = 0
        End Select
    End If
    
    '�Բ���ID��Ϊ���㵥�š���Ժ�վݺŽ���סԺԤ����
    gstr����� = Left(rsTemp!����ID & "_" & rsTemp!��ҳID, 16) & "_" & Mid(CStr(Get����(rsTemp!����ID)), 1, 3)
    '��ɾ����ǰδ����Ľ����¼
'    gstrSQL = "Delete InHosBalance Where InHosBalanceNO='" & gstr����� & "' And InHosRegisterNO='" & strסԺ�� & "'"
'    gcn�Թ�.Execute gstrSQL
    
    'IsRecru:�����Ա���Ϊ�����Ҽ��˲о��ˣ�4����1��ʾԭ�˸���
    'CureKindCode:�������0-����;1-��ת;2-δ��;3-�������н�һ��Ϊ2�����ʼ��HIS����Ϊǰ�û���CureKind��
    intRecur = 0
    gstrSQL = "Select ��Ա��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ��˲о���", lng����ID, TYPE_�Ĵ��Թ�)
    If Nvl(rsTemp!��Ա���, 1) = 4 Then
        '�˲о���
        If MsgBox("�ò�����ԭ�˸���סԺ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intRecur = 1
        End If
    End If
    
    gstrSQL = " Insert Into InHosBalance(InHosBalanceNO,InHosRegisterNO,InvoiceNO,PayType,RedBillFlag,occurdate,IsRecur,CureKindCode)" & _
              " Values ('" & gstr����� & "','" & strסԺ�� & "','" & gstr����� & "'," & int�������� & ",1," & _
              " to_date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & intRecur & "," & intCureKindCode & ")"
    gcn�Թ�.Execute gstrSQL
    
    '����סԺԤ����ӿ�
    '�ܽ��|ȫ�Է�|���������Ը�|����ͳ�����(��������) |ִ������|ʵ�����߽��|ͳ��֧��|ͳ���Ը�|���ⶥ�߷���|�����ʻ�֧��|�ֽ�֧��
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.סԺԤ����, GetPass(lng����ID) & "|" & strסԺ�� & "|" & int�������� & "|3|0|" & IIf(bln�����ʻ�, 1, 0), strReturn) Then Exit Function
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = Val(arrBalance(int�ܽ��))
    cur_������Ϣ.ͳ��֧�� = Val(arrBalance(intͳ��֧��))
    cur_������Ϣ.�����ʻ� = Val(arrBalance(int�ʻ�֧��))
    If Format(cur_������Ϣ.�ܽ��, "#####0.00") <> Format(cur�ܷ���, "#####0.00") Then
        If MsgBox("HIS�����ܶ���ҽ�������ܶ�ȣ��Ƿ�������㣿" & vbCrLf & _
        "HIS��" & Format(cur�ܷ���, "#####0.00") & Space(10) & "ҽ����" & Format(cur_������Ϣ.�ܽ��, "#####0.00"), vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    If cur_������Ϣ.ͳ��֧�� <> 0 Then סԺ�������_�Թ� = סԺ�������_�Թ� & "|ҽ������;" & cur_������Ϣ.ͳ��֧�� & ";0"
    If cur_������Ϣ.�����ʻ� <> 0 Then סԺ�������_�Թ� = סԺ�������_�Թ� & "|�����ʻ�;" & cur_������Ϣ.�����ʻ� & ";0"
    If סԺ�������_�Թ� <> "" Then סԺ�������_�Թ� = Mid(סԺ�������_�Թ�, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�Թ�(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim int�������� As Integer
'    Dim str��ʱ���㵥�� As String
'    Dim str���㵥�� As String
    Dim strסԺ�� As String
    Dim str��Ʊ�� As String
    Dim lng��ҳID As Long
    
    Dim strReturn As String
    Dim arrBalance
    
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�ܷ���_HIS As Currency, cur�ܷ��� As Currency, curҽ������ As Currency, cur�����ʻ� As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '���±����Ͳ����й�
    Dim lng���� As Long, str���� As String
    Dim rsSelected As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    
    Const int�ܽ�� As Integer = 0
    Const intȫ�Է� As Integer = 1
    Const int�����Ը� As Integer = 2
    Const int����ͳ�� As Integer = 3
    Const int���� As Integer = 4
    Const int�������� As Integer = 5
    Const intͳ��֧�� As Integer = 6
    Const intͳ���Ը� As Integer = 7
    Const int���ⶥ As Integer = 8
    Const int�ʻ�֧�� As Integer = 9
    Const int�ֽ�֧�� As Integer = 10
    On Error GoTo errHand
    '��������ѳ�Ժ�����ǳ�Ժ���㣻����Ϊ��;����(�������͡�0���㣬1�н�)
    gstrSQL = "Select A.����ID,A.��ҳID,��Ժ���� From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ����", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    strסԺ�� = GetסԺ��(rsTemp!����ID)
'    str��ʱ���㵥�� = "NO" & lng����ID
    If IsNull(rsTemp!��Ժ����) Then
        int�������� = 1
    Else
        int�������� = 0
    End If
    
    '--ҽ�����Ĳ����ķ�Ʊ��--
'    '��ȡ���ν��㷢Ʊ������㵥��
'    gstrSQL = "Select NO,ʵ��Ʊ�� From ���˽��ʼ�¼ Where ID=" & lng����ID
'    Call OpenRecordset(rsTemp, "��ȡ���㵥��")
'    str���㵥�� = rsTemp!NO
'    str��Ʊ�� = Nvl(rsTemp!ʵ��Ʊ��)
'    If str��Ʊ�� = "" Then
'        MsgBox "��Ʊ�Ų���Ϊ�գ�", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '�����㵥�Ÿ�Ϊ���ʵ��ţ���Ժ�վݺŸ�Ϊ��Ʊ��
'    gstrSQL = "Update InHosBalance Set InHosBalanceNO='" & str���㵥�� & "',InvoiceNO='" & str��Ʊ�� & "' Where InHosBalanceNO='" & str��ʱ���㵥�� & "'"
'    gcn�Թ�.Execute gstrSQL
    '-------------------------
    
    'ÿ�ν��㶼Ҫѡ���֣���ȷ��һЩ�����շ���Ŀ
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS ����,A.SickName AS ����,A.SickSpell AS ���� " & _
            " From SickDefine A Where 1=2"
    Call OpenRecordset_OtherBase(rsSelected, "��ȡ��ѡ��Ĳ���", gstrSQL, gcn�Թ�)
    gstrSQL = " Select A.SickSerialNo AS ID,A.SickNum AS ����,A.SickName AS ����,A.SickSpell AS ���� " & _
            " From SickDefine A Where 1=1"
    Set rs���� = New ADODB.Recordset
    Call OpenRecordset_OtherBase(rs����, "�����֤", gstrSQL, gcn�Թ�)
    
    If rs����.RecordCount > 0 Then
VirusSelect:
        If frm�ಡ��ѡ��_�Թ�.ShowSelect(rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�", rsSelected, False, gcn�Թ�) = True Then
            lng���� = 0
            str���� = ""
            With rs����
                If .RecordCount <> 0 Then .MoveFirst
                lng���� = rs����("ID")
                Do While Not .EOF
                    str���� = str���� & "|" & rs����!ID
                    .MoveNext
                Loop
                If str���� <> "" Then str���� = Mid(str����, 2)
            End With
        Else
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "����Ҫѡ���֣�", vbInformation, gstrSysName
            GoTo VirusSelect
        End If
    End If
    
    Call InsertDisease("InHosSick", gstr�����, str����)
    
    '����סԺ����ӿ�
    '����סԺԤ����ӿ�
    '�ܽ��|ȫ�Է�|���������Ը�|����ͳ�����(��������) |ִ������|ʵ�����߽��|ͳ��֧��|ͳ���Ը�|���ⶥ�߷���|�����ʻ�֧��|�ֽ�֧��
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.סԺ����, GetPass(lng����ID) & "|" & gstr����� & "|" & int�������� & "|3|0|" & IIf(Val(cur_������Ϣ.�����ʻ�) = 0, 0, 1), strReturn) Then Exit Function
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = Val(arrBalance(int�ܽ��))
    cur_������Ϣ.ȫ�Է� = Val(arrBalance(intȫ�Է�))
    cur_������Ϣ.�����Ը� = Val(arrBalance(int�����Ը�))
    cur_������Ϣ.����ͳ�� = Val(arrBalance(int����ͳ��))
    cur_������Ϣ.�������� = Val(arrBalance(int����))
    cur_������Ϣ.ʵ������ = Val(arrBalance(int��������))
    cur_������Ϣ.ͳ��֧�� = Val(arrBalance(intͳ��֧��))
    cur_������Ϣ.���ⶥ���� = Val(arrBalance(int���ⶥ))
    cur_������Ϣ.�����ʻ� = Val(arrBalance(int�ʻ�֧��))
    
    '���汣�ս����¼
    Call Get�ʻ���Ϣ(TYPE_�Ĵ��Թ�, cur_������Ϣ.����ID, Year(zlDatabase.Currentdate()), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & cur_������Ϣ.����ID & "," & TYPE_�Ĵ��Թ� & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & _
        cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    'g��������.�����Ը�����б���������ﲡ�˾������ͣ�������ⲡ�������ͨ����������¼�ı�ע������ǲ��ֵ�����
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ĵ��Թ� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & "," & cur_������Ϣ.�������� & ",0," & cur_������Ϣ.ʵ������ & "," & cur_������Ϣ.�ܽ�� & "," & cur_������Ϣ.ȫ�Է� & "," & cur_������Ϣ.�����Ը� & "," & _
        cur_������Ϣ.����ͳ�� & "," & cur_������Ϣ.ͳ��֧�� & ",0," & cur_������Ϣ.���ⶥ���� & "," & cur_������Ϣ.�����ʻ� & ",'" & gstr����� & "'," & lng��ҳID & "," & int�������� & ",'" & strסԺ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    סԺ����_�Թ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_�Թ�(ByVal lng����ID As Long) As Boolean
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim lng����ID As Long
    Dim int�������� As Integer
    Dim lng��ҳID As Long, lng����ID As Long
    Dim str���㵥�� As String, str������㵥�� As String
    Dim strסԺ�� As String
    
    Dim strReturn As String
    Dim arrBalance
    Dim rsTemp As New ADODB.Recordset
    
    Const int�ܽ�� As Integer = 0
    Const intȫ�Է� As Integer = 1
    Const int���� As Integer = 2
    Const int�������� As Integer = 3
    Const intͳ��֧�� As Integer = 4
    Const intͳ���Ը� As Integer = 5
    Const int���ⶥ As Integer = 6
    Const int�ʻ�֧�� As Integer = 7
    Const int�ֽ�֧�� As Integer = 8
    On Error GoTo errHand
    
    '��ȡ����ID
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    '��ȡ���㵥�ţ���������
    gstrSQL = "Select ����ID,��ҳID,֧��˳���,��;���� From ���ս����¼ Where ����=2 And ����=[1] And ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���㵥��", TYPE_�Ĵ��Թ�, lng����ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "û���ҵ�ԭʼ���շѼ�¼���޷����סԺ���������", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp!��ҳID
    cur_������Ϣ.����ID = rsTemp!����ID
    strסԺ�� = GetסԺ��(rsTemp!����ID)
    int�������� = Nvl(rsTemp!��;����, 0)
    str������㵥�� = Nvl(rsTemp!֧��˳���)
    gstrSQL = "Select NO from ���˽��ʼ�¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���㵥��", lng����ID)
    str���㵥�� = "HCZY" & rsTemp!NO
    
    On Error Resume Next
    '��������¼������Ʊ��־Ϊ-1
    gstrSQL = " Insert Into InHosBalance(InHosBalanceNO,InHosRegisterNO,InvoiceNO,PayType,RedBillFlag,StrikedBillNO,occurdate)" & _
              " Values ('" & str���㵥�� & "','" & strסԺ�� & "','" & str���㵥�� & "'," & int�������� & ",-1,'" & str������㵥�� & "',to_date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    gcn�Թ�.Execute gstrSQL
    On Error GoTo errHand
    
    '����סԺ��������ӿ�(�ܽ��|ȫ�Է�|����ִ������|����ʵ�����߽��|����ͳ��֧��|����ͳ���Ը�|���峬�ⶥ�߷���|��������˻�֧��|����������˻����|�����ֽ�֧��)
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.סԺ�������, GetPass(cur_������Ϣ.����ID) & "|" & str���㵥�� & "|" & str������㵥��, strReturn) Then Exit Function
    
    arrBalance = Split(strReturn, "|")
    cur_������Ϣ.�ܽ�� = -1 * Val(arrBalance(int�ܽ��))
    cur_������Ϣ.ȫ�Է� = 0
    cur_������Ϣ.�����Ը� = 0
    cur_������Ϣ.����ͳ�� = 0
    cur_������Ϣ.�������� = -1 * Val(arrBalance(int����))
    cur_������Ϣ.ʵ������ = -1 * Val(arrBalance(int��������))
    cur_������Ϣ.ͳ��֧�� = -1 * Val(arrBalance(intͳ��֧��))
    cur_������Ϣ.���ⶥ���� = -1 * Val(arrBalance(int���ⶥ))
    cur_������Ϣ.�����ʻ� = -1 * Val(arrBalance(int�ʻ�֧��))
    
    '���汣�ս����¼
    Call Get�ʻ���Ϣ(TYPE_�Ĵ��Թ�, cur_������Ϣ.����ID, Year(zlDatabase.Currentdate()), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & cur_������Ϣ.����ID & "," & TYPE_�Ĵ��Թ� & "," & Year(zlDatabase.Currentdate()) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & _
        cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    
    'g��������.�����Ը�����б���������ﲡ�˾������ͣ�������ⲡ�������ͨ����������¼�ı�ע������ǲ��ֵ�����
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ĵ��Թ� & "," & cur_������Ϣ.����ID & "," & _
        Year(zlDatabase.Currentdate()) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur_������Ϣ.�����ʻ� & "," & cur����ͳ���ۼ� + cur_������Ϣ.����ͳ�� & "," & _
        curͳ�ﱨ���ۼ� + cur_������Ϣ.ͳ��֧�� & "," & intסԺ�����ۼ� & "," & cur_������Ϣ.�������� & ",0," & cur_������Ϣ.ʵ������ & "," & cur_������Ϣ.�ܽ�� & "," & cur_������Ϣ.ȫ�Է� & "," & cur_������Ϣ.�����Ը� & "," & _
        cur_������Ϣ.����ͳ�� & "," & cur_������Ϣ.ͳ��֧�� & ",0," & cur_������Ϣ.���ⶥ���� & "," & cur_������Ϣ.�����ʻ� & ",'" & str���㵥�� & "'," & lng��ҳID & "," & int�������� & ",'" & strסԺ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�Թ�ҽ��")
    סԺ�������_�Թ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Sub InsertDisease(ByVal strTable As String, ByVal strKey As String, ByVal strDisease As String)
    Dim arrDisease
    Dim intDO As Integer, intCOUNT As Integer
    Dim strInsert As String
    Dim rsTemp As New ADODB.Recordset
    '��ָ�����в��벡�����ݣ���Ժ��RegHosSick�����㲡����InHosSick��
    On Error Resume Next
    
    arrDisease = Split(strDisease, "|")
    intCOUNT = UBound(arrDisease)
    
    '�򿪲��ּ�¼��
    gstrSQL = "Select SickSerialNO,SickName,SickKindCode From SickDefine Where SickSerialNO in (" & Replace(strDisease, "|", ",") & ")"
    Call OpenRecordset_OtherBase(rsTemp, "�򿪲��ּ�¼��", gstrSQL, gcn�Թ�)
    
    '׼����������
    gstrSQL = "Insert Into " & strTable & "(" & IIf(UCase(strTable) = "REGHOSSICK", "InHosRegisterNO", "InHosBalanceNO") & _
              ",SickSerialNO,HosSickName,RowNO) Values ('" & strKey & "',"
    For intDO = 0 To intCOUNT
        rsTemp.Filter = "SickSerialNO=" & Val(arrDisease(intDO))
        strInsert = arrDisease(intDO) & ",'" & rsTemp!SickName & "'," & intDO + 1 & ")"
        gcn�Թ�.Execute gstrSQL & strInsert
    Next
    rsTemp.Filter = 0
End Sub

Private Function Get����(ByVal lng����ID As Long) As Integer
    Dim rsTemp As New ADODB.Recordset
    '����סԺʹ�ã����ڲ�������
    gstrSQL = "Select Nvl(����֤��,0) AS ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", TYPE_�Ĵ��Թ�, lng����ID)
    Get���� = rsTemp!���� + 1
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�Ĵ��Թ� & ",'����֤��','" & Get���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
End Function

Private Function GetסԺ��(ByVal lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡ���˵�סԺ�Ǽ���ˮ��
    gstrSQL = "Select ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵�סԺ��ˮ��", lng����ID, TYPE_�Ĵ��Թ�)
    GetסԺ�� = Nvl(rsTemp!˳���)
End Function

Private Function GetPass(ByVal lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡ���˵����루����������룩
    gstrSQL = "Select ���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", lng����ID, TYPE_�Ĵ��Թ�)
    GetPass = Nvl(rsTemp!����)
End Function
