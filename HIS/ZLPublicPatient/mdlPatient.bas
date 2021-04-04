Attribute VB_Name = "mdlPatient"
Option Explicit

Public gobjSquare As SquareCard  '�����㲿��

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

Public Function GetPatiColor(ByVal strPatiType As String, Optional ByVal lngColor As Long = 0) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ɫ
    '���:strPatiType:��������
    '����:������ɫ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPatiType <> "" Then
        GetPatiColor = gobjDatabase.GetPatiColor(strPatiType)
    Else
        GetPatiColor = lngColor
    End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '���:strAge:��������
    '     strBirthDay:��������
    '     bytTag:����zl_Age_Check�������ص�ѯ�����͵���Ϣ���Ƿ�Ҫǿ����ֹ�����Ǳ���ѯ��.0-����ѯ��,1-��ֹ
    '     strCalcDate:��������,Ĭ�ϼ�������Ϊϵͳ����.
    '���أ�TRUE��FALSE��TRUE:����,FALSE:��ֹ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String, lngTmp As Long
    
    On Error GoTo ErrHand
    strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthday) Then
        If strCalcDate = "" Then
            strSQL = "select Zl_Age_Check([1],[2]) From dual"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthday))
        Else
            strCalcDate = Format(strCalcDate, "YYYY-MM-DD HH:mm")
            strSQL = "select Zl_Age_Check([1],[2],[3]) From dual"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthday), CDate(strCalcDate))
        End If
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    strInfo = Nvl(rsTemp.Fields(0).Value)
    If InStr(1, strInfo, "|") > 0 Then
        lngTmp = Val(Split(strInfo, "|")(0)) '1��ֹ,0��ʾ
        strInfo = Split(strInfo, "|")(1)
        If lngTmp = 1 Or (lngTmp = 0 And bytTag = 1) Then
            MsgBox strInfo & vbCrLf & vbCrLf & "��������������!", vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strInfo & vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    CheckAge = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckIdcard(ByVal strIdcard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, _
    Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����֤����Ϸ���У��
    '��Σ�strIdCard ���֤����
    '���Σ�strBirthday  ��������TrueΪ��������
    '         strSex ��������TrueΪ�Ա�
    '         strErrInfo ��������FalseΪ������Ϣ
    '         datCalc �������� ȱʡ��ϵͳʱ�����
    '���أ�True/False  ���֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)�����򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim xmlDoc As DOMDocument
    Dim xmlRoot As IXMLDOMElement
    
    On Error GoTo errH
     '������֤���Ƿ�Ϸ�
    '--<OUTPUT>
    '--       <BIRTHDAY></BIRTHDAY> //��������
    '--       <SEX></SEX>           //�Ա�
    '--       <AGE></AGE>          //����
    '--     <MSG></MSG>         //���֤�Ϸ����ؿ�(�ɴ����֤���л�ȡ�������ں��Ա�)�����򷵻ش�����Ϣ
    '--</OUTPUT>
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Fun_Checkidcard([1]) As Info From Dual"
    Else
        strSQL = "Select Zl_Fun_Checkidcard([1],[2]) As Info From Dual"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Fun_Checkidcard", strIdcard, datCalc)
    strXML = Trim(Nvl(rsTmp!Info))
    If strXML = "" Then Exit Function
    
    Set xmlDoc = New DOMDocument
    xmlDoc.loadXML (strXML)
    '��ȡXML����
    Set xmlRoot = xmlDoc.selectSingleNode("OUTPUT")
    strErrInfo = xmlRoot.selectSingleNode("MSG").Text
    If strErrInfo <> "" Then Exit Function
    
    strBirthday = xmlRoot.selectSingleNode("BIRTHDAY").Text
    strSex = xmlRoot.selectSingleNode("SEX").Text
    strAge = xmlRoot.selectSingleNode("AGE").Text
    
    CheckIdcard = True
    Exit Function
errH:
 If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function SaveBaseInfo(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal strģ�� As String, Optional ByVal int���� As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������˻�����Ϣ(��ҵ�����ݵ�ͬ������)
    '��Σ�lng����ID-����ID (����Ϊ��/0)
    '         lng����ID-�Һ�ID����ҳID(��Ϊ0)
    '         strName-���� (����Ϊ��)
    '         strSex-�Ա� (����Ϊ��)
    '         strAge-���� (����Ϊ��)
    '         strBirthDay-�������� (����Ϊ��)
    '         strģ��-���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '         int���� 1-����;2-סԺ(lng����ID=0,��Ĭ��Ϊ1;lng����ID<>0,1-lng����IDΪ�Һ�ID,2-lng����IDΪ��ҳID)
    '         strInfo-���˻�����Ϣ���������޸�ԭ��
    '         blnXWHIS-������Ϣ����ʱ�Ƿ����RIS�Ľӿ� =True���ã��ò������ڱ��ⲡ����Ϣ���ظ�����RIS�ӿڣ�
    '         blnEMPI-T EMPIƽ̨�Ѿ�������F-EMPIƽ̨δ����
    ' ���Σ�strInfo:���³ɹ�-��Ϣ�������µı仯��Ϣ(����True); ����ʧ��-��Ϣ����δ�ɹ���ԭ��
    ' ���أ�TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Dim strSQL As String, strSQLProc As String
    Dim blnTrans As Boolean
    Dim lngAgeNum As Long, strAgeUnit As String
    Dim str�Һ�NO As String
    Dim strErr As String, strTip As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRet As Long
    
    strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
    Set cmdTmp = New ADODB.Command
    strSQLProc = "Zl_������Ϣ_������Ϣ����("
'   ����id_In ������Ϣ�䶯.����id%Type,
    strSQLProc = strSQLProc & "" & lng����ID & ","
    Set cmdPara = cmdTmp.CreateParameter("����ID", adVarNumeric, adParamInput, 18, lng����ID)
    cmdTmp.Parameters.Append cmdPara
'   ����id_In Number := Null,
    strSQLProc = strSQLProc & "" & lng����ID & ","
    Set cmdPara = cmdTmp.CreateParameter("����ID", adVarNumeric, adParamInput, 18, lng����ID)
    cmdTmp.Parameters.Append cmdPara
'   ģ��_In   ������Ϣ�䶯.�䶯ģ��%Type,
    strSQLProc = strSQLProc & "'" & strģ�� & "',"
    Set cmdPara = cmdTmp.CreateParameter("�䶯ģ��", adVarChar, adParamInput, 100, strģ��)
    cmdTmp.Parameters.Append cmdPara
'   ����_In   ������Ϣ.����%Type,
    strSQLProc = strSQLProc & "'" & strName & "',"
    Set cmdPara = cmdTmp.CreateParameter("����", adVarChar, adParamInput, 100, strName)
    cmdTmp.Parameters.Append cmdPara
'   �Ա�_In   ������Ϣ.�Ա�%Type,
    strSQLProc = strSQLProc & "'" & strSex & "',"
    Set cmdPara = cmdTmp.CreateParameter("�Ա�", adVarChar, adParamInput, 100, strSex)
    cmdTmp.Parameters.Append cmdPara
'   ����_In   ������Ϣ.����%Type
    strSQLProc = strSQLProc & "'" & strAge & "',"
    Set cmdPara = cmdTmp.CreateParameter("����", adVarChar, adParamInput, 100, strAge)
    cmdTmp.Parameters.Append cmdPara
'   ��������_In ������Ϣ.��������%Type,
    strSQLProc = strSQLProc & "" & "TO_Date('" & strBirthday & "','YYYY-MM-DD HH24:mi')" & ","
    If Not IsDate(strBirthday) Then
        Set cmdPara = cmdTmp.CreateParameter("��������", adVarChar, adParamInput, 18, strBirthday)
    Else
        Set cmdPara = cmdTmp.CreateParameter("��������", adDBTimeStamp, adParamInput, , CDate(strBirthday))
    End If
    cmdTmp.Parameters.Append cmdPara
'   ����_In   number(1)  --1-����;2-סԺ
    strSQLProc = strSQLProc & "" & int���� & ","
    Set cmdPara = cmdTmp.CreateParameter("����", adVarNumeric, adParamInput, 1, int����)
    cmdTmp.Parameters.Append cmdPara
    '�޸�ԭ��_IN    varchar2
    strSQLProc = strSQLProc & "'" & strInfo & "',"
    Set cmdPara = cmdTmp.CreateParameter("�޸�ԭ��", adVarChar, adParamInput, 100, strInfo)
    cmdTmp.Parameters.Append cmdPara
'   ˵��_Out    Out ������Ϣ�䶯.˵��%Type --����
    strSQLProc = strSQLProc & "" & "" & ")"
    Set cmdPara = cmdTmp.CreateParameter("˵��", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_������Ϣ_������Ϣ����"
    
    strInfo = ""
    On Error GoTo errH
    
    'LIS ����֮���������ʼ��������Ϊ����LIS������Ҫ����LIS������ʼ���ӿڣ������п��ܻᵯ����Ϣ�����Է�������֮ǰ��
    Call InitObjLis
    If Not gobjLIS Is Nothing Then
        lngAgeNum = CLng(Val(Trim(strAge)))
        strAgeUnit = Mid(strAge, InStr(strAge, CStr(lngAgeNum)) + Len(CStr(lngAgeNum)))
        If lng����ID <> 0 And int���� = 1 Then
            strSQL = "select NO from ���˹Һż�¼ where ����ID = [1] and ID = [2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˹Һŵ�NO", lng����ID, lng����ID)
            If rsTmp.RecordCount > 0 Then
                str�Һ�NO = rsTmp!NO & ""
            End If
        End If
    End If
    Call CreatePlugInOK(glngModule)
    If (Not gobjLIS Is Nothing) Or (Not gobjPlugIn Is Nothing) Or blnXWHIS Then gcnOracle.BeginTrans: blnTrans = True
    
    Call gobjComlib.SQLTest(App.ProductName, "Zl_������Ϣ_������Ϣ����", strSQLProc)
    cmdTmp.Execute
    Call gobjComlib.SQLTest
    '90816,�޸��°�LIS����
    If Not gobjLIS Is Nothing Then
        If lng����ID <> 0 And int���� = 1 Then
            If Not gobjLIS.ModifyPatientBaseintoLIS(lng����ID, str�Һ�NO, int����, strName, strSex, lngAgeNum, strAgeUnit, strģ��, UserInfo.����, strInfo) Then
                gcnOracle.RollbackTrans: blnTrans = False
                strInfo = "LIS ϵͳ������Ϣ�޸�ʧ�ܣ�����" & strInfo
                Exit Function
            Else
                If strInfo <> "" Then strTip = strInfo  '�ɹ������ʾ
            End If
        Else
            If Not gobjLIS.ModifyPatientBaseintoLIS(lng����ID, CStr(lng����ID), int����, strName, strSex, lngAgeNum, strAgeUnit, strģ��, UserInfo.����, strInfo) Then
                gcnOracle.RollbackTrans: blnTrans = False
                strInfo = "LIS ϵͳ������Ϣ�޸�ʧ�ܣ�����" & strInfo
                Exit Function
            Else
                If strInfo <> "" Then strTip = strInfo
            End If
        End If
    End If
    'EMPI
    If Not gobjPlugIn Is Nothing Then
        If blnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModule, lng����ID, IIf(int���� = 2, lng����ID, 0), IIf(int���� = 1, lng����ID, 0), strInfo)  '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strErr)
            If Err.Number = 438 Then lngRet = 1
            Err.Clear: On Error GoTo 0
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModule, lng����ID, IIf(int���� = 2, lng����ID, 0), IIf(int���� = 1, lng����ID, 0), strInfo)  '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strErr)
            If Err.Number = 438 Then lngRet = 1
            Err.Clear: On Error GoTo 0
        End If
        If strErr <> "" Or lngRet = 0 Then
            gcnOracle.RollbackTrans
            strInfo = IIf(blnEMPI, "��EMPIƽ̨���²�����Ϣʧ�ܣ�", "��EMPIƽ̨����������Ϣʧ�ܣ�") & vbCrLf & IIf(strErr <> "", strErr, strInfo)
            Exit Function
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    If blnXWHIS Then
        'RIS 118004
        If CreateXWHIS() Then
            If gobjXWHIS.HISModPati(int����, lng����ID, lng����ID) <> 1 Then
                strTip = strTip & "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��"
            End If
        ElseIf gblnXW = True Then
            strTip = strTip & "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��"
        End If
    End If
    
    strInfo = Trim(Nvl(cmdTmp.Parameters("˵��"), "")) & IIf(strTip <> "", vbCrLf & strTip, "")
    SaveBaseInfo = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Get������Ϣ�ӱ�(ByVal lng����ID As Long, Optional ByVal str��Ϣ�� As String = "") As ADODB.Recordset
'���ܣ�
'    ��ȡ������Ϣ�ӱ���
'����:
    Dim strSQL As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str��Ϣ��, ","))
    If intRet = -1 Then '��ȡ�������дӱ���Ϣ
        strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID =[1] And ��Ϣֵ is Not Null"
    ElseIf intRet = 0 Then '��ȡָ��ĳ���ӱ���Ϣ
        strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID =[1] And ��Ϣ��='" & Split(str��Ϣ��, ",")(0) & "'" & " And ��Ϣֵ is Not Null "
    ElseIf intRet > 0 Then '��ȡָ���Ķ���ӱ���Ϣֵ
        strSQL = "Select ��Ϣ��, ��Ϣֵ" & vbNewLine & _
            "From ������Ϣ�ӱ�" & vbNewLine & _
            "Where ����id = [1] And" & vbNewLine & _
            "      ��Ϣ�� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And ��Ϣֵ is Not Null "
    End If
    
    On Error GoTo errH
    Set Get������Ϣ�ӱ� = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˴ӱ�", lng����ID, str��Ϣ��)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

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

Public Function InitObjLis(Optional ByVal blnMsg As Boolean) As Boolean
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = GetObject(, "zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
    
        On Error Resume Next
        If gobjLIS Is Nothing Then Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
        
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If blnMsg Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
                Exit Function
            End If
        End If
    End If
    InitObjLis = True
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
