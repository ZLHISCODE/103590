VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************************************************************
'���ܣ�ҩƷ������������ش���
'   *CreateDrugAndStuffSvr:����ҩƷ���ķ��񹫹�����
'   *CreateDrugPacker:����������ҩ��(�Զ���ҩ��)
'   *GetDefaultWindow:��ȡȱʡ��ҩ����������
'   DrugMachine_Charge:�շѻ���ʺ�ҩ�������ϴ�
'   DrugMachine_Del:�˷ѻ����˺�ҩ�������ϴ�
'   Get��ҩ����:���䷢ҩ����
'   InitCommon:��ʼ����ػ�������
'*************************************************************************************************************************************************

Private mlngModule As Long, mstrPrivs As String

Private mobjService As Object 'zlPublicExpense.clsService

Private mobjDrugPacker  As Object ' �Զ���ҩ��(���·�ҩ����)
Private mobjDrugMachine As Object '�Զ���ҩ��(�£�

Private mcllWindows As Collection 'array(ҩ��ID,����)
Private mstr���� As String, mstr�ɴ� As String, mstr�д� As String '��¼���ﲡ�������շѵĴ��ڷ���
Private mlng��ҩ�� As Long, mlng��ҩ�� As Long, mlng��ҩ�� As Long '��¼���ﲡ�������շѵ�ҩ������

Public Function InitCommon(ByVal lngModule As Long, strPrivs As String, ByVal blnCliniqueRoomPay As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ػ�������
    '���:
    '   blnCliniqueRoomPay �Ƿ�����޴���֧��
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    mlngModule = lngModule: mstrPrivs = strPrivs
    If CreateDrugAndStuffSvr(lngModule) = False Then Exit Function
    
    If blnCliniqueRoomPay Then
        'ҩ�������ڷ��䷽ʽ
        mstr�д� = zlDatabase.GetPara(49, glngSys, lngModule)
        mstr���� = zlDatabase.GetPara(50, glngSys, lngModule)
        mstr�ɴ� = zlDatabase.GetPara(51, glngSys, lngModule)
        
        mlng��ҩ�� = Val(zlDatabase.GetPara(18, glngSys, lngModule))
        mlng��ҩ�� = Val(zlDatabase.GetPara(19, glngSys, lngModule))
        mlng��ҩ�� = Val(zlDatabase.GetPara(20, glngSys, lngModule))
    Else
        mstr���� = "": mstr�д� = "": mstr�ɴ� = ""
        mlng��ҩ�� = 0: mlng��ҩ�� = 0: mlng��ҩ�� = 0
    End If
    
    InitCommon = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateDrugAndStuffSvr(ByVal lngMoudle As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩƷ���ķ��񹫹�����
    '���:
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjService Is Nothing Then
        Set mobjService = CreateObject("zlPublicExpense.clsService")
        If mobjService Is Nothing Then
            MsgBox "���ù�������(zlPublicExpense.clsService)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    'zlInitCommon(ByVal lngSys As Long, ByVal lngMoudle As Long, ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String, _
        Optional objService As Object, Optional objStuffService As Object)
    If mobjService.zlInitCommon(glngSys, lngMoudle, gcnOracle, gstrDBUser) = False Then
        MsgBox "���ù�������(zlPublicExpense.clsService)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Set mobjService = Nothing: Exit Function
    End If
    
    CreateDrugAndStuffSvr = True
End Function

Private Function CreateDrugPacker() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ҩ��(�Զ���ҩ��)
    '���:
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String, strMessage As String
    Dim blnSuccess As Boolean

    If Not mobjDrugMachine Is Nothing Or Not mobjDrugPacker Is Nothing Then CreateDrugPacker = True: Exit Function

    On Error Resume Next
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err <> 0 Then Set mobjDrugMachine = Nothing
        
        On Error GoTo 0
        If Not mobjDrugMachine Is Nothing Then
            'Ȩ�޼��
            strPrivs = GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))
            If zlstr.IsHavePrivs(strPrivs, "����") = False Then Set mobjDrugMachine = Nothing: Exit Function
            blnSuccess = mobjDrugMachine.Init(1, objComLib, strMessage)
            If blnSuccess = False Then Set mobjDrugMachine = Nothing: Exit Function
            CreateDrugPacker = True: Exit Function
        End If
    End If
    
    '�ɲ���
    On Error Resume Next
    Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
    If Err <> 0 Then Set mobjDrugMachine = Nothing
    If mobjDrugPacker Is Nothing Then Exit Function
    
    On Error GoTo 0
    blnSuccess = mobjDrugPacker.DYEY_MZ_IniSoap()
    If blnSuccess = False Then Set mobjDrugPacker = Nothing: Exit Function
    
    CreateDrugPacker = True
End Function

Public Sub DrugMachine_Charge(ByVal bytBillType As Byte, ByVal strNOs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շѻ���ʺ�ҩ�������ϴ�
    '���:
    '   bytBillType �������ͣ�1-�շѵ�,2-���ʵ�
    '   strNos ���õ��ݣ���ʽ��A001,A002
    '����:
    '����:
    '˵��:110319,95366
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, strReturn As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '�շѵ�
        If Not mobjDrugMachine Is Nothing Then
            '�����ʽ��1|����1,������1;����2,������2
            strData = "1|" & "8," & Replace(strNOs, ",", ";8,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        '��ʽ������1,������1|����2,������2
        strData = "8," & Replace(strNOs, ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.���, UserInfo.����, 0, strData, strReturn)
    Case 2: '���ʵ�
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        '�����ʽ��1|����1,������1;����2,������2
        strData = "1|" & "9," & Replace(strNOs, ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DrugMachine_Del(ByVal bytBillType As Byte, ByVal varValue As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷ѻ����˺�ҩ�������ϴ�
    '���:
    '   bytBillType �������ͣ�1-�շѵ�,2-���ʵ�
    '   varValue bytBillType=0,������ţ�bytBillType=2,���ݺ�
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim strReturn As String, strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
    Dim lng������� As Long, strNo As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '�շѵ�
        lng������� = varValue
        If Not mobjDrugMachine Is Nothing Then
            '�����˵ļ�ȥ���յľ���ʵ���˵�
            strSQL = "Select Max(Decode(a.��¼״̬, 2, a.Id, 0)) As ����id, -1 * Nvl(Sum(a.���� * a.����), 0) As ��ҩ����" & _
                    " From ������ü�¼ A,(Select Distinct ����ID From ����Ԥ����¼ Where ������� = [1]) B" & _
                    " Where a.����id = b.����ID And Mod(a.��¼����, 10) = 1 And a.�շ���� In('5','6','7')" & _
                    " Group By NO, Nvl(�۸񸸺�, ���)" & _
                    " Having Nvl(Sum(a.���� * a.����), 0) <> 0"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�˷�����", lng�������)
            If rsData.RecordCount = 0 Then Exit Sub
            
            Do While Not rsData.EOF
                strData = strData & ";" & nvl(rsData!����id) & "," & nvl(rsData!��ҩ����)
                rsData.MoveNext
            Loop
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        strSQL = "Select a.No, a.ִ�в���id" & _
                " From ������ü�¼ A, ����Ԥ����¼ B" & _
                " Where a.����id = b.����id And a.��¼״̬=2  And a.�շ���� In('5','6','7') And b.������� = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�˷�����", lng�������)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            If InStr(strData & "|", "|" & nvl(rsData!NO) & "," & nvl(rsData!ִ�в���ID) & "|") = 0 Then
                strData = strData & "|" & nvl(rsData!NO) & "," & nvl(rsData!ִ�в���ID)
            End If
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.���, UserInfo.����, strData, strReturn)
        
    Case 2: '���ʵ�
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        strNo = varValue
        strSQL = "Select Id As ����id, -1 * Nvl(����, 1) * ���� As ��ҩ����" & _
                " From ������ü�¼" & _
                " Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1] And �շ���� In('5','6','7')" & _
                "       And �Ǽ�ʱ�� + 0 = (Select Max(�Ǽ�ʱ��)" & _
                "                       From ������ü�¼" & _
                "                       Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1])"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������", strNo)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            strData = strData & ";" & nvl(rsData!����id) & "," & nvl(rsData!��ҩ����)
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
        
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDefaultWindow(ByVal str��� As String, ByVal lngҩ��ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ��ҩ����������
    '���:
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str���
        Case "5"
            If InStr(mstr����, ":") > 0 Then '������û�д�ҩ��ID
                strTmp = mstr����
            ElseIf mlng��ҩ�� > 0 And mstr���� <> "" Then
                strTmp = mlng��ҩ�� & ":" & mstr����
            End If
        Case "6"
            If InStr(mstr�д�, ":") > 0 Then
                strTmp = mstr�д�
            ElseIf mlng��ҩ�� > 0 And mstr�д� <> "" Then
                strTmp = mlng��ҩ�� & ":" & mstr�д�
            End If
        Case "7"
            If InStr(mstr�д�, ":") > 0 Then
                strTmp = mstr�ɴ�
            ElseIf mlng��ҩ�� > 0 And mstr�ɴ� <> "" Then
                strTmp = mlng��ҩ�� & ":" & mstr�ɴ�
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str���
                Case "5"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function

Public Function Get��ҩ����(ByVal lng����ID As Long, ByVal str��� As String, ByVal lngҩ��ID As Long, _
    Optional ByVal blnFirst As Boolean, Optional ByVal bytBillType As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���䷢ҩ����
    '���:
    '   bytBillType �������ͣ�1-�շѵ�,2-���ʵ�
    '����:
    '����:����ɹ�������true,���򷵻�False
    '˵��:
    '   ����ҩ��ID��ȷ��,��ͬ��ҩ��ID������ͬ�ķ�ҩ����
    '   �޸�ʱ����ԭ�з�ҩ���ڣ��������
    '����˵����
    '    1.ǰ�浥�ݺ�ǰ�����ѷ��䷢ҩ���ڣ�����ǰ���Ϊ׼����֤��ͬҩ����ͬһ�����ڣ�
    '    2.�ж�ָ��������ָ��ҩ����δ��ҩƷ��¼���Ƿ���������ϰ�ķ�ҩ����
    '      a.��ҩ���ڴ��ڣ�����������������ķ�ҩ����
    '      b.��ҩ���ڲ�����:
    '        i:�������ȱʡ�ķ�ҩ���ڣ��������ϰ࣬�򷵻�ȱʡ�ķ�ҩ���ڣ��������δ�ϰ��򷵻�null
    '        ii:���������ȱʡ�ķ�ҩ���ڣ�����ݶ�̬�������0-��æ;1-ƽ������ȡ��ר�ҵķ�ҩ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str��ҩ���� As String
    Dim rsData As ADODB.Recordset, strȱʡ As String
    
    On Error GoTo ErrHandler
    If blnFirst Then Set mcllWindows = New Collection
     
    If InStr(",5,6,7,", str���) = 0 Then Exit Function
        
    '���䴰��ʱ���������ҩ�������ŵ��ݲ�ͬ�������ȱʡ����,����ҩ����ͬ���䵽��ͬ����
    If str��� = "5" Then
        If lngҩ��ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr���� = ""
        mlng��ҩ�� = lngҩ��ID '��¼�ò���ʹ�õ�ҩ��(�����Ѷ�)
    ElseIf str��� = "6" Then
        If lngҩ��ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr�ɴ� = ""
        mlng��ҩ�� = lngҩ��ID
    ElseIf str��� = "7" Then
        If lngҩ��ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr�д� = ""
        mlng��ҩ�� = lngҩ��ID
    End If
    
    '2.���ѷ�����ȡ:����Ƿ���ڸô��ڣ���֤��ͬҩ����ͬһ������
    For i = 1 To mcllWindows.Count
        If mcllWindows(i)(0) = lngҩ��ID Then
            Get��ҩ���� = mcllWindows(i)(1): Exit Function
        End If
    Next
    
    '3.ָ��ʱ�̶�����(ָ����ָû�ж�Ӧҩ���ϰ�ʱָ��)
    Select Case str���
    Case "5"
        If mstr���� <> "" Then
            strȱʡ = mstr����
        ElseIf mlng��ҩ�� > 0 Then
            strȱʡ = GetDefaultWindow(str���, lngҩ��ID)
        End If
    Case "6"
        If mstr�ɴ� <> "" Then
            strȱʡ = mstr�ɴ�
        ElseIf mlng��ҩ�� > 0 Then
            strȱʡ = GetDefaultWindow(str���, lngҩ��ID)
        End If
    Case "7"
        If mstr�д� <> "" Then
            strȱʡ = mstr�д�
        ElseIf mlng��ҩ�� > 0 Then
            strȱʡ = GetDefaultWindow(str���, lngҩ��ID)
        End If
    End Select
    
    '4.��ҩƷϵͳ�Ǳ߷���
    If mobjService.zlDrugSvr_GetSendWindows(1, lng����ID, 0, _
         lngҩ��ID & "," & strȱʡ, rsData, mlngModule) = False Then Exit Function
    If Not rsData.EOF Then str��ҩ���� = nvl(rsData!��ҩ����)
    
    If str��ҩ���� <> "" Then
        Select Case str���
            Case "5"
                mstr���� = str��ҩ����
            Case "6"
                mstr�ɴ� = str��ҩ����
            Case "7"
                mstr�д� = str��ҩ����
        End Select
        
        mcllWindows.Add Array(lngҩ��ID, str��ҩ����), "K" & lngҩ��ID
    End If
    Get��ҩ���� = str��ҩ����
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


 


Private Sub Class_Terminate()
    '�ͷŶ���
    Set mobjService = Nothing
    Set mobjDrugPacker = Nothing
    Set mobjDrugMachine = Nothing
End Sub