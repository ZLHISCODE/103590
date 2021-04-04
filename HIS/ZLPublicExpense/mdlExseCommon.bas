Attribute VB_Name = "mdlExseCommon"
Option Explicit
'*********************************************************************************************************************************************
'����������ش�����
'�ӿ�˵��:
'   1.zlGetSpecialItemFee-���������ѡ����￨����дסԺ���ü�¼ʱ�ı�����Ϣ(�շ����,�շ�ϸĿID,���㵥λ,������ĿID,������Ŀ,�վݷ�Ŀ,ԭ��,�ּ�,�Ƿ���,���ұ�־)
'   2.GetAllAdviceIDsFromDiagnoID���������ID,��ȡ�漰������ҽ��ID
'����:
'����:�ɹ�����true,���򷵻�False
'����:���˺�
'����:2019-11-23 17:33:30
'*********************************************************************************************************************************************
Public grs���㷽ʽ As ADODB.Recordset
Public gobjService As clsService
Public gobjExpenceSvr As clsExpenceSvr
Public gobjBillPrint As Object

Public Function zlGetExpenceSvrObject(ByRef objExpenceSvr As clsExpenceSvr) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:objExpenceSvr-���ط������
    '����:��ȡ����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gobjExpenceSvr Is Nothing Then Set objExpenceSvr = gobjExpenceSvr: zlGetExpenceSvrObject = True: Exit Function
    Set objExpenceSvr = New clsExpenceSvr
    Call objExpenceSvr.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser)
    Set gobjExpenceSvr = objExpenceSvr
    zlGetExpenceSvrObject = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Set objExpenceSvr = Nothing
End Function
Public Function zlGetServiceObject(ByRef objService As clsService) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:objService-���ط������
    '����:��ȡ����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gobjService Is Nothing Then Set objService = gobjService: zlGetServiceObject = True: Exit Function
    Set objService = New clsService
    Call objService.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser)
    Set gobjService = objService
    zlGetServiceObject = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Set objService = Nothing
End Function



Public Function zlGetSpecialItemFee(strClass As String, Optional ByVal strPriceGrade As String, Optional ByVal lng�շ�ϸĿid As Long) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ѡ����￨����дסԺ���ü�¼ʱ�ı�����Ϣ(�շ����,�շ�ϸĿID,���㵥λ,������ĿID,������Ŀ,�վݷ�Ŀ,ԭ��,�ּ�,�Ƿ���,���ұ�־)
    '���:
    '   strClass=�����ѡ����￨��������
    '   strPriceGrade ��ͨ�۸�ȼ�
    '����:ָ�����������ķ��ü�
    '����:���˺�
    '����:2011-07-07 02:17:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strWherePriceGrade As String
    Dim rsTmp As New ADODB.Recordset
  
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    
    If lng�շ�ϸĿid = 0 Then
        strSql = _
            "Select a.��� As �շ����, a.Id As �շ�ϸĿid, a.���㵥λ, c.Id As ������Ŀid, Nvl(a.���ηѱ�, 0) As ���ηѱ�, c.���� As ������Ŀ, c.�վݷ�Ŀ, b.ԭ��, b.�ּ�," & vbNewLine & _
            "       Nvl(b.ȱʡ�۸�, 0) ȱʡ�۸�, Nvl(a.�Ƿ���, 0) As �Ƿ���, Nvl(a.ִ�п���, 0) As ���ұ�־" & vbNewLine & _
            "From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շ��ض���Ŀ D" & vbNewLine & _
            "Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And d.�շ�ϸĿid = a.Id And d.�ض���Ŀ = [1]" & vbNewLine & _
            "      And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    Else
        strSql = _
            "Select a.��� As �շ����, a.Id As �շ�ϸĿid, a.���㵥λ, c.Id As ������Ŀid, Nvl(a.���ηѱ�, 0) As ���ηѱ�, c.���� As ������Ŀ, c.�վݷ�Ŀ, b.ԭ��, b.�ּ�," & vbNewLine & _
            "       Nvl(b.ȱʡ�۸�, 0) ȱʡ�۸�, Nvl(a.�Ƿ���, 0) As �Ƿ���, Nvl(a.ִ�п���, 0) As ���ұ�־" & vbNewLine & _
            "From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C " & vbNewLine & _
            "Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And A.ID = [3]" & vbNewLine & _
            "      And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ض���Ŀ�ķ��ü�", strClass, strPriceGrade, lng�շ�ϸĿid)
    If Not rsTmp.EOF Then Set zlGetSpecialItemFee = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlGetUnitID(bytFlag As Byte, lngID As Long) As Long
'���ܣ������շ��ض���Ŀ��ִ�п���
'������bytFlag=ִ�п��ұ�־,lngID=�շ�ϸĿID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '����ȷ����
            zlGetUnitID = UserInfo.����ID 'ȡ����Ա���ڿ���
        Case 4 'ָ������
            strSql = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                zlGetUnitID = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
            Else
                zlGetUnitID = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
            End If
        Case 1, 2, 3 '���˿���,����Ա����
            zlGetUnitID = UserInfo.����ID '��ȡ����Ա����
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetCardFeeExcuteDeptID(ByVal lng�շ�ϸĿid As Long, ByVal byt���ұ�־ As Byte, Optional ByVal lng���˿���ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ��ұ�־����ȡ��Ӧ��ִ�в���ID
    '���:lng�շ�ϸĿID-�շ�ϸĿID
    '     byt���ұ�־-���ұ�־(0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���)
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-18 11:31:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngUnitID  As Long
    
    On Error GoTo errHandle
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
     Select Case byt���ұ�־
         Case 4 'ָ������
             lngUnitID = zlGetUnitID(byt���ұ�־, lng�շ�ϸĿid)
         Case 1, 2 '���˿���
             If lng���˿���ID <> 0 Then
                 lngUnitID = lng���˿���ID
             Else
                 lngUnitID = UserInfo.����ID
             End If
         Case 0, 3, 5, 6
             lngUnitID = UserInfo.����ID
     End Select
     zlGetCardFeeExcuteDeptID = lngUnitID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFromBalanceName(ByVal str���㷽ʽ As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��㷽ʽ��ȡ������
    '���:str���㷽ʽ-���㷽ʽ����
    '����:
    '����:���ؿ�����
    '����:���˺�
    '����:2018-03-30 15:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    Set objCard = New Card
    Set rsTemp = zlGet���㷽ʽ
    With objCard
        .���㷽ʽ = str���㷽ʽ
        .���� = str���㷽ʽ
        rsTemp.Filter = "����='" & str���㷽ʽ & "'"
        If Not rsTemp.EOF Then
            .�������� = Val(Nvl(rsTemp!����))
            .ȱʡ��־ = Val(Nvl(rsTemp!����ȱʡ)) = 1
        End If
    End With
    Set zlGetCardFromBalanceName = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet���㷽ʽ() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㷽ʽ
    '����:���ؽ��㷽ʽ��Ϣ��
    '����:���˺�
    '����:2018-03-29 17:35:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    
    If Not grs���㷽ʽ Is Nothing Then
        If grs���㷽ʽ.State = 1 Then
            grs���㷽ʽ.Filter = 0
            Set zlGet���㷽ʽ = grs���㷽ʽ: Exit Function
        End If
    End If
    On Error GoTo errHandle
    strSql = "" & _
    "   Select a.����,a.����, a.����,b.Ӧ�ó���,nvl(a.Ӧ����,0) as Ӧ����,nvl(a.Ӧ�տ�,0) as Ӧ�տ�,nvl(a.ȱʡ��־,0) as ȱʡ,nvl(b.ȱʡ��־,0) as  ����ȱʡ" & vbNewLine & _
    "   From ���㷽ʽ a, ���㷽ʽӦ�� b" & vbNewLine & _
    "   Where a.���� = b.���㷽ʽ(+)    "
        
    Set grs���㷽ʽ = zlDatabase.OpenSQLRecord(strSql, "��ȡ���㷽ʽ")
    Set zlGet���㷽ʽ = grs���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get���㷽ʽ(str���� As String, Optional str���� As String) As ADODB.Recordset
    Dim strSql As String, strIF As String
    
    On Error GoTo errH
    
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strIF = "And Instr(','||[2]||',',','||B.����||',')>0 "
        Else
            strIF = "And B.���� = [2]"
        End If
    End If
    strSql = _
        " Select B.����,B.����,Nvl(Nvl(A.ȱʡ��־,B.ȱʡ��־),0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ " & _
        " And  B.����<>7    " & strIF
    If InStr(1, str����, ",9") > 0 Then
        strSql = strSql & " Union " & _
                 " Select ����,����,Nvl(ȱʡ��־,0) As ȱʡ,Nvl(����,1) as ����,Nvl(Ӧ����,0) as Ӧ���� " & _
                 " From ���㷽ʽ " & _
                 " Where ����=9 " & _
                 " Order by ����,����"
    Else
        strSql = strSql & " Order by ����,lpad(����,3,' ')"
    End If
    Set Get���㷽ʽ = zlDatabase.OpenSQLRecord(strSql, App.ProductName, str����, str����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAllAdviceIDsFromDiagnoID(ByVal str���IDs As String, ByRef strҽ��Ids_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ID,��ȡ�漰������ҽ��ID
    '���:str���IDs-���ID,����ö���
    '����:strҽ��Ids_Out-ҽ��ID,����ö���
    '����:��ȡ��Ϸ���true,���򷵻�False
    '����:���˺�
    '����:2019-01-23 18:11:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    If zlGetServiceObject(objService) = False Then Exit Function
    GetAllAdviceIDsFromDiagnoID = objService.zlCisSvr_GetAdviceidsFromDiag(str���IDs, strҽ��Ids_Out)
End Function

 
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False, _
    Optional strLogFunName As String, Optional strLogName As String)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '     strLogName-��־�����
    '     strLogFunName-��־������
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSql As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSql = cllProcs(i)
        If strLogFunName <> "" Then Call WritLog(strLogName, strLogFunName, "zlExecuteProcedureArrAy", strSql)
        Call zlDatabase.ExecuteProcedure(strSql, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Sub zlAddArray(ByRef clldata As Collection, ByVal strSql As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = clldata.Count + 1
    clldata.Add strSql, "K" & i
End Sub





Public Sub zlBillPrint_Initialize(Optional ByVal lngModul As Long)
    '����:���÷�Ʊ��ӡ���֮��ʼ���ӿ�
    '���:
    '   lngModul=ģ��ţ������շ�=1121�����ղ������=1124������=1137
    '�����:140948
    Dim blnInitSuccess As Boolean
    
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    If gobjBillPrint Is Nothing Then
        Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    End If
    If gobjBillPrint Is Nothing Then Exit Sub
    If lngModul = 0 Then Exit Sub
    
    blnInitSuccess = gobjBillPrint.zlInitialize(gcnOracle, glngSys, lngModul, UserInfo.���, UserInfo.����)
    If blnInitSuccess = False Then
        '��ʼ�������ɹ�,����Ϊ�����ڴ���
        Set gobjBillPrint = Nothing: Exit Sub
    End If
End Sub

Public Function zlBillPrint_EraseBill(ByVal strNOs As String, ByVal lngBalanceID As Long) As Boolean
    '����:���÷�Ʊ��ӡ���֮���Ϸ�Ʊ
    '���:
    '   strNOs=�����շѡ����ղ�����㣺�Զ��ŷָ��Ĵ����ŵĶ�����ݺ�:'F0000001','F0000002',...
    '   lngBalanceId=���ʣ����ʵ�ID
    '�����:140948
    
    On Error GoTo ErrHandler
    If gobjBillPrint Is Nothing Then zlBillPrint_EraseBill = True: Exit Function
    If gobjBillPrint.zlEraseBill(strNOs, lngBalanceID) = False Then Exit Function
    
    zlBillPrint_EraseBill = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlBillPrint_Terminate() As Boolean
    '����:���÷�Ʊ��ӡ���֮��ֹ�ӿ�
    '�����:140948
    
    On Error GoTo ErrHandler
    If gobjBillPrint Is Nothing Then zlBillPrint_Terminate = True: Exit Function
    If gobjBillPrint.zlTerminate() = False Then Exit Function
    Set gobjBillPrint = Nothing
    
    zlBillPrint_Terminate = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMedicalGroupID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��������ID As Long, ByVal str������ As String, ByVal dt����ʱ�� As Date) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ȡ��Ӧ��ҽ��С��ID
    '���:
    '   dt����ʱ��=���÷���ʱ��
    '����:
    '����:��ȡ����ҽ��С��ID
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng��id As Long
    
    On Error GoTo ErrHandler
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetMedicalGroupID(lng����ID, lng��ҳID, _
        lng��������ID, str������, dt����ʱ��, lng��id) = False Then Exit Function
        
    ZlGetMedicalGroupID = lng��id
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

