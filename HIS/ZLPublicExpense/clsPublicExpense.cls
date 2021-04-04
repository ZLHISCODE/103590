VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public zlBillingWarn As clsBillingWarn  '���ʱ������
Public zlBillOperation As clsBillOperation '���ݲ������
Public zlDockExpense  As clsDockExpense 'ҽ���������
Public zlInvoice     As clsInvoice   '��Ʊ���
'***********************************************************************
'�ӿ�����
'   zlShowMsgBox-���ѿ�
'   zlSelectWholeItems-������Ŀѡ����
'***********************************************************************
 
Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
    If zlBillingWarn Is Nothing Then Set zlBillingWarn = New clsBillingWarn
    If zlBillOperation Is Nothing Then Set zlBillOperation = New clsBillOperation
    If zlDockExpense Is Nothing Then Set zlDockExpense = New clsDockExpense
    If zlInvoice Is Nothing Then Set zlInvoice = New clsInvoice
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not zlBillingWarn Is Nothing Then Set zlBillingWarn = Nothing
    If Not zlBillOperation Is Nothing Then Set zlBillOperation = Nothing
    If Not zlDockExpense Is Nothing Then Set zlDockExpense = Nothing
    If Not zlInvoice Is Nothing Then Set zlInvoice = Nothing
    
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetPatiMoney(ByVal bytType As Byte, ByVal lng����ID As Long, ByRef objPatiFee As clsPatiFeeinfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵���ط�����Ϣ
    '���:bytType-0-����;1-סԺ
    '     lng����ID-����ID
     '����:
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiMoney = GetPatiMoney(bytType, lng����ID, objPatiFee)
End Function
Public Function zlGetFeeMoneyFromAdviceIDs(ByVal strҽ��IDs As String, _
    ByRef dblOutӦ�ս�� As Double, ByRef dblOutʵ�ս�� As Double) As Boolean
    '����:����ҽ��IDs����ȡӦ�պ�ʵ�ս��
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '����:dblOutӦ�ս��-Ӧ�ս��
    '     dblOutʵ�ս��-ʵ�ս��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetFeeMoneyFromAdviceIDs = GetFeeMoneyFromAdviceIDs(strҽ��IDs, dblOutӦ�ս��, dblOutʵ�ս��)
End Function

Public Function zlAdviceIsCharged(ByVal strҽ��IDs As String, _
    ByVal strNos As String, ByRef bytOutChargeStatus As Byte, Optional ByRef strOutδ��ҽ��IDs As String, _
    Optional ByRef bytOutBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ���Ƿ��Ѿ��շ�
    '���:strҽ��IDs-ҽ��ID,����ö��ŷ���
    '     strNOs-���ݺ�,����ö��ŷ���
    '����:bytOutChargeStatus-�շ�״̬(0-�����շ�,1-��ȫ�շ�;2-�����շ�)
    '     strOutδ��ҽ��IDs-����δ�շѻ�δ����˵�ҽ��ID
    '     bytOutBillType:���ص�ǰ�ĵ�������(1-�շѵ�;2-���ʵ�;3-�շѺͼ��ʶ���)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 09:48:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlAdviceIsCharged = AdviceIsCharged(strҽ��IDs, strNos, bytOutChargeStatus, strOutδ��ҽ��IDs, bytOutBillType)
End Function
Public Function zlBillExistNotBalance(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��շѵ����Ƿ����δ�շѵ�
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '����:
    '����:�����д���δ�շѵ�,����true,���򷵻�False
    '����:Ƚ����
    '����:2016-08-25 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBillExistNotBalance = BillExistNotBalance(strNos)
End Function
Public Function zlGetBillChargeStatus(ByVal strNos As String, ByRef bytOutStatus As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ��ݵļƷ�״̬
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '����:bytOutStatus:0-δ�շ�;1-�����շ�/�˷�;2-ȫ���շ�;3-ȫ���˷�
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBillChargeStatus = GetBillChargeStatus(strNos, bytOutStatus)
End Function
Public Function zlGetBalanceStatus(ByVal strNos As String, ByRef bytOutStatus As Byte, _
    Optional bln���� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϼ��ʵ��Ƿ��Ѿ�����(ֻ����ʵ�)
    '���:strNOs:ָ���ĵ��ݺ�,�������,�ö��ŷ���
    '     bln����-������ʵ�
    '����:bytOutStatus:0-δ����;1-���ֽ���;2-ȫ������
    '����:��ȡ�ɹ�,����true,���򷵻�False(��δ�ҵ����ݲ���)
    '����:���˺�
    '����:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceStatus = GetBalanceStatus(strNos, bytOutStatus, bln����)
End Function
Public Function zlGetBalanceExpenseDetails(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutDetails As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵķ�����ϸ����
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������(���õ��ţ��շ�����շ����ơ��շ����������ʽ��շѵ��ۡ����㵥λ��ִ�п��ң�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceExpenseDetails = GetBalanceExpenseDetails(frmMain, lngModule, lng����ID, rsOutDetails)
End Function
Public Function zlGetBalanceInfor(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng����ID As Long, ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����������
    '���:frmMain -����������
    '    lngModule -ģ���
    '    lng����id -����ID
    '����:rsOutDetails-��������( ���㷽ʽ��������������,ҽ�ƿ����ID,���ѿ�,������ˮ��,����˵��,ˢ�����ţ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetBalanceInfor = GetBalanceInfor(frmMain, lngModule, lng����ID, rsOutBalance)
End Function
Public Function zlIsExiseApplyNotAudit(ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    ByRef blnOutExistApplyed As Boolean, ByRef rsOutMsgInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϲ����Ǵ������������¼
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '����:blnOutExistApplyed-true:��ʾ��������������δ��˵����ʼ�¼
    '     rsOutMsgInfor-���������¼(NO,��Ŀ����,��˲���)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-01 14:40:01
    '˵��:���ݲ���ID����ҳID,�жϸò����Ƿ���������˵�,��δ��˵ķ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strInfo As String
    
    On Error GoTo errHandle
    strSql = "" & _
    "   Select distinct A.NO ,D.���� ��Ŀ����,C.���� ��˿���  " & _
    "   From סԺ���ü�¼ A, ���˷������� B,���ű� C,�շ���ĿĿ¼ D" & vbNewLine & _
    "   Where A.����ID=[1] And A.��ҳID=[2] And A.Id = B.����ID And b.״̬=0  " & _
    "           And B.��˲���ID=C.ID And B.�շ�ϸĿID=D.ID" & _
    "   Order by A.NO,C.����"
    Set rsOutMsgInfor = gobjDatabase.OpenSQLRecord(strSql, "����Ƿ������������", lng����ID, lng��ҳId)
    blnOutExistApplyed = Not rsOutMsgInfor.EOF
    zlIsExiseApplyNotAudit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetInvoiceNoToNOs(ByVal strInvioceNo As String, _
    ByRef strOutNos As String, Optional ByVal blnDataMoved As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ�Ʊ�Ż�ȡ�÷�Ʊ����Ӧ�Ĵ�����(��Ҫ������շѵ�)
    '���:strInvioceNo-��Ʊ��
    '     blnDataMoved-�Ƿ��������ʷ�ռ���в�ѯ
    '����: strOutNOs-�����漰���շѵ���(����ö��ŷ���)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-04-12 15:59:32
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strOutNos = ""
    If gSysPara.bytƱ�ݷ������ <> 0 Then
        strSql = "" & _
        "   Select distinct A.NO" & _
        "   From Ʊ�ݴ�ӡ��ϸ A" & _
        "   Where A.Ʊ��=[1] and Ʊ��=1"
        If blnDataMoved Then
            strSql = strSql & vbCrLf & " UNION  ALL " & vbCrLf & _
             Replace(strSql, "Ʊ�ݴ�ӡ��ϸ", "HƱ�ݴ�ӡ��ϸ")
        End If
        strSql = strSql & " Order by NO"
    Else
        strSql = "" & _
         "   Select distinct NO  " & _
         "   From Ʊ�ݴ�ӡ���� A," & _
         "           (   Select Max(M.��ӡID) as ��ӡID " & _
         "               From  Ʊ��ʹ����ϸ M  " & _
         "               Where M.Ʊ��=1 And M.����=1 And M.����=[1]  " & _
         "               Group by M.����" & _
         "               )  Q" & _
         "   Where A.��������=1  And ID=Q.��ӡID "
        If blnDataMoved Then
            strSql = strSql & vbCrLf & " UNION  ALL " & vbCrLf & _
             Replace(Replace(strSql, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����"), "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
        End If
        strSql = strSql & " Order by NO"
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��Ӧ��Ʊ���շѵ��ݺ�", strInvioceNo)
    With rsTemp
        Do While Not .EOF
            strOutNos = strOutNos & "," & Nvl(!NO)
            .MoveNext
        Loop
        If strOutNos <> "" Then strOutNos = Mid(strOutNos, 2)
    End With
    zlGetInvoiceNoToNOs = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetPatiAuditRecord(ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    ByRef rsOutAudit As ADODB.Recordset, Optional ByVal lng��Ŀid As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵ķ���������Ŀ
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     lng��ĿID-��ĿID(��ѡ,���Բ���)
    '����:rsOutAudit-����ָ�����˵ķ���������Ŀ��Ϣ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-01 16:12:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select ��ĿId,ʹ������,��������,ʹ������-�������� ��������  " & _
    "   From ����������Ŀ  " & _
    "   Where ����ID=[1] And ��ҳID=[2]" & IIf(lng��Ŀid <> 0, " And ��ĿID=[3]", "")
    Set rsOutAudit = gobjDatabase.OpenSQLRecord(strSql, "mdlInExse", lng����ID, lng��ҳId, lng��Ŀid)
    zlGetPatiAuditRecord = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlIsExiseNotAuditPriceBill(ByVal bytType As Byte, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    ByVal strҽ��IDs As String, _
    ByRef blnOutExisePriceBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����δ��˵Ļ��۵�
    '���:bytType-1-�շѻ��۵�;2-���ʻ��۵�
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     strҽ��IDs-����ö��ŷ���,����Ϊ��,��ʾ�����������в�ѯ
    '����:blnOutExisePriceBill-true:��ʾ���ڻ��۵�;False:��ʾ�����ڻ��۵�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-01 16:20:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim varPara As Variant
    
    On Error GoTo errHandle
    If strҽ��IDs = "" Then
        If bytType = 1 Then
            strSql = "SELECT 1 From ������ü�¼ where ��¼����=1 and ��¼״̬=0 and ����ID=[1] and Rownum  <2"
        Else
            If lng��ҳId = 0 Then
                strSql = "" & _
                "   SELECT 1 From ������ü�¼ where  ��¼״̬=0 And nvl( ���ʷ���,0)=1 and ����ID=[1]  and Rownum  <2 " & _
                "   UNION ALL " & _
                "   Select 1 From סԺ���ü�¼ Where  ��¼״̬=0 and  nvl( ���ʷ���,0)=1  and �����־<>2 and ����ID=[1] and Rownum  <2"
            Else
                strSql = "SELECT 1 From סԺ���ü�¼ where ��¼״̬=0 and  nvl( ���ʷ���,0)=1   and ����ID=[1] and ��ҳID=[2] and Rownum  <2 "
            End If
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���ݲ���ID�ж��Ƿ���ڻ��۵�", lng����ID, lng��ҳId)
    Else
        If gobjCommFun.ActualLen(strҽ��IDs) >= 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_����, strҽ��IDs, varPara, strSql, "ҽ��ID") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */  1 " & _
            " From (With ҽ������ As (" & strSql & ") " & _
            "        SELECT 1 From ������ü�¼ A,ҽ������ B  Where a.ҽ�����=B.ҽ��ID And A.��¼״̬=0 And nvl(A.���ʷ���,0)=1 and Rownum  <2 " & _
            "        UNION ALL " & _
            "        SELECT 1 From סԺ���ü�¼ A,ҽ������ B  Where a.ҽ�����=B.ҽ��ID And A.��¼״̬=0 And nvl(A.���ʷ���,0)=1 and Rownum  <2 )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID�ж��Ƿ���ڻ��۵�", varPara)
            
        Else
            strSql = "" & _
            " Select /*+ RULE */  1 " & _
            " From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2list([1]))) " & _
            "        SELECT 1 From ������ü�¼ A,ҽ������ B  Where a.ҽ�����=B.ҽ��ID And A.��¼״̬=0 And nvl(A.���ʷ���,0)=1 and Rownum  <2   " & _
            "        UNION ALL " & _
            "        SELECT 1 From סԺ���ü�¼ A,ҽ������ B  Where a.ҽ�����=B.ҽ��ID And A.��¼״̬=0 And nvl(A.���ʷ���,0)=1 and Rownum  <2  )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "����ҽ��ID�ж��Ƿ���ڻ��۵�", strҽ��IDs)
        End If
    End If
    blnOutExisePriceBill = Not rsTemp.EOF
    zlIsExiseNotAuditPriceBill = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetPriceGradeStartType() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�۸�ȼ�����������
    '����: 0-δ����,1-ֻ������վ��,2-ֻ������ҽ�Ƹ��ʽ,3-վ���ҽ�ƿʽ��������
    '����:���˺�
    '����:2016-09-23 11:08:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPriceGradeStartType = GetPriceGradeStartType()
End Function

Public Function zlGetPriceGrade(ByVal strվ�� As String, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    Optional ByVal strҽ�Ƹ��ʽ As String, _
    Optional ByRef strҩƷ�۸�ȼ�_Out As String, _
    Optional ByRef str���ļ۸�ȼ�_Out As String, _
    Optional ByRef str��ͨ��Ŀ�۸�ȼ�_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ҽ�Ƹ��ʽ��վ�㣬��ȡ��Ӧ�ļ۸�ȼ�
    '���:strվ��-��½��վ�㣬���봫�룬����NULLʱ���۸�ȼ�Ϊ���ؿ�
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     strҽ�Ƹ��ʽ:�������ǿգ����Դ���ҽ�Ƹ��ʽ_In��ʽ����ȡ�۸�ȼ�;�����Բ���ID_In����ҳID����ȡ��Ӧ�Ĳ��˵�ҽ�Ƹ��ʽ��
    
    '����:strҩƷ�۸�ȼ�_out-����ҩƷ�۸�ȼ�
    '     str���ļ۸�ȼ�_out-�������ļ۸�ȼ�
    '     str��ͨ��Ŀ�۸�ȼ�_out-������ͨ�շ���Ŀ�۸�ȼ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPriceGrade = GetPriceGrade(strվ��, lng����ID, lng��ҳId, strҽ�Ƹ��ʽ, _
        strҩƷ�۸�ȼ�_Out, str���ļ۸�ȼ�_Out, str��ͨ��Ŀ�۸�ȼ�_out)
End Function

Public Function zlGetRetailPrice(ByVal lng�շ�ϸĿID As Long, _
    ByVal str�۸�ȼ� As String, ByRef dbl���ۼ�_out As Double, ByRef dblδ�ֽ���_out As Double, _
    Optional ByVal lng�ⷿID As Long = 0, _
    Optional ByVal dbl���� As Double = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ���ݼ۸�ȼ���ȡָ���շ���Ŀ�����ۼ۵������Ϣ
    '���:lng�շ�ϸĿid-�շ�ϸĿID
    '     str�۸�ȼ�-�շѼ۸�ȼ�
    '     lng�ⷿid-�ⷿID��ҩƷ���������ϴ���)
    '     dbl����:��ǰ��������(ҩƷ���������ϴ���)��
    '����:dbl���ۼ�_out-�������ۼ۸�
    '     dblδ�ֽ���_out-���ҩƷ������������Ч����ʾ���ݵ�ǰ����ĳ���������������dbl����)���зֽ�ʱ��δ�ֽ��������.
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetRetailPrice = GetRetailPrice(lng�շ�ϸĿID, str�۸�ȼ�, dbl���ۼ�_out, dblδ�ֽ���_out, lng�ⷿID, dbl����)
End Function

Public Function zlShowMsgBox(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intType As Integer) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ϣ��
    '���:frmMain-���õ�������
    '     strInfo=��ʾ��Ϣ,��Ҫ���Ѵ�������,����"^"��ʾ�س�,">"��ʾ����
    '     intType=��Ϣ������=0(ȱʡ)=MsgBox����,1-Ƥ������
    '     blnNoAsk="intType=0"ʱ��Ч����ʾ�Ƿ�ֻ��ʾһ��ȷ����ť,����ѯ�ʷ�ʽ��ʾ�Ǻͷ�
    '����:
    '    intType=0��vbIgnore=���Ҳ�����ʾ,vbCancel=���Ҳ�����ʾ,vbYes=��,vbNo=��
    '    intType=1��vbYes=����,vbNo=����,vbCancel=ȡ��
    '����:���˺�
    '����:2017-11-08 11:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
  zlShowMsgBox = ShowMsgBox_Custom(frmMain, strInfo, blnNoAsk, intType)
End Function
Public Function zlSelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ŀѡ����(ѡ�������)
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:rsOutSel-�ɹ�ʱ,����ѡ��ĳ�����Ŀ(���ֶ�:ϸĿID,����,����,���,��������,ִ�п���....)
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-08 16:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSelectWholeItems = SelectWholeItems(frmMain, lngModule, strPrivs, rsOutSel)
End Function

Public Function zlPatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng����ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln�˷� As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ����֤
    '���:lngModlue-ģ���
    '     dblMoney-���
    '     lng����ID-����ID
    '     bln�˷�-��ǰ�Ƿ��˷Ѳ���
    '     bytDepositShowMode- Ԥ����ʾ��ʽ(0-��������ʾ;1-ֻ��ʾ�������;2-ֻ��ʾסԺ���)
    '     lngDefaultCardTypeID-ȱʡ��ˢ�����
    '     blnFamilyMoney-�Ƿ��ȡ����Ԥ�����
    '     blnOlnyFamilyIDs-true:���鿨��ֻ��ȡ����IDs;False-��Ҫ��ȡ���鿨
    '����:strFamilyPatiIDs-���˼���ID,����ö��ŷָ���79868
    '����:������֤�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-10-24 14:55:59
    '˵��:
    '   һ�������鿨�����������bln�˷�=falseʱ):
    '       1.������ˢ����֤,ֱ�ӷ���True
    '       2.��������ʱ����Ҫ����ˢ����֤��ͬʱ��Ҫ�������루������ʱ,���Ҫ���������)
    '       3.��������ʱ��������ģ������ˢ���鿨���������룬������ʱ,����Ҫ�鿨��������
    '       4.��ʾ����������NԪ�ڱ���ˢ��,�����������뼴��֧��;���������������(������ʱ�����Ҫ���������)
    '  �����˷��鿨��bln�˷�=trueʱ):
    '       1.������ˢ�����ƣ�ֱ�ӷ���true
    '       2.���������˷�ʱ��Ҫˢ����֤,ͬʱ��Ҫ�������루������ʱ,���Ҫ���������)
    '       3.���������˷�ʱ��������ģ������ˢ����֤,������ʱ,����Ҫ�鿨��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiIdentify = PatiIdentify(lngModlue, frmMain, lng����ID, curMoney, bln�˷�, bytDepositShowMode, lngDefaultCardTypeID, _
                                               blnFamilyMoney, blnOlnyFamilyIDs, strFamilyPatiIDs_Out)
End Function



