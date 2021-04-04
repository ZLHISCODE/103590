VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_DIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ����ͣ���
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjSOAP As Object

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mobjSOAP = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'˵���������ӿ�Ϊ�״̬

    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mobjSOAP Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "���ӵ��ͷ�����ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ���ӵ��ͷ�����", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        Call mobjSOAP.MSSoapInit(mobjBase.URL)
        mobjBase.mobjLog.Add "�ɹ����ӵ��ͷ�����", 1
    Else
        Set mobjSOAP = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Add "���ӵ��ͷ������쳣", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����Filter����
    
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim objDB As Object
    
    On Error GoTo hErr
    
    strInfo = Split(strData, "|")(1)

    Select Case intType
    Case Val("22-��ʼ��ҩ"), Val("23-��ɷ�ҩ")
        If TypeName(mobjBase.mobjComLib) = "clsComLib" Then
            'ZLHIS����������ҩƷ�Զ����豸��������
            Set objDB = mobjBase.mobjComLib.zlDatabase
        ElseIf TypeName(mobjBase.mobjComLib) = "clsPublic" Then
            '֧����
            Set objDB = mobjBase.mobjComLib
        Else
            mobjBase.mobjLog.Add "�������Ͳ���ȷ����ֹ��ȡ��׼����", 1
            mobjBase.mobjLog.Add "" & intType, 1, 1
            mobjBase.mobjLog.Save
            Exit Function
        End If
        
        If Val(strData) = 1 Then
            strSQL = _
                "Select count(1) REC " & vbNewLine & _
                "From All_Tab_Cols " & vbNewLine & _
                "Where Table_Name = 'δ��ҩƷ��¼' And (Lower(Column_Name) = 'groupno' Or Lower(Column_Name) = 'ordertype')"
            mobjBase.mobjLog.Add strSQL, 1, 1
            
            Set rsResult = objDB.OpenSQLRecord(strSQL, "��ȡDIH��д��GroupNO�ֶ��Ƿ����")

            If rsResult!Rec >= 2 Then
                strSQL = _
                    "Select Distinct a.����, a.No ������, a.�ⷿid, b.����id, g.���� ���ڱ���, f.GroupNO, f.OrderType," & _
                    "    Zl_Fun_Drug_Machine(a.�ⷿid, d.ҩƷ����) �ӿڱ�� " & vbNewLine & _
                    "From ҩƷ�շ���¼ A, ������ü�¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ�շ������־ E, δ��ҩƷ��¼ F, ��ҩ���� G " & vbNewLine & _
                    "   , Table(f_Str2list2([1], ';', ',')) I " & vbNewLine & _
                    "Where a.����id = b.Id And a.ҩƷid = c.ҩƷid And c.ҩ��id = d.ҩ��id(+) And a.���� = i.C1 " & vbNewLine & _
                    "    And a.NO = e.������(+) And a.���� = e.����(+) And a.�ⷿid = e.�ⷿid(+) " & vbNewLine & _
                    "    And a.NO = f.NO And a.���� = f.���� And a.�ⷿid = f.�ⷿid And f.��ҩ���� = g.����(+) " & vbNewLine & _
                    "    And e.ҵ�����(+) = [2] And a.No = i.C2 And (Mod(a.��¼״̬, 3) = 0 Or a.��¼״̬ = 1) " & vbNewLine & _
                    "Order By �ӿڱ��, a.NO, a.����, a.�ⷿid "
            Else
                mobjBase.mobjLog.Add "��Ϊ��δ��ҩƷ��¼�����ӡ�GROUPNO��ORDERTYPE�������ֶ�", 1, 1
                strSQL = _
                    "Select Distinct a.����, a.No ������, a.�ⷿid, b.����id, g.���� ���ڱ���, '' GroupNO, '' OrderType, " & _
                    "    Zl_Fun_Drug_Machine(a.�ⷿid, d.ҩƷ����) �ӿڱ�� " & vbNewLine & _
                    "From ҩƷ�շ���¼ A, ������ü�¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ�շ������־ E, δ��ҩƷ��¼ F, ��ҩ���� G " & vbNewLine & _
                    "   , Table(f_Str2list2([1], ';', ',')) I " & vbNewLine & _
                    "Where a.����id = b.Id And a.ҩƷid = c.ҩƷid And c.ҩ��id = d.ҩ��id(+) And a.���� = i.C1 " & vbNewLine & _
                    "    And a.NO = e.������(+) And a.���� = e.����(+) And a.�ⷿid = e.�ⷿid(+) " & vbNewLine & _
                    "    And a.NO = f.NO And a.���� = f.���� And a.�ⷿid = f.�ⷿid And f.��ҩ���� = g.����(+) " & vbNewLine & _
                    "    And e.ҵ�����(+) = [2] And a.No = i.C2 And (Mod(a.��¼״̬, 3) = 0 Or a.��¼״̬ = 1) " & vbNewLine & _
                    "Order By �ӿڱ��, a.NO, a.����, a.�ⷿid "
            End If
            rsResult.Close

            mobjBase.mobjLog.Add strSQL, 1, 1
            Set rsResult = objDB.OpenSQLRecord(strSQL, "��ȡ��ʼ��ҩ����", strInfo, intType - 20)
            mobjBase.mobjLog.Add "��ȡ��ʼ��ҩ�������", 1
        Else
            Set rsResult = Nothing
            mobjBase.mobjLog.Add "סԺ��֧��", 1
        End If
        
        Set clsINF_GetData = rsResult
    
    Case Else
        'ʹ�û��������
        Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    End Select
    
    Exit Function
    
hErr:
    blnFinish = False
    mobjBase.mobjLog.Add "��ȡ��׼����ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����SendData����

'    Const STR_TAKE_END As String = "</outpOrderTake>"
'    Const STR_COMP_END As String = "</outpOrderCompletion>"
    
    Dim strInput As String
    Dim strSingle As String, strKey As String, strTmp As String
    Dim strReturnCode As String, strWindows As String
    Dim blnFinish As Boolean
    Dim objDB As Object
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 5, 21 To 23
        rsData.Filter = ""
    Case Else
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
makExit:
        mobjBase.mobjLog.Add "�����ݿɴ���", 1
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    '��׼���ݴ��Ϳ�ʼ
    mobjBase.mobjLog.Add "��׼���ݴ��Ϳ�ʼ", 1
    
    '��������
    If intType = Val("22-��ʼ��ҩ") Or intType = Val("23-��ɷ�ҩ") Then
        '��ͬ���������ܵ����⴦��
        
        rsData.MoveFirst
        Do While rsData.EOF = False
            blnFinish = True
        
            '��ʽ�����ݣ�������¼��
            mobjBase.mobjLog.Add "��ʽ������", 1
            strInput = FormatData(intType, intIO, strCode, rsData, Val("1-����¼") = 1)
            If strInput = "" Then
                strMessage = "��ʽ������ʧ��"
                'Exit Function
            End If
            
            If SendDataSingle(intType, strInput, strReturnCode, strWindows, strMessage) Then
                '��׼���ݴ������
                mobjBase.mobjLog.Add "��׼���ݴ������", 1
            Else
                '��׼���ݴ���ʧ��
                blnFinish = False
                strMessage = strWindows
                mobjBase.mobjLog.Add strMessage, 1
                mobjBase.mobjLog.Add "���ͽӿ��쳣", 1
                
                '׼����¼������
                With rsTemp
                    mobjBase.mobjLog.Add "׼����¼������", 2
                    If .Fields.Count <= 0 Then
                        .LockType = adLockOptimistic
                        .CursorType = adOpenStatic
                        .CursorLocation = adUseClient
                        
                        .Fields.Append "�ⷿID", adBigInt, , adFldIsNullable
                        .Fields.Append "����", adInteger, , adFldIsNullable
                        .Fields.Append "������", adVarChar, 10, adFldIsNullable
                        .Open
                    End If
                    
                    mobjBase.mobjLog.Add "׼��ʧ�ܱ�ǵ����ݿ�ʼ", 2
                    mobjBase.mobjLog.Add "�ⷿID:" & mdlDrugMachine.NVL(rsData!�ⷿid), 3, 1
                    mobjBase.mobjLog.Add "����:" & mdlDrugMachine.NVL(rsData!����), 3, 1
                    mobjBase.mobjLog.Add "������:" & mdlDrugMachine.NVL(rsData!������), 3, 1
                    
                    .AddNew
                    .Fields!�ⷿid = mdlDrugMachine.NVL(rsData!�ⷿid, 0)
                    .Fields!���� = mdlDrugMachine.NVL(rsData!����)
                    .Fields!������ = mdlDrugMachine.NVL(rsData!������)
                    .Update
                    
                    mobjBase.mobjLog.Add "׼��ʧ�ܱ�ǵ��������", 2
                End With
                                
            End If
            
            '��ǳɹ���ҩƷ�շ���¼����֮����ʧ��Ҳ��ǣ���Ϊ���ýӿڹ������ߵĶ�ʱ���Ϳɲ��䴫�͡�
            If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsTemp, mobjBase, blnFinish) Then
                mobjBase.mobjLog.Add "�����͡���ǳɹ�", 1
            Else
                mobjBase.mobjLog.Add "�����͡����ʧ��", 1
            End If
            If rsTemp.State = adStateOpen Then rsTemp.Close
            
            rsData.MoveNext
        Loop
        
    Else
    
        '��ʽ������
        mobjBase.mobjLog.Add "��ʽ������", 1
        strInput = FormatData(intType, intIO, strCode, rsData)
        If strInput = "" Then
            strMessage = "��ʽ������ʧ��"
            Exit Function
        End If
   
        If SendDataSingle(intType, strInput, strReturnCode, strWindows, strMessage) Then
            If intType = 21 Then
                '��д��ҩ����
                If intAppType = Val("3-֧����") Then
                    Set objDB = mobjBase.mobjComLib
                Else
                    Set objDB = mobjBase.mobjComLib.zlDatabase
                End If
                If mdlDrugMachine.UpdateDispenseWindow(rsData, strWindows, objDB, mobjBase.mobjLog) Then
                    mobjBase.mobjLog.Add "��д��ҩ���ڳɹ�", 1
                Else
                    mobjBase.mobjLog.Add "��д��ҩ����ʧ��", 1
                End If
            End If
        Else
            strMessage = strWindows
            mobjBase.mobjLog.Add strMessage, 1
            mobjBase.mobjLog.Add "���ͽӿ��쳣", 1
        End If
        
        '��ǳɹ���ҩƷ�շ���¼����֮����ʧ��Ҳ��ǣ���Ϊ���ýӿڹ������ߵĶ�ʱ���Ϳɲ��䴫�͡�
        If intType > 20 Then
            If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
                mobjBase.mobjLog.Add "�����͡���ǳɹ�", 1
            Else
                mobjBase.mobjLog.Add "�����͡����ʧ��", 1
            End If
        End If
        
        '��׼���ݴ������
        mobjBase.mobjLog.Add "��׼���ݴ������", 1
        mobjBase.mobjLog.Save
    End If
    
    clsINF_SendData = True
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "��׼���ݴ���ʧ��", 1
    mobjBase.mobjLog.Save
    Exit Function
End Function

Private Function SendDataSingle(ByVal intType As Integer, ByVal strInput As String, _
    ByRef strReturnCode As String, ByRef strWindows As String, _
    ByRef strMessage As String) As Boolean
'���ܣ������������豸�ӿ�
'������
'���أ�True�ɹ���Falseʧ��
    
    Dim strOutput As String
    Dim objXML As New clsXML
    
    On Error GoTo hErr
    
    strReturnCode = ""
    
    Select Case intType
    Case Val("5-��ҩ����")
        strOutput = mobjSOAP.outpWindowNotify(strInput)
        GoSub hResult
        strReturnCode = Val(strOutput)
        
    Case Val("21-��ҩ")
        '�ٷ�����ҩ
        strOutput = mobjSOAP.outpOrderDispense(strInput)
        GoSub hResult
        
        '����������Ϣ
        Call objXML.OpenXMLDocument(strOutput)
        Call objXML.GetNodePropertyValue("status", "code", strReturnCode)
        Call objXML.GetNodePropertyValue("status", "message", strMessage)
        Call objXML.GetSingleNodeValue("windowNo", strWindows)
        objXML.CloseXMLDocument
        
    Case Val("22-��ʼ��ҩ")
        strOutput = mobjSOAP.outpOrderTakeNotify(strInput)
        GoSub hResult
        
        '����������Ϣ
        Call objXML.OpenXMLDocument(strOutput)
        Call objXML.GetNodePropertyValue("status", "code", strReturnCode)
        Call objXML.GetNodePropertyValue("status", "message", strMessage)
        objXML.CloseXMLDocument
        
    Case Val("23-��ɷ�ҩ")
        strOutput = mobjSOAP.outpOrderCompletionNotify(strInput)
        GoSub hResult
        strReturnCode = Val(strOutput)
    
    End Select
    
    SendDataSingle = Val(strReturnCode) = 0 _
                     Or Val(strReturnCode) = 5 And intType = Val("21-��ҩ")
    
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "��׼���ݴ���ʧ��", 1
    mobjBase.mobjLog.Save
    Exit Function
    
hResult:
    mobjBase.mobjLog.Add "�ӿڷ�����Ϣ��", 2
    mobjBase.mobjLog.Add strOutput, 2
    Return
End Function

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, _
    ByVal rsData As ADODB.Recordset, _
    Optional ByVal blnSingle As Boolean = False) As String
'���ܣ�����¼������ļ�¼ת������Ϣƽ̨����θ�ʽ
'������
'  intType��ҵ������
'  intIO��������סԺ
'  strCode���ӿڱ��
'  rsData����¼������
'  blnSingle��True����¼��ʽ����False���м�¼��ʽ��
'���أ���׼��������
    
    Dim objXML As New clsXML
    Dim lngPatientID As Long
    Dim strNO As String
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    lngPatientID = -9999        '��¼���ġ�����ID���ֶ�ֵ����ΪNull����ֹ��֯XML����ȷ
    
    With rsData
        If blnSingle = False Then
            .MoveFirst
        End If
        Do
    
            Select Case intType
            Case Val("5-��ҩ����")
                'ע�⣺����Ҫ��windowNo��Ϊ���֣���ȷ������ҩ���ڡ����ġ����롱ҲΪ�����֡�
            
                objXML.AppendNode "outpWindow"
                
                objXML.AppendData "windowNo", mdlDrugMachine.NVL(!����)
                objXML.AppendData "windowStatus", IIf(Val(mdlDrugMachine.NVL(!�ϰ�)) = 1, "open", "close")
                objXML.AppendData "windowType", "indirect"
                
                objXML.AppendNode "outpWindow", True
                
            Case Val("21-��ҩ")
                '����
                If lngPatientID <> Val(mdlDrugMachine.NVL(!����id)) Then
                    objXML.AppendNode "outpOrder"
                    
                    objXML.AppendNode "patient"
                    objXML.AppendData "windowNo", ""
                    objXML.AppendData "patientID", mdlDrugMachine.NVL(!����id)
                    objXML.AppendData "patientName", mdlDrugMachine.NVL(!����)
                    objXML.AppendData "patientGender", mdlDrugMachine.NVL(!�Ա�)
                    objXML.AppendData "patientAge", mdlDrugMachine.NVL(!����)
                    objXML.AppendData "identity", mdlDrugMachine.NVL(!����)
                    objXML.AppendData "insuranceType", mdlDrugMachine.NVL(!ҽ�����)
                    objXML.AppendData "chargeType", mdlDrugMachine.NVL(!�ѱ�)
                    objXML.AppendNode "patient", True
                
                    objXML.AppendNode "prescriptions"
                    
                    GoTo makRecipe
                    
                Else
makRecipe:
                    '����
                    If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                        objXML.AppendNode "prescription no=""" & mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) & """ " & _
                                          "type=""" & IIf(Val(mdlDrugMachine.NVL(!�Ƿ���)) = 1, "J", "M") & """ " & _
                                          "paymentDT=""" & Format(mdlDrugMachine.NVL(!�ɷ�ʱ��), "yyyy-MM-DD hh:mm:ss") & """"
                        
                        objXML.AppendData "outpNo", ""
                        objXML.AppendData "visitNo", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                        objXML.AppendData "deptCode", mdlDrugMachine.NVL(!��������id)
                        objXML.AppendData "deptName", mdlDrugMachine.NVL(!��������)
                        objXML.AppendData "doctCode", ""
                        objXML.AppendData "doctName", mdlDrugMachine.NVL(!����ҽ��)
                        objXML.AppendData "diagnosis", mdlDrugMachine.NVL(!�������)
                        objXML.AppendData "paymentDT", mdlDrugMachine.NVL(!�ɷ�ʱ��), xsDateTime
                        objXML.AppendNode "drugList"
                        
                        GoTo makDetail
                        
                    Else
makDetail:
                        objXML.AppendNode "drug"
                        
                        objXML.AppendData "drugCode", mdlDrugMachine.NVL(!ҩƷid)
                        objXML.AppendData "drugName", mdlDrugMachine.NVL(!ͨ����)
                        objXML.AppendData "drugSpec", mdlDrugMachine.NVL(!���)
                        objXML.AppendData "firmName", mdlDrugMachine.NVL(!������)
                        objXML.AppendData "amount", mdlDrugMachine.NVL(!���ﵥλ����)
                        objXML.AppendData "takeUnit", mdlDrugMachine.NVL(!���ﵥλ)
                        objXML.AppendData "takeMethod", mdlDrugMachine.NVL(!�÷�)
                        If Val(mdlDrugMachine.NVL(!��������)) = 0 Then
                            objXML.AppendData "takeDosage", ""
                        Else
                            objXML.AppendData "takeDosage", mdlDrugMachine.NVL(!��������) & _
                                                            mdlDrugMachine.NVL(!������λ) & "��" & _
                                                            mdlDrugMachine.NVL(!ִ��Ƶ��)
                        End If
                        objXML.AppendData "takeType", ""
                        objXML.AppendData "takeNote", mdlDrugMachine.NVL(!ҽ������)
                        objXML.AppendData "pharmacyCode", mdlDrugMachine.NVL(!�ⷿid)
                        objXML.AppendData "sortNo", mdlDrugMachine.NVL(!���)
                        
                        objXML.AppendNode "drug", True
                    End If
                        
                End If
            
                lngPatientID = Val(mdlDrugMachine.NVL(!����id))
                strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
            
            Case Val("22-��ʼ��ҩ")
                
                objXML.AppendNode "outpOrderTake"
                
                objXML.AppendData "windowNo", mdlDrugMachine.NVL(!���ڱ���)
                objXML.AppendData "patientID", mdlDrugMachine.NVL(!����id)
                objXML.AppendData "groupNo", mdlDrugMachine.NVL(!groupNo)
                objXML.AppendData "orderType", mdlDrugMachine.NVL(!OrderType)
                
                objXML.AppendNode "outpOrderTake", True
                
            Case Val("23-��ɷ�ҩ")
            
                objXML.AppendNode "outpOrderCompletion"
                
                objXML.AppendData "patientID", mdlDrugMachine.NVL(!����id)
                objXML.AppendData "groupNo", mdlDrugMachine.NVL(!groupNo)
                
                objXML.AppendNode "outpOrderCompletion", True
                
            End Select
        
            
            If blnSingle = False Then
                .MoveNext
            End If
            
            '�����β
            Select Case intType
            Case 21
                If .EOF Then
                    objXML.AppendNode "drugList", True
                    objXML.AppendNode "prescription", True
                    objXML.AppendNode "prescriptions", True
                    objXML.AppendNode "outpOrder", True
                Else
                    If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                        objXML.AppendNode "drugList", True
                        objXML.AppendNode "prescription", True
                    End If
                    If lngPatientID <> Val(mdlDrugMachine.NVL(!����id)) Then
                        objXML.AppendNode "prescriptions", True
                        objXML.AppendNode "outpOrder", True
                    End If
                End If
            Case 22, 23
                If blnSingle Then
                    '������¼��ʽ��
                    Exit Do
                End If
            End Select
            
        Loop While .EOF = False
        
    End With
    
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ���������", 1
    mobjBase.mobjLog.Save
    FormatData = objXML.XmlText
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ������ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
'
End Property

Private Property Get clsINF_Token() As String
'
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
        
    Call CreateSOAP(mobjSOAP, mobjBase)
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
    
    'Call mdlDrugMachine.GetUserInfo(mobjBase.UserName, mobjBase.mobjComLib, mobjBase.mobjLog, UserInfo)
End Property