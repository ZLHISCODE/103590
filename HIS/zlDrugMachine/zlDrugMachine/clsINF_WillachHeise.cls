VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_WillachHeise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ�Τ�ֺ��ģ���
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjSOAP As Object
Private UserInfo As TYPE_USER_INFO

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
            mobjBase.mobjLog.Add "����Τ�ֺ��ķ�����ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ����Τ�ֺ��ķ�����", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        Call mobjSOAP.MSSoapInit(mobjBase.URL)
        mobjBase.mobjLog.Add "�ɹ�����Τ�ֺ��ķ�����", 1
    Else
        Set mobjSOAP = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Add "����Τ�ֺ��ķ������쳣", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Public Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����Filter����
    
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)

End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '�����ʹ�û���ļ�¼�������뽫ǰ��������Σ����������д�µĴ���
    'TODO��
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

Public Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����SendData����

    Dim strInput As String, strOutput As String
    Dim strVal As String, strReturn As String, strWindows As String
    Dim objXML As New clsXML
    Dim blnFinish As Boolean
    Dim objDB As Object
    
    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "ͣ��=0"
    Case 4 To 5
        rsData.Filter = ""
    Case 21 To 24
        rsData.Filter = "�Ƿ���=1"
    Case 25
        rsData.Filter = "�Ƿ���=1 And ����=8"
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
    
    '��ʽ������
    mobjBase.mobjLog.Add "��ʽ������", 1
    strInput = FormatData(intType, intIO, strCode, rsData)
    If strInput = "" Then
        strMessage = "�����ݿɸ�ʽ��"
        Exit Function
    End If
    
    '��׼���ݴ��Ϳ�ʼ
    mobjBase.mobjLog.Add "��׼���ݴ��Ϳ�ʼ(HisTransData)", 1
    
    '��������
    On Error GoTo makFinish
    strOutput = mobjSOAP.HisTransData(strInput)
    On Error GoTo hErr
    mobjBase.mobjLog.Add "�ӿڷ�����Ϣ��", 2
    mobjBase.mobjLog.Add strOutput, 2
    
    '����������Ϣ
    Call objXML.OpenXMLDocument(strOutput)
    Call objXML.GetSingleNodeValue("RETVAL", strVal)                '�豸��������ĵڼ��Ŵ���
    Call objXML.GetSingleNodeValue("RETCODE", strReturn)            '1-ִ�гɹ�������ֵ-ִ��ʧ��
    Call objXML.GetSingleNodeValue("RETMSG", strWindows)            '��ǰ������ҩ���ڱ���������Ϣ
    objXML.CloseXMLDocument
    blnFinish = Val(strReturn) = 1
    
    If blnFinish Then
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
        mobjBase.mobjLog.Add "Τ�ֺ��Ľӿ��쳣", 1
    End If
    
makFinish:
    If Err.Number <> 0 Then
        blnFinish = False
        Err.Clear: On Error GoTo hErr
        mobjBase.mobjLog.Add "Τ�ֺ��Ľӿ��쳣���������쳣", 1
    End If
    
    '��ǳɹ���ҩƷ�շ���¼����֮����ʧ��Ҳ��ǣ���Ϊ���ýӿڹ������ߵĶ�ʱ���Ϳɲ��䴫�͡�
    If InStr(";21;22;23;25;", ";" & intType & ";") > 0 Then
        If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
            mobjBase.mobjLog.Add "�����͡���ǳɹ�", 1
        Else
            mobjBase.mobjLog.Add "�����͡����ʧ��", 1
        End If
    End If
    
    '��׼���ݴ������
    mobjBase.mobjLog.Add "��׼���ݴ������", 1
    
    mobjBase.mobjLog.Save
    
    'ִ�л���Ĵ���
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add "��׼���ݴ���ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset) As String
'���ܣ�����¼������ļ�¼ת������Ϣƽ̨����θ�ʽ
'������
'  intType��ҵ������
'  intIO��������סԺ
'  strCode���ӿڱ��
'  rsData����¼������
'���أ���׼��������

    Dim objXML As New clsXML
    Dim strIP As String, strPCode As String, strPName As String, strTmp As String
    Dim strItem As String, strNO As String
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    With rsData
    
        strIP = mdlDrugMachine.IP
        strPCode = UserInfo.���
        strPName = UserInfo.����
        
        Select Case intType
        Case 1      '����
            strTmp = "104"
        Case 2      '��Ա
            strTmp = "107"
        Case 3      'ҩƷ
            strTmp = "101"
        Case 4      '������λ
            strTmp = "102"
        Case 5      '��ҩ����
            strTmp = "106"
        Case 21     '��ҩ
            strTmp = "201"
        Case 22     '��ʼ��ҩ
            strTmp = "202"
        Case 23     '��ɷ�ҩ
            strTmp = "203"
        Case 24     '��ҩ
            strTmp = "204"
        Case 25     '������ҩ
            strTmp = "205"
        Case Else
            Exit Function
        End Select
    
        .MoveFirst
        
        objXML.AppendNode "ROOT"
        objXML.AppendData "OPSYSTEM", "HIS"
        Select Case intType
        Case 22, 23
            objXML.AppendData "OPWINID", mdlDrugMachine.NVL(!���ڱ���)
        Case Else
            objXML.AppendData "OPWINID", ""
        End Select
        
        objXML.AppendData "OPTYPE", strTmp
        objXML.AppendData "OPIP", strIP
        objXML.AppendData "OPMANNO", SpecialChar(strPCode)
        objXML.AppendData "OPMANNAME", SpecialChar(strPName)
        
        Do
            Select Case intType
            Case 1      '����
                strItem = "CONSIS_BASIC_DEPTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DEPTCODE", mdlDrugMachine.NVL(!ID)
                objXML.AppendData "DEPTNAME", SpecialChar(mdlDrugMachine.NVL(!����))
                objXML.AppendData "OUTP_OR_INP", mdlDrugMachine.NVL(!�������)
                
                objXML.AppendNode strItem, True
                
            Case 2      '��Ա
                strItem = "CONSIS_BASIC_STAFFERVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "STAFFERID", mdlDrugMachine.NVL(!ID)
                objXML.AppendData "STAFFERNO", SpecialChar(mdlDrugMachine.NVL(!���))
                objXML.AppendData "DEPTCODE", mdlDrugMachine.NVL(!����ID)
                objXML.AppendData "STAFFERNAME", SpecialChar(mdlDrugMachine.NVL(!����))
'                objXML.AppendData "JOB", ""
'                objXML.AppendData "TITLE", ""
                
                objXML.AppendNode strItem, True
            
            Case 3      'ҩƷ
                strItem = "CONSIS_BASIC_DRUGSVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
                objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!ͨ����))
                objXML.AppendData "TRADE_NAME", SpecialChar(mdlDrugMachine.NVL(!��Ʒ��))
                objXML.AppendData "ENGLISHNAME", SpecialChar(mdlDrugMachine.NVL(!Ӣ����))
                objXML.AppendData "PINYIN", mdlDrugMachine.NVL(!ͨ��������)
                
                objXML.AppendData "SORTTYPE1", mdlDrugMachine.NVL(!ҩƷ���)
                objXML.AppendData "SORTTYPE2", SpecialChar(mdlDrugMachine.NVL(!ҩƷ����))
                objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!���))
                objXML.AppendData "MINSPECS", ""
                
                objXML.AppendData "UNIT", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                objXML.AppendData "MAXUNIT", SpecialChar(mdlDrugMachine.NVL(!ҩ�ⵥλ))
                objXML.AppendData "MINUNIT", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                objXML.AppendData "DOSAGE", mdlDrugMachine.NVL(!����ϵ��)
                objXML.AppendData "DOSAGEUNIT", SpecialChar(mdlDrugMachine.NVL(!������λ))
                
                objXML.AppendData "PRICE1", "0"
                objXML.AppendData "CONVERTION1", Round(mdlDrugMachine.NVL(!ҩ���װ, 0) / mdlDrugMachine.NVL(!�����װ), 2)
                objXML.AppendData "CONVERTION2", mdlDrugMachine.NVL(!�����װ)
                objXML.AppendData "FIRM_ID", SpecialChar(mdlDrugMachine.NVL(!�����̱���))
                objXML.AppendData "FIRM_NAME", SpecialChar(mdlDrugMachine.NVL(!������))
                
                objXML.AppendData "PASSNO", SpecialChar(mdlDrugMachine.NVL(!��׼�ĺ�))
                objXML.AppendData "BARCODE", "0"
                objXML.AppendData "STORAGECONDITION", ""
                objXML.AppendData "STORAGETYPE", ""
                objXML.AppendData "ALLOWIND", IIf(mdlDrugMachine.NVL(!ͣ��) = "1", "Y", "N")
                
                objXML.AppendNode strItem, True
            
            Case 4      '������λ
                strItem = "CONSIS_BASIC_LOCATIONVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!ҩƷ����))
                objXML.AppendData "LOCATIONINFO", SpecialChar(mdlDrugMachine.NVL(!�ⷿ��λ))
                objXML.AppendData "BATCHID", mdlDrugMachine.NVL(!����)
                objXML.AppendData "BATCHNO", mdlDrugMachine.NVL(!�ϴ�����)
                
                objXML.AppendData "PRODUCEDATE", mdlDrugMachine.NVL(!�ϴ���������), xsDateTime
                objXML.AppendData "DISABLEDDATE", mdlDrugMachine.NVL(!Ч��), xsDateTime
                objXML.AppendData "DRUG_QUANTITY", mdlDrugMachine.NVL(!ʵ������)
                
                objXML.AppendNode strItem, True
                
            Case 5      '��ҩ����
                strItem = "CONSIS_BASIC_OPWINIDVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "OPWINID_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!ҩ��id)
                objXML.AppendData "OPWINID_NAME", SpecialChar(mdlDrugMachine.NVL(!����))
                                
                objXML.AppendNode strItem, True
            
            Case 21     '��ҩ
                
                If Not strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW"  '����
                    
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                    objXML.AppendData "PATIENT_ID", SpecialChar(mdlDrugMachine.NVL(!���￨��))
                    objXML.AppendData "PATIENT_NAME", SpecialChar(mdlDrugMachine.NVL(!����))
                    
                    objXML.AppendData "INVOICE_NO", ""
                    objXML.AppendData "PATIENT_TYPE", IIf(Val(mdlDrugMachine.NVL(!���ȼ�)) = 1, "01", "00")
                    objXML.AppendData "DATE_OF_BIRTH", mdlDrugMachine.NVL(!��������)
                    objXML.AppendData "SEX", SpecialChar(mdlDrugMachine.NVL(!�Ա�))
                    objXML.AppendData "PRESC_IDENTITY", SpecialChar(mdlDrugMachine.NVL(!����))
                    
                    objXML.AppendData "CHARGE_TYPE", mdlDrugMachine.NVL(!ҽ�����)
                    objXML.AppendData "PRESC_ATTR", ""
                    objXML.AppendData "PRESC_INFO", ""
                    objXML.AppendData "RCPT_INFO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "RCPT_REMARK", ""
                    
                    objXML.AppendData "REPETITION", "1"
                    objXML.AppendData "COSTS", "0"
                    objXML.AppendData "PAYMENTS", "0"
                    objXML.AppendData "ORDERED_BY", mdlDrugMachine.NVL(!��������id)
                    objXML.AppendData "ORDERED_BY_NAME", SpecialChar(mdlDrugMachine.NVL(!��������))
                    
                    objXML.AppendData "PRESCRIBED_BY", SpecialChar(mdlDrugMachine.NVL(!����ҽ��))
                    objXML.AppendData "ENTERED_BY", SpecialChar(mdlDrugMachine.NVL(!����ҽ��))
                    objXML.AppendData "DISPENSE_PRI", IIf(Val(mdlDrugMachine.NVL(!���ȼ�)) = 1, "1", "2")
                
                    GoTo makDetail
                Else
makDetail:
                    
                    strItem = "CONSIS_PRESC_DTLVW"
                    objXML.AppendNode strItem           '������ϸ
                
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "ITEM_NO", mdlDrugMachine.NVL(!���)
                    objXML.AppendData "ADVICE_CODE", mdlDrugMachine.NVL(!ҽ��id)
                    objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
                    objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!���))
                    
                    objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!ͨ����))
                    objXML.AppendData "FIRM_ID", ""
                    objXML.AppendData "FIRM_NAME", SpecialChar(mdlDrugMachine.NVL(!������))
                    objXML.AppendData "PACKAGE_SPEC", SpecialChar(mdlDrugMachine.NVL(!���))
                    objXML.AppendData "PACKAGE_UNITS", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                    
                    objXML.AppendData "QUANTITY", mdlDrugMachine.NVL(!���ﵥλ����)
                    objXML.AppendData "UNIT", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                    objXML.AppendData "COSTS", "0"
                    objXML.AppendData "PAYMENTS", "0"
                    objXML.AppendData "DOSAGE", SpecialChar(mdlDrugMachine.NVL(!��������))
                    
                    objXML.AppendData "DOSAGE_UNITS", SpecialChar(mdlDrugMachine.NVL(!������λ))
                    objXML.AppendData "ADMINISTRATION", SpecialChar(mdlDrugMachine.NVL(!�÷�))
                    objXML.AppendData "FREQUENCY", SpecialChar(mdlDrugMachine.NVL(!ִ��Ƶ��))
                    objXML.AppendData "ADDITIONUSAGE", ""
                    objXML.AppendData "RCPT_REMARK", ""
                    
                    objXML.AppendNode strItem, True
                End If
                
                strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                
            Case 22, 23    '22��ʼ��ҩ��23-��ɷ�ҩ
                strItem = "CONSIS_PRESC_MSTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                objXML.AppendData "INVOICE_NO", ""
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                
                objXML.AppendNode strItem, True
            
            Case 24     '��ҩ
                If Not strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW"  '����
                    
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "INVOICE_NO", ""
                    objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                
                    GoTo makReturn
                Else
makReturn:
                    strItem = "CONSIS_PRESC_DTLVW"
                    objXML.AppendNode strItem           '������ϸ
                    
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "ADVICE_CODE", mdlDrugMachine.NVL(!ҽ��id)
                    objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!ҩƷ����))
                    objXML.AppendData "QUANTITY", mdlDrugMachine.NVL(!���ﵥλ����)
                    
                    objXML.AppendNode strItem, True
                End If
                
                strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                
            Case 25     '������ҩ
                strItem = "CONSIS_PRESC_MSTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                objXML.AppendData "INVOICE_NO", ""
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                
                objXML.AppendNode strItem, True
            End Select
            
            'XML��β
            .MoveNext
            
            Select Case intType
            Case 21, 24    '21-��ҩ��24-��ҩ
                If .EOF Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW", True
                Else
                    If Not strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                        objXML.AppendNode "CONSIS_PRESC_MSTVW", True
                    End If
                End If
            End Select
            
        Loop While .EOF = False
    End With
    
    objXML.AppendNode "ROOT", True
    
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ���������", 1
    mobjBase.mobjLog.Save
    FormatData = objXML.XmlText
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ������ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description
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
    Call mdlDrugMachine.GetUserInfo(mobjBase.UserName, mobjBase.mobjComLib, mobjBase.mobjLog, UserInfo)
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property


