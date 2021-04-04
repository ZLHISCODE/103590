VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_WillachHeiseOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ�Τ�ֺ��ģ��ɽӿ��� v2.2
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private Type typOption
    OptWinID As String
    OptType As Integer
End Type

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
    Dim usrOpt As typOption
    Dim intResult As Integer
    Dim strIP As String
    
    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 1, 3
        rsData.Filter = "ͣ��=0"
    Case 4
        rsData.Filter = ""
    Case 21 To 23
        rsData.Filter = "�Ƿ���=1"
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
    strInput = FormatData(intType, intIO, strCode, rsData, usrOpt)
    If strInput = "" Then
        strMessage = "�����ݿɸ�ʽ��"
        Exit Function
    End If
    
    '��׼���ݴ��Ϳ�ʼ
    mobjBase.mobjLog.Add "��׼���ݴ��Ϳ�ʼ(TransConsisData)", 1
    
    '��������
    strIP = mdlDrugMachine.IP()
    On Error GoTo makFinish
    strOutput = "" & mobjSOAP.TransConsisData(usrOpt.OptWinID, _
                                              usrOpt.OptType, _
                                              strInput, _
                                              strIP, _
                                              UserInfo.���, _
                                              UserInfo.����, _
                                              intResult, _
                                              strVal)
    On Error GoTo hErr
    mobjBase.mobjLog.Add "�ӿڷ�����Ϣ��", 2
    mobjBase.mobjLog.Add strOutput, 2
    
    blnFinish = Val(strOutput) = 1
    If blnFinish Then
        If intType = 21 Then
            '��д��ҩ����
            If intAppType = Val("3-֧����") Then
                Set objDB = mobjBase.mobjComLib
            Else
                Set objDB = mobjBase.mobjComLib.zlDatabase
            End If
            If mdlDrugMachine.UpdateDispenseWindow(rsData, CStr(intResult), objDB, mobjBase.mobjLog) Then
                mobjBase.mobjLog.Add "��д��ҩ���ڳɹ�", 1
            Else
                mobjBase.mobjLog.Add "��д��ҩ����ʧ��", 1
            End If
        End If
    Else
        strMessage = strVal
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
    If InStr(";21;22;23;", ";" & intType & ";") > 0 Then
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

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, _
    ByVal rsData As ADODB.Recordset, ByRef usrOpt As typOption) As String
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
    
        usrOpt.OptWinID = ""
        Select Case intType
        Case 1      '����
            usrOpt.OptType = "104"
        Case 3      'ҩƷ
            usrOpt.OptType = "101"
        Case 4      '������λ
            usrOpt.OptType = "102"
        Case 21     '��ҩ
            usrOpt.OptType = "201"
        Case 22     '��ʼ��ҩ
            usrOpt.OptType = "202"
            usrOpt.OptWinID = mdlDrugMachine.NVL(!���ڱ���)
        Case 23     '��ɷ�ҩ
            usrOpt.OptType = "203"
            usrOpt.OptWinID = mdlDrugMachine.NVL(!���ڱ���)
        Case Else
            Exit Function
        End Select
        
        .MoveFirst
        
        objXML.AppendNode "ROOT"
        
        Do
            Select Case intType
            Case 1      '����
                strItem = "CONSIS_BASIC_DEPTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DEPTCODE", mdlDrugMachine.NVL(!ID)
                objXML.AppendData "DEPTNAME", SpecialChar(mdlDrugMachine.NVL(!����))
                objXML.AppendData "OUTP_OR_INP", mdlDrugMachine.NVL(!�������)
                
                objXML.AppendNode strItem, True
            
            Case 3      'ҩƷ
                strItem = "CONSIS_BASIC_DRUGSVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
                objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!ͨ����))
                objXML.AppendData "TRADE_NAME", SpecialChar(mdlDrugMachine.NVL(!��Ʒ��))
                objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!���))
                objXML.AppendData "DRUG_PACKAGE", mdlDrugMachine.NVL(!�����װ)
                
                objXML.AppendData "DRUG_UNIT", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                objXML.AppendData "FIRM_ID", SpecialChar(mdlDrugMachine.NVL(!�����̱���))
                objXML.AppendData "DRUG_PRICE", ""
                objXML.AppendData "DRUG_FORM", SpecialChar(mdlDrugMachine.NVL(!ҩƷ����))
                objXML.AppendData "DRUG_SORT", SpecialChar(mdlDrugMachine.NVL(!��������))
                
                objXML.AppendData "BARCODE", ""
                objXML.AppendData "LAST_DATE", mobjBase.mobjComLib.Currentdate, xsDateTime
                objXML.AppendData "PINYIN", SpecialChar(mdlDrugMachine.NVL(!ͨ��������))
                objXML.AppendData "DRUG_CONVERTATION", Round(mdlDrugMachine.NVL(!ҩ���װ, 0) / mdlDrugMachine.NVL(!�����װ), 2)
                
                objXML.AppendNode strItem, True
            
            Case 4      '������λ
                strItem = "CONSIS_PHC_STORAGEVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!ҩƷ����))
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                objXML.AppendData "DRUG_QUANTITY", Round(mdlDrugMachine.NVL(!ʵ������, 0) / mdlDrugMachine.NVL(!�����װ), 2)
                objXML.AppendData "LOCATIONINFO", SpecialChar(mdlDrugMachine.NVL(!�ⷿ��λ))
                
                objXML.AppendNode strItem, True
                
'            Case 5      '��ҩ����
'                strItem = "CONSIS_BASIC_OPWINIDVW"
'                objXML.AppendNode strItem
'
'                objXML.AppendData "OPWINID_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
'                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!ҩ��id)
'                objXML.AppendData "OPWINID_NAME", SpecialChar(mdlDrugMachine.NVL(!����))
'
'                objXML.AppendNode strItem, True
            
            Case 21     '��ҩ
                
                If Not strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW"  '����
                    
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!�ⷿid)
                    objXML.AppendData "PATIENT_ID", SpecialChar(mdlDrugMachine.NVL(!���￨��))
                    objXML.AppendData "PATIENT_NAME", SpecialChar(mdlDrugMachine.NVL(!����))
                    
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
                    objXML.AppendData "PRESCRIBED_BY", SpecialChar(mdlDrugMachine.NVL(!����ҽ��))
                    objXML.AppendData "ENTERED_BY", SpecialChar(mdlDrugMachine.NVL(!����ҽ��))
                    
                    objXML.AppendData "DISPENSE_PRI", IIf(Val(mdlDrugMachine.NVL(!���ȼ�)) = 1, "1", "2")
                
                    GoTo makDetail
                Else
makDetail:
                    
                    strItem = "CONSIS_PRESC_DTLVW"
                    objXML.AppendNode strItem           '������ϸ
                
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��, 0), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                    objXML.AppendData "ITEM_NO", mdlDrugMachine.NVL(!���)
                    objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!����))
                    objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!ͨ����))
                    
                    objXML.AppendData "TRADE_NAME", SpecialChar(mdlDrugMachine.NVL(!��Ʒ��))
                    objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!���))
                    objXML.AppendData "DRUG_PACKAGE", SpecialChar(mdlDrugMachine.NVL(!���))
                    objXML.AppendData "DRUG_UNIT", SpecialChar(mdlDrugMachine.NVL(!���ﵥλ))
                    objXML.AppendData "FIRM_ID", SpecialChar(mdlDrugMachine.NVL(!������))
                    
                    objXML.AppendData "QUANTITY", mdlDrugMachine.NVL(!���ﵥλ����)
                    objXML.AppendData "COSTS", "0"
                    objXML.AppendData "PAYMENTS", "0"
                    objXML.AppendData "DOSAGE", SpecialChar(mdlDrugMachine.NVL(!��������))
                    objXML.AppendData "DOSAGE_UNITS", SpecialChar(mdlDrugMachine.NVL(!������λ))
                    objXML.AppendData "ADMINISTRATION", SpecialChar(mdlDrugMachine.NVL(!�÷�))
                    objXML.AppendData "FREQUENCY", SpecialChar(mdlDrugMachine.NVL(!ִ��Ƶ��))
                    
                    objXML.AppendNode strItem, True
                End If
                
                strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                
            Case 22, 23    '22��ʼ��ҩ��23-��ɷ�ҩ
                strItem = "CONSIS_PRESC_MSTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����)
                
                objXML.AppendNode strItem, True
            
            End Select
            
            'XML��β
            .MoveNext
            
            Select Case intType
            Case 21
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
