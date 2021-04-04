VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_Grani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ��Ͼ�-�����ᣩ��
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjSOAP As Object
Private mobjHTTP As Object

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mobjSOAP = Nothing
    Set mobjHTTP = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'˵���������ӿ�Ϊ�״̬

    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mobjSOAP Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "���ӹ����������ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ���ӹ����������", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        Call mobjSOAP.MSSoapInit(mobjBase.URL)
        mobjBase.mobjLog.Add "�ɹ����ӹ����������", 1
    Else
        Set mobjSOAP = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Add "���ӹ�����������쳣", 1
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

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, _
    Optional strWinInfo As String) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����SendData����

    Dim strOutput As String, strInput As String, strURL As String
    Dim blnFinish As Boolean
    Dim objDB As Object

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 3
        rsData.Filter = "ͣ��=0"
    Case 21, 23
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
    strInput = FormatData(intType, intIO, strCode, rsData)
    If strInput = "" Then
        strMessage = "��ʽ������ʧ��"
        Exit Function
    End If
    
    '��׼���ݴ��Ϳ�ʼ
    Select Case intType
    Case 3
        mobjBase.mobjLog.Add "��׼���ݴ��Ϳ�ʼ(syncDrugInfo)", 1
        On Error GoTo makFinish
        strOutput = mobjSOAP.syncDrugInfosByXml(strInput)
        On Error GoTo hErr
    Case 21
        mobjBase.mobjLog.Add "��׼���ݴ��Ϳ�ʼ(syncPrescription)", 1
        On Error GoTo makFinish
        strOutput = mobjSOAP.syncPrescriptionByXml(strInput)
        On Error GoTo hErr
    Case 23
        On Error GoTo makFinish
        Call CreateHTTP(mobjHTTP, mobjBase)
        On Error GoTo hErr
        If mobjHTTP Is Nothing Then
            Err.Number = 1
            GoTo makFinish
        Else
            On Error Resume Next
            strURL = mobjBase.URL
            Call mobjHTTP.Open("POST", strURL)
            If Err.Number = -2147012891 Then
                mobjBase.mobjLog.Add mobjHTTP.ResponseText, 1, 1
                mobjBase.mobjLog.Add "���ӷ������쳣", 1
                GoTo makFinish
            End If
            On Error GoTo hErr
            
            ''''''''''''
            mobjHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            mobjHTTP.Send "message=" & strInput & "&flag=1"
            strOutput = mobjHTTP.ResponseText
            mobjBase.mobjLog.Add strOutput, 1, 1
            mobjBase.mobjLog.Add "�������ݵ���Ϣ����ƽ̨���", 1
        End If
    End Select
    mobjBase.mobjLog.Add "�ӿڷ�����Ϣ��", 2
    mobjBase.mobjLog.Add strOutput, 2
    
    '����������Ϣ
    '???????
    blnFinish = True
    
makFinish:
    If Err.Number <> 0 Then
        blnFinish = False
        Err.Clear: On Error GoTo hErr
        mobjBase.mobjLog.Add "���ӷ������쳣���������쳣", 1
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
    
    'ִ�л���Ĵ���
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add "��׼���ݴ���ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
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
End Property

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset) As String
'���ܣ�����¼������ļ�¼ת������Ϣƽ̨����θ�ʽ
'������
'  intType��ҵ������
'  intIO��������סԺ
'  strCode���ӿڱ��
'  rsData����¼������
'���أ���׼��������

    Const STR_ENCODING As String = "?xml version=""1.0"" encoding=""UTF-8""?"

    Dim objXML As New clsXML
    Dim lngPatientID As Long
    Dim strNO As String, strTmp As String
    Dim strNOs As String
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    objXML.AppendNode STR_ENCODING
    If intType = 3 Then objXML.AppendNode "list"
    
    With rsData
        .MoveFirst
        Do
            Select Case intType
            Case Val("3-ҩƷĿ¼")
                strTmp = "org.grani.dao.pojo.DrugInfo"
                objXML.AppendNode strTmp
                
                '�����ֶ����Լ���д�������ĵ�����ע����ò�Ҫת�����۸��Է�����ȷ����ҩ�ʹ�ӡ��ǩ��Ϣ������
                'objXML.AppendData "itemcd", mdlDrugMachine.NVL(!x)
                objXML.AppendData "itemcd", mdlDrugMachine.NVL(rsData!ID)
                objXML.AppendData "antibioticlevel", 0
                objXML.AppendData "attention", ""
                objXML.AppendData "barcode", mdlDrugMachine.NVL(rsData!����)
                objXML.AppendData "createtime", Format(Now, "yyyy-mm-dd") '��
                
                objXML.AppendData "creatorid", 0 '��
                objXML.AppendData "deleteflag", 0
                objXML.AppendData "deletetime", "" '��
                If InStr(mdlDrugMachine.NVL(rsData!ҩƷ����), "��") > 0 Then
                    objXML.AppendData "dosageformcd", "2"
                Else
                    objXML.AppendData "dosageformcd", ""
                End If
                
                objXML.AppendData "dosageformname", ""
                
                objXML.AppendData "downlimit", 0
                objXML.AppendData "expire", "" '?
                objXML.AppendData "genericname", mdlDrugMachine.NVL(rsData!ͨ����)
                objXML.AppendData "height", 0 '��
                objXML.AppendData "image", 0 '��
                
                objXML.AppendData "interval", 0
                objXML.AppendData "interval1", 0
                objXML.AppendData "isminpackunit", 0
                objXML.AppendData "length", 0 '?
                objXML.AppendData "manufacturercd", ""
                
                objXML.AppendData "manufacturername", ""
                objXML.AppendData "medicinegroove", 0
                objXML.AppendData "minpackunit", mdlDrugMachine.NVL(rsData!���㵥λ)
                objXML.AppendData "modulus", 0
                objXML.AppendData "nerve1", 0
                
                objXML.AppendData "nerve2", 0
                objXML.AppendData "notesex1", ""
                objXML.AppendData "notesex2", ""
                objXML.AppendData "noxiousflg", 0
                objXML.AppendData "perioddate", ""
                
                objXML.AppendData "pharmtype", 0
                objXML.AppendData "putunit", ""
                objXML.AppendData "remark", "" '?
                objXML.AppendData "reportpharm", 0
                objXML.AppendData "searchname", ""
                
                objXML.AppendData "searchspell", mdlDrugMachine.NVL(rsData!ͨ��������)
                objXML.AppendData "serialno", 1
                objXML.AppendData "setitemflg", ""
                objXML.AppendData "singlemedicine", 0
                objXML.AppendData "skintest", 0
                
                objXML.AppendData "splitnum", mdlDrugMachine.NVL(rsData!ҩ���װ)
                objXML.AppendData "strength", mdlDrugMachine.NVL(rsData!���)
                objXML.AppendData "terminalcd", ""
                objXML.AppendData "tradename", mdlDrugMachine.NVL(rsData!ͨ����)
                objXML.AppendData "updatetime", "" '?
                
                objXML.AppendData "updatorid", "" '?
                objXML.AppendData "uplimit", 0
                objXML.AppendData "usagename", ""
                objXML.AppendData "userattrcd", ""
                objXML.AppendData "userattrname", ""
                
                objXML.AppendData "useruse1", ""
                objXML.AppendData "useruse2", ""
                objXML.AppendData "useruse3", ""
                objXML.AppendData "useruse4", ""
                objXML.AppendData "weight", 0 '?
                
                objXML.AppendData "width", 0 '?
                objXML.AppendData "width0", 0 '?
                
                
                objXML.AppendNode strTmp, True


            Case Val("21-��ҩ")
                If strNO <> mdlDrugMachine.NVL(rsData!������) & "_" & mdlDrugMachine.NVL(rsData!����) & "_" & mdlDrugMachine.NVL(rsData!�ⷿid) Then
                    '����
                    objXML.AppendNode "Prescription"
                    If NVL(rsData!���￨��) = "" Then
                        objXML.AppendData "citycardcd", mdlDrugMachine.NVL(rsData!����id)
                    Else
                        objXML.AppendData "citycardcd", mdlDrugMachine.NVL(rsData!���￨��)
                    End If
                    objXML.AppendData "doctorName", mdlDrugMachine.NVL(rsData!����ҽ��)
                    objXML.AppendData "doctorcd", ""
                    objXML.AppendData "doctordeptcd", mdlDrugMachine.NVL(rsData!�ⷿid)
                    objXML.AppendData "emergencypatientflg", mdlDrugMachine.NVL(rsData!�Ƿ���)
                    
                    objXML.AppendData "enteroper", ""
                    objXML.AppendData "entertime", Format(rsData!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                    objXML.AppendData "feeamount", ""
                    objXML.AppendData "feedatetime", Format(rsData!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                    objXML.AppendData "feeoper", ""
                    
                    objXML.AppendData "hisprescd", mdlDrugMachine.NVL(rsData!������)
                    objXML.AppendData "infusionflg", ""
                    objXML.AppendData "invoiceno", ""
                    objXML.AppendData "ledgerno", mdlDrugMachine.NVL(rsData!���)
                    objXML.AppendData "patientage", mdlDrugMachine.NVL(rsData!����)
                    
                    objXML.AppendData "patientbirthday", Format(rsData!��������, "yyyy-mm-dd hh:mm:ss")
                    objXML.AppendData "patientcardcd", mdlDrugMachine.NVL(rsData!����id)
                    objXML.AppendData "patientheight", ""
                    objXML.AppendData "patienticd", ""
                    objXML.AppendData "patientiname", mdlDrugMachine.NVL(rsData!�������)
                    
                    objXML.AppendData "patientinotes", ""
                    objXML.AppendData "patientname", mdlDrugMachine.NVL(rsData!����)
                    objXML.AppendData "patientphone", ""
                    objXML.AppendData "patientsex", mdlDrugMachine.NVL(rsData!�Ա�)
                    objXML.AppendData "patientsortcd", ""
                    
                    objXML.AppendData "patientsortname", mdlDrugMachine.NVL(rsData!����)
                    objXML.AppendData "patienttimes", ""
                    objXML.AppendData "patientweight", ""
                    objXML.AppendData "patientworkunit", ""
                    objXML.AppendData "pharmacycd", mdlDrugMachine.NVL(rsData!�ⷿid)
                    
                    objXML.AppendData "presamount", ""
                    objXML.AppendData "prescount", ""
                    objXML.AppendData "presexpiredate", ""
                    objXML.AppendData "pressort", ""
                    objXML.AppendData "prestype", ""
                    
                    objXML.AppendData "vipflg", ""
                    objXML.AppendData "windowcd", "2"
                    
                    objXML.AppendNode "drugitems"
                    
                    GoTo makDetail
                    
                Else
makDetail:
                    'ҩƷ��ϸ
                    objXML.AppendNode "Rowdrug"
                        
                    objXML.AppendData "amount", ""
                    objXML.AppendData "buyprice", ""
                    
                    objXML.AppendData "doctorAdvice", mdlDrugMachine.NVL(rsData!�÷�) & "," & mdlDrugMachine.NVL(rsData!ִ��Ƶ��) & ",��" & mdlDrugMachine.NVL(rsData!����) & "��"
                    objXML.AppendData "doseunit", mdlDrugMachine.NVL(rsData!������λ)
                    objXML.AppendData "groupcd", ""
                    
                    objXML.AppendData "groupserialno", mdlDrugMachine.NVL(rsData!���)
                    objXML.AppendData "item_cd", mdlDrugMachine.NVL(rsData!����)
                    objXML.AppendData "notes", ""
                    objXML.AppendData "outunit", mdlDrugMachine.NVL(rsData!���ﵥλ)
                    objXML.AppendData "presnum", mdlDrugMachine.NVL(rsData!���ﵥλ����)
                    
                    objXML.AppendData "saleprice", ""
                    objXML.AppendData "serialno", "1"
                    objXML.AppendData "serialnumber", mdlDrugMachine.NVL(rsData!���)
                    objXML.AppendData "skintest", ""
                    objXML.AppendData "usagecd", ""
                    
                    objXML.AppendData "usedays", ""
                    objXML.AppendData "usedose", ""
                    objXML.AppendData "usetimescd", ""
                    objXML.AppendData "veinhigh", ""
                    objXML.AppendData "veinlower", ""
                    
                    objXML.AppendData "veinsort", ""
                    objXML.AppendData "veintype", ""
                    
                    objXML.AppendNode "Rowdrug", True
                    
                End If
            
                strNO = mdlDrugMachine.NVL(rsData!������) & "_" & mdlDrugMachine.NVL(rsData!����) & "_" & mdlDrugMachine.NVL(rsData!�ⷿid)
            
            Case Val("23-��ҩ���")
                strNO = mdlDrugMachine.NVL(rsData!����id, "")
            
            End Select
        
            .MoveNext
            
            '�����β
                If .EOF Then
makEOF:
                    If intType = 3 Then
                        objXML.AppendNode "list", True
                    ElseIf intType = 21 Then
                        objXML.AppendNode "drugitems", True
                        objXML.AppendNode "Prescription", True
                    ElseIf intType = 23 And Trim(strNO) <> "" Then
                        '����ID�ظ��ж�
                        If InStr(";" & strNOs & ";", ";" & Trim(strNO) & ";") = 0 Then
                            strNOs = strNOs & ";" & Trim(strNO)
                        End If
                    End If
                Else
                    If intType = 21 Then
                        If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                            GoTo makEOF
                        End If
                    ElseIf intType = 23 Then
                        If Val(strNO) <> Val(mdlDrugMachine.NVL(!����id)) Then
                            GoTo makEOF
                        End If
                    End If
                End If
            
        Loop While .EOF = False
        
    End With
    
    If intType = 23 Then
        If Left(strNOs, 1) = ";" Then strNOs = Mid(strNOs, 2)
        FormatData = strNOs
    Else
        mobjBase.mobjLog.Add objXML.XmlText, 1, 1
        mobjBase.mobjLog.Add "��ʽ���������", 1
        mobjBase.mobjLog.Save
        FormatData = objXML.XmlText
    End If
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ������ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function