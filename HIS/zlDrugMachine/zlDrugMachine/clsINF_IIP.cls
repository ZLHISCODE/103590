VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_IIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ�������Ϣ����ƽ̨����
'IIP��Information Interchange Platform
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjHTTP As Object

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'˵���������ӿ�Ϊ�״̬

    mobjBase.Action = RHS
    
    If RHS Then
        If mobjHTTP Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "����WinHTTP����ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ������Ϣ����ƽ̨", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        
        On Error Resume Next
        Call mobjHTTP.Open("POST", mobjBase.URL)
        If Err.Number = -2147012891 Then
            mobjBase.mobjLog.Add mobjHTTP.ResponseText, 1, 1
            mobjBase.mobjLog.Add "������Ϣ����ƽ̨�쳣", 1
        Else
            mobjBase.mobjLog.Add "�ɹ�������Ϣ����ƽ̨", 1
        End If
        On Error GoTo 0
    Else
        Set mobjHTTP = Nothing
    End If
    
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Public Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����Filter����

    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)

End Function

Public Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, ByRef blnFinish As Boolean) As ADODB.Recordset
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

    Dim arrFilter As Variant
    Dim lngPatientID As Long, lngStockID As Long
    Dim intBill As Integer
    Dim strNO As String
    Dim i As Long
    Dim rsNew As New ADODB.Recordset
    
    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "ͣ��=0"
    Case 21 To 25
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

    If intType = 21 And intIO = 1 Then
        '��Ҫ���ش�����Ϣ���������ˡ��ⷿ���ֱ���
        
        arrFilter = Array()
        
        'ȡ������ⷿ
        Do While rsData.EOF = False
            If Not (lngPatientID = rsData!����id And lngStockID = rsData!�ⷿid) Then
                ReDim Preserve arrFilter(UBound(arrFilter) + 1)
                arrFilter(UBound(arrFilter)) = "����id=" & rsData!����id & " And �ⷿid=" & rsData!�ⷿid & " And �Ƿ���=1"
            End If
            lngPatientID = rsData!����id
            lngStockID = rsData!�ⷿid
            
            rsData.MoveNext
        Loop
        
        '��������
        For i = LBound(arrFilter) To UBound(arrFilter)
            rsData.Filter = arrFilter(i)
            Call SendDataEx(intAppType, intType, intIO, strCode, rsData, strMessage)
        Next
    
    ElseIf intType >= 22 And intType <= 24 And intIO = 1 Then
        '22-��ʼ��ҩ��23-��ɷ�ҩ��24-������ҩ���������ݡ������š��ֱ���
        
        'ȡ�����봦����
        Do While rsData.EOF = False
            If Not (strNO = Trim(rsData!������) And intBill = rsData!����) Then
                ReDim Preserve arrFilter(UBound(arrFilter) + 1)
                arrFilter(UBound(arrFilter)) = "������='" & Trim(rsData!������) & "' And ����=" & intBill & " And �Ƿ���=1"
            End If
            strNO = Trim(rsData!������)
            intBill = rsData!����
            
            rsData.MoveNext
        Loop
        
        '��������
        For i = LBound(arrFilter) To UBound(arrFilter)
            rsData.Filter = arrFilter(i)
            Call SendDataEx(intAppType, intType, intIO, strCode, rsData, strMessage)
        Next
        
    Else
        '����¼�д���
        
        Set rsNew = mdlDrugMachine.CopyStructure(rsData.Fields)
        If rsNew.State = adStateOpen Then
            i = 1
            Do While rsData.EOF = False
                '���Ƽ�¼
                strMessage = CopyRecord(rsData.Fields, rsNew)
                If strMessage = "" Then
                    '��50�д���1��
                    If i Mod 50 = 0 Then
                        '��������
                        Call SendDataEx(intAppType, intType, intIO, strCode, rsNew, strMessage)
                        '�������
                        If rsNew.RecordCount > 0 Then Call ClearRecord(rsNew)
                    End If
                    i = i + 1
                Else
                    mobjBase.mobjLog.Add "������¼ʧ�ܡ�" & strMessage, 1
                End If
                
                rsData.MoveNext
            Loop
            If rsNew.State = adStateOpen Then
                If rsNew.RecordCount > 0 Then
                    Call SendDataEx(intAppType, intType, intIO, strCode, rsNew, strMessage)
                    Call ClearRecord(rsNew)
                End If
                rsNew.Close
            End If
        Else
            mobjBase.mobjLog.Add "������¼������ʧ��", 1
        End If
    End If
    
    mobjBase.mobjLog.Add "�����������", 1
    mobjBase.mobjLog.Save
    
    'ִ�л���Ĵ���
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
hErr:
    '���ʧ�ܣ�ҩƷ�շ���¼��
    strMessage = Err.Number & "��" & Err.Description
    mobjBase.mobjLog.Add strMessage, 1
    mobjBase.mobjLog.Add "��������ʧ��", 1
    mobjBase.mobjLog.Save
End Function

Private Sub SendDataEx(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String)
    
    Const WinHttpRequestOption_SslErrorIgnoreFlags As Long = 4
    
    Dim strInput As String, strOutput As String, strVal As String
    Dim strCData As String, strWindows As String, strTmp As String
    Dim strErrCode As String, strErrMsg As String
    Dim strFuncURL As String, strService As String
    Dim blnFinish As Boolean
    Dim objXML As New clsXML, objData As New clsXML
    Dim objDB As Object, objEncrypt As Object
    
    On Error GoTo hErr
    
    '��ʽ������
    strTmp = "��ʽ������"
    mobjBase.mobjLog.Add "��ʼ" & strTmp, 1
    strInput = FormatData(intType, intIO, strCode, rsData)
    If strInput = "" Then
        strMessage = "��ʽ������ʧ��"
        Exit Sub
    End If
    
    '��׼���ݴ��Ϳ�ʼ
    strTmp = "��׼���ݴ���"
    mobjBase.mobjLog.Add "��ʼ" & strTmp, 1
    
    'ƽ̨�ӿ�
    strService = "Drug/Machine"
    If Right(mobjBase.URL, 1) = "/" Then
        strFuncURL = mobjBase.URL & strService
    Else
        strFuncURL = mobjBase.URL & "/" & strService
    End If
    
    '��������
    On Error Resume Next
    Call mobjHTTP.Open("POST", strFuncURL, False)
'    mobjHTTP.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300     '���Դ���
    If Err.Number <> -2147012891 Then
        If Err.Number = 0 Then
            On Error GoTo hErr
        
            mobjHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'mobjHTTP.SetRequestHeader "Content-Length", LenB(StrConv(strInput, vbFromUnicode))
            mobjHTTP.Send "redata=" & strInput
            strOutput = mobjHTTP.ResponseText
            mobjBase.mobjLog.Add strOutput, 1, 1
            mobjBase.mobjLog.Add "�������ݵ���Ϣ����ƽ̨���", 1
        Else
            mobjBase.mobjLog.Add Err.Description, 1, 1
            mobjBase.mobjLog.Add "WinHttpRequest.Openʧ��", 1
        End If
    Else
        mobjBase.mobjLog.Add "WinHTTP����ʧ�ܣ��������ݵ���Ϣ����ƽ̨���ɹ�", 1
    End If
    On Error GoTo hErr
    
'<ROOT>
'   <STATE><![CDATA[״̬]]></STATE>
'   <DATAPARAM><![CDATA[ҵ������]]></DATAPARAM>
'   <ERROR>
'       <ERRCODE><![CDATA[������]]></ERRCODE>
'       <MSG><![CDATA[������Ϣ]]></MSG>
'   </ERROR>
'</ROOT>
    
    '����������Ϣ
    Call objXML.OpenXMLDocument(strOutput)
    Call objXML.GetSingleNodeValue("STATE", strVal)
    
    If UCase(strVal) = "T" Then
        '�ɹ�
        Call objXML.GetSingleNodeValue("DATAPARAM", strCData)
        If intType = Val("21-��ҩ") And intIO = Val("1-����") Then
            On Error Resume Next
            Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
            If Err.Number <> 0 Then
                mobjBase.mobjLog.Add "zlEncryptPub����δע�ᣬ�ܾ���ʽ������", 1
                Err.Clear
                On Error GoTo hErr
            Else
                On Error GoTo hErr
                strCData = objEncrypt.AES.DecryptStr(strCData, objEncrypt.Base64Decode(mobjBase.KEY))
                Call objData.OpenXMLDocument(strCData)
                Call objData.GetSingleNodeValue("FYCK", strWindows)       '��ҩ����
                objData.CloseXMLDocument
            End If
        End If
        blnFinish = True
        
    ElseIf UCase(strVal) = "F" Then
        'ʧ��
        Call objXML.GetSingleNodeValue("ERRCODE", strErrCode)
        Call objXML.GetSingleNodeValue("MSG", strErrMsg)
        
        strMessage = "[" & strErrCode & "]" & strErrMsg
        
    Else
        '������Ϣ�쳣
        strMessage = "������Ϣ�쳣"
    End If
    
    objXML.CloseXMLDocument
    
    If blnFinish Then
        '�ɹ�
        If intType = Val("21-��ҩ") And intIO = Val("1-����") Then
            '��д��ҩ����
            If strWindows <> "" Then
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
        End If
    ElseIf Val(strVal) = Val("2-��Ϣ����ƽ̨�쳣�����������쳣") Then
        mobjBase.mobjLog.Add strMessage, 1
        mobjBase.mobjLog.Add "��Ϣ����ƽ̨�쳣", 1
    Else
        mobjBase.mobjLog.Add strMessage, 1
        mobjBase.mobjLog.Add "�������ӿ��쳣", 1
    End If
    
    '��׼���ݴ������
    mobjBase.mobjLog.Add "���" & strTmp, 1
    
    '��ǳɹ���ҩƷ�շ���¼����֮����ʧ��Ҳ��ǣ���Ϊ���ýӿڹ������ߵĶ�ʱ���Ϳɲ��䴫�͡�
    If intType > 20 Then
        If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
            mobjBase.mobjLog.Add "�����͡���ǳɹ�", 1
        Else
            mobjBase.mobjLog.Add "�����͡����ʧ��", 1
        End If
    End If
    mobjBase.mobjLog.Save
    
    Exit Sub
    
hErr:
    strMessage = Err.Number & ":" & Err.Description
    mobjBase.mobjLog.Add strMessage, 1
    mobjBase.mobjLog.Add strTmp & "ʧ��", 1
    mobjBase.mobjLog.Save
End Sub

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset) As String
'���ܣ�����¼������ļ�¼ת������Ϣƽ̨����θ�ʽ
'������
'  intType��ҵ������
'  intIO��������סԺ
'  strCode���ӿڱ��
'  rsData����¼������
'���أ���׼��������

    Dim strTmp As String, strNO As String, strService As String
    Dim l As Long, lngDeptID As Long, lngPatientID As Long, lngDeptTakeID As Long
    Dim arrTmp As Variant
    Dim objEncrypt As Object
    Dim objXML As New clsXML, objData As New clsXML

    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error Resume Next
    Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        mobjBase.mobjLog.Add "zlEncryptPub����δע�ᣬ�ܾ���ʽ������", 1
        mobjBase.mobjLog.Save
        Exit Function
    End If
    On Error GoTo hErr
    
    lngPatientID = -9999        '��¼���ġ�����ID���ֶ�ֵ����ΪNull����ֹ��֯XML����ȷ
    
    '��׼������
    strService = "NotSupport"

    Select Case intType
    Case 1      '����
        strService = "Drug.Machine.Base.Dept"
    Case 2      '��Ա
        strService = "Drug.Machine.Base.Worker"
    Case 3      'ҩƷ
        strService = "Drug.Machine.Base.Drug"
    Case 4      '���
        strService = "Drug.Machine.Base.Stock"
    Case 5      '����
        strService = "Drug.Machine.Base.Window"
    Case 21     '��ҩ
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.Dosage"
        Else
            strService = "Drug.Machine.Inpatient.Dosage"
        End If
    Case 22     '��ʼ��ҩ
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.SendStart"
        End If
    Case 23     '��ɷ�ҩ
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.SendComplete"
        End If
    Case 24     '��ҩ
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.Back"
        End If
    End Select
    
'<ROOT>
'    <TOKEN><![CDATA[����]]></TOKEN>
'    <SERVICE><![CDATA[��������]]></SERVICE>
'    <DATAPARAM>
'       <![CDATA[ҵ������]]>
'    </DATAPARAM>
'</ROOT>
    
    objXML.AppendNode "ROOT"
    strTmp = "<![CDATA[" & objEncrypt.AES.EncryptStr(objEncrypt.Base64Decode(mobjBase.Token), objEncrypt.Base64Decode(mobjBase.KEY)) & "]]>"
    objXML.AppendData "TOKEN", strTmp, xsString, True
    objXML.AppendData "SERVICE", "<![CDATA[" & strService & "]]>", xsString, True
    
    objData.AppendData "JKBH", strCode
    objData.AppendNode "ITEMS"
    
    With rsData
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            
            Select Case intType
            Case 1 To 5
                objData.AppendNode "ITEM"
            End Select
            
            Select Case intType
            Case 1      '������Ϣ
                
                objData.AppendData "ID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "BM", mdlDrugMachine.NVL(!����)
                objData.AppendData "MC", mdlDrugMachine.NVL(!����)
                objData.AppendData "FWDX", mdlDrugMachine.NVL(!�������)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!ͣ��)
                
                arrTmp = Split(mdlDrugMachine.NVL(!��������), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "BMXZ"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "XZ", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "BMXZ", True
                    
            Case 2      '��Ա��Ϣ
            
                objData.AppendData "ID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "BH", mdlDrugMachine.NVL(!���)
                objData.AppendData "XM", mdlDrugMachine.NVL(!����)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!ͣ��)
                
                '��Ա����
                arrTmp = Split(mdlDrugMachine.NVL(!��Ա����), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "RYXZ"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "XZ", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "RYXZ", True
                
                '��������
                arrTmp = Split(mdlDrugMachine.NVL(!����ID), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "SSBM"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "BMID", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "SSBM", True
            
            Case 3      'ҩƷĿ¼
            
                objData.AppendData "YPID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "YPBM", mdlDrugMachine.NVL(!����)
                objData.AppendData "YPGG", mdlDrugMachine.NVL(!���)
                objData.AppendData "YPMC", mdlDrugMachine.NVL(!ͨ����)
                objData.AppendData "SPMC", mdlDrugMachine.NVL(!��Ʒ��)
                
                objData.AppendData "YWMC", mdlDrugMachine.NVL(!Ӣ����)
                objData.AppendData "PYJM", mdlDrugMachine.NVL(!ͨ��������)
                objData.AppendData "YPLB", mdlDrugMachine.NVL(!ҩƷ���)
                objData.AppendData "YPJX", mdlDrugMachine.NVL(!ҩƷ����)
                objData.AppendData "YKDW", mdlDrugMachine.NVL(!ҩ�ⵥλ)
                
                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!סԺ��λ)
                objData.AppendData "MZDW", mdlDrugMachine.NVL(!���ﵥλ)
                objData.AppendData "ZXDW", mdlDrugMachine.NVL(!���㵥λ)
                objData.AppendData "JLXS", mdlDrugMachine.NVL(!����ϵ��)
                objData.AppendData "JLDW", mdlDrugMachine.NVL(!������λ)
                
                objData.AppendData "YKBZ", mdlDrugMachine.NVL(!ҩ���װ)
                objData.AppendData "ZYBZ", mdlDrugMachine.NVL(!סԺ��װ)
                objData.AppendData "MZBZ", mdlDrugMachine.NVL(!�����װ)
                objData.AppendData "DLFL", mdlDrugMachine.NVL(!��������)
                objData.AppendData "PZWH", mdlDrugMachine.NVL(!��׼�ĺ�)
                
                objData.AppendData "SCSDM", mdlDrugMachine.NVL(!�����̱���)
                objData.AppendData "SCS", mdlDrugMachine.NVL(!������)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!ͣ��)
            
            Case 4      'ҩƷ������λ
            
                objData.AppendData "KFID", mdlDrugMachine.NVL(!�ⷿid)
                objData.AppendData "YPID", mdlDrugMachine.NVL(!ҩƷid)
                objData.AppendData "YPBM", mdlDrugMachine.NVL(!ҩƷ����)
                objData.AppendData "YPWZ", mdlDrugMachine.NVL(!�ⷿ��λ)
                objData.AppendData "PC", mdlDrugMachine.NVL(!����)
                
                objData.AppendData "PH", mdlDrugMachine.NVL(!�ϴ�����)
                objData.AppendData "SCRQ", mdlDrugMachine.NVL(!�ϴ���������)
                objData.AppendData "SXRQ", mdlDrugMachine.NVL(!Ч��)
                objData.AppendData "YKDW", mdlDrugMachine.NVL(!ҩ�ⵥλ)
                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!סԺ��λ)
                
                objData.AppendData "MZDW", mdlDrugMachine.NVL(!���ﵥλ)
                objData.AppendData "SJDW", mdlDrugMachine.NVL(!���㵥λ)
                objData.AppendData "YKBZ", mdlDrugMachine.NVL(!ҩ���װ)
                objData.AppendData "ZYBZ", mdlDrugMachine.NVL(!סԺ��װ)
                objData.AppendData "MZBZ", mdlDrugMachine.NVL(!�����װ)
                
                objData.AppendData "SJDWSL", mdlDrugMachine.NVL(!ʵ������)
                objData.AppendData "SCSBM", mdlDrugMachine.NVL(!�����̱���)
                objData.AppendData "SCS", mdlDrugMachine.NVL(!������)
                objData.AppendData "GYSBM", mdlDrugMachine.NVL(!��Ӧ�̱���)
                objData.AppendData "GYS", mdlDrugMachine.NVL(!��Ӧ��)
                
            Case 5      '��ҩ����
            
                objData.AppendData "KFID", mdlDrugMachine.NVL(!ҩ��id)
                objData.AppendData "CKBM", mdlDrugMachine.NVL(!����)
                objData.AppendData "CKMC", mdlDrugMachine.NVL(!����)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!�ϰ�)
            
            Case 21     '��ҩ
            
                If intIO = 1 Then
                    '�����䣨�ڣ�ҩ
                    
                    If lngPatientID <> Val(mdlDrugMachine.NVL(!����id)) Then
                        '����
                        objData.AppendNode "BRXX"
                        
                        objData.AppendData "ID", mdlDrugMachine.NVL(!����id)
                        objData.AppendData "XM", mdlDrugMachine.NVL(!����)
                        objData.AppendData "XB", mdlDrugMachine.NVL(!�Ա�)
                        objData.AppendData "NL", mdlDrugMachine.NVL(!����)
                        objData.AppendData "JZKH", mdlDrugMachine.NVL(!���￨��)
                        
                        objData.AppendData "YBLB", mdlDrugMachine.NVL(!ҽ�����)
                        objData.AppendData "FB", mdlDrugMachine.NVL(!�ѱ�)
                        objData.AppendData "YXJ", mdlDrugMachine.NVL(!���ȼ�)
                        
                        objData.AppendNode "BRXX", True
                        
                        GoTo makBill
                    Else
makBill:
                        '����
                        If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                            objData.AppendNode "DJ"
                            
                            objData.AppendData "CFSJ", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                            objData.AppendData "CFH", mdlDrugMachine.NVL(!������)
                            objData.AppendData "DJLX", mdlDrugMachine.NVL(!����)
                            objData.AppendData "FYYFID", Val(mdlDrugMachine.NVL(!�ⷿid))
                            objData.AppendData "FYYF", mdlDrugMachine.NVL(!��ҩҩ��)
                            
                            objData.AppendData "CFZD", mdlDrugMachine.NVL(!�������)
                            objData.AppendData "CFBZ", ""
                            objData.AppendData "JFSJ", mdlDrugMachine.NVL(!�ɷ�ʱ��), xsDateTime
                            objData.AppendData "KFKSID", Val(mdlDrugMachine.NVL(!��������id))
                            objData.AppendData "KFKS", mdlDrugMachine.NVL(!��������)
                            
                            objData.AppendData "KFYS", mdlDrugMachine.NVL(!����ҽ��)
                            
                            GoTo makDetail
                        Else
makDetail:
                            '������ϸ
                            objData.AppendNode "MX"
                            
                            objData.AppendData "XH", Val(mdlDrugMachine.NVL(!���))
                            objData.AppendData "YZID", Val(mdlDrugMachine.NVL(!ҽ��id))
                            objData.AppendData "YPID", Val(mdlDrugMachine.NVL(!ҩƷid))
                            objData.AppendData "YPBM", mdlDrugMachine.NVL(!����)
                            objData.AppendData "YPGG", mdlDrugMachine.NVL(!���)
                            
                            objData.AppendData "YPMC", mdlDrugMachine.NVL(!ͨ����)
                            objData.AppendData "SPMC", mdlDrugMachine.NVL(!��Ʒ��)
                            objData.AppendData "SCS", mdlDrugMachine.NVL(!������)
                            objData.AppendData "MZDWSL", Val(mdlDrugMachine.NVL(!���ﵥλ����))
                            objData.AppendData "MZDW", mdlDrugMachine.NVL(!���ﵥλ)
                            
                            objData.AppendData "YF", mdlDrugMachine.NVL(!�÷�)
                            objData.AppendData "DL", mdlDrugMachine.NVL(!��������)
                            objData.AppendData "DLDW", mdlDrugMachine.NVL(!������λ)
                            objData.AppendData "ZXPC", mdlDrugMachine.NVL(!ִ��Ƶ��)
                            objData.AppendData "YPJX", mdlDrugMachine.NVL(!ҩƷ����)
                            
                            objData.AppendData "DLFL", mdlDrugMachine.NVL(!��������)
                            objData.AppendData "YSZT", mdlDrugMachine.NVL(!ҽ������)
                            
                            objData.AppendNode "MX", True
                            
                        End If
                        strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                            
                    End If
                    lngPatientID = Val(mdlDrugMachine.NVL(!����id))
                
                ElseIf intIO = 2 Then
                    'סԺ�䣨�ڣ�ҩ
                    
                    '��ҩҩ��
                    If lngDeptID <> Val(mdlDrugMachine.NVL(!�ⷿid)) Then
                        objData.AppendNode "FYYF"
                        
                        objData.AppendData "FYYFID", mdlDrugMachine.NVL(!��ҩҩ��id)
                        objData.AppendData "FYYF", mdlDrugMachine.NVL(!��ҩҩ��)
                        objData.AppendData "FYH", mdlDrugMachine.NVL(!��ҩ��)
                        objData.AppendData "CSSJ", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                        
                        GoTo makDeptTake
                    Else
makDeptTake:
                        '��ҩ����
                        If lngDeptTakeID <> Val(mdlDrugMachine.NVL(!��ҩ����ID)) Then
                            objData.AppendNode "LYKS"
                            
                            objData.AppendData "LYKSID", mdlDrugMachine.NVL(!��ҩ����ID)
                            objData.AppendData "LYKSMC", mdlDrugMachine.NVL(!��ҩ����)
                            
                            '����
                            If lngPatientID <> Val(mdlDrugMachine.NVL(!����id)) Then
                                objData.AppendNode "BR"
                                
                                objData.AppendData "BRID", mdlDrugMachine.NVL(!����id)
                                objData.AppendData "BRXM", mdlDrugMachine.NVL(!��������)
                                objData.AppendData "BRXB", mdlDrugMachine.NVL(!�����Ա�)
                                objData.AppendData "BRNL", mdlDrugMachine.NVL(!��������)
                                objData.AppendData "JZKH", mdlDrugMachine.NVL(!���￨��)
                                
                                objData.AppendData "BQID", mdlDrugMachine.NVL(!����id)
                                objData.AppendData "BQ", mdlDrugMachine.NVL(!����)
                                objData.AppendData "CH", mdlDrugMachine.NVL(!����)
                                objData.AppendData "YXJ", mdlDrugMachine.NVL(!���ȼ�)
                                
                                GoTo makMedical
                            Else
makMedical:
                                'ҽ��
                                objData.AppendNode "MX"
                                
                                objData.AppendData "KZYS", mdlDrugMachine.NVL(!����ҽ��)
                                objData.AppendData "KZSJ", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                                objData.AppendData "YZID", mdlDrugMachine.NVL(!����id)
                                objData.AppendData "YPJX", mdlDrugMachine.NVL(!ҩƷ����)
                                objData.AppendData "YPID", mdlDrugMachine.NVL(!ҩƷid)
                                
                                objData.AppendData "YPMC", mdlDrugMachine.NVL(!ҩƷ����)
                                objData.AppendData "SPMC", mdlDrugMachine.NVL(!��Ʒ����)
                                objData.AppendData "YF", mdlDrugMachine.NVL(!�÷�)
                                objData.AppendData "DL", mdlDrugMachine.NVL(!����)
                                objData.AppendData "DLDW", mdlDrugMachine.NVL(!������λ)
                                
                                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!סԺ��λ)
                                objData.AppendData "ZYDWSL", mdlDrugMachine.NVL(!סԺ��λ����)
                                objData.AppendData "YZQX", mdlDrugMachine.NVL(!ҽ����Ч)
                                objData.AppendData "ZXSJ", mdlDrugMachine.NVL(!ִ��ʱ��), xsDateTime
                                objData.AppendData "YSZT", mdlDrugMachine.NVL(!ҽ������)
                                
                                objData.AppendData "DLFL", mdlDrugMachine.NVL(!��������)
                                objData.AppendData "HSJDSJ", mdlDrugMachine.NVL(!��ʿУ��ʱ��), xsDateTime
                                objData.AppendData "SFTZR", mdlDrugMachine.NVL(!�շ�������)
                                objData.AppendData "SFSHR", mdlDrugMachine.NVL(!�շ������)
                                
                                objData.AppendNode "MX", True
                                
                            End If
                            lngPatientID = Val(mdlDrugMachine.NVL(!����id))
                            
                        End If
                        lngDeptTakeID = Val(mdlDrugMachine.NVL(!��ҩ����ID))
                        
                    End If
                    lngDeptID = Val(mdlDrugMachine.NVL(!�ⷿid))
                    
                End If
            
            Case 22, 23     '22-��ʼ��ҩ��23-��ɷ�ҩ
                
                If intIO = 1 Then
                    objData.AppendNode "DJ"
                    
                    objData.AppendData "CFSJ", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                    objData.AppendData "CFH", mdlDrugMachine.NVL(!������)
                    objData.AppendData "DJLX", mdlDrugMachine.NVL(!����)
                    objData.AppendData "FYYFID", Val(mdlDrugMachine.NVL(!�ⷿid))
                    objData.AppendData "BRID", Val(mdlDrugMachine.NVL(!����id))
                    objData.AppendData "YPJX", mdlDrugMachine.NVL(!ҩƷ����)
                    
                    objData.AppendNode "DJ", True
                End If
            
            Case 24     '��ҩ
            
                If intIO = 1 Then
                    
                    If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                        objData.AppendNode "DJ"
                        
                        objData.AppendData "CFSJ", mdlDrugMachine.NVL(!����ʱ��), xsDateTime
                        objData.AppendData "CFH", mdlDrugMachine.NVL(!������)
                        objData.AppendData "DJLX", mdlDrugMachine.NVL(!����)
                        objData.AppendData "FYYFID", mdlDrugMachine.NVL(!�ⷿid)
                        objData.AppendData "BRID", Val(mdlDrugMachine.NVL(!����id))
                        
                        GoTo makReturnDetail
                    Else
makReturnDetail:
                        objData.AppendNode "YPMX"
                        
                        objData.AppendData "YPJX", mdlDrugMachine.NVL(!ҩƷ����)
                        objData.AppendData "YZID", Val(mdlDrugMachine.NVL(!ҽ��id))
                        objData.AppendData "YPID", Val(mdlDrugMachine.NVL(!ҩƷid))
                        objData.AppendData "YPBM", mdlDrugMachine.NVL(!ҩƷ����)
                        objData.AppendData "MZDWSL", mdlDrugMachine.NVL(!���ﵥλ����)
                        
                        objData.AppendData "MZDW", mdlDrugMachine.NVL(!���ﵥλ)
                        objData.AppendData "DLFL", mdlDrugMachine.NVL(!��������)
                        
                        objData.AppendNode "YPMX", True
                    End If
                    
                    strNO = mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid)
                    
                End If
            
            End Select
            
            Select Case intType
            Case 1 To 5
                objData.AppendNode "ITEM", True
            End Select
            
            .MoveNext
            
            'XML��β

            Select Case intType
            Case 21, 24    '21-��ҩ��24-��ҩ
                If intIO = 1 Then
                    '����
                    If .EOF Then
                        objData.AppendNode "DJ", True
                    Else
                        If strNO <> mdlDrugMachine.NVL(!������) & "_" & mdlDrugMachine.NVL(!����) & "_" & mdlDrugMachine.NVL(!�ⷿid) Then
                            objData.AppendNode "DJ", True
                        End If
                    End If
                    
                ElseIf intIO = 2 Then
                    'סԺ
                    If .EOF Then
                        objData.AppendNode "BR", True
                        objData.AppendNode "LYKS", True
                        objData.AppendNode "FYYF", True
                    Else
                        If lngPatientID <> Val(!����id) Then objData.AppendNode "BR", True
                        If lngDeptTakeID <> Val(!��ҩ����ID) Then objData.AppendNode "LYKS", True
                        If lngDeptID <> Val(!�ⷿid) Then objData.AppendNode "FYYF", True
                    End If
                End If
            End Select
            
        Loop
    End With
    
    objData.AppendNode "ITEMS", True
    
    '����
    strTmp = objData.XmlText
    mobjBase.mobjLog.Add Replace(strTmp, vbCrLf, ""), 1, 1
    objData.ClearXmlText
    
    If Not objEncrypt Is Nothing Then
        strTmp = objEncrypt.AES.EncryptStr(strTmp, objEncrypt.Base64Decode(mobjBase.KEY))
        strTmp = "<![CDATA[" & strTmp & "]]>"
        mobjBase.mobjLog.Add "���ܳɹ�", 1
    Else
        strTmp = "Encrypt.Exception"
        mobjBase.mobjLog.Add "����ʧ�ܣ�zlEncryptPub����δע�ᣩ", 1
    End If
    
    objXML.AppendData "DATAPARAM", strTmp, xsString, True
    objXML.AppendNode "ROOT", True
    
    strTmp = objXML.XmlText
    objXML.ClearXmlText
    
    mobjBase.mobjLog.Add "��ʽ���������", 1
    mobjBase.mobjLog.Save
    FormatData = strTmp
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objData.XmlText, 1, 1
    mobjBase.mobjLog.Add "��ʽ������ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    mobjBase.Token = RHS
End Property

Private Property Get clsINF_Token() As String
    clsINF_Token = mobjBase.Token
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
        
    Call CreateHTTP(mobjHTTP, mobjBase)
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

Private Function GetXML_CDATA(ByVal strVar As String) As String
    Const STR_CDATA_S As String = "<![CDATA["
    Const STR_CDATA_E As String = "]]>"
    
    If InStr(UCase(strVar), STR_CDATA_S) > 0 Then
        strVar = Mid(strVar, InStr(UCase(strVar), STR_CDATA_S) + Len(STR_CDATA_S))
        If InStr(strVar, STR_CDATA_E) > 0 Then
            strVar = Left(strVar, InStr(strVar, STR_CDATA_E) - Len(STR_CDATA_E))
        Else
            strVar = ""
        End If
    Else
        strVar = ""
    End If
End Function