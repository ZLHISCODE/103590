VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private Const STRSPLIT As String = "���"
Private Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mstrTitle As String
Private mstrError As String
Private mstrStation As String
Private mstrStationName As String
Private mstrStationIP As String
Private mstrUser As String
Private mstrUserName As String
Private mstrHisVersion As String
Private mstrMspServerIP As String
Private mlngMspServerPort As Long
Private mstrMspUser As String
Private mstrMspUserPassword As String
Private mdocXML As Object
Private mclsMipSendMessage As clsMipSendMessage
Private mclsMipServiceData As clsMipServiceData

Public Event AfterInfo(ByVal strInfoType As String, ByVal strInfoContent As String)

'######################################################################################################################
'�ӿڹ���/����
Public Function Initialize(ByVal strUser As String, Optional ByVal strUserName As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strPara As String
    Dim varPara As Variant
    Dim rsCondition As New ADODB.Recordset
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
            
    On Error GoTo errHand
    
    mstrUser = UCase(strUser)
    mstrUserName = strUserName
        
    '��oracle���ӵ�IP��ַΪ��
    mstrStationName = AnalyseComputer
    mstrStationIP = GetIpAddress(AnalyseIP)
    
    mstrStation = mstrStationName & "(" & mstrStationIP & ")"
    
    Set mdocXML = InitXMLDoc
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsMipServiceData = New clsMipServiceData
    
    If mclsMipServiceData.OpenFile(App.Path & "\Data\zlMspPollService.db") = True Then
        'ȡ����
        Set rsCondition = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(rsCondition, "�������", "5")
        rsTmp = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            strPara = zlCommFun.NVL(rsTmp.DataSet("Content").Value)
            If strPara <> "" Then
                varPara = Split(strPara, ";")
                mstrMspUser = varPara(0)
                mstrMspUserPassword = varPara(1)
                mstrMspServerIP = varPara(2)
                mlngMspServerPort = Val(varPara(3))
            End If
        End If
        mclsMipServiceData.CloseFile
    End If
        
    Initialize = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function GetLastError() As String
    '******************************************************************************************************************
    '���ܣ��ӿ�
    '������
    '���أ�
    '******************************************************************************************************************
    GetLastError = mstrError
End Function

Public Function ConnectMip() As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿ�
    '������
    '���أ�
    '******************************************************************************************************************
'    Dim strTemp As String
    
'    strTemp = gclsBusiness.GetOptionsValue(17)
'
'    If strTemp <> "" Then
'        MsgBox "��ѯ�����Ѿ��ڡ�" & strTemp & "���������������ٴ�������"
'        Exit Function
'    End If
    
    If mclsMipSendMessage Is Nothing Then Set mclsMipSendMessage = New clsMipSendMessage
    If mclsMipSendMessage.ConnectMsp(mstrMspUser, mstrMspUserPassword, mstrMspServerIP, mlngMspServerPort) = False Then
        MsgBox "������ѯ���������ԭ��" & mclsMipSendMessage.GetLastError
        Exit Function
    Else
        Call gclsBusiness.UpdateOptionsValue(17, mstrStation)
    End If
    
    ConnectMip = True
    
End Function

Public Function DisConnectMip() As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿ�
    '������
    '���أ�
    '******************************************************************************************************************
    If Not (mdocXML Is Nothing) Then Set mdocXML = Nothing
    
    If Not (mclsMipSendMessage Is Nothing) Then
        Call mclsMipSendMessage.DisConnectMsp
        Set mclsMipSendMessage = Nothing
        Call gclsBusiness.UpdateOptionsValue(17, "")
    End If
    
    DisConnectMip = True
End Function

Public Sub RunPoll()
    '******************************************************************************************************************
    '���ܣ��ӿ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    '1.�Ƚ�����ѯ��鴥����Ϣ����
    Call PollCheckMessage
    
    '2.�ٽ����ط���Ϣ����
    Call PollRepeatMessage
    
End Sub

'######################################################################################################################
'�ڲ�����/����
Private Sub PollCheckMessage()
    '******************************************************************************************************************
    '���ܣ���ѯ�����Ҫ��鴥����ͨ����Ϣ��Ŀ
    '��������
    '���أ���
    '******************************************************************************************************************
    Dim rsCondResult As New ADODB.Recordset
    Dim rsMessage As New ADODB.Recordset
    Dim strCondSQL As String
    
    On Error GoTo errHand
    
    '1.��ȡ��Ҫ���ڼ���ͨ����Ϣ
    Set rsMessage = gclsBusiness.GetPollCommonItem
        
    '2.ѭ�����ͨ����Ϣ�Ĵ��������Ƿ����㣬���������������Ϣ���ݲ�������Ϣ
    If rsMessage.RecordCount > 0 Then
        rsMessage.MoveFirst
        Do While Not rsMessage.EOF
            
            strCondSQL = zlCommFun.NVL(rsMessage("trigger_condition").Value)
            
            '��鴥������,SQLΪ��ʱ�Զ���ʾ��������
            If strCondSQL = "" Then
                Call SendOneMessage(zlCommFun.NVL(rsMessage("id").Value), _
                                    zlCommFun.NVL(rsMessage("trigger_frequency").Value, 0), _
                                    CStr(zlCommFun.NVL(rsMessage("start_date").Value)), _
                                    CStr(zlCommFun.NVL(rsMessage("stop_date").Value)), _
                                    zlCommFun.NVL(rsMessage("item_request").Value, 0))
            '----------------------------------------------------------------------------------------------------------
            Else
                If CheckPollItemCondition(strCondSQL, rsCondResult) = True Then
                    If rsCondResult.RecordCount > 0 Then
                        rsCondResult.MoveFirst
                        Do While Not rsCondResult.EOF
                            Call SendOneMessage(zlCommFun.NVL(rsMessage("id").Value), _
                                                zlCommFun.NVL(rsMessage("trigger_frequency").Value, 0), _
                                                zlCommFun.NVL(rsMessage("start_date").Value), _
                                                zlCommFun.NVL(rsMessage("stop_date").Value), _
                                                zlCommFun.NVL(rsMessage("item_request").Value, 0), _
                                                rsCondResult)
                            rsCondResult.MoveNext
                        Loop
                    End If
                End If
            End If
            rsMessage.MoveNext
        Loop
    End If
    
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    RaiseEvent AfterInfo("����", Err.Description)
'    Resume
End Sub

Private Sub PollRepeatMessage()
    '******************************************************************************************************************
    '���ܣ���ѯ�����Ҫ�ط���ģ���ͨ����Ϣ
    '��������
    '���أ���
    '******************************************************************************************************************
    Dim rsAgainMessage As New ADODB.Recordset
    Dim strMessgeContent As String
    Dim blnSucced As Boolean
    Dim strMessageKey As String
    Dim strTemp As String
    Dim lngRepeatOrder As Long
    Dim strPara As String
    
    On Error GoTo errHand
    
    If Not (mclsMipSendMessage Is Nothing) Then
        Set rsAgainMessage = gclsBusiness.GetPollAgainMessage
        If rsAgainMessage.RecordCount > 0 Then
            rsAgainMessage.MoveFirst
            Do While Not rsAgainMessage.EOF
                
                strMessageKey = rsAgainMessage("id").Value
                lngRepeatOrder = rsAgainMessage("repeatorder").Value + 1
                
                strPara = "<root>"
                strPara = strPara & "<send_station_name>" & mstrStationName & "</send_station_name>"
                strPara = strPara & "<send_station_ip>" & mstrStationIP & "</send_station_ip>"
                strPara = strPara & "<send_program>" & App.EXEName & "</send_program>"
                strPara = strPara & "<send_mipuser>" & UCase(mstrMspUser) & "</send_mipuser>"
                
                strTemp = IIf(mstrUserName = "", UCase(mstrUser), UCase(mstrUser) & "/" & mstrUserName)
                strPara = strPara & "<send_hisuser>" & strTemp & "</send_hisuser>"
                strPara = strPara & "</root>"
                
                blnSucced = gclsBusiness.RepeatMessageContent(strMessageKey, strPara, lngRepeatOrder)
                If blnSucced Then
                    strMessgeContent = gclsBusiness.GetSendMessgeContent(strMessageKey, lngRepeatOrder)
                    blnSucced = mclsMipSendMessage.SendMessage(strMessgeContent, (zlCommFun.NVL(rsAgainMessage("item_request").Value, 0) = 1))
                    If blnSucced = True Then
                        Call gclsBusiness.SendLogEdit(strMessageKey, 1, mstrMspUser, mstrUser, mstrUserName, mstrStation)
                        RaiseEvent AfterInfo("��Ϣ", "������Ϣ�ɹ���")
                    Else
                        Call gclsBusiness.SendLogEdit(strMessageKey, 2, mstrMspUser, mstrUser, mstrUserName, mstrStation, mclsMipSendMessage.GetLastError)
                        strTemp = "������Ϣʧ�ܣ�"
                        strTemp = strTemp & vbCrLf & "SendLogID��" & strMessageKey
                        strTemp = strTemp & vbCrLf & "ԭ��" & mclsMipSendMessage.GetLastError
                        RaiseEvent AfterInfo("����", strTemp)
                    End If
                End If
                
                rsAgainMessage.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Sub

Private Sub SendOneMessage(ByVal strMessageItemKey As String, _
                                ByVal lngTriggerFrequency As Long, _
                                Optional ByVal strStartDate As String, _
                                Optional ByVal strEndDate As String, _
                                Optional ByVal bytItemRequest As Byte, _
                                Optional ByVal rsCondResult As ADODB.Recordset)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    Dim strMessageKey As String
    Dim blnSucced As Boolean
    Dim strMessageContent As String
    Dim strTemp As String
    Dim strPara As String
    
    On Error GoTo errHand
    
    '������Ϣ��¼
    blnSucced = gclsBusiness.BuildMessageRecord(strMessageKey, strMessageItemKey, lngTriggerFrequency, strStartDate, strEndDate, rsCondResult)
    If blnSucced = True Then
        '������Ϣ����
        If strMessageKey <> "" Then
                    
            strPara = "<root>"
            strPara = strPara & "<send_station_name>" & mstrStationName & "</send_station_name>"
            strPara = strPara & "<send_station_ip>" & mstrStationIP & "</send_station_ip>"
            strPara = strPara & "<send_program>" & App.EXEName & "</send_program>"
            strPara = strPara & "<send_instance>" & App.hInstance & "</send_instance>"
            strPara = strPara & "<send_mipuser>" & UCase(mstrMspUser) & "</send_mipuser>"
            
            strTemp = IIf(mstrUserName = "", UCase(mstrUser), UCase(mstrUser) & "/" & mstrUserName)
            strPara = strPara & "<send_hisuser>" & strTemp & "</send_hisuser>"
            strPara = strPara & "</root>"
        
            If gclsBusiness.BuildMessageContent(strMessageKey, strPara) = True Then
                '������Ϣ
                blnSucced = False
                If Not (mclsMipSendMessage Is Nothing) Then
                    strMessageContent = gclsBusiness.GetSendMessgeContent(strMessageKey)
                    blnSucced = mclsMipSendMessage.SendMessage(strMessageContent, (bytItemRequest = 1))
                    If blnSucced = True Then
                        Call gclsBusiness.SendLogEdit(strMessageKey, 1, mstrMspUser, mstrUser, mstrUserName, mstrStation)
                        RaiseEvent AfterInfo("��Ϣ", "������Ϣ�ɹ���")
                    Else
                        Call gclsBusiness.SendLogEdit(strMessageKey, 2, mstrMspUser, mstrUser, mstrUserName, mstrStation, mclsMipSendMessage.GetLastError)
                        strTemp = "������Ϣʧ�ܣ�"
                        strTemp = strTemp & vbCrLf & "SendLogID��" & strMessageKey
                        strTemp = strTemp & vbCrLf & "ԭ��" & mclsMipSendMessage.GetLastError
                        RaiseEvent AfterInfo("����", strTemp)
                    End If
                End If
            Else
                strTemp = "������Ϣ����ʧ�ܣ�"
                strTemp = strTemp & vbCrLf & "SendLogID��" & strMessageKey
                strTemp = strTemp & vbCrLf & "ԭ��" & mclsMipSendMessage.GetLastError
                RaiseEvent AfterInfo("����", strTemp)
            End If
        End If
    Else
        strTemp = "������Ϣ��¼ʧ�ܣ�"
        strTemp = strTemp & vbCrLf & "��Ϣ��" & strMessageItemKey
        strTemp = strTemp & vbCrLf & "ԭ��" & mclsMipSendMessage.GetLastError
        RaiseEvent AfterInfo("����", strTemp)
    End If
    
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    RaiseEvent AfterInfo("����", Err.Description)
End Sub

Private Function RepairMessage(ByVal strMsgContent As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    Dim strQueueName As String
    Dim strRequestResult As String
    Dim strXml As String
    Dim objNode As Object
    
    On Error GoTo errHand
        
    RepairMessage = strMsgContent
    
                    
    If mdocXML.loadXML(strMsgContent) Then

        Set objNode = mdocXML.selectSingleNode(".//head/send_station_name")
        objNode.Text = mstrStationName
                
        Set objNode = mdocXML.selectSingleNode(".//head/send_station_ip")
        objNode.Text = mstrStationIP
        
        Set objNode = mdocXML.selectSingleNode(".//head/send_program")
        objNode.Text = App.EXEName
        
        Set objNode = mdocXML.selectSingleNode(".//head/send_mipuser")
        objNode.Text = UCase(mstrMspUser)
        
        Set objNode = mdocXML.selectSingleNode(".//head/send_hisuser")
        If mstrUserName = "" Then
            objNode.Text = UCase(mstrUser)
        Else
            objNode.Text = UCase(mstrUser) & "/" & mstrUserName
        End If
                
        Set objNode = mdocXML.selectSingleNode(".//head/send_time")
        objNode.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")
        
        strXml = mdocXML.xml
        strXml = Replace(strXml, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""gbk""?>")
        
        RepairMessage = strXml

    End If
      
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    RaiseEvent AfterInfo("����", Err.Description)
End Function

Private Function CheckPollItemCondition(ByVal strConditionSQL As String, ByRef rsCondResult As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set rsCondResult = zlDataBase.OpenSQLRecord(strConditionSQL, mstrTitle)
        
    CheckPollItemCondition = (rsCondResult.RecordCount > 0)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Private Function AnalyseComputer() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Private Function AnalyseIP() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    
    
    On Error GoTo END1
        GetIpAddrTable ByVal 0&, Ret, True
    
    
        If Ret <= 0 Then Exit Function
        ReDim bBytes(0 To Ret - 1) As Byte
        ReDim TempList(0 To Ret - 1) As String
        
        'retrieve the data
        GetIpAddrTable bBytes(0), Ret, False
          
        'Get the first 4 bytes to get the entry's.. ip installed
        CopyMemory Listing.dEntrys, bBytes(0), 4
        
        For Tel = 0 To Listing.dEntrys - 1
            'Copy whole structure to Listing..
            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Next Tel
        'Sort Out The IP For WAN
            TempIP = TempList(0)
            For Tempi = 0 To Listing.dEntrys - 1
                L3 = Left(TempList(Tempi), 3)
                If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                    TempIP = TempList(Tempi)
                End If
            Next Tempi
            AnalyseIP = TempIP 'Return The TempIP
    
    Exit Function
END1:
    AnalyseIP = ""
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Private Function GetIpAddress(Optional ByVal strDefaultIp_Address As String = "") As String
    '******************************************************************************************************************
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '******************************************************************************************************************
    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
    Err = 0: On Error GoTo errHand:
     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "��ȡIP��ַ")
    If rsTemp.EOF = False Then
        strIp_Address = zlCommFun.NVL(rsTemp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    GetIpAddress = strIp_Address
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitXMLDoc() As Object

    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split("6.0,4.0", ",")
    
    On Error Resume Next
    For intLoop = 0 To UBound(varXMLVersion)
        Err = 0
        Set objXML = CreateObject("MSXML2.DOMDocument." & varXMLVersion(intLoop))
        If Err = 0 Then
            strXMLVer = varXMLVersion(intLoop)
            Exit For
        End If
    Next
    On Error GoTo errHand
    
    If strXMLVer = "" Then
        MsgBox "����MSXML2.DOMDocument����ʧ��"
        Exit Function
    End If
    
    Set InitXMLDoc = objXML
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Sub Class_Terminate()
    Set mclsMipServiceData = Nothing
End Sub