VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BS800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BS-800,BS-400
'����:����
'�û�:����ʡ�����н���������ҽԺ
'������Ա��������
'�������ڣ�2014-09-03
Private strDeviceString '������Ϣ���豸��
Private mstrCode As String
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    On Error GoTo hErr '�򿪴�������
    '------------------------------------------------
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, int_Loop As Long, i As Integer
    Dim str_Code As String, str_Line As String, var_Code As Variant, strTmp As String, varTmp As Variant, strTmp2 As String
    Dim str_ʱ�� As String, str_������ As String, str_������ As String, str_�걾 As String, str_�Ƿ��ʿ� As String, str_��� As String
    Dim str_���� As String
    Dim str_��Ϣ���� As String, str_Ӧ�� As String
    Dim iResultType As Integer  '1Ϊ��������2Ϊ������������
    '--����
    mstrCode = mstrCode & strSource
    lng_S = InStr(strSource, Chr(&HB))
    
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
    End If
    
    '��������
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 2)
        var_Code = Split(str_Code, Chr(&HD))
        
        If str_Code <> "" Then
            str_��Ϣ���� = Split(var_Code(0), "|")(8)
        End If
        Select Case Mid(str_��Ϣ����, 1, 3)
            Case "ORU"         '������
              If Split(var_Code(0), "|")(15) = 2 Then '�ж��ʿ�
                  iResultType = 1
                
                
                str_Line = var_Code(int_Loop)
                    'If Mid(str_Line, 2, 4) = "MSH|" Then
                    '���汾����Ϣ�����Ϣ
                    varTmp = Split(str_Line, "|")
                    str_��Ϣ���� = varTmp(8)
                    str_Ӧ�� = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                    For i = 6 To UBound(varTmp)
                        strTmp2 = varTmp(i)
                        If InStr(1, strTmp2, "ORU^") > 0 Then
                            strTmp2 = Replace(strTmp2, "ORU", "ACK")
                        End If
                        '����Ӧ����Ϣ
                        str_Ӧ�� = str_Ӧ�� & "|" & strTmp2
                    Next
                    str_Ӧ�� = str_Ӧ�� & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                    str_Ӧ�� = str_Ӧ�� & Chr(&H1C) & Chr(&HD)
                        
                
                
                
                  str_ʱ�� = Split(var_Code(0), "|")(6)
                  str_ʱ�� = Mid(str_ʱ��, 1, 4) & "-" & Mid(str_ʱ��, 5, 2) & "-" & Mid(str_ʱ��, 7, 2) & " " & Mid(str_ʱ��, 9, 2) & ":" & Mid(str_ʱ��, 11, 2) & ":00"
                  str_��� = str_ʱ�� & "|" & Mid(Split(var_Code(0), "|")(6), 1, 8) & "| |Ѫ��|1" & "|" & Split(var_Code(1), "|")(2) & "|" & Split(var_Code(1), "|")(20)
                  strResult = strResult & "||" & str_���
                  Exit Do
              Else
                iResultType = 1
                For int_Loop = LBound(var_Code) To UBound(var_Code)
                    str_Line = var_Code(int_Loop)
                    If Mid(str_Line, 2, 4) = "MSH|" Then
                    '���汾����Ϣ�����Ϣ
                    varTmp = Split(str_Line, "|")
                    str_��Ϣ���� = varTmp(8)
                    str_Ӧ�� = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                    For i = 6 To UBound(varTmp)
                        strTmp2 = varTmp(i)
                        If InStr(1, strTmp2, "ORU^") > 0 Then
                            strTmp2 = Replace(strTmp2, "ORU", "ACK")
                        End If
                        '����Ӧ����Ϣ
                        str_Ӧ�� = str_Ӧ�� & "|" & strTmp2
                    Next
                    str_Ӧ�� = str_Ӧ�� & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                    str_Ӧ�� = str_Ӧ�� & Chr(&H1C) & Chr(&HD)
                        
                    ElseIf Mid(str_Line, 1, 4) = "OBR|" Then
                        'ȡ������
                        str_������ = Split(str_Line, "|")(3)
                        If Split(str_Line, "|")(5) = "Y" Then
                            str_������ = str_������ & "^1"
                        End If
                        If Split(str_Line, "|")(2) <> "" Then
                            If InStr(str_������, "^1") > 0 Then
                                str_������ = str_������ & "^" & Split(str_Line, "|")(2)
                            Else
                                str_������ = str_������ & "^^" & Split(str_Line, "|")(2)
                            End If
                            
                        End If
                        
                        'ȡ����ʱ��
                        str_ʱ�� = Split(str_Line, "|")(7)
                        If str_ʱ�� = "" Then
                            str_ʱ�� = Split(str_Line, "|")(14)
                        End If
                        If Len(str_ʱ��) = 14 Then
                            str_ʱ�� = Mid(str_ʱ��, 1, 4) & "-" & Mid(str_ʱ��, 5, 2) & "-" & Mid(str_ʱ��, 7, 2) & " " & Mid(str_ʱ��, 9, 2) & ":" & Mid(str_ʱ��, 11, 2)
                        ElseIf Len(str_ʱ��) = 8 Then
                            str_ʱ�� = Mid(str_ʱ��, 1, 4) & "-" & Mid(str_ʱ��, 5, 2) & "-" & Mid(str_ʱ��, 7, 2)
                        End If
                        
                    ElseIf Mid(str_Line, 1, 4) = "OBX|" Then
                        str_��� = str_��� & "|" & Split(str_Line, "|")(3) & "|" & Split(str_Line, "|")(5)
                    End If
                Next
                strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |Ѫ��|0" & str_���
                lng_S = InStr(mstrCode, Chr(&HB))
                If lng_S > 0 Then
                    lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
                Else
                    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
                End If
             End If
            Case "QRY"         '��������
                iResultType = 2
                
                For int_Loop = LBound(var_Code) To UBound(var_Code)
                    str_Line = var_Code(int_Loop)
                    If Mid(str_Line, 2, 4) = "MSH|" Then
                        '���汾����Ϣ�����Ϣ
                        varTmp = Split(str_Line, "|")
                        str_��Ϣ���� = varTmp(8)
                        str_Ӧ�� = varTmp(0) & "|" & varTmp(1) & "|||" & varTmp(2) & "|" & varTmp(3)
                        For i = 6 To UBound(varTmp)
                            strTmp2 = varTmp(i)
                            If InStr(1, strTmp2, "QRY^") > 0 Then
                                strTmp2 = Replace(strTmp2, "QRY", "QCK")
                            End If
                            '����Ӧ����Ϣ
                            str_Ӧ�� = str_Ӧ�� & "|" & strTmp2
                        Next
                        str_Ӧ�� = str_Ӧ�� & Chr(&HD) & "MSA|AA|" & varTmp(9) & "|Message accpted|||0|" & Chr(&HD)
                        str_Ӧ�� = str_Ӧ�� & "ERR|0|" & Chr(&HD)
                        str_Ӧ�� = str_Ӧ�� & "QAK|SR|OK|" & Chr(&HD)
                        str_Ӧ�� = str_Ӧ�� & Chr(&H1C) & Chr(&HD)
                        
                    ElseIf Mid(str_Line, 1, 4) = "QRD|" Then
                        str_���� = Split(str_Line, "|")(8)
                    ElseIf InStr(str_Line, "QRF") > 0 Then
                        str_������ = Split(str_Line, "|")(4)
                    End If
                Next
                strDeviceString = str_Ӧ��   '������Ϣ���豸��
                
                ''��Ϊʵ�ʲ���˫��ʱ������Ҫ��������ţ��յ�����Ż������Ӧ�����⣬Ϊ�˱�֤�����ܹ��������������������Ŵ���걾��
                ''�����Ҫ�ڽ��н�������ʱ�򣬶�����źͱ걾����ͬ�ı걾���д���
                If Mid(Trim(str_����), 1, 8) = Format(Now, "yyyymmdd") Then
                    str_���� = ""
                End If
                
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & str_������ & "^^" & Trim(str_����)
                lng_S = InStr(mstrCode, Chr(&HB))
                If lng_S > 0 Then
                    lng_E = InStr(lng_S, mstrCode, Chr(&H1C) & Chr(&HD))
                Else
                    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
                End If
            Case Else
                lng_S = 0
        End Select
    Loop
    Select Case iResultType
            Case 1
                strCmd = "0|" & str_Ӧ��
            Case 2
                strCmd = "1|" & str_Ӧ��
            Case Else
                strCmd = "0|"
                strResult = ""
        End Select
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    '------------------------------------------------
    Exit Sub
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_BS800"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    Dim strBanCode As String '�Զ�������
    Dim int_Loop As Integer
    blnSuccess = True
    
    strBanCode = Format(Now, "yyyymmdd")
    Select Case iSendStep
      '  Case 0 '��ʼ����
     '       iSendStep = 1
        Case Else '��������
            If strResponse Like "*MSH*" Or Len(Trim(strResponse)) = 0 Then '��������
                If strResponse Like "*MSH*" Or Len(Trim(strResponse)) = 0 Then
                    '�������
                    '�����豸��
                    aItem = Split(strSample, "|")
                    '�걾��SampleID
                    strBuffer = aItem(1)
                    strSampleID = aItem(5)
                    '��ȡ�����ļܺźͱ���
                    strRackNO = ""
                    If strDeviceString <> "" Then
                        SendSample = Chr(&HB) & "MSH|^~\&|||Mindray|BS-800|" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "||DSR^Q03|" & Split(strDeviceString, "|")(9) & "|P|2.3.1||||||ASCII|||" & Chr(&HD)
                    
                        SendSample = SendSample & "MSA|AA|" & Split(strDeviceString, "|")(9) & "|Message accepted|||0|" & Chr(&HD)
                        
                        SendSample = SendSample & "ERR|0|" & Chr(&HD)
                        
                        SendSample = SendSample & "QAK|SR|OK|" & Chr(&HD)
                        
                        SendSample = SendSample & "QRD|" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "|R|D|1|||RD|" & strSampleID & "|OTH|||T|" & Chr(&HD)
                        
                        Dim strTimeTmp As String
                        strTimeTmp = Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "")
                        SendSample = SendSample & "QRF|BS-800|" & strTimeTmp & "|" & strTimeTmp & "|||RCT|COR|ALL||" & Chr(&HD)
                        
                        For int_Loop = 1 To 28
                            Dim strTmp As String
                            Select Case int_Loop
                                Case 3              '��������
                                    strTmp = "DSP|3||" & aItem(2) & "|||" & Chr(&HD)
                                Case 21             '��������
                                    strTmp = "DSP|21||" & IIf(Trim(aItem(5)) = "", strBanCode & aItem(1), aItem(5)) & "|||" & Chr(&HD)
                                Case 22             '�������
                                    strTmp = "DSP|22||" & aItem(1) & "|||" & Chr(&HD)
                                Case 23             '�ͼ�ʱ��
                                    strTmp = "DSP|23||" & Replace(Replace(Replace(Now, "-", ""), " ", ""), ":", "") & "|||" & Chr(&HD)
                                Case 24             '�Ƿ���
                                    If aItem(4) = "1" Then
                                        strTmp = "DSP|24||Y|||" & Chr(&HD)
                                    Else
                                        strTmp = "DSP|24||N|||" & Chr(&HD)
                                    End If
                                Case 26
                                    strTmp = "DSP|26||" & aItem(3) & "|||" & Chr(&HD)
                                Case Else
                                    strTmp = "DSP|" & int_Loop & "|||||" & Chr(&HD)
                            End Select
                            SendSample = SendSample & strTmp
                        Next
                        
                        For i = 10 To UBound(aItem)
                            SendSample = SendSample & "DSP|" & int_Loop & "||" & aItem(i) & "^^^|||" & Chr(&HD)
                            int_Loop = int_Loop + 1
                        Next
                        SendSample = SendSample & "DSC||" & Chr(&HD) & Chr(&H1C) & Chr(&HD)
                    End If
                    iSendStep = 0
                Else
                    SendSample = Chr(6)
                    iSendStep = iSendStep + 1
                End If
            Else
                SendSample = Chr(6)
                iSendStep = 0
                blnSuccess = False '����ʧ��
            End If
    End Select
End Function

