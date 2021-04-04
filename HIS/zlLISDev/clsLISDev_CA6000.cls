VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA6000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���� ˫��
Private mstrSendString As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim strSampleType As String '��������
    Dim strItemValue As String
    Dim iPos  As Long, intDataType As Integer '�������� 0�������1������
    Dim strBarcode As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            intDataType = -1
            iPos = InStr(iStrPoint, strSource, "R")
            If iPos > 0 Then
                '��������
                intDataType = 1
            Else
                '�������
                iPos = InStr(iStrPoint, strSource, "D")
                If iPos > 0 Then intDataType = 0
            End If
            Select Case intDataType
                Case 0 '���
                    '��¼һ���������
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '�����Ƿ�����
                    If InStr(iStrPoint, strSource, "") > 0 Then
                        '����ʱ��
                        iStrPoint = iStrPoint + 8
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '��������
                        
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) 'Сʱ
                        strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '����
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 6
                        strBarcode = Trim(Replace(UCase(ReadSubString(strSource, iStrPoint, 13)), "QC", ""))
                        strData = strData & "|" & Val(Right(strBarcode, 9))
                        If Len(strBarcode) > 10 Then
                            strData = strData & "^^" & strBarcode
                        End If
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:���桢E:STAT������S:��׼���ߡ�C:�ʿ�
                        
                        '׼������
                        iStrPoint = iStrPoint + 12
                        blnReadItem = True
                    Else
                        Exit Do
                    End If
                Case 1 '����
                    '��¼һ���������
                    iStrPoint = iPos + 1
                    iCurrGrpEntry = iStrPoint - 1
                    '�����Ƿ�����
                    iPos = InStr(iStrPoint, strSource, "")
                    If iPos > 0 Then
                        '����ʱ��
                        iStrPoint = iStrPoint + 8
                        strSampleType = ReadSubString(strSource, iStrPoint, 1)   '��������
                        
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) 'Сʱ
                        strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '����
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 6
                        strBarcode = Trim(ReadSubString(strSource, iStrPoint, 13))
                        strData = strData & "|" & Val(Right(strBarcode, 9))
                        If Len(strBarcode) > 10 Then
                            strData = strData & "^^" & strBarcode
                        End If
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:���桢E:STAT������S:��׼���ߡ�C:�ʿ�
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        '�������ѽ���
                        iCurrGrpEntry = 0
                        If InStr(iStrPoint, strSource, "R") = 0 Then Exit Do
                    Else
                        Exit Do
                    End If
                Case Else
                    Exit Do
            End Select
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '�������
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                '������Ŀ
                iStrPoint = iStrPoint - 1
                strBuffer = ReadSubString(strSource, iStrPoint, 3)
                strData = strData & "|" & strBuffer
                'ȷ��С��λ��
                Select Case Mid(strBuffer, 3, 1)
                    Case "3", "4", "5" '���ʡ���׼��
                        strItemValue = ReadSubString(strSource, iStrPoint, 6)
                        If IsNumeric(strItemValue) Then
                            strData = strData & "|" & Val(strItemValue) / 100
                        Else
                            strData = strData & "|*"
                        End If
                    Case Else
                        Select Case strBuffer
                            Case "062" 'Fpg
                                strItemValue = ReadSubString(strSource, iStrPoint, 6)
                                If IsNumeric(strItemValue) Then
                                    strData = strData & "|" & Val(strItemValue) / 100
                                Else
                                    strData = strData & "|*"
                                End If
                            Case Else
                                strItemValue = ReadSubString(strSource, iStrPoint, 6)
                                If IsNumeric(strItemValue) Then
                                    strData = strData & "|" & Val(strItemValue) / 10
                                Else
                                    strData = strData & "|*"
                                End If
                        End Select
                End Select
            End If
        End If
    Loop
            
    If blnReadItem Then
        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
        '�������ѽ���
        iCurrGrpEntry = 0
    End If
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '���豸���͡������յ�������
        Select Case intDataType
            Case 0 '���ؽ��
                strCmd = "0|" & Chr(&H6&)
            Case 1 '������������
                strCmd = "1|" & Chr(&H6&)
            Case Else
                strCmd = ""
        End Select
    End If
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
    Dim iPos As Long, strItem As String, strSampleNO As String
    
    blnSuccess = True
    Select Case iSendStep
        Case 0 '��ʼ����
            If strResponse Like "*R*" Or Len(Trim(strResponse)) = 0 Then '��������
                If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                    If Len(Trim(strSample)) > 0 Then
                        '�������
                        '�����豸��
                        aItem = Split(strSample, "|")
                        '�걾��SampleID
                        strBuffer = aItem(1)
                        strSampleID = aItem(5)
                        '��ȡ�����ļܺźͱ���
                        strBuffer = "210101"
                        If Len(Trim(strResponse)) > 0 Then
                            iPos = InStr(strResponse, "R")
                            strBuffer = Mid(strResponse, iPos + 3, 6)
                        End If
                        SendSample = "S2" & Right(Space(6) & strBuffer, 6) & "U"
                        If Len(Trim(strResponse)) > 0 Then
                            iPos = InStr(strResponse, "R")
                            strBuffer = Mid(strResponse, iPos + 10, 29)
                            SendSample = SendSample & strBuffer & "B" & Space(11)
                        Else
                            SendSample = SendSample & Format(aItem(0), "yymmdd") & Format(Now, "hhmm") & "000101" & _
                                Right(Space(13) & strSampleID, 13) & "B" & Space(11)
                        End If
                        
                        For i = 10 To UBound(aItem)
                            strItem = Right(String(2, "0") & Left(aItem(i), 2), 2) & "0" & Space(6)
                            If InStr(SendSample, strItem) = 0 Then
                                SendSample = SendSample & strItem
                            End If
                        Next
                        SendSample = SendSample & ""
                    
                        iSendStep = 0
                        mstrSendString = SendSample
                    ElseIf Len(Trim(strResponse)) > 0 Then
                        '�������
                        '�����豸��
                        '��ȡ�����ļܺźͱ���
                        iPos = InStr(strResponse, "R")
                        strBuffer = Mid(strResponse, iPos + 3, 6)
                        
                        SendSample = "S2" & Right(Space(6) & strBuffer, 6) & "U"
                            
                        iPos = InStr(strResponse, "R")
                        strBuffer = Mid(strResponse, iPos + 10, 16)
                        strSampleNO = Mid(strResponse, iPos + 29, 10)
                        If Mid(strResponse, iPos + 27, 2) <> "88" Then
                            '���ʿر걾,�걾�ż�10000
                            strSampleNO = Right(Space(10) & (Val(strSampleNO) + 10000), 10)
                        Else
                            strSampleNO = Right(Space(10) & "QC" & Val(strSampleNO), 10)
                        End If
                        SendSample = SendSample & strBuffer & "   " & strSampleNO & "C" & Space(11) & "040      050      060      510      "
                        
                        SendSample = SendSample & ""
                    
                        iSendStep = 0
                        mstrSendString = SendSample
                    End If
                Else
                    SendSample = ""
                    iSendStep = 0
                    blnSuccess = False '����ʧ��
                End If
            Else
                SendSample = ""
                iSendStep = 0
            End If
        Case Else
            If InStr(strResponse, Chr(&H6)) = 0 Then
                'δ�ɹ����ط�
                SendSample = Chr(&H15)
                iSendStep = 0
            Else
                '���ͳɹ�
                SendSample = ""
                iSendStep = 0
            End If
    End Select
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function