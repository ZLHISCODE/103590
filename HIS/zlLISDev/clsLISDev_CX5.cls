VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CX5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'˫��
Private strDeviceString '������Ϣ���豸��
Private strLastString As String '�ϴη��͵��豸��

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
    Dim strTP As Long, strALP As Long, strTB As Long, strDB As Long, strIB As Long, strAG As Long, strGLO As Long
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strChennel As String, lngResult As Double
    Dim intDataType As Integer '�������� 0�������1������
    Dim strbarcode As String
    Dim intRequest As Integer   '0���� 1���
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1
    strTB = 0
    strDB = 0
    strTP = 0
    strALP = 0
    Do While iStrPoint <= Len(strSource)
        strBuffer = ReadSubString(strSource, iStrPoint, 1)
        '�ж�һ�����Ƿ�ʼ
        If strBuffer = "[" Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint - 1
            
            iStrPoint = iStrPoint + 3
            strBuffer = ReadSubString(strSource, iStrPoint, 3)
            Select Case strBuffer
            
            Case "702"   '701������702�����
                intRequest = 1
                intDataType = 0
                iStrPoint = iStrPoint + 1
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "03" Then
                    If Len(strSource) - iStrPoint >= 80 Then
                        '����ʱ��
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '��
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        
                        '������ţ�����Ϊ����
                        iStrPoint = iStrPoint + 30
                        
                        strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                        strData = strData & "|" & IIf(Len(strSampleID) = 0, strTmpID, strSampleID)
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|ѪҺ" '��������֧��
                        '�ʿ�Ʒ
                        strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                        
                        '������Ŀ
                        iStrPoint = iStrPoint + 1
                        strChennel = ReadSubString(strSource, iStrPoint, 3)
                        strData = strData & "|" & strChennel
                        iStrPoint = iStrPoint + 19
                        lngResult = Val(ReadSubString(strSource, iStrPoint, 9))
                        strData = strData & "|" & lngResult
                        Select Case strChennel
                            Case "11A"
                                strTB = lngResult
                            Case "12A"
                                strDB = lngResult
                            Case "07A"
                                strTP = lngResult
                            Case "08A"
                                strALP = lngResult
                        End Select
                        If strTB <> 0 And strDB <> 0 Then
                            strIB = strTB - strDB
                            strTB = 0
                            strDB = 0
                            strData = strData & "|IBIL|" & strIB
                        End If
                        If strTP <> 0 And strALP <> 0 Then
                            strGLO = strTP - strALP
                            strAG = strALP / strGLO
                            strTP = 0
                            strALP = 0
                            strData = strData & "|A/G|" & strAG & "|GLO|" & strGLO
                        End If
                        '�������ѽ���
                        iCurrGrpEntry = 0
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                    End If
                Else
                    '��������Ч
                    If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                End If
            Case "701" '�ǽ��
                intRequest = 0
                intDataType = 1
                iStrPoint = iStrPoint + 1
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "06" Then '�����������룬һ������߸� [-0,701,06,samp1------,samp2------,samp3------,samp4------,samp5------,samp6------,samp7------]
                    Do While iStrPoint < Len(strSource)
                        
                        iStrPoint = iStrPoint + 1
                        strbarcode = Trim(ReadSubString(strSource, iStrPoint, 11)) '����
                        strSampleID = Val(Mid(strbarcode, Len(strbarcode) - 1))
                        If Val(strbarcode) <> 0 Then
                            strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            strData = strData & "|" & strSampleID & "^0^" & strbarcode
                                    
                            '������
                            strData = strData & "| " '��������֧��
                            '�걾
                            strData = strData & "|ѪҺ" '��������֧��
                            '�ʿ�Ʒ
                            strData = strData & "|0"
                            strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        '�������ѽ���
                        End If
                        iCurrGrpEntry = 0
                    Loop
                End If
            End Select
            'ָ�붨λ����һ������֮��
            MoveToNextLine strSource, iStrPoint
        End If
    Loop
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    If intRequest = 0 Then
        '���豸���͡������յ�������
        strCmd = "1|" & Chr(&H6) & Chr(&H3)
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
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim blnEnd As Boolean
    blnSuccess = True
    Select Case iSendStep
        Case 0 '��ʼ����
            SendSample = Chr(4) & Chr(1)
            iSendStep = 1
        Case Else '��������
            If strResponse Like "*" & Chr(6) & "*" Then '��������
                If iSendStep = 1 Then '��һ�η���
                    '�����豸��
                    strDeviceString = GetDeviceString(strSample, blnUndo, iType)
                End If
'                strLastString = Mid(strDeviceString, (iSendStep - 1) * 240 + 1, 240)
'                strLastString = iSendStep & strLastString
'                If iSendStep * 240 >= Len(strDeviceString) Then '���һ�λ�ֻ��һ��
'                    strLastString = strLastString & Chr(&H3)
'                    blnEnd = True
'                Else
'                    strLastString = strLastString & Chr(&H17)
'                    blnEnd = False
'                End If
                strLastString = strDeviceString & GetCheckCode(strDeviceString) & Chr(&HD) & Chr(&HA)
                
                SendSample = strLastString
'                If blnEnd Then
'                    SendSample = SendSample & Chr(&H4)  '���ͽ���
'                    iSendStep = 0
'                Else
'                    iSendStep = iSendStep + 1
'                End If
            Else
'                If iSendStep = 1 Then '�ڿ�ʼ������������
                    SendSample = Chr(&H4) '���ͽ���
                    iSendStep = 0
                    blnSuccess = False '����ʧ��
'                Else '���·���
'                    SendSample = strLastString '���·����ϴ�����
'                End If
            End If
    End Select
End Function
Private Function GetDeviceString(ByVal strSample As String, Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    Dim I As Long, aItem() As String
    Dim strItem As String, iSampleType As Integer
    Dim blnNewMode As Boolean
    
    '    [00,701,01,01,03,1,RO,SE,samp3------,CONTROL_NAME________,
'-------------------------,SAMPLE_COMMENT_TWO_______,
'------------------,PATIENT_F_NAME_,M,------------,
'------------------,------,----,--------------------,
'000,4,------,M,-------------------------,-------,
'----,----,------,001,01B-,0]2F<CR><LF>

    blnNewMode = False
    
    If blnNewMode Then
        'New Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "[00,701,01,00,00,1,RO,SE," & aItem(1) & Space(11 - Len(aItem(1))) & "," & _
        aItem(1) & Space(20 - Len(aItem(1))) & "," & Space(25) & "," & Space(25) & "," & Space(18) & "," & Space(15) & "," & Space(1) & "," & Space(12) & _
         "," & Space(18) & "," & Format(aItem(0), "ddmmyy") & "," & Format(aItem(0), "hhmm") & "," & Space(20) & "," & Space(3) & "," & Space(1) & "," & Space(6) & "," & Space(1) & "," & Space(25) & _
         "," & Space(7) & "," & Space(4) & "," & Space(4) & "," & Space(6) & "," & Space(3)
            
        'ȡͨ������
        strItem = ""
        For I = 10 To UBound(aItem)
            strItem = strItem & "," & aItem(I) & ",0"
        Next
        GetDeviceString = GetDeviceString & strItem
        GetDeviceString = GetDeviceString & "]"
    End If
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = "]" Then
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

Private Function GetCheckCode(ByVal strSource As String) As String
'��ȡУ����
    Dim lngAsc As Long, I As Integer
    
    For I = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, I, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function
