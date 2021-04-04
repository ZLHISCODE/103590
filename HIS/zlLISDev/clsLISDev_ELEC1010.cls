VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ELEC1010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���� ˫��
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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnMod2 As Boolean 'ģʽ2�ڱ걾��ǰ��01
    Dim blnQc As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode ��2:New Mode
    Dim iResultType As Integer '������ͣ�1����������2����������
    Dim i As Long
    Dim str39 As String, str32 As String, strItemCode As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 0
    
    iStrPoint = 1: blnReadItem = False
    'ȡ�������
    strSource = Replace(strSource, Chr(5), "")
    'ȡ��������
    strSource = Replace(strSource, Chr(4), "")
    'ȡ����ʼ��
    strSource = Replace(strSource, Chr(2), "")
    'ȡ���س����з�
    strSource = Replace(strSource, vbCrLf, "")
    'ȡ�����з�
    strSource = Replace(strSource, Chr(10), "")
    'ȡ���Ʊ���
    strSource = Replace(strSource, Chr(9), "")
    'ȡ��ͨ�����ӷ�
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(23))
        If i = 0 Then Exit Do
        If i + 3 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 4)
        Else
            Exit Do
        End If
    Loop
    
    '��¼�ָ���
    strFldDelim = "|"
    strRepeatDelim = "\" '�ظ����ݷָ���
    strCompDelim = "^" '������ݷָ���
    Do While iStrPoint <= Len(strSource)
        '������Ϣ�Ŀ�ʼλ�ã���H��ʼ
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Right(Split(strBuffer, strFldDelim)(0), 2) Like "*H" Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint
            
            '�����Ƿ�����
            If InStr(iStrPoint, strSource, "L" & strFldDelim & "1") > 0 Then
                aItemData = Split(strBuffer, strFldDelim)
                '��Ϣ���������ٱ���
                iCurrGrpEntry = 0
                iResultType = 1
                
                '���ұ걾��Ϣ
                iStrPoint = iPos: blnReadItem = False
                Do While iStrPoint <= Len(strSource)
                    If Not blnReadItem Then '��ǰ�Ƿ����ڶ�ȡ������Ŀ
                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Right(Split(strBuffer, strFldDelim)(0), 1) = "O" Then Exit Do
                        Loop
                        If iStrPoint > Len(strSource) Then
                            '��ǰ��Ϣ��Ч�����ٴ�����
                            Exit Do
                        End If
                        
                        aItemData = Split(strBuffer, strFldDelim)
                        '����ʱ��
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                
                        '������ţ�����Ϊ����
                        strData = strData & "|" & Val(aItemData(2))
                                
                        '������
                        strData = strData & "| " '��������֧��
                        '�걾
                        strData = strData & "|Ѫ��"
                        '�ʿ�Ʒ
                        strData = strData & "|0"
                        
                        '׼������
                        blnReadItem = True
                        
                        str39 = "": str32 = ""
                    Else
                        Do While iStrPoint <= Len(strSource)
                            iPos = iStrPoint '������ǰλ��
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Right(Split(strBuffer, strFldDelim)(0), 1) = "R" Then
                            '������
                                aItemData = Split(strBuffer, strFldDelim)
                                strItemCode = Split(aItemData(2), strCompDelim)(3)
                                strData = strData & "|" & Val(strItemCode)
                                If InStr(aItemData(3), "^") > 0 Then
                                    strData = strData & "|" & IIf(Len(Split(aItemData(3), "^")(1)) = 0, " ", Split(aItemData(3), "^")(1)) 'Format(Val(Split(aItemData(3), "^")(1)), "0.00")
                                    If strItemCode Like "39?" Then
                                        str39 = IIf(Len(Split(aItemData(3), "^")(1)) = 0, " ", Split(aItemData(3), "^")(1))
                                    End If
                                    If strItemCode Like "32?" Then
                                        str32 = IIf(Len(Split(aItemData(3), "^")(1)) = 0, " ", Split(aItemData(3), "^")(1))
                                    End If
                                Else
                                    strData = strData & "|" & IIf(Len(aItemData(3)) = 0, " ", aItemData(3)) 'Format(Val(aItemData(3)), "0.00")
                                    If strItemCode Like "39?" Then
                                        str39 = IIf(Len(aItemData(3)) = 0, " ", aItemData(3))
                                    End If
                                    If strItemCode Like "32?" Then
                                        str32 = IIf(Len(aItemData(3)) = 0, " ", aItemData(3))
                                    End If
                                End If
                            Else
                                If Right(Split(strBuffer, strFldDelim)(0), 1) <> "C" Then
                                    blnReadItem = False
                                    Exit Do
                                End If
                            End If
                        Loop
                        'һ���걾����
                        '������Ŀ
                        If Trim(str39) <> "" And Trim(str32) <> "" And Val(str32) <> 0 Then
                            strData = strData & "|990|" & (Val(str39) / Val(str32))
                        End If
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        
                        iStrPoint = iPos
                        If InStr("PO", Right(Split(strBuffer, strFldDelim)(0), 1)) = 0 Then
                        '������һ���걾
                            Exit Do
                        End If
                    End If
                Loop
            Else
                Exit Do
            End If
        Else
            iStrPoint = iPos
        End If
    Loop
            
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
        '���豸���͡������յ�������
        Select Case iResultType
            Case 1 '���ؽ��
                strCmd = "0|" & Chr(&H6&)
            Case 2 '������������
                strCmd = "1|" & Chr(&H6&)
            Case Else
                strCmd = Chr(&H6&)
        End Select
    Else
        '���豸���͡������յ�������
        Select Case iResultType
            Case 1 '���ؽ��
                strCmd = "0|" & Chr(&H6&)
            Case 2 '������������
                strCmd = "1|" & Chr(&H6&)
            Case Else
                strCmd = Chr(&H6&)
        End Select
    End If
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
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
            SendSample = Chr(5)
            iSendStep = 1
        Case Else '��������
            If strResponse Like "*" & Chr(6) & "*" Then '��������
                If iSendStep = 1 Then '��һ�η���
                    '�����豸��
                    strDeviceString = GetDeviceString(strSample, blnUndo, iType)
                End If
                strLastString = Mid(strDeviceString, (iSendStep - 1) * 240 + 1, 240)
                strLastString = iSendStep & strLastString
                If iSendStep * 240 >= Len(strDeviceString) Then '���һ�λ�ֻ��һ��
                    strLastString = strLastString & Chr(&H3)
                    blnEnd = True
                Else
                    strLastString = strLastString & Chr(&H17)
                    blnEnd = False
                End If
                strLastString = Chr(&H2) & strLastString & GetCheckCode(strLastString) & Chr(&HD) & Chr(&HA)
                
                SendSample = strLastString
                If blnEnd Then
                    SendSample = SendSample & Chr(&H4)  '���ͽ���
                    iSendStep = 0
                Else
                    iSendStep = iSendStep + 1
                End If
            Else
'                If iSendStep = 1 Then '�ڿ�ʼ������������
                    SendSample = Chr(&H6) & Chr(&H4) '���ͽ���
                    iSendStep = 0
                    blnSuccess = False '����ʧ��
'                Else '���·���
'                    SendSample = strLastString '���·����ϴ�����
'                End If
            End If
    End Select
End Function

Private Function GetDeviceString(ByVal strSample As String, Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    Dim i As Long, aItem() As String
    Dim strItem As String, iSampleType As Integer
    Dim blnNewMode As Boolean
    
    blnNewMode = False
    
    If blnNewMode Then
        'New Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1|||||||||||||"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & Space(22) & "|" & _
            aItem(1) & "^^^^"
        Select Case True
            Case aItem(3) Like "*Ѫ*"
                GetDeviceString = GetDeviceString & "S1"
                iSampleType = 1
            Case aItem(3) Like "*��*"
                GetDeviceString = GetDeviceString & "S2"
                iSampleType = 2
            Case aItem(3) Like "*����*"
                GetDeviceString = GetDeviceString & "S3"
                iSampleType = 3
            Case Else
                GetDeviceString = GetDeviceString & "S5"
                iSampleType = 5
        End Select
        GetDeviceString = GetDeviceString & "^SC"
    
        GetDeviceString = GetDeviceString & "|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||" & iSampleType & "||||||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "C|1|I|^^^^|G" & vbCr & "L|1|N" & vbCr
    Else
        'Current Mode
        aItem = Split(strSample, "|")
        
        GetDeviceString = "H|\^&|||HOST^1|||||H7600|TSDWN^BATCH|P|1"
        GetDeviceString = GetDeviceString & vbCr & "P|1"
        GetDeviceString = GetDeviceString & vbCr & "O|1|" & _
            aItem(1) & "^" & Space(22) & "^"
        GetDeviceString = GetDeviceString & "1" '�걾��Ĭ��ΪѪҺ�������������ϣ�
        iSampleType = 1
'        Select Case True
'            Case aItem(3) Like "*Ѫ*"
'                GetDeviceString = GetDeviceString & "1"
'                iSampleType = 1
'            Case aItem(3) Like "*��*"
'                GetDeviceString = GetDeviceString & "2"
'                iSampleType = 2
'            Case aItem(3) Like "*����*"
'                GetDeviceString = GetDeviceString & "3"
'                iSampleType = 3
'            Case Else
'                GetDeviceString = GetDeviceString & "5"
'                iSampleType = 5
'        End Select
        GetDeviceString = GetDeviceString & "^^"
    
        GetDeviceString = GetDeviceString & "|R1|"
        strItem = ""
        For i = 10 To UBound(aItem)
            strItem = strItem & "\^^^" & aItem(i) & "/" & IIf(blnUndo, "clr", "")
        Next
        If Len(strItem) > 0 Then strItem = Mid(strItem, 2)
        GetDeviceString = GetDeviceString & strItem
        
        GetDeviceString = GetDeviceString & "|" & IIf(iType = 1, "S", "R")
        GetDeviceString = GetDeviceString & "||" & Format(aItem(0), "yyyymmddhhMMss")
        GetDeviceString = GetDeviceString & "||||N" & "||||||||^^^^||||||O" & vbCr
        
        GetDeviceString = GetDeviceString & "L|1|N" & vbCr
    End If
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
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

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
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

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '��ָ��λ�ÿ�ʼ��ȡһ�У�ͬʱָ�붨λ����һ������֮��
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function
