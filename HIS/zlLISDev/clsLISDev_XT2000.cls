VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XT2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����
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
    Dim i As Integer
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
    '06/ 8/1117:50           97ZRSZ   Q  8.3 3.66  1240.358 97.8 33.9  346   890.2090.0550.736  1.7  0.5  6.10.154 55.6---.----.--.--- 51.2161.7               +                                 MPMPMP
    '[kx21]D1U0610125000000000099000000S010400522001240038700741002380032000298003690005400577000380000600060004150010900089001700

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strBarcode As String
    Dim intLoop As Integer
    Dim lngFind As Long
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    intSpot = InStr(strSource, "")
    
    iStrPoint = 1
    lngFind = 1
    
    'û�н�����˳�
'    If InStr(strSource, "D2U") = 0 Then
'        strReserved = strSource
'        '���豸���͡������յ�������
'        strCmd = Chr(&H6&)
'        Exit Sub
'    End If
'    strTmp = Mid(strSource, 1, InStr(strSource, "D2U"))
        
    
    
    
    Do While iStrPoint <= Len(strSource)
        'û�н�����˳�
        If InStr(iStrPoint, strSource, "D2U") = 0 Then
            strReserved = Mid(strSource, iStrPoint)
            '���豸���͡������յ�������
            strCmd = Chr(&H6&)
            Exit Sub
        End If
        strTmp = Mid(strSource, iStrPoint, InStr(iStrPoint, strSource, "D2U"))
        
        For intLoop = 0 To UBound(Split(strTmp, "D1U"))
            If intLoop <> UBound(Split(strTmp, "D1U")) Then
                lngFind = InStr(lngFind, strTmp, "D1U")
                lngFind = lngFind + 4
            Else
                lngFind = lngFind - 4
                If lngFind <= 0 Then
                    lngFind = 1
                End If
            End If
        Next
        iStrPoint = iStrPoint + lngFind - 1
        iStrPoint = InStr(iStrPoint, strSource, "D1U")
        '�ж�һ�����Ƿ�ʼ
        If iStrPoint > 0 Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint
            
            iPos = InStr(iStrPoint, strSource, "D2U")
            If iPos = 0 Then Exit Do
            If InStr(iPos, strSource, "") > 0 Then
                '������ţ�����Ϊ���� 5λ
                iStrPoint = iStrPoint + 33
                strBarcode = Trim(ReadSubString(strSource, iStrPoint, 15))
                
                '����ʱ��
                strData = ReadSubString(strSource, iStrPoint, 4)   '��
                strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                
                strData = strData & " " & Format(Time, "hh:mm:00") 'ʱ��
                'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                
                strData = strData & "|" & Replace(Val(Right(strBarcode, 9)), "-", "")
                If Len(strBarcode) > 10 Then
                    strData = strData & "^^" & strBarcode
                End If
                
                '������
                strData = strData & "| " '��������֧��
                '�걾
                strData = strData & "|ѪҺ" '��������֧��
                '�ʿ�Ʒ
                strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                
                '������Ŀ
                iStrPoint = iPos + 48
                For i = 1 To 30
                    Select Case i
                        Case 3, 7, 8
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 10
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                        Case 5, 6, 9, 10, 11, 12, 13, 20, 21, 22, 24
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 100
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                        Case 1, 2, 14, 15, 16, 17, 18, 30
                            strItemValue = ReadSubString(strSource, iStrPoint, IIf((i >= 14 And i <= 18) Or i = 1, 6, 5))
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 1000
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                        Case 4, 19, 23, 26, 27, 28, 29
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 10000
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                        Case 25
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 100000
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                        Case Else
                            strItemValue = ReadSubString(strSource, iStrPoint, 5)
                            If IsNumeric(strItemValue) Then
                                strData = strData & "|" & i
                                strData = strData & "|" & Val(strItemValue) / 100
                            ElseIf InStr(strItemValue, "*") > 0 Then
                                strData = strData & "|" & i
                                strData = strData & "|*"
                            End If
                    End Select
                Next
                '�������ѽ���
                iCurrGrpEntry = 0
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
            Else
                Exit Do
            End If
          
            'ָ�붨λ����һ������֮��
        Else
            Exit Do
        End If
    Loop
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '���豸���͡������յ�������
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
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