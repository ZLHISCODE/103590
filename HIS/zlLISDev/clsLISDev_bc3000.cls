VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_bc3000"
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

    Dim iStrPoint As Integer    '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer   '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    
    iStrPoint = 1
    Do While iStrPoint <= Len(strSource)
        strBuffer = ReadSubString(strSource, iStrPoint, 1)
        '�ж�һ�����Ƿ�ʼ
        If strBuffer = "A" Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint - 1 '63530918
            
            iStrPoint = iStrPoint + 8
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If InStr(iStrPoint, strSource, Chr(&H3)) > 0 Or InStr(iStrPoint, strSource, Chr(&H1A)) > 0 Then
                '����ʱ�� '060520061632
                strData = ReadSubString(strSource, iStrPoint, 2)   '��
                strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
                strData = ReadSubString(strSource, iStrPoint, 4) & "-" & strData '��
                strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) & ":" & ReadSubString(strSource, iStrPoint, 2) 'ʱ��
                'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                
                '������ţ�����Ϊ���� 8λ
                iStrPoint = iStrPoint - 21 '��λ���걾λ
               ' strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 8))
                '������
                strData = strData & "| " '��������֧��
                '�걾
                strData = strData & "|ѪҺ" '��������֧��
                '�ʿ�Ʒ
                strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                '������Ŀ
                iStrPoint = iStrPoint + 13
                For i = 1 To 19
                    strData = strData & "|" & i
                    Select Case i
                        Case 1, 2, 3, 4, 11, 12, 19
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 4)) / 10
                        Case 10, 15
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 4))
                        Case 5, 6
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 3)) / 10
                        Case 7, 13, 14, 16, 17
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 3)) / 10
                        Case 8
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 3)) / 100
                        Case 9
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 3))
                        Case 18
                           strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 3)) / 1000
                    End Select
                Next
                '�������ѽ���
                iCurrGrpEntry = 0
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                '����ͼ������
                lngSector = InStr(iStrPoint, strSource, Chr(&H3))
                If lngSector = 0 Then
                    lngSector = InStr(iStrPoint, strSource, Chr(&H1A))
                End If
                If lngSector = 0 Then
                    iStrPoint = Len(strSource) + 1
                Else
                    iStrPoint = lngSector + 1
                End If
            Else
                Exit Do
            End If
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