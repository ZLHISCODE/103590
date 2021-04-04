VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ATB"
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

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strBarcode As String                    '����
    Dim strVerifyDate As String                 '����ʱ��
    Dim strSampleNumber As String               '�������
    Dim strVerifyName As String                 '������
    Dim intStart As Integer                     '��¼��ʱ�ֽ⿪ʼλ��
    Dim intEnd As Integer                       '��¼��ʱ�ֽ����λ��
    Dim strSpecimen As String                   '�ͼ�걾
    Dim intLoop As Integer                      '��ʱѭ������
    
    Dim blnReadItem As Boolean                  '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    intSpot = InStr(strSource, "")
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
        
            iStrPoint = InStr(iStrPoint, strSource, "NO:")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iCurrGrpEntry = iStrPoint
                
                iPos = InStr(iStrPoint, strSource, "")        '���������˳�
                If iPos = 0 Then Exit Do
                If InStr(iPos, strSource, "") > 0 Then
                
                    iStrPoint = InStr(iStrPoint, strSource, "�걾���:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '�걾���
                    strSampleNumber = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    iStrPoint = InStr(iStrPoint, strSource, "�������:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '�ͼ�걾������ָϸ����
                    strSpecimen = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    iStrPoint = InStr(iStrPoint, strSource, "��������:") + 5
                    intEnd = InStr(iStrPoint, strSource, vbCrLf)
                    
                    '����ʱ��
                    strVerifyDate = ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                    
                    strData = strData & strVerifyDate & " " & Format(Time, "hh:mm:ss")
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    
                    strData = strData & "|" & strSampleNumber & "^0"
                    
                    strData = strData & "| |" & strSpecimen & "|0"
                    
                    '�ҵ���Ŀ��ʼ��λ��
                    
                    iStrPoint = InStr(iStrPoint, strSource, vbCrLf & vbCrLf)
                    
                    For intLoop = 1 To 6
                        strTmp = ReadSubString(strSource, iStrPoint, 1)
                        If Asc(strTmp) >= 65 And Asc(strTmp) <= 122 Then
                            Exit For
                        End If
                    Next
                    
                    iStrPoint = iStrPoint - 1
                    blnReadItem = True
                End If
            Else
                Exit Do
            End If
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '�������
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                strData = ""
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                                        
                'ͨ����
                iStrPoint = iStrPoint - 1
                intEnd = InStr(iStrPoint, strSource, " ")
                strData = strData & "|" & ReadSubString(strSource, iStrPoint, intEnd - iStrPoint)
                
                
                '��ҩ
                iStrPoint = InStr(iStrPoint, strSource, "   ") + 3
                intEnd = InStr(iStrPoint, strSource, " ")
                strTmp = Trim(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                Select Case strTmp
                    Case "��ҩ"
                        strTmp = "R"
                    Case "�н�"
                        strTmp = "I"
                    Case "����"
                        strTmp = "S"
                End Select
                strData = strData & "|" & strTmp
                
                '����
                intEnd = InStr(iStrPoint, strSource, ":") + 1
                strTmp = Trim(ReadSubString(strSource, iStrPoint, intEnd - iStrPoint))
                Select Case strTmp
                    Case "MIC:"
                        strTmp = 1
                    Case "DISK:"
                        strTmp = 2
                    Case "K-B:"
                        strTmp = 3
                End Select
                strData = strData & "^" & strTmp
                
                '���
                intEnd = InStr(iStrPoint, strSource, " ")
                
                strTmp = ReadSubString(strSource, iStrPoint, IIf(intEnd - iStrPoint = 0, 1, intEnd - iStrPoint))
                strData = strData & "^" & Trim(strTmp)
                iStrPoint = InStr(iStrPoint, strSource, vbCrLf) + 2
                
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

