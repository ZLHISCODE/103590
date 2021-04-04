VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    
    Dim iStrPoint As Integer, iPos As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String, strChannel As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    Dim str����ʱ�� As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1
    If InStr(1, strSource, "[88]") > 0 Then
        Do While iStrPoint <= Len(strSource)
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            '�ж�һ�����Ƿ�ʼ
            If strBuffer = "[" Then
                '��¼һ���������
                iCurrGrpEntry = iStrPoint - 1
                
                'iStrPoint = iStrPoint + 7
                strBuffer = ReadSubString(strSource, iStrPoint, 2)
                If strBuffer = "20" Then
                    If Len(strSource) - iStrPoint >= 80 Then
                        '����ʱ��
                        iStrPoint = iStrPoint + 1
                        strData = ReadSubString(strSource, iStrPoint, 2)   '��
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '��
                        strData = ReadSubString(strSource, iStrPoint, 2) & "-" & strData '��
                        iStrPoint = iStrPoint + 1
                        strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) & ":" & ReadSubString(strSource, iStrPoint, 2)
                        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                        strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                        str����ʱ�� = strData
                        iStrPoint = iStrPoint + 3

                        '������ţ�����Ϊ����
                         Do While iStrPoint <= Len(strSource)
                             strData = ""
                             Do While iStrPoint <= Len(strSource)
                                iPos = iStrPoint '������ǰλ��
                                strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                If Mid(strBuffer, 3, 2) = "21" Then
                                   iStrPoint = iPos + 8
                                   strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                                   strData = str����ʱ�� & "|" & strSampleID
                                    '������
                                   strData = strData & "| " '��������֧��
                                    '�걾
                                   strData = strData & "|ѪҺ" '��������֧��
                                    '�ʿ�Ʒ
                                   strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                                   Exit Do
                                End If
                             Loop
                            
                            '������Ŀ
                            Do While iStrPoint <= Len(strSource)
                                iPos = iStrPoint '������ǰλ��
                                strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                                If Mid(strBuffer, 3, 2) = "22" Then
                                   
                                   strChannel = Trim(ReadSubString(strBuffer, 10, 9))
                                   strData = strData & "|" & strChannel
                                   strData = strData & "|" & Trim(ReadSubString(strBuffer, 22, 6))
                                ElseIf Mid(strBuffer, 3, 2) = "99" Then
                                   Exit Do
                                End If
                             Loop
        
                            '�������ѽ���
                            iCurrGrpEntry = 0
                            
                            strResult = IIf(Len(strResult) > 0, IIf(Len(strData) > 0, strResult & "||", strResult), "") & strData
                        Loop
                    End If
                Else
                    '��������Ч
                    If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                End If
                'ָ�붨λ����һ������֮��
                MoveToNextLine strSource, iStrPoint
            End If
        Loop
        '����δ������ԭʼ��
        If iCurrGrpEntry > 0 Then
            strReserved = Mid(strSource, iCurrGrpEntry)
        End If
        '���豸���͡������յ�������
        strCmd = Chr(&H6&)
    Else
        strReserved = strSource
    End If
    
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
Private Function ReadLine(ByVal strSource, iPoint As Integer) As String
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

