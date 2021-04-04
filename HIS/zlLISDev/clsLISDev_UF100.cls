VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UF100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: UF-100
'����: sysmex
'����: (δ֪)
'�û�: ҽ���Ժ
'�޸�: 2007-08-16 �¶�
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
    Dim strTmpReturn As String, blnEnd As Boolean
    Dim lngLength As Long, str���� As String, str�걾�� As String, str�����˱걾�ʿ� As String, strLine As String
    Dim strItem As String, i As Integer, blnLineEnd As Boolean
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "DS*" & Chr(3) & Chr(2) & "DP*" & Chr(3) & Chr(2) & "DC*" & Chr(3) & "*" Then
        strTmpReturn = strSource
        
        Do While Not blnEnd
            
            lngLength = InStr(strTmpReturn, Chr(3)) - InStr(strTmpReturn, Chr(2)) + 1
            strLine = Mid(strTmpReturn, InStr(strTmpReturn, Chr(2)), lngLength)
            strTmpReturn = Mid(strTmpReturn, InStr(strTmpReturn, Chr(3)) + 1)
            
            If strLine Like "*DS*" Then
                strItem = "": str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": blnLineEnd = False
                str���� = Format(CDate(Mid(strLine, 11, 2) & "-" & Mid(strLine, 13, 2) & "-" & Mid(strLine, 15, 2)), "yyyy-MM-dd")
                If Mid(strLine, 40, 1) = " " Then
                    '��ҽ��Ժ��ʽ
                    str�걾�� = Val(Mid(strLine, 27, 13))
                ElseIf Mid(strLine, 44, 1) = " " Then
                    '2008-3-11 ���ӳɶ�������ʽ
                    str�걾�� = Val(Mid(strLine, 31, 13))
                End If
                str�����˱걾�ʿ� = " |��Һ|" & IIf(Mid(strLine, 10, 1) = "C", 1, 0)
                
                Select Case Mid(strLine, 47, 1)
                    Case "0"
                        strItem = "|RBCInfo|RBC Negative"
                    Case "1"
                        strItem = "|RBCInfo|Normocytic"
                    Case "2"
                        strItem = "|RBCInfo|Microcytic"
                    Case "3"
                        strItem = "|RBCInfo|Non-classified"
                End Select
            ElseIf strLine Like "*DP*" Then
                strLine = Replace(Mid(strLine, 12), Chr(3), "")
                For i = 0 To 4
                    strItem = strItem & "|" & Mid(strLine, 1 + 12 * i, 4) & "|" & Val(Mid(strLine, 5 + 12 * i, 8))
                Next
            ElseIf strLine Like "*DC*" Then
                strLine = Replace(Mid(strLine, 12), Chr(3), "")
                If strLine <> "" Then
                    For i = 0 To Len(strLine) / 4
                        If Mid(strLine, 1 + 4 * i, 4) <> "" Then
                            strItem = strItem & "|" & Mid(strLine, 1 + 4 * i, 4) & "|+"
                        End If
                    Next
                End If
                blnLineEnd = True
            End If
            
            If Val(str�걾��) > 0 And strItem <> "" And blnLineEnd Then
                strResult = strResult & "||" & str���� & "|" & Val(str�걾��) & "|" & str�����˱걾�ʿ� & strItem
            End If
            
            If Not strTmpReturn Like "*" & Chr(2) & "*" & Chr(3) & "*" Or blnLineEnd Then
                blnEnd = True
            End If
        Loop
'        strTmpReturn = Mid(strSource, InStr(strSource, Chr(2)) + 192)
'        strSource = Mid(strSource, InStr(strSource, Chr(2)), 192)
        
            'Beging ԭ���ĳ���
'            iStrPoint = 1: blnReadItem = False
'            Do While iStrPoint <= Len(strSource)
'                If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
'                    strBuffer = ReadSubString(strSource, iStrPoint, 2)
'                    iStrPoint = iStrPoint - 1
'                    '�ж�һ�����Ƿ�ʼ
'                    If strBuffer = "DS" Then
'                        '��¼һ���������
'                        iCurrGrpEntry = iStrPoint - 1
'                        '�����Ƿ�����
'                        If Len(strSource) - iStrPoint >= 117 Then
'                            '����ʱ��
'                            iStrPoint = iStrPoint + 7
'                            strSampleType = ReadSubString(strSource, iStrPoint, 1)   '��������
'
'                            strData = ReadSubString(strSource, iStrPoint, 2)   '��
'                            strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
'                            strData = strData & "-" & ReadSubString(strSource, iStrPoint, 2) '��
'                            strData = strData & " " & ReadSubString(strSource, iStrPoint, 2) 'Сʱ
'                            strData = strData & ":" & ReadSubString(strSource, iStrPoint, 2) & ":00" '����
'                            'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
'                            strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
'
'                            '������ţ�����Ϊ����
'                            iStrPoint = iStrPoint + 6
'                            strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 13))
'
'                            '������
'                            strData = strData & "| " '��������֧��
'                            '�걾
'                            strData = strData & "| " '��������֧��
'                            '�ʿ�Ʒ
'                            strData = strData & "|" & IIf(strSampleType = "C", 1, 0) 'U:���桢E:STAT������S:��׼���ߡ�C:�ʿ�
'
'                            '׼������
'                            iStrPoint = iStrPoint + 80
'                            blnReadItem = True
'                        End If
'                    End If
'                Else
'                    strBuffer = ReadSubString(strSource, iStrPoint, 1)
'                    If strBuffer = "" Then '�������
'                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'
'                        blnReadItem = False
'                        '�������ѽ���
'                        iCurrGrpEntry = 0
'                    Else
'                        '������Ŀ:0201��ϸ����0202��ϸ��
'                        iStrPoint = iStrPoint - 1
'                        strBuffer = ReadSubString(strSource, iStrPoint, 4)
'                        strData = strData & "|" & strBuffer
'                        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 8))
'                    End If
'                End If
'            Loop
'
'            If blnReadItem Then
'                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'                '�������ѽ���
'                iCurrGrpEntry = 0
'            End If
'            '����δ������ԭʼ��
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '���豸���͡������յ�������
'                strCmd = Chr(&H6&)
'            End If
            'End ԭ���ĳ���
            '
            If strTmpReturn <> "" Then strReserved = strTmpReturn
    Else
        strReserved = strSource
        strCmd = Chr(&H6&)
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
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