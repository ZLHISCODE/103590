VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_U168"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: U168
'����:
'����: ����
'�û�: �Ͼ���������ҽԺ
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
    Dim strTmp As String, intStart As Integer, intEnd As Integer
    Dim intPosition As Integer                  '��ǰ��
    Dim intloop As Integer
    Dim astrLine() As String
    Dim strItemVal As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û�лس����з�ʱ�˳�
    If InStr(strSource, vbCr) <= 0 Then strReserved = strSource: Exit Sub
    '���ҿ�ʼλ��
    astrLine = Split(strSource, vbCrLf)
    intStart = 0
    For intloop = 0 To UBound(astrLine)
        If IsNumeric(Trim(Mid(astrLine(intloop), 1, 2))) = True Then
            intStart = InStr(strSource, Mid(astrLine(intloop), 1, 4))
            Exit For
        End If
    Next
    'û���ҵ���ʼλ��ʱ�˳�
    If intStart = 0 Then strReserved = strSource: Exit Sub
    
    '���ҽ���λ��
    intEnd = InStr(strSource, "ASC")
    If intEnd <= 0 Then strReserved = strSource: Exit Sub
    'δ�ҵ�����λ���˳�
    'If InStr(intEnd, strSource, "neg") <= 0 Then strReserved = strSource: Exit Sub
    
    strLine = Mid(strSource, intStart, intEnd + 24 - intStart)
    
    'ֻ����һ��,���������´ν���
    strReserved = Replace(Mid(strSource, intEnd + 30), vbCrLf & vbCrLf, "")
    
    astrLine = Split(strLine, vbCrLf)
    
    strData = Format(Mid(astrLine(0), 4, 2) & "-" & TrunMonth(Mid(astrLine(0), 7, 3)) & "-" & Mid(astrLine(0), 11, 2), "yyyy-mm-dd") & " " & Trim(Mid(astrLine(0), 13))
    strData = strData & "|" & Val(Mid(astrLine(2), 10, 4)) & "| |��|0"
    
    For intloop = 3 To UBound(astrLine)
        If intloop = 7 Or intloop = 9 Then
            strItemVal = Format(Val(Trim(Mid(astrLine(intloop), 6))), "0.00")
            strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItemVal
        Else
            strItem = Replace(Trim(Mid(astrLine(intloop), 6, 9)), "-", "��")
            strItemVal = Format(Val(Trim(Mid(astrLine(intloop), 15))), "0.00")
            If Val(strItemVal) = 0 Then
                strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItem & "  " & Trim(Mid(astrLine(intloop), 15))
            Else
                strData = strData & "|" & Trim(Mid(astrLine(intloop), 2, 3)) & "|" & strItem & "  " & strItemVal
            End If
        End If
    Next
    strResult = strData
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


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function

