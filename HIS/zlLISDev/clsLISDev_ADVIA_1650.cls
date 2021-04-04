VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_1650"
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
    '1R 010102420070803N11            01-01                                  M   20070803 1.013  1M    21.2     4M     4.5     7M   43.20    10M    23.2    13M    37.7 *  16M    60.5 *R 19M     174    22M   76.42    28M    4.17    31M    7.47    34M     300    37M   221.4    40M      37    43M   87.40    49M    1.62    52M    3.72    67M    1.67    70M    1.00    73M    0.97    76M  131.39    85M   74.94 D 324M     121l  325M     3.0l  326M      90l   
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim iStrEndLeng As Integer '�ִ�����
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim strSampleType As String '��������
    Dim varTmp() As String      '�걾��
    Dim intloop As Integer      'ѭ����
    Dim blnVar As Boolean       '�Ƿ���ֵ
    Dim intHightLow As Integer  '=1�� =0��
    Dim blnQc  As Boolean       '�Ƿ��ʿ�
    Dim intStartPos As Integer
    Dim strHead  As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    '��Ӧ����
    If InStr(strSource, Chr(&H5&)) > 0 Then
        strCmd = Chr(&H6&)
        strReserved = ""
        Exit Sub
    End If
    
    'û������ʱ�˳���������������� &H17 �� &H3)
    If InStr(strSource, Chr(&H17&)) = 0 And InStr(strSource, Chr(&H3&)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strHead = Mid(strSource, 1, 30)
    
    iStrPoint = 1
    
    iStrPoint = IIf(InStr(strHead, "N") = 0, InStr(strHead, "I"), InStr(strHead, "N"))
    If iStrPoint = 0 Then
        iStrPoint = InStr(strHead, "C")
    End If
    iStrPoint = iStrPoint - 8
    
'    iStrPoint = InStr(strSource, Chr(&H20&)) + 2
'    iStrPoint = iStrPoint + 6
    strBuffer = ReadSubString(strSource, iStrPoint, 8)
    strBuffer = Mid(strBuffer, 1, 4) & "-" & Mid(strBuffer, 5, 2) & "-" & Mid(strBuffer, 7)
    'ʱ��
    strData = Format(strBuffer, "YYYY-MM-DD") & " " & Format(Now, "HH:MM:SS")
    
    iStrPoint = InStr(strHead, "N") + 1
    If iStrPoint = 1 Then iStrPoint = InStr(strHead, "I") + 1
    '�����ʿ�
    If iStrPoint = 1 Then
        If InStr(strSource, "PA") > 0 Then iStrPoint = InStr(strSource, "PA") + 2: intHightLow = 0
        If InStr(strSource, "PB") > 0 Then iStrPoint = InStr(strSource, "PB") + 2: intHightLow = 1
        If InStr(strSource, "PC") > 0 Then iStrPoint = InStr(strSource, "PC") + 2: intHightLow = 2
        If InStr(strSource, "PD") > 0 Then iStrPoint = InStr(strSource, "PD") + 2: intHightLow = 3
        If InStr(strSource, "PE") > 0 Then iStrPoint = InStr(strSource, "PE") + 2: intHightLow = 4
        If InStr(strSource, "PF") > 0 Then iStrPoint = InStr(strSource, "PF") + 2: intHightLow = 5
        blnQc = True
    End If
    
    iStrEndLeng = InStr(iStrPoint, strSource, Chr(&H20&))
    iStrEndLeng = iStrEndLeng - iStrPoint
    strBuffer = ReadSubString(strSource, iStrPoint, iStrEndLeng)
    '�걾��
    If blnQc = False Then
        strData = strData & "|" & Val(Mid(strBuffer, 2))
    Else
        If intHightLow = 0 Then
            strData = strData & "|" & 9100 + Val(strBuffer)
        ElseIf intHightLow = 1 Then
            strData = strData & "|" & 9200 + Val(strBuffer)
        ElseIf intHightLow = 2 Then
            strData = strData & "|" & 9300 + Val(strBuffer)
        ElseIf intHightLow = 3 Then
            strData = strData & "|" & 9400 + Val(strBuffer)
        ElseIf intHightLow = 4 Then
            strData = strData & "|" & 9500 + Val(strBuffer)
        Else
            strData = strData & "|" & 9600 + Val(strBuffer)
        End If
    End If
    '������
    strData = strData & "| "
    'Ѫ��
    strData = strData & "|Ѫ��"
    '�ʿ�Ʒ
    strData = strData & IIf(blnQc, "|1", "|0")
    
    varTmp = Split(strSource, "M")
    
    If strSource Like "*1R*" Then intStartPos = 1
    If strSource Like "*2R*" Then intStartPos = 0
    
    For intloop = intStartPos To UBound(varTmp)
        If intloop = intStartPos Then
            '��ʼ��������
            strData = strData & "|" & Trim(Right(varTmp(intloop), 3)) & "M"
        ElseIf intloop = UBound(varTmp) Then
            '������������
            strData = strData & "|" & Val(Trim(Mid(varTmp(intloop), 1, Len(varTmp(intloop)) - 4)))
        Else
            '�м䲿�ݴ���
            strData = strData & "|" & Val(Trim(Mid(varTmp(intloop), 1, Len(varTmp(intloop)) - 4))) & "|" & Trim(Right(varTmp(intloop), 3)) & "M"
            
        End If
        
    Next
    Debug.Print strData
'    varTmp = Split(Mid(strSource, iStrPoint))
'    For intloop = 0 To UBound(varTmp)
'        '����������ʱ�˳�
'        If InStr(varTmp(intloop), Chr(&H3&)) > 0 Or InStr(varTmp(intloop), Chr(&H17&)) > 0 Then
'            Exit For
'        End If
'
'        If blnVar = True And varTmp(intloop) <> "" And InStr(varTmp(intloop), "M") = 0 Then
'            strData = strData & "|" & Val(varTmp(intloop))
'            blnVar = False
'        End If
'
'        If blnVar = False Then
'            If InStr(varTmp(intloop), "M") Then
'                If varTmp(intloop) <> "M" Then
'                    strData = strData & "|" & varTmp(intloop)
'                    blnVar = True
'                Else
'                    blnVar = False
'                End If
'            Else
'                blnVar = False
'            End If
'        Else
'            If InStr(varTmp(intloop), "M") Then
'                If varTmp(intloop) <> "M" Then
'                    strData = strData & "| "
'                    blnVar = False
'                End If
'            End If
'        End If
'    Next
'    If Mid(strData, Len(strData)) = "M" Then
'        strData = strData & "| "
'    End If
    strResult = strData
    strReserved = ""
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
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



