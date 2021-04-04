VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_PCR_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim blnEndSample As Boolean '�걾����
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As Variant, strItemResult As String, i As Integer
    Dim lngNo As Long
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "Sample information:")
        '�ж�һ�����Ƿ�ʼ
        If iStrPoint > 0 Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint
            '�����Ƿ�����
            iPos = InStr(iStrPoint, strSource, "Cycles" & Chr(vbKeyTab) & "F1")
            If iPos > 0 Then
                '���ҵ�ǰ�걾��ʼλ��
                Do While iStrPoint < iPos
                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                    If UCase(Mid(strBuffer, 1, 6)) = "NUMBER" Then Exit Do
                Loop
                If iStrPoint < iPos Then
                    Do While True
                        strBuffer = Trim(ReadLine(strSource, iStrPoint))
                        If Len(strBuffer) = 0 Then Exit Do
                        
                        aDatas = getArray(strBuffer)
                        If UBound(aDatas) >= 4 Then
                            If aDatas(1) = "Sample" Then
                                '����ʱ��
                                'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                                strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                                '������ţ�����Ϊ����
                                Select Case Mid(aDatas(2), 1, 2)
                                    Case "BQ", "BS", "NQ", "NS", "TQ", "TS", "CT"
                                        strData = strData & "|" & Val(Mid(aDatas(2), 3))
                                    Case Else
                                        strData = strData & "|" & Val(Mid(aDatas(2), 2))
                                End Select
                                'strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", Val(Mid(aDatas(2), 3)), Val(Mid(aDatas(2), 2)))
                                '������
                                strData = strData & "| " '��������֧��
                                '�걾
                                strData = strData & "|Ѫ��" '��������֧��
                                '�ʿ�Ʒ
                                strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", 1, 0)
                                
                                'ͨ����
                                Select Case Mid(aDatas(2), 1, 2)
                                    Case "BQ", "BS", "NQ", "NS", "TQ", "TS", "CT"
                                        strData = strData & "|" & Mid(aDatas(2), 1, 2)
                                    Case Else
                                        strData = strData & "|" & Mid(aDatas(2), 1, 1)
                                End Select
                            
'                                strData = strData & "|" & IIf(Left(aDatas(2), 1) = "Q", Mid(aDatas(2), 1, 2), Mid(aDatas(2), 1, 1))
                                
                                '���
                                strData = strData & "|" & aDatas(3)
                
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                            End If
                        End If
                    Loop
                End If
                
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                Exit Do
            End If
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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject, objStream As TextStream
    Dim lngResultCount As Long
    Dim strSource As String
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iPos As Long
    Dim aDatas() As Variant
    Dim strIDCode As String  '��ʶ��
    Dim strItem As String
    
    strItem = ";HBV;HCV;T B;HBY;HCF;HPV;BQ0;"
    
    Set objStream = objFile.OpenTextFile(strFile)
    strSource = objStream.ReadAll
    objStream.Close
        
    '��ʼ������ֵ
    iStrPoint = 1
    lngResultCount = -1
        
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "Sample information:")
        '�ж�һ�����Ƿ�ʼ
        If iStrPoint > 0 Then
            '�����Ƿ�����
            iPos = InStr(iStrPoint, strSource, "Standard Line Data")
            If iPos > 0 Then
                '���ҵ�ǰ�걾��ʼλ��
                Do While iStrPoint < iPos
                    strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                    If UCase(Mid(strBuffer, 1, 6)) = "NUMBER" Then Exit Do
                Loop
                If iStrPoint < iPos Then
                    Do While True
                        strBuffer = Trim(ReadLine(strSource, iStrPoint))
                        If Len(strBuffer) = 0 Then Exit Do
                        
                        aDatas = getArray(strBuffer)
                        If UBound(aDatas) >= 1 And InStr(strItem, ";" & Mid(aDatas(1), 1, 3)) > 0 Then
                            
                            If UBound(aDatas) >= 2 Then
                                strData = Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & Mid(aDatas(1), 4) & "| |Ѫ��|0|" & Mid(aDatas(1), 1, 3) & "|" & aDatas(2)
                            Else
                                strData = Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & Mid(aDatas(1), 4) & "| |Ѫ��|0|" & Mid(aDatas(1), 1, 3) & "| "
                            End If
                            
                            lngResultCount = lngResultCount + 1
                                ReDim Preserve aResults(lngResultCount)
                                aResults(lngResultCount) = strData
                        End If
                    Loop
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    ResultFromFile = aResults
    For iPos = 0 To UBound(aResults)
        Debug.Print aResults(iPos)
    Next
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ���ַ��������ַ�ָ���Ƶ���һ�п�ʼλ��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos > 0 Then
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 2
    Else
        iPos = InStr(iPoint, strSource, Chr(13))
        If iPos = 0 Then iPos = InStr(iPoint, strSource, Chr(10))
        If iPos = 0 Then iPos = Len(strSource) + 1
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function getArray(ByVal strData As String) As Variant()
    Dim i As Long
    Dim aTmp() As String, aData() As Variant
    
    strData = Replace(strData, Chr(vbKeyTab), " ")
    aTmp = Split(strData, " "): aData = Array()
    
    For i = 0 To UBound(aTmp)
        If Len(Trim(aTmp(i))) > 0 Then
            ReDim Preserve aData(UBound(aData) + 1)
            aData(UBound(aData)) = aTmp(i)
        End If
    Next
    getArray = aData
End Function
Private Function CheckNumber(strNumber As String) As Integer
    '����:          ȡ��һ��Ϊ���ֵ��ַ�λ��
    '����:          ȡ�ַ����ִ�
    
    Dim intloop As Integer
    
    For intloop = 1 To Len(strNumber)
        If IsNumeric(Mid(strNumber, intloop, 1)) = True Then
            If intloop > 1 Then
                CheckNumber = intloop - 1
            Else
                CheckNumber = intloop
            End If
            Exit Function
        End If
    Next
    CheckNumber = Len(strNumber)
End Function

