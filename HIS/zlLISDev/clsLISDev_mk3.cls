VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MK3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LABITEMS As String = ",pH,PCO2,PO2,pHt,PCO2t,PO2t,ctHb,O2Hb,HHb,COHb,MetHb,SulfHb,pHst,cHCO3st,cHCO3,ctCO2(P),ctCO2(B),BE,BEecf,BB,SO2,PAO2t,AaDO2t,a/AO2t,RI,"
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
    strCmd As String, strMBChannel As String, lngMBNo As Long, strMBPosition As String)
   ' Analyse(strInput, strResult, strReserved, strCmd, strMBChannel, lngMBNo, strMBPosition)
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
    Dim CutOff As Double '�ٽ�ֵ
    Dim str��־ As String '����С�ڱ�־
    Dim aItemData As Variant, blnReadItem As Boolean
    Dim i As Integer, j As Integer, intRow As Integer, intCol As Integer
    Dim iPos As Integer
    Dim strDate As String, lngSampleNO As Long, strItemValue As String
    Dim strTmp As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    strTmp = "    1     2     3     4     5     6     7     8     9    10    11    12 " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  '  strSource = Replace(strSource, Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
    strSource = Replace(strSource, strTmp, "")
    intRow = Asc(Mid(strMBPosition, 1, 1)) - 65 + 1
    intCol = Val(Mid(strMBPosition, 2, 2))
    
    '������ʼλ�ã�NEGATIVE CONTROL��Ϊ��ʼ,iStrPoint = 1:
    iStrPoint = InStr(strSource, "NEGATIVE CONTROL")
    strSource = Mid(strSource, iStrPoint)
    blnReadItem = False: str��־ = ""
    strDate = "": lngSampleNO = -1
    If Len(strSource) > 780 Then '����δ����
        Do While iStrPoint <= Len(strSource)
            If Not blnReadItem Then '���ұ걾������Ϣ
                strBuffer = ReadLine(strSource, iStrPoint)
                Do While Len(strBuffer) > 0
                    Select Case True
                        Case InStr(UCase(strBuffer), UCase("CUT OFF")) > 0 '�ٽ�ֵ
                            iPos = InStr(UCase(strBuffer), UCase("CUT OFF"))
                            CutOff = Val(Mid(strBuffer, 17, 5))
                            strBuffer = ""
                        Case InStr(UCase(strBuffer), UCase("INTERPRETATION")) > 0
                            If InStr(strBuffer, ">") > InStr(strBuffer, "+") Then
                                 str��־ = "����"
                            Else
                                 str��־ = "����"
                            End If
                            strBuffer = ""
                        Case InStr(UCase(strBuffer), UCase("ABSORBANCES")) > 0
                            blnReadItem = True
                            Exit Do
                        Case Else
                            strBuffer = ""
                    End Select
                Loop
            Else
                If str��־ <> "" Then  '������Ϣ������ɣ���λ��������Ŀ��ʼλ��
                    strBuffer = ReadLine(strSource, iStrPoint) '�ھ��У�����ʱ��
                    strDate = Format(Replace(strBuffer, ".", ""), "yyyy-mm-dd hh:mm:ss")
                    strBuffer = ReadLine(strSource, iStrPoint) '��ʮ��
               
                '�걾��Ϣ��ʼ����ʼ�걾��Ϊ��������
                   For i = intRow To 8
                      strBuffer = ReadLine(strSource, iStrPoint) '��ʮһ�У����ݿ�ʼ
                      aItemData = Split(strBuffer, " ")
                      For j = intCol To 12
                        lngSampleNO = lngMBNo + (i - intRow) * 12 + j - intCol
                       'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                       strData = strDate
                       '������ţ�����Ϊ����
                       strData = strData & "|" & lngSampleNO
                       '������
                       strData = strData & "| " '��������֧��
                       '�걾
                       strData = strData & "|ѪҺ" '��������֧��
                       '�ʿ�Ʒ
                       strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                       
                       strData = strData & "|" & strMBChannel
                       
                       '�������
                       If Abs(Val(aItemData(j))) < 0.004 Then Exit Do
                       If str��־ = "����" Then
                            If (Val(aItemData(j)) - CutOff) > 0 Then
                               strItemValue = "����"
                            Else
                               strItemValue = "����"
                            End If
                       Else
                            If (Val(aItemData(j)) - CutOff) > 0 Then
                               strItemValue = "����"
                            Else
                               strItemValue = "����"
                            End If
                       End If
                       
                       strData = strData & "|" & strItemValue
                       strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                   Next
                Next
              End If
           End If
        Loop
   Else
      iCurrGrpEntry = 1
   End If
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '���豸���͡������յ�������
        strCmd = ""
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

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ���ַ��������ַ�ָ���Ƶ���һ�п�ʼλ��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function








