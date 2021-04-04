VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Roche_9180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Roche 9180
'����:����
'����:
'�û�:����ҽԺ
Private Const LABITEMS As String = "Na,K,Cl"
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
    Dim iPos As Long, iEndPos As Long, i As Integer, aItems() As String
    Dim strDate As String, lngSampleNO As Long, strItemName As String, dblItemValue As Double
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    strDate = "": lngSampleNO = -1
    Do While iStrPoint <= Len(strSource)
        iStrPoint = InStr(iStrPoint, strSource, "")
        '�ж�һ�����Ƿ�ʼ
        If iStrPoint = 0 Then Exit Do
        '��¼һ���������
        iCurrGrpEntry = iStrPoint
        
        iEndPos = InStr(iStrPoint, strSource, "")
        If iEndPos = 0 Then Exit Do
        
        '����ʱ��
        'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
        strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        '������ţ�����Ϊ����
        iStrPoint = InStr(iStrPoint, strSource, "Sample No.") + 10
        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 5))
        
        '������
        strData = strData & "| " '��������֧��
        '�걾
        strData = strData & "|ѪҺ" '��������֧��
        '�ʿ�Ʒ
        strData = strData & "|0"
        
        aItems = Split(LABITEMS, ",")
        For i = 0 To UBound(aItems)
            iPos = InStr(iStrPoint, strSource, Mid(aItems(i) & "  ", 1, 2) & "=")
            If iPos > 0 Then
                strData = strData & "|" & aItems(i)
                strData = strData & "|" & Val(Mid(strSource, iPos + 3, 7))
            End If
        Next
    
        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
        
        '�������ѽ���
        iCurrGrpEntry = 0
        iStrPoint = iEndPos + 1
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
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 2
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function
