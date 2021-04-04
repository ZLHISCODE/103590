VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MI921"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:MI-921
'����:Խ��
'����:����
'�û�:�Ͼ���������ҽԺ
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
    Dim strBarcode As String
    Dim strLine As String
    Dim intStart As Integer, intEnd As Integer
    Dim adata() As String
    Dim intloop As Integer
    Dim strItem As String
    
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    adata = Split(strSource, vbCrLf)
    
    For intloop = 0 To UBound(adata)
        strItem = adata(intloop)
        If Len(strItem) >= 49 Then
            If Val(Mid(strItem, 1, 3)) > 0 Then
                strData = strData & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(Mid(strItem, 1, 3)) & "| |Ѫ��|0" & _
                       "|3|" & Val(Mid(strItem, 18, 4)) & "|4|" & Val(Mid(strItem, 23, 5)) & _
                       "|5|" & Val(Mid(strItem, 29, 5)) & "|6|" & Val(Mid(strItem, 35, 4)) & _
                       "|7|" & Val(Mid(strItem, 40, 4)) & "|8|" & Val(Mid(strItem, 46, 4)) & _
                       "|9|" & Val(Mid(strItem, 35, 4)) * 1.95 & "|10|" & Val(Mid(strItem, 23, 5)) - Val(Mid(strItem, 29, 5)) - Val(Mid(strItem, 46, 4))
            End If
        Else
            strReserved = strReserved & strItem
        End If
'        If UBound(strItem) >= 11 Then
'            strData = strData & "||" & Now & "|" & strItem(0) & "| |Ѫ��|0"
'            strData = strData & "|1|" & Val(strItem(2)) & "|2|" & Val(strItem(4)) & "|3|" & Val(strItem(5)) & _
'                      "|4|" & Val(strItem(6)) & "|5|" & Val(strItem(7)) & "|6|" & Val(strItem(8)) & _
'                      "|7|" & Val(strItem(9)) & "|8|" & Val(strItem(11))
'        End If
    Next
    
    strResult = Mid(strData, 3)
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



