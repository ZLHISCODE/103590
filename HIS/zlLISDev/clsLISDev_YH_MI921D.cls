VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_YH_MI921D"
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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim aDatas() As String, i As Integer, strSampleID As String, strTmpID As String
    Dim dbliCa As Double
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    '����ʱ��
    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    strSource = Replace(strSource, vbCr, ""): strSource = Replace(strSource, vbLf, "")
    aDatas = Split(Trim(strSource), " ")
    If UBound(aDatas) >= 6 Then
    '�����ļ�������
        '�������
        strData = strData & "|" & Val(aDatas(0))
        '������
        strData = strData & "| " '��������֧��
        '�걾
        strData = strData & "|ѪҺ" '��������֧��
        '�ʿ�Ʒ
        strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
        For i = 1 To 6
            strData = strData & "|" & i
            strData = strData & "|" & Val(aDatas(i))
            If i = 4 Then dbliCa = Val(aDatas(i)) 'ָ��iCa��ֵ
        Next
        'ָ��tCa
        strData = strData & "|" & i
        strData = strData & "|" & (2 * dbliCa - 0.06)
        
        '�����豣�������ݣ�����һ������ʹ��
        strReserved = ""
        For i = 7 To UBound(aDatas)
            strReserved = strReserved & " " & aDatas(i)
        Next
        If Len(strReserved) > 0 Then strReserved = Mid(strReserved, 2) & IIf(Len(strSource) > Len(Trim(strSource)), " ", "")
        '�������ѽ���
        iCurrGrpEntry = 0
        
        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
    Else
        iCurrGrpEntry = 1 '�������ݲ��������������ݱ���
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