VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XT1800i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:xs-1800i
'����:Sysmex
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
    Dim aData() As String
    Dim intLoop As Integer
    Dim strItem As String, strItemVal As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    strCmd = Chr(&H6&)
    
    'û���ҵ���ʼλ�����˳�
    If InStr(strSource, "") <= 0 Then strReserved = strSource: Exit Sub
    
    'û���ҵ�����λ��ʱ�˳�
    If InStr(strSource, "") <= 0 Then strReserved = strSource: Exit Sub
    
    'ȡһ���걾��Ϣ���н������������ȷ����´ν��н���
    intStart = InStr(strSource, "")
    intEnd = InStr(strSource, "") + 1
    strLine = Mid(strSource, intStart, intEnd - intStart)
    
    '����û�н������ִ�
    strReserved = Mid(strSource, intEnd)
    If InStr(strReserved, "") <= 0 Then
        'û���ҵ���ʼ��ʱ����Ϊ��
        strReserved = ""
    End If
    
    
    '��ʼ����
    aData = Split(strLine, vbCrLf)
    
    '����ʱ��
    strTmp = Split(aData(6), "|")(12)
    strData = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        
    '�걾��

    
    If aData(2) Like "*^^2*" Or aData(2) Like "*QC*" Then
        strTmp = Split(aData(2), "|")(3)
        strTmp = Trim(Split(strTmp, "^")(2))
        
        
        strData = strData & "|" & IIf(Val(strTmp) = 0, "1000", Val(strTmp)) & "| |Ѫ��|1"
        For intLoop = 4 To UBound(aData) - 2
            
            strItem = Split(Split(aData(intLoop), "|")(2), "^")(4)
            strItemVal = Split(aData(intLoop), "|")(3)
            
            If InStr(aData(2) & "^PCT^P-LCR^", strItem) > 0 Then
                strData = strData & "|" & strItem & "|" & strItemVal
            End If
        Next
    Else
        strTmp = Split(aData(3), "|")(3)
        strTmp = Trim(Split(strTmp, "^")(2))
        strData = strData & "|" & strTmp & "| |Ѫ��|0"
        
        For intLoop = 6 To UBound(aData) - 2
            
            strItem = Split(Split(aData(intLoop), "|")(2), "^")(4)
            Debug.Print strItem
            strItemVal = Split(aData(intLoop), "|")(3)
            
            If InStr(aData(3) & "^PCT^P-LCR^", strItem) > 0 Or strItem = "P-LCR" Then
                strData = strData & "|" & strItem & "|" & strItemVal
            End If
        Next
    End If
    

    
    strResult = strData
    'strCmd = Chr(&H6&)
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

