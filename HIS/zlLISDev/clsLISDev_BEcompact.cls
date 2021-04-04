VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BEcompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�����ж�Ժ

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
    Const strSplit = "|"
    Dim i As Integer, iPos As Integer
    Dim aItemData() As String
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String, blnEnd As Boolean
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    If strSource Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
        '----- ��ʼ�������
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ����������Ϣ
            strLine = GetStr_Section(strReserved, Chr(&H2), Chr(&H3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)

    
            iStrPoint = 1
            Do While iStrPoint <= Len(strLine)
                strBuffer = Trim(ReadLine(strLine, iStrPoint))
                Do While Len(strBuffer) > 10
                    aItemData = Split(strBuffer, strSplit)
                    If UBound(aItemData) < 7 Then '��Ч����
                       iCurrGrpEntry = iPos
                       Exit Do
                    End If
                    i = Trim(aItemData(2)) 'ÿһ�ж�ȡ�걾��Ϣ��i��Ϊ����ͨ���Ļ���
                      '����ʱ��
                       'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    
                       strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                       iStrPoint = iStrPoint + 1
                       '������ţ�����Ϊ����
                     
                       strSampleID = Val(Trim(aItemData(1)))
                       strData = strData & "|" & strSampleID
                       '������
                       strData = strData & "| " '��������֧��
                       '�걾
                       strData = strData & "|ѪҺ" '��������֧��
                       '�ʿ�Ʒ
                       strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                          
                        strData = strData & "|" & 3 * i + 1
                        strData = strData & "|" & Val(Trim(aItemData(5)))
                        strData = strData & "|" & 3 * i + 2
                        strData = strData & "|" & Val(Trim(aItemData(6)))
                        strData = strData & "|" & 3 * i + 3
                        strData = strData & "|" & Val(Trim(aItemData(7)))
                        '�������ѽ���
                        iCurrGrpEntry = 0
                
                    If Len(strData) > 0 Then strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                     strData = ""
                     iPos = iStrPoint
                     strBuffer = Trim(ReadLine(strLine, iStrPoint))
                 Loop
            Loop
            '����δ������ԭʼ��
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '���豸���͡������յ�������
'                strCmd = Chr(&H6&)
'            End If
            
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
        strCmd = Chr(&H6&)
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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



