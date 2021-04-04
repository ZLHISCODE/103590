VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU1000"
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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnMod2 As Boolean 'ģʽ2�ڱ걾��ǰ��01
    Dim blnQc As Boolean
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
'            iStrPointBK = iStrPoint
'            iStrPoint = InStr(iStrPoint, strSource, "DB")
'            If iStrPoint > 0 Then
'                blnMod2 = False
'            Else
'                iStrPoint = iStrPointBK
'                iStrPoint = InStr(iStrPoint, strSource, "D01B")
'                If iStrPoint > 0 Then blnMod2 = True
'            End If
            strSource = Replace(strSource, "DB", "")
            strSource = Replace(strSource, "DE", "")
            strSource = Replace(strSource, "D01B", "")
            strSource = Replace(strSource, "D01E", "")
            blnMod2 = True
            
            iStrPoint = InStr(iStrPoint, strSource, "D")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iStrPoint = iStrPoint + 1
                iCurrGrpEntry = iStrPoint - 1
                '�����Ƿ�����
'                If Len(strSource) - iStrPoint >= 192 Then
                iPos = InStr(iStrPoint, strSource, "")
                If iPos > 0 Then
                    '���ҵ�ǰ�걾��ʼλ��
                    Do While iStrPoint > 1 And iStrPoint < iPos
                        iStrPointBK = iStrPoint
                        iStrPoint = InStr(iStrPoint, strSource, "D") + 1
                    Loop
                    iStrPoint = iStrPointBK
                    
                    '����ʱ��
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            
                    '������ţ�����Ϊ����
                    iStrPoint = iStrPoint + IIf(blnMod2, 3, 1)
                    If Mid(strSource, iStrPoint, 1) = "Q" Then '�ʿ�
                        blnQc = True
                        iStrPoint = iStrPoint + 1
                        strData = strData & "|" & (Val(Mid(strSource, iStrPoint, 3)) + _
                            Val(Mid(strSource, iStrPoint + 3, 21)) * 100 + 1000)
                    Else
                        blnQc = False
                        strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 4))
                    End If
                            
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|ѪҺ" '��������֧��
                    '�ʿ�Ʒ
                    strData = strData & IIf(blnQc, "|1", "|0")
                    
                    '׼������
                    iStrPoint = iStrPoint + IIf(blnQc, 24, 180)
                    blnReadItem = True
                End If
            Else
                Exit Do
            End If
        Else
            strBuffer = ReadSubString(strSource, iStrPoint, 1)
            If strBuffer = "" Then '�������
                '������򵰰�(3-4)��ͨ���룺50
                iPos = InStr(28, strData, "|03|")
                sng1 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
                iPos = InStr(28, strData, "|04|")
                sng2 = IIf(iPos = 0, 0, Val(Mid(strData, iPos + 4, 8)))
                strData = strData & "|50|" & Mid(CStr(sng1 - sng2), 1, 8)
                '�ǻ�Ѫ�쵰��()
                
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
            Else
                '������Ŀ
                iStrPoint = iStrPoint - 1
                strData = strData & "|" & ReadSubString(strSource, iStrPoint, 2)
                strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 8))
            End If
        End If
    Loop
            
'    If blnReadItem Then
'        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
'        '�������ѽ���
'        iCurrGrpEntry = 0
'    End If
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