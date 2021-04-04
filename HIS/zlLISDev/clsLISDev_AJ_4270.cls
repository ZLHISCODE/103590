VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AJ_4270"
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
    
    Dim sng1 As Single, sng2 As Single, iPos As Long, iEndPos As Long
    Dim blnMod2 As Boolean 'ģʽ2�ڱ걾��ǰ��01
    Dim blnQc As Boolean
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '����һ�����ݵĿ�ʼλ��
            strSource = Replace(strSource, "", "")
            
            iStrPoint = InStr(iStrPoint, strSource, "")
            '�ж�һ�����Ƿ�ʼ
            If iStrPoint > 0 Then
                '��¼һ���������
                iStrPoint = iStrPoint + 1
                iCurrGrpEntry = iStrPoint - 1
                '�����Ƿ�����
                iPos = InStr(iStrPoint, strSource, "")
                If iPos > 0 Then
                    '���ҵ�ǰ�걾��ʼλ��
                    iStrPoint = iStrPoint + 32
                    
                    '����ʱ��
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    strData = ReadSubString(strSource, iStrPoint, 10)
                    iStrPoint = iStrPoint + 1
                    strData = strData & " " & ReadSubString(strSource, iStrPoint, 5)
                    strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                            
                    '������ţ�����Ϊ����
                    iStrPoint = iStrPoint + 4
                    strData = strData & "|" & Val(ReadSubString(strSource, iStrPoint, 4))
                            
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|��Һ" '��������֧��
                    '�ʿ�Ʒ
                    strData = strData & "|0"
                    
                    '׼������
                    iStrPoint = iStrPoint + 18
                    blnReadItem = True
                    iEndPos = iPos
                End If
            Else
                Exit Do
            End If
        Else
            '������Ŀ
            If iEndPos - iStrPoint < 50 Then '����COLOR��һ���걾����
                strBuffer = Trim(ReadSubString(strSource, iStrPoint, 14))
                strData = strData & "|COLOR"
                strData = strData & "|" & IIf(Len(strBuffer) = 0, " ", strBuffer)
            
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & Replace(strData, "||", "| |")
            
                blnReadItem = False
                '�������ѽ���
                iCurrGrpEntry = 0
                '���տ��ܵ���һ���걾
                iStrPoint = iEndPos + 1
            Else
                strBuffer = Trim(ReadSubString(strSource, iStrPoint, 6))
                If Len(strBuffer) = 0 Then
                    iStrPoint = iStrPoint + 19
                Else
                    strBuffer = Replace(strBuffer, ".", "")
                    strData = strData & "|" & strBuffer
                    strData = strData & "|" & Trim(ReadSubString(strSource, iStrPoint, 13))
                    iStrPoint = iStrPoint + 6 '���Լ�����λ
                End If
            End If
        End If
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

