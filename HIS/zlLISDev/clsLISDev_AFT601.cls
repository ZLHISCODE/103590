VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AFT601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstrCode As String
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
    
    Dim str_Date As String, str_No As String, str_val As String
    Dim lngA As Long, lngB As Long, M As Long, N As Long
    Dim str_Hc As String, str_tmp As String, str_Item As String
    
    On Error GoTo errhand
    
    mstrCode = mstrCode & strSource
    
    '�ж��Ƿ���һ������������
    If mstrCode Like "*" & Chr(&HA) & Chr(&HD) Then
        '����֧���������䣬���������н���
        lngA = UBound(Split(mstrCode, Chr(&HA) & Chr(&HD)))
        For M = 0 To lngA - 1
            str_Hc = Split(mstrCode, Chr(&HA) & Chr(&HD))(M)
            str_Date = Mid(str_Hc, 2, 16) & ":00"
            str_No = Split(str_Hc, Chr(&HA))(3)
            str_No = Mid(str_No, 5)
            '�������ݷ���
            strResult = strResult & "||" & str_Date & "|" & Val(str_No) & "| |����Ѫ|0"
            
            '�������
            str_Hc = Mid(str_Hc, 56)
            lngB = UBound(Split(str_Hc, Chr(&HA)))
            For N = 0 To lngB
                If Split(str_Hc, Chr(&HA))(N) <> "" Then
                    str_tmp = Split(str_Hc, Chr(&HA))(N)
                    str_Item = Trim(Mid(str_tmp, 1, InStr(1, str_tmp, "=") - 1))
                    str_val = Trim(Mid(str_tmp, InStr(1, str_tmp, "=") + 1, 7))
                    strResult = strResult & "|" & str_Item & "|" & str_val
                End If
            Next
        Next
        strResult = Mid(strResult, 3)
        '������ɣ���ջ���
        mstrCode = ""
    End If
    Exit Sub
errhand:
    WriteLog "clsLISDev_AFT601", strSource, Err.Description
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



