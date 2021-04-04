VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_PA900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private mstrCode As String

'����:PA900
'����:
'����:������
'�û�:��ɽ�ظ��ױ���վ

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
        
    Dim strCode As String
    Dim lngA As Long, lngB As Long, N As Integer, M As Integer
    Dim str�걾�� As String, str��Ŀ As String, str��� As String, str���� As String, str�ʿ� As String
    Dim str_Hc As String, strItem As String  '��Ҫ�����Ĳ������ݷ���һ�����������ڽ���
    
    On Error GoTo errhand
    mstrCode = mstrCode & strSource
    
    '�ж��Ƿ���һ������������
    If mstrCode Like "lis_star*lis_end" & Chr(&HD) & Chr(&HA) Then
        '��������˵����û��˵���������Ƿ�֧���������䣬���а�����������ķ�ʽ���н���
        lngA = UBound(Split(mstrCode, "lis_end" & Chr(&HD) & Chr(&HA)))
        For N = 0 To lngA - 1
            '�������걾����Ϣ��ȡ����,���н���
            str_Hc = Split(mstrCode, "lis_end" & Chr(&HD) & Chr(&HA))(N)
            lngB = UBound(Split(str_Hc, Chr(&HD) & Chr(&HA)))
            For M = 0 To lngB
                strItem = Split(str_Hc, Chr(&HD) & Chr(&HA))(M)
                
                If InStr(strItem, "lis_type") > 0 Then '�Ƿ��ʿ�
                    str�ʿ� = Split(strItem, "=")(1)
                    If str�ʿ� = "qc" Then
                        str�ʿ� = 1
                    Else
                        str�ʿ� = 0
                    End If
                ElseIf InStr(strItem, "lis_s_item") > 0 Then '������Ŀ
                    str��Ŀ = Split(strItem, "=")(1)
                ElseIf InStr(strItem, "lis_s_sample") > 0 Then '�걾��
                    str�걾�� = Val(Split(strItem, "=")(1))
                ElseIf InStr(strItem, "lis_s_date") > 0 Then  '��������
                    str���� = Replace(Split(strItem, "=")(1), "/", "-")
                ElseIf InStr(strItem, "lis_s_result") > 0 Then '������
                    str��� = Trim(Mid(Split(strItem, "=")(1), 1, InStr(Split(strItem, "=")(1), " ")))
                    '���ñ걾Ϊ�ʿ�ʱ���������ᴫ�ر걾�ţ������ֹ�ָ��һ��
                    If str�ʿ� = 1 Then str�걾�� = 800
                    strResult = strResult & "||" & str���� & "|" & str�걾�� & "| |����Ѫ|" & str�ʿ� & "|" & str��Ŀ & "|" & str���
                End If
            Next
        Next
        strResult = Mid(strResult, 3)
        '������ɣ���ջ���
        mstrCode = ""
    End If
    Exit Sub
errhand:
    WriteLog "clsLISDev_ACCUTE", strSource, Err.Description
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





