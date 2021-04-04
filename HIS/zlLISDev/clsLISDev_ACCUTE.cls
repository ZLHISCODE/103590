VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACCUTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private mstrCode As String

'����:ACCUTE
'����:��֥
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
        
    Dim lngA As Long, M As Long, N As Long
    Dim str�걾�� As String, str��Ŀ As String, str��� As String, str���� As String
    Dim str_Hc As String, str_Hc1 As String
    
    On Error GoTo errhand:
    
    mstrCode = mstrCode & strSource
    
    If mstrCode Like "*" & Chr(&H3) Then
        '�������������������䣬���԰�����������ķ�ʽ���н���
        lngA = UBound(Split(mstrCode, Chr(&H3)))
        For M = 0 To lngA - 1
            '��ȡ�����걾��Ϣ���н���
            str_Hc = Split(mstrCode, Chr(&H3))(M)
            
            '��ȡ�걾��
            str�걾�� = Split(str_Hc, Chr(&HA))(0)
            str�걾�� = Trim(Mid(str�걾��, 4))
            str�걾�� = Mid(str�걾��, 1, InStr(str�걾��, Chr(&H20)) - 1)
            str���� = Format(Now, "yyyy-mm-dd hh:mm:ss")
 
            strResult = strResult & "||" & str���� & "|" & Val(str�걾��) & "| |����Ѫ|0"
            
            'ȡ�����Ϣ
            str_Hc = Split(str_Hc, Chr(&HA))(1)
            str_Hc = Mid(str_Hc, 5)
            
            For N = 1 To 100
                str��Ŀ = Mid(str_Hc, 1, 10)
                str��� = Val(Trim(Mid(str��Ŀ, 4)))
                str��Ŀ = Trim(Mid(str��Ŀ, 1, 3))
                
                strResult = strResult & "|" & str��Ŀ & "|" & str���
                
                '���Ѿ���ȡ�Ľ��ɾ��
                str_Hc = Mid(str_Hc, 11)
                If Len(str_Hc) < 10 Then
                    Exit For
                End If
            Next
        Next
        
        strResult = Mid(strResult, 3)
        '������ɣ���ջ���
        mstrCode = ""
    End If
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





