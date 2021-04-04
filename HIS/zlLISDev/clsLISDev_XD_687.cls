VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XD_687"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  XD687 ����� Ѹ��
'ҽԺ:  ������ҽԺ(̫ԭ)
'����:  ������
'ʱ��:  2013/04/7

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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim str_Code As String
    Dim str_Line As String               '��ǰҪ�������ִ�
    Dim lng_S As Long                  '��ʼλ��
    Dim lng_E As String                '����λ��
    
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String
    
    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    If mstrCode Like "*XD  687*AG*" & Chr(&HA) & "*" Then
        
        Do While InStr(mstrCode, "AG") > 0
            lng_S = InStr(mstrCode, "XD  687")
            lng_E = InStr(lng_S, mstrCode, "AG")
            lng_E = InStr(lng_E, mstrCode, Chr(&HA))
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            str_���� = "": str_�걾�� = "": str_��Ŀ = ""
            Do While str_Code <> ""
                lng_S = 1
                lng_E = InStr(str_Code, Chr(&HA))
                str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                str_Code = Mid(str_Code, lng_E + 1)
                
                If str_Line Like "DATE*" Then
                    str_���� = Trim(Mid(str_Line, 5))
                ElseIf str_Line Like "Sample No*" Then
                    str_�걾�� = Val(Mid(str_Line, 12))
                ElseIf str_Line Like "*=*" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, InStr(str_Line, "=") - 1)) & "|" & Trim(Mid(str_Line, InStr(str_Line, "=") + 1, 5))
                End If
                
            Loop
            
            If Val(str_�걾��) <> 0 And IsDate(str_����) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0" & str_��Ŀ
            End If
        Loop
    End If
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
End Function










