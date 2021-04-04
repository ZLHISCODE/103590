VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BECompactX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:Compact-X Ѫ����
'����:�¹��ݶ�
'����:�¶�
'�û�:��������ҽԺ
'�޸���:���
'�޸�����:2009-08-31

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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim varItem As Variant, str_Line As String, i As Integer
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String
        
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------
        str_��Ŀ = "": str_�걾�� = ""
        varItem = Split(str_Code, Chr(&HD) & Chr(&HA))
        For i = LBound(varItem) To UBound(varItem)
            str_Line = varItem(i)
            If UBound(Split(str_Line, "|")) = 7 Then
                str_�걾�� = Split(str_Line, "|")(1)
                str_��Ŀ = str_��Ŀ & "|" & "1" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(5))
                str_��Ŀ = str_��Ŀ & "|" & "2" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(6))
                str_��Ŀ = str_��Ŀ & "|" & "3" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(7))
            End If
        Next
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_��Ŀ
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
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



