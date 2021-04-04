VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XL600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:�¹�ŷ��XL-600ȫ�Զ�����������
'�ͺ�:XL-600
'����:�¹�ŷ��
'Author:���
'User:���ճ���
'Date:2010-05-13

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
    
    Dim lng_S As Long, lng_E As Long
    Dim str_�걾�� As String, str_�걾 As String
    Dim str_�������� As String
    Dim str_������ As String
    Dim var_Tmp() As String
    Dim str_Data As String
    Dim blnEnd As Boolean
    
    strReserved = strReserved & strSource
    lng_S = InStr(strReserved, "<")
    lng_E = InStr(strReserved, ">")
    
    If lng_S > 0 And lng_E > 0 Then
        str_Data = Mid(strReserved, lng_S, lng_E)
        strReserved = Mid(strReserved, lng_E + 1)
        
        str_Data = Replace(Replace(str_Data, "<", ""), ">", "")
        
        var_Tmp = Split(str_Data, "|")
        str_�걾�� = var_Tmp(2)
        
        str_�걾 = "Ѫ��"
        str_�������� = Format(Now, "yyyy-MM-dd hh:mm:ss")
        str_������ = str_������ & "|" & Trim(var_Tmp(3)) & "|" & Trim(var_Tmp(4))
        
        If Val(str_�걾��) > 0 And Trim(str_������) <> "" Then
            strResult = str_�������� & "|" & str_�걾�� & "| |" & str_�걾 & "|0" & str_������
        End If
        strCmd = "  " & Chr(&H3C) & Chr(&H53) & Chr(&H61) & Chr(&H76) & Chr(&H65) & Chr(&H64) & Chr(&H7C) & Chr(&H31) & Chr(&H46) & Chr(&H33) & Chr(&H3E)
    End If
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
End Function