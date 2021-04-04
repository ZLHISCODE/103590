VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_PSD_16A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:PSD-16A����ʷ�����
'����:�Ͼ����´�����������޹�˾
'����:�¶�
'�û�:�ɶ���˾
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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_���� As String, str_������ As String, str_��Ŀ As String, str_��� As String
    Dim lng_str As Long, lng_End As Long, str_Line As String, var_Tmp As Variant, i As Integer, str_Item As String
    
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    strResult = ""
    lng_End = InStr(mstrCode, Chr(&HD) & Chr(&HA))
    Do While lng_End > 0
        str_Line = Mid(mstrCode, 1, lng_End + 1)
        mstrCode = Mid(mstrCode, lng_End + 2)
        
        str_Line = DelSapce(str_Line)
        str_Line = Replace(str_Line, Chr(&H0), "")
        
        If str_Line Like "####-##-## ##:## #### *=*" Then
            str_Line = Replace(str_Line, "= ", "=")
            var_Tmp = Split(str_Line, " ")
            str_���� = var_Tmp(0) & " " & var_Tmp(1)
            str_������ = Trim(var_Tmp(2))
            str_��Ŀ = ""
            For i = 3 To UBound(var_Tmp)
                str_Item = var_Tmp(i)
                If InStr(str_Item, "=") > 0 Then
                    str_��� = Trim(Split(str_Item, "=")(1))
                    If str_��� <> "" And Trim(Split(str_Item, "=")(0)) <> "" Then
                        str_��Ŀ = str_��Ŀ & "|" & Trim(Split(str_Item, "=")(0)) & "|" & str_���
                        
                    End If
                End If
            Next
            If str_��Ŀ <> "" And str_���� <> "" And str_������ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & str_������ & "| |Ѫ��|0" & str_��Ŀ
            End If
        End If
        lng_End = InStr(mstrCode, Chr(&HD) & Chr(&HA))
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





