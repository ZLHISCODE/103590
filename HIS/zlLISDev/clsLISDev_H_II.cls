VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H_II"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����H-II��Һ������
'����:
'����:�¶�
'�û�:׼���������ҽԺ

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
    Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String
    Dim lon_str As Long, lon_end As Long
    Dim str_Data As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String
    
    
    
    lon_str = InStr(strSource, Chr(&H2))
    lon_end = InStr(strSource, Chr(&H3))
    If lon_str > 0 And lon_end > 0 Then
        str_Data = Mid(strSource, lon_str, lon_end)
        strReserved = Mid(strSource, lon_end + 1)
        str_��Ŀ = ""
        str_ʱ�� = ""
        
        varLine = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = LBound(varLine) To UBound(varLine)
            str_Line = varLine(i)
            If str_Line Like "*NO*" Then
                str_������ = Val(Mid(str_Line, InStr(str_Line, "NO") + 2, 4))
                str_ʱ�� = Trim(Mid(str_Line, InStr(str_Line, "NO") + 6))
                If UBound(Split(str_ʱ��, "-")) = 2 Then
                    str_ʱ�� = Split(str_ʱ��, "-")(2) & "-" & Split(str_ʱ��, "-")(0) & "-" & Split(str_ʱ��, "-")(1)
                Else
                    str_ʱ�� = Format(Now, "yyyy-MM-dd")
                End If
                
            Else
                If Len(str_Line) > 13 And InStr(str_Line, ":") <= 0 Then
                    If Trim(Mid(str_Line, 2, 4)) <> "" And Trim(Mid(str_Line, 5)) <> "" Then
                        str_��� = Trim(Mid(str_Line, 5))
                        str_��� = Replace(str_���, "umol/L", "")
                        str_��� = Replace(str_���, "mmol/L", "")
                        str_��� = Replace(str_���, "cells/uL", "")
                        str_��� = Replace(str_���, "Cell/uL", "")
                        
                        
                        str_��� = Replace(str_���, "mg/dL", "")
                        str_��� = Replace(str_���, "g/L", "")
                        
                        str_��� = Replace(str_���, "PRo", "PRO")
                        str_��� = Replace(str_���, "URo", "URO")
                        
                        str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 2, 4)) & "|" & Trim(str_���)
                    End If
                ElseIf Trim(str_Line) Like "##:##" Then
                    str_ʱ�� = str_ʱ�� & " " & Trim(str_Line)
                End If
            End If
            
        Next
        If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then

            strResult = str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ" & "|0" & str_��Ŀ
        End If
    Else
        strReserved = strSource
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



