VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ADVIA120Ѫ�������
'����:�ݶ�ADVIA 120Ѫ�������
'����:���
'�û�:����ҽѧ����ҽԺ
'����:2010-03-12

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
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String
    Dim varline As Variant, blnEnd As Boolean
    Dim varTemp As Variant
    Dim i As Integer, str_Line As String
    Dim lngLength As Long
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        
        Do While Not blnEnd
            lngLength = InStr(strReserved, Chr(3)) - InStr(strReserved, Chr(2)) + 1
            str_Data = Mid(strReserved, InStr(strReserved, Chr(2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            
            varline = Split(str_Data, Chr(&HD) & Chr(&HA))
            
            str_��Ŀ = ""
            
            For i = LBound(varline) To UBound(varline) - 1
                str_Line = varline(i)
                If str_Line Like "*0R*" Then
                    varTemp = Split(str_Line, " ")
                    str_ʱ�� = Format(Mid(varTemp(2), 7, 2) & "-" & Mid(varTemp(2), 4, 2) & "-" & Mid(varTemp(2), 1, 2), "yyyy-MM-dd")
                    str_������ = Val(varTemp(5))
                ElseIf i = 1 Then
                    str_��Ŀ = str_��Ŀ & "|WBC|" & Val(Mid(str_Line, 1, 6)) & "|RBC|" & Val(Mid(str_Line, 7, 6)) & "|HGB|" & Val(Mid(str_Line, 14, 6)) & "|HCT|" & _
                                Val(Mid(str_Line, 21, 6)) & "|MCV|" & Val(Mid(str_Line, 28, 6)) & "|MCH|" & Val(Mid(str_Line, 35, 6)) & "|MCHC|" & Val(Mid(str_Line, 42, 6)) & "|RDW|" & _
                                Val(Mid(str_Line, 49, 6)) & "|HDW|" & Val(Mid(str_Line, 56))
                ElseIf i = 2 Then
                    str_��Ŀ = str_��Ŀ & "|PLT|" & Val(Mid(str_Line, 1, 6)) & "|MPV|" & Val(Mid(str_Line, 7, 6)) & "|PDW|" & Val(Mid(str_Line, 14, 6)) & "|PCT|" & _
                                Val(Mid(str_Line, 21, 6))
                ElseIf i = 3 Then
                    str_��Ŀ = str_��Ŀ & "|NEUT#|" & Val(Mid(str_Line, 1, 6)) & "|LYMPH#|" & Val(Mid(str_Line, 7, 6)) & "|MONO#|" & Val(Mid(str_Line, 14, 6)) & "|EO#|" & _
                                Val(Mid(str_Line, 21, 6)) & "|BASE#|" & Val(Mid(str_Line, 28, 6)) & "|LUC#|" & Val(Mid(str_Line, 35, 6))
                ElseIf i = 4 Then
                    str_��Ŀ = str_��Ŀ & "|NEUT%|" & Val(Mid(str_Line, 1, 6)) & "|LYMPH%|" & Val(Mid(str_Line, 7, 6)) & "|MONO%|" & Val(Mid(str_Line, 14, 6)) & "|EO%|" & _
                                Val(Mid(str_Line, 21, 6)) & "|BASO%|" & Val(Mid(str_Line, 28, 6)) & "|LUC%|" & Val(Mid(str_Line, 35, 6))
                ElseIf i = 5 Then
                    str_��Ŀ = str_��Ŀ & "|CHCM|" & Val(Mid(str_Line, 1, 6)) & "|CH|" & Val(Mid(str_Line, 7, 6))
                End If
                
            Next
            If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "| |ȫѪ|0" & str_��Ŀ
            End If
            If Not strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
       strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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



