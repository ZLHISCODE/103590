VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cd3700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ���� cell-DYN3700
'����:��������
'����:�¶�
'�û�:�»�
'�޸�ʱ��:2009-09-02
'�޸���:���

Private mstrCode As String    '��������

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
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_���� As String, var_line As Variant
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim str_WaitSendLine As String, str_���� As String, str_�ʿ� As String
    Dim str_Last�걾�� As String, str_Tmp As String
    '����ͼ������
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str���������ʽ As String       '������������ʽ
    
    str_in = strSource
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = 0
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    End If
    
    '��ȡ���������ʽ
    str���������ʽ = Trim(ReadIni("CD3700", "���������ʽ", App.Path & "\CD3700.ini"))
    If str���������ʽ = "" Then str���������ʽ = "Ĭ��": Call WriteIni("CD3700", "���������ʽ", str���������ʽ, App.Path & "\CD3700.ini")
    
    Do While lng_S > 0 And lng_E > 0
        
        str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "��ʼ", "�յ���������", str_Line
        str_Line = Replace(Replace(str_Line, """", ""), Chr(&H3), "")
       
        var_line = Split(str_Line, ",")
        str_Tmp = Trim(var_line(0))
        str_�걾�� = Trim(var_line(4))
        str_���� = var_line(7)
        str_���� = Format(CDate(Replace(str_����, "/", "-")), "yyyy-MM-dd") & " " & var_line(8)
        If Not IsDate(str_����) Then str_���� = ""
        '�жϽ��������Ǽ���������ͼ��
        Select Case str_Tmp
            Case Chr(&H2) & ""  '������
                For i_Count = 9 To UBound(var_line)
                    str_��� = Trim(var_line(i_Count))
                    If str_��� <> "" Then
                        Select Case str���������ʽ
                            Case "Ĭ��"
                                str_��Ŀ = str_��Ŀ & "|" & i_Count - 8 & "|" & str_���
                            Case "�������ַ�"
                                str_��Ŀ = str_��Ŀ & "|" & i_Count - 8 & "|" & IIf(IsNumeric(str_���), str_���, " ")
                        End Select
                       ' str_��Ŀ = str_��Ŀ & "|" & i_Count - 8 & "|" & str_���
                    End If
                Next
            Case Chr(&H2) & "WIC"   'WBCͼ��
                str_��Ŀ = "|ͼ��ͨ��|0"
                For i_Count = 10 To UBound(var_line) - 1
                     str_WBC = str_WBC & ";" & var_line(i_Count)
                Next
            Case Chr(&H2) & "RBC"   'RBCͼ��
                str_��Ŀ = "|ͼ��ͨ��|0"
                For i_Count = 10 To UBound(var_line) - 1
                    str_RBC = str_RBC & ";" & var_line(i_Count)
                Next
            Case Chr(&H2) & "PLT"   'PLTͼ��
                str_��Ŀ = "|ͼ��ͨ��|0"
                For i_Count = 10 To UBound(var_line) - 1
                    str_PLT = str_PLT & ";" & var_line(i_Count)
                Next
        End Select

        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = 0
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        End If
        
        If Val(str_�걾��) > 0 And str_���� <> "" And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |ѪҺ|0" & str_��Ŀ
            '����ͼ������
            If str_WBC <> "" Then strResult = strResult & vbCrLf & "WBC;0;256,67.5,7.5-,15-100,22.5-," & _
                                                            "30-200,37.5-,45-300,52.5-,60-400,67.5-" & str_WBC
            If str_RBC <> "" Then strResult = strResult & vbCrLf & "RBC;0;256,70,7-,14-50,21-,28-100," & _
                                                            "35-,42-150,49-,56-200,63-,70-250" & _
                                                            str_RBC
            If str_PLT <> "" Then strResult = strResult & vbCrLf & "PLT;0;256,64,8-5,16-10," & _
                                                            "24-15,32-20," & _
                                                            "40-25,48-30,56-35," & _
                                                            "64-40" & str_PLT
        End If
        str_�걾�� = "": str_���� = "": str_��Ŀ = "": str_WBC = "": str_RBC = "": str_PLT = ""
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    WriteLog "����", "������", strResult

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

    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String
    Dim str_SendItem As String, str_Line As String
    
    SendSample = ""
End Function



