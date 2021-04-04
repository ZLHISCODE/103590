VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SEAC_DELTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�����SEAC DELTA ���ֵ��׷�����
'����:�����SEAC
'����:�¶�
'�û�:��������ҽԺ
'Ҫ��Ϊ��ģʽ����
Private mstrCode As String
Private mstrTotal As String '����Ӧ��

Public Function GetAnswerCmd() As String
    '�Զ�Ӧ��ָ��
    GetAnswerCmd = ",06"
End Function

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
    Dim str_Code As String, lng_S As Long, lng_E As Long
    Dim str_Tmp As String
    Dim var_Code As Variant, str_Line As String, i As Long, k As Long
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim str_last As String
    
    On Error GoTo errH
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    If strSource <> "" Then mstrTotal = mstrTotal & strSource
    '----------------------------------------
    lng_S = InStr(mstrTotal, ",01")
    Do While lng_S > 0
        mstrTotal = Mid(mstrTotal, lng_S + 1)
        lng_S = InStr(mstrTotal, ",01")
    Loop
    If Len(mstrTotal) / 3 >= 129 Then strCmd = ",06"
    '----------------------------------------
    lng_S = InStr(mstrCode, ",01")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ",04")
    Else
        lng_E = InStr(mstrCode, ",04")
    End If
    If lng_E > 0 Then
        If InStr(lng_E + 3, mstrCode, ",04") > 0 Then
            lng_E = InStr(lng_E + 3, mstrCode, ",04")
        Else
            If Mid(mstrCode, lng_E + 3) <> "" Then
                lng_E = 0
            End If
        End If
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 3)
        mstrCode = Mid(mstrCode, lng_E + 3)
        
        '-------------------------------------------------------
        str_Tmp = ""
        Do While InStr(str_Code, ",01") > 0
            lng_S = InStr(str_Code, ",01")
            If str_last <> Mid(str_Code, lng_S + 3 * 3, 128 * 3) Then
                str_tmp = str_tmp & Mid(str_Code, lng_S + 3 * 3, 128 * 3)
            End If
            str_last = Mid(str_Code, lng_S + 3 * 3, 128 * 3)
            
            str_Code = Mid(str_Code, lng_S + 3 * 3 + 129 * 3)
        Loop
        If str_Tmp <> "" Then str_Code = Mid_bin(str_Tmp, 1)
        WriteLog "seac_delta", "", str_Code
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like "[[]*[]]" Then
                str_�걾�� = Split(var_Code(i + 1), "=")(1)
                str_���� = Split(var_Code(i + 13), "=")(1) & "-" & Split(var_Code(i + 12), "=")(1) & "-" & Split(var_Code(i + 11), "=")(1)
                str_��Ŀ = Split(var_Code(i + 14), "=")(1)
                str_��� = Split(var_Code(i + 15), "=")(1)
                
                If Val(str_�걾��) > 0 And IsDate(str_����) And str_��Ŀ <> "" And str_��� <> "" Then
                    strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0"
                    For k = LBound(Split(str_��Ŀ, ",")) To UBound(Split(str_��Ŀ, ","))
                        If Trim(Split(str_���, ",")(k)) <> "" And Trim(Split(str_��Ŀ, ",")(k)) <> "" Then
                            strResult = strResult & "|" & Trim(Split(str_��Ŀ, ",")(k)) & "|" & Trim(Split(str_���, ",")(k))
                        End If
                    Next
                End If
            End If
        Next
        '-------------------------------------------------------
        lng_S = InStr(mstrCode, ",01")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ",04")
        Else
            lng_E = InStr(mstrCode, ",04")
        End If
        If lng_E > 0 Then
            If InStr(lng_E + 3, mstrCode, ",04") > 0 Then
                lng_E = InStr(lng_E + 3, mstrCode, ",04")
            Else
                If Mid(mstrCode, lng_E + 3) <> "" Then
                    lng_E = 0
                End If
            End If
        End If
    
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "����ʱ����", Err.Description, str_Code
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
