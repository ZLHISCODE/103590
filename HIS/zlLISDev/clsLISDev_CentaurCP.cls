VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CentaurCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ADVIA centaur CP ��ѧ�������߷�����
'����:�ݶ�
'����:�¶�
'�û�:����������ҽԺ

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
    Dim var_Code As Variant, str_Line As String, int_Loop As Integer
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String, str_��� As String
    Dim str_Item As String
    On Error GoTo errH
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If str_Line Like "?O|*" Then
            
               If str_�걾�� <> Split(str_Line, "|")(2) Then
                    If str_Item <> "" And Val(str_�걾��) > 0 And str_���� <> "" Then
                        strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_Item
                    End If
                    str_�걾�� = Split(str_Line, "|")(2)
                    str_Item = ""
               End If
               str_��Ŀ = Split(str_Line, "|")(4)
            ElseIf str_Line Like "?R|*" Then
                If InStr(Split(str_Line, "|")(2), str_��Ŀ) = 1 And InStr(Split(str_Line, "|")(2), "^DOSE") > 0 Then
                    str_���� = Split(str_Line, "|")(12)
                    If str_���� <> "" And Len(str_����) > 14 And IsNumeric(Mid(str_����, 1, 14)) Then
                        str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                        If Not IsDate(str_����) Then str_���� = ""
                    Else
                        str_���� = ""
                    End If
                    str_��� = Trim(Split(str_Line, "|")(3))
                    If str_��� <> "" Then
                        str_Item = str_Item & "|" & Replace(str_��Ŀ, "^", "") & "|" & str_���
                    End If
                End If
            End If
        Next
        
        If str_Item <> "" And Val(str_�걾��) > 0 And str_���� <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_Item
        End If
        '--------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    Call WriteLog("CentaurCP", Err.Description & vbCrLf & mstrCode, str_Code)
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    SendSample = ""
End Function
