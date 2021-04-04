VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACLAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ACLAdvance
'����:�������������
'����:�¶�
'�û�:�����е�������ҽԺ

Private mstrCode As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim var_Code As Variant, intLoop As Integer
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim lng_S As Long, lng_E As Long
    strResult = ""
    
    
    If InStr(strSource, Chr(&H5)) > 0 Then                                         '�յ���ʼ�� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(strSource, Chr(&HA)) > 0 Then                                         '�յ��н����� 0A ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    
   
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))

    If lng_S > 0 And lng_E > 0 Then WriteLog "���� ACLAdvance", mstrCode, ""
    Do While lng_S > 0 And lng_E > 0                       '�յ�һ�������ļ�¼,��ʼ����
        
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '-------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For intLoop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(intLoop)
            If str_Line Like Chr(&H2) & "?O|*" Then                           '�걾��,��Ŀ ��
                str_�걾�� = Val(Split(str_Line, "|")(2))
                If str_�걾�� = 0 Then str_�걾�� = ""
                str_��Ŀ = Split(str_Line, "|")(4)
               
            ElseIf str_Line Like Chr(&H2) & "?R|*" Then                           '��Ŀ ��� ��
                If str_�걾�� <> "" And str_��Ŀ <> "" Then
                    If str_��Ŀ = Split(str_Line, "|")(2) Then
                        str_���� = Split(str_Line, "|")(12)
                        str_��� = Trim(Split(str_Line, "|")(3))
                        'str_��Ŀ = Replace(str_��Ŀ, "^^^", "") & Trim(Split(str_Line, "|")(1))
                        If str_���� <> "" And str_��� <> "" Then
                            str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|0|" & Replace(str_��Ŀ, "^^^", "") & Trim(Split(str_Line, "|")(1)) & "|" & str_���
                        End If
                    End If
                End If
              End If
        Next
        '-------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "���� ACLAdvance End", strResult, ""
    End If
        
End Sub
