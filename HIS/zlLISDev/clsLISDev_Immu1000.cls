VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Immu1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:IMMULITE 1000 ȫ�Զ��������߷�����
'����: DPC
'����: �¶�
'�û�: ����3201ҽԺ
Private mstrCode As String      '��������ԭʼ����
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
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
    Dim lng_S As Long, lng_E As Long, str_Code As String, str_Line As String, var_Code As Variant
    Dim int_Loop As Integer, str_��� As String, str_�걾�� As String, str_��Ŀ As String, str_���� As String
    
    '-- ���� Ӧ��
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    'If InStr(strSource, Chr(&H4)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)

    If Len(strSource) <> 0 Then mstrCode = mstrCode & strSource
    
    '--����
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 1) = Chr(&H2) And Mid(str_Line, 3, 2) = "O|" Then
                If str_�걾�� <> Split(str_Line, "|")(2) Then
                
                    If str_�걾�� <> "" And str_��� <> "" And str_���� <> "" Then
                        strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_���
                        str_��� = "": str_���� = ""
                    End If
                    str_�걾�� = Split(str_Line, "|")(2)
                    
                End If
                str_��Ŀ = Trim(Split(str_Line, "|")(4))
            ElseIf Left(str_Line, 1) = Chr(&H2) And Mid(str_Line, 3, 2) = "R|" Then
                
                If str_�걾�� <> "" And str_��Ŀ = Trim(Split(str_Line, "|")(2)) Then
                    str_���� = Trim(Split(str_Line, "|")(11))
                    
                    If IsNumeric(str_����) And Len(str_����) > 13 Then
                        str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                    End If
                    str_��� = str_��� & "|" & Replace(str_��Ŀ, "^", "") & "|" & Trim(Split(str_Line, "|")(3))
                End If
            End If
        Next
        
        If str_�걾�� <> "" And str_��� <> "" And str_���� <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_���
            str_���� = "": str_��� = ""
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    
End Sub

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







