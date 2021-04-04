VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ULT_1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:������1200
'����:
'�û�:³����ҽԺ
Dim mstrCode As String

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

    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long

    Dim str_Segment As String
    Dim var_Filed As Variant

    Dim lng_S As Long
    Dim lng_E As Long

    Dim str_�걾�� As String, str_���� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If

    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))

    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)

        var_Code = Split(str_Code, Chr(&HD))

        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 1, 3) = "OBR" Then
                '�걾��
                var_Filed = Split(str_Segment, "|")
                str_�걾�� = var_Filed(3)
                str_���� = var_Filed(7)
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2)
            ElseIf Mid(str_Segment, 1, 3) = "OBX" Then
                var_Filed = Split(str_Segment, "|")
                If Trim(var_Filed(2)) = "NM" Then
                    If var_Filed(3) Like "*^*" Then
                        str_��Ŀ = str_��Ŀ & "|" & Split(var_Filed(3), "^")(0) & "|" & var_Filed(5)
                    Else
                        str_��Ŀ = str_��Ŀ & "|" & var_Filed(3) & "|" & var_Filed(5)
                    End If
                End If
            End If
        Next

        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then

            str_��Ŀ = Replace(str_��Ŀ, "mmol/L", "")
            str_��Ŀ = Replace(str_��Ŀ, "Cell/uL", "")
            str_��Ŀ = Replace(str_��Ŀ, "umol/L", "")
            str_��Ŀ = Replace(str_��Ŀ, "g/L", "")

            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |��Һ|0" & str_��Ŀ
        End If
        str_��Ŀ = ""
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "DIRUIH_FUS200_analyse", strSource, Err.Number & " " & Err.Description
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