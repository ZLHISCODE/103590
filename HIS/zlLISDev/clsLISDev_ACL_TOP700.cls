VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL_TOP700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:������ACL700TOPѪ����
'�û�:�����ﶫ������ҽԺ
'����:hhd
'����
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)

    Dim lng_S As Long    '��ʼ
    Dim lng_E As Long    '����
    Dim lng_SS As Long, lng_ss1 As Long
    'һ�������ļ�¼
    Dim str_�걾�� As String, str_���� As String, str_�����˱걾�ʿ� As String, str_������Ŀ As String, str_������ As String
    Dim str_Line As String, strTmp As String
    Dim lng_�걾�� As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_��� As String
    str_�����˱걾�ʿ� = " |ѪҺ|0"
    On Error GoTo errH
    '-----Ӧ��
    strSource = Replace(strSource, Chr(&H4) & Chr(&H5), Chr(&H5))
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
        '    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        '        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If strSource <> "" Then strCmd = Chr(&H6)

    mstrCode = mstrCode & strSource
    '---ȡ����������
    lng_S = InStr(mstrCode, Chr(&H5))   '��&H5��ͷ��������������Ŀ�ʼλ��
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '��&H4������������������Ľ���λ��
    Else
        lng_E = 0
    End If


    Do While lng_S > 0 And lng_E > 0    '�����ʼ����������Ϊ��
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)    '���شӿ�ʼλ�õ�����λ��֮����ַ�
        mstrCode = Mid(mstrCode, lng_E + 1)
        Do While InStr(str_Code, Chr(&H17)) > 0
            lng_SS = InStr(str_Code, Chr(&H17))
            str_Code = Mid(str_Code, 1, lng_SS - 1) & Mid(str_Code, lng_SS + 7)    'ȡ7λУ����
        Loop

        var_data = Split(str_Code, Chr(&HD))    'ȡÿ������
        For i = LBound(var_data) To UBound(var_data)
            str_Line = var_data(i)

            If InStr(str_Line, "H|") > 0 Then
                str_���� = Trim(Split(str_Line, "|")(13))    '��ȡ����ʱ��
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid( _
                         str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, _
                                                                             2) & ":" & Mid(str_����, 13, 2)    '������ڳ�����yyyy-mm-dd��
                str_���� = Format(str_����, "yyyy-MM-dd hh:mm:ss")
                If Not IsDate(str_����) Then    '�������Ϊ�գ�����ʾΪ��ǰ����
                    str_���� = Format(Now, "yyyy-MM-dd hh:mm:ss")
                End If
            ElseIf InStr(str_Line, "O|") = 4 Then
                str_�걾�� = Trim(Split(str_Line, "|")(2))    '�ӡ�2P|1|||||�� ��|���1��ʼ��ʼȡ
            ElseIf InStr(str_Line, "R|") = 4 Then    'ȡ����ļ�����Ŀ

                strTmp = Split(str_Line, "|")(4)
                strTmp = Replace$(strTmp, "%", "1")
                strTmp = Replace$(strTmp, "/", "2")
                strTmp = Replace$(strTmp, "", "3")  ' �����ո�
                str_������Ŀ = Trim(Split(str_Line, "|")(2)) & strTmp
                str_������Ŀ = Replace$(str_������Ŀ, "^^^", "")
                'n = n + 1
                str_������ = Trim(Split(str_Line, "|")(3))

                '�����ƴ������
                If Val(str_�걾��) > 0 And str_������Ŀ <> "" And str_������ <> "" Then
                    str_��� = str_��� & "||" & str_���� & "|" & str_�걾�� & "|" & str_�����˱걾�ʿ� & "|" & str_������Ŀ & "|" & str_������
                End If

            End If

        Next

        lng_S = InStr(mstrCode, Chr(&H5))   '��&H5��ͷ��������������Ŀ�ʼλ��
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '��&H4������������������Ľ���λ��
        Else
            lng_E = 0
        End If
    Loop

    strResult = Mid(str_���, 3)
    Exit Sub
errH:
    WriteLog "ALC_TOP700", strSource, Err.Description
End Sub

Public Function SendSample(ByVal strSample As String) As String
    SendSample = ""
End Function




