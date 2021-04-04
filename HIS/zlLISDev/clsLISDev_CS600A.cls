VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS600A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����CS600A
'����:��������
'�û�:�±�����
'����:hhd
'ʹ�÷���:���⴮�ڽ���
'ʱ�䣺2013/12/19
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
    '��׼���������ļ���ȡ������
    '   strFile������·�����ļ���
    '   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    '   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    '   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
    '   ���صļ���������(������������밴���±�׼��֯���)
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������

    Dim lng_S As Long    '��ʼ
    Dim lng_E As Long    '����
    Dim lng_SS As Long, lng_ss1 As Long
    'һ�������ļ�¼
    Dim str_�걾�� As String, str_�걾��1 As String, str_���� As String, str_�����˱걾�ʿ� As String, str_������Ŀ As String, str_������ As String
    Dim str_������1 As String
    Dim str_Line As String
    Dim lng_�걾�� As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_��� As String
    str_�����˱걾�ʿ� = " |Ѫ��|0"
    
    On Error GoTo errH
    
    '-----Ӧ��
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6&)
    End If
    If InStr(strSource, Chr(&HD) & Chr(&HA)) > 0 Then
        strCmd = Chr(&H6&)
    End If

    If strSource <> "" Then mstrCode = mstrCode & strSource

    '---ȡ����������
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
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

            '��ȡ����ʱ��
            If InStr(str_Line, "H|\") > 0 Then    'H|\
                str_���� = Trim(Split(str_Line, "|")(12))
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid( _
                         str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, _
                                                                             2) & ":" & Mid(str_����, 13, 2)    '������ڳ�����yyyy-mm-dd��
                str_���� = Format(str_����, "yyyy-MM-dd hh:mm:ss")

                If Not IsDate(str_����) Then    '�������Ϊ�գ�����ʾΪ��ǰ����
                    str_���� = Format(Now, "yyyy-MM-dd hh:mm:ss")
                End If

                'ȡ�걾��
            ElseIf InStr(str_Line, "O|") = 4 Then
                str_�걾��1 = Trim(Split(str_Line, "|")(2))
                str_�걾�� = Trim(Val(Split(str_�걾��1, "^")(1)))

                'ȡ����ļ�����Ŀ�ͼ�����
            ElseIf InStr(str_Line, "R|") = 4 Then
                str_������Ŀ = Replace(Replace(Trim(Split(str_Line, "|")(2)), "^^^", ""), "��ϸ��", "rbc")
                str_������1 = Trim(Split(str_Line, "|")(3))
                str_������ = IIf(str_������1 = "", "*.**", str_������1)

                If str_�걾�� > 0 And str_������Ŀ <> "" Then
                    str_��� = IIf(Trim(str_���) = "", "", str_��� & "||") & str_���� & "|" & str_�걾�� & "|" & str_�����˱걾�ʿ� & "|" & str_������Ŀ & "|" & str_������
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
    If str_��� <> "" Then strResult = Mid(str_���, 3)
    Exit Sub
errH:
    WriteLog "CS600A", "CS", Err.Number & ":" & Err.Description
End Sub

