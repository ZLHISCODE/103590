VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_sysmex_CS2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  sysmex CS-2000i
'ҽԺ:  ������ҽԺ
'����:  lwf
'ʱ��:  2014/11/24

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
    '   ΢��������ʽ ͨ����|�������^��ʽ^���ֵ
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim str_Code As String
    Dim str_Line As String               '��ǰҪ�������ִ�
    Dim lng_S As Long                  '��ʼλ��
    Dim lng_E As String                '����λ��
    Dim varLine As Variant
    Dim ii As Long
    Dim dec_bd As Double
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String

    mstrCode = strSource
    On Error GoTo errH
    strResult = ""

    lng_S = InStr(strSource, "")
    lng_E = InStr(lng_S, strSource, "")

    If lng_S > 0 And lng_E > 0 And lng_S < lng_E Then
        varLine = Split(mstrCode, "M")
        str_���� = "": str_�걾�� = "": str_��Ŀ = ""
        str_���� = Mid(Split(varLine(0), "U")(1), 1, 6)
        str_���� = "20" & Mid(str_����, 1, 2) & "-" & Mid(str_����, 3, 2) & "-" & Mid(str_����, 5, 2)
        str_�걾�� = Val(Right(varLine(0), 4))

        For ii = 0 To 10 Step 1
            If ii = 0 Then
                dec_bd = 0.1
            ElseIf ii = 3 Then
                dec_bd = 0.01
            ElseIf ii = 5 Then
                dec_bd = 0.1
            ElseIf ii = 7 Then
                dec_bd = 0.001
            ElseIf ii = 8 Then
                dec_bd = 0.1
            ElseIf ii = 10 Then
                dec_bd = 0.01
            Else
                dec_bd = 1
            End If
            str_��Ŀ = str_��Ŀ & "|" & Mid(Trim(varLine(1)), ii * 9 + 1, 3) & "|" & Val(Trim(Mid(Trim(varLine(1)), ii * 9 + 4, 6))) * dec_bd
        Next
        If Val(str_�걾��) > 0 Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0" & str_��Ŀ
        End If
    Else
        strReserved = mstrCode
    End If

    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function