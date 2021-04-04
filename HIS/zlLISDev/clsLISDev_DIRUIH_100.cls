VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DIRUIH_100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ��Һ������
'ҽԺ:  �����зο�ҽԺ
'����:  �¶�

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
    Dim var_Code As Variant, i As Integer, str_Line As String
    Dim str_�걾�� As String, str_���� As String, str_Tmp As String, str_��Ŀ As String
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        str_��Ŀ = "": str_Tmp = "": str_�걾�� = "": str_���� = ""
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If Trim(Left(str_Line, 4)) = "No." Then
                str_�걾�� = Mid(str_Line, 5, 3)
                str_���� = Mid(str_Line, 10)
            ElseIf InStr(",WBC,KET,URO,BIL,PRO,BLD,SG,Vc,GLU,NIT,pH,", "," & Trim(Left(str_Line, 4)) & ",") > 0 Then
                str_Tmp = Mid(str_Line, 5)
                str_Tmp = Replace(str_Tmp, "mmol/L", "")
                str_Tmp = Replace(str_Tmp, "umol/L", "")
                str_Tmp = Replace(str_Tmp, "Cell/uL", "")
                str_Tmp = Replace(str_Tmp, "g/L", "")
                str_Tmp = Trim(str_Tmp)
                If str_Tmp <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Left(str_Line, 4)) & "|" & str_Tmp
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_���� <> "" And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |��Һ|0" & str_��Ŀ
        End If
        '----------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
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