VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IMS972"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:����ʷ����� IMS 972
'����:
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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_Line As String
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim intLoop As Integer
    
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "PAT_ID")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "PH")
        If lng_E > 0 Then
            lng_E = InStr(lng_E, mstrCode, Chr(&HD))
        End If
    End If
        
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------------
        var_Code = Split(str_Code, Chr(&HD))
        str_���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
        For intLoop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(intLoop)
            If Trim(Left(str_Line, 6)) = "PAT_ID" Then
                str_�걾�� = Val(Mid(str_Line, 7))
                str_��Ŀ = ""
            ElseIf InStr("K,Na,Cl,nCa,iCa,TCa,TCO2,AG,PH", Trim(Left(str_Line, 6))) > 0 Then
                str_��� = Trim(Replace(Mid(str_Line, 7), "mmol/L", ""))
                If str_��� <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Trim(Left(str_Line, 6)) & "|" & str_���
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|0" & str_��Ŀ
        End If
        '----------------------------------------------
        lng_S = InStr(mstrCode, "PAT_ID")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "PH")
            If lng_E > 0 Then
                lng_E = InStr(lng_E, mstrCode, Chr(&HD))
            End If
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