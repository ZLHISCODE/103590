VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IMS972_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:����ʷ����� IMS 972
'����:
'����:����
'�û�:
'�޸ģ�����
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
    On Error GoTo hErr '�򿪴�������
    '------------------------------------------------
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_Line As String
    Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
    Dim intLoop As Integer
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "Sample")
    lng_E = InStr(mstrCode, "End")
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid$(mstrCode, lng_S, lng_E - lng_S + 3) '���������һ��ѭ��
        mstrCode = Mid$(mstrCode, lng_E + 3)
        var_Code = Split(str_Code, Chr$(&H20) & Chr$(&H20))
        For intLoop = LBound(var_Code) To UBound(var_Code)
            If intLoop = 2 Then '��ȡ����
                str_���� = Format$(var_Code(intLoop), "yyyy-mm-dd")
            End If
            If intLoop = 4 Then '��ȡʱ��
                str_���� = Format$(str_���� & " " & Format$(var_Code(intLoop), "hh:mm:ss"), "yyyy-mm-dd HH:mm:ss")
            End If
            If intLoop >= 5 And Trim$(var_Code(intLoop)) <> "" Then
                str_��Ŀ = str_��Ŀ & "|" & Trim$(var_Code(intLoop))
                If InStr(Trim$(var_Code(intLoop)), "PAT") > 0 Then
                    str_�걾�� = Mid$(str_Code, InStr(str_Code, "PAT") + 10, 3)
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|0" & str_��Ŀ
            str_��Ŀ = ""
            If Right$(strResult, 4) = "|End" Then
                strResult = Mid$(strResult, 1, Len(strResult) - 4)
            End If
        End If
        lng_S = InStr(mstrCode, "Sample")
        lng_E = InStr(mstrCode, "End")
    Loop
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_IMS972_02"
    Close #lngErrFileNo
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