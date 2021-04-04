VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BioxunSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 

Option Explicit

'�ӿƴ�һ��Ժ
'C_ZY
'ѪҺ������
Private mstr_�ϴν�� As String
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    On Error GoTo ErrHand '�򿪴�������
    '------------------------------------------------
    Dim str_�걾�� As String, str_���� As String, str_�걾 As String
    Dim sz
    Dim sz_�ֶ�
    Dim i As Integer
    Dim str_��� As String
    Dim mdata As String
    Dim j As Integer
    Dim str_ȫ�� As String
    mdata = mdata & strSource

    If mdata Like "MSH*" Then

        sz = Split(mdata, Chr(13))
        str_���� = Mid(Split(sz(0), "|")(6), 1, 4) & "-" & Mid(Split(sz(0), "|")(6), 5, 2) & "-" & Mid(Split(sz(0), "|")(6), 7, 2) & " " & Mid(Split(sz(0), "|")(6), 9, 2) & ":" & Mid(Split(sz(0), "|")(6), 11, 2)
        For j = 1 To UBound(sz) - 1
            If Mid(sz(j), 2) > Chr(0) Then
                If Split(sz(j), "|")(0) = "OBR" Then
                    str_�걾�� = Split(sz(j), "|")(2)
                    str_��� = str_��� & "||" & str_���� & "|" & str_�걾�� & "| |ѪҺ|0"
                End If
                If Split(sz(j), "|")(0) = "OBX" And Split(sz(j), "|")(3) = "ABDR" Then
                    If InStr(1, Split(sz(j), "|")(5), "/") > 0 Then
                        '��ABO�ֳ�ABOZ��ABOF,���ǽ�ȡ�����ͬ,�������̨���ָ����ͬͨ����,ֻ��һ�����յ�.
                        str_��� = str_��� & "|ABOZ|" & Split(Split(sz(j), "|")(5), "/")(0) & "|ABOF|" & Split(Split(sz(j), "|")(5), "/")(0) & "|RH|" & Split(Split(sz(j), "|")(5), "/")(1)
                    Else
                        str_��� = str_��� & "|ABOZ|" & Split(sz(j), "|")(5) & "|ABOF|" & Split(sz(j), "|")(5)
                    End If
                End If
                If Split(sz(j), "|")(0) = "OBX" And Split(sz(j), "|")(3) = "AS" Then
                    str_��� = str_��� & "|AS|" & Split(sz(j), "|")(5)
                End If
            End If
        Next
        If str_��� <> "" Then
            str_��� = Replace(Replace(str_���, "Positive", "����"), "Negative", "����")
            strResult = Mid(str_���, 3)
            mdata = ""
            str_��� = ""
        End If
    End If
    '------------------------------------------------
    Exit Sub
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Sub
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