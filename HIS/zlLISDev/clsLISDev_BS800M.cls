VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BS800M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BS800M
'����:��������
'����:KX21
'����:
'�û�:
'ʱ��:20140307

Private mstrCode As String '������յ�������

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
    Dim lngS As Long, lngE As Long, i As Long, J As Long
    Dim strDate As String, str As String
    Dim strCode() As String, str���� As String, str�걾�� As String, str��Ŀ As String, str��� As String
    Dim str������ As String
    Dim var() As String
    mstrCode = mstrCode & strSource
    If InStr(mstrCode, "MSH|^~\&") > 0 Then
        strCode = Split(mstrCode, "MSH|^~\&")
        For J = LBound(strCode) To UBound(strCode)
            If Trim$(strCode(J)) <> "" Then
                strCode(J) = "MSH|^~\&" & strCode(J)
                If UBound(Split(strCode(J), "|")) >= 6 Then
                    '                   strDate = Split(strCode(J), "|")(6)
                    strDate = Split(Split(strCode(J), "OBR")(1), "|")(6)

                    str���� = Mid$(strDate, 1, 4) & "-" & Mid$(strDate, 5, 2) & "-" & Mid$(strDate, 7, 2) & " " & Mid$(strDate, 9, 2) & ":" & Mid$(strDate, 11, 2) & ":" & Mid$(strDate, 13, 2)

                    str�걾�� = Split(Split(strCode(J), "OBR")(1), "|")(3)
                    str������ = ""

                    var = Split(strCode(J), "OBX")
                    For i = LBound(var) + 1 To UBound(var)
                        If UBound(Split(var(i), "|")) >= 5 Then
                            str��Ŀ = Split(var(i), "|")(4)
                            str��� = Split(var(i), "|")(5)
                            str������ = str������ & str��Ŀ & "|" & str��� & "|"
                        End If
                    Next i
                    If Right$(str������, 1) = "|" Then
                        str������ = Mid$(str������, 1, Len(str������) - 1)
                    End If
                    If Trim$(str������) <> "" Then
                        strResult = IIf(strResult = "", str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str������, _
                                    strResult & "||" & str���� & "|" & str�걾�� & "| |Ѫ��|0|" & str������)
                    End If
                End If
            End If
        Next J
    End If
    '------------------------------------------------
    Exit Sub
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_BS800M"
    Close #lngErrFileNo
End Sub