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
'仪器:BS800M
'厂商:深圳迈瑞
'仪器:KX21
'作者:
'用户:
'时间:20140307

Private mstrCode As String '缓存接收到的数据

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo hErr '打开错误陷阱
    '------------------------------------------------
    Dim lngS As Long, lngE As Long, i As Long, J As Long
    Dim strDate As String, str As String
    Dim strCode() As String, str日期 As String, str标本号 As String, str项目 As String, str结果 As String
    Dim str检验结果 As String
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

                    str日期 = Mid$(strDate, 1, 4) & "-" & Mid$(strDate, 5, 2) & "-" & Mid$(strDate, 7, 2) & " " & Mid$(strDate, 9, 2) & ":" & Mid$(strDate, 11, 2) & ":" & Mid$(strDate, 13, 2)

                    str标本号 = Split(Split(strCode(J), "OBR")(1), "|")(3)
                    str检验结果 = ""

                    var = Split(strCode(J), "OBX")
                    For i = LBound(var) + 1 To UBound(var)
                        If UBound(Split(var(i), "|")) >= 5 Then
                            str项目 = Split(var(i), "|")(4)
                            str结果 = Split(var(i), "|")(5)
                            str检验结果 = str检验结果 & str项目 & "|" & str结果 & "|"
                        End If
                    Next i
                    If Right$(str检验结果, 1) = "|" Then
                        str检验结果 = Mid$(str检验结果, 1, Len(str检验结果) - 1)
                    End If
                    If Trim$(str检验结果) <> "" Then
                        strResult = IIf(strResult = "", str日期 & "|" & str标本号 & "| |血清|0|" & str检验结果, _
                                    strResult & "||" & str日期 & "|" & str标本号 & "| |血清|0|" & str检验结果)
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
