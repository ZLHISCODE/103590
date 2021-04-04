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

'河科大一附院
'C_ZY
'血液鉴定仪
Private mstr_上次结果 As String
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
    Dim str_标本号 As String, str_日期 As String, str_标本 As String
    Dim sz
    Dim sz_分段
    Dim i As Integer
    Dim str_结果 As String
    Dim mdata As String
    Dim j As Integer
    Dim str_全部 As String
    mdata = mdata & strSource

    If mdata Like "MSH*" Then

        sz = Split(mdata, Chr(13))
        str_日期 = Mid(Split(sz(0), "|")(6), 1, 4) & "-" & Mid(Split(sz(0), "|")(6), 5, 2) & "-" & Mid(Split(sz(0), "|")(6), 7, 2) & " " & Mid(Split(sz(0), "|")(6), 9, 2) & ":" & Mid(Split(sz(0), "|")(6), 11, 2)
        For j = 1 To UBound(sz) - 1
            If Mid(sz(j), 2) > Chr(0) Then
                If Split(sz(j), "|")(0) = "OBR" Then
                    str_标本号 = Split(sz(j), "|")(2)
                    str_结果 = str_结果 & "||" & str_日期 & "|" & str_标本号 & "| |血液|0"
                End If
                If Split(sz(j), "|")(0) = "OBX" And Split(sz(j), "|")(3) = "ABDR" Then
                    If InStr(1, Split(sz(j), "|")(5), "/") > 0 Then
                        '把ABO分成ABOZ和ABOF,但是截取结果相同,解决导航台多个指标相同通道码,只能一个接收到.
                        str_结果 = str_结果 & "|ABOZ|" & Split(Split(sz(j), "|")(5), "/")(0) & "|ABOF|" & Split(Split(sz(j), "|")(5), "/")(0) & "|RH|" & Split(Split(sz(j), "|")(5), "/")(1)
                    Else
                        str_结果 = str_结果 & "|ABOZ|" & Split(sz(j), "|")(5) & "|ABOF|" & Split(sz(j), "|")(5)
                    End If
                End If
                If Split(sz(j), "|")(0) = "OBX" And Split(sz(j), "|")(3) = "AS" Then
                    str_结果 = str_结果 & "|AS|" & Split(sz(j), "|")(5)
                End If
            End If
        Next
        If str_结果 <> "" Then
            str_结果 = Replace(Replace(str_结果, "Positive", "阳性"), "Negative", "阴性")
            strResult = Mid(str_结果, 3)
            mdata = ""
            str_结果 = ""
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
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function
