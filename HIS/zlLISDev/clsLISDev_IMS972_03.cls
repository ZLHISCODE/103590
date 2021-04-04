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

'仪器:电解质分析仪 IMS 972
'厂商:
'作者:敬秦
'用户:
'修改：敬秦
Private mstrCode As String

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_Line As String
    Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
    Dim intLoop As Integer
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "Sample")
    lng_E = InStr(mstrCode, "End")
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid$(mstrCode, lng_S, lng_E - lng_S + 3) '缓存进入下一次循环
        mstrCode = Mid$(mstrCode, lng_E + 3)
        var_Code = Split(str_Code, Chr$(&H20) & Chr$(&H20))
        For intLoop = LBound(var_Code) To UBound(var_Code)
            If intLoop = 2 Then '读取日期
                str_日期 = Format$(var_Code(intLoop), "yyyy-mm-dd")
            End If
            If intLoop = 4 Then '读取时间
                str_日期 = Format$(str_日期 & " " & Format$(var_Code(intLoop), "hh:mm:ss"), "yyyy-mm-dd HH:mm:ss")
            End If
            If intLoop >= 5 And Trim$(var_Code(intLoop)) <> "" Then
                str_项目 = str_项目 & "|" & Trim$(var_Code(intLoop))
                If InStr(Trim$(var_Code(intLoop)), "PAT") > 0 Then
                    str_标本号 = Mid$(str_Code, InStr(str_Code, "PAT") + 10, 3)
                End If
            End If
        Next
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_项目
            str_项目 = ""
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
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function
