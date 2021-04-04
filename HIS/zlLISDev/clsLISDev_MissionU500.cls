VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MissionU500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Mission U500
'厂商:艾康
'作者:罗云龙
'用户:新疆阿勒泰地区人民医院
'时间:2014-05-29

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
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strLine As String
    Dim iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H3&)) > 0 Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, Chr(&H3)) - InStr(strReserved, Chr(&H2))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H2)), lngLength + 3)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)

            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = ""
            '解析
            iCount = 0
            varLine = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = 0 To UBound(varLine)

                Select Case iCount
                    Case 1 '日期
                    str日期 = Mid(Trim(varLine(iRow)), 1, 4) & "-" & TrunMonth(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 1, 3)) & "-" & Mid(Trim(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 4)), 1, 2) & " " & Trim(Mid(Trim(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 4)), 3))
                    Case 3 '标本号
                    If InStr(varLine(iRow), ":") > 0 Then varLine(iRow) = Replace(varLine(iRow), ":", "")
                    str标本号 = Val(Trim(Left(Mid(varLine(iRow), InStr(varLine(iRow), "-") + 1), 6)))
                    str检验人标本质控 = " |尿液|0"
                    Case 4 To 14
                    strItem = strItem & "|" & Trim(MidUni(varLine(iRow), 2, 3)) & "|" & Trim(MidUni(varLine(iRow), 5))
                End Select
                If varLine(iRow) <> "" Then iCount = iCount + 1
            Next
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = IIf(strResult <> "", strResult & "||", "") & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If InStr(strReserved, Chr(&H3&)) = 0 Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
    Close #lngErrFileNo
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function MidUni(ByVal strTemp As String, ByVal Start As Long, Optional ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    If Length = 0 Then Length = LenB(strTemp)
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
    MidUni = Trim(Replace(MidUni, Chr(0), ""))
    MidUni = Trim(Replace(MidUni, "0.56mmol/L", ""))
    MidUni = Trim(Replace(MidUni, "3.5umol/L", ""))
    MidUni = Trim(Replace(MidUni, "neg", ""))
    MidUni = Trim(Replace(MidUni, "g/L", ""))
    MidUni = Trim(Replace(MidUni, "mg/dL", ""))
    MidUni = Trim(Replace(MidUni, "Leu/uL", ""))
End Function

Private Function TrunMonth(ByVal strMonth As String) As String
    '转换英文的月份为数字
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function
