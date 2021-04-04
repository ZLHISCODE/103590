VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_TECAN_F50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstrCode As String

'厂商：瑞士帝肯
'型号：酶标仪TECAN，型号：F50
'用户：都安人民医院
'接口方式：监测.asc文件
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
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    '仪器数据以H02开始H03结束()

    Dim str_S As Long, str_E As Long, str_SS As Long
    Dim str_Code As String, strLine As String, varLine As Variant
    Dim VAR_DAT As Variant, aItem As Variant
    Dim i As Integer, intLoop As Integer
    Dim str_time As String
    Dim str日期 As String, str标本号 As String, str检验结果 As String, strData As String, str项目 As String, str通道码 As String
    Dim str_Line As String

    str日期 = ""
    str标本号 = ""
    str项目 = ""
    str检验结果 = ""
    strData = ""

    '---数据缓冲
    mstrCode = mstrCode & strSource
    str_S = 0
    str_E = 0

    '对仪器传输过来的数据进行应答

    strCmd = ""

    '判断字符串的完整性
    str_S = InStr(mstrCode, Chr(&H2))

    If str_S > 0 Then
        str_E = InStr(str_S, mstrCode, Chr(&H3))
    Else
        str_E = 0
    End If

    If str_S > 0 And str_E > 0 And Len(mstrCode) > 10 Then

        Do While str_S > 0 And str_E > 0 '判断一个数据中有多个开始和结束的开始

            str_Code = Mid(mstrCode, str_S, str_E - str_S)
            mstrCode = Mid(mstrCode, str_E + 1)
            'HOA分段
            VAR_DAT = Split(str_Code, Chr(&HA))
            '文件没有数据，跳出
            If InStr(str_Code, "-") = 0 Then
                Exit Do
            End If
            '判断点位符个数，防止不出标本号的时候解析错误
            If InStr(str_Code, vbTab) > 0 And Len(str_Code) - Len(Replace(str_Code, vbTab, "")) > 2 Then
                '取通道码
                If InStr(str_Code, ".mth") > 0 Then
                    str项目 = Replace(Right(Split(str_Code, ".m")(0), 5), vbCrLf, "")
                End If

                For i = LBound(VAR_DAT) To UBound(VAR_DAT) '最后一段为空，避免下标越界
                    strLine = VAR_DAT(i)
                    If Len(strLine) - Len(Replace(strLine, vbTab, "")) > 3 Then
                        str_Line = trimString(strLine)

                        str_time = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        str日期 = str_time

                        If Len(str_Line) - Len(Replace(str_Line, vbTab, "")) > 2 Then
                            str标本号 = Val(Split(str_Line, vbTab)(2))
                        End If

                        If str标本号 <> 0 Then
                            str检验结果 = Replace(Replace(Replace(Replace(Trim(Split(str_Line, vbTab)(0)), "?", ""), "", "*"), "-", "阴性"), "+", "阳性") & "^0^0^" & Trim(Split(str_Line, vbTab)(3))
                            strData = IIf(Trim(strData) = "", "", strData & "||") & str日期 & "|" & str标本号 & "| |血液|0" & "|" & Trim(str项目) & "|" & Replace(Trim(str检验结果), "", "")
                        End If
                    End If
                Next
                '循环条件判断
                str_S = InStr(mstrCode, Chr(&H2))
                If str_S > 0 Then
                    str_E = InStr(str_E, mstrCode, Chr(&H3))
                Else
                    str_E = 0
                End If

            End If
        Loop
    End If
    '向仪器返回结果
    strResult = strData
    strReserved = ""
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

Private Function trimString(str As String) As String
        Dim strTemp       As String
        
        strTemp = Trim(str)
        Do While InStr(1, strTemp, Space(2)) > 0
                strTemp = Replace(strTemp, Space(2), Space(1))
        Loop
        trimString = strTemp
End Function
