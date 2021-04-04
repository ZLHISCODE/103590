VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_Roche_ZLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'仪器:罗氏朱莉亚尿机
'用户:广西容县人民医院
'作者:hhd

Option Explicit

Public Function GetStartCmd() As String
'设备标准接口
'功能：获取设备开始发送命令
'GetStartCmd = ""
    GetStartCmd = Chr(&H6)    ' Chr(6)
End Function

Public Function GetEndCmd() As String
'设备标准接口
'功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
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

    Dim iStrPoint As Integer, iStrPointBK As Integer    '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer    '当前这组结果的入口字符位置

    Dim blnReadItem As Boolean    '当前是否正在读取检验项目
    Dim blnEndSample As Boolean    '标本结束

    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer
    Dim strNowData As String        '当前能解析的完整的字串
    Dim aLine() As String
    Dim intLoop As Integer
    Dim str_标本号 As String
    Dim str_时间 As String
    Dim str_time As String
    Dim str_日期 As String
    Const strStartChr As String = ""
    Const strEndChr As String = ""
    
    On Error GoTo errHandle
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False

    '没有找到开始符和结束符时退出
    If InStr(strSource, strStartChr) = 0 Or InStr(strSource, strEndChr) = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If

    strNowData = Mid(strSource, InStr(strSource, strStartChr), InStr(strSource, strEndChr))
    strReserved = Mid(strSource, InStr(strSource, strEndChr) + 1)
    '解析
    aLine = Split(strNowData, vbCrLf)

    '标本号
    str_标本号 = Val(Trim(Mid(aLine(1), InStr(aLine(1), "Seq.no.") + 7)))

    '时间
    str_时间 = "20" & Mid(aLine(2), 1, 2) & "-" & Mid(aLine(2), 4, 2) & "-" & Mid(aLine(2), 7, 2)
    str_time = Mid(Trim(aLine(2)), 20) & ":00"    'Trim(aLine(2))
    str_日期 = str_时间 & " " & str_time
    If Not IsDate(str_日期) Then
        str_时间 = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Else
        str_时间 = Format(CDate(str_日期), "yyyy-mm-dd hh:mm:ss")
    End If
    strData = str_日期 & "|" & str_标本号 & "| |尿液|0"

    For intLoop = 3 To UBound(aLine)
        If Trim(aLine(intLoop)) <> "" And Trim(aLine(intLoop)) <> "" And InStr(aLine(intLoop), "Color:") = 0 Then
            strData = strData & "|" & Trim(Replace(Mid(aLine(intLoop), 1, InStr(aLine(intLoop), ":") - 1), "*", "")) & "|" & Trim(Mid(aLine(intLoop), InStr(aLine(intLoop), ":") + 1))
            'strData = strData & "|" & Replace(Replace(Replace(Trim(Mid(aLine(intLoop), 1, 6)), "*", ""), " ", ""), ":", "") & "|" & Trim$(Mid(aLine(intLoop), 7))
        End If
    Next

    strData = Replace(strData, "Brown Yellow", "棕黄色")
    strData = Replace(strData, "Thin Yellow", "淡黄色")
    strData = Replace(strData, "Yellow", "黄色")
    strData = Replace(strData, "Green", "绿色")
    strData = Replace(strData, "Primrose", "淡黄色")
    strData = Replace(strData, "Red", "红色")
    strData = Replace(strData, "Orange", "橙色")
    strData = Replace(strData, "Blue", "蓝色")
    strData = Replace(strData, "Brown", "棕色")
    strData = Replace(strData, "Unselected", "未选")
    strData = Replace(strData, "Other", "其它")
    strData = Replace(strData, "Filemot", "黄褐色")
    strData = Replace(strData, "Little Dark", "稍模糊")
    strData = Replace(strData, "Dark", "模糊")
    strData = Replace(strData, "Clear", "清澈")
    strData = Replace(strData, "Feculent", "浑浊")
    '    strData = Replace(strData, "Negative", "阴性")
    'strData = Replace(strData, "Normal", "正常")
    '    strData = Replace(strData, "Positive", "阳性")
    '    strData = Replace(strData, "Small", "1+")
    '    strData = Replace(strData, "Moderate", "2+")
    '    strData = Replace(strData, "Large", "3+")
    '    strData = Replace(strData, "Trace", "微量")
    '    strData = Replace(strData, "Thin", "浅黄色")
    '    strData = Replace$(strData, "Cells/uL", "")
    '    strData = Replace$(strData, "g/L", "")
    '    strData = Replace$(strData, "/ul", "")
    '    strData = Replace$(strData, "mmol/L", "")
    'strData = Replace$(strData, " ", "")

    If Len(strData) > 0 Then
        strResult = strData
    End If
    'strCmd = Chr(&H6&)
    Exit Sub
errHandle:
    WriteLog "Roche_ZLR", "Roche_ZLR", Err.Number & " " & Err.Description
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



