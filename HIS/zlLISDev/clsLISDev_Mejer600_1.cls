VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mejer600_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'仪器:Mejer600
'医院:XXXXXXXX
'作者:曾超
'完成时间:20090710

Option Explicit

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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim blnEndSample As Boolean '标本结束
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim aDatas() As String, strItemResult As String, i As Integer
    Dim strNow As String        '当前能解析的完整的字串
    Dim aLine() As String
    Dim intLoop As Integer
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    
    '没有找到开始结束位置时退出
    If InStr(strSource, "") = 0 Or InStr(strSource, "") = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    strReserved = Mid(strSource, InStr(strSource, "") + 1)
    strNow = Mid(strSource, InStr(strSource, "") + 1, InStr(strSource, "") - 1)
    
    aLine = Split(strNow, vbCrLf)
    
    '时间
    strData = Format(Trim(Mid(aLine(1), 8)) & " " & Trim(aLine(2)), "yyyy-mm-dd hh:mm:ss") & _
                "|" & Val(Mid(aLine(1), 5, 3)) & "| |尿液|0"
    
    
    For intLoop = 4 To UBound(aLine)
        If Trim(aLine(intLoop)) <> "" And Trim(aLine(intLoop)) <> "" Then
            strData = strData & "|" & Trim(Mid(aLine(intLoop), 1, 4)) & "|" & Val(Mid(aLine(intLoop), 9)) & _
                IIf(Trim(Mid(aLine(intLoop), 6, 3)) = "", "", "(" & Trim(Mid(aLine(intLoop), 6, 3)) & ")")
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
    strData = Replace(strData, "Negative", "阴性")
    strData = Replace(strData, "Normal", "正常")
    strData = Replace(strData, "Positive", "阳性")
    strData = Replace(strData, "Small", "1+")
    strData = Replace(strData, "Moderate", "2+")
    strData = Replace(strData, "Large", "3+")
    strData = Replace(strData, "Trace", "微量")
    strData = Replace(strData, "Thin", "浅黄色")
    
    If Len(strData) > 0 Then
        strResult = strData
    End If
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

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos > 0 Then
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 2
    Else
        iPos = InStr(iPoint, strSource, Chr(13))
        If iPos = 0 Then iPos = InStr(iPoint, strSource, Chr(10))
        If iPos = 0 Then iPos = Len(strSource) + 1
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function



