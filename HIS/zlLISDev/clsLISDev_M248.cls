VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_M248"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:M248 血气分析仪
'厂商:德国拜尔(Bayer)
'作者:陈东
'用户:重庆市南桐医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim varline As Variant, strItem As String, strItem1 As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Const strStartChr As String = "---- Bayer Diagnostics  248 ----"
    Const strEndChr As String = "--------------------------------"
    Dim strNowData As String
    Dim strData As String
    Dim intloop As Integer
    Dim strLineData As String
    Dim aItem As Variant
    
    Dim lng_E As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    '没有找到开始符和结束符时退出
    
    If InStr(strSource, strStartChr) = 0 Or InStr(strSource, strEndChr) = 0 Then
        strReserved = strSource
    End If
    
    '取得当前一段，剩下的返回。下次再解析。
    lng_E = InStr(strSource, strEndChr)
    lng_E = InStr(lng_E, strSource, "")
    strNowData = Mid(strSource, InStr(strSource, strStartChr), lng_E - InStr(strSource, strStartChr) + 1)
    strReserved = Mid(strSource, lng_E + 1)
    
    '开始解析
    varline = Split(strNowData, vbCrLf)
    If UBound(varline) < 27 Then Exit Sub
    
    '日期
    strLineData = DelSapce(CStr(varline(2)))
    aItem = Split(strLineData, Space(1))
    strData = Format(aItem(4) & "-" & aItem(3) & "-" & Replace(aItem(2), "m", "") & " " & aItem(1), "yyyy-mm-dd hh:mm:ss")
    
    
    '标本号
    strLineData = DelSapce(CStr(varline(3)))
    aItem = Split(strLineData, Space(1))
    strData = strData & "|" & aItem(1) & "| |血清|0"
    
    '指标
    For intloop = 8 To 10
        strLineData = CStr(varline(intloop))
        If Trim(Mid(strLineData, 15, 5)) <> "" Then
            strData = strData & "|" & Trim(Mid(strLineData, 3, 4)) & "|" & Trim(Mid(strLineData, 15, 5))
        End If
    Next
    
    
    For intloop = 19 To 27
        strLineData = CStr(varline(intloop))
        If Trim(Mid(strLineData, 15, 5)) <> "" Then
            strData = strData & "|" & Trim(Mid(strLineData, 3, 8)) & "|" & Trim(Mid(strLineData, 15, 5))
        End If
    Next

    strResult = strData
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

Private Function DelSapce(strLine As String) As String
    '功能       删除多余的空格
    Dim intloop  As Integer
    Dim strNow As String
    strNow = strLine
    For intloop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intloop), Space(1))
    Next
    DelSapce = strNow
End Function
Private Function TrunMonth(ByVal strMonth As String) As String
    '转换中文的月份为数字
    Select Case UCase(strMonth)
        Case "一月": TrunMonth = "01"
        Case "二月": TrunMonth = "02"
        Case "三月": TrunMonth = "03"
        Case "四月": TrunMonth = "04"
        Case "五月": TrunMonth = "05"
        Case "六月": TrunMonth = "06"
        Case "七月": TrunMonth = "07"
        Case "八月": TrunMonth = "08"
        Case "九月": TrunMonth = "09"
        Case "十月": TrunMonth = "10"
        Case "十一": TrunMonth = "11"
        Case "十二月": TrunMonth = "12"
        
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
