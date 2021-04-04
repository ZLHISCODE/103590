VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HT2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:尿分仪 HT2000
'厂商:
'作者:曾超
'用户:南京江北人民医院

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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Dim strImage As String  '图像数据
    Dim varImage As Variant '图像数组
    Dim strItemVal As String   '结果值
    Dim strData As String
    Dim intloop As Integer
    Dim strTitle As String
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    '没有找到结束符时退出
    If InStr(strSource, Chr(&H3)) <= 0 Then strReserved = strSource: Exit Sub
    strLine = Mid(strSource, InStr(strSource, Chr(&H2)), InStr(strSource, Chr(&H3)))
    '只解析一个标本，后继先返回
    strReserved = Mid(strSource, InStr(strSource, Chr(&H3)) + 1)
    If InStr(strReserved, "") <= 0 Then
        strReserved = ""
    End If
    
    varLine = Split(strLine, vbCr)
    str日期 = Mid(varLine(1), 10) & Trim(Mid(varLine(2), 12))
    str标本号 = Trim(Mid(varLine(1), 6, 5))
    
    strTitle = str日期 & "|" & str标本号 & "| |血清|0"
    
    '写入结果
    For intloop = 4 To UBound(varLine) - 1
        
        If intloop = 9 Then 'SG
            strItemVal = Trim(Mid(varLine(intloop), 7, 7))
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & strItemVal
        ElseIf intloop = 5 Then 'NIT
            strItem = Trim(Mid(varLine(intloop), 8, 2))            '结果
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & strItem
        Else    '其他
            strItemVal = Trim(Mid(varLine(intloop), 10, 5))            '结果值
            strItem = Trim(Mid(varLine(intloop), 8, 2))                '结果
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & _
                        IIf(strItem = "", "", strItem) & "  " & strItemVal
        End If
    Next
    strResult = strTitle & Replace(Replace(strData, "*", ""), "-", "－")
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




