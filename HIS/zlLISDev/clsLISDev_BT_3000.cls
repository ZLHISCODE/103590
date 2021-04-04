VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BT_3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BT-3000
'厂商:意大利
'用户:湖南邵阳正骨医院
'开发:曾超
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
    
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean
    Dim intQC As Integer       '质控位数
    Dim intDataType As Integer '数据类型 0＝结果、1＝请求
    Dim intEnd As Integer      '最后位置
    Dim strStartChr As String, strEndChr As String
    Dim strNowData As String
    Dim strTmp As String
    Dim blnEmergency As String
    Dim str标本号 As String
    Dim strItem As String
    
    On Error GoTo errHandle
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    strStartChr = Chr(&O2)
    strEndChr = Chr(&O3)
    
    '没有开始和结束时退出
    If InStr(strSource, strStartChr) <= 0 Then strReserved = strSource
    If InStr(strSource, strEndChr) <= 0 Then strReserved = strSource
    
    strNowData = Mid(strSource, InStr(strSource, strStartChr), InStr(strSource, strEndChr) + Len(strEndChr))
    strReserved = Mid(strSource, InStr(strSource, strEndChr) + Len(strEndChr))
    
    strNowData = Mid(strNowData, InStr(strNowData, Chr(&H54)) + 1)
    strTmp = Mid(strNowData, 1, 10)
    blnEmergency = IIf(InStr(strTmp, "Batch") > 0, True, False)
    str标本号 = Replace(Replace(strTmp, "Batch", ""), " ", "")
    
    strData = Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & str标本号 & "| |血清|0"
    strTmp = Mid(strNowData, 12)
    Do While strTmp <> ""
        strItem = Mid(strTmp, 1, 21)
        If Len(strItem) >= 21 Then
            strTmp = Mid(strTmp, 22)
            strData = strData & "|" & Trim(Mid(strItem, 1, 3)) & "|" & Val(Mid(strItem, 16))
        Else
            strTmp = ""
        End If
    Loop
    
    If strData <> "" Then strResult = strData
    Exit Sub
errHandle:
    Call WriteLog("BT_3000 analyse", "解码错误：" & Err.Description, "")
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    
    
    
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

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




