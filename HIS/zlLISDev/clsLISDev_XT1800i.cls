VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XT1800i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:xs-1800i
'厂商:Sysmex
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
    Dim i As Integer
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

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strBarcode As String
    Dim strLine As String
    Dim intStart As Integer, intEnd As Integer
    Dim aData() As String
    Dim intLoop As Integer
    Dim strItem As String, strItemVal As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    strCmd = Chr(&H6&)
    
    '没有找到开始位置是退出
    If InStr(strSource, "") <= 0 Then strReserved = strSource: Exit Sub
    
    '没有找到结束位置时退出
    If InStr(strSource, "") <= 0 Then strReserved = strSource: Exit Sub
    
    '取一个标本信息进行解析，其他的先返回下次进行解析
    intStart = InStr(strSource, "")
    intEnd = InStr(strSource, "") + 1
    strLine = Mid(strSource, intStart, intEnd - intStart)
    
    '返回没有解析的字串
    strReserved = Mid(strSource, intEnd)
    If InStr(strReserved, "") <= 0 Then
        '没有找到开始符时返回为空
        strReserved = ""
    End If
    
    
    '开始解析
    aData = Split(strLine, vbCrLf)
    
    '日期时间
    strTmp = Split(aData(6), "|")(12)
    strData = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        
    '标本号

    
    If aData(2) Like "*^^2*" Or aData(2) Like "*QC*" Then
        strTmp = Split(aData(2), "|")(3)
        strTmp = Trim(Split(strTmp, "^")(2))
        
        
        strData = strData & "|" & IIf(Val(strTmp) = 0, "1000", Val(strTmp)) & "| |血清|1"
        For intLoop = 4 To UBound(aData) - 2
            
            strItem = Split(Split(aData(intLoop), "|")(2), "^")(4)
            strItemVal = Split(aData(intLoop), "|")(3)
            
            If InStr(aData(2) & "^PCT^P-LCR^", strItem) > 0 Then
                strData = strData & "|" & strItem & "|" & strItemVal
            End If
        Next
    Else
        strTmp = Split(aData(3), "|")(3)
        strTmp = Trim(Split(strTmp, "^")(2))
        strData = strData & "|" & strTmp & "| |血清|0"
        
        For intLoop = 6 To UBound(aData) - 2
            
            strItem = Split(Split(aData(intLoop), "|")(2), "^")(4)
            Debug.Print strItem
            strItemVal = Split(aData(intLoop), "|")(3)
            
            If InStr(aData(3) & "^PCT^P-LCR^", strItem) > 0 Or strItem = "P-LCR" Then
                strData = strData & "|" & strItem & "|" & strItemVal
            End If
        Next
    End If
    

    
    strResult = strData
    'strCmd = Chr(&H6&)
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = "]" Then
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


