VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_XI921CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:XI-921CT
'厂商:深圳凯特
'作者:陈德荣
'用户:四川省林业医院
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
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strLog As String
    Dim strLine As String
    Dim intStart As Integer, intEnd As Integer
    Dim aData() As String, aItem() As String
    Dim intLoop As Integer
    Dim strItem As String
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": strTmp = "": strData = ""
    
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " 接收原始数据：" & strSource
    Call writelislog(strLog)
    
    If Len(strSource) < 62 Then strReserved = strSource: Exit Sub
    
    If InStr(strSource, Chr(&HD)) = 0 Then
      strSource = strSource & Chr(&HD) & Chr(&HA)
    End If
   
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " 处理原始数据：" & strSource
    Call writelislog(strLog)
    
    aData = Split(strSource, Chr(&HD) & Chr(&HA))
    For intLoop = 0 To UBound(aData)
        strTmp = aData(intLoop)
        strTmp = Replace(strTmp, Chr(&H20) & Chr(&H20), Chr(&H20)) 'Replace(strTmp, "  ", " ")
        aItem = Split(Trim(strTmp), Chr(&H20))
        For i = 0 To UBound(aItem)
            If Val(aItem(0)) <> 0 Then
               Select Case i
                 Case 0: strData = strData & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(aItem(i)) & "| |血清|0"
                 Case 3: strData = strData & "|K|" & Val(aItem(i))
                 Case 4: strData = strData & "|Na|" & Val(aItem(i))
                 Case 5: strData = strData & "|Cl|" & Val(aItem(i))
                 Case 6: strData = strData & "|iCa|" & Val(aItem(i)) & "|TCa|" & Val(Val(aItem(i)) * 1.95)
                 Case 7: strData = strData & "|pH|" & Val(aItem(i))
               End Select
            End If
        Next
    Next
    strResult = Mid(strData, 3)
    strLog = "解析后数据：" & strResult
    Call writelislog(strLog)
    strCmd = ""
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_XI921CT_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function
