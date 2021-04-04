VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_ACL_Advance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
        
    '设备标准接口-----U200
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim IntStar As Integer
    Dim IntGetStar As Integer
    Dim IntGetEnd As Integer
    Dim Intpos1 As Integer
    Dim Intpos2 As Integer
    Dim Intpos3 As Integer
    Dim lngloop As Long
    Dim Intlen As Integer
    
    Dim StrStar As String
    Dim StrRes As String
    Dim StrItem As String
    Dim StrItemRes As String
    Dim StrTime As String
    Dim StrNum As String
    Dim StrInspe As String
    Dim StrResultOne As String '一个样本号结果字符串
    Dim StrWbc As String
    Dim StrRbc As String
    Dim StrPLt As String
    Dim aData() As String
    Dim aItem() As String
    Dim aResult() As String
    Dim aNum() As String
    
   strReserved = ""
   strResult = ""
    
   Intlen = Len(strSource)
  
    If Intlen = 1 Then '根据发送字符判断是否要发送询问字符
       If Asc(strSource) = &H3 Then strCmd = &H6
       Exit Sub
    End If
      
    If InStr(strSource, "O|") > 0 Then
        aData = Split(strSource, "O|")    '样本数组
        For i = 1 To UBound(aData)
            aItem = Split(aData(i), "R|") '项目数组
            If UBound(aItem) < 4 Then
               strReserved = strSource
               Exit Sub
            End If
            aNum = Split(aItem(0), "|")
            StrNum = Trim(Mid(aNum(1), 2, Len(aNum(1)) - 1))
            StrTime = Trim(Mid(aItem(1), 1, 12))
            StrRes = ""
            For j = 2 To 4
                StrItem = Trim(Mid(aItem(j), InStr(aItem(j), "|^^^") + 4, 3) + "-" + Trim(Str(j))) '一个项目两个结果
                aResult = Split(aItem(j), "|") '结果数组
                StrItemRes = Trim(aResult(2))
                StrRes = StrRes + StrItem + "|" + StrItemRes + "|"
            Next
            strResult = strResult + StrTime + "|" + StrNum + "|" + " | |0|" + StrRes + "|"
            
        Next
        
    Else
        strReserved = strSource
    End If
  
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








