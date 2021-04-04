VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstr_Tmp As String '延时应答

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
    Dim lngLoop As Long
    Dim Intlen As Integer
    
    Dim StrStar As String
    Dim StrRes As String
    Dim strItem As String
    Dim StrItemRes As String
    Dim StrTime As String
    Dim StrNum As String
    Dim StrInspe As String
    Dim StrResultOne As String '一个样本号结果字符串
    Dim StrWbc As String
    Dim StrRbc As String
    Dim StrPLt As String
    Dim aData() As String
    Dim aitem() As String
    Dim aResult() As String
    
   strReserved = ""
   strResult = ""
    
   Intlen = Len(strSource)
  
'    If Intlen = 1 Then '根据发送字符判断是否要发送询问字符
'       strCmd = Chr(&H6)
'       Exit Sub
'    End If
    If Replace(strSource, mstr_Tmp, "") Like "***" & Chr(&HD) & Chr(&HA) & "*" Then
        strCmd = Chr(&H6)
        mstr_Tmp = mstr_Tmp & strSource
    ElseIf Replace(strSource, mstr_Tmp, "") Like "*" & Chr(&H5) & "*" Then
        strCmd = Chr(&H6)
        mstr_Tmp = mstr_Tmp & strSource
    End If
    
    If InStr(strSource, "O|") > 0 And InStr(strSource, "L|1") Then
        aData = Split(strSource, "O|")    '样本数组
        For i = 1 To UBound(aData)
            aitem = Split(aData(i), "R|") '项目数组
            If UBound(aitem) < 1 Then
                strReserved = strSource
                Exit Sub
            End If
            StrNum = Trim(Mid(aData(i), InStr(aData(i), "1|") + 2, InStr(aData(i), "|^") - InStr(aData(i), "1|") - 2))
            StrRes = ""
            For j = 1 To UBound(aitem)
                aResult = Split(aitem(j), "|") '结果数组
                If InStr(aResult(1), "^^^") + 3 > 0 And InStr(aResult(1), "^1") > 0 Then
                    strItem = Mid(aResult(1), InStr(aResult(1), "^^^") + 3, InStr(aResult(1), "^1") - InStr(aResult(1), "^^^") - 3) '一个项目结果
                    StrItemRes = aResult(2)
                    StrRes = StrRes + "|" & strItem + "|" + StrItemRes
                End If
            Next
            If StrRes <> "" Then
                strResult = strResult + "||" + Format(Now, "yyyy-MM-dd") + "|" + StrNum + "|" + " | |0|" + Mid(StrRes, 2)
            End If
            
        Next
        
    Else
        strReserved = strSource
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)

    
    
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







