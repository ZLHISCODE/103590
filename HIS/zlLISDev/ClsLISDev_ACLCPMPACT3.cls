VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_ACLCPMPACT3"
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
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim Intlen As Long
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
    Dim adata() As String
    Dim aitem() As String
    Dim lngloop As Long
    Dim intloop As Integer
    Dim strData As String
    Dim strTmp As String
    
    
    
    IntGetStar = 0
    IntGetEnd = 0
    pos2 = 1
    
   Intlen = Len(strSource)
  
    If InStr(strSource, "") > 0 Then
        
        adata = Split(strSource, "")
        
        For lngloop = 1 To UBound(adata)
            
            aitem = Split(adata(lngloop), vbCr) 'vbCr 代表回车
            
            '没有传完数据时退出
            If UBound(aitem) < 45 And lngloop = UBound(adata) Then
                If strData <> "" Then
                    strResult = strData
                End If
                strReserved = adata(lngloop)
                Exit Sub
            End If
            
             StrRes = ""

            For intloop = 13 To 34
                StrItem = ""
                StrItemRes = ""
                If intloop = 13 Then StrTime = Trim(Mid(aitem(intloop), 5, 15))
                If intloop = 15 Then StrNum = Trim(Mid(aitem(intloop), 6, 11))
                If intloop = 19 Or intloop = 21 Or intloop = 23 Or intloop = 24 Then
                    StrItem = Trim(Mid(aitem(intloop), 1, 6))
                    StrItemRes = Trim(Mid(aitem(intloop), 7, 6))
                End If
                If intloop = 16 Or intloop = 18 Or intloop = 22 Or intloop = 25 Or _
                   intloop = 26 Or intloop = 27 Or intloop = 28 Or intloop = 29 Or _
                   intloop = 30 Or intloop = 31 Or intloop = 32 Or intloop = 34 Then
                   StrItem = Trim(Mid(aitem(intloop), 1, 9))
                   StrItemRes = Trim(Mid(aitem(intloop), 10, 10))
                End If
                
                If StrItem <> "" Then StrRes = StrRes + StrItem + "|" + StrItemRes + "|"
                
            Next
               strResult = strResult + CStr(Now) + "|" + StrNum + "|" + " | |0|" + StrRes + "|"
        Next
    Else
        strReserved = strSource
    End If
    strResult = Left(strResult, Len(strResult) - 2) '去除最后两个字符
    strCmd = Chr(&H6&)
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








