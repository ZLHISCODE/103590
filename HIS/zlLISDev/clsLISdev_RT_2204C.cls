VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_RT_2204C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血凝分析仪 RT-2204C
'厂商:Rayto(雷杜)
'作者:陈东
'用户:綦江人民医院

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
    
    Dim strLine As String, varLine As Variant, var项目 As Variant, strData As String
    Dim str_标本号 As String, str_日期 As String, str_项目 As String
    Dim intRow As Integer, strTmp As String, strR As String
    
    strData = strSource
    strResult = ""
    
    If strData Like "*R,*" & Chr(&HA) & "*" Then
        
        varLine = Split(strData, Chr(&HA))
        str_标本号 = "": str_日期 = "": str_项目 = ""
        
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            var项目 = Split(strLine, ",")
            
            If UBound(var项目) >= 17 Then
                str_日期 = var项目(8)
                str_标本号 = var项目(1)
                str_项目 = var项目(9)
                
                If str_项目 = "PT" Then
                    str_项目 = "|PT|" & Val(var项目(12))
                ElseIf str_项目 = "DFIB" Then
                    str_项目 = "|DFIB|" & Val(var项目(14))
                ElseIf str_项目 = "APTT" Then
                    str_项目 = "|APTT|" & Val(var项目(11))
                ElseIf str_项目 = "TT" Then
                    str_项目 = "|APTT|" & Val(var项目(11))
                End If
                
                If IsDate(str_日期) And Val(str_标本号) > 0 And str_项目 <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血液|0" & str_项目
                End If
            Else
                If strLine Like "R,*" Or strLine = "R" Then
                    strReserved = strReserved & strLine
                End If
            End If
        Next
    End If
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function
