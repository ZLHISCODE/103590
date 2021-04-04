VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VIT_DT60"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Dt60
'厂商:强生
'作者:陈东
'用户:曲阳医院

Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串
Private mStr位置 As String  '双向通信用。

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
    '!1113          554         1UREA0    20.3MMOL/L             6A
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String
    Dim str年 As String, str月 As String, str日 As String, str指标 As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*!*" Then
        '----- 开始解析结果
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的消息
            If Len(Mid(strReserved, InStr(strReserved, "!"))) >= 62 Then
                strLine = Mid(strReserved, InStr(strReserved, "!"), 62)
                strReserved = Mid(strReserved, InStr(strReserved, "!") + 62)
    
                '--解析结果
                str标本号 = Val(Mid(strLine, 19, 10))
                str日期 = Format(Now(), "YYYY-MM-DD")
                strItem = "|" & Mid(strLine, 29, 4) & "|" & Trim(Mid(strLine, 34, 8))
                str检验人标本质控 = " |血液|0"
                
                If strItem <> "" And str标本号 <> "0" Then
                    strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                End If
            Else
                blnEnd = True
            End If
            
            If Not strReserved Like "*!*" Then
                blnEnd = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    '向设备发送‘数据收到’命令
    'strCmd = Chr(&H6&)
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
    
    SendSample = ""
    
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

Private Function getNumber_Right(ByVal strSource As String) As String
    '从右开始取数字，遇到空格结束
    Dim strTmp As String, lngLength As Long
    Dim strIn As String
    strIn = strSource
    
    If Len(strIn) > 0 Then
        For lngLength = 1 To Len(strSource)
            strTmp = Right(strIn, 1)
            If IsNumeric(strTmp) Then
                getNumber_Right = strTmp & getNumber_Right
                strIn = Mid(strIn, 1, Len(strIn) - 1)
            ElseIf strTmp = "." Then
                If InStr(getNumber_Right, ".") <= 0 Then
                    getNumber_Right = strTmp & getNumber_Right
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                Else
                    Exit For
                End If
            ElseIf strTmp = " " Then
                If getNumber_Right <> "" Then
                    Exit For
                Else
                    strIn = Mid(strIn, 1, Len(strIn) - 1)
                End If
            End If
        Next
    End If
    If getNumber_Right <> "" Then
        If InStr(strSource, "- ") > 0 Then getNumber_Right = "-" & getNumber_Right
    End If
End Function


