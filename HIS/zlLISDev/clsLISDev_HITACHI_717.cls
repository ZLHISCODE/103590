VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITACHI_717"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HITACHI 717
'厂商:(未知)
'作者:陈东
'用户:北京朝阳二院

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
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str条码 As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
            strTmp = Left(strLine, 4) '03E
            
            strLine = Mid(strLine, 5)
            str检验人标本质控 = " |血液|0"
            str标本号 = Val(Left(strLine, 4)) '  58
            If Right(strTmp, 1) = "E" Then
                str标本号 = str标本号 & "^1"
            ElseIf Right(strTmp, 1) = "C" Then
                str检验人标本质控 = " |血液|1"
                str标本号 = Val(Right(Left(strLine, 4), 2))
            End If
            strLine = Mid(strLine, 21)
            Do While Not (strLine Like "?*" Or strLine = "")
                '指标
                If Not Trim(Left(strLine, 3)) Like " *" Then
                    strItem = strItem & "|" & Trim(Left(strLine, 3))
                    strLine = Mid(strLine, 4)
                    strItem = strItem & "|" & Trim(Left(strLine, 6))
                    strLine = Mid(strLine, 7)
                Else
                    strLine = Mid(strLine, 4)
                End If
            Loop
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If Not strReserved Like "***" Then
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



