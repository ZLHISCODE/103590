VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS6400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:迪瑞CS6400生化仪
'厂商:迪瑞
'作者:罗云龙
'用户:库车县第二人民医院
'日期:2015.01.15
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


    Dim varLine As Variant, varData As Variant, str日期 As String, str标本号 As String
    Dim iRow As Integer, jRow As Integer, str检验人标本质控 As String, str项目 As String, strResultTmp As String
    Dim strTemp As String, strCode As String, strCmdR As String
    Dim lngS As Long, lngE As Long
    On Error GoTo ErrH
    '初始化返回值
    strResult = "": strReserved = "": strCmd = "": str日期 = "": str标本号 = "": str检验人标本质控 = "": str项目 = ""
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, Chr(&H0), "")
    strCode = Replace(strSource, " ", "")

    If strCode Like "*" & Chr(11) & "MSH|*OBX*" & Chr(13) & Chr(28) & Chr(13) & "*" Then
        Do While strCode <> ""
            iRow = iRow + 1
            lngS = InStr(strCode, "MSH")
            lngE = InStr(lngS + 3, strCode, Chr(13) & Chr(28) & Chr(13))
            If lngE > 0 Then
                strTemp = Mid(strCode, lngS - 1, lngE + 1)
                strTemp = Trim(strTemp)
                strCode = Mid(strCode, lngE + 3)
            Else
                strTemp = Mid(strCode, IIf(lngS = 0, 1, lngS), lngE + 1)
                strTemp = Trim(strTemp)
                strCode = ""
            End If
            varData = Split(strTemp, Chr(13))
            If UBound(varData) > 0 And varData(0) Like "*ORU*" Then
                strCmdR = Replace(varData(0), "ORU", "ACK") & Chr(13) & "MSA|AA|" & Mid(Mid(varData(0), InStr(varData(0), "ORU^R01") + 8), 1, InStr(Mid(varData(0), InStr(varData(0), "ORU^R01") + 8), "|") - 1) & "||||0|" & Chr(13) & Chr(28) & Chr(13)
                strCmd = "0|" & GetStrCmd(strCmdR)
            End If
            For jRow = 0 To UBound(varData)
                If Mid(varData(jRow), 2) Like "MSH*" Then
                    varLine = Split(varData(jRow), "|")
                    If varLine(15) = 0 Then
                        str检验人标本质控 = " |血液|0"
                    Else
                        str检验人标本质控 = " |血液|1"
                    End If
                End If
                If varData(jRow) Like "OBR*" Then
                    varLine = Split(varData(jRow), "|")
                    str日期 = Mid(varLine(7), 1, 4) & "-" & Mid(varLine(7), 5, 2) & "-" & Mid(varLine(7), 7, 2) & " " & Mid(varLine(7), 9, 2) & ":" & Mid(varLine(7), 11, 2) & ":" & Mid(varLine(7), 13, 2)
                    str标本号 = varLine(3)
                ElseIf varData(jRow) Like "OBX|?|*" Or varData(jRow) Like "OBX|??|*" Then
                    varLine = Split(varData(jRow), "|")
                    If str项目 = "" Then
                        str项目 = Replace(varLine(3), "^1", "") & "|" & varLine(4)
                    Else
                        str项目 = str项目 & "|" & Replace(varLine(3), "^1", "") & "|" & varLine(4)
                    End If
                End If
            Next
            If str标本号 <> "" And str项目 <> "" And str日期 <> "" Then
                strResultTmp = str日期 & "|" & str标本号 & "|" & str检验人标本质控 & "|" & str项目
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strResultTmp
                str项目 = "": str标本号 = "": str日期 = ""
            End If
        Loop
    Else
        strReserved = strCode
    End If
'    strCmd = strCmdR
    Exit Sub
ErrH:
    If Error = 1 Then
        Resume
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

Private Function GetStrCmd(strCmd As String) As String
    Dim lngCur As Long
    Do While strCmd <> ""
        GetStrCmd = GetStrCmd & Mid(strCmd, 1, 1) & Chr(&H0)
        strCmd = Mid(strCmd, 2)
    Loop
    GetStrCmd = GetStrCmd
End Function


