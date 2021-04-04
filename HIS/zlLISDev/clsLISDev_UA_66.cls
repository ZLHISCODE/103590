VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UA_66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:UA-66
'厂商:
'作者:陈东
'用户:成都儿童专科医院(东城根)

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    'GetStartCmd = Chr(6)
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
    '   图形数据放在检验结果后，用回车换行符(vbcrlf)分隔
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, iReturn As Integer
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, vbCrLf) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    varLine = Split(strSource, vbCrLf)
    strItem = ""
    For iRow = LBound(varLine) To UBound(varLine)
        If IsDate(varLine(iRow)) Then
            If UBound(varLine) > iRow + 14 Then
                str日期 = varLine(iRow)
                str标本号 = Val(varLine(iRow + 2))
                str检验人标本质控 = " |尿液|0"
                For iReturn = iRow + 4 To iRow + 14
                    'If InStr(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1), "-") <= 0 Then
                    str项目 = Trim(Mid(varLine(iReturn), 1, InStr(varLine(iReturn), vbTab) - 1))
                    strValues = Trim(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1))
                    
                    If UCase(strValues) = "NEG." Then
                        strValues = "-"
                    ElseIf UCase(strValues) = "NORM" Then
                        strValues = "norm"
                    ElseIf strValues = "-" Then
                        strValues = "-"
                    Else
                        Select Case UCase(str项目)
                        Case "LEU"
                            If Val(strValues) <= 15 Then
                                strValues = "± " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 70 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 70 And Val(strValues) <= 125 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 125 And Val(strValues) <= 500 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 500 Then
                                strValues = ">3＋ " & Split(strValues, " ")(0)
                            End If
                        Case "URO"
                            If Val(strValues) <= 2 Then
                                strValues = "norm " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 2 And Val(strValues) <= 4 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 4 And Val(strValues) <= 8 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 8 And Val(strValues) <= 12 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 12 Then
                                strValues = "4＋ " & Split(strValues, " ")(0)
                            End If
                        Case "BLD"
                            If Val(strValues) <= 0.06 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.06 And Val(strValues) <= 0.15 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.15 And Val(strValues) <= 0.75 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.75 Then
                                strValues = ">3＋ " & Split(strValues, " ")(0)
                            End If
                        Case "BIL"
                            If Val(strValues) <= 0.5 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 0.5 And Val(strValues) <= 1 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 1 And Val(strValues) <= 2 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 2 Then
                                strValues = ">3＋ " & Split(strValues, " ")(0)
                            End If
                        Case "KET"
                            If Val(strValues) <= 15 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 30 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 30 And Val(strValues) <= 80 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 80 Then
                                strValues = ">3＋ " & Split(strValues, " ")(0)
                            End If
                        Case "GLU"
                            If Val(strValues) <= 60 Then
                                strValues = "± " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 60 And Val(strValues) <= 100 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 100 And Val(strValues) <= 250 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 250 And Val(strValues) <= 500 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 500 Then
                                strValues = ">=3＋ " & Split(strValues, " ")(0)
                            End If
                        Case "PRO"
                            If Val(strValues) <= 15 Then
                                strValues = "± " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 15 And Val(strValues) <= 30 Then
                                strValues = "1＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 30 And Val(strValues) <= 100 Then
                                strValues = "2＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 100 And Val(strValues) <= 300 Then
                                strValues = "3＋ " & Split(strValues, " ")(0)
                            ElseIf Val(strValues) > 300 Then
                                strValues = ">=3＋ " & Split(strValues, " ")(0)
                            End If
                        Case Else
                            strValues = Split(strValues, " ")(0)
                        End Select
                    End If
                    strItem = strItem & "|" & str项目 & "|" & strValues
                    'End If)
                Next
                If UBound(varLine) > iRow + 14 Then
                    For iReturn = iRow + 15 To UBound(varLine)
                        If iReturn < UBound(varLine) Then
                            strReserved = strReserved & varLine(iReturn) & vbCrLf
                        Else
                            If varLine(iReturn) <> "" Then
                                strReserved = strReserved & varLine(iReturn)
                            End If
                        End If
                    Next
                End If
                Exit For
            Else
                strReserved = strSource
                Exit Sub
            End If
            
        End If
    Next
    
    If Val(str标本号) > 0 And strItem <> "" Then
        strResult = str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
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






