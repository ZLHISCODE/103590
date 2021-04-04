VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Spife3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:spife 3000
'厂商:美国Helena
'作者:陈东
'用户:3201医院

Private mstrCode As String

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, strImg As String
    Dim lngS As Long, lngE As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    Do While mstrCode Like "*H|\^&*L|1*"
        '----- 开始解析结果

        '取一个完整的消息
        lngS = InStr(mstrCode, "H|\")
        lngE = InStr(lngS, mstrCode, Chr(&HD) & "L|1")
        strLine = Mid(mstrCode, lngS, lngE - lngS + 4)
        mstrCode = Mid(mstrCode, lngE + 4)
        
        str日期 = ""
        '--解析结果
        varline = Split(strLine, Chr(&HD))
        For iRow = LBound(varline) To UBound(varline)
            If varline(iRow) Like "P|*" Then
                varItem = Split(varline(iRow), "|")
'                    str标本号 = Val(Split(varItem(3), "^")(2))
                str标本号 = Val(varItem(4))
                str检验人标本质控 = " |血液|0"
                strItem = ""
                
            ElseIf varline(iRow) Like "R|*^^^*^*" Then
                varItem = Split(varline(iRow), "|")
                
                strTmp = Replace(varItem(2), "^^^", "")
                If InStr(strTmp, "TraceValues") > 0 Then
                    strImg = varItem(3)
                Else
                    If UBound(Split(strTmp, "^")) > 1 Then
                        If Split(strTmp, "^")(2) = "%" Then
                            strItem = strItem & "|" & Split(strTmp, "^")(1) '通道码
                            strTmp = varItem(3)
                            strItem = strItem & "|" & strTmp
                            If str日期 = "" Then
                                str日期 = Trim(varItem(11))
                                If IsNumeric(str日期) And Len(str日期) > 12 Then
                                    str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2) & " " & Mid(str日期, 9, 2) & ":" & Mid(str日期, 11, 2)
                                Else
                                    str日期 = Format(Now, "yyyy-MM-dd")
                                End If
                            End If
                        
                        End If
                    Else
                        strItem = strItem & "|" & Split(strTmp, "^")(1) '通道码
                        strTmp = varItem(3)
                        strItem = strItem & "|" & strTmp
                        If str日期 = "" Then
                            str日期 = Trim(varItem(11))
                            If IsNumeric(str日期) And Len(str日期) > 12 Then
                                str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2) & " " & Mid(str日期, 9, 2) & ":" & Mid(str日期, 11, 2)
                            Else
                                str日期 = Format(Now, "yyyy-MM-dd")
                            End If
                        End If
                        
                    End If

                End If
            End If
        Next
        If strItem <> "" Then
            strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
        End If
        If strImg <> "" Then
            strImg = Replace(strImg, "^", ";")
            strImg = Replace(strImg, "0.", "")
            strResult = strResult & vbCrLf & "　;0;100,200,0-0,25-,50-5,75-,100-10,125-,150-15,175-;" & strImg
        End If

    Loop
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
    
End Function


