VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Elecsys2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Elecsys 2010
'厂商:罗氏
'作者:陈东
'用户:塘厦医院 双向

Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串
Private mStr位置 As String  '双向通信用。
Private strTmp As String

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str条码 As String
    Dim blnEnd As Boolean '是否已解析完
    Dim lngS As Long, lngE As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
        strReserved = Mid(strSource, InStr(strSource, Chr(&H5)) + 1)
    Else
        strReserved = strSource
        If Replace(strReserved, strTmp, "") Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
            strCmd = Chr(&H6)
            strTmp = strReserved
        End If
    End If
    
    If strSource Like "*" & Chr(2) & "1H|\^&**" Then
        '----- 开始解析结果
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的消息
            lngS = InStr(strReserved, Chr(2) & "1H")
            lngE = InStr(strReserved, "")
            strLine = Mid(strReserved, lngS, lngE - lngS)
            strReserved = Mid(strReserved, lngE + 1)
            If strReserved = Chr(4) Then strReserved = ""
            If strLine Like "*?Q|*L|1*" Then
                If strResult <> "" Then
                    '将已解析的结果先返回,再来解析双向的结果
                    strReserved = strLine & strReserved
                    blnEnd = True
                    Exit Sub
                End If
                '双向通信: 收到仪器发来的标本号，条码号。
                varLine = Split(strLine, vbCrLf)
                For iRow = LBound(varLine) To UBound(varLine)
                    If varLine(iRow) Like "?Q|*" Then
                        varItem = Split(varLine(iRow), "|")
                        str条码 = Split(varItem(2), "^")(1)
                        str标本号 = Val(Split(varItem(2), "^")(2))
                        str日期 = Format(Now, "yyyy-MM-dd")
                        
                        mStr位置 = Replace(varItem(2), "^" & str条码 & "^", "") 'SendSample过程中要用。
                        mStr位置 = Mid(mStr位置, 1, InStr(mStr位置, "^^") - 1)
                        strResult = str日期 & "|" & str标本号 & "^0^" & str条码
                        strCmd = "1|"
                    End If
                Next
                blnEnd = True
                Exit Sub
            End If
            
            '--解析结果
            varLine = Split(strLine, vbCrLf)
            For iRow = LBound(varLine) To UBound(varLine)
                If varLine(iRow) Like "?O|*" Then
                    varItem = Split(varLine(iRow), "|")
'                    str标本号 = Val(Split(varItem(3), "^")(2))
                    str标本号 = Val(varItem(2))
                    str日期 = Mid(varItem(6), 1, 8)
                    str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2)
                    str检验人标本质控 = " |血液|0"
                    strItem = ""
                    
                ElseIf varLine(iRow) Like "?R|*^^^*" Then
                    varItem = Split(varLine(iRow), "|")
                    strTmp = Split(varItem(2), "^")(3)
                    strItem = strItem & "|" & strTmp  '通道码
                    strTmp = varItem(3)
                    If InStr(strTmp, "^") > 0 Then
                        '定性的结果
                        strItem = strItem & "|" & Mid(strTmp, InStr(strTmp, "^") + 1)
                        
                    Else
                        '定量的结果
                        strItem = strItem & "|" & IIf(Trim(strTmp) = "", " ", strTmp)
                    End If
                End If
            Next
            If strItem <> "" Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If Not strReserved Like "*1H|\^&**" Then
                blnEnd = True
            End If
        Loop
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
    
    Dim varItem As Variant, varRecord As Variant, str检测项目 As String
    Dim i As Integer
    If iSendStep = 0 Then
        SendSample = Chr(5)
        iSendStep = iSendStep + 1
        Exit Function
    End If
    
    If strResponse Like "*" & Chr(6) & "*" Then
        If iSendStep = 1 Then
            '返回消息头: [STX]1H|\^&|||ASTM-Host[CR][ETX]59[CR][LF]
            SendSample = Chr(2) & "1H|\^&|||ASTM-Host" & Chr(13) & Chr(3) & "59" & Chr(13) & Chr(10)
        ElseIf iSendStep = 2 Then
            '返回病人信息：  [STX]2P|1||000004[CR][ETX]5B[CR][LF]
            varItem = Split(strSample, "|")
            SendSample = Chr(2) & "2P|1||" & varItem(1) & Chr(13) & Chr(3) & "5B" & Chr(13) & Chr(10)
        ElseIf iSendStep = 3 Then
            '返回要检测的项目 : [STX]3O|1|000004|278^0^19|^^^10^0\^^^20^0|R||||||N||||||||||||||O[CR][ETX]58[CR][LF]
            varItem = Split(strSample, "|")
            SendSample = Chr(2) & "3O|1|" & varItem(1) & "|" & mStr位置 & "|"
            str检测项目 = ""
            For i = 10 To UBound(varItem)
                str检测项目 = str检测项目 & "\^^^" & varItem(i) & "^0"
            Next
            str检测项目 = Mid(str检测项目, 2)
            SendSample = SendSample & str检测项目 & "|R||||||N||||||||||||||O" & Chr(13) & Chr(3) & "58" & Chr(13) & Chr(10)
        ElseIf iSendStep = 4 Then
            '返回消息尾: [STX]4L|1[CR][ETX]3D[CR][LF]
            SendSample = Chr(2) & "4L|1" & Chr(13) & Chr(3) & "3D" & Chr(13) & Chr(10)
        ElseIf iSendStep = 5 Then
            '返回结束符 : 
            SendSample = Chr(4)
        End If
        If iSendStep >= 5 Then
            iSendStep = 0
        Else
            iSendStep = iSendStep + 1
        End If
    End If
    
End Function








