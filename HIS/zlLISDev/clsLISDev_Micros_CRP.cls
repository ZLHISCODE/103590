VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Micros_CRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ABX Micros CRP
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
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    str项目 = ",!,@,#,$,%,&,',"",A,B,C,2,3,4,5,6,7,8,"
    If strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            varLine = Split(strLine, Chr(&HD&))
            
            For iRow = 0 To UBound(varLine)
                strTmp = varLine(iRow)
                Select Case Left(strTmp, 1)
                    Case "q"  '日期
                        str日期 = Format(CDate(Mid(strTmp, 9, 2) & "-" & Mid(strTmp, 6, 2) & "-" & Mid(strTmp, 3, 2)), "yyyy-MM-dd")
                    Case "u"  '标本号
                        str标本号 = Val(Mid(strTmp, 3))
                        str检验人标本质控 = " |血液|0"
                    Case Else
                        If InStr(str项目, "," & Left(strTmp, 1) & ",") > 0 Then
                            strValues = Mid(strTmp, 3, 5)
                            If strValues <> "--.--" Then
                                strItem = strItem & "|" & SwapWord(Left(strTmp, 1)) & "|" & Val(strValues)
                            End If
                        End If
                End Select
            
            Next
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

Private Function SwapWord(ByVal strWord) As String
    '将仪器返回的特殊字符转换为普通字符,避免返回到HIS中出错.
    Select Case strWord
    Case "!"
        SwapWord = "1"
    Case "@"
        SwapWord = "9"
    Case "#"
        SwapWord = "D"
    Case "%"
        SwapWord = "E"
    Case "'"
        SwapWord = "F"
    Case """"
        SwapWord = "G"
    Case "$"
        SwapWord = "H"
    Case "&"
        SwapWord = "I"
    Case Else
        SwapWord = strWord
    End Select
End Function

