VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CD3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Cell Dyn 3200
'厂商:美国雅培
'作者:陈东
'用户:贵州省肿瘤医院

'Data

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
    
    Dim varline As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String
    Dim lngLength As Long
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, Chr(3)) - InStr(strReserved, Chr(2)) + 1
            strLine = Mid(strReserved, InStr(strReserved, Chr(2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            varline = Split(strLine, ",")
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            If UBound(varline) >= 73 Then
                If varline(0) Like "*" & Chr(2) & """   ""*" Then
                    '是数据行
                    strTmp = Replace(varline(13), """", "")
                    str日期 = Format(CDate(Right(strTmp, 2) & "-" & Replace(Left(strTmp, 5), "/", "-")), "yyyy-MM-dd")
                    strTmp = Replace(varline(14), """", "")
                    str日期 = str日期 & " " & strTmp
                    
                    'str标本号 = Replace(varline(6), """", "")
                    str标本号 = Replace(varline(3), """", "")
                    If str标本号 Like "*R*T*" Then
'                        strTmp = Mid(str标本号, InStr(str标本号, "R") + 1, InStr(str标本号, "T") - InStr(str标本号, "R"))
'                        str标本号 = Val(CStr(Val(strTmp) - 1) & Mid(str标本号, InStr(str标本号, "T") + 1))
                        strTmp = Val(Mid(str标本号, InStr(str标本号, "T") + 1))
                        str标本号 = (Val(Mid(str标本号, 2)) - 1) * 10 + Val(strTmp)
                    Else
                        str标本号 = Val(str标本号)
                    End If
                    str检验人标本质控 = " |血液|" & IIf(str标本号 >= 2000, 1, 0)
                    
                    For iRow = 18 To 39
                        If Val(varline(iRow)) <> 0 Then
                            strItem = strItem & "|" & iRow & "|" & Val(varline(iRow))
                        End If
                    Next
                    
                    If strItem <> "" And Val(str标本号) <> 0 Then
                        strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                    End If
                    
                ElseIf varline(0) Like "*" & Chr(2) & """PLT""*" Then
                    'PLT图形 暂不解析
                ElseIf varline(0) Like "*" & Chr(2) & """RBC""*" Then
                    'RBC图形 暂不解析
                ElseIf varline(0) Like "*" & Chr(2) & """WBC""*" Then
                    'WBC图形 暂不解析
                End If
            End If
            If Not strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
                blnEnd = True
                'strReserved = Replace(strReserved, "^", "#") ' 将^还原为#
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

