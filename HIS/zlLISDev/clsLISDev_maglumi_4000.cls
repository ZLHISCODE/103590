VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_maglumi_4000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'仪器:MAGLUMI_4000
'厂商:深圳新产业
'作者:蒋坤
'用户:玉溪二院

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "*H|\*L|1|N*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, "L|1|N") - InStr(strReserved, "H|\")
            strLine = Mid(strReserved, InStr(strReserved, "H|\"), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, "L|1|N") + 5)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            varline = Split(strLine, Chr(&HD))
            strItem = ""
            For iRow = LBound(varline) To UBound(varline)
                If InStr(varline(iRow), "|") > 0 Then
                    varItem = Split(varline(iRow), "|")
                    
                    If InStr(varline(iRow), "H|\") Then '取检验日期
                        str日期 = Mid(varItem(13), 1, 4) & "-" & Mid(varItem(13), 5, 2) & "-" & Mid(varItem(13), 7, 2)
                    ElseIf InStr(varline(iRow), "O|") Then '取标本号
                        str标本号 = Val(varItem(2))
                        str检验人标本质控 = " |血液|0"
                        strItem = ""
                    ElseIf InStr(varline(iRow), "R|") Then '取检验结果
                        strItem = strItem & "|" & Replace(varItem(2), "^^^", "") & "|" & varItem(3)
                    End If

                End If
            Next
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If Not strReserved Like "*H|\*L|1|N*" Then
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
    strCmd = Chr(6)
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









