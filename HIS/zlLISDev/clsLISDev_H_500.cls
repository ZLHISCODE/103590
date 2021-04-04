VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H_500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:H-500
'厂商:长春瑞迪
'作者:陈东
'用户:贵阳肿瘤医院

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
    
    Dim varline As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H3&)) > 0 Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, Chr(&H3)) - InStr(strReserved, Chr(&H2))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            iCount = 0
            varline = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = 0 To UBound(varline)
                
                Select Case iCount
                    Case 1 '日期
                       
                        str日期 = Trim(Replace(varline(iRow), "Date:", ""))
                        If InStr(str日期, ",") > 0 Then
                             '北京的是:  Date:08,16,2007 15:10 格式
                            str日期 = Mid(str日期, 7, 4) & "-" & Mid(str日期, 1, 2) & "-" & Mid(str日期, 4, 2)
                        End If
                    Case 2 '标本号
                        If InStr(varline(iRow), ":") > 0 Then varline(iRow) = Replace(varline(iRow), ":", "")
                        str标本号 = Val(Replace(varline(iRow), "No.", ""))
                        str检验人标本质控 = " |尿液|0"
                    Case 3 To 13
                        'If blnReturn(MidUni(varLine(iRow), 8)) Then
                            If Trim(MidUni(varline(iRow), 5)) <> "" And Trim(MidUni(varline(iRow), 2, 3)) <> "" Then
                                strItem = strItem & "|" & Trim(MidUni(varline(iRow), 2, 3)) & "|" & Trim(MidUni(varline(iRow), 5))
                            End If
                        'End If

                End Select
                If varline(iRow) <> "" Then iCount = iCount + 1
            Next
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strItem = Replace(strItem, "umol/L", "")
                strItem = Replace(strItem, "mmol/L", "")
                strItem = Replace(strItem, "Ery/uL", "")
                strItem = Replace(strItem, "g/L", "")
                strItem = Replace(strItem, "Leu/uL", "")
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            If InStr(strReserved, Chr(&H3&)) = 0 Then
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


Private Function MidUni(ByVal strTemp As String, ByVal Start As Long, Optional ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    If Length = 0 Then Length = LenB(strTemp)
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Private Function blnReturn(ByVal strValues As String) As Boolean
    '是否返回检验结果，如果未做，则不返回
    blnReturn = True
    If Val(strValues) = 0 Then
        If InStr(strValues, "0") <= 0 Then
            blnReturn = False
        End If
    End If
End Function

