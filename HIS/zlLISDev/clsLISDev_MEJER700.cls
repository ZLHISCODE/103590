VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MEJER700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MEJER700尿液分析
'厂商:美国Mejer公司
'作者:陈德荣
'用户:广元市中医院

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
    
    '请在这个过程中写解析程序，上面的注释行请不要删除。
    Dim str_时间 As String, str_样本号 As String, str_项目 As String, str_结果 As String, str_通道号 As String
    Dim lon_str As Long, lon_end As Long
    Dim str_Data As String, strLog As String
    Dim varLine As Variant
    Dim i As Integer, str_Line As String, ITYP As Integer
    Dim str_日期 As String
    Dim str_sj As String
    Dim strIniFile As String
    Dim objStream As Scripting.TextStream
    Dim objFileSystem As New Scripting.FileSystemObject
    
    strResult = ""
    
    strIniFile = App.Path & "\MEJER700.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "ITYP=0 ;0=包含定量定性结果 1=只有定性结果"
       objStream.Close
       Set objStream = Nothing
    End If
    ITYP = Val(Replace(ReadIni("Options", "ITYP", strIniFile, 0), ";0=包含定量定性结果 1=只有定性结果", ""))
    
    lon_str = InStr(strSource, Chr(&H2))
    lon_end = InStr(strSource, Chr(&H3))
    If lon_str > 0 And lon_end > 0 Then
        str_Data = Mid(strSource, lon_str, lon_end)
        strReserved = Mid(strSource, lon_end + 1)
        str_项目 = "": str_时间 = "": str_通道号 = "": str_结果 = ""
        str_Data = Replace(str_Data, Chr(&H2), "")
        str_Data = Replace(str_Data, Chr(&H3), "")
        varLine = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = 1 To UBound(varLine)
            str_Line = varLine(i)
            If InStr(str_Data, "#") > 0 Then
              If i = 1 Then
                 str_样本号 = Val(Mid(str_Line, InStr(str_Line, "#") + 1, 4))
                 str_日期 = Trim(Mid(str_Line, InStr(str_Line, "#") + 8))
              End If
              If i = 2 Then
                 str_sj = Trim(Mid(str_Line, InStr(str_Line, "#") + 8))
                 str_时间 = str_sj & " " & str_日期
              End If
            End If
            If i >= 3 And i <= 13 Then
                  '替换单位为空值
                  str_Line = Replace(str_Line, "mmol/L", "")
                  str_Line = Replace(str_Line, "Cell/uL", "")
                  str_Line = Replace(str_Line, "g/L", "")
                  str_Line = Replace(str_Line, "leu/uL", "")
                  str_Line = Replace(str_Line, "ery/uL", "")
                  str_Line = Replace(str_Line, "umol/L", "")
                  str_Line = Trim(str_Line)
                  
                  str_通道号 = Trim(Mid(str_Line, 1, 4))
                  str_通道号 = Trim(Replace(str_通道号, "*", ""))
                  If InStr(str_Data, "#") > 0 Then
                     If ITYP = 0 Then
                       str_结果 = Trim(Mid(str_Line, 5))
                     Else
                       If InStr(Trim(Mid(str_Line, 5)), "+") > 0 Or InStr(Trim(Mid(str_Line, 5)), "-") > 0 Then
                         str_结果 = Trim(Mid(str_Line, 5, 5))
                       Else
                         str_结果 = Trim(Mid(str_Line, 5))
                       End If
                     End If
                  Else
                    If (UCase(str_通道号) = "URO") Or (UCase(str_通道号) = "VC") Then
                      str_结果 = Trim(Mid(str_Line, 5, 6))
                    Else
                      str_结果 = Trim(Mid(str_Line, 5, 9))
                    End If
                  End If
                  
                  str_结果 = Replace(str_结果, "Brown Yellow", "棕黄色")
                  str_结果 = Replace(str_结果, "Thin Yellow", "淡黄色")
                  str_结果 = Replace(str_结果, "Yellow", "黄色")
                  str_结果 = Replace(str_结果, "Green", "绿色")
                  str_结果 = Replace(str_结果, "Primrose", "淡黄色")
                  str_结果 = Replace(str_结果, "Red", "红色")
                  str_结果 = Replace(str_结果, "Orange", "橙色")
                  str_结果 = Replace(str_结果, "Blue", "蓝色")
                  str_结果 = Replace(str_结果, "Brown", "棕色")
                  str_结果 = Replace(str_结果, "Unselected", "未选")
                  str_结果 = Replace(str_结果, "Other", "其它")
                  str_结果 = Replace(str_结果, "Filemot", "黄褐色")
                  str_结果 = Replace(str_结果, "Little Dark", "稍模糊")
                  str_结果 = Replace(str_结果, "Dark", "模糊")
                  str_结果 = Replace(str_结果, "Clear", "清澈")
                  str_结果 = Replace(str_结果, "Feculent", "浑浊")
                  str_结果 = Replace(str_结果, "Negative", "阴性")
'                  str_结果 = Replace(str_结果, "Normal", "正常")
                  str_结果 = Replace(str_结果, "Positive", "阳性")
                  str_结果 = Replace(str_结果, "Small", "1+")
                  str_结果 = Replace(str_结果, "Moderate", "2+")
                  str_结果 = Replace(str_结果, "Large", "3+")
                  str_结果 = Replace(str_结果, "Trace", "微量")
                  str_结果 = Replace(str_结果, "Thin", "浅黄色")
                  str_结果 = Trim(str_结果)
                  str_项目 = str_项目 & "|" & str_通道号 & "|" & str_结果
            End If
        Next
        If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
            strResult = str_时间 & "|" & str_样本号 & "| |尿液|0" & str_项目
        End If
    Else
        strReserved = strSource
    End If
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function
