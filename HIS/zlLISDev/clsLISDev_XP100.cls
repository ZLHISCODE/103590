VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XP100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:XP100血气分析
'厂商:
'作者:蒋淑君
'用户:广元第三人民医院
'修改：敬秦 增加ASTM格式处理

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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    '请在这个过程中写解析程序，上面的注释行请不要删除。
    Dim str_时间 As String, str_样本号 As String, str_项目 As String, str_结果 As String, str_质控 As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String, iRows As Integer
    Dim varLine As Variant, blnEnd As Boolean
    Dim i As Integer, str_Line As String, iRow As Integer
    Dim lngLength As Long, strData As Variant
    Dim strRequest As String
    Dim strLine As String, strVarry() As String
    
    '初始化返回值
'    If strSource <> "" Then strReserved = strReserved & strSource
On Error GoTo errh
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    Call writelislog(" 收到：" & strSource)
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '收到开始符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&HA) & Chr(&HD)) > 0 Then                              '收到行结束符 03 ,应答 06
         strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H3)) > 0 Then                                         '收到开始符 03 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H4)) > 0 Then                                         '收到结束符 04 ,应答 06
        strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
'    lng_S = InStr(strReserved, Chr(&H2))
'    lng_E = InStr(strReserved, "")
    Do While lng_S > 0 And lng_E > 0
        If lng_S <= lng_E Then
            str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            str_项目 = ""
'            If str_Data Like "*/*/*" Then
'                str_质控 = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
'                str_时间 = Format(Trim(Mid(str_Data, 2, 8)), "yyyy-MM-dd") & " " & Format(Trim(Mid(str_Data, 10, 5)), "HH:mm")
'                str_样本号 = Val(Mid(str_Data, 25, 3))
'                str_项目 = str_项目 & "|" & "WBC" & "|" & Val(Trim(Mid(str_Data, 83, 5)))
'                str_项目 = str_项目 & "|" & "RBC" & "|" & Val(Trim(Mid(str_Data, 88, 5)))
'                str_项目 = str_项目 & "|" & "HGB" & "|" & Val(Trim(Mid(str_Data, 93, 5)))
'                str_项目 = str_项目 & "|" & "HCT" & "|" & Val(Trim(Mid(str_Data, 98, 5))) * (10 ^ 2)
'                str_项目 = str_项目 & "|" & "MCV" & "|" & Val(Trim(Mid(str_Data, 103, 5)))
'                str_项目 = str_项目 & "|" & "MCH" & "|" & Val(Trim(Mid(str_Data, 108, 5)))
'                str_项目 = str_项目 & "|" & "MCHC" & "|" & Val(Trim(Mid(str_Data, 113, 5)))
'                str_项目 = str_项目 & "|" & "PLT" & "|" & Val(Trim(Mid(str_Data, 118, 5)))
'                str_项目 = str_项目 & "|" & "LYM%" & "|" & Val(Trim(Mid(str_Data, 123, 5))) * (10 ^ 2)
'                str_项目 = str_项目 & "|" & "MDX%" & "|" & Val(Trim(Mid(str_Data, 128, 5))) * (10 ^ 2)
'                str_项目 = str_项目 & "|" & "NEUT%" & "|" & Val(Trim(Mid(str_Data, 133, 5))) * (10 ^ 2)
'                str_项目 = str_项目 & "|" & "LYM#" & "|" & Val(Trim(Mid(str_Data, 138, 5)))
'                str_项目 = str_项目 & "|" & "MXD#" & "|" & Val(Trim(Mid(str_Data, 143, 5)))
'                str_项目 = str_项目 & "|" & "NEUT#" & "|" & Val(Trim(Mid(str_Data, 148, 5)))
'                str_项目 = str_项目 & "|" & "RDW-C" & "|" & Val(Trim(Mid(str_Data, 153, 5))) * (10 ^ 2)
'                str_项目 = str_项目 & "|" & "RDW-S" & "|" & Val(Trim(Mid(str_Data, 158, 5)))
'                str_项目 = str_项目 & "|" & "PDW" & "|" & Val(Trim(Mid(str_Data, 163, 5)))
'                str_项目 = str_项目 & "|" & "MPV" & "|" & Val(Trim(Mid(str_Data, 168, 5)))
'                str_项目 = str_项目 & "|" & "P-LCR" & "|" & Val(Trim(Mid(str_Data, 173, 5)))
'                str_项目 = str_项目 & "|" & "1" & "|" & Val(Trim(Mid(str_Data, 178, 5)))
'                str_项目 = str_项目 & "|" & "2" & "|" & Val(Trim(Mid(str_Data, 183, 5)))
'                str_项目 = str_项目 & "|" & "3" & "|" & Val(Trim(Mid(str_Data, 188, 5)))
'                If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
'                    strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "| |" & "血清|" & str_质控 & str_项目
'                    Debug.Print strResult
'                End If
'            End If
            If str_Data Like "*1H|*" Then 'XP100添加ASTM格式处理
                strVarry = Split(str_Data, Chr(&HD))
                For i = LBound(strVarry) To UBound(strVarry)
                    strLine = strVarry(i)
                    If InStr(strLine, "R|") > 0 Then '获取检验数据
                        str_项目 = str_项目 & "|" & Split(Split(strLine, "|")(2), "^")(4) & "|" & Val(Split(strLine, "|")(3))
                        str_时间 = Split(strLine, "|")(12)
                        str_时间 = Format(Left(str_时间, 4) & "-" & Mid(str_时间, 5, 2) & "-" & Mid(str_时间, 7, 2) & " " & Mid(str_时间, 9, 2) & ":" & Mid(str_时间, 11, 2) & ":" & Right(str_时间, 2), "yyyy-mm-dd HH:mm:ss")
                    End If
                    If InStr(strLine, "O|") > 0 Then '获取标本号
                        str_样本号 = Val(Split(Split(strLine, "|")(3), "^")(2))
                    End If
                Next
                If str_样本号 <> "" And str_项目 <> "" Then
                    strResult = str_时间 & "|" & str_样本号 & "| |血液|0" & str_项目
'                    Call writelislog(" 接口返回：" & strResult)
                End If
            End If
            If str_Data Like "/*/*" Then
                str_质控 = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
                str_时间 = Format(Trim(Mid(str_Data, 2, 8)), "yyyy-MM-dd") & " " & Format(Trim(Mid(str_Data, 10, 5)), "HH:mm")
                str_样本号 = Val(Mid(str_Data, 25, 3))
                strRequest = Mid$(str_Data, InStr(str_Data, "") + 2)
                str_项目 = str_项目 & "|" & "WBC" & "|" & Val(Trim(Mid(strRequest, 1, 5)))
                str_项目 = str_项目 & "|" & "RBC" & "|" & Val(Trim(Mid(strRequest, 6, 5)))
                str_项目 = str_项目 & "|" & "HGB" & "|" & Val(Trim(Mid(strRequest, 11, 5)))
                str_项目 = str_项目 & "|" & "HCT" & "|" & Val(Trim(Mid(strRequest, 16, 5)))
                str_项目 = str_项目 & "|" & "MCV" & "|" & Val(Trim(Mid(strRequest, 21, 5)))
                str_项目 = str_项目 & "|" & "MCH" & "|" & Val(Trim(Mid(strRequest, 26, 5)))
                str_项目 = str_项目 & "|" & "MCHC" & "|" & Val(Trim(Mid(strRequest, 31, 5)))
                str_项目 = str_项目 & "|" & "PLT" & "|" & Val(Trim(Mid(strRequest, 36, 5)))
                str_项目 = str_项目 & "|" & "LYM%" & "|" & Val(Trim(Mid(strRequest, 43, 4)))
                str_项目 = str_项目 & "|" & "MDX%" & "|" & Val(Trim(Mid(strRequest, 48, 5)))
                str_项目 = str_项目 & "|" & "NEUT%" & "|" & Val(Trim(Mid(strRequest, 53, 5)))
                str_项目 = str_项目 & "|" & "LYM#" & "|" & Val(Trim(Mid(strRequest, 58, 5)))
                str_项目 = str_项目 & "|" & "MXD#" & "|" & Val(Trim(Mid(strRequest, 63, 5)))
                str_项目 = str_项目 & "|" & "NEUT#" & "|" & Val(Trim(Mid(strRequest, 67, 5)))
                str_项目 = str_项目 & "|" & "RDW-C" & "|" & Val(Trim(Mid(strRequest, 72, 5)))
                str_项目 = str_项目 & "|" & "RDW-S" & "|" & Val(Trim(Mid(strRequest, 77, 5)))
                str_项目 = str_项目 & "|" & "PDW" & "|" & Val(Trim(Mid(strRequest, 82, 5)))
                str_项目 = str_项目 & "|" & "MPV" & "|" & Val(Trim(Mid(strRequest, 87, 5)))
                str_项目 = str_项目 & "|" & "P-LCR" & "|" & Val(Trim(Mid(strRequest, 91, 5)))
                If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
                    strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "| |" & "血清|" & str_质控 & str_项目
                    Debug.Print strResult
                End If
            End If
            If str_Data Like "*D1U*" Then
                str_质控 = IIf(Trim(Mid(str_Data, 24, 2)) = "QC", 1, 0)
                str_时间 = Mid(str_Data, 45, 8)
                str_时间 = Format(Left(str_时间, 4) & "-" & Mid(str_时间, 5, 2) & "-" & Mid(str_时间, 7, 2), "yyyy-mm-dd")
                str_样本号 = Val(Mid(str_Data, 25, 3))
                strRequest = Mid$(str_Data, 68)
                str_项目 = str_项目 & "|" & "WBC" & "|" & Val(Trim(Mid(strRequest, 10, 4))) / 10 ^ 2
                str_项目 = str_项目 & "|" & "RBC" & "|" & Val(Trim(Mid(strRequest, 15, 4))) / 10 ^ 3
                str_项目 = str_项目 & "|" & "HGB" & "|" & Val(Trim(Mid(strRequest, 19, 5))) / 10
                str_项目 = str_项目 & "|" & "HCT" & "|" & Val(Trim(Mid(strRequest, 24, 4))) / 10 ^ 3
                str_项目 = str_项目 & "|" & "MCV" & "|" & Val(Trim(Mid(strRequest, 29, 4))) / 10
                str_项目 = str_项目 & "|" & "MCH" & "|" & Val(Trim(Mid(strRequest, 33, 5))) / 10
                str_项目 = str_项目 & "|" & "MCHC" & "|" & Val(Trim(Mid(strRequest, 39, 4)))
                str_项目 = str_项目 & "|" & "PLT" & "|" & Val(Trim(Mid(strRequest, 44, 4)))
                str_项目 = str_项目 & "|" & "LYM%" & "|" & Val(Trim(Mid(strRequest, 49, 4))) / 10 ^ 4
                str_项目 = str_项目 & "|" & "MDX%" & "|" & Val(Trim(Mid(strRequest, 54, 4))) / 10 ^ 3
                str_项目 = str_项目 & "|" & "NEUT%" & "|" & Val(Trim(Mid(strRequest, 59, 4))) / 10 ^ 3
                str_项目 = str_项目 & "|" & "LYM#" & "|" & Val(Trim(Mid(strRequest, 64, 5))) / 10 ^ 2
                str_项目 = str_项目 & "|" & "MXD#" & "|" & Val(Trim(Mid(strRequest, 69, 5))) / 10 ^ 2
                str_项目 = str_项目 & "|" & "NEUT#" & "|" & Val(Trim(Mid(strRequest, 74, 4))) / 10
                str_项目 = str_项目 & "|" & "RDW-C" & "|" & Val(Trim(Mid(strRequest, 84, 5))) / 10 ^ 4
                str_项目 = str_项目 & "|" & "RDW-S" & "|" & Val(Trim(Mid(strRequest, 79, 4))) / 10
'                str_项目 = str_项目 & "|" & "PDW" & "|" & Val(Trim(Mid(strRequest, 82, 5)))
'                str_项目 = str_项目 & "|" & "MPV" & "|" & Val(Trim(Mid(strRequest, 87, 5)))
'                str_项目 = str_项目 & "|" & "P-LCR" & "|" & Val(Trim(Mid(strRequest, 91, 5)))
                If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
                    strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "| |" & "血清|" & str_质控 & str_项目
                    Debug.Print strResult
                End If
            End If
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = InStr(mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errh:
    WriteLog "XP100", "解码错误", Err.Description
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

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_XP100_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function






