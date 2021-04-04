VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_PA900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private mstrCode As String

'仪器:PA900
'厂商:
'作者:王冬建
'用户:霍山县妇幼保健站

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
        
    Dim strCode As String
    Dim lngA As Long, lngB As Long, N As Integer, M As Integer
    Dim str标本号 As String, str项目 As String, str结果 As String, str日期 As String, str质控 As String
    Dim str_Hc As String, strItem As String  '将要解析的部分数据放入一个变量，便于解析
    
    On Error GoTo errhand
    mstrCode = mstrCode & strSource
    
    '判断是否是一个完整的数据
    If mstrCode Like "lis_star*lis_end" & Chr(&HD) & Chr(&HA) Then
        '由于仪器说明上没有说明该仪器是否支持批量传输，所有按照批量传输的方式进行解析
        lngA = UBound(Split(mstrCode, "lis_end" & Chr(&HD) & Chr(&HA)))
        For N = 0 To lngA - 1
            '将单个标本的信息提取出来,进行解析
            str_Hc = Split(mstrCode, "lis_end" & Chr(&HD) & Chr(&HA))(N)
            lngB = UBound(Split(str_Hc, Chr(&HD) & Chr(&HA)))
            For M = 0 To lngB
                strItem = Split(str_Hc, Chr(&HD) & Chr(&HA))(M)
                
                If InStr(strItem, "lis_type") > 0 Then '是否质控
                    str质控 = Split(strItem, "=")(1)
                    If str质控 = "qc" Then
                        str质控 = 1
                    Else
                        str质控 = 0
                    End If
                ElseIf InStr(strItem, "lis_s_item") > 0 Then '检验项目
                    str项目 = Split(strItem, "=")(1)
                ElseIf InStr(strItem, "lis_s_sample") > 0 Then '标本号
                    str标本号 = Val(Split(strItem, "=")(1))
                ElseIf InStr(strItem, "lis_s_date") > 0 Then  '检验日期
                    str日期 = Replace(Split(strItem, "=")(1), "/", "-")
                ElseIf InStr(strItem, "lis_s_result") > 0 Then '检验结果
                    str结果 = Trim(Mid(Split(strItem, "=")(1), 1, InStr(Split(strItem, "=")(1), " ")))
                    '当该标本为质控时，仪器不会传回标本号，所以手工指定一个
                    If str质控 = 1 Then str标本号 = 800
                    strResult = strResult & "||" & str日期 & "|" & str标本号 & "| |静脉血|" & str质控 & "|" & str项目 & "|" & str结果
                End If
            Next
        Next
        strResult = Mid(strResult, 3)
        '解析完成，清空缓存
        mstrCode = ""
    End If
    Exit Sub
errhand:
    WriteLog "clsLISDev_ACCUTE", strSource, Err.Description
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






