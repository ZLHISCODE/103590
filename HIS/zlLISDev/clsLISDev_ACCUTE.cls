VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACCUTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private mstrCode As String

'仪器:ACCUTE
'厂商:东芝
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
        
    Dim lngA As Long, M As Long, N As Long
    Dim str标本号 As String, str项目 As String, str结果 As String, str日期 As String
    Dim str_Hc As String, str_Hc1 As String
    
    On Error GoTo errhand:
    
    mstrCode = mstrCode & strSource
    
    If mstrCode Like "*" & Chr(&H3) Then
        '由于仪器可以批量传输，所以按照批量传输的方式进行解析
        lngA = UBound(Split(mstrCode, Chr(&H3)))
        For M = 0 To lngA - 1
            '提取单个标本信息进行解析
            str_Hc = Split(mstrCode, Chr(&H3))(M)
            
            '提取标本号
            str标本号 = Split(str_Hc, Chr(&HA))(0)
            str标本号 = Trim(Mid(str标本号, 4))
            str标本号 = Mid(str标本号, 1, InStr(str标本号, Chr(&H20)) - 1)
            str日期 = Format(Now, "yyyy-mm-dd hh:mm:ss")
 
            strResult = strResult & "||" & str日期 & "|" & Val(str标本号) & "| |静脉血|0"
            
            '取结果信息
            str_Hc = Split(str_Hc, Chr(&HA))(1)
            str_Hc = Mid(str_Hc, 5)
            
            For N = 1 To 100
                str项目 = Mid(str_Hc, 1, 10)
                str结果 = Val(Trim(Mid(str项目, 4)))
                str项目 = Trim(Mid(str项目, 1, 3))
                
                strResult = strResult & "|" & str项目 & "|" & str结果
                
                '将已经提取的结果删除
                str_Hc = Mid(str_Hc, 11)
                If Len(str_Hc) < 10 Then
                    Exit For
                End If
            Next
        Next
        
        strResult = Mid(strResult, 3)
        '解析完成，清空缓存
        mstrCode = ""
    End If
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






