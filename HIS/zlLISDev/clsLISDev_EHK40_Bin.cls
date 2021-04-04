VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EHK40_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:EHK-40
'厂商:天津市亚坤电子科技发展有限公司
'作者:陈东
'用户:商丘第一人民医院
Dim mStrInput  As String
Dim mlngStep As Long '接收步骤
Dim mlng_通道数 As Long '总的通道数
Dim mlng_采样数量 As Long
Dim mlng_间隔时间 As Long
Dim mlng_Total As Long '总的数据量

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
    Dim lngS As Long, lngE As Long, i As Integer, str标本号 As String, strEsr As String
    Dim strRecord As String
    
    If strSource <> "" Then
        strCmd = strSource
        mStrInput = mStrInput & strSource
    Else
        Exit Sub
    End If
    
    lngS = InStr(mStrInput, ",69,00")
    If lngS > 0 And mlngStep < 2 Then
        mlngStep = 2 '开始
        
        mStrInput = Mid(mStrInput, lngS)
    Else
        
        If mlngStep = 2 Then
            mlng_通道数 = Val("&H" & Mid(Mid(mStrInput, 7, 3), 2))
            mlngStep = 3
        ElseIf mlngStep = 3 Then
            mlng_采样数量 = Val("&H" & Mid(Mid(mStrInput, 10, 3), 2))
            mlngStep = 4
            
            mlng_Total = (6 + mlng_采样数量) * mlng_通道数
        ElseIf mlngStep = 4 Then
            mlng_间隔时间 = Val("&H" & Mid(Mid(mStrInput, 13, 3), 2))
            mlngStep = 5
        ElseIf mlngStep = 5 Then
            '收记录数据
            If Len(Mid(mStrInput, 16)) / 3 >= mlng_Total Then
                '收完了
                mStrInput = Mid(mStrInput, 16)
                For i = 1 To mlng_通道数
                    strRecord = Mid(mStrInput, 1, (6 + mlng_采样数量) * 3)
                    str标本号 = Val("&H" & Mid(strRecord, 2, 2))
                    strEsr = "ESR|" & Val("&H" & Mid(strRecord, 14, 2))
                    
                    If Val(str标本号) <> 0 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(str标本号) & "| |血液|0|" & strEsr
                    End If
                    mStrInput = Mid(mStrInput, (6 + mlng_采样数量) * 3 + 1)
                Next
                mlngStep = 0
                mlng_Total = 0
                mlng_通道数 = 0
                mlng_采样数量 = 0
                mlng_间隔时间 = 0
                
                str标本号 = ""
                strEsr = ""
            End If
            
        End If
        
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
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
