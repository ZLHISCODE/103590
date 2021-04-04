VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UF50HT2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: UF-50,H2000联在一个电脑上。
'厂商: sysmex
'作者: 陈东
'用户: 南京江北人民医院
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
    

    '初始化返回值
    Static objUF50 As clsLISDev_UF50   'UF50解析程序
    Static objHT2000 As clsLISDev_HT2000   'HT2000解析程序
    
    Dim StrInput As String
    Dim strOut   As String
    
    StrInput = strSource
    
    strReserved = "": strResult = "": strCmd = ""
    
    If objUF50 Is Nothing Then
        Set objUF50 = New clsLISDev_UF50
    End If
    If objHT2000 Is Nothing Then
        Set objHT2000 = New clsLISDev_HT2000
    End If
    
    'strCmd = Chr(&H6)
    If (StrInput Like "*" & Chr(&H2) & "*GLU*" & Chr(&H3) & "*") And Not ((StrInput Like "*DS*") Or (StrInput Like "*O*" & Chr(&H3) & "*")) Then
        Call objHT2000.Analyse(StrInput, strResult, strReserved, strCmd)
    ElseIf StrInput Like "**" & Chr(&H3) & "*" And (Not StrInput Like "*" & Chr(&H2) & Chr(&HA) & Chr(&HD) & "*") Then
        Call objUF50.Analyse(StrInput, strResult, strReserved, strCmd)
    Else
        If InStr(strSource, Chr(&H6)) > 0 Then
            strCmd = Chr(&H6)
            strReserved = Replace(strSource, Chr(&H6), "")
        Else
            strReserved = strSource
        End If
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function


