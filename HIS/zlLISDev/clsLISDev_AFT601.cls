VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AFT601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim str_Date As String, str_No As String, str_val As String
    Dim lngA As Long, lngB As Long, M As Long, N As Long
    Dim str_Hc As String, str_tmp As String, str_Item As String
    
    On Error GoTo errhand
    
    mstrCode = mstrCode & strSource
    
    '判断是否是一条完整的数据
    If mstrCode Like "*" & Chr(&HA) & Chr(&HD) Then
        '仪器支持批量传输，按批量进行解析
        lngA = UBound(Split(mstrCode, Chr(&HA) & Chr(&HD)))
        For M = 0 To lngA - 1
            str_Hc = Split(mstrCode, Chr(&HA) & Chr(&HD))(M)
            str_Date = Mid(str_Hc, 2, 16) & ":00"
            str_No = Split(str_Hc, Chr(&HA))(3)
            str_No = Mid(str_No, 5)
            '多组数据返回
            strResult = strResult & "||" & str_Date & "|" & Val(str_No) & "| |静脉血|0"
            
            '解析结果
            str_Hc = Mid(str_Hc, 56)
            lngB = UBound(Split(str_Hc, Chr(&HA)))
            For N = 0 To lngB
                If Split(str_Hc, Chr(&HA))(N) <> "" Then
                    str_tmp = Split(str_Hc, Chr(&HA))(N)
                    str_Item = Trim(Mid(str_tmp, 1, InStr(1, str_tmp, "=") - 1))
                    str_val = Trim(Mid(str_tmp, InStr(1, str_tmp, "=") + 1, 7))
                    strResult = strResult & "|" & str_Item & "|" & str_val
                End If
            Next
        Next
        strResult = Mid(strResult, 3)
        '解析完成，清空缓存
        mstrCode = ""
    End If
    Exit Sub
errhand:
    WriteLog "clsLISDev_AFT601", strSource, Err.Description
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
        If Mid(strSource, iPoint, 1) = "]" Then
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




