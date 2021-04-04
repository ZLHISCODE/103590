VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UT8020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:URIT-8020生化分析仪
'厂商:优利特
'作者:陈德荣
'用户:成都市慢性病医院

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
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String, mstrCode As String
    Dim str_质控 As String, str_Return As String
    Dim str_项目 As String
    Dim str_结果 As String
    
    strResult = "": strReserved = ""
    
    lng_S = InStr(strSource, Chr(&H7E))
    lng_E = InStr(strSource, Chr(&H40))
    
    If lng_S > 0 And lng_E > 0 Then
        str_Code = Mid(strSource, lng_S, lng_E)
        strReserved = Mid(strSource, lng_E + 1)
        
        str_Return = "": str_标本号 = "": str_日期 = ""
        var_Code = Split(str_Code, Chr(&HA))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If lng_count = 0 Then
                '标本号，日期
                var_Filed = Split(str_Segment, "`")
                str_标本号 = var_Filed(5)
                str_日期 = Trim(var_Filed(11))
                str_日期 = Replace(str_日期, Chr(&HD), "")
                str_日期 = Trim(str_日期) & " " & Format(Now, "HH:MM:SS")
                If Not IsDate(str_日期) Then str_日期 = Format(Now, "YYYY-MM-DD HH:MM:SS")
            Else
               If str_Segment <> "@" Then
                  var_Filed = Split(str_Segment, "`")
                  str_项目 = Trim(var_Filed(1))
                  str_结果 = Trim(var_Filed(3))
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
               End If
            End If
        Next
        If Val(str_标本号) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
        End If
        If strResult <> "" Then strResult = Mid(strResult, 3)
    Else
      strReserved = strSource
    End If
End Sub

