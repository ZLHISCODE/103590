VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BECompactX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:Compact-X 血凝仪
'厂商:德国拜尔
'作者:陈东
'用户:永川人民医院
'修改人:李建飞
'修改日期:2009-08-31

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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim varItem As Variant, str_Line As String, i As Integer
    Dim str_标本号 As String, str_日期 As String, str_项目 As String
        
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------
        str_项目 = "": str_标本号 = ""
        varItem = Split(str_Code, Chr(&HD) & Chr(&HA))
        For i = LBound(varItem) To UBound(varItem)
            str_Line = varItem(i)
            If UBound(Split(str_Line, "|")) = 7 Then
                str_标本号 = Split(str_Line, "|")(1)
                str_项目 = str_项目 & "|" & "1" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(5))
                str_项目 = str_项目 & "|" & "2" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(6))
                str_项目 = str_项目 & "|" & "3" & Trim(Split(str_Line, "|")(2)) & "|" & Trim(Split(str_Line, "|")(7))
            End If
        Next
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & Val(str_标本号) & "| |血清|0" & str_项目
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    
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




