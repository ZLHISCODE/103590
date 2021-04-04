VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IMS972"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:电解质分析仪 IMS 972
'厂商:
'作者:陈东
'用户:玉屏县人民医院

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_Line As String
    Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
    Dim intLoop As Integer
    
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "PAT_ID")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "PH")
        If lng_E > 0 Then
            lng_E = InStr(lng_E, mstrCode, Chr(&HD))
        End If
    End If
        
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------------
        var_Code = Split(str_Code, Chr(&HD))
        str_日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
        For intLoop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(intLoop)
            If Trim(Left(str_Line, 6)) = "PAT_ID" Then
                str_标本号 = Val(Mid(str_Line, 7))
                str_项目 = ""
            ElseIf InStr("K,Na,Cl,nCa,iCa,TCa,TCO2,AG,PH", Trim(Left(str_Line, 6))) > 0 Then
                str_结果 = Trim(Replace(Mid(str_Line, 7), "mmol/L", ""))
                If str_结果 <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Left(str_Line, 6)) & "|" & str_结果
                End If
            End If
        Next
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_项目
        End If
        '----------------------------------------------
        lng_S = InStr(mstrCode, "PAT_ID")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "PH")
            If lng_E > 0 Then
                lng_E = InStr(lng_E, mstrCode, Chr(&HD))
            End If
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
