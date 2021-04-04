VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DIRUIH_100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  尿液分析仪
'医院:  贵阳市肺科医院
'作者:  陈东

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
    Dim var_Code As Variant, i As Integer, str_Line As String
    Dim str_标本号 As String, str_日期 As String, str_Tmp As String, str_项目 As String
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        str_项目 = "": str_Tmp = "": str_标本号 = "": str_日期 = ""
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If Trim(Left(str_Line, 4)) = "No." Then
                str_标本号 = Mid(str_Line, 5, 3)
                str_日期 = Mid(str_Line, 10)
            ElseIf InStr(",WBC,KET,URO,BIL,PRO,BLD,SG,Vc,GLU,NIT,pH,", "," & Trim(Left(str_Line, 4)) & ",") > 0 Then
                str_Tmp = Mid(str_Line, 5)
                str_Tmp = Replace(str_Tmp, "mmol/L", "")
                str_Tmp = Replace(str_Tmp, "umol/L", "")
                str_Tmp = Replace(str_Tmp, "Cell/uL", "")
                str_Tmp = Replace(str_Tmp, "g/L", "")
                str_Tmp = Trim(str_Tmp)
                If str_Tmp <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Left(str_Line, 4)) & "|" & str_Tmp
                End If
            End If
        Next
        If Val(str_标本号) > 0 And str_日期 <> "" And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |尿液|0" & str_项目
        End If
        '----------------------------------------------
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
