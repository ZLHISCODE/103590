VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MC2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MC-2000 二通道血凝仪
'厂商:美国美创
'作者:陈东
'用户:洛阳202医院
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
    Dim str_时间 As String, str_样本号 As String, str_项目 As String, str_结果 As String
    Dim str_Line As String, lng_S As Long, lng_E As Long, var_line As Variant, str_Item As String
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))
    If lng_E > 0 Then
        lng_S = lng_E - 57
        If lng_S > 0 Then
            str_Line = Mid(mstrCode, lng_S, lng_E - lng_S)
            mstrCode = Mid(mstrCode, lng_E + 2)
        
            var_line = Split(str_Line, Chr(&H9))
            If UBound(var_line) >= 10 Then
                str_时间 = Format(Now, "yyyy-MM-dd")
                str_样本号 = var_line(0)
                
                If Val(str_样本号) > 0 Then
                    str_Item = var_line(1)
                    Select Case str_Item
                    Case "01"
                        strResult = str_时间 & "|" & Val(str_样本号) & "| |血液|0|PT|" & Val(var_line(5)) / 10 & "|INR|" & Val(var_line(7)) / 100
                    Case "02"
                        strResult = str_时间 & "|" & Val(str_样本号) & "| |血液|0|APTT|" & Val(var_line(5)) / 10
                    Case "03"
                        strResult = str_时间 & "|" & Val(str_样本号) & "| |血液|0|TT|" & Val(var_line(5)) / 10
                    Case "04"
                        strResult = str_时间 & "|" & Val(str_样本号) & "| |血液|0|FIB|" & Val(var_line(10)) / 10
                    End Select
                End If
            End If
        Else
            mstrCode = Mid(mstrCode, lng_E + 2)
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


