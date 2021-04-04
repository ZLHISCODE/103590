VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_PSD_16A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:PSD-16A电解质分析仪
'厂商:南京攀事达电子仪器有限公司
'作者:陈东
'用户:成都公司
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
    Dim str_日期 As String, str_样本号 As String, str_项目 As String, str_结果 As String
    Dim lng_str As Long, lng_End As Long, str_Line As String, var_Tmp As Variant, i As Integer, str_Item As String
    
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    strResult = ""
    lng_End = InStr(mstrCode, Chr(&HD) & Chr(&HA))
    Do While lng_End > 0
        str_Line = Mid(mstrCode, 1, lng_End + 1)
        mstrCode = Mid(mstrCode, lng_End + 2)
        
        str_Line = DelSapce(str_Line)
        str_Line = Replace(str_Line, Chr(&H0), "")
        
        If str_Line Like "####-##-## ##:## #### *=*" Then
            str_Line = Replace(str_Line, "= ", "=")
            var_Tmp = Split(str_Line, " ")
            str_日期 = var_Tmp(0) & " " & var_Tmp(1)
            str_样本号 = Trim(var_Tmp(2))
            str_项目 = ""
            For i = 3 To UBound(var_Tmp)
                str_Item = var_Tmp(i)
                If InStr(str_Item, "=") > 0 Then
                    str_结果 = Trim(Split(str_Item, "=")(1))
                    If str_结果 <> "" And Trim(Split(str_Item, "=")(0)) <> "" Then
                        str_项目 = str_项目 & "|" & Trim(Split(str_Item, "=")(0)) & "|" & str_结果
                        
                    End If
                End If
            Next
            If str_项目 <> "" And str_日期 <> "" And str_样本号 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & str_样本号 & "| |血清|0" & str_项目
            End If
        End If
        lng_End = InStr(mstrCode, Chr(&HD) & Chr(&HA))
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






