VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AC9900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AC9900 全自动电解质分析仪
'厂商:镇江奥迪康医疗仪器有限责任公司
'作者:李建飞
'用户:贵阳公司-王伟(思蓝县人民医院)

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
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String
    Dim varLine As Variant, blnEnd As Boolean
    Dim i As Integer, str_Line As String
    Dim lngLength As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        
        Do While Not blnEnd
            lngLength = InStr(strReserved, Chr(3)) - InStr(strReserved, Chr(2)) + 1
            str_Data = Mid(strReserved, InStr(strReserved, Chr(2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
            
            str_项目 = ""
            str_时间 = Trim(Mid(str_Data, 2, 20))
            str_样本号 = Trim(Mid(str_Data, 22, 6))
            
            str_Line = Mid(str_Data, 28, 15)
            If Trim(Mid(str_Line, 1, 6)) <> "" And Trim(Mid(str_Line, 7)) <> "" Then
                str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 6)) & "|" & Trim(Mid(str_Line, 7))
            End If
            varLine = Split(Mid(str_Data, 29), Chr(&H2A))
            
            For i = LBound(varLine) To UBound(varLine) - 1
                str_Line = varLine(i)
                If Trim(Mid(str_Line, 1, 6)) <> "" And Trim(Mid(str_Line, 7)) <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 6)) & "|" & Trim(Mid(str_Line, 7))
                End If
            Next
            If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
                strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "|" & " |" & "尿液|0" & str_项目
            End If
            If Not strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
                blnEnd = True
            End If
           
        Loop
    Else
       strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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




