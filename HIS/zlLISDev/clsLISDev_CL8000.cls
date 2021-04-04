VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CL8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:导津 CL-8000  生化分析仪
'厂商:日本导津
'作者:陈东
'用户:成都XX医院（陈辉采的数据）
'修改人:李建飞
'修改日期:2010-01-29

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
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_日期 As String, str_质控 As String, str_结果 As String
    Dim str_Tmp As String
    Dim strSmpStart As String
    
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    '不应答
    '-------------------------------------------------------------------------------------------------------------
    strSmpStart = ReadIni("CL8000", "标本号起始位置", App.Path & "\CL8000.ini")
    If strSmpStart = "" Then strSmpStart = "7": Call WriteIni("CL8000", "标本号起始位置", "7", App.Path & "\CL8000.ini")
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    Do While mstrCode Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*"                         '收到一个完整的记录,开始解析
    
        lng_S = InStrRev(mstrCode, Chr(&H2))
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_质控 = Mid(str_Code, 2, 1)
        
        If str_质控 = "Q" Then
            str_质控 = "1"
        ElseIf str_质控 = "D" Then
            str_质控 = "0"
        Else
            GoTo errFormat
        End If
        
        If str_质控 = "0" Then
            str_标本号 = Val(Mid(str_Code, Val(strSmpStart), 4))
            str_日期 = Format(Now, "yyyy-MM-dd HH:MM")
        
            str_Line = Mid(str_Code, 26)
            Do While Len(str_Line) > 22
                str_Tmp = Left(str_Line, 22)
                str_Line = Mid(str_Line, 23)
                str_结果 = Trim(Mid(str_Tmp, 7, 8))
                If str_结果 <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Tmp, 2, 3)) & "|" & Val(str_结果) / 1000
                End If
            Loop
            
            If Val(str_标本号) <> 0 And str_项目 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_项目
            End If
        ElseIf str_质控 = "1" Then
            str_标本号 = Mid(str_Code, 9, 3)
            str_日期 = Format(Now, "yyyy-MM-dd HH:MM")
        
            str_Line = Mid(str_Code, 19)
            Do While Len(str_Line) > 24
                str_Tmp = Left(str_Line, 24)
                str_Line = Mid(str_Line, 25)
                str_结果 = Trim(Mid(str_Tmp, 5, 8))
                If str_结果 <> "" Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Tmp, 2, 3)) & "|" & Val(str_结果) / 1000
                End If
            Loop
            
            If Val(str_标本号) <> 0 And str_项目 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|1" & str_项目
            End If

        End If
    Loop

    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        
    End If
    Exit Sub
errFormat:
    
    
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    

    SendSample = ""
    
End Function
