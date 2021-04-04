VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_S8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'产品名称: S8 急诊生化分析仪
'产品型号: S8(XD-S8)
'厂商:成都贝斯达仪器有限公司
'作者:李建飞
'日期:2009-09-21

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
    Dim str_时间 As String, str_样本号 As String, str_检验项目 As String
    Dim str_标本 As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String, mstrCode As String
    Dim varline As Variant, varTemp As Variant
    Dim str_项目名称 As String
    Dim i As Integer, j As Integer, str_Line As String
    If strSource <> "" Then mstrCode = mstrCode & strReserved & strSource
    
    lng_S = InStr(mstrCode, Chr(&H1B) & Chr(&H49) & Chr(&H43))
    If InStr(mstrCode, "检验日期") Then
        If InStr(mstrCode, "检验日期") + 21 < Len(mstrCode) Then    'And Mid(mstrCode, InStr(mstrCode, "检验日期"), 25) Like "*####/##/## ##:##*"
            lng_E = InStr(mstrCode, "检验日期") + 21
        Else
            lng_E = 0
        End If
    End If
    If Not (lng_S > 0 And lng_E > 0) Then
        strReserved = mstrCode
        Exit Sub
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Data = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E + 1)
        str_检验项目 = ""
        
        varline = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = LBound(varline) To UBound(varline)
            
            str_Line = varline(i)
            If str_Line Like "*样品编号*" Then
                str_样本号 = Trim(Mid(str_Line, InStr(str_Line, "样品编号") + 5, _
                            (InStr(Mid(str_Line, InStr(str_Line, "样品编号") + 5), "科别")) - 1))
                str_标本 = "血清"         'Trim(Mid(str_Line, InStr(str_Line, "标本种类") + 5))
            ElseIf InStr(str_Line, "/L") Then
                varTemp = Split(str_Line, "/L")
                For j = LBound(varTemp) To UBound(varTemp) - 1
                    str_Tmp = Trim(varTemp(j))
                    str_项目名称 = Trim(Mid(str_Tmp, 1, 7))
                    Select Case str_项目名称
                        Case "葡萄糖"
                            str_项目名称 = "PTT"
                        Case "谷丙转氨酶"
                            str_项目名称 = "GBZAM"
                        Case "尿素"
                            str_项目名称 = "NS"
                        Case "白蛋白"
                            str_项目名称 = "BDB"
                        Case "谷氨酰转移酶"
                            str_项目名称 = "GAXZYM"
                        Case "肌酐"
                            str_项目名称 = "JG"
                        Case "淀粉酶"
                            str_项目名称 = "DFM"
                        Case "二氧化碳"
                            str_项目名称 = "EYHT"
                    End Select
                    
                    str_检验项目 = str_检验项目 & "|" & str_项目名称 & "|" & Val(Trim(Mid(str_Tmp, 8, 13)))
                Next
            ElseIf InStr(str_Line, "检验日期") > 0 Then
                str_时间 = Format(Trim(Mid(str_Line, InStr(str_Line, "检验日期") + 5)), "yyyy-MM-dd hh:mm")
            End If
        Next
        If Val(str_样本号) > 0 And IsDate(str_时间) And str_检验项目 <> "" Then
            strResult = strResult & "||" & str_时间 & "|" & str_样本号 & "|" & " |" & str_标本 & "|0" & str_检验项目
        End If
        
        lng_S = InStr(mstrCode, Chr(&H1B) & Chr(&H49) & Chr(&H43))
        If InStr(mstrCode, "检验日期") Then
            If InStr(mstrCode, "检验日期") + 21 < Len(mstrCode) Then    'And Mid(mstrCode, InStr(mstrCode, "检验日期"), 25) Like "*####/##/## ##:##*"
                lng_E = InStr(mstrCode, "检验日期") + 21
            Else
                lng_E = 0
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




