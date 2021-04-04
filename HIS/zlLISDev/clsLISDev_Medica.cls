VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Medica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:电解质
'厂商:??
'作者:曾超
'用户:重庆市二院
Private mstrLastrTRAY As String
Private mstrCode As String

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
    mstrLastrTRAY = ""
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
    Dim strLine As String, lngS As Long, lngE As Long, strSampleNo As String, strQC As String, strItem As String, strValue As String
    mstrCode = mstrCode & strSource
    
    Do While InStr(mstrCode, Chr(&HA)) > 0
        lngE = InStr(mstrCode, Chr(&HA))
        strLine = Mid$(mstrCode, 1, lngE)
        mstrCode = Mid$(mstrCode, lngE + 1)
        If InStr(strLine, "TRAY =") > 0 Then
            mstrLastrTRAY = Val(Split(strLine, "TRAY =")(1))
        ElseIf strLine Like "SAMPLE*Cl*" And mstrLastrTRAY <> "" Then
            strSampleNo = Val(mstrLastrTRAY & Mid(strLine, 8, 3))
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |血清|0" & strItem
            End If
        ElseIf strLine Like "NORM CNTRL*Cl*" Then
            strSampleNo = "1"
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |血清|1" & strItem
            End If
            mstrLastrTRAY = ""
        ElseIf strLine Like "ABNRM CNTRL*Cl*" Then
            strSampleNo = "2"
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            mstrLastrTRAY = ""
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |血清|1" & strItem
            End If
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    '转换中文的月份为数字
    Select Case UCase(strMonth)
        Case "一月": TrunMonth = "01"
        Case "二月": TrunMonth = "02"
        Case "三月": TrunMonth = "03"
        Case "四月": TrunMonth = "04"
        Case "五月": TrunMonth = "05"
        Case "六月": TrunMonth = "06"
        Case "七月": TrunMonth = "07"
        Case "八月": TrunMonth = "08"
        Case "九月": TrunMonth = "09"
        Case "十月": TrunMonth = "10"
        Case "十一": TrunMonth = "11"
        Case "十二月": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function

