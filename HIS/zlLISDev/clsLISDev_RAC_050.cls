VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RAC_050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:全自动凝血分析仪
'厂商:RAC-050
'作者:王振涛
'用户:太原武乡县医院
'时间:2013-1-30

Private mstrCode As String    '缓存数据
Private mstrLastData As String
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
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    '数据格式："标题;图像类型;Y高度,X长度;上下左右边框留白(用于画刻度);X轴刻度[|Y刻度];曲线1数据[|曲线2数据...][;界标数据]"
    '其中 曲线数据：是y座标数据，以,分隔，多条曲线数据以|分隔
    '界标数据：是x座标数据，以,号分隔
    Dim strDate As String
    Dim strNo As String, strNow As String
    Dim strItems As String, strItem As String
    Dim varItem As Variant
    Dim strLine As String
    Dim str_Sample As String
    Dim lng_S As Long
    Dim lng_E As Long

    On Err GoTo errHandle
    If mstrLastData <> "" Then
        mstrCode = mstrLastData & mstrCode & strSource    '保存缓存的数据
        mstrLastData = ""
    Else
        mstrCode = mstrCode & strSource
    End If
    lng_E = 0: lng_S = InStr(mstrCode, "R")
    If lng_S > 0 Then lng_E = InStr(mstrCode, "FINISH")
    If lng_E = 0 Then lng_E = Len(mstrCode)
    Do While lng_S > 0 And lng_E > 0
        lng_E = InStr(mstrCode, "0R")
        If lng_E = 0 Then lng_E = Len(mstrCode)
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, ",")
        If IsNumeric(varItem(1)) = True And IsDate(varItem(2)) = True Then
            strNow = varItem(1)
            strDate = varItem(2)
            If InStr(varItem(11), "PT") > 0 And InStr(varItem(11), "APTT") = 0 Then
                strItem = varItem(11) & "|" & varItem(12) & "|INR|" & varItem(14)
            ElseIf InStr(varItem(11), "D-Dimer") > 0 Or InStr(varItem(11), "FIB") > 0 Then
                strItem = varItem(11) & "|" & varItem(15)
            Else
                strItem = varItem(11) & "|" & varItem(12)
            End If
        End If
        
        
        If strNo = "" Or strNo = strNow Then
            strItems = strItems & "|" & strItem
            strNo = strNow
        Else
            strResult = strResult & "||" & strDate & "|" & strNo & "| |血清|0" & strItems
            strItems = ""
            strItems = strItems & "|" & strItem
            strNo = strNow
        End If
        

        lng_E = 0: lng_S = InStr(mstrCode, "R")
        If lng_S > 0 Then lng_E = InStr(mstrCode, "0R")
        If InStr(mstrCode, "0R") = 0 And mstrCode <> "" Then
            mstrLastData = mstrCode
            mstrCode = ""
            If UBound(Split(mstrLastData, ",")) = 17 Then
                If Split(mstrLastData, ",")(17) <> "" Then
                    varItem = Split(mstrLastData, ",")
                    If IsNumeric(varItem(1)) = True And IsDate(varItem(2)) = True Then
                        strNow = varItem(1)
                        strDate = varItem(2)
                        If InStr(varItem(11), "PT") > 0 And InStr(varItem(11), "APTT") = 0 Then
                            strItem = varItem(11) & "|" & varItem(12) & "|INR|" & varItem(14)
                        ElseIf InStr(varItem(11), "D-Dimer") > 0 Or InStr(varItem(11), "FIB") > 0 Then
                            strItem = varItem(11) & "|" & varItem(15)
                        Else
                            strItem = varItem(11) & "|" & varItem(12)
                        End If
                    End If
                    
                    
                    If strNo = "" Or strNo = strNow Then
                        strItems = strItems & "|" & strItem
                        strNo = strNow
                    End If
                    If strItems <> "" And Val(strNo) > 0 Then
                        strResult = strResult & "||" & strDate & "|" & strNo & "| |血清|0" & strItems
                    End If
                            
                End If
            End If
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
    Exit Sub
errHandle:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
End Sub










