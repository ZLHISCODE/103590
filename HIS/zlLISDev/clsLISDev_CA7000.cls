VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CA7000 凝血分析
'厂商:希森美康
'作者:郝立功
'用户:阳煤集团总医院
'时间:2015--08-18
Private mstrCode As String '缓存数据
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

    Dim strDate As String
    Dim strNo As String
    Dim strItems As String
    Dim varItem As Variant
    Dim strItem As String
    Dim strLine As String
    Dim str_Sample As String
    Dim lng_S As Long
    Dim lng_E As Long
    Dim intLoop As Integer
    Dim strAllItem As String
        
    On Err GoTo errHandle
    mstrCode = mstrCode & strSource    '保存缓存的数据
    '应答处理
    
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, vbCrLf) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If

    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H5)) '"1H"
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4)) '"07" & Chr(&HD) & Chr(&HA))
    
    Do While lng_S > 0 And lng_E > 0
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, Chr(&HD) & Chr(&HA))
        For intLoop = LBound(varItem) To UBound(varItem)
            strLine = Trim(varItem(intLoop))
            
            If strLine <> "" Then
                If InStr(Split(strLine, "|")(0), "R") > 0 Then
                    
                    strDate = Split(strLine, "|")(12)
                    strDate = strDate
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2)
                    strDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
                    
                    
                    strItem = Split(strLine, "|")(2)
                    strItem = Trim(Split(strItem, "^")(3))
                    strItem = strItem & "|" & Trim(Split(strLine, "|")(3))
                    
                    If strItem <> "" Then
                        strItems = strItems & "|" & strItem
                    End If
                    
                ElseIf InStr(Split(strLine, "|")(0), "O") > 0 Then
                    strResult = ""
                     strNo = Split(strLine, "|")(3)
                     strNo = Trim(Split(strNo, "^")(2))
                
                    
                End If
            End If
        Next
        
       
            
    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H5)) '"1H"
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Loop
    
    If strDate <> "" And strNo <> "" And strItems <> "" Then
           strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |血液|0" & strItems
        End If
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
errHandle:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
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












