VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AVL_OPTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:OPTI CCA 血气分析仪
'厂商:Osmetech
'作者:王振涛
'用户:太原炉厂
'时间:2012-12-17
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
    Dim intloop As Integer
    Dim strAllItem As String
    
    On Err GoTo errHandle
    mstrCode = mstrCode & strSource    '保存缓存的数据
    '应答处理
    strAllItem = ",ACID/BASE,PH,PCO2,PO2,BE,tCO2,HCO3,BB,BEact,BEecf,stHCO3,st.pH,cH+,Na+,K+,Ca++,nCa++,Cl-,AnGap,tHb,SO2,Hct(c),AaDO2,O2Ct,P50(c),"
    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, Chr(&HA) & Chr(&HD))
        For intloop = LBound(varItem) To UBound(varItem)
            strLine = Trim(varItem(intloop))
            strLine = Replace(strLine, "", "")
            strLine = Replace(strLine, "", "")
            
            If strLine <> "" Then
                If strLine Like "*##-???-##*" Then
                    strLine = Replace(strLine, " ", "")
                    strLine = Mid(strLine, 8, 2) & "-" & Mid(strLine, 4, 3) & "-" & Mid(strLine, 1, 2) & " " & Mid(strLine, 10, 2) & ":" & Mid(strLine, 13, 2)
                    strDate = Format(strLine, "yyyy-mm-dd hh:mm:ss")
                ElseIf InStr(strLine, "Sample") > 0 Then
                    strNo = Mid(strLine, InStr(strLine, ":") + 1)
                ElseIf InStr(strLine, " -") > 0 Then
                '----------------------------------不处理-----------------------------------
                ElseIf InStr(strLine, "Ref.Lim") > 0 Then
                '----------------------------------不处理-----------------------------------
                ElseIf InStr(strAllItem, Split(Trim(strLine), " ")(0)) > 0 Then
                    
                    If Split(strLine, " ")(0) Like "ACID/BASE" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 10, 4)
                    ElseIf Split(strLine, " ")(0) Like "pH" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 3, 4)
                    ElseIf Split(strLine, " ")(0) Like "st.pH" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 6, 5)
                    ElseIf Split(strLine, " ")(0) Like "tHb" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 4, 3)
                    ElseIf Split(strLine, " ")(0) Like "SO2" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 4, 2)
                    ElseIf Split(strLine, " ")(0) Like "Hct(c)" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 7, 2)
                    ElseIf Split(strLine, " ")(0) Like "O2Ct" Then
                        strItem = Split(strLine, " ")(0) & "|"
                        strLine = Replace(strLine, " ", "")
                        strItem = strItem & Mid(strLine, 5, 4)
                    ElseIf Split(strLine, " ")(0) Like "Hb" Then
                        strItem = ""
                    ElseIf Split(strLine, " ")(0) Like "P50" Then
                        strItem = ""
                    Else
                        strItem = Trim(Mid(strLine, Len(Split(strLine, " ")(0)) + 1, InStr(strLine, "mm") - Len(Split(strLine, " ")(0)) - 1))
                        strItem = Split(strLine, " ")(0) & "|" & strItem
                    End If
                    If strItem <> "" Then
                        strItems = strItems & "|" & strItem
                    End If
                End If
            End If
        Next
        
        If strDate <> "" And strNo <> "" And strItems <> "" Then
           strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |血液|0" & strItems
        End If
            
        lng_E = 0: lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Loop
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






