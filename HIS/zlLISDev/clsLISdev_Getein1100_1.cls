VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Getein1100_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'血凝仪器
'仪器型号：Getein1100_1
'用户：介休市人民医院
'作战中：郝立功
'日期：2014-09-22
'
Private mstrCode As String    '缓存数据

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


'----------------------文本方式接口函数---------------------------------------------------
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '标准函数：从文件获取检验结果
    '   strFile：包含路径的文件名
    '   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
    '   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    '   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    '   返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果

End Function


'----------------------------------------------------------------------------------------

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
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
    Dim lng_S As Long, lng_E As Long
    Dim str_Sample As String, str_Day As String
    Dim str_No As String, str_Item As String, str_Value As String
    Dim INR_Value As String
    Dim str结果 As String
    Dim strCode As String
    Dim J As Integer
    Dim lngID As String
    Dim strLine As String
    Dim Var_date As Variant
    Dim var_date1 As Variant

    '    mstrCode = ""
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If

    Do While lng_S > 0 And lng_E > 0

        strCode = Mid(mstrCode, lng_S, lng_E - lng_S)
        mstrCode = Mid(mstrCode, lng_E)

        If UBound(Split(strCode, "|")) > 2 Then

            Var_date = Split(strCode, "|")
            str_Day = Var_date(13)
            str_Day = Format(str_Day, "yyyy/mm/dd hh:mm:ss")
            str_No = Trim(Var_date(3)) '提取标本号

            str_Item = Trim(Var_date(5))
            str_Value = Trim(Var_date(7))

            If InStr(str_Item, "^") > 0 Then

                var_date1 = Split(str_Item, "^")

                For J = LBound(var_date1) To UBound(var_date1)
                    '                    strLine = ""
                    str_Value = Trim(Var_date(7))
                    str_Value = Split(str_Value, "^")(J)
                    If str_Item = "CK-MB" Then
                        If Val(str_Value) < 2.5 Then
                            str_Value = "<2.5"
                        ElseIf Val(str_Value) > 80 Then
                            str_Value = ">80"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "cTnI" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 50 Then
                            str_Value = ">50"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "Myo" Then
                        If Val(str_Value) < 30 Then
                            str_Value = "<30"
                        ElseIf Val(str_Value) > 600 Then
                            str_Value = ">600"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "D-Dimer" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 30 Then
                            str_Value = ">30"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "NT-proBNP" Then
                        If Val(str_Value) < 100 Then
                            str_Value = "<100"
                        ElseIf Val(str_Value) > 35000 Then
                            str_Value = ">35000"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "hs-CRP" Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 200 Then
                            str_Value = ">200"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "PCT" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 50 Then
                            str_Value = ">50"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "mAlb" Then
                        If Val(str_Value) < 10 Then
                            str_Value = "<10"
                        ElseIf Val(str_Value) > 200 Then
                            str_Value = ">200"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "CysC" Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 10 Then
                            str_Value = ">10"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf InStr(str_Item, "MG") > 0 Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 20 Then
                            str_Value = ">20"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "NGAL" Then
                        If Val(str_Value) < 50 Then
                            str_Value = "<50"
                        ElseIf Val(str_Value) > 1500 Then
                            str_Value = ">1500"
                        Else
                            str_Value = str_Value
                        End If
                    End If
                    strLine = strLine & "|" & Split(str_Item, "^")(J) & "|" & str_Value

                Next
            Else
                If str_Item = "CK-MB" Then
                    If Val(str_Value) < 2.5 Then
                        str_Value = "<2.5"
                    ElseIf Val(str_Value) > 80 Then
                        str_Value = ">80"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "cTnI" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 50 Then
                        str_Value = ">50"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "Myo" Then
                    If Val(str_Value) < 30 Then
                        str_Value = "<30"
                    ElseIf Val(str_Value) > 600 Then
                        str_Value = ">600"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "D-Dimer" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 30 Then
                        str_Value = ">30"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "NT-proBNP" Then
                    If Val(str_Value) < 100 Then
                        str_Value = "<100"
                    ElseIf Val(str_Value) > 35000 Then
                        str_Value = ">35000"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "hs-CRP" Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 200 Then
                        str_Value = ">200"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "PCT" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 50 Then
                        str_Value = ">50"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "mAlb" Then
                    If Val(str_Value) < 10 Then
                        str_Value = "<10"
                    ElseIf Val(str_Value) > 200 Then
                        str_Value = ">200"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "CysC" Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 10 Then
                        str_Value = ">10"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf InStr(str_Item, "MG") > 0 Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 20 Then
                        str_Value = ">20"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "NGAL" Then
                    If Val(str_Value) < 50 Then
                        str_Value = "<50"
                    ElseIf Val(str_Value) > 1500 Then
                        str_Value = ">1500"
                    Else
                        str_Value = str_Value
                    End If
                End If
                strLine = strLine & "|" & str_Item & "|" & str_Value

            End If
            str结果 = str结果 & "||" & str_Day & "|" & str_No & "|检验人|血液|0"
            str结果 = str结果 & strLine
            '                  str结果 = ""
            strLine = ""
        End If

        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop

    strResult = Mid(str结果, 3)
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    
End Function


'---------------------------酶标仪接口自定义指令--------------------
Public Function CmdAnalyse(strCmdType As String) As String

End Function

Public Function StrValue(strN As String) As String

If InStr(strN, "(") > 0 Then
    StrValue = Split(strN, "(")(0)
Else
    StrValue = strN
End If


End Function
