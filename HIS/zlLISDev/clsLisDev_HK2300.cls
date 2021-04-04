VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_HK2300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'电解质
'仪器型号：HK2300
'用户：运城卫校医院
'作战中：郝立功
'日期：2014-07-11
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
    Dim str_Day As String
    Dim str_No As String, Str_Item As String, str_Value As String
  
    Dim strResultTmp As String
    Dim strCode As String
    Dim j As Integer
   
    Dim StrLine As String
    Dim var_Date As Variant
  
    mstrCode = mstrCode & strSource
    On Error GoTo errHand
    
    lng_S = InStr(mstrCode, vbCrLf)
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, vbCrLf)
    Else
        lng_E = InStr(mstrCode, Chr(&HD))
    End If
    
    Do While lng_S > 0 And lng_E > 0
    
        strCode = Mid(mstrCode, lng_S, lng_E - lng_S)
        mstrCode = Mid(mstrCode, lng_E)
        '提取标本号
        var_Date = Split(strCode, " ")
        For j = LBound(var_Date) To UBound(var_Date)
            
            If j = 0 Then
               
                str_Day = Format(var_Date(0) & " " & var_Date(1), "yyyy/mm/dd hh:mm:ss")
    
            ElseIf j = 3 Then
                
                str_No = Mid(var_Date(3), 4, 3)
                strResultTmp = strResultTmp & "||" & str_Day & "|" & str_No & "|检验人|血液|0"
                
            ElseIf j > 3 Then
                    Str_Item = Trim(var_Date(j))
                    
                    If Str_Item = "K" Then
                        str_Value = var_Date(j + 2)
                        str_Value = Trim(StrValue(str_Value))
                        j = j + 2
                    ElseIf Str_Item = "Ca=" Or Str_Item = "pH=" Or Str_Item = "Cl=" Then
                        Str_Item = Mid(var_Date(j), 1, InStr(var_Date(j), "=") - 1)
                        str_Value = var_Date(j + 1)
                        str_Value = Trim(StrValue(str_Value))
                        j = j + 1
                    Else
                        Str_Item = Mid(var_Date(j), 1, InStr(var_Date(j), "=") - 1)
                        str_Value = ""
                        str_Value = Mid(var_Date(j), InStr(var_Date(j), "=") + 1)
                        str_Value = StrValue(str_Value)
                    End If
                    
                    strResultTmp = strResultTmp & "|" & Str_Item & "|" & str_Value
            End If
        Next
        lng_S = InStr(mstrCode, Chr(&HD))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, Chr(&HD))
        Else
            lng_E = InStr(mstrCode, Chr(&HD))
        End If
    Loop
    strResult = Mid(strResultTmp, 3)
    Exit Sub
errHand:
    WriteLog "clsLisDev_HK2300", Err.Description, "待解析结果：" & mstrCode
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






