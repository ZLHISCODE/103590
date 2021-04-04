VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  新成XC8001生化分析仪
'医院:  巩留县中医院
'作者:  郭进东
'时间:  2016/01/10

Private mstrCode As String '缓存接收到的数据

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
    Dim intloop As Integer
    Dim strDate As String               '日期
    Dim strSampleNO As String           '标本号
    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    Dim str_Sample As String
    Dim str_No As String, Str_Item As String, str_Value As String
    Dim varLine As Variant, varData As Variant, iRow As Integer, varData1 As Variant
    
    On Error GoTo errhand
    
    '第一步，缓存接收到的数据
    '缓存是因为串口通讯速度较慢，接收数据时，有可能会接收多次才能接收到一个
    '完整的数据
    If InStr(strSource, Chr(&H5)) > 0 Then
    strCmd = Chr(6)
     ElseIf InStr(strSource, vbCrLf) > 0 Then
    strCmd = Chr(6)
     ElseIf InStr(strSource, Chr(&H4)) > 0 Then
    strCmd = Chr(6)
    End If
    mstrCode = mstrCode & strSource
   
    '第二步，判断是否收到了一个完整的数据，是完整的数据才开始解码，
    '否则，什么都不做。
'    mstrCode = Replace(mstrCode, Chr(13), "")
'    mstrCode = Replace(mstrCode, Chr(10), "")
'    mstrCode = Replace(mstrCode, Chr(28), "")
    lngBegin = InStr(mstrCode, "H|\^&|")
    lngEnd = InStr(lngBegin, mstrCode, "L|1|N")
   
    Do While lngBegin > 0 And lngEnd > 0
        '第三步，收到完整的数据，先把这个数据从缓存中取出来。
        str_Sample = Mid(mstrCode, lngBegin, lngEnd - lngBegin + 10)
        mstrCode = Mid(mstrCode, lngEnd + 12)
 
        '第四步，取出日期，标本号，项目，检验结果
        varLine = Split(str_Sample, Chr(10))
        
        For iRow = 0 To UBound(varLine)
            If varLine(iRow) Like "H|\^&|*" Then
                varData = Split(varLine(iRow), "|")
                If varData(12) <> "" Then
                    strDate = Mid(varData(12), 1, 4) & "-" & Mid(varData(12), 5, 2) & "-" & Mid(varData(12), 7, 2) & " " & Mid(varData(12), 9, 2) & ":" & Mid(varData(12), 11, 2)
                End If
            End If
            
'            If strSampleNO = "" And varLine(iRow) Like "?P|*|*" Then
'                varData = Split(varLine(iRow), "|")
'                strSampleNO = Val(varData(1))
'            End If
            
            If varLine(iRow) Like "?O|*|*" Then
                varData = Split(varLine(iRow), "|")
                strSampleNO = Val(Split(varData(2), "^")(1))
            End If
            
            If varLine(iRow) Like "*R*^^^*" Then
                varData = Split(varLine(iRow), "|")
                If varData(3) <> "No Result" Then
                    varData1 = Split(varData(2), "^")
                    Str_Item = Str_Item & "|" & varData1(3) & "|" & varData(3)
                    If strDate = "" Then
                        strDate = Mid(varData(12), 1, 4) & "-" & Mid(varData(12), 5, 2) & "-" & Mid(varData(12), 7, 2) & " " & Mid(varData(12), 9, 2) & ":" & Mid(varData(12), 11, 2)
                    End If
                End If
            End If
        Next
           
        '第五步 将取出的数据组织成规定格式存入strResult变量。
        If Str_Item <> "" And Val(strSampleNO) <> 0 Then
            strResult = strResult & "||" & strDate & "|" & strSampleNO & "| |血清|0" & Str_Item
                        
        End If
        Str_Item = ""
        strSampleNO = ""
        strDate = ""
        lngBegin = 0
        lngEnd = 0
        '第六步 检查缓冲中的数据，如果还有完整的记录，则继续解码
        lngBegin = InStr(mstrCode, "H|\^&|")
        If lngBegin > 0 Then lngEnd = InStr(lngBegin, mstrCode, "L|1|N")
    Loop
    '解码完成，将多余的||字符去掉。
    If strResult <> "" Then strResult = Mid(strResult, 3)
    strReserved = mstrCode
    Exit Sub
errhand:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
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




