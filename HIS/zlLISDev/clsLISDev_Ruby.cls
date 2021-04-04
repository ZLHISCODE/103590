VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Ruby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'血常规
'仪器型号：Ruby
'厂家：美国雅培
'用户：晋城古矿医院
'郝立功
'日期：2014-11-28
Private mstrCode As String    '缓存数据
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


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
    Dim j As Integer
    Dim lngID As String
    Dim strLine As String
    Dim Var_date As Variant
    Dim str项目 As String
    Dim strVLineLeft  As String '左侧分界线
    Dim strVLineRight As String  '右侧分界线
    Dim strImageName  As String   '图像名称
    Dim strImage  As String       '图像数据
    '    mstrCode = ""
    mstrCode = mstrCode & strSource

    '应答
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If

    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
    Else
        lng_E = InStr(mstrCode, "L|1|N")

    End If

    Do While lng_S > 0 And lng_E > 0

        strLine = ""
        strCode = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E)

        Var_date = Split(strCode, Chr(&HD))
        For j = LBound(Var_date) To UBound(Var_date)
            strLine = Var_date(j)
            strLine = strLine

            If InStr(strLine, "No_ID") > 0 Then

                str_Day = "": str_No = "":  str项目 = ""
                str_Day = Format(Now, "yyyy-mm-dd hh:mm:ss")

                '标本号
                str_No = Split(strLine, "|")(3)

            ElseIf InStr(strLine, "R|") > 0 And InStr(strLine, "F|") > 0 Then
                '日期
                If str_Day = "" Then
                    str_Day = Split(strLine, "|")(11)
                    str_Day = Mid(str_Day, 1, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Mid(str_Day, 7, 2) & "　" & _
                              Mid(str_Day, 9, 2) & ":" & Mid(str_Day, 11, 2) & ":" & Mid(str_Day, 13, 2)
                    str_Day = Format(str_Day, "yyyy-mm-dd hh:mm:ss")
                End If

                str_Item = Split(strLine, "|")(2)
                str_Item = Split(str_Item, "^")(6)

                str_Value = Split(strLine, "|")(3)

                If str_Value = "" Then
                    str_Value = 0
                End If

                str项目 = str项目 & "|" & str_Item & "|" & str_Value
            ElseIf Split(strLine, "|")(0) = "M" Then
                '直方图
                If Split(strLine, "|")(3) = "WB1" Or Split(strLine, "|")(3) = "RBC" Or Split(strLine, "|")(3) = "PLT" Then
                    strVLineLeft = Split(Split(strLine, "|")(5), "\")(0)
                    strVLineRight = Split(Split(strLine, "|")(5), "\")(1)
                    strImageName = IIf(Split(strLine, "|")(3) = "WB1", "WBC", Split(strLine, "|")(3))
                    strLine = Split(strLine, "|")(6)
                    strImage = strImage & "^" & strImageName & ";5;255,255;10,50,50,10;0-0,50-50,100-100,150-150,200-200,250-250|0-0,50-50,100-100,150-150,200-200,250-250;" & _
                               Base64Decode(strLine) & ";" & strVLineLeft & "," & strVLineRight
                End If
            End If

        Next

        str结果 = str结果 & "||" & str_Day & "|" & str_No & "|检验人|血液|0" & str项目 & vbCrLf & Mid(strImage, 2)

        lng_S = InStr(mstrCode, Chr(&HD))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, "L|1|N")
        Else
            lng_E = InStr(mstrCode, "L|1|N")
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
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_Ruby"
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


Public Function StrValue(strItem As String, strC As Integer) As String

Dim str_j As String
str_j = strItem

StrValue = Split(strItem, "^")(strC)

End Function



Private Function Base64Decode(strEncoded As String) As String '解码

    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT, bRet() As Byte
    Dim i As Long, j As Long
    Dim strResult   As String
    Dim var()   As String
'    Dim BASE64CHR As String
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    Dim bRetTemp(3) As Byte
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63
        
        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
        bRet((i \ 4) * 3) = bT \ 65536
        If (i \ 4) * 3 Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3)
        End If
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        If ((i \ 4) * 3 + 1) Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3 + 1)
        End If
        bRet((i \ 4) * 3 + 2) = bT And 255
        If ((i \ 4) * 3 + 2) Mod 2 = 0 Then
            strResult = strResult & "," & bRet((i \ 4) * 3 + 2)
        End If
    Next
    Base64Decode = strResult
End Function
