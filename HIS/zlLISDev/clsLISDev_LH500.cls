VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LH500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String
Const conPicPath = "\draw.bmp"
'仪器:LH-500/LH750
'2007-9-28吉林第二中心医院: tqh新增
'修改图形算法
'修改人：王振涛
'2012-11-20
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

Public Sub Analyse(ByVal strSource As String, strResult As String, Reserved As String, _
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
    '补充图像的方式：
    '            1.图像数据跟随指标数据后，使用回车换行符来分隔。
    '            2.有多个图像数据时使用"^"来分隔
    '            3.单个图像数据格式: 图像画法 0=直方图  1=散点图
    '              a) 直方图: 图像名称;图像画法(0=直方图  1=散点图);X1;X2;X3;X4;X5...
    '              b) 散点图: 图像名称;图像画法(0=直方图  1=散点图):
    '                 例:00000100001000010000100010;00000100001000010000100010;
    '                 说明:1.散点图以点阵方式保存每一行使用分号来分隔.
    '                      2.有多少个分号就有多少行
    '                      3.每一行有多少个点由每一行的长度来确定
    '                      3.画图的方向是从最上边向下画，如有65*65的图就是从65行开始画(最上边开始画)
    Dim str_Line As String
    Dim lngS As Long        '开始长度
    Dim lngE As String      '结束长度
    Dim strDate As String   '日期
    Dim strNo As String     '标本号
    Dim strItems As String   '检验项目
    Dim StrWbc As String    '直方图
    Dim StrRbc As String    '带界标直方图
    Dim StrPLt As String
    Dim strDiff As String   '散点图
    Dim strImgs As String   '图形字符串
    Dim strSample As String
    Dim strDIFFLine As String
    Dim varDIFF As Variant
    Dim varItem As Variant
    Dim strLine As String
    Dim i As Integer
    Dim j As Integer
    Dim intloop As Integer
    Dim strNowData As Integer
    Dim str_条码 As String
    Dim str_getData As String
    Dim str_图形数据 As String
    str_图形数据 = ";RBCH;PLTH;PLTF;WBCT;WBCF;DIFF;"
    mstrCode = mstrCode & strSource

    mstrCode = Replace(mstrCode, "", "|")    ' 去掉字符串内的字符
    lngE = 0
    lngS = InStr(mstrCode, "")
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, "")
    Else
        lngE = InStr(mstrCode, "")
    End If
    Do While lngS > 0 And lngE > 0
        strSample = Mid(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid(mstrCode, lngE + 1)
        varItem = Split(strSample, vbCrLf)
        For i = LBound(varItem) To UBound(varItem)
            strLine = varItem(i)
            If Len(strLine) > 0 Then
                Select Case Split(strLine, " ")(0)
                Case "DATE"
                    strDate = Format(Split(strLine, " ")(1), "YYYY-MM-DD")
                Case "TIME"
                    strDate = strDate & " " & Split(strLine, " ")(1)
                    '                    strDate = Format(strDate, "yyyy-mm-dd hh;mm:ss")
                Case "ID1"
                    strNo = Split(strLine, " ")(1)
                Case "CASSPOS"
                    If Val(strNo) = 0 Then
                        str_条码 = Split(strLine, " ")(1)
                        For j = 1 To Len(str_条码)
                            If IsNumeric(Left(str_条码, 1)) Then
                                Exit For
                            Else
                                str_条码 = Mid(str_条码, 2)
                            End If
                        Next
                    End If
                Case "WBC", "RBC", "HGB", "MCV", "MCH", "RDW", "PLT", "MPV", "PDW", "PCT", "MO#", "NE#", "EO#", "BA#", "LY%", "MO%", "NE%", "EO%", "BA%"
                    strItems = strItems & "|" & Split(strLine, " ")(0) & "|" & Split(strLine, " ")(1)
                Case "HCT"
                    strItems = strItems & "|" & Split(strLine, " ")(0) & "|" & Mid(Split(strLine, " ")(1), 1, 5)
                Case "RBCH", "PLTH", "WBCT", "WBCF"  '解析RBCH图
                    strImgs = strImgs & "^" & getStrImg(Trim(Split(strLine, " ")(0)), Trim(Split(strLine, " ")(1)))
                Case "PLTF"
                    strImgs = strImgs & "|" & getStrImg(Trim(Split(strLine, " ")(0)), Trim(Split(strLine, " ")(1)))
                Case "DIFF"
                    If Len(strLine) > 4900 Then    '解析散点图
                        strImgs = strImgs & "^" & getStrImg("DIFF", Trim(Split(strLine, " ")(1)))
                    End If
                End Select
            End If
        Next
        If strDate <> "" And strNo <> "" And strImgs <> "" And strItems <> "" Then
            If str_条码 <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strNo) & "^^" & str_条码 & "| |血清|0" & strItems & IIf(strImgs <> "", vbCrLf & Mid(strImgs, 2), "")
            Else
                strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |血清|0" & strItems & IIf(strImgs <> "", vbCrLf & Mid(strImgs, 2), "")
            End If
        End If
        lngS = InStr(mstrCode, "")
        If lngS > 0 Then
            lngE = InStr(lngS, mstrCode, "")
        Else
            lngE = InStr(mstrCode, "")
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

Public Function getStrImg(ByVal strType As String, ByVal strData As String) As String
    Dim str_R As String
    Dim varDIFF As Variant
    Dim strItem As String
    Dim strLine As String
    Dim j As Integer
    Select Case strType
    Case "DIFF"    '解析DIFF图形数据
        varDIFF = Split(strData, "|")
        For j = 1 To UBound(varDIFF) - 1
            str_R = str_R & Mid(varDIFF(j), 3, Len(varDIFF(j)) - 2 - 4)
        Next
        '        Debug.Print str_R
    Case Else    '其他图形数据
        If InStr(strData, "|") > 0 Then
            varDIFF = Split(strData, "|")
            strLine = Mid(varDIFF(0), 1, Len(varDIFF(0)) - 4)
            For j = 1 To UBound(varDIFF) - 1
                strLine = strLine & Mid(varDIFF(j), 3, Len(varDIFF(j)) - 2 - 4)
            Next
            strLine = strLine & Mid(varDIFF(UBound(varDIFF)), 3)
        Else
            strLine = strData
        End If

        Do While Len(strLine) > 0
            strItem = Left(strLine, 2)
            strLine = Mid(strLine, 3)
            If CByte("&H" & strItem) >= 0 Then
                If InStr(strType, "PLT") > 0 Then
                    str_R = str_R & "," & CByte("&H" & strItem)
                Else
                    str_R = str_R & ";" & CByte("&H" & strItem)
                End If
            End If
        Loop
    End Select
    If str_R <> "" Then
        Select Case strType
        Case "RBCH"
            getStrImg = "RBC;0;32,350,52-50,104-100,156-,208-200,260-,302-300" & str_R
        Case "WBCT"
            getStrImg = "WBCT;0;32,70,15-30,30-60" & str_R
        Case "WBCF"
            getStrImg = "WBCF;0;100,250,20-50,50-100,100-,150-200" & str_R
        Case "PLTH"
            getStrImg = "PLT;4;100,200,4-2,50-10,100-20,150-30;" & Mid(str_R, 2)
        Case "PLTF"
            getStrImg = Mid(str_R, 2)
        Case "DIFF"
            getStrImg = "DIFF;100;1;" & drawDIFF(str_R)
        End Select
    End If
End Function

Public Function drawDIFF(ByVal strImgData As String) As String
    Dim x As Integer
    Dim Y As Integer
    Dim strLine As String
    Dim strTmp As String
    frmGraph.Picture1.Scale (0, 0)-(256, 256)
    frmGraph.Picture1.BackColor = vbWhite
    strLine = strImgData
    '    Debug.Print Len(strLine)
    For Y = 1 To 64
        For x = 1 To 64
            If Mid(strLine, 1, 2) <> "00" And Mid(strLine, 1, 2) <> "" Then
                Call DrawPoint(Mid(strLine, 1, 2), x, Y)
            End If
            strLine = Mid(strLine, 3)
        Next
    Next
    If Dir(App.Path & conPicPath) <> "" Then
        Kill App.Path & conPicPath
    End If
    SavePicture frmGraph.Picture1.Image, App.Path & conPicPath
    drawDIFF = App.Path & conPicPath
End Function
Public Function DrawPoint(ByVal str_in As String, ByVal x As Integer, ByVal Y As Integer)
    Dim strColor As String
    Dim i As Integer
    Dim j As Integer
    frmGraph.Picture1.DrawWidth = 1
    Select Case str_in
    Case "01", "02", "03", "04", "05", "06", "07", "08", "09"
        frmGraph.Picture1.PSet (x * 4, Y * 4), vbWhite
    Case "0B"    '橙色
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H40C0&
    Case "0C"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "0D"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H40C0&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
    Case "0E"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "0F"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H80FF&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H80FF&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H40C0&
    Case "10"
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H80FF&
    Case "11"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H80FF&
    Case "15"    '水红色
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF00FF
    Case "16", "17", "18", "19"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00FF
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF00FF
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1A"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1B"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF00FF

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1C", "1D"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H800080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC000C0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &H400040

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000C0
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF80FF
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &H400040

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &H400040
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H800080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000C0

    Case "1F"    '青色
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
    Case "20"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC000&
    Case "21"
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H8000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF00&
    Case "22"
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &H8000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H8000&
    Case "23", "24"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF00&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H8000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &H8000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00&
    Case "25"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &H8000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &H4000&
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HC000&

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &H4000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &H8000&

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HC000&
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF00&
        frmGraph.Picture1.PSet (x * 3 - 3, Y * 4 - 3), &H4000&

    Case "29"    '深蓝
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF0000
    Case "2A"
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HC00000
    Case "2C"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &H0
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HC00000
    Case "2B"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
    Case "2D"
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFFFF00
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF8080
    Case "2E"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &HFF0000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HC00000

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &HFF0000
    Case "2F"
        frmGraph.Picture1.PSet (x * 3, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 2), &HC00000
        frmGraph.Picture1.PSet (x * 3, Y * 4 - 3), &H800000

        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4), &H800000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 1), &HC00000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 2), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 1, Y * 4 - 3), &HFF8080

        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4), &HFF8080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 1), &HFF0000
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 2), &HFF8080
        frmGraph.Picture1.PSet (x * 3 - 2, Y * 4 - 3), &H800000

    Case Else
        '            Debug.Print str_in
    End Select

End Function


