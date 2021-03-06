VERSION 5.00
Begin VB.Form frmGraph 
   BorderStyle     =   0  'None
   Caption         =   "Graph"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picImg 
      AutoRedraw      =   -1  'True
      Height          =   2500
      Left            =   15
      ScaleHeight     =   2445
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   -15
      Width           =   2500
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Diff
    'Diff AL 作图 要用
    NoL = 0: NoN: NoE: LN: RN: LL: AL: LMU: LMD: LMN: MN: RM: NL: NE: RMN: FNE: FMN: FLN
End Enum

Public Function Draw_HMX_DF1(ByVal str_Line As String, ByVal str_Data As String) As String
    '画DF1图
    '入参
    '   str_Line:画线的坐标，用，分隔，一共5个
    '   str_Data:散点图数据
    '出参
    '   绘图成功，返回图形文件名。
    
    picImg.Scale (0, 0)-(256, 256)
    picImg.BackColor = vbWhite
    Dim X As Integer, Y As Integer
    Dim i_L1 As Integer, i_L2 As Integer, i_L3 As Integer, i_L4 As Integer, i_L5 As Integer
    Dim str_Img As String
    str_Img = str_Data
    i_L1 = Split(str_Line, ",")(0)
    i_L2 = Split(str_Line, ",")(1)
    i_L3 = Split(str_Line, ",")(2)
    i_L4 = Split(str_Line, ",")(3)
    i_L5 = Split(str_Line, ",")(4)
    picImg.Line (i_L2, 0)-(i_L2, 256 - i_L4), vbBlack, BF
    picImg.Line (i_L1, 0)-(i_L1, 256 - i_L5), vbBlack, BF
    picImg.Line (0, 256 - i_L3)-(i_L1, 256 - i_L3), vbBlack, BF
    picImg.Line (i_L1, 256 - i_L4)-(256, 256 - i_L4), vbBlack, BF
    picImg.Line (0, 256 - i_L5)-(256, 256 - i_L5), vbBlack, BF
    
    
    For X = 1 To 64
        For Y = 64 To 1 Step -1
            If Mid(str_Img, 1, 1) <> "0" Then
                Call DrawPoint(Mid(str_Img, 1, 1), X, Y)
            End If
            str_Img = Mid(str_Img, 2)
        Next
    Next
    If Dir(App.Path & "\DF1_Tmp.Bmp") <> "" Then
        Kill App.Path & "\DF1_Tmp.Bmp"
    End If
    Draw_HMX_DF1 = App.Path & "\DF1_Tmp.Bmp"
    SavePicture picImg.Image, Draw_HMX_DF1
    
End Function

Public Function Draw_HMX_DF2(ByVal str_Data As String) As String
    '画DF2图
    '入参
    '   str_Data:散点图数据
    '出参
    '   绘图成功，返回图形文件名。
    
    
    picImg.Scale (0, 0)-(256, 256)
    picImg.BackColor = vbWhite
    Dim X As Integer, Y As Integer
    Dim str_Line As String
    
    str_Line = str_Data
    For X = 1 To 64
        For Y = 64 To 1 Step -1
            If Mid(str_Line, 1, 1) <> "0" Then
                Call DrawPoint(Mid(str_Line, 1, 1), X, Y)
            End If
            str_Line = Mid(str_Line, 2)
        Next
    Next
    If Dir(App.Path & "\DF2_Tmp.Bmp") <> "" Then
        Kill App.Path & "\DF2_Tmp.Bmp"
    End If
    Draw_HMX_DF2 = App.Path & "\DF2_Tmp.Bmp"
    SavePicture picImg.Image, Draw_HMX_DF2
    
End Function

Private Function DrawPoint(ByVal str_in As String, ByVal X As Integer, ByVal Y As Integer)
    '画点函数
    Dim lng_x As Long, lng_y As Long
    Select Case str_in
    Case "1"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If lng_x = 1 And lng_y = 1 Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "2"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 2) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "3"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 3) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "4"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x = 1 And lng_y <= 4) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    
    Case "5"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y = 1) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "6"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y <= 2) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "7"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y <= 3) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "8"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x <= 2 And lng_y >= 2 And lng_y <= 4) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "9"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y = 1) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    Case "A"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y >= 2 And lng_y <= 2) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbBlack
                End If
            Next
        Next
    ', "C", "D", "E", "F"
    '问题：29348
    '修改颜色
    Case "B"
        For lng_x = 1 To 4
            For lng_y = 1 To 4
                If (lng_x >= 2 And lng_x <= 3 And lng_y >= 2 And lng_y <= 3) Then
                    picImg.PSet ((X - 1) * 4 + lng_x, (Y - 1) * 4 + lng_y), vbYellow
                End If
            Next
        Next
    
    Case "C"

    Case "D"
    Case "E"
    Case "F"
    End Select
End Function

Private Sub Form_Load()
   Me.Hide
End Sub

Public Function Draw_Bc5500(ByVal str_bin As String, ByVal strFilename As String, ByVal strColor) As Boolean
    
    Dim lngX As Long, lngY As Long, lngV As Long, i As Integer
    Dim strByte As String, strV As String
    Dim strData As String, lngCount As Long, lngDawPoint As Long
    Dim strColorPoint As String, lngPointColor As Long
    Dim strInColor As String, lngMaxType As Long
    
    strData = str_bin
    strInColor = strColor
    picImg.Scale (0, 0)-(256, 256)
    picImg.BackColor = vbWhite
    
    picImg.Line (0, 0)-(0, 255)
    picImg.Line (0, 255)-(255, 255)
    
    Do While Len(strInColor) > 0
        For i = 0 To 1
            strByte = Mid(Left(strInColor, 3), 2)
            strInColor = Mid(strInColor, 4)
            If i = 0 Then
                strV = strByte
            Else
                strColorPoint = strColorPoint & "," & Val("&H" & strV & strByte)
            End If
        Next
    Loop
    
    If strColorPoint <> "" Then
        strColorPoint = Mid(strColorPoint, 2)
        lngMaxType = UBound(Split(strColorPoint, ","))
        
    End If

     
    Do While Len(strData) > 0

        
        strByte = Mid(Left(strData, 3), 2)
        lngX = Val("&H" & strByte)
        strData = Mid(strData, 4)

        strByte = Mid(Left(strData, 3), 2)
        lngY = Val("&H" & strByte)

        strData = Mid(strData, 10)
        
        lngCount = lngCount + 1

        If lngCount > lngDawPoint Then
            '换色
            If InStr(strFilename, "BASO") > 0 Then
                If UBound(Split(strColorPoint, ",")) = lngMaxType Then
                    lngPointColor = vbBlue
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 1 Then
                    lngPointColor = vbGreen
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 2 Then
                    lngPointColor = vbCyan
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 3 Then
                    lngPointColor = vbRed
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 4 Then
                    lngPointColor = vbMagenta
                End If
            Else
                If UBound(Split(strColorPoint, ",")) = lngMaxType Then
                    lngPointColor = vbBlue
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 1 Then
                    lngPointColor = vbGreen
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 2 Then
                    lngPointColor = vbMagenta
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 3 Then
                    lngPointColor = vbRed
                ElseIf UBound(Split(strColorPoint, ",")) = lngMaxType - 4 Then
                    lngPointColor = vbCyan
                End If
            
            End If

            If strColorPoint <> "" Then
                If InStr(strColorPoint, ",") > 0 Then
                    lngDawPoint = lngDawPoint + Mid(strColorPoint, 1, InStr(strColorPoint, ",") - 1)
                    strColorPoint = Mid(strColorPoint, InStr(strColorPoint, ",") + 1)
                Else
                    lngDawPoint = lngDawPoint + strColorPoint
                    strColorPoint = ""
                End If
            End If
        End If
        picImg.PSet (lngX, 256 - lngY), RGB(lngPointColor Mod 256, lngPointColor / 256 Mod 256, lngPointColor / 256 / 256)
    Loop
    
    If Dir(strFilename) <> "" Then
        Kill strFilename
    End If
    SavePicture picImg.Image, strFilename
    Draw_Bc5500 = True
End Function

Public Function DrawP60(ByVal str_in As String, ByVal strFilename As String, ByVal strFlag As String) As Boolean
    Dim str_Line As String, X As Integer, Y As Integer
    picImg.Scale (0, 0)-(128, 128)
    picImg.BackColor = vbWhite
    str_Line = str_in
    For Y = 0 To 127
        For X = 0 To 127
            If Val(Replace(Mid(str_Line, 1, 3), ",", "")) <> 0 Then
                picImg.PSet (X, Y), vbBlack
            End If
            str_Line = Mid(str_Line, 4)
            If str_Line = "" Then Exit For
        Next
        If str_Line = "" Then Exit For
    Next
    '---
    Dim strA As String, intLoop As Integer
    Dim intA(18) As Integer
    Dim X1 As Currency, X2 As Currency, Y1 As Currency, Y2 As Currency
    strA = strFlag ' "022,025,048,035,118,030,068,078,090,070,090,118,029,071,051,002,002,002"
                     '022 025 048 035 118 030 068 078 090 070 090 118 029 063 038 002 002 002
    For intLoop = LBound(Split(strA, ",")) To UBound(Split(strA, ","))
        intA(intLoop) = Split(strA, ",")(intLoop)
    Next
    
    '第一块
    X1 = intA(Diff.NoL): Y1 = 127: X2 = intA(Diff.NoL): Y2 = 127 - intA(Diff.NL)
    picImg.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.NoL): Y1 = 127 - intA(Diff.NL): X2 = intA(Diff.LMU): Y2 = 127 - intA(Diff.NL)
    picImg.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.LMU): Y1 = 127 - intA(Diff.NL): X2 = intA(Diff.LMD): Y2 = 127
    picImg.Line (X1, Y1)-(X2, Y2), vbRed
    
    X1 = intA(Diff.LL): Y1 = 127: X2 = intA(Diff.LL): Y2 = 127 - intA(Diff.NL)
    picImg.Line (X1, Y1)-(X2, Y2), vbMagenta
    
    X1 = intA(Diff.AL): Y1 = 127: X2 = intA(Diff.AL): Y2 = 127 - intA(Diff.NL)
    picImg.Line (X1, Y1)-(X2, Y2), vbMagenta
    '第二块
    picImg.Line (intA(Diff.RM), 127)-(intA(Diff.RM), 127 - intA(Diff.RMN)), vbRed
    picImg.Line (intA(Diff.LMN), 127 - intA(Diff.NL))-(intA(Diff.MN), 127 - intA(Diff.RMN)), vbRed
    picImg.Line (intA(Diff.MN), 127 - intA(Diff.RMN))-(127, 127 - intA(Diff.RMN)), vbRed
    
    '第三块
    picImg.Line (intA(Diff.NoN), 127 - intA(Diff.NL))-(intA(Diff.NoN), 127 - intA(Diff.NE)), vbRed
    picImg.Line (intA(Diff.NoN), 127 - intA(Diff.NE))-(127, 127 - intA(Diff.NE)), vbRed
    picImg.Line (intA(Diff.LN), 127 - intA(Diff.NL))-(intA(Diff.LN), 127 - intA(Diff.NE)), vbMagenta
    picImg.Line (intA(Diff.RN), 127 - intA(Diff.RMN))-(intA(Diff.RN), 127 - intA(Diff.NE)), vbRed
    
    '第四块
    picImg.Line (intA(Diff.NoE), 127 - intA(Diff.NE))-(intA(Diff.NoE), 127 - 127), vbRed
    '虚线
    picImg.Line (intA(Diff.NoE), 127 - (intA(Diff.NE) + intA(Diff.FNE)))-(127, 127 - (intA(Diff.NE) + intA(Diff.FNE))), vbBlue
    picImg.Line (intA(Diff.NoN), 127 - (intA(Diff.NL) + intA(Diff.FLN)))-(intA(Diff.LMN), 127 - (intA(Diff.NL) + intA(Diff.FLN))), vbBlue
    picImg.Line (intA(Diff.LMN), 127 - (intA(Diff.NL) + intA(Diff.FLN)))-(intA(Diff.MN), 127 - (intA(Diff.RMN) + intA(Diff.FMN))), vbBlue

    SavePicture picImg.Image, strFilename
    DrawP60 = True
End Function


Public Function DrawDiff5AL(ByVal strCode As String, ByVal strFilename As String, ByVal strFlag As String) As Boolean
    Dim X As Integer, Y As Integer, str_in As String
    Dim strBit As String
    
    str_in = strCode
    
    picImg.Scale (0, 0)-(128, 128)
    picImg.BackColor = vbWhite
    
    For Y = 0 To 127
        For X = 0 To 127
            strBit = Left(str_in, 1)
            If Val(strBit) <> 0 Then
                picImg.PSet (X, Y), vbBlack
            End If
            str_in = Mid(str_in, 2)
            If str_in = "" Then Exit For
        Next
        If str_in = "" Then Exit For
    Next
    
    '---
    Dim strA As String, intLoop As Integer
    Dim intA(18) As Integer
    Dim X1 As Currency, X2 As Currency, Y1 As Currency, Y2 As Currency
    strA = strFlag ' "022,025,048,035,118,030,068,078,090,070,090,118,029,071,051,002,002,002"
    For intLoop = LBound(Split(strA, ",")) To UBound(Split(strA, ","))
        intA(intLoop) = Split(strA, ",")(intLoop)
    Next
    picImg.DrawMode = vbCopyPen
    picImg.DrawStyle = vbSolid
    picImg.DrawWidth = 1.5
    '第一块
    X1 = 0: Y1 = 127 - intA(Diff.NoL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.NoL)
    picImg.Line (X1, Y1)-(X2, Y2), vbBlue
    
    X1 = intA(Diff.NL): Y1 = 127 - intA(Diff.NoL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.LMU)
    picImg.Line (X1, Y1)-(X2, Y2), vbBlue
    
    X1 = intA(Diff.NL): Y1 = 127 - intA(Diff.LMU): X2 = 0: Y2 = 127 - intA(Diff.LMD)
    picImg.Line (X1, Y1)-(X2, Y2), vbBlue
    

    '第二块
    picImg.Line (0, 127 - intA(Diff.RM))-(intA(Diff.RMN), 127 - intA(Diff.RM)), vbBlue
    picImg.Line (intA(Diff.NL), 127 - intA(Diff.LMN))-(intA(Diff.RMN), 127 - intA(Diff.MN)), vbBlue
    picImg.Line (intA(Diff.RMN), 127 - intA(Diff.MN))-(intA(Diff.RMN), 127 - 127), vbBlue
    
    '第三块
    picImg.Line (intA(Diff.NL), 127 - intA(Diff.NoN))-(intA(Diff.NE), 127 - intA(Diff.NoN)), vbBlue
    picImg.Line (intA(Diff.NE), 127 - intA(Diff.NoN))-(intA(Diff.NE), 127 - 127), vbBlue
   
    picImg.Line (intA(Diff.RMN), 127 - intA(Diff.RN))-(intA(Diff.NE), 127 - intA(Diff.RN)), vbBlue
    
    '第四块
    picImg.Line (intA(Diff.NE), 127 - intA(Diff.NoE))-(127, 127 - intA(Diff.NoE)), vbBlue
    
    picImg.DrawWidth = 1
    '外框
    picImg.Line (0, 0)-(128, 0), vbBlack
    picImg.Line (0, 0)-(0, 128), vbBlack
    picImg.Line (127, 0)-(127, 128), vbBlack
    picImg.Line (0, 127)-(127, 127), vbBlack
    '虚线
    picImg.DrawStyle = vbDot
    
    X1 = 0: Y1 = 127 - intA(Diff.LL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.LL)
    picImg.Line (X1, Y1)-(X2, Y2), vbBlack
    
    X1 = 0: Y1 = 127 - intA(Diff.AL): X2 = intA(Diff.NL): Y2 = 127 - intA(Diff.AL)
    picImg.Line (X1, Y1)-(X2, Y2), vbBlack
    
    picImg.Line (intA(Diff.NL), 127 - intA(Diff.LN))-(intA(Diff.NE), 127 - intA(Diff.LN)), vbBlack
    
    picImg.Line ((intA(Diff.NE) + intA(Diff.FNE)), 127 - intA(Diff.NoE))-((intA(Diff.NE) + intA(Diff.FNE)), 127 - 127), vbBlack
    picImg.Line ((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.NoN))-((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.LMN)), vbBlack
    picImg.Line ((intA(Diff.NL) + intA(Diff.FLN)), 127 - intA(Diff.LMN))-((intA(Diff.RMN) + intA(Diff.FMN)), 127 - intA(Diff.MN)), vbBlack
    

    SavePicture picImg.Image, strFilename
    DrawDiff5AL = True
End Function


Public Function Draw_YDA_111(arrHigh() As Double, arrVAL() As Double, arrLow() As Double, strImgPath As String, str标本号 As String) As String
    Dim X As Integer, Y As Double
    Const int_左边距 = 20, int_下边距 = 2
    picImg.Width = 5115: picImg.Height = 3795
    picImg.Scale (0, 18)-(250, 0)
    picImg.BackColor = vbWhite

'    For x = 30 To 210
'        picDraw.PSet (x + int_左边距 / 2, (arrHigh(0) / x ^ 2 + arrHigh(1) / x + arrHigh(2)) + int_下边距), vbRed
'        picDraw.PSet (x + int_左边距 / 2, (arrVAL(0) / x ^ 2 + arrVAL(1) / x + arrVAL(2)) + int_下边距), vbBlack
'        picDraw.PSet (x + int_左边距 / 2, (arrLow(0) / x ^ 2 + arrLow(1) / x + arrLow(2)) + int_下边距), vbGreen
'    Next
    For X = 30 To 210
        '高限曲线
        picImg.Line (X + int_左边距, (arrHigh(0) / X ^ 2 + arrHigh(1) / X + arrHigh(2)) + int_下边距)- _
                    (X - 1 / 2 + int_左边距, (arrHigh(0) / (X - 1 / 2) ^ 2 + arrHigh(1) / (X - 1 / 2) + _
                    arrHigh(2)) + int_下边距), vbRed
        picImg.Line (X + int_左边距, (arrHigh(0) / X ^ 2 + arrHigh(1) / X + arrHigh(2)) + int_下边距)- _
                    (X + 1 / 2 + int_左边距, (arrHigh(0) / (X + 1 / 2) ^ 2 + arrHigh(1) / (X + 1 / 2) + _
                    arrHigh(2)) + int_下边距), vbRed
        '检验结果曲线
        picImg.DrawWidth = 2
        picImg.Line (X + int_左边距, (arrVAL(0) / X ^ 2 + arrVAL(1) / X + arrVAL(2)) + int_下边距)- _
                    (X - 1 / 2 + int_左边距, (arrVAL(0) / (X - 1 / 2) ^ 2 + arrVAL(1) / (X - 1 / 2) + _
                    arrVAL(2)) + int_下边距), vbBlack
        picImg.Line (X + int_左边距, (arrVAL(0) / X ^ 2 + arrVAL(1) / X + arrVAL(2)) + int_下边距)- _
                    (X + 1 / 2 + int_左边距, (arrVAL(0) / (X + 1 / 2) ^ 2 + arrVAL(1) / (X + 1 / 2) + _
                    arrVAL(2)) + int_下边距), vbBlack
        picImg.DrawWidth = 1
        
        '低限曲线
        picImg.Line (X + int_左边距, (arrLow(0) / X ^ 2 + arrLow(1) / X + arrLow(2)) + int_下边距)- _
                    (X - 1 / 2 + int_左边距, (arrLow(0) / (X - 1 / 2) ^ 2 + arrLow(1) / (X - 1 / 2) + _
                    arrLow(2)) + int_下边距), vbGreen
        picImg.Line (X + int_左边距, (arrLow(0) / X ^ 2 + arrLow(1) / X + arrLow(2)) + int_下边距)- _
                    (X + 1 / 2 + int_左边距, (arrLow(0) / (X + 1 / 2) ^ 2 + arrLow(1) / (X + 1 / 2) + _
                    arrLow(2)) + int_下边距), vbGreen
    Next
    
    'X 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)
    'X 轴刻度
    picImg.Line (int_左边距 + 10, int_下边距)-(int_左边距 + 10, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 30, int_下边距)-(int_左边距 + 30, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 70, int_下边距)-(int_左边距 + 70, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 120, int_下边距)-(int_左边距 + 120, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 150, int_下边距)-(int_左边距 + 150, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 200, int_下边距)-(int_左边距 + 200, int_下边距 + 0.3)
    
    'Y 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 14)
    picImg.Line (int_左边距, int_下边距 + 14)-(int_左边距 - 3, int_下边距 + 14 - 0.5)
    picImg.Line (int_左边距, int_下边距 + 14)-(int_左边距 + 3, int_下边距 + 14 - 0.5)
    'Y 轴刻度
    picImg.Line (int_左边距, int_下边距 + 5)-(int_左边距 + 3, int_下边距 + 5)
    picImg.Line (int_左边距, int_下边距 + 10)-(int_左边距 + 3, int_下边距 + 10)
    picImg.Line (int_左边距, int_下边距 + 12)-(int_左边距 + 3, int_下边距 + 12)
    
'    '标题
'    With picImg
'        .CurrentX = int_左边距 + 130
'        .CurrentY = int_下边距 + 15
'        .FontSize = 12
'        .FontBold = True
'    End With
'    picImg.Print "血液粘度曲线"
    
    
    'X 周标签
    picImg.FontBold = False
    With picImg
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 0
    
    With picImg
        .CurrentX = int_左边距 + 10 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 10
    
    With picImg
        .CurrentX = int_左边距 + 30 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 30
    
    With picImg
        .CurrentX = int_左边距 + 70 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 70
    
    With picImg
        .CurrentX = int_左边距 + 120 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 120
    
    With picImg
        .CurrentX = int_左边距 + 150 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 150
    
    With picImg
        .CurrentX = int_左边距 + 200 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 200
    
    With picImg
        .CurrentX = int_左边距 + 230 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print "V"
    
    
    'Y 轴标签
    With picImg
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 5 + 0.5
        .FontSize = 10
    End With
    picImg.Print 5
    
    With picImg
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 10 + 0.5
        .FontSize = 10
    End With
    picImg.Print 10
    
    With picImg
        .CurrentX = int_左边距 - 17
        .CurrentY = int_下边距 + 12 + 0.5
        .FontSize = 10
    End With
    picImg.Print 12
    
    With picImg
        .CurrentX = 0
        .CurrentY = int_下边距 + 15 + 0.5
        .FontSize = 10
    End With
    picImg.Print "mpa·s"
    
    If Dir(strImgPath & "\YDA-111_" & str标本号 & ".JPG") <> "" Then
        Kill strImgPath & "\YDA-111_" & str标本号 & ".JPG"
    End If
    Draw_YDA_111 = strImgPath & "\YDA-111_" & str标本号 & ".JPG"
    
    SavePic picImg.Image, Draw_YDA_111, "JPG"
    
    'Call SavePicture(picImg.Image, Draw_YDA_111)
End Function

'clsLISDev_File_Fascow   2010D 血流变曲线图形
Public Function Draw_2010D(arrHigh() As Double, arrVAL() As Double, arrLow() As Double, arrNianDu() As Double, strImgPath As String, str标本号 As String) As String
    Dim intLoop As Integer
    Dim X As Integer
   
    Dim dblAA As Double, dblBB As Double, dblC As Double
    Dim dblAA1 As Double, dblBB1 As Double, dblC1 As Double
    Dim dblAA2 As Double, dblBB2 As Double, dblC2 As Double
    
    Const int_左边距 = 20, int_下边距 = 2
    
    Call ImageCalc(arrLow(0), arrVAL(0), arrHigh(0), dblAA2, dblBB2, dblC2)
    Call ImageCalc(arrLow(1), arrVAL(1), arrHigh(1), dblAA, dblBB, dblC)
    Call ImageCalc(arrLow(2), arrVAL(2), arrHigh(2), dblAA1, dblBB1, dblC1)
    
    picImg.Width = 5115: picImg.Height = 3795
    picImg.Scale (0, 36)-(250, 0)
    picImg.BackColor = vbWhite
    
    For intLoop = 15 To 200
        X = intLoop
        
        '高限曲线
        picImg.Line (X + int_左边距, (dblAA1 * Exp(dblBB1 * X ^ dblC1)) + int_下边距)- _
                    (X - 1 / 2 + int_左边距, (dblAA1 * Exp(dblBB1 * (X - 1 / 2) ^ dblC1)) + int_下边距), vbRed
        picImg.Line (X + int_左边距, (dblAA1 * Exp(dblBB1 * X ^ dblC1)) + int_下边距)- _
                    (X + 1 / 2 + int_左边距, (dblAA1 * Exp(dblBB1 * (X + 1 / 2) ^ dblC1)) + int_下边距), vbRed
        
       '检验结果曲线
        picImg.DrawWidth = 2
        picImg.Line (X + int_左边距, (dblAA * Exp(dblBB * X ^ dblC)) + int_下边距)- _
                    (X - 1 + int_左边距, (dblAA * Exp(dblBB * (X - 1) ^ dblC)) + int_下边距), vbBlack
        picImg.Line (X + int_左边距, (dblAA * Exp(dblBB * X ^ dblC)) + int_下边距)- _
                    (X + 1 + int_左边距, (dblAA * Exp(dblBB * (X + 1) ^ dblC)) + int_下边距), vbBlack
        picImg.DrawWidth = 1
        
        '低限曲线
        picImg.Line (X + int_左边距, (dblAA2 * Exp(dblBB2 * X ^ dblC2)) + int_下边距)- _
                    (X - 1 / 2 + int_左边距, (dblAA2 * Exp(dblBB2 * (X - 1 / 2) ^ dblC2)) + int_下边距), vbGreen
        picImg.Line (X + int_左边距, (dblAA2 * Exp(dblBB2 * X ^ dblC2)) + int_下边距)- _
                    (X + 1 / 2 + int_左边距, (dblAA2 * Exp(dblBB2 * (X + 1 / 2) ^ dblC2)) + int_下边距), vbGreen
                    
        '血浆粘度
        picImg.Line (int_左边距, arrNianDu(1) + int_下边距)-(int_左边距 + 200, arrNianDu(1) + int_下边距)
        If intLoop Mod 20 = 0 Then
            picImg.Line (X + int_左边距, arrNianDu(0) + int_下边距)-(X + int_左边距 - 2.5, arrNianDu(0) + int_下边距)
            picImg.Line (X + int_左边距, arrNianDu(0) + int_下边距)-(X + int_左边距 + 2.5, arrNianDu(0) + int_下边距)
            
            picImg.Line (X + int_左边距, arrNianDu(2) + int_下边距)-(X + int_左边距 - 2.5, arrNianDu(2) + int_下边距)
            picImg.Line (X + int_左边距, arrNianDu(2) + int_下边距)-(X + int_左边距 + 2.5, arrNianDu(2) + int_下边距)
            
        End If
    Next
    
     'X 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)
    
    'Y 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 34)
    picImg.Line (int_左边距, int_下边距 + 34)-(int_左边距 - 4, int_下边距 + 34 - 0.5)
    picImg.Line (int_左边距, int_下边距 + 34)-(int_左边距 + 4, int_下边距 + 34 - 0.5)


    'X 周标签，刻度线
    For intLoop = 0 To 200 Step 20
        picImg.Line (int_左边距 + intLoop, int_下边距)-(int_左边距 + intLoop, int_下边距 + 0.5)
        picImg.FontBold = False
        With picImg
            .CurrentX = int_左边距 - 8 + intLoop
            .CurrentY = int_下边距 - 0.3
            .FontSize = 10
        End With
        picImg.Print intLoop
    Next
    
    With picImg
        .CurrentX = int_左边距 + 210 - 8
        .CurrentY = int_下边距 + 2
        .FontSize = 10
    End With
    picImg.Print "(1/S)"
    
    
    'Y 轴标签
    For intLoop = 6 To 30 Step 6
        picImg.Line (int_左边距, int_下边距 + intLoop)-(int_左边距 + 3, int_下边距 + intLoop)
        With picImg
            .CurrentX = int_左边距 - 17
            .CurrentY = int_下边距 + intLoop + 0.5
            .FontSize = 10
        End With
        picImg.Print intLoop
    Next
    
    With picImg
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    picImg.Print "(mpas)"
    
    If Dir(strImgPath & "\2010D_" & str标本号 & ".JPG") <> "" Then
        Kill strImgPath & "\2010D_" & str标本号 & ".JPG"
    End If
    Draw_2010D = strImgPath & "\2010D_" & str标本号 & ".JPG"
    
    SavePic picImg.Image, Draw_2010D, "JPG"

End Function

Private Sub ImageCalc(dblQlow As Double, dblQmID As Double, dblQhigh As Double, dblAA As Double, dblBB As Double, dblC As Double)
    Dim dblE As Double
    Dim dblC1 As Double
    Dim dblC2 As Double
    Dim dblD As Double
    Dim dblY As Double
    Dim dblY1 As Double
    Dim dblY2 As Double
 

    dblE = 0.0000001
    dblC1 = 1
    dblC2 = -5
    
    dblD = Log(dblQlow / dblQmID) / Log(dblQlow / dblQhigh)
    dblY1 = (1 - (30 / 3) ^ dblC1) / (1 - (200 / 3) ^ dblC1) - dblD
    dblY2 = (1 - (30 / 3) ^ dblC2) / (1 - (200 / 3) ^ dblC2) - dblD
    
    While Abs(dblY2 - dblY1) > dblE
        dblC = (dblC1 + dblC2) / 2
        dblY = (1 - (30 / 3) ^ dblC) / (1 - (200 / 3) ^ dblC) - dblD
        
        If dblY * dblY1 > 0 Then
            dblY1 = dblY
            dblC1 = dblC
        Else
            dblY2 = dblY
            dblC2 = dblC
        End If
    Wend
    
    dblBB = Log(dblQlow / dblQmID) / (3 ^ dblC - 30 ^ dblC)
    dblAA = dblQhigh / Exp(dblBB * 200 ^ dblC)
End Sub

Public Function Draw_SA6000(strLow As String, strVal As String, strHigh As String, strImgPath As String, str标本号 As String) As String
    Dim intLoop As Integer
    Dim X As Single
    Dim sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single
    Dim varTmp As Variant
    Const int_左边距 = 20, int_下边距 = 3

    picImg.Width = 5115: picImg.Height = 3795
    picImg.Scale (0, 42)-(255, 0)
    picImg.BackColor = vbWhite

    For intLoop = 1 To 220
        X = intLoop
        '高限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
    
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 - 7 / 50 / 2) + int_下边距), vbRed
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 + 7 / 50 / 2) + int_下边距), vbRed

       '检验结果曲线
        sngX1 = 1: sngY1 = Split(Split(strVal, ",")(1), "-")(1): sngX2 = 200: sngY2 = Split(Split(strVal, ",")(4), "-")(1)
        
        picImg.DrawWidth = 1
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 - 7 / 50 / 2) + int_下边距), vbGreen
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 + 7 / 50 / 2) + int_下边距), vbGreen
        picImg.DrawWidth = 1

        '低限曲线
        sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
        
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 - 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 - 7 / 50 / 2) + int_下边距), vbMagenta
        picImg.Line (X + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 3 + 7 / 50 / 2 + int_左边距, GetY_SA6000(sngX1, sngY1, sngX2, sngY2, X + 3 + 7 / 50 / 2) + int_下边距), vbMagenta
    Next

     'X 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)

    'Y 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 38)
    picImg.Line (int_左边距, int_下边距 + 38)-(int_左边距 - 4, int_下边距 + 38 - 0.5)
    picImg.Line (int_左边距, int_下边距 + 38)-(int_左边距 + 4, int_下边距 + 38 - 0.5)


    'X 轴刻度
    picImg.Line (int_左边距 + 3, int_下边距)-(int_左边距 + 3, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 10, int_下边距)-(int_左边距 + 10, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 30, int_下边距)-(int_左边距 + 30, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 100, int_下边距)-(int_左边距 + 100, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 200, int_下边距)-(int_左边距 + 200, int_下边距 + 0.5)
    
    picImg.FontBold = False
    With picImg
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 1
    
    With picImg
        .CurrentX = int_左边距 + 3 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 3
    
    With picImg
        .CurrentX = int_左边距 + 10 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 10
    
    With picImg
        .CurrentX = int_左边距 + 30 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 30
    
    With picImg
        .CurrentX = int_左边距 + 100 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 100
    
    With picImg
        .CurrentX = int_左边距 + 200 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 200
    
    With picImg
        .CurrentX = int_左边距 + 220 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print "(1/s)"

    'Y 轴标签
    For intLoop = 0 To 36 Step 2
        picImg.Line (int_左边距, int_下边距 + intLoop)-(int_左边距 + 3, int_下边距 + intLoop)
        With picImg
            .CurrentX = int_左边距 - 17
            .CurrentY = int_下边距 + intLoop + 0.5
            .FontSize = 10
        End With
        picImg.Print intLoop
    Next

    With picImg
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    picImg.Print "(mpas)"

    If Dir(strImgPath & "\SA6000_" & str标本号 & ".jpg") <> "" Then
        Kill strImgPath & "\SA6000_" & str标本号 & ".jpg"
    End If
    Draw_SA6000 = strImgPath & "\SA6000_" & str标本号 & ".jpg"

    SavePic picImg.Image, Draw_SA6000, "jpg"
End Function

Private Function GetY_SA6000(sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single, sngX As Single) As Single
    Dim dblA As Double, dblB As Double, sngY As Single
    
    dblB = (Sqr(sngY1) - Sqr(sngY2)) / (Sqr(1 / sngX1) - Sqr(1 / sngX2))
    dblA = Sqr(sngY1) - dblB * Sqr(1 / sngX1)
    sngY = dblA ^ 2 + dblB ^ 2 / sngX + 2 * dblA * dblB * Sqr(1 / sngX)
    GetY_SA6000 = sngY
End Function

Public Function Draw_ZL6000C(strLow As String, strVal As String, strHigh As String, strImgPath As String, str标本号 As String) As String
    Dim intLoop As Integer
    Dim X As Single
    Dim sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single
    Dim varTmp As Variant
    Dim Z As Single
    Dim U As Single
    Const int_左边距 = 20, int_下边距 = 3

    picImg.Width = 5115: picImg.Height = 3795
    picImg.Scale (0, 42)-(255, 0)
    picImg.BackColor = vbWhite

    For intLoop = 20 To 220
        X = intLoop
            '高限曲线
'        If x = 20 Then
'            sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
'            Z = x + 7 / 8 + int_左边距 - 15
'            U = GetY_SA6000_541(sngX1, sngY1, sngX2, sngY2, x) + int_下边距
'            picImg.Line (x + int_左边距 - 15, GetY_SA6000_541(sngX1, sngY1, sngX2, sngY2, x) + int_下边距)-(Z, U), vbRed
'        Else
'            sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
'            picImg.Line (Z, U)-(x + 7 / 8 + int_左边距 - 15, GetY_SA6000_541(sngX1, sngY1, sngX2, sngY2, x) + int_下边距), vbRed
'            Z = x + 7 / 8 + int_左边距 - 15
'            U = GetY_SA6000_541(sngX1, sngY1, sngX2, sngY2, x) + int_下边距
'        End If
'
            
            
            '高限曲线
            sngX1 = 1: sngY1 = Split(Split(strLow, ";")(1), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(1), ",")(0)
            picImg.Line (X + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 7 / 8 + 0.2 + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距), vbRed

           '检验结果曲线
            sngX1 = 1: sngY1 = Split(Split(strVal, ",")(1), "-")(1): sngX2 = 200: sngY2 = Split(Split(strVal, ",")(4), "-")(1)
            picImg.Line (X + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距 - 0.1415926)-(X + 7 / 8 + 0.2 + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距 - 0.1415926), vbGreen

            '低限曲线
            sngX1 = 1: sngY1 = Split(Split(strLow, ";")(0), ",")(0): sngX2 = 200: sngY2 = Split(Split(strHigh, ";")(0), ",")(0)
            picImg.Line (X + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距)-(X + 7 / 8 + 0.2 + int_左边距 - 15, GetY_ZL6000C(sngX1, sngY1, sngX2, sngY2, X) + int_下边距), vbMagenta
    Next

     'X 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距 + 220, int_下边距)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 + 0.3)
    picImg.Line (int_左边距 + 220, int_下边距)-(int_左边距 + 215, int_下边距 - 0.3)

    'Y 轴线
    picImg.Line (int_左边距, int_下边距)-(int_左边距, int_下边距 + 38)
    picImg.Line (int_左边距, int_下边距 + 38)-(int_左边距 - 4, int_下边距 + 38 - 0.5)
    picImg.Line (int_左边距, int_下边距 + 38)-(int_左边距 + 4, int_下边距 + 38 - 0.5)


    'X 轴刻度
    picImg.Line (int_左边距 + 3, int_下边距)-(int_左边距 + 3, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 10, int_下边距)-(int_左边距 + 10, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 30, int_下边距)-(int_左边距 + 30, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 100, int_下边距)-(int_左边距 + 100, int_下边距 + 0.5)
    picImg.Line (int_左边距 + 200, int_下边距)-(int_左边距 + 200, int_下边距 + 0.5)
    
    picImg.FontBold = False
    With picImg
        .CurrentX = int_左边距 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 1
    
    With picImg
        .CurrentX = int_左边距 + 30 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 3
    
    With picImg
        .CurrentX = int_左边距 + 60 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 10
    
    With picImg
        .CurrentX = int_左边距 + 90 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 30
    
    With picImg
        .CurrentX = int_左边距 + 140 - 8
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 100
    
    With picImg
        .CurrentX = int_左边距 + 200 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print 200
    
    With picImg
        .CurrentX = int_左边距 + 220 - 18
        .CurrentY = int_下边距 - 0.3
        .FontSize = 10
    End With
    picImg.Print "(1/s)"

    'Y 轴标签
    For intLoop = 0 To 36 Step 4
        picImg.Line (int_左边距, int_下边距 + intLoop)-(int_左边距 + 3, int_下边距 + intLoop)
        With picImg
            .CurrentX = int_左边距 - 17
            .CurrentY = int_下边距 + intLoop + 0.5
            .FontSize = 10
        End With
        picImg.Print intLoop
    Next

    With picImg
        .CurrentX = int_左边距 + 1
        .CurrentY = int_下边距 + 32 + 0.5
        .FontSize = 10
    End With
    picImg.Print "(mpas)"

    If Dir(strImgPath & "\ZL6000C_" & str标本号 & ".jpg") <> "" Then
        Kill strImgPath & "\ZL6000C_" & str标本号 & ".jpg"
    End If
    Draw_ZL6000C = strImgPath & "\ZL6000C_" & str标本号 & ".jpg"

    SavePic picImg.Image, Draw_ZL6000C, "jpg"
End Function

Private Function GetY_ZL6000C(sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single, sngX As Single) As Single
    Dim dblA As Double, dblB As Double, sngY As Single
    
    dblB = (Sqr(sngY1) - Sqr(sngY2)) / (Sqr(1 / sngX1) - Sqr(1 / sngX2))
    dblA = Sqr(sngY1) - dblB * Sqr(1 / sngX1)
    sngY = dblA ^ 2 + dblB ^ 2 / sngX + 18 * dblA * dblB * Sqr(1 / sngX)
    GetY_ZL6000C = sngY / 1.88888888888888
End Function
