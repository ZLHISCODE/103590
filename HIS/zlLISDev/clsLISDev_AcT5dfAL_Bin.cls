VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AcT5dfAL_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AcT5dfAL
'厂商:美国贝克曼库尔特 Beckman coulter
'作者:陈东
'用户:烟台山医院
Private mstrCode As String
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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varLine As Variant, intRow As Integer, strLine As String
    Dim str标本号 As String, str日期 As String, strItems As String, lngItem As Long
    Dim StrWbc As String, StrRbc As String, StrPLt As String, strWBCFLAG As String, strRBCFlag As String, strPLTFLAG As String
    Dim strBASO As String, strBasoFlag As String, strDiffFlag As String
    Dim strimg As String, strDiff As String, strValues As String
    
    mstrCode = mstrCode & strSource
    
    lngS = Instr_Bin(mstrCode, Chr(&H2))
    If lngS > 0 Then
        lngE = Instr_Bin(mstrCode, Chr(&H3), lngS)
    End If
    Do While lngS > 0 And lngE > 0
        strCode = Mid_bin(mstrCode, lngS, lngE - lngS + 1, False)
        mstrCode = Mid_bin(mstrCode, lngE + 1, , False)
        
        varLine = Split(strCode, ",0D")
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            If InStr(",!,0,1,2,3,4,5,6,7,8,@,A,B,C,-,,,+,*,(,),%,$,#,"""",.,/,", "," & Mid_bin(strLine, 1, 1)) & "," > 0 Then
                '项目
                lngItem = lngItem + 1
                strValues = Mid_bin(strLine, 3)
                If InStr(strValues, ".....") <= 0 Then
                    strItems = strItems & "|" & lngItem & "|" & Val(strValues)
                End If
            ElseIf Mid_bin(strLine, 1, 1) = """" Then
                lngItem = lngItem + 1
                strValues = Mid_bin(strLine, 3)
                If InStr(strValues, ".....") <= 0 Then
                    strItems = strItems & "|" & lngItem & "|" & Val(strValues)
                End If
            ElseIf Mid_bin(strLine, 1, 1) = "q" Then
                '日期
                strLine = Mid_bin(strLine, 1)
                If InStr(strLine, "/") > 0 Then
                    str日期 = Format(CDate(Mid(strLine, 9, 2) & "-" & Mid(strLine, 6, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd")
                Else
                    str日期 = Format(CDate(Mid(strLine, 7, 4) & "-" & Mid(strLine, 5, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd")
                End If
            ElseIf Mid_bin(strLine, 1, 1) = "u" Then
                '标本号
                str标本号 = Val(Mid_bin(strLine, 3))
                If strDiff <> "" And strDiffFlag <> "" Then strDiff = DrawBmpToFile(strDiff, "Diff", str标本号, strDiffFlag)
            ElseIf Mid_bin(strLine, 1, 1) = "v" Then
                If str标本号 = "" Then
                    str标本号 = Val(Mid_bin(strLine, 3))
                    If strDiff <> "" And strDiffFlag <> "" Then strDiff = DrawBmpToFile(strDiff, "Diff", str标本号, strDiffFlag)
                End If
            ElseIf Mid_bin(strLine, 1, 1) = "W" Then    'WBC
                StrWbc = Mid_bin(strLine, 2, , False)
                StrWbc = Graph_Text(StrWbc)
            ElseIf Mid_bin(strLine, 1, 1) = "X" Then    'RBC
                StrRbc = Mid_bin(strLine, 2, , False)
                StrRbc = Graph_Text(StrRbc)
            ElseIf Mid_bin(strLine, 1, 1) = "Y" Then    'PLT
                StrPLt = Mid_bin(strLine, 2, , False)
                StrPLt = Graph_Text(StrPLt)
            ElseIf Mid_bin(strLine, 1, 1) = "Z" Then    'BASO
                strBASO = Mid_bin(strLine, 2, , False)
                strBASO = Graph_Text(strBASO)
            ElseIf Mid_bin(strLine, 1, 1) = "[" Then    'DIFFPLOT
                strDiff = Mid_bin(strLine, 3, , False)
                'strDiff = Graph_Diff(strDiff)
            ElseIf Mid_bin(strLine, 1, 1) = "]" Then    'WBC Thresholds
                strWBCFLAG = Trim(Mid_bin(strLine, 2))
                If strWBCFLAG <> "" Then strWBCFLAG = Graph_Flag(strWBCFLAG, 130)
            ElseIf Mid_bin(strLine, 1, 1) = "^" Then    'RBC Thresholds
                strRBCFlag = Mid_bin(strLine, 2)
                If strRBCFlag <> "" Then strRBCFlag = Graph_Flag(strRBCFlag, 130)
            ElseIf Mid_bin(strLine, 1, 1) = "_" Then    'PLT Thresholds
                strPLTFLAG = Mid_bin(strLine, 2)
                If strPLTFLAG <> "" Then strPLTFLAG = Graph_Flag(strPLTFLAG, 130)
            ElseIf Mid_bin(strLine, 1, 1) = "`" Then    'BASO Thresholds
                strBasoFlag = Mid_bin(strLine, 2)
                If strBasoFlag <> "" Then strBasoFlag = Graph_Flag(strBasoFlag, 130)
            ElseIf Mid_bin(strLine, 1, 1) = "a" Then    'DIFFPLOT Thresholds
                strDiffFlag = Trim(Mid_bin(strLine, 2))
                strDiffFlag = Replace(strDiffFlag, " ", ",")
            End If
        Next
        
        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = strResult & "||" & str日期 & "|" & str标本号 & "| |全血|0" & strItems
            strimg = ""
                        
            If StrWbc <> "" Then strimg = strimg & "^" & "WBC;4;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400;" & Mid(StrWbc, 2) '& IIf(strWBCFLAG <> "", "|" & strWBCFLAG, "")
            If StrRbc <> "" Then strimg = strimg & "^" & "RBC;4;140,135,15.45-30,51.5-100,103-200;" & Mid(StrRbc, 2) '& IIf(strRBCFlag <> "", "|" & strRBCFlag, "")
            If StrPLt <> "" Then strimg = strimg & "^" & "PLT;4;70,135,8.6-2,21.5-5,43-10,86-20,129-30;" & Mid(StrPLt, 2) '& IIf(strPLTFlag <> "", "|" & strPLTFlag, "")
            If strBASO <> "" Then strimg = strimg & "^" & "BASO;4;70,135,43-150,86-300,129-450;" & Mid(strBASO, 2) '& IIf(strBasoFlag <> "", "|" & strRBCFlag, "")
            If strDiff <> "" Then strimg = strimg & "^" & "V;107;1;" & strDiff
            
            If strimg <> "" Then
                strResult = strResult & vbCrLf & Mid(strimg, 2)
            End If
        End If
        
        lngS = Instr_Bin(mstrCode, Chr(&H2))
        If lngS > 0 Then
            lngE = Instr_Bin(mstrCode, Chr(&H3), lngS)
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

Private Function Graph_Text(ByVal str_bin As String) As String
    '生成图形文本数据
    Dim lng_len As Long, lng_i As Long, str_Bit As String
    Dim str_Return As String, str_in As String
    lng_len = Len(str_bin)

    If lng_len < 3 Then
        Graph_Text = ""
        Exit Function
    Else
        str_in = str_bin
    End If

    For lng_i = 1 To lng_len / 3
        
        str_Bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        
        If IsNull(str_Bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & "," & Format((CByte("&H" & str_Bit) - 32), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

Private Function Graph_Flag(ByVal strFlag As String, lngMax As String) As String
    '生成直方图界标数据
    Dim i As Integer, strRetu As String, strIN As String, strV As String, intLoop As Integer
    strIN = strFlag
    Do While strIN <> ""
        If IsNumeric(Mid(strIN, 1, 3)) Then
            strRetu = strRetu & "," & Mid(strIN, 1, 3)
        End If
        If Mid(strIN, 4, 1) = " " Then
            strIN = Mid(strIN, 5)
        Else
            strIN = Mid(strIN, 4)
        End If
    Loop
    If strRetu <> "" Then strRetu = Mid(strRetu, 2)
    
    For i = 1 To lngMax
        strV = "0"
        For intLoop = LBound(Split(strRetu, ",")) To UBound(Split(strRetu, ","))
            If CLng(Split(strRetu, ",")(intLoop)) = i Then
                strV = lngMax
                Exit For
            End If
        Next
        Graph_Flag = Graph_Flag & "," & strV
    Next
    Graph_Flag = Replace(Graph_Flag, ",0," & lngMax, ",1," & lngMax)
    If Graph_Flag <> "" Then Graph_Flag = Mid(Graph_Flag, 2)
End Function

Private Function DrawBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str标本号 As String, strFlag As String) As String
    Dim strBinData As String
    Dim lngCount As Long
    Dim strFilename As String
    
    Dim frmPic As New frmGraph
    strBinData = HexToStr(str_bin)
    lngCount = Len(strBinData)
    
    If lngCount <= 0 Then Exit Function
    strBinData = getGraphDot(strBinData)
    strFilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    If frmPic.DrawDiff5AL(strBinData, strFilename, strFlag) Then
        DrawBmpToFile = strFilename
    End If
    Unload frmPic
    Set frmPic = Nothing
End Function


'----
Private Function HexToStr(ByVal strCode As String) As String
    Dim lngS As Long, strBit As String
    Dim StrInput As String, strReturn As String
    strReturn = ""
    StrInput = strCode
    Do While StrInput <> ""
        strBit = Mid(Left(StrInput, 3), 2)
        strReturn = strReturn & Chr("&H" & strBit)
        StrInput = Mid(StrInput, 4)
    Loop
    HexToStr = strReturn
End Function

Private Function getGraphDot(ByVal strCode As String) As String
    Dim lngStx As Long, lngEnd As Long, i As Integer, lngPos2 As Long
    Dim strIN As String, strReturn As String
    strIN = strCode
    For i = 1 To 4096 Step 2

        If Left(strIN, 2) = "00" Then
            strReturn = strReturn & "00000000"
        Else
            strReturn = strReturn & Format(H_To_B(Left(strIN, 2)), "00000000")
        End If
        strIN = Mid(strIN, 3)
    Next
    getGraphDot = strReturn
End Function



  Private Function D_To_B(ByVal Dec As Long) As String
          Do
                  D_To_B = Dec Mod 2 & D_To_B
                  Dec = Dec \ 2
          Loop While Dec
  End Function
    
  Private Function B_To_D(ByVal Bin As String) As Currency
          Dim i     As Long
          For i = 1 To Len(Bin)
                  B_To_D = B_To_D * 2 + Val(Mid(Bin, i, 1))
          Next i
  End Function
    
  Private Function H_To_B(ByVal Hex As String) As String
          Dim i     As Long
          Dim B     As String
            
          Hex = UCase(Hex)
          For i = 1 To Len(Hex)
                  Select Case Mid(Hex, i, 1)
                          Case "0":     B = B & "0000"
                          Case "1":     B = B & "0001"
                          Case "2":     B = B & "0010"
                          Case "3":     B = B & "0011"
                          Case "4":     B = B & "0100"
                          Case "5":     B = B & "0101"
                          Case "6":     B = B & "0110"
                          Case "7":     B = B & "0111"
                          Case "8":     B = B & "1000"
                          Case "9":     B = B & "1001"
                          Case "A":     B = B & "1010"
                          Case "B":     B = B & "1011"
                          Case "C":     B = B & "1100"
                          Case "D":     B = B & "1101"
                          Case "E":     B = B & "1110"
                          Case "F":     B = B & "1111"
                  End Select
          Next i
          While Left(B, 1) = "0"
                  B = Right(B, Len(B) - 1)
          Wend
          H_To_B = B
  End Function
    
  Private Function B_To_H(ByVal Bin As String) As String
          Dim i     As Long
          Dim H     As String
          If Len(Bin) Mod 4 <> 0 Then
                  Bin = String(4 - Len(Bin) Mod 4, "0") & Bin
          End If
            
          For i = 1 To Len(Bin) Step 4
                  Select Case Mid(Bin, i, 4)
                          Case "0000":     H = H & "0"
                          Case "0001":     H = H & "1"
                          Case "0010":     H = H & "2"
                          Case "0011":     H = H & "3"
                          Case "0100":     H = H & "4"
                          Case "0101":     H = H & "5"
                          Case "0110":     H = H & "6"
                          Case "0111":     H = H & "7"
                          Case "1000":     H = H & "8"
                          Case "1001":     H = H & "9"
                          Case "1010":     H = H & "A"
                          Case "1011":     H = H & "B"
                          Case "1100":     H = H & "C"
                          Case "1101":     H = H & "D"
                          Case "1110":     H = H & "E"
                          Case "1111":     H = H & "F"
                  End Select
          Next i
          B_To_H = H
  End Function
