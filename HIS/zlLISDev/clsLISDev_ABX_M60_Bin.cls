VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_M60_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ABX Micros 60
'厂商:ABX
'作者:陈东
'用户:洛阳202医院

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
    Dim varline As Variant, intRow As Integer, strLine As String
    Dim str标本号 As String, str日期 As String, strItems As String, lngItem As Long
    Dim StrWbc As String, StrRbc As String, StrPLt As String, strWBCFLAG As String, strPLTFLAG As String
    Dim strBASO As String, strDiff As String, strDiffFlag As String
    Dim strimg As String
    Dim str日期格式 As String
    Dim str仪器型号 As String
    
    On Error GoTo errH
    
    str仪器型号 = Trim(ReadIni("ABX_M60_Bin", "仪器型号", App.Path & "\ABX_M60_Bin.ini"))
    If str仪器型号 = "" Then str仪器型号 = "":   Call WriteIni("ABX_M60_Bin", "仪器型号", "", App.Path & "\ABX_M60_Bin.ini")
            
    str日期格式 = Trim(ReadIni("ABX_M60_Bin", "日期格式", App.Path & "\ABX_M60_Bin.ini"))
    If str日期格式 = "" Then str日期格式 = "":   Call WriteIni("ABX_M60_Bin", "日期格式", "", App.Path & "\ABX_M60_Bin.ini")
    
    lngS = Instr_Bin(strSource, Chr(&H2))
    lngE = Instr_Bin(strSource, Chr(&H3))
    
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid_bin(strSource, lngS, lngE - lngS + 1, False)
        strReserved = Mid_bin(strSource, lngE + 1, , False)
        varline = Split(strCode, ",0D")
        For intRow = LBound(varline) To UBound(varline)
            strLine = varline(intRow)
            If InStr(",!,0,1,2,3,4,5,6,7,8,@,A,B,C,#,%,',"""",$,&,(,),*,+,-,.,/,,,", "," & Mid_bin(strLine, 1, 1) & ",") > 0 Then
                '项目
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid_bin(strLine, 3))
            ElseIf Mid_bin(strLine, 1, 1) = """" Then
                lngItem = lngItem + 1
                strItems = strItems & "|" & lngItem & "|" & Val(Mid_bin(strLine, 3))
            ElseIf Mid_bin(strLine, 1, 1) = "q" Then
                '日期
                If str日期格式 = "" Then
                    str日期 = Format(CDate(Mid_bin(strLine, 9, 2) & "-" & Mid_bin(strLine, 6, 2) & "-" & Mid_bin(strLine, 3, 2)), "yyyy-MM-dd")
                Else
                    str日期 = Format(CDate(Mid_bin(strLine, 3, 2) & "-" & Mid_bin(strLine, 6, 2) & "-" & Mid_bin(strLine, 9, 2)), "yyyy-MM-dd")
                End If
            ElseIf Mid_bin(strLine, 1, 1) = "s" Then
                'ABX-120 太原格式
                If str标本号 = "" Then str标本号 = Trim(Mid_bin(strLine, 2))
            ElseIf Mid_bin(strLine, 1, 1) = "u" Then
                '标本号
                If Val(Trim(Replace(Mid_bin(strLine, 3), "AUTO_SID", ""))) > 0 Then
                    str标本号 = Val(Trim(Replace(Mid_bin(strLine, 3), "AUTO_SID", "")))
                End If
                If strDiff <> "" And strDiffFlag <> "" Then strDiff = DrawBmpToFile(strDiff, "Diff", str标本号, strDiffFlag)
                
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
                strDiff = Mid_bin(strLine, 2, , False)
                strDiff = Graph_Diff(strDiff)
            ElseIf Mid_bin(strLine, 1, 1) = "a" Then    'DIFFPLOT Thresholds
                strDiffFlag = Trim(Mid_bin(strLine, 2))
                strDiffFlag = Replace(strDiffFlag, " ", ",")
            End If
        Next
        
        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = str日期 & "|" & str标本号 & "| |全血|0" & strItems
            
            Select Case str仪器型号
            Case "P80"
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,15.45-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BAS;0;32,130,12-,25-100,38-,51-200,64-,77-300,90-,103-400,116-;" & Mid(strBASO, 2)
                If strDiff <> "" Then strimg = strimg & "^" & "LMNE;107;1;" & strDiff
            Case "P60"
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,9-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BASO;0;70,135,43-150,86-300,129-450;" & Mid(strBASO, 2, Len(strBASO) - 2 - 4)
                If strDiff <> "" Then strimg = strimg & "^" & "V;107;1;" & strDiff
            Case Else
                If StrWbc <> "" Then strimg = strimg & "^" & "WBC;0;140,135,16.125-50,32.25-100,64.5-200,96.75-300,129-400" & StrWbc
                If StrRbc <> "" Then strimg = strimg & "^" & "RBC;0;140,135,15.45-30,51.5-100,103-200" & StrRbc
                If StrPLt <> "" Then strimg = strimg & "^" & "PLT;0;70,135,8.6-2,21.5-5,43-10,86-20,129-30" & StrPLt
                
                If strBASO <> "" Then strimg = strimg & "^" & "BASO;4;70,135,43-150,86-300,129-450;" & Mid(strBASO, 2)
                If strDiff <> "" Then strimg = strimg & "^" & "V;107;1;" & strDiff
            End Select
            If strimg <> "" Then
                strResult = strResult & vbCrLf & Mid(strimg, 2)
            End If
        End If
    End If
    
    Exit Sub
errH:
    WriteLog "M60_Bin", Err.Description, ""
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
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String, str_in As String
    lng_len = Len(str_bin)

    If lng_len < 3 Then
        Graph_Text = ""
        Exit Function
    Else
        str_in = str_bin
    End If

    For lng_i = 1 To lng_len / 3
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format((CByte("&H" & str_bit) - 32), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

Private Function Graph_Diff(ByVal str_bin As String) As String
    '生成散点图数据
    Dim lng_len As Long, str_in As String, lng_i As Long, str_bit As String
    Dim int_Stx As Integer, int_Etx As Integer, int_Pos As Integer
    Dim int_Bit As Integer, int_i As Integer
    lng_len = Len(str_bin)
    If lng_len < 3 Then
        Graph_Diff = ""
        Exit Function
    Else
        str_in = str_bin
    End If
    
    int_Bit = 1
    str_in = Mid(str_in, 4)
   Do While Len(str_in) > 0
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        int_Stx = CByte("&H" & str_bit)
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        int_Etx = CByte("&H" & str_bit)
            
        int_Pos = (int_Stx Mod 16) * 16 + (int_Etx Mod 16)
        If int_Pos = 0 Then
            Graph_Diff = Graph_Diff & ",00,00,00,00,00,00,00,00"
        Else
            For int_i = 1 To 8
                Graph_Diff = Graph_Diff & ",0" & (int_Pos Mod 2)
                int_Pos = Int(int_Pos / 2)
            Next
        End If
         
    Loop

End Function

Private Function DrawBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str标本号 As String, strFlag As String) As String
    Dim strBinData As String
    Dim lngCount As Long
    Dim strFileName As String
    
    Dim frmPic As New frmGraph
    strBinData = str_bin
    lngCount = Len(strBinData)
    If lngCount <= 0 Then Exit Function
    
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    If frmPic.DrawP60(strBinData, strFileName, strFlag) Then
        DrawBmpToFile = strFileName
    End If
    Unload frmPic
    Set frmPic = Nothing
End Function
