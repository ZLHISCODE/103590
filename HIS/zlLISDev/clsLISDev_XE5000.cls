VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XE5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:XE-5000
'厂商:Sysmex
'作者:陈东
'用户:商丘第一人民医院
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
    Dim lng_Stx As Long, lng_Etx As Long, bln_QC As Boolean
    Dim str_标本号 As String, str_日期 As String, int_Index As Integer
    Dim str_Code As String, str_Line As String, str_Item As String
    Dim strD1U As String, strTmp As String
    Dim str_Path As String, str_File  As String, str_imgName As String, str_RetuImage As String
    Dim bln_Rec As Boolean, strBacordID As String
    strCmd = Chr(&H6)
    mstrCode = mstrCode & strSource
    
    
        
    If Not mstrCode Like "*DI*" & Chr(&H3) & "DR*" & Chr(&H3) & "*" Then Exit Sub
    lng_Stx = InStr(mstrCode, "DI")
    lng_Etx = InStr(lng_Stx, mstrCode, Chr(&H3))
    str_Code = Mid(mstrCode, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    lng_Stx = InStr(mstrCode, "DR")
    lng_Etx = InStr(lng_Stx, mstrCode, Chr(&H3))
    strResult = ""
    
    Do While lng_Stx > 0 And lng_Etx > 0
        str_Code = str_Code & Mid(mstrCode, lng_Stx + 1, lng_Etx - lng_Stx - 1)
        mstrCode = Mid(mstrCode, lng_Etx + 1)
        
        
        If Left(str_Code, 2) = "DI" Then
            bln_QC = False
        ElseIf ((Left(str_Code, 3) = "D1C") Or (Left(str_Code, 3) = "D2C")) Then
            bln_QC = True
        End If
        If bln_QC = False And Left(str_Code, 2) = "DI" Then
            str_标本号 = Trim(Mid(str_Code, 65, 15))
            If Len(str_标本号) <= 0 Then Exit Do
            If Val(str_标本号) <= 0 Then
                Do While Left(str_标本号, 1) = "0"
                    str_标本号 = Mid(str_标本号, 2)
                Loop
                bln_QC = True
            End If

            str_日期 = Mid(str_Code, 43, 4)
            str_日期 = str_日期 & "-" & Mid(str_Code, 47, 2)
            str_日期 = str_日期 & "-" & Mid(str_Code, 49, 2)
            
            str_日期 = str_日期 & " " & Mid(str_Code, 51, 2)
            str_日期 = str_日期 & ":" & Mid(str_Code, 53, 2)
            str_日期 = str_日期 & ":" & Mid(str_Code, 55, 2)
            
            If IsDate(str_日期) Then
                str_日期 = Format(CDate(str_日期), "yyyy-MM-dd HH:mm:ss")
            Else
                str_日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
            strD1U = Mid(str_Code, InStr(str_Code, "D1U"), 132)
            str_Line = Mid(str_Code, InStr(str_Code, "D2U"), 216)
            str_Line = str_Line & Mid(str_Code, InStr(str_Code, "D6U"), 255)
            str_Item = ""
            
            If Len(strD1U) >= 132 Then
                strTmp = Trim(GetValues(Mid(strD1U, 93, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC-BF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 100, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RBC-BF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 113, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|MN#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 120, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|MN%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 125, 7), 3))
                If strTmp <> "" Then str_Item = str_Item & "|PMN#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(strD1U, 132, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PMN%|" & strTmp
            End If
            If Len(str_Line) >= 471 Then
                
                strTmp = Trim(GetValues(Mid(str_Line, 11, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 17, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 22, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|HGB|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 27, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HCT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 32, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MCV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 37, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MCH|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 42, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|MCHC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 47, 5), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PLT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 52, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LYMPH%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 57, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MONO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 62, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 67, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|EO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 72, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|BASO%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 77, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|LYMPH#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 83, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|MONO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 89, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 95, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|EO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 101, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|BASO#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 107, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RDW-CV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 112, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RDW-SD|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 117, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|PDW|" & strTmp
                 
                strTmp = Trim(GetValues(Mid(str_Line, 122, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MPV|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 127, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|P-LCR|" & strTmp
                
                bln_Rec = Not (Val(Mid(str_Line, 132, 5)) / 10 ^ 3 = 0)
                
                strTmp = Trim(GetValues(Mid(str_Line, 132, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RET%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 137, 5), 5))
                If strTmp <> "" Then str_Item = str_Item & "|RET#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 142, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IRF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 147, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 152, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 157, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HFR|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 162, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|PCT|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 167, 6), 1))
                If strTmp <> "" Then str_Item = str_Item & "|NRBC%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 173, 6), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NRBC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 179, 6), 3))
                If strTmp <> "" Then str_Item = str_Item & "|IG#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 185, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IG%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 190, 6), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HPC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 196, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RET-He|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 201, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IPF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 228, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC-B|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 234, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|WBC-D|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 240, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT#&|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 246, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT%&|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 251, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|LYMP#&|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 257, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LYMP%&|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 262, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|HFLC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 268, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HFLC%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 273, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|Area#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 279, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|Area%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 284, 5), 3))
                If strTmp <> "" Then str_Item = str_Item & "|NRBC+W|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 289, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|NEUT-X|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 295, 4), 3))
                If strTmp <> "" Then str_Item = str_Item & "|RBC-O|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 300, 4), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PLT-O|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 305, 4), 1))
                If strTmp <> "" Then str_Item = str_Item & "|PLT-I|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 310, 6), 1))
                If strTmp <> "" Then str_Item = str_Item & "|HF-BF#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 317, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HF-BFB|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 322, 6), 1))
                If strTmp <> "" Then str_Item = str_Item & "|TC-BF#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 340, 4), 4))
                If strTmp <> "" Then str_Item = str_Item & "|FRC#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 345, 4), 3))
                If strTmp <> "" Then str_Item = str_Item & "|FRC%|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 350, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RBC-He|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 355, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|Delta-He|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 360, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RET-Y|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 365, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|RBC-Y|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 370, 5), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IRF-Y|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 376, 4), 1))
                If strTmp <> "" Then str_Item = str_Item & "|RPI|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 381, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|LScRBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 386, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|HScRBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 391, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MicroRBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 396, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|MacroRBC|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 401, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|H-IPF|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 406, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|IPF#|" & strTmp
                
                strTmp = Trim(GetValues(Mid(str_Line, 411, 4), 2))
                If strTmp <> "" Then str_Item = str_Item & "|PLT-X|" & strTmp
                
            End If
            
            
            If str_标本号 <> "" And str_Item <> "" Then
            
                If bln_QC Then
                    strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|1" & str_Item
                Else
                    str_标本号 = Val(str_标本号)
                    If Len(str_标本号) >= 12 Then
                        Call ReadConf(App.Path & "\XE5000.ini", "B_ID", strBacordID)
                        If strBacordID = "" Then
                            '塘厦特殊处理，条码第3位后为医嘱ID，当做标本号返回。
                            If Val(Mid(str_标本号, 3)) <> 0 Then
                                strResult = strResult & "||" & str_日期 & "|" & Val(Mid(str_标本号, 3)) & "^^" & str_标本号 & "| |血清|0" & str_Item
                            Else
                                strResult = strResult & "||" & str_日期 & "|" & "^^" & str_标本号 & "| |血清|0" & str_Item
                            End If
                        Else
                            strResult = strResult & "||" & str_日期 & "|" & "^^" & str_标本号 & "| |血清|0" & str_Item
                        End If
                    Else
                        strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_Item
                    End If
                End If
            Else
                Exit Sub
            End If
            
            '图形数据
            str_Path = App.Path & "\Gram"
            If Dir(str_Path, vbDirectory) = "" Then
                MkDir str_Path
            End If
            If InStr(str_Code, "D3U") > 0 Then
                str_imgName = "HRBC"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3U"))
            End If
            If InStr(str_Code, "D4U") > 0 Then
                str_imgName = "HPLT"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4U"))
            End If
            If InStr(str_Code, "D5U") > 0 Then
                str_imgName = "WBC DISCRI"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5U"))
            End If
            If InStr(str_Code, "D1G") > 0 Then
                str_imgName = "SDIFF"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D1G"))
            End If
            If InStr(str_Code, "D2G") > 0 Then
                str_imgName = "SBASO"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D2G"))
            End If
            If InStr(str_Code, "D3G") > 0 Then
                str_imgName = "SRET"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3G"))
            End If
            If InStr(str_Code, "D4G") > 0 Then
    '            If bln_Rec Then
                    str_imgName = "SPLT"
                    str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                    str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4G"))
    '            End If
            End If
            If InStr(str_Code, "D5G") > 0 Then
                str_imgName = "SIMI"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5G"))
            End If
            If InStr(str_Code, "D6G") > 0 Then
                str_imgName = "SRET-E"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D6G"))
            End If
            If InStr(str_Code, "D7G") > 0 Then
                str_imgName = "SNRBC"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D7G"))
            End If
            If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
        Else
            '没数据，暂时不处理
        End If
        If mstrCode Like "*DI*" & Chr(&H3) & "DR*" & Chr(&H3) & "*" Then
            lng_Stx = InStr(mstrCode, "DI")
            lng_Etx = InStr(lng_Stx, mstrCode, Chr(&H3))
            str_Code = Mid(mstrCode, lng_Stx + 1, lng_Etx - lng_Stx - 1)
            lng_Stx = InStr(mstrCode, "DR")
            lng_Etx = InStr(lng_Stx, mstrCode, Chr(&H3))
        Else
             lng_Stx = 0
             lng_Etx = 0
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

Private Function Save_image(ByVal str_img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_img = "HRBC" Or str_img = "HPLT" Or str_img = "WBC DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Mid(str_Data, lng_Start + 29, 4)
        T_ImgInfo.lngupper = Mid(str_Data, lng_Start + 33, 4)
        T_ImgInfo.lngmaxx = Mid(str_Data, lng_Start + 16, 3)
        T_ImgInfo.lngmaxy = Mid(str_Data, lng_Start + 19, 3)
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_img = "PLT DISCRI", 40, 46)
        
        
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
             Save_image = "^" & str_img & ";105;1;" & str_FileName
        End If
    Else
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 1
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 29, lng_Lens)
        If uf_xescat(str_ImgData, lng_Lens, "c:\Temp2.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_img & ";104;1;" & str_FileName
        End If
    End If
    
End Function

Private Function GetValues(ByVal str_In As String, int_i As Integer) As String
    '功能:格式化返回值
    'str_in:待格式化的字符串
    'int_i:当入参是数字时,保留的小数位数
    
    If str_In = "*0000" Then
        GetValues = "----"
    ElseIf str_In = Space(5) Or str_In = Space(6) Or str_In = Space(7) Then
        GetValues = " "
    Else
        GetValues = Val(str_In) / 10 ^ int_i
    End If
    
End Function


Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intLoop, 0) Then
                        arrItems(intLoop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub
