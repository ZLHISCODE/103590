VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XN1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:XN1000
'厂商:Sysmex
'作者:蒋淑君
'用户:荧经县人民医院
'日期:2013-08-08
' base 64 encoder string
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim lng_Stx As Long, lng_Etx As Long, bln_QC As Boolean
    Dim str_标本号 As String, str_日期 As String, int_Index As Integer
    Dim str_Code As String, str_Line As String, str_Item As String
    Dim str_日期2 As String, str_质控 As String, str_标本 As String
    
    
    Dim str_Path As String, str_File  As String, str_imgName As String, str_RetuImage As String
    Dim bln_Rec As Boolean
    
    On Error GoTo errHand
    
    strCmd = Chr(&H6)
    strReserved = strSource

    strReserved = Replace(strReserved, "DR", "")
    If Not strReserved Like "**" & Chr(&H3) & "*" Then Exit Sub
    lng_Stx = InStr(strReserved, "")
    lng_Etx = InStr(strReserved, Chr(&H3))
    
    str_Code = Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    strReserved = Mid(strReserved, lng_Etx + 1)
    strResult = ""
    If Left(str_Code, 2) = "DI" Then
        bln_QC = False
    ElseIf ((Left(str_Code, 3) = "D1C") Or (Left(str_Code, 3) = "D2C")) Then
        bln_QC = True
    End If

    If bln_QC = False And Left(str_Code, 2) = "DI" Then
        str_标本 = Mid(str_Code, 75, 15)
        If str_标本 Like "*QC*" Then
            str_标本号 = Split(str_标本, "-")(1)
            str_质控 = 1
        Else
            str_标本号 = str_标本
            str_质控 = 0
        End If
        If Val(str_标本号) <= 0 Then Exit Sub
        
        str_日期 = Mid(str_Code, 46, 4) '日期应该后移3位
        str_日期 = str_日期 & "-" & Mid(str_Code, 50, 2)
        str_日期 = str_日期 & "-" & Mid(str_Code, 52, 2)
        str_日期 = str_日期 & " " & Mid(str_Code, 54, 2)
        str_日期 = str_日期 & ":" & Mid(str_Code, 56, 2)
        str_日期 = str_日期 & ":" & Mid(str_Code, 58, 2)
        If IsDate(str_日期) Then
            str_日期 = Format(str_日期, "yyyy-MM-dd HH:mm:ss")
        Else
            str_日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
        End If
        
        str_日期2 = Format(str_日期, "yyyy-MM-dd")
        
        str_Line = Mid(str_Code, InStr(str_Code, "D2U"), 216)
        If Len(str_Line) >= 216 Then
            str_Item = "|WBC|" & Val(Mid(str_Line, 11, 6)) / 10 ^ 3
            str_Item = str_Item & "|RBC|" & GetValues(Mid(str_Line, 17, 5), 3)
            str_Item = str_Item & "|HGB|" & GetValues(Mid(str_Line, 22, 5), 1)
            str_Item = str_Item & "|HCT|" & GetValues(Mid(str_Line, 27, 5), 2)
            str_Item = str_Item & "|MCV|" & GetValues(Mid(str_Line, 32, 5), 2)
            str_Item = str_Item & "|MCH|" & GetValues(Mid(str_Line, 37, 5), 2)
            str_Item = str_Item & "|MCHC|" & GetValues(Mid(str_Line, 42, 5), 1)
            str_Item = str_Item & "|PLT|" & GetValues(Mid(str_Line, 47, 5), 1)
            str_Item = str_Item & "|LYMPH%|" & GetValues(Mid(str_Line, 52, 5), 2)
            str_Item = str_Item & "|MONO%|" & GetValues(Mid(str_Line, 57, 5), 2)
            str_Item = str_Item & "|NEUT%|" & GetValues(Mid(str_Line, 62, 5), 2)
            str_Item = str_Item & "|EO%|" & GetValues(Mid(str_Line, 67, 5), 2)
            str_Item = str_Item & "|BASO%|" & GetValues(Mid(str_Line, 72, 5), 2)
            str_Item = str_Item & "|LYMPH#|" & GetValues(Mid(str_Line, 77, 6), 3)
            str_Item = str_Item & "|MONO#|" & GetValues(Mid(str_Line, 83, 6), 3)
            str_Item = str_Item & "|NEUT#|" & GetValues(Mid(str_Line, 89, 6), 3)
            str_Item = str_Item & "|EO#|" & GetValues(Mid(str_Line, 95, 6), 3)
            str_Item = str_Item & "|BASO#|" & GetValues(Mid(str_Line, 101, 6), 3)
            str_Item = str_Item & "|RDW-CV|" & GetValues(Mid(str_Line, 107, 5), 2)
            str_Item = str_Item & "|RDW-SD|" & GetValues(Mid(str_Line, 112, 5), 2)
            str_Item = str_Item & "|PDW|" & GetValues(Mid(str_Line, 117, 5), 2)
            str_Item = str_Item & "|MPV|" & GetValues(Mid(str_Line, 122, 5), 2)
            str_Item = str_Item & "|P-LCR|" & GetValues(Mid(str_Line, 127, 5), 2)
            
            bln_Rec = Not (Val(Mid(str_Line, 132, 5)) / 10 ^ 3 = 0)
            str_Item = str_Item & "|RET%|" & GetValues(Mid(str_Line, 132, 5), 3)
            str_Item = str_Item & "|RET#|" & GetValues(Mid(str_Line, 137, 5), 5)
            str_Item = str_Item & "|IRF|" & GetValues(Mid(str_Line, 142, 5), 2)
            str_Item = str_Item & "|LFR|" & GetValues(Mid(str_Line, 147, 5), 2)
            str_Item = str_Item & "|MFR|" & GetValues(Mid(str_Line, 152, 5), 2)
            str_Item = str_Item & "|HFR|" & GetValues(Mid(str_Line, 157, 5), 2)
            
            str_Item = str_Item & "|PCT|" & GetValues(Mid(str_Line, 162, 5), 3)
            str_Item = str_Item & "|NRBC%|" & GetValues(Mid(str_Line, 167, 6), 1)
            str_Item = str_Item & "|NRBC#|" & GetValues(Mid(str_Line, 173, 6), 2)
            
            str_Item = str_Item & "|IG#|" & GetValues(Mid(str_Line, 179, 6), 3)
            str_Item = str_Item & "|IG%|" & GetValues(Mid(str_Line, 185, 5), 2)
            str_Item = str_Item & "|HPC#|" & GetValues(Mid(str_Line, 190, 6), 2)
            str_Item = str_Item & "|RET-He|" & GetValues(Mid(str_Line, 196, 5), 2)
            str_Item = str_Item & "|IPF|" & GetValues(Mid(str_Line, 201, 5), 2)
        End If
        If Val(str_标本号) > 0 And str_Item <> "" Then
            If str_质控 = "0" Then
                strResult = str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Item
            Else
                strResult = str_日期 & "|" & Val(str_标本号) & "| |血清|1" & str_Item
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
            str_imgName = "RBC DISCRI"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3U"))
        End If
        If InStr(str_Code, "D4U") > 0 Then
            str_imgName = "PLT DISCRI"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4U"))
        End If
        If InStr(str_Code, "D5U") > 0 Then
            str_imgName = "WBC DISCRI"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5U"))
        End If
        If InStr(str_Code, "D1G") > 0 Then
            str_imgName = "WDF SCAT"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D1G"))

        End If
        If InStr(str_Code, "D2G") > 0 Then
            str_imgName = "WNR SCAT"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D2G"))
        End If
        If InStr(str_Code, "D3G") > 0 Then
            str_imgName = "SIMI SCAT"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3G"))
        End If
        If InStr(str_Code, "D4G") > 0 Then
            If bln_Rec Then
                str_imgName = "RET SCAT"
                str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4G"))
            End If
        End If
        If InStr(str_Code, "D5G") > 0 Then
            str_imgName = "SPLT-O"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5G"))
        End If
        If InStr(str_Code, "D6G") > 0 Then
            str_imgName = "SRET-E"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D6G"))
        End If
        If InStr(str_Code, "D7G") > 0 Then
            str_imgName = "SNRBC"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D7G"))
        End If
        If InStr(str_Code, "D6U") > 0 Then
            str_imgName = "RBCY DISCRI"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D6U"))
        End If
        If InStr(str_Code, "D7U") > 0 Then
            str_imgName = "WDFY DISCRI"
            str_File = Replace(str_日期2, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D7U"))
        End If
        If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
    Else
        '没数据，暂时不处理
    End If
    
    Exit Sub
errHand:
    WriteLog "clsLISDev_XN1000", Err.Number & ":" & Err.Description, ""
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

Private Function Save_image(ByVal str_Img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_Img = "RBC DISCRI" Or str_Img = "PLT DISCRI" Or str_Img = "WBC DISCRI" Or str_Img = "RBCY DISCRI" Or str_Img = "WDFY DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Mid(str_Data, lng_Start + 29, 4)
        T_ImgInfo.lngupper = Mid(str_Data, lng_Start + 33, 4)
        T_ImgInfo.lngmaxx = Mid(str_Data, lng_Start + 16, 3)
        T_ImgInfo.lngmaxy = Mid(str_Data, lng_Start + 19, 3)
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_Img = "PLT DISCRI", 40, 46)
        
        
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
             Save_image = "^" & str_Img & ";105;1;" & str_FileName
        End If
    Else
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 1
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 29, lng_Lens)
        If uf_xescat(str_ImgData, lng_Lens, "c:\Temp2.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_Img & ";104;1;" & str_FileName
        End If
    End If
    
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '功能:格式化返回值
    'str_in:待格式化的字符串
    'int_i:当入参是数字时,保留的小数位数
    
    If str_in = "*0000" Then
        GetValues = "----"
    ElseIf str_in = "     " Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function

