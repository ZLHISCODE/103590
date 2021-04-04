Attribute VB_Name = "mdlBarCode128Auto"
Option Explicit

Private astr128Code()
Private astr128A() As Variant
Private astr128B() As Variant
Private astr128C() As Variant
Private astr128ID() As Variant
Public Function PrintBarCode128Auto(ByVal PrintObj As Object, ByVal strBarCode As String, ByVal X As Long, ByVal Y As Long, sngPrintHeight As Single, _
                            Optional ByVal intLineWidth As Integer = 2, Optional ByVal blnShowBarCodeTxt As Boolean = True) As StdPicture
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       传入图片对象画图像
    '参数                       PicObj              Picture对象(用于画图）
    '                           strBarCode          生成条码的内容
    '                           sngPrintWidth       打印时使用的固定长度单位（mm)(返回）
    '                   可选
    '                           intLineWidth        线宽，用于控制条码的宽度 默认为2
    '                           blnShowBarCodeTxt   是否显示条码内容，默认True为显示
    '返回                       Image对象
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strGetBarCode As String
    Dim intVal As Integer
    Dim lngTxtHeight As Long
    
    Dim lngX As Long
    Dim lngXWidth As Long
    

    Dim lngY As Long
    Dim lngH As Long
    
    Dim intX As Integer
    Dim intY As Integer
    
    
    Dim intScalMode As Integer
    Dim intDrawWidth As Integer
    Dim intFontSize As Integer
    Dim strFontFontName As String
    
    intScalMode = PrintObj.ScaleMode
    intDrawWidth = PrintObj.DrawWidth
    intFontSize = PrintObj.FontSize
    strFontFontName = PrintObj.FontName
    
    If intLineWidth = 0 Then intLineWidth = 2
    
    PrintObj.ScaleMode = vbPixels
    PrintObj.DrawWidth = intLineWidth
    PrintObj.FontName = "Arial"
    
    
    strGetBarCode = GetBarCode(strBarCode)
    
    
    

    lngH = sngPrintHeight / PrintObj.TwipsPerPixelY
    lngH = lngH - PrintObj.TextHeight(strBarCode)
    
    lngX = X / PrintObj.TwipsPerPixelX
    lngY = Y / PrintObj.TwipsPerPixelY
    


    For intLoop = 1 To Len(strGetBarCode)
        intVal = Mid$(strGetBarCode, intLoop, 1)
        PrintObj.Line (lngX + (intLoop * intLineWidth), lngY)-(lngX + (intLoop * intLineWidth), lngY + lngH), IIF(intVal = 1, vbBlack, vbWhite), BF
    Next
    
    If blnShowBarCodeTxt = True Then
        PrintObj.FontSize = 5 + intLineWidth
        PrintObj.CurrentX = (lngX + (intLoop * intLineWidth) / 2) - (PrintObj.TextWidth(strBarCode) / 2)
        PrintObj.CurrentY = lngY + lngH + 1
        

        PrintObj.Print strBarCode
    End If
    
    PrintObj.ScaleMode = intScalMode
    PrintObj.DrawWidth = intDrawWidth
    PrintObj.FontSize = intFontSize
    PrintObj.FontName = strFontFontName

End Function

Public Function DrawBarCode128Auto(ByVal PicObj As Object, ByVal strBarCode As String, sngPrintWidth As Single, _
                            Optional ByVal intLineWidth As Integer = 2, Optional ByVal blnShowBarCodeTxt As Boolean = True) As StdPicture
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       传入图片对象画图像
    '参数                       PicObj              Picture对象(用于画图）
    '                           strBarCode          生成条码的内容
    '                           sngPrintWidth       打印时使用的固定长度单位（mm)(返回）
    '                   可选
    '                           intLineWidth        线宽，用于控制条码的宽度 默认为2
    '                           blnShowBarCodeTxt   是否显示条码内容，默认True为显示
    '返回                       Image对象
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strGetBarCode As String
    Dim intVal As Integer
    Dim lngTxtHeight As Long
    
    If intLineWidth = 0 Then intLineWidth = 2
    
    With PicObj
        .Cls
        .BackColor = vbWhite
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .DrawWidth = intLineWidth
        .Height = 1335
    End With
    
    strGetBarCode = GetBarCode(strBarCode)
    PicObj.Width = (2 + Len(strGetBarCode) * intLineWidth) * Screen.TwipsPerPixelX
    
    If blnShowBarCodeTxt Then
        PicObj.FontSize = 12
        PicObj.FontName = "Arial"
        lngTxtHeight = PicObj.TextHeight(strBarCode)
        
        PicObj.CurrentX = (PicObj.ScaleWidth / 2) - PicObj.TextWidth(strBarCode) / 2   ' 水平坐标
        PicObj.CurrentY = PicObj.ScaleHeight - lngTxtHeight - 1 ' 垂直坐标
        PicObj.Print strBarCode
    End If
    
    For intLoop = 1 To Len(strGetBarCode)
        intVal = Mid$(strGetBarCode, intLoop, 1)
        If blnShowBarCodeTxt Then
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - lngTxtHeight - 2), IIF(intVal = 1, vbBlack, vbWhite), BF
        Else
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - 1), IIF(intVal = 1, vbBlack, vbWhite), BF
        End If
    Next
    
    PicObj.Width = (2 + intLoop * intLineWidth) * Screen.TwipsPerPixelX
    
    sngPrintWidth = PicObj.ScaleWidth / 8
    
    Set DrawBarCode128Auto = PicObj.Image
    
    '恢复缺省以避免其它作图受影响
    PicObj.ScaleMode = vbTwips
    PicObj.DrawWidth = 1
    PicObj.FontName = "宋体"
    PicObj.FontSize = 9
End Function

Private Sub initBarCode()
    '//539
    '初始化相当内容
    astr128Code = Array( _
             "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
             "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
             "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
             "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
             "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
             "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
             "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
             "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
             "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
             "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
             "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
             "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
             "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
             "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
             "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
             "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
             "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
             "11110101110", "11010000100", "11010010000", "11010011100", "1100011101011" _
             )
             
    astr128A = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "NUL", "SOH", _
             "STX", "ETX", "EOT", "ENQ", "ACK", "BEL", _
             "BS", "HT", "LF", "VT", "FF", "CR", _
             "SO", "SI", "DLE", "DC1", "DC2", "DC3", _
             "DC4", "NAK", "SYN", "ETB", "CAN", "EM", _
             "SUB", "ESC", "FS", "GS", "RS", "US", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "CODEB", "FNC4", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    
    astr128B = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "`", "a", _
             "b", "c", "d", "e", "f", "g", _
             "h", "i", "j", "k", "I", "m", _
             "n", "o", "p", "q", "r", "s", _
             "t", "u", "v", "w", "x", "y", _
             "z", "{", "|", "}", "~", "DEL", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "FNC4", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
             
    astr128C = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "CODEB", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    astr128ID = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "100", "101", _
             "102", "103", "104", "105", "106" _
             )
End Sub

Private Function FindCode(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:              按规则查找固定的字符
    '参数:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar 传入字符
    '返回:              对应的编码规则
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 4 Then
        intIndex = FindArray(strChar, astr128ID)
        FindCode = astr128Code(intIndex)
    End If
End Function

Private Function FindID(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:              按规则查找固定的字符的ID
    '参数:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar 传入字符
    '返回:              对应的编码的ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindID = Val(astr128ID(intIndex))
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindID = Val(astr128ID(intIndex))
    End If
End Function


Private Function FindArray(strChar As String, strArray() As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           查找字符在数据中的位置
    '参数           strChar 传入字符
    '               strArray() 要查找的数组
    '返回           位置index
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    FindArray = 0
    For intLoop = 0 To UBound(strArray)
        If strChar = strArray(intLoop) Then
            FindArray = intLoop
            Exit For
        End If
    Next
    
End Function

Private Function GetBarCode(strChar As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           按传入字符成生条码线规则
    '参数           strChar = 传入字符串
    '返回           条码成生的规则
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strTmp As String
    Dim intType As Integer
    Dim lngCheckCount As Long
    Dim intRow As Integer
    intType = 0
    intRow = 0
    
    For intLoop = 1 To Len(strChar) Step 2
        strTmp = Mid$(strChar, intLoop, 2)
        If Len(strTmp) = 2 And IsNumeric(strTmp) = True Then
            '128C码
            If intType = 0 Then
                GetBarCode = FindCode(3, "StartC")
                lngCheckCount = FindID(3, "StartC")
                intRow = intRow + 1
                intType = 3
            End If
            If intType <> 3 Then
                '转为128C码
                GetBarCode = GetBarCode & FindCode(1, "CODEC")
                lngCheckCount = lngCheckCount + intRow * FindID(1, "CODEC")
                intRow = intRow + 1
            End If
            
            
            GetBarCode = GetBarCode & FindCode(3, Val(strTmp))
            lngCheckCount = lngCheckCount + intRow * FindID(3, Val(strTmp))
            intRow = intRow + 1
        
            intType = 3
        Else
             '128A码
            If intType = 0 Then
                GetBarCode = FindCode(1, "StartA")
                lngCheckCount = FindID(1, "StartA")
                intRow = intRow + 1
                intType = 1
            End If
            If intType <> 1 Then
                '转为128A码
                GetBarCode = GetBarCode & FindCode(3, "CODEA")
                lngCheckCount = lngCheckCount + intRow * FindID(3, "CODEA")
                intRow = intRow + 1
            End If
            If Len(strTmp) = 1 Then
                GetBarCode = GetBarCode & FindCode(1, strTmp)
                lngCheckCount = lngCheckCount + intRow * FindID(1, strTmp)
                intRow = intRow + 1
            Else
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 1, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 1, 1))
                intRow = intRow + 1
              
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 2, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 2, 1))
                intRow = intRow + 1
            End If
            intType = 1
        End If
    Next
    lngCheckCount = lngCheckCount Mod 103
    GetBarCode = GetBarCode & FindCode(4, CStr(lngCheckCount)) & FindCode(3, "Stop")
End Function
