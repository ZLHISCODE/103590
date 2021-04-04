Attribute VB_Name = "mEPREditor"
Option Explicit
'判断当前是否某个虚拟键按下或者放开

'################################################################################################################
'## 功能：  搜索整个文本给出指定关键字区域的定位信息
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey           :   IN  ，给定欲查找的关键字ID号。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果找到该关键字具体位置，则返回True，否则返回False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能：  判断给定位置是否在任何一个关键字对之间，如果是，给出关键字相关位置和ID号
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         lngCurPosition  :   OUT ，指定的当前位置
'##         strKeyType      :   OUT ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey          :   OUT ，给定欲查找的关键字Key。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果包含于某个关键字对之间，则返回True，否则返回False
'################################################################################################################
Public Function IsBetweenAnyKeys(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean

    '基本方法：使用 Instr() 和 InstrRev() 进行查找！
    Dim N As Long, i As Long, j As Long, k As Long
    Dim lFirst As Long
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    strKeyType = ""
    lngKSS = 0
    lngKSE = 0
    lngKES = 0
    lngKEE = 0
    lngKey = 0
    blnNeeded = False

    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        For N = 1 To UBound(gKeyWords)     '共 5 对保留关键字
            '看是否是关键字
            i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
            i = InStr(i, sText, gKeyWords(N).KeyEnd)    '先向后搜索结尾关键字
            If i <> 0 Then
                If .Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                    i = i + 1
                    GoTo LL1
                End If
                j = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL2:
                j = InStr(j, sText, gKeyWords(N).KeyStart) '若找到结尾关键字，再找同名的开始关键字
                If j <> 0 Then
                    If .Range(j - 1, j).Font.Hidden = False Then
                        j = j + 1
                        GoTo LL2
                    End If
                End If
                If (j = 0) Or (j > 0 And i < j) Then '即在关键字对之间
                    k = lngCurPosition
LL3:
                    k = InStrRev(sText, gKeyWords(N).KeyStart, k, vbTextCompare)     '找匹配的开始关键字
                    If k <> 0 Then
                        If .Range(k - 1, k).Font.Hidden = False Then
                            k = k - 1
                            GoTo LL3
                        End If
                        strKeyType = Left(gKeyWords(N).KeyStart, 1)
                        lngKSS = k - 1
                        lngKSE = k + 15
                        lngKES = i - 1
                        lngKEE = i + 15
                        lngKey = Val(.Range(k + 2, k + 10).Text)
                        blnNeeded = -Val(.Range(k + 11, k + 12).Text)
                        IsBetweenAnyKeys = True
                        Exit For
                    End If
                End If
            End If
        Next N
    End With
End Function

'################################################################################################################
'## 功能：  判断给定位置是否在指定的关键字对之间，如果是，给出关键字相关位置和ID号
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         lngCurPosition  :   IN  ，指定的当前位置（由1开始编号）
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         lngKey          :   OUT ，给定欲查找的关键字Key。
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果包含于指定关键字对之间，则返回True，否则返回False
'################################################################################################################
Public Function IsBetweenKeys(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean

    Dim N As Long, i As Long, j As Long, k As Long
    Dim lFirst As Long
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    lngKSS = 0
    lngKSE = 0
    lngKES = 0
    lngKEE = 0
    lngKey = 0
    blnNeeded = False

    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        
        '看是否是关键字
        If lngCurPosition = 0 Then lngCurPosition = edtThis.SelStart + 1
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, strKeyType & "E")    '先向后搜索结尾关键字
        If i <> 0 Then
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            j = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL2:
            j = InStr(j, sText, strKeyType & "S") '若找到结尾关键字，再找同名的开始关键字
            If j <> 0 Then
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
            End If
            If (j = 0) Or (j > 0 And i < j) Then '即在关键字对之间
                k = lngCurPosition
LL3:
                k = InStrRev(sText, strKeyType & "S", k, vbTextCompare)    '找匹配的开始关键字
                If k <> 0 Then
                    If .TOM.TextDocument.Range(k - 1, k).Font.Hidden = False Then
                        k = k - 1
                        GoTo LL3
                    End If
                    lngKSS = k - 1
                    lngKSE = k + 15
                    lngKES = i - 1
                    lngKEE = i + 15
                    lngKey = Val(.TOM.TextDocument.Range(k + 2, k + 10))
                    blnNeeded = -Val(.TOM.TextDocument.Range(k + 11, k + 12))
                    IsBetweenKeys = True
                End If
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能：  搜索指定位置后的第一个给定类型的关键字位置
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         lngCurPosition  :   IN  ，指定的当前位置（由1开始编号）
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey          :   OUT ，给定欲查找的关键字Key。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果包含于指定关键字对之间，则返回True，否则返回False
'################################################################################################################
Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能：  搜索指定位置前的第一个给定类型的关键字位置
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         lngCurPosition  :   IN  ，指定的当前位置（由1开始编号）
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey          :   OUT ，给定欲查找的关键字Key。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果包含于指定关键字对之间，则返回True，否则返回False
'################################################################################################################
Public Function FindPrevKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStrRev(sText, sTMP, i)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindPrevKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能：  获取给出位置的下一个任意关键字位置，如果发现，给出关键字相关位置和ID号
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         lngCurPosition  :   IN  ，指定的当前位置（由1开始编号）
'##         strKeyType      :   OUT ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey          :   OUT ，给定欲查找的关键字Key。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果包含于指定关键字对之间，则返回True，否则返回False
'################################################################################################################
Public Function FindNextAnyKey(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 '转换为0开始的坐标位置。
                lngKSE = i + 14
                lngKES = j - 2
                lngKEE = j + 14
                lngKey = Val(.TOM.TextDocument.Range(i + 1, i + 9))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 10, i + 11))
                FindNextAnyKey = True
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能：  设置指定文本段的常用样式（包括字体格式和段落格式）
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         varStyle        :   IN  ，指定的样式（如果为数字，则表示常用样式编号；如果是字符串，则表示常用样式名称）
'##         lStart          :   IN  ，需要设置样式的文本起始位置
'##         lEnd            :   IN  ，需要设置样式的文本结束位置
'##         ForceParaFmt    :   IN  ，是否强制设置指定范围的文本的段落属性
'##
'## 说明：  通常指定位置位于段首时，同时设置字体属性和段落属性；如果是段内文本，则只设置字体属性；
'##         如果需要同时设置段内文本字体属性和段落属性，则入参 ForceParaFmt 应该置为 True。
'##
'##         若 varStyle =0 或者“清除格式”，则同时清除段落格式（ForceParaFmt = True时）和字体格式
'##
'##   另：  如果某个属性为 tomUndefined = -9999999 ，表示不改变已有属性值。比如，段落行间距为 -9999999，表示不改变行间距。
'################################################################################################################
Public Sub SetCommonStyle(ByRef edtThis As Object, _
        ByVal varStyle As Variant, _
        ByVal lStart As Long, _
        ByVal lEnd As Long, _
        Optional ForceParaFmt As Boolean = False)
        
    Dim rs As New ADODB.Recordset, blnForceEdit As Boolean
    Dim strFont As String, strPara As String
    Dim T As Variant
    Dim blnBeginWithCRLF As Boolean
    
    blnForceEdit = edtThis.ForceEdit
    
    If IsNumeric(varStyle) Then
        If varStyle = 0 Then
            If ForceParaFmt Then edtThis.TOM.TextDocument.Selection.Para.Reset tomDefault
            edtThis.TOM.TextDocument.Selection.Font.Reset tomDefault
            Exit Sub
        End If
        gstrSQL = "select * from 病历常用样式 where 编号=[1]"
    Else
        If varStyle = "清除格式" Then
            If ForceParaFmt Then edtThis.TOM.TextDocument.Selection.Para.Reset tomDefault
            edtThis.TOM.TextDocument.Selection.Font.Reset tomDefault
            Exit Sub
        End If
        gstrSQL = "select * from 病历常用样式 where 名称=[1]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, varStyle)
    
    If Not rs.EOF Then
        strFont = rs("字体样式")
        strPara = rs("段落样式")
        
        If edtThis.Range(lStart - 2, lStart) = vbCrLf Or edtThis.SelStart = 0 Then
            blnBeginWithCRLF = True
        Else
            blnBeginWithCRLF = False
        End If
        
        T = Split(strFont, ";")
        If UBound(T) > 0 Then
            With edtThis.Range(lStart, lEnd).Font
                edtThis.ForceEdit = True
                If Trim(T(0)) <> "" Then .Name = T(0)
                If Val(T(1)) > 0 Then .Size = Val(T(1))
                
                .Bold = IIf(Mid(T(2), 1, 1) = 1, True, False)
                .Italic = IIf(Mid(T(2), 2, 1) = 1, True, False)
'               .Hidden = IIf(Mid(T(2), 3, 1) = 1, True, False)
'                .Protected = IIf(Mid(T(2), 4, 1) = 1, True, False)
'                .Link = IIf(Mid(T(2), 5, 1) = 1, True, False)
'                .Strikethrough = IIf(Mid(T(2), 6, 1) = 1, True, False)
                .Superscript = IIf(Mid(T(2), 7, 1) = 1, True, False)
                .Subscript = IIf(Mid(T(2), 8, 1) = 1, True, False)
                
'                .Underline = Val(T(3))
'                .BackColor = Val(T(4))
'                .ForeColor = Val(T(5))
                edtThis.ForceEdit = blnForceEdit
            End With
        End If

        T = Split(strPara, ";")
        If UBound(T) > 0 Then
            If blnBeginWithCRLF Or ForceParaFmt Then
            '设置段落样式
                With edtThis.Range(lStart, lEnd).Para
                    edtThis.ForceEdit = True
                    '如果为9，则不改变值
                    If Mid(T(0), 2, 1) < 9 Then .ListAlignment = Mid(T(0), 2, 1)                       '正常取值为：0、1、2
'                   If Mid(T(0), 3, 1) < 9 Then .LineSpacingRule = IIf(Mid(T(0), 3, 1) = 1, True, False) '正常取值为：0、1、2、3、4、5
                    
                    If Val(T(1)) <> -9999999 Then .Style = Val(T(1))        '正常取值： -1 ~ -10
                    If Val(T(2)) <> -9999999 Then
                        .ListType = Val(T(2))     '正常取值：0 ～ 6、65536、131072、196608
                        .ListStart = Val(T(3))
                    End If
                    If Val(T(4)) <> tomUndefined Then .FirstLineIndent = Val(T(4)) '首行缩进一般是正数
                    If Val(T(5)) <> tomUndefined Then .LeftIndent = Val(T(5))
                    If Val(T(6)) <> tomUndefined Then .RightIndent = Val(T(6))
'                   If Val(T(7)) <> tomUndefined Then .LineSpacing = Val(T(7))
                    If Val(T(8)) <> tomUndefined Then .ListTab = Val(T(8))
                    If Val(T(9)) <> tomUndefined Then .SpaceBefore = Val(T(9))
                    If Val(T(10)) <> tomUndefined Then .SpaceAfter = Val(T(10))
                    
                    If Mid(T(0), 3, 1) < 9 And Val(T(7)) <> tomUndefined Then .SetLineSpacing Mid(T(0), 3, 1), Val(T(7))
                    If Mid(T(0), 1, 1) < 9 Then .Alignment = Mid(T(0), 1, 1)                           '正常取值为：0、1、2
                    edtThis.ForceEdit = blnForceEdit
                End With
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

