VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFText"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'保持属性值的局部变量
Private mvarText As String
Private mvarBold As Boolean
Private mvarItalic As Boolean
Private mvarForeColor As ColorEnum
Private mvarFontSizeFixup As FontSizeEnum
Private mvarFontStyle As FontStyleEnum
Private mvarCharStyle As CharStyleEnum
Private mvarAnimType As AnimTypeEnum
Private mvarScaleX As Long
Private mvarEmbo As EmboEnum
Private mvarExpand As Long
Private mvarOutLine As Boolean
Private mvarShadow As Boolean
Private mvarStrike As Boolean
Private mvarStrikeD As Boolean
Private mvarSubChar As Boolean
Private mvarSuperChar As Boolean
Private mvarUnderLineType As UnderLineTypeEnum
Private mvarUnderLineColor As ColorEnum
Private mvarHided As Boolean
Private mvarWebHidden As Boolean
Private mvarCharBorder As Boolean
Private mvarCharShadow As Long
Private mvarCharPatForeColor As ColorEnum
Private mvarCharPatBackColor As ColorEnum
Private mvarPatType As PatTypeEnum
Private mvarHighlight As Boolean
Private mvarHighlightColor As ColorEnum
Private mvarAlignment As AlignTypeEnum
Private mvarListType As ListTypeEnum
Private mvarIsInTable As Boolean
Private mvarFontName As String              '目前只支持“宋体”和“黑体”
Private mvarSpaceAfter As Long
Private mvarSpaceBefore As Long
Private mvarSpaceLine As Long
Private mvarSpaceFirstLine As Long
Private mvarFontSizeDefault As Long
Private mvarFontSize As Long

Public Property Let FontSize(ByVal vData As Long)
Attribute FontSize.VB_Description = "字体大小。整形值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property




Public Property Let FontSizeDefault(ByVal vData As Long)
Attribute FontSizeDefault.VB_Description = "默认字体尺寸。一般是18。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontSizeDefault = 5
    mvarFontSizeDefault = vData
End Property


Public Property Get FontSizeDefault() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontSizeDefault
    FontSizeDefault = mvarFontSizeDefault
End Property




Public Property Let SpaceFirstLine(ByVal vData As Long)
Attribute SpaceFirstLine.VB_Description = "首行缩进的字符数。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SpaceFirstLine = 5
    mvarSpaceFirstLine = vData
End Property


Public Property Get SpaceFirstLine() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SpaceFirstLine
    SpaceFirstLine = mvarSpaceFirstLine
End Property




Public Property Let SpaceLine(ByVal vData As Long)
Attribute SpaceLine.VB_Description = "行间距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SpaceLine = 5
    mvarSpaceLine = vData
End Property


Public Property Get SpaceLine() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SpaceLine
    SpaceLine = mvarSpaceLine
End Property



Public Property Let SpaceBefore(ByVal vData As Long)
Attribute SpaceBefore.VB_Description = "段后间距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SpaceBefore = 5
    mvarSpaceBefore = vData
End Property


Public Property Get SpaceBefore() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property



Public Property Let SpaceAfter(ByVal vData As Long)
Attribute SpaceAfter.VB_Description = "段前间距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SpaceAfter = 5
    mvarSpaceAfter = vData
End Property


Public Property Get SpaceAfter() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property





Public Property Let FontName(ByVal vData As String)
Attribute FontName.VB_Description = "字体名称。目前只支持宋体和黑体。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontName = 5
    mvarFontName = vData
End Property


Public Property Get FontName() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontName
    FontName = mvarFontName
End Property




Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "是否包含与表格中。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.IsInTable = 5
    mvarIsInTable = vData
End Property


Public Property Get IsInTable() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property




Public Property Let ListType(ByVal vData As ListTypeEnum)
Attribute ListType.VB_Description = "项目符号类型。1～15。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ListType = 5
    mvarListType = vData
End Property


Public Property Get ListType() As ListTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ListType
    ListType = mvarListType
End Property


Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "段落文本对齐方式。1～4，默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As AlignTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Function GetTextRTF() As String
'获取最终RTF文本
    Dim strR As String
    
    Dim strAlignment As String
    Select Case mvarAlignment
    Case alm2_居中对齐
        strAlignment = "\qc"
    Case alm3_居右对齐
        strAlignment = "\qr"
    Case alm4_两端对齐
        strAlignment = "\qj"
    Case alm5_分散对齐
        strAlignment = "\qd"
    Case Else
        strAlignment = "\ql"
    End Select
    
    Dim strIsInTable As String
    If mvarIsInTable Then strIsInTable = "\intbl"
    
    '默认字体大小及间距处理
    Dim strSB As String, strSA As String, strSL As String, strFL As String
    Dim strFontSize As String
    
    If mvarFontSize <> 0 Then   '字体大小优先权：FontSize > FontSizeFixup > FontSizeDefault
        strFontSize = "\fs" & mvarFontSize * 2
    Else
        If mvarFontSizeFixup <> fs00_默认 Then
            strFontSize = "\fs" & mvarFontSizeFixup * 2
        Else
            If mvarFontSizeDefault <> 0 Then
                strFontSize = "\fs" & mvarFontSizeDefault * 2
            Else
                strFontSize = "\fs18"
            End If
        End If
    End If
    If mvarSpaceBefore <> 0 Then
        strSB = "\sb" & mvarSpaceBefore
    Else    '默认段前间距为0
        If mvarIsInTable Then
            strSB = "\sb0"
        Else
            strSB = "\sb0"
        End If
    End If
    If mvarSpaceAfter <> 0 Then
        strSA = "\sa" & mvarSpaceAfter
    Else    '默认段后间距为0
        If mvarIsInTable Then
            strSA = "\sa0"
        Else
            strSA = "\sa0"
        End If
    End If
    If mvarSpaceLine <> 0 Then
        strSL = "\sl" & mvarSpaceLine & "\slmult0"
    Else    '默认行间距为30
        If mvarIsInTable Then
            strSL = "\sl0\slmult0"
        Else
            strSL = "\sl30\slmult0"
        End If
    End If
    If mvarSpaceFirstLine <> 0 Then
        strFL = "\fi" & mvarSpaceFirstLine * 210
    Else    '默认首行缩进为0
        If mvarIsInTable Then
            strFL = "\fi0"
        Else
            strFL = "\fi0"
        End If
    End If
    
    If ListType <> ltp00_无 Then
        Dim strListString As String
        Select Case ListType
        Case ltp09_符号1
            strListString = "\'6c"
        Case ltp10_符号2
            strListString = "\'6e"
        Case ltp11_符号3
            strListString = "\'75"
        Case ltp12_符号4
            strListString = "\'a8"
        Case ltp13_符号5
            strListString = "\'fc"
        Case ltp14_符号6
            strListString = "\'d8"
        Case ltp15_符号7
            strListString = "\'b2"
        End Select
        
        If ListType <= ltp08_文本8 Then
            strR = "\pard\plain " & strAlignment & strIsInTable & " \fi-420\li3360\ri0\nowidctlpar\jclisttab\tx260\aspalpha\aspnum\faauto\ls1\ilvl" & CStr(mvarListType - 1) & "\adjustright\rin0\lin240\itap0 "
        Else
            strR = "\pard\plain " & strAlignment & strIsInTable & " \li0\ri0\nowidctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 " & _
                    "{\listtext\pard\plain\f10\fs18\kerning2 \loch\af10\dbch\af13\hich\f10 " & strListString & " \tab}" & _
                    "\pard\plain " & strAlignment & strIsInTable & " \fi-420\li3360\ri0\nowidctlpar\jclisttab\tx260\aspalpha\aspnum\faauto\ls2\ilvl" & CStr(mvarListType - 9) & "\adjustright\rin0\lin240\itap0"
        End If
    Else
        If mvarFontStyle = ft0_正文 Then
            '处理间距
            
            strR = strR & "\pard\plain " & strAlignment & strIsInTable & strFL & "\li0\ri0" & strSB & strSA & strSL & "\nowidctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft1_标题1 Then
            strR = strR & "\pard\plain \s1" & strAlignment & strIsInTable & " \li0\ri0\sb340\sa330\sl578\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel0\adjustright\rin0\lin0\itap0 \b\fs44\lang1033\langfe2052\kerning44\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft2_标题2 Then
            strR = strR & "\pard\plain \s2" & strAlignment & strIsInTable & " \li0\ri0\sb260\sa260\sl416\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel1\adjustright\rin0\lin0\itap0 \b\fs32\lang1033\langfe2052\kerning2\loch\af1\hich\af1\dbch\af17\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft3_标题3 Then
            strR = strR & "\pard\plain \s3" & strAlignment & strIsInTable & " \li0\ri0\sb260\sa260\sl416\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel2\adjustright\rin0\lin0\itap0 \b\fs32\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        End If
    End If
    
    If mvarFontStyle = ft0_正文 Then
        '字体：\fN，黑体或者宋体
        If mvarFontName = "黑体" Then strR = strR & "\f17"
        
        '字体尺寸
        strR = strR & strFontSize
        
        If mvarBold Then
            strR = strR & "\b"
        Else
            strR = strR & "\b0"
        End If
        If mvarItalic Then
            strR = strR & "\i"
        Else
            strR = strR & "\i0"
        End If
        If mvarForeColor <> clr00_默认 Then
            strR = strR & "\cf" & mvarForeColor
        End If
        If mvarCharStyle = 0 Then
            strR = strR & "\cs10"
        Else
            strR = strR & "\cs" & mvarCharStyle
        End If
        If mvarAnimType <> ani0_无 Then
            strR = strR & "\animtext" & mvarAnimType
        End If
        If mvarScaleX <> 100 Then
            strR = strR & "\charscalex" & mvarScaleX
        End If
        If mvarEmbo = emb1_浮雕效果 Then
            strR = strR & "\embo"
        ElseIf mvarEmbo = emb2_雕刻效果 Then
            strR = strR & "\impr"
        End If
        If mvarExpand <> 0 Then
            strR = strR & "\expnd" & mvarExpand & "\expndtw" & mvarExpand
        End If
        
        If mvarOutLine Then
            strR = strR & "\outl"
        Else
            strR = strR & "\outl0"
        End If
        If mvarShadow Then
            strR = strR & "\shad"
        Else
            strR = strR & "\shad0"
        End If
        If mvarStrike Then
            strR = strR & "\strike"
        Else
            strR = strR & "\strike0"
        End If
        If mvarStrikeD Then
            strR = strR & "\striked1"
        Else
            strR = strR & "\striked0"
        End If
        If mvarSubChar Then strR = strR & "\sub"
        If mvarSuperChar Then strR = strR & "\super"
        
        If mvarUnderLineType <> udl00_无 Then
            Select Case mvarUnderLineType
                Case udl01_连续线:
                    strR = strR & "\ul"
                Case udl02_点线:
                    strR = strR & "\uld"
                Case udl03_短划线:
                    strR = strR & "\uldash"
                Case udl04_点划线:
                    strR = strR & "\uldashd"
                Case udl05_双点划线:
                    strR = strR & "\uldashdd"
                Case udl06_双线:
                    strR = strR & "\uldb"
                Case udl07_加重波浪线:
                    strR = strR & "\ulhwave"
                Case udl08_长划线:
                    strR = strR & "\ulldash"
                Case udl09_粗线:
                    strR = strR & "\ulth"
                Case udl10_粗点线:
                    strR = strR & "\ulthd"
                Case udl11_粗短划线:
                    strR = strR & "\ulthdash"
                Case udl12_粗点划线:
                    strR = strR & "\ulthdashd"
                Case udl13_粗双点划线:
                    strR = strR & "\ulthdashdd"
                Case udl14_粗长划线:
                    strR = strR & "\ulthldash"
                Case udl15_双波浪线:
                    strR = strR & "\ululdbwave"
                Case udl16_字下加线:
                    strR = strR & "\ulw"
                Case udl17_波浪线:
                    strR = strR & "\ulwave"
                Case Else
                    strR = strR & "\ulnone"
            End Select
            If mvarUnderLineColor <> clr00_默认 Then
                strR = strR & "\ulc" & mvarUnderLineColor
            Else
                strR = strR & "\ulc" & mvarUnderLineColor
            End If
        End If
        If mvarHided Then strR = strR & "\v"
        If mvarWebHidden Then strR = strR & "\webhidden"
        If mvarCharBorder Then strR = strR & "\chbrdr"
        If mvarPatType <> ptt00_无 Then
            Select Case mvarPatType
            Case ptt01_横线:
                strR = strR & "\chbghoriz"
            Case ptt02_竖线:
                strR = strR & "\chbgvert"
            Case ptt03_下斜线:
                strR = strR & "\chbgfdiag"
            Case ptt04_上斜线:
                strR = strR & "\chbgbdiag"
            Case ptt05_网格线:
                strR = strR & "\chbgcross"
            Case ptt06_斜网格线:
                strR = strR & "\chbgdcross"
            Case ptt07_粗横线:
                strR = strR & "\chbgdkhoriz"
            Case ptt08_粗竖线:
                strR = strR & "\chbgdkvert"
            Case ptt09_粗下斜线:
                strR = strR & "\chbgdkfdiag"
            Case ptt10_粗上斜线:
                strR = strR & "\chbgdkbdiag"
            Case ptt11_粗网格线:
                strR = strR & "\chbgdkcross"
            Case ptt12_粗斜网格线:
                strR = strR & "\chbgdkdcross"
            End Select
            If mvarCharShadow Then strR = strR & "\chshdng" & mvarCharShadow '字符阴影百分比
            If mvarCharPatForeColor <> clr00_默认 Then
                strR = strR & "\chcfpat" & mvarCharPatForeColor
            End If
            If mvarCharPatBackColor <> clr00_默认 Then
                strR = strR & "\chcbpat" & mvarCharPatBackColor
            End If
        End If
        If mvarHighlight Then
            strR = strR & "\highlight" & mvarHighlightColor
        End If
'    Else
'        If mvarFontSizeFixup <> 0 Then strR = strR & "\fs" & mvarFontSizeFixup
    End If
    strR = strR & vbCrLf & StrToASC(mvarText)
    GetTextRTF = strR
End Function


Public Property Let HighlightColor(ByVal vData As ColorEnum)
Attribute HighlightColor.VB_Description = "高亮显示背景颜色。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.HighlightColor = 5
    mvarHighlightColor = vData
End Property


Public Property Get HighlightColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.HighlightColor
    HighlightColor = mvarHighlightColor
End Property



Public Property Let Highlight(ByVal vData As Boolean)
Attribute Highlight.VB_Description = "高亮显示。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Highlight = 5
    mvarHighlight = vData
End Property


Public Property Get Highlight() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Highlight
    Highlight = mvarHighlight
End Property



Public Property Let PatType(ByVal vData As PatTypeEnum)
Attribute PatType.VB_Description = "背景图案类型，1～12，默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatType = 5
    mvarPatType = vData
End Property


Public Property Get PatType() As PatTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatType
    PatType = mvarPatType
End Property



Public Property Let CharPatBackColor(ByVal vData As ColorEnum)
Attribute CharPatBackColor.VB_Description = "背景图案前景色。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CharPatBackColor = 5
    mvarCharPatBackColor = vData
End Property


Public Property Get CharPatBackColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CharPatBackColor
    CharPatBackColor = mvarCharPatBackColor
End Property



Public Property Let CharPatForeColor(ByVal vData As ColorEnum)
Attribute CharPatForeColor.VB_Description = "背景图案填充色。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CharPatForeColor = 5
    mvarCharPatForeColor = vData
End Property


Public Property Get CharPatForeColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CharPatForeColor
    CharPatForeColor = mvarCharPatForeColor
End Property



Public Property Let CharShadow(ByVal vData As Boolean)
Attribute CharShadow.VB_Description = "字符阴影。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CharShadow = 5
    mvarCharShadow = vData
End Property


Public Property Get CharShadow() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CharShadow
    CharShadow = mvarCharShadow
End Property



Public Property Let CharBorder(ByVal vData As Boolean)
Attribute CharBorder.VB_Description = "字符边框。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CharBorder = 5
    mvarCharBorder = vData
End Property


Public Property Get CharBorder() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CharBorder
    CharBorder = mvarCharBorder
End Property



Public Property Let WebHidden(ByVal vData As Boolean)
Attribute WebHidden.VB_Description = "Web方式下完全隐藏，包括在保存时。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WebHidden = 5
    mvarWebHidden = vData
End Property


Public Property Get WebHidden() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WebHidden
    WebHidden = mvarWebHidden
End Property



Public Property Let Hided(ByVal vData As Boolean)
Attribute Hided.VB_Description = "隐藏文本，默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Hided = 5
    mvarHided = vData
End Property


Public Property Get Hided() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Hided
    Hided = mvarHided
End Property



Public Property Let UnderLineColor(ByVal vData As ColorEnum)
Attribute UnderLineColor.VB_Description = "下划线颜色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.UnderLineColor = 5
    mvarUnderLineColor = vData
End Property


Public Property Get UnderLineColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.UnderLineColor
    UnderLineColor = mvarUnderLineColor
End Property



Public Property Let UnderLineType(ByVal vData As UnderLineTypeEnum)
Attribute UnderLineType.VB_Description = "下划线类型，1～17，默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.UnderLineType = 5
    mvarUnderLineType = vData
End Property


Public Property Get UnderLineType() As UnderLineTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.UnderLineType
    UnderLineType = mvarUnderLineType
End Property



Public Property Let SuperChar(ByVal vData As Boolean)
Attribute SuperChar.VB_Description = "上标。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SuperChar = 5
    mvarSuperChar = vData
End Property


Public Property Get SuperChar() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SuperChar
    SuperChar = mvarSuperChar
End Property



Public Property Let SubChar(ByVal vData As Boolean)
Attribute SubChar.VB_Description = "下标。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SubChar = 5
    mvarSubChar = vData
End Property


Public Property Get SubChar() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SubChar
    SubChar = mvarSubChar
End Property



Public Property Let StrikeD(ByVal vData As Boolean)
Attribute StrikeD.VB_Description = "双删除线，默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.StrikeD = 5
    mvarStrikeD = vData
End Property


Public Property Get StrikeD() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.StrikeD
    StrikeD = mvarStrikeD
End Property



Public Property Let Strike(ByVal vData As Boolean)
Attribute Strike.VB_Description = "删除线，默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Strike = 5
    mvarStrike = vData
End Property


Public Property Get Strike() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Strike
    Strike = mvarStrike
End Property



Public Property Let Shadow(ByVal vData As Boolean)
Attribute Shadow.VB_Description = "阴影，默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Shadow = 5
    mvarShadow = vData
End Property


Public Property Get Shadow() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Shadow
    Shadow = mvarShadow
End Property



Public Property Let OutLine(ByVal vData As Boolean)
Attribute OutLine.VB_Description = "边框。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.OutLine = 5
    mvarOutLine = vData
End Property


Public Property Get OutLine() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.OutLine
    OutLine = mvarOutLine
End Property



Public Property Let Expand(ByVal vData As Long)
Attribute Expand.VB_Description = "字符间距，负值表示压缩。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Expand = 5
    mvarExpand = vData
End Property


Public Property Get Expand() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Expand
    Expand = mvarExpand
End Property



Public Property Let Embo(ByVal vData As EmboEnum)
Attribute Embo.VB_Description = "浮雕雕刻效果。0:无;1:浮雕;2:雕刻。默认:0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Embo = 5
    mvarEmbo = vData
End Property


Public Property Get Embo() As EmboEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Embo
    Embo = mvarEmbo
End Property



Public Property Let ScaleX(ByVal vData As Long)
Attribute ScaleX.VB_Description = "横向缩放比例，默认：100。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ScaleX = 5
    mvarScaleX = vData
End Property


Public Property Get ScaleX() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ScaleX
    ScaleX = mvarScaleX
End Property



Public Property Let AnimType(ByVal vData As AnimTypeEnum)
Attribute AnimType.VB_Description = "动态文本类型，1～6。默认：0。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.AnimType = 5
    mvarAnimType = vData
End Property


Public Property Get AnimType() As AnimTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.AnimType
    AnimType = mvarAnimType
End Property



Public Property Let CharStyle(ByVal vData As CharStyleEnum)
Attribute CharStyle.VB_Description = "字符样式0，1，2...，默认：-1。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CharStyle = 5
    mvarCharStyle = vData
End Property


Public Property Get CharStyle() As CharStyleEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CharStyle
    CharStyle = mvarCharStyle
End Property



Public Property Let FontStyle(ByVal vData As FontStyleEnum)
Attribute FontStyle.VB_Description = "字体样式0，1，2，...，默认：-1"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontStyle = 5
    mvarFontStyle = vData
End Property


Public Property Get FontStyle() As FontStyleEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontStyle
    FontStyle = mvarFontStyle
End Property


Public Property Let FontSizeFixup(ByVal vData As FontSizeEnum)
Attribute FontSizeFixup.VB_Description = "字体尺寸。默认：21"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontSizeFixup = 5
    mvarFontSizeFixup = vData
End Property


Public Property Get FontSizeFixup() As FontSizeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontSizeFixup
    FontSizeFixup = mvarFontSizeFixup
End Property



Public Property Let ForeColor(ByVal vData As ColorEnum)
Attribute ForeColor.VB_Description = "前景色，默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ForeColor = 5
    mvarForeColor = vData
End Property


Public Property Get ForeColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property



Public Property Let Italic(ByVal vData As Boolean)
Attribute Italic.VB_Description = "斜体。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Italic = 5
    mvarItalic = vData
End Property


Public Property Get Italic() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Italic
    Italic = mvarItalic
End Property



Public Property Let Bold(ByVal vData As Boolean)
Attribute Bold.VB_Description = "粗体。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property



Public Property Let Text(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



