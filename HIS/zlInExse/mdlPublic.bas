Attribute VB_Name = "mdlPublic"
Option Explicit '要求变量声明
'系统公用临时变量
Public gfrmMain As Object
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrSQL As String

Public glngSys As Long
Public glngModul As Long
Public gblnShowInTaskBar As Boolean

Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public gstrProductName As String

Public gstrUnitName As String '用户单位名称
Public gstrDBUser As String '当前数据库用户名

Public gstrIme As String '自动的开启输入法
Public gblnOK As Boolean
Public Const LONG_MAX = 2147483647 'Long型最大值
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'公共报表参数------------------------------------
Public Const conLineWide As Integer = 30 '横线所占宽度(单位为缇)占两条线宽度
Public Const conLineHigh As Integer = 30 '竖线所占高度(单位为缇)占两条线高度
Public gobjOutTo As Object

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p住院记帐 = 1133
    p病人结帐 = 1137
    p费用查询 = 1139
    p一日清单 = 1141
    p记帐操作 = 1150
    p住院医嘱下达 = 1253
    p预交款 = 1103
End Enum

'API信息-----------------------------------------
Public lngTXTProc As Long '保存默认的消息函数的地址
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'系统方案设置----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const WM_VSCROLL = &H115
Public Const SB_TOP = 6

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Windows风格----------------------------------
Public Const ETO_OPAQUE = 2
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
 
'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'作图API
Public Const PS_SOLID = 0
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Enum Em_Appearance
    Show_3D = 1     '3D显示
    Show_Flat = 0   '平面
End Enum
Public Enum Em_BorderStyle
    Show_Fixed_Single = 1
    Show_None = 0   '无边框线
End Enum

Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum EM_DrawStyle
    DW_Flat = 0  '= 平面
    Dw_SubKen = -1 '= 凹下
    Dw_Heave = 1  '= 凸起
    Dw_Deepen_Subken = -2 '= 深凹下,
    Dw_Deepen_Heave = 2 ' = 深凸起
End Enum
Private Type TY_System_para_Balance
    bln刷卡输入密码 As Boolean  '是否刷卡输入密码
    bln在院不准结帐 As Boolean '1-在院不准结帐,0-在院允许结帐
    bytAuditing As Byte  '记帐未审核单据的结帐处理:0-不检查,1-检查并提示,2-检查并禁止
    byt检查未执行 As Byte    '出院和结帐出院时检查是否有未执行项目及未发药品:0-不检查,1-检查并提示,2-检查并禁止
    byt检查未发药 As Byte   '在出院结帐及病人入出管理中出院时是否检查病人的未发药品项目:0-不检查,1-检查并提示,2-检查并禁止
    bln医生允许才能出院 As Boolean '医生下达出院医嘱才允许病人出院

End Type

Public Enum Em_InputMode
    InPut_Chars = 0 '字符
    InPut_Numbers = 1   '负数
    Input_Moneys = 2    '金额(不包含负数)
    Input_NegativeMoneys = 4 '含负数金额
End Enum


'控件定位
Public Type ty_ctlObject_Locale
    '控件的位置
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '下拉列表的最小高度和宽度
    minWidth As Single
    minHeight As Single
    
    '下接列表的实际位置
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '屏模相关
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type

Public Function OpenIme(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "不自动开启" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function SetWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
'功能：设置 Combo 下拉的宽度,单位为 pixels
    '以twips为单位传入
    Call cbo.SetListWidth(cboHwnd, NewWidthPixel * Screen.TwipsPerPixelX)
    
    SetWidth = True
End Function

Public Function GetWidth(cboHwnd As Long) As Long
'功能： 取得 Combo 下拉的宽度,单位为 pixels
    Dim lRetVal As Long
    lRetVal = cbo.ListWidth(cboHwnd)
    If lRetVal <> -1 Then
        GetWidth = lRetVal / Screen.TwipsPerPixelX
    Else
        GetWidth = 0
    End If
End Function

Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'功能：获取单据费目合计金额
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.实收金额
        Next
    Next
End Function

Public Function GetBillRowTotal(objBillInComes As BillInComes) As Currency
'功能：获取单据费目合计金额
    Dim objBillIncome As New BillInCome
    For Each objBillIncome In objBillInComes
        GetBillRowTotal = GetBillRowTotal + objBillIncome.实收金额
    Next
End Function

Public Function GetFirstRow(curBill As ExpenseBill, Optional strClass As String) As Integer
'功能：获取当前单据中第一个为药品的收费行号
'参数：strClass=取第一中药或西药行,空为药品
'返回：0=没有药品收费行
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstRow = 0
    For i = 1 To curBill.Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", curBill.Details(i).收费类别) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If curBill.Details(i).收费类别 = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function GetFirstClass(curBill As ExpenseBill) As String
'功能：获取当前单据中第一个为药品的收费行号
'返回：0=没有药品收费行
    Dim i As Long
    If curBill.Details.Count = 0 Then GetFirstClass = ""
    For i = 1 To curBill.Details.Count
        If InStr(",5,6,7,", curBill.Details(i).收费类别) > 0 Then
            GetFirstClass = curBill.Details(i).收费类别: Exit Function
        End If
    Next
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function strPad(ByVal strPre As String, ByVal intLen As Integer, ByVal strFill As String, ByVal bytAlign As Byte, Optional ByVal blnTrim As Boolean) As String
'功能：填充字符串
'参数：
'     strPre=要填充的字符串
'     intLen=填充后的长度
'     strFill=要填充的字符
'     bytAlign=1,2/左,右对齐，左对齐时，在原字符串右边填充
'     blnTrim=当字符串超长时，是否强行按指定长度截取。
'返回：处理后的字符串
'说明：一个汉字当作两个字符长度处理
    Dim i As Long
    
    If LenB(StrConv(strPre, vbFromUnicode)) >= intLen Then
        If blnTrim Then
            For i = 1 To Len(strPre)
                strPad = strPad & Mid(strPre, i, 1)
                If LenB(StrConv(strPad, vbFromUnicode)) >= intLen Then Exit For
            Next
        Else
            strPad = strPre
        End If
    Else
        If Len(strFill) > 1 Then strFill = Left(strFill, 1)
        If bytAlign = 1 Then
            strPad = strPre
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
        ElseIf bytAlign = 2 Then
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
            strPad = strPad & strPre
        End If
    End If
End Function

Public Sub PrintCell(ByVal Text As String, _
    ByVal X As Single, ByVal Y As Single, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic)
    '------------------------------------------------
    '功能： 按指定坐标打印一个数据单元,并将当前坐标移动到单元右上角位置
    '参数：
    '   Text:    输出的字符串,其中不包含回车或换行符
    '   X:       左上角X坐标
    '   Y:       左上角Y坐标
    '   Wide:    输出宽度
    '   High:    输出高度
    '   Alignment:    对齐模式，0-左对齐(缺省),1-右对齐,2-居中
    '   ForeColor前景色,缺省为黑色
    '   GridColor边线色,缺省为黑色
    '   FillColor填充色,缺省为设备背景色,由于系统采用了黑色的色码，所以将不允许填充黑色
    '   LineStyle:依序分别为上左右下的线条宽度
    '           0-无线，1-9依序加粗，1为缺省
    '   FontName,FontSize,FontBold,FontItalic:字体属性
    '返回：
    '------------------------------------------------
    Dim aryString() As String       '回车分割的字符串
    Dim lngOldForeColor As Long     '输出设备缺省前景色
    Dim intRow As Integer, intAllRow As Integer
    Dim strRest As String, sngYMove As Single
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    lngOldForeColor = gobjOutTo.ForeColor
    
    On Error Resume Next
    With gobjOutTo
        If Not IsMissing(FontName) Then
            oldFontName = gobjOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = gobjOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = gobjOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = gobjOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = gobjOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = gobjOutTo.TextHeight(Text) + 2 * conLineHigh
'    Wide = CLng(Wide)
'    High = CLng(High)
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '------------------------------------------
    '   边线打印
    '------------------------------------------
    If Mid(LineStyle, 1, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 1, 1)
        gobjOutTo.Line (X, Y)-(X + Wide, Y), GridColor
    End If
    
    If Mid(LineStyle, 2, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 2, 1)
        gobjOutTo.Line (X, Y)-(X, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 3, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 3, 1)
        gobjOutTo.Line (X + Wide, Y)-(X + Wide, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 4, 1) <> 0 Then
        gobjOutTo.DrawWidth = Mid(LineStyle, 4, 1)
        gobjOutTo.Line (X, Y + High)-(X + Wide, Y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   底色填充
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            gobjOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   文字打印
        '------------------------------------------
        gobjOutTo.ForeColor = ForeColor
    
        If InStr(Text, vbCrLf) = 0 And InStr(Text, Chr(13)) = 0 Then
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    '小于一个字符
                intAllRow = 1
            Else
                If gobjOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide)
                Else
                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
                End If
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= gobjOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While gobjOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(Text)) / 2
                Case 1
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(Text)
                Case Else
                    gobjOutTo.CurrentX = X + conLineWide / 2
                End Select
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(Text)
                gobjOutTo.Print Text
                Text = strRest
            Next
        Else
            If InStr(Text, vbCrLf) > 0 Then
                aryString = Split(Trim(Text), vbCrLf)
            Else
                aryString = Split(Trim(Text), Chr(13))
            End If

            intAllRow = UBound(aryString)
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                strRest = aryString(intRow)
                Select Case Alignment
                Case 2
                    Dim blnLR As Boolean
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        blnLR = Not blnLR
                        strRest = IIf(blnLR, Left(strRest, Len(strRest) - 1), Right(strRest, Len(strRest) - 1))
                    Loop
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(strRest)) / 2
                Case 1
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Right(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(strRest)
                Case Else
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Left(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X + conLineWide / 2
                End Select
                
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(strRest)
                If gobjOutTo.CurrentY + gobjOutTo.TextHeight(strRest) > Y + High Then Exit For
                If gobjOutTo.CurrentY >= Y Then gobjOutTo.Print strRest
            
            Next
        End If
    End If
    gobjOutTo.CurrentX = X + Wide
    gobjOutTo.CurrentY = Y
    gobjOutTo.DrawStyle = 0
    gobjOutTo.DrawWidth = 1
    gobjOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then gobjOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then gobjOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then gobjOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then gobjOutTo.FontItalic = oldFontItalic
End Sub

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'功能：判断输入金额是否在原价和现从限定的范围内
'参数：varL=原价,varR=现价,varI=输入金额
'返回：如果不在范围内,则为提示信息,否则为空串
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "输入的价格绝对值不在范围(" & FormatEx(Abs(varL), gintFeePrecision) & "-" & FormatEx(Abs(varR), gintFeePrecision) & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If varI < varL Or varI > varR Then
            CheckScope = "输入的价格值不在范围(" & FormatEx(varL, gintFeePrecision) & "-" & FormatEx(varR, gintFeePrecision) & ")内."
        End If
    End If
End Function

Public Sub SetGridWidth(msh As Control, frmParent As Object)
'功能：自动调整表格列宽,以最小适合为准
    Dim blnRedraw As Boolean
    Dim blnDo As Boolean, i As Long, j As Long, strText As String
    Dim lngStart As Long, lngEnd As Long, lngMaxLen As Long, lngCurLen As Long, lngMWRow As Long
        
    On Local Error Resume Next
    
    blnRedraw = msh.Redraw
    msh.Redraw = False
    lngStart = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1)
    lngEnd = msh.Rows - 1
    
    For i = 0 To msh.Cols - 1
        lngMaxLen = LenB(StrConv(msh.TextMatrix(0, i), vbFromUnicode))  '至少为列名宽度
        lngMWRow = 0
        For j = lngStart To lngEnd
            blnDo = True
            strText = msh.TextMatrix(j, i)
            
            If msh.MergeRow(j) Then
                If i > 0 Then If strText = msh.TextMatrix(j, i - 1) Then blnDo = False
                If blnDo Then
                    If i < msh.Cols - 1 Then If strText = msh.TextMatrix(j, i + 1) Then blnDo = False
                End If
            End If
            If blnDo Then
                lngCurLen = LenB(StrConv(strText, vbFromUnicode))
                If lngCurLen > lngMaxLen Then
                    lngMaxLen = lngCurLen
                    lngMWRow = j
                End If
            End If
        Next
        msh.ColWidth(i) = frmParent.TextWidth(msh.TextMatrix(lngMWRow, i)) + 100
        If msh.ColWidth(i) > 3090 Then msh.ColWidth(i) = 3000
    Next
    
    msh.Redraw = blnRedraw
End Sub

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub SaveRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'功能：对指定金额按分币处理规则进行处理,返回处理后的金额
'参数：curMoney=要进行分币处理的金额(为应缴金额,2位小数)
'      gBytMoney=
'         0.不处理
'         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
'         2.补整收法,eg:0.51=0.60,0.56=0.60
'         3.舍分收法,eg:0.51=0.50,0.56=0.50
'         4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约
'           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
'         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.29(含)以下都舍掉角，0.80(含)以上都进角，0.3-0.79处理为0.5。
'         6.五舍六入:eg:0.15=0.10:0.16=0.2:   刘兴洪 问题:34519  日期:2010-12-06 09:58:02
'91385,调整“5.三七作五、二舍八入”规则：先对分币进行四舍五入，即0.24(含)以下都舍掉角，0.75(含)以上都进角，0.25-0.74都处理为0.5
'       分币先四舍五入，那么0.00～0.24=0，0.25～0.5=0.50, 0.50～0.74=0.50，0.75～1.00=1，这样舍和入各占50%的比例

    Dim intSign As Integer, curTmp As Currency

    If gBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf gBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '先取两位金额,再处理分币,如:0.248 得0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf gBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf gBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf gBytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf gBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf gBytMoney = 6 Then
         '刘兴洪 问题:34519 五舍六入:eg:0.15=0.10:0.16=0.2:    日期:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

Public Sub ExChangeLocate(objA As Object, objB As Object)
'功能：交换医生和开单科室的输入位置
    Dim X1 As Long, Y1 As Long, w1 As Long, t1 As Integer
    Dim X2 As Long, Y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    X1 = objA.Left
    Y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    X2 = objB.Left
    Y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    If TypeName(objB) = "Label" Then
        objB.Left = X1 + w1 - objB.Width
    Else
        objB.Left = X1
        objB.Width = w1
    End If
    objB.Top = Y1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    If TypeName(objA) = "Label" Then
        objA.Left = X2 + w2 - objA.Width
    Else
        objA.Left = X2
        objA.Width = w2
    End If
    objA.Top = Y2
    objA.TabIndex = t2
End Sub

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'功能：检查单据合计金额是否溢出
'说明：以Currency上限922337203685477为准
    Dim dbl应收 As Double, dbl实收 As Double
    Dim i As Integer, j As Integer
    
    '要用VAL转为Double进行运算
    For i = 1 To objBill.Details.Count
        For j = 1 To objBill.Details(i).InComes.Count
            If Abs(dbl应收 + Val(objBill.Details(i).InComes(j).应收金额)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            If Abs(dbl实收 + Val(objBill.Details(i).InComes(j).实收金额)) > 922337203685477# Then
                MoneyOverFlow = True: Exit Function
            End If
            dbl应收 = dbl应收 + Val(objBill.Details(i).InComes(j).应收金额)
            dbl实收 = dbl实收 + Val(objBill.Details(i).InComes(j).实收金额)
        Next
    Next
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    GetTaskbarHeight = OS.TaskbarHeight
End Function


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln负数检查, bln零检查, hWnd, str项目)
End Function

Public Function Where撤档时间(Optional strAlias As String) As String
    If strAlias = "" Then
        Where撤档时间 = " (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null) "
    Else
        Where撤档时间 = " (" & strAlias & ".撤档时间=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".撤档时间 is null) "
    End If
End Function

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub


Public Sub zlRaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '功能：将PictureBox模拟成3D平面按钮
    'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
    Dim PicRect As RECT
    Dim lngTmp As Long
    If picBox Is Nothing Then Exit Sub
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            If intStyle = 2 Then
                    DrawEdge .hDC, PicRect, EDGE_RAISED Or BF_SOFT, BF_RECT
            ElseIf intStyle = -2 Then
                    DrawEdge .hDC, PicRect, EDGE_SUNKEN Or BF_SOFT, BF_RECT
            Else
                DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
            End If
        End If
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) '-10
            End If
            picBox.Print strName
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub
 
'去掉TextBox的默认右键菜单
Public Function NotRightMenuMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then NotRightMenuMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function


