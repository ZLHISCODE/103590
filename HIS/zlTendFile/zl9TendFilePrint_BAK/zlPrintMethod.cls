VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlPrintMethod"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Public objOutTo As Object           '打印输出的目标对象,可能是printer或PictureBox的一页
Private WithEvents frmView As frmPreView       '预览窗体
Attribute frmView.VB_VarHelpID = -1
Private frmView2 As frmPreview2     '提供给外部编程的预览窗体
Private frmExcel As frmOutExcel
Private frmSet As frmPrintSet
Private mvarViewPage As Byte        '预览页数
Private mvarSuccess As Boolean      '执行成功标志
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'zlPrintSet             进行打印机的常规设置
'zlPrintAsk             进行打机对象的有关属性设置
'zlPrintOrViewTends     打印或预览指定护理记录单
'zlPrintCell            按指定坐标打印一个数据单元,并将当前坐标移动到单元右上角位置
'PaperName              根据当前打印机的设置，获取纸张名称
'PaperSource            根据当前打印机的设置，获取送纸方式描述
'zlInitPrn              初始化打印设备，用于外部直接编写面向打印设备的程序时使用
'zlNewPage              打印设备新增一页，用于外部直接编写面向打印设备的程序时使用
'zlPreview              对已经完成的预览提供可视效果，用于外部直接编写面向打印设备的程序时使用
'zlPrintEnd             结束或杀死打印文档
'zlPrintReset           恢复缺省打印设置
'zlGetColor             通过调色板获取RGB颜色值
'zlGetPrinterSet        读取本系统注册表的打印缺省设置
'zlPutPrinterSet        向系统注册表中保存打印缺省设置
'SwitchPrintSet         读取用户对当前表的打印设置，作为当前缺省打印设置
'                       或将当前打印缺省设置保存为用户对当前表的打印设置
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Event zlAfterPrint()

Public Sub InitPrint(ByVal cnMain As ADODB.Connection, ByVal strDBUser As String)
'功能：对数据库连接进行初始化
    Set gcnOracle = cnMain
    gstrDBUser = strDBUser
    gstrPrivs = GetPrivFunc(0, 16)
End Sub

Public Property Let Success(ByVal vData As Boolean)
    mvarSuccess = vData
End Property

Public Property Get OutTo() As Object
    Set OutTo = objOutTo
End Property

Public Property Let ViewPage(ByVal vData As Byte)
    If vData > 0 Then
        mvarViewPage = vData
    End If
End Property

Public Property Get ViewPage() As Byte
    ViewPage = mvarViewPage
End Property

Private Sub Class_Initialize()
    Set objOutTo = Printer
    Set frmAsk = New frmTendPrintAsk
    Set frmView = New frmPreView
    Set frmView2 = New frmPreview2
    Set frmSet = New frmPrintSet
    mvarViewPage = 6
    mvarSuccess = True
    
    gstrGrant = GetSetting("ZLSOFT", "注册信息", "Kind", "")
    If gstrGrant = "正式" Then gstrGrant = ""
    Call GetOEM
End Sub

Private Sub Class_Terminate()
    Set frmAsk = Nothing
'    Unload frmView
    Set frmView = Nothing
    Set frmView2 = Nothing
    Set frmSet = Nothing
    Set gobjOutTo = Nothing
    Set gobjSend = Nothing
    Set objOutTo = Nothing
    Set gcnOracle = Nothing
End Sub


Public Function zlPrintSet() As Boolean
    '------------------------------------------------
    '功能： 进行打印机的常规设置
    '参数：
    '返回： 成功设置，返回true；系统未安装打印机，返回false
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        zlPrintSet = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    zlGetPrinterSet
    frmSet.Show 1
    zlPutPrinterSet
    
    zlPrintSet = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintSet = False

End Function

Public Function zlPrintAsk(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, ByVal lng文件ID As Long, Optional ByVal blnBatch As Boolean = False) As Byte
    '------------------------------------------------
    '功能： 进行打机对象的有关属性设置
    '参数：
    '   objSend:打印对象
    '返回： 放弃或不能设置=0；打印=1；预览=2
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        zlPrintAsk = 0
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    glng文件ID = lng文件ID
    glng病人ID = lng病人ID
    glng主页ID = lng主页ID
    gint婴儿 = int婴儿
    '打印结束前关闭询问窗体
    If blnBatch Then
        Load frmAsk
    Else
        frmAsk.Show vbModal
    End If
    If frmAsk.byRunMode = 0 Then
        zlPrintAsk = 0
        Exit Function
    End If
    
    zlPrintAsk = frmAsk.byRunMode
    Exit Function
errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintAsk = 0
    
End Function
'
'Public Function zlPrintOrViewTends(objSend As zlPrintTends, Optional byRunMode As Byte = 1) As Boolean
'    '------------------------------------------------
'    '功能： 打印或预览指定msFlexGrid打印对象
'    '参数：
'    '   objSend:打印对象
'    '   byRunMode:输出模式，1-打印，2-预览
'    '返回： 成功返回true ；错误返回false
'    '------------------------------------------------
'    Dim GridRedraw As Boolean
'
'    If byRunMode = 3 Then
'        If gstrGrant <> "" Then
'            MsgBox "试用或测试版本不能使用该功能。", vbInformation, gstrSysName
'            Exit Function
'        End If
'        If Not HaveExcel() Then
'            MsgBox "请安装好Excel或WPS(ET)后，再使用本功能。", vbCritical + vbOKOnly, gstrSysName
'            zlPrintOrViewTends = False
'            Exit Function
'        End If
'        If objSend.Privileged = False Then
'            If InStr(1, ";" & gstrPrivs & ";", ";Excel输出;") = 0 Then
'                MsgBox "你没有Excel输出的权限，不能使用该功能。", vbInformation + vbOKOnly, gstrSysName
'                zlPrintOrViewTends = False
'                Exit Function
'            End If
'        End If
'
'        Set gobjSend = objSend
'        Set frmExcel = New frmOutExcel
'
'        GridRedraw = objSend.Body.Redraw
'        objSend.Body.Redraw = False
'        frmExcel.Show 1
'        objSend.Body.Redraw = GridRedraw
'
'        Set frmExcel = Nothing
'        Exit Function
'    End If
'
'    If Printers.Count = 0 Then
'        MsgBox "注意：" & Chr(13) _
'            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
'            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
'            zlPrintOrViewTends = False
'        Exit Function
'    End If
'    zlGetPrinterSet
'
'    Err = 0
'    On Error GoTo errHand
'    '----------------------------------------------------
'    '   对象特性保存
'    '----------------------------------------------------
'    GridRedraw = objSend.Body.Redraw
'    objSend.Body.Redraw = False
'
'    '----------------------------------------------------
'    '初始化输出目标对象
'    '----------------------------------------------------
'    gintObjType = gintTends
'    Set gobjSend = objSend
'    If Not ReadVar Then
'        objSend.Body.Redraw = GridRedraw
'        zlPrintOrViewTends = False
'        Exit Function
'    End If
'
'    '最多不过一列就是一页面
'    Call CalculateHeight
'    Call CalculateRC
'
'    If byRunMode = 1 Then
'        If objSend.Privileged = False Then
'            If InStr(1, ";" & gstrPrivs & ";", ";打印;") = 0 Then
'                MsgBox "你没有打印权限，不能使用该功能。", vbInformation + vbOKOnly, gstrSysName
'                zlPrintOrViewTends = False
'                Exit Function
'            End If
'        End If
'        RealPrint 1, gintRowTotal * gintColTotal
'    Else
'        Set frmView = New frmPreView
'        frmView.Show 1
'        Set frmView = Nothing
'    End If
'
'
'    If Not mvarSuccess Then zlPrintOrViewTends = False: Exit Function
'
'    '网格特性恢复
'    objSend.Body.Redraw = GridRedraw
'    zlPrintOrViewTends = True
'    Exit Function
'
'errHand:
'    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
'    objSend.Body.Redraw = GridRedraw
'    zlPrintOrViewTends = False
'End Function

Public Function zlPrintOrViewTends(ByVal blnBatch As Boolean, Optional byRunMode As Byte = 1) As Boolean
    '------------------------------------------------
    '功能： 打印或预览指定msFlexGrid打印对象
    '参数：
    '   byRunMode:输出模式，1-打印，2-预览
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim lngMaxPage As Long, lngStartPage As Long, lngPages As Long
    gblnPrintMode = (byRunMode = 1)
    If byRunMode = 3 Then
        If gstrGrant <> "" Then
            MsgBox "试用或测试版本不能使用该功能。", vbInformation, gstrSysName
            GoTo ex_d
        End If
        If Not HaveExcel() Then
            MsgBox "请安装好Excel后再使用本功能。", vbCritical + vbOKOnly, gstrSysName
            GoTo ex_d
        End If

        Load frmTendFileReader
        lngMaxPage = frmTendFileReader.GetPages
        lngStartPage = frmTendFileReader.GetStartPage
        lngPages = lngMaxPage - lngStartPage + 1

        If gobjSend.Privileged = False Then
            If InStr(1, ";" & gstrPrivs & ";", ";Excel输出;") = 0 Then
                MsgBox "你没有Excel输出的权限，不能使用该功能。", vbInformation + vbOKOnly, gstrSysName
                GoTo ex_d
            End If
        End If

        Set frmExcel = New frmOutExcel
        frmExcel.Show 1

        Set frmExcel = Nothing
        GoTo ex_d
    End If
    
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
            zlPrintOrViewTends = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    Load frmTendFileReader
    lngMaxPage = frmTendFileReader.GetPages
    lngStartPage = frmTendFileReader.GetStartPage
    lngPages = lngMaxPage - lngStartPage + 1
'    If lngPages = 0 Then
'        MsgBox "没有打印数据！", vbInformation, gstrSysName
'        GoTo ex_d
'    End If
    If Not ReadVar Then GoTo ex_d
    
    If byRunMode = 1 Then
        If InStr(1, ";" & gstrPrivs & ";", ";打印;") = 0 Then
            MsgBox "你没有打印权限，不能使用该功能。", vbInformation + vbOKOnly, gstrSysName
            GoTo ex_d
        End If
        RealPrint 1, lngPages
    Else
        Set frmView = New frmPreView
        frmView.Show 1
        Set frmView = Nothing
    End If
    If Not mvarSuccess Then GoTo ex_d
    
    zlPrintOrViewTends = True
    GoTo ex_d
errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPrintOrViewTends = False
ex_d:
    Unload frmTendFileReader
    Unload frmAsk
End Function

Public Sub zlPrintCell(ByVal Text As String, _
    ByVal x As Long, ByVal Y As Long, _
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
    Dim intRow As Long, intAllRow As Long
    Dim strRest As String, lngYMove As Long
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    lngOldForeColor = objOutTo.ForeColor
    
    On Error Resume Next
    With objOutTo
        If Not IsMissing(FontName) Then
            oldFontName = objOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = objOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = objOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = objOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = objOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = objOutTo.TextHeight(Text) + 2 * conLineHigh
    
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '------------------------------------------
    '   边线打印
    '------------------------------------------
    If Mid(LineStyle, 1, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 1, 1)
        objOutTo.Line (x, Y)-(x + Wide, Y), GridColor
    End If
    
    If Mid(LineStyle, 2, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 2, 1)
        objOutTo.Line (x, Y)-(x, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 3, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 3, 1)
        objOutTo.Line (x + Wide, Y)-(x + Wide, Y + High), GridColor
    End If
    
    If Mid(LineStyle, 4, 1) <> 0 Then
        objOutTo.DrawWidth = Mid(LineStyle, 4, 1)
        objOutTo.Line (x, Y + High)-(x + Wide, Y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   底色填充
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            objOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   文字打印
        '------------------------------------------
        objOutTo.ForeColor = ForeColor
    
        If InStr(1, Text, Chr(13)) = 0 Then
            If Wide - conLineWide < objOutTo.TextWidth("1") Then    '小于一个字符
                intAllRow = 1
            Else
                If objOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
                    intAllRow = objOutTo.TextWidth(Text) \ (Wide - conLineWide)
                Else
                    intAllRow = objOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
                End If
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= objOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            lngYMove = (High - conLineHigh - objOutTo.TextHeight(Text) * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While objOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    objOutTo.CurrentX = x + (Wide - objOutTo.TextWidth(Text)) / 2
                Case 1
                    objOutTo.CurrentX = x - conLineWide / 2 + Wide - objOutTo.TextWidth(Text)
                Case Else
                    objOutTo.CurrentX = x + conLineWide / 2
                End Select
                objOutTo.CurrentY = Y + conLineHigh / 2 + lngYMove + intRow * objOutTo.TextHeight(Text)
                objOutTo.Print Text
                Text = strRest
            Next
        Else
            intAllRow = 1
            aryString = Split(Trim(Text), Chr(13))
            intAllRow = UBound(aryString)
            lngYMove = (High - conLineHigh - objOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                Text = aryString(intRow)
                Select Case Alignment
                Case 2
                    objOutTo.CurrentX = x + (Wide - objOutTo.TextWidth(Text)) / 2
                Case 1
                    objOutTo.CurrentX = x - conLineWide / 2 + Wide - objOutTo.TextWidth(Text)
                Case Else
                    objOutTo.CurrentX = x + conLineWide / 2
                End Select
                objOutTo.CurrentY = Y + conLineHigh / 2 + lngYMove + intRow * objOutTo.TextHeight(Text)
                objOutTo.Print Text
                Text = strRest
            Next
        
        End If
    End If
    objOutTo.CurrentX = x + Wide
    objOutTo.CurrentY = Y
    objOutTo.DrawStyle = 0
    objOutTo.DrawWidth = 1
    objOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then objOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then objOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then objOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then objOutTo.FontItalic = oldFontItalic
End Sub


Public Function PaperName() As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取纸张名称
    '参数：
    '返回： 纸张名称
    '------------------------------------------------
    Dim mSize As Integer
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperSize = 256 Then
        PaperName = "用户自定义，" _
            & Printer.Width / 56.6857142857143 & "x" _
            & Printer.Height / 56.6857142857143 & "毫米"
        Exit Function
    End If
    If Printer.PaperSize >= 1 And Printer.PaperSize <= 41 Then
        mSize = Printer.PaperSize
        PaperName = IIf(Printer.Orientation = 1, "纵向", "横向") & Space(2) _
            & Switch( _
            mSize = 1, conSize1, mSize = 2, conSize2, mSize = 3, conSize3, mSize = 4, conSize4, mSize = 5, conSize5, _
            mSize = 6, conSize6, mSize = 7, conSize7, mSize = 8, conSize8, mSize = 9, conSize9, mSize = 10, conSize10, _
            mSize = 11, conSize11, mSize = 12, conSize12, mSize = 13, conSize13, mSize = 14, conSize14, mSize = 15, conSize15, _
            mSize = 16, conSize16, mSize = 17, conSize17, mSize = 18, conSize18, mSize = 19, conSize19, mSize = 20, conSize20, _
            mSize = 21, conSize21, mSize = 22, conSize22, mSize = 23, conSize23, mSize = 24, conSize24, mSize = 25, conSize25, _
            mSize = 26, conSize26, mSize = 27, conSize27, mSize = 28, conSize28, mSize = 29, conSize29, mSize = 30, conSize30, _
            mSize = 31, conSize31, mSize = 32, conSize32, mSize = 33, conSize33, mSize = 34, conSize34, mSize = 35, conSize35, _
            mSize = 36, conSize36, mSize = 37, conSize37, mSize = 38, conSize38, mSize = 39, conSize39, mSize = 40, conSize40, _
            mSize = 41, conSize41)
        Exit Function
    End If
errHand:
    PaperName = "不可测的纸张"
    
End Function

Public Function PaperSource() As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取送纸方式描述
    '参数：
    '返回： 送纸方式字符串
    '------------------------------------------------
    Dim mBin As Integer
    
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperBin = 14 Then
        PaperSource = "附加的卡式纸盒进纸"
        Exit Function
    End If
    If Printer.PaperBin >= 1 And Printer.PaperBin <= 11 Then
        PaperSource = Switch( _
            mBin = 1, conBin1, mBin = 2, conBin2, mBin = 3, conBin3, mBin = 4, conBin4, mBin = 5, conBin5, _
            mBin = 6, conBin6, mBin = 7, conBin7, mBin = 8, conBin8, mBin = 9, conBin9, mBin = 10, conBin10, _
            mBin = 11, conBin11)
        Exit Function
    End If
errHand:
    PaperSource = "不可测的进纸方式"

End Function

Public Function zlInitPrn(Optional byRunMode As Byte = 1, _
    Optional lngHeight As Long = 0, _
    Optional lngWidth As Long = 0) As Boolean
    '------------------------------------------------
    '功能： 初始化打印设备，用于外部直接编写面向打印设备的程序时使用
    '参数：
    '   byRunMode:打印模式。缺省和1，设备指向打印机；2-指向预览页面
    '返回：成功，返回true;失败，返回false,且打印设备为空
    '------------------------------------------------
    Dim iPage As Integer
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlInitPrn = False
        Exit Function
    End If
    
    Err = 0
    On Error Resume Next
    If Printer.Orientation = 1 Then
        If lngHeight <> 0 Then Printer.Height = lngHeight
        If lngWidth <> 0 Then Printer.Width = lngWidth
    Else
        If lngHeight <> 0 Then Printer.Width = lngHeight
        If lngWidth <> 0 Then Printer.Height = lngWidth
    End If
    Err = 0
    On Error GoTo errHand
    If byRunMode = 1 Then
        Set objOutTo = Printer
    Else
        If frmView2.PctPage.Count > 1 Then
            Set frmView2 = Nothing
            Set frmView2 = New frmPreView
        End If
        frmView2.PctPage(0).Height = Printer.Height
        frmView2.PctPage(0).Width = Printer.Width
        Load frmView2.PctPage(1)
        Set objOutTo = frmView2.PctPage(1)
    End If
    objOutTo.FillStyle = 1          '设置填充透明
    
    zlInitPrn = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Set objOutTo = Nothing
    zlInitPrn = False

End Function


Public Function zlNewPage() As Boolean
    '------------------------------------------------
    '功能： 打印设备新增一页，用于外部直接编写面向打印设备的程序时使用
    '参数：
    '返回：成功，返回true;失败，返回false
    '------------------------------------------------
    If Printers.Count = 0 Then
        MsgBox "注意：" & Chr(13) _
            & "    未安装打印机，请通过系统设置的打印机" & Chr(13) _
            & "管理添加安装打印机。", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlNewPage = False
        Exit Function
    End If
    
    If objOutTo Is Nothing Then
        MsgBox "注意：" & Chr(13) _
            & "没有初始化打印设备。", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlNewPage = False
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand
    If objOutTo Is Printer Then
        objOutTo.NewPage
    Else
        Dim intPageNo As Integer
        intPageNo = frmView2.PctPage.Count
        Load frmView2.PctPage(intPageNo)
        frmView2.PctPage(intPageNo).Height = Printer.Height
        frmView2.PctPage(intPageNo).Width = Printer.Width
        Set objOutTo = frmView2.PctPage(intPageNo)
    End If
    zlNewPage = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Set objOutTo = Nothing
    zlNewPage = False

End Function

Public Function zlPreview() As Boolean
    '------------------------------------------------
    '功能：对已经完成的预览提供可视效果，用于外部直接
    '   编写面向打印设备的程序时使用
    '参数：
    '返回：成功，返回true;失败，返回false
    '------------------------------------------------
    If frmView2.PctPage.Count < 2 Then
        MsgBox "注意：" & Chr(13) _
            & "错误调用了打印方法", vbCritical + vbOKOnly, gstrSysName
        Set objOutTo = Nothing
        zlPreview = False
        Exit Function
    End If
    
    frmView2.Toolbar.Buttons("print").Enabled = False
    frmView2.Show 1
    Unload frmView2
    zlPreview = True
    Exit Function
    
errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    zlPreview = False

End Function

Public Sub zlPrintEnd(Optional EndMode As Byte = 0)
    '----------------------------
    '功能：结束或杀死打印文档
    '----------------------------
    If EndMode <> 0 Then
        Printer.KillDoc
    Else
        Printer.EndDoc
    End If
End Sub

Public Sub zlPrintReset()
    '----------------------------
    '功能：恢复缺省打印设置
    '----------------------------
    Class_Terminate
    Class_Initialize
End Sub


Public Function zlGetColor(Optional Color As Long) As Long
    '------------------------------------------------
    '功能：通过调色板获取RGB颜色值
    '参数：
    '   color:缺省颜色
    '返回：确定，返回颜色;取消，返回缺省颜色或黑色
    '------------------------------------------------
    Dim blnDefault As Boolean
    
    Err = 0
    On Error GoTo errHand
    If Not IsMissing(Color) Then
        blnDefault = True
        If LCase(TypeName(Color)) <> "long" And _
            LCase(TypeName(Color)) <> "integer" And _
            LCase(TypeName(Color)) <> "byte" Then
            MsgBox "缺省颜色错误!", vbExclamation, gstrSysName
            blnDefault = False
        End If
    Else
        blnDefault = False
    End If
    
    With frmAsk
        If blnDefault Then .comDlg.Color = Color
        .comDlg.CancelError = True
        .comDlg.Flags = cdlCCRGBInit
        .comDlg.ShowColor
        zlGetColor = .comDlg.Color
    End With
    Exit Function
        
errHand:
    If blnDefault Then
        zlGetColor = Color
    Else
        zlGetColor = 0
    End If

End Function

Public Function zlGetPrinterSet() As Boolean
    
    zlGetPrinterSet = GetPrinterSet
End Function

Public Function zlPutPrinterSet() As Boolean
    '------------------------------------------------
    '功能：向系统注册表中保存打印缺省设置
    '------------------------------------------------
    If Printers.Count = 0 Then
        zlPutPrinterSet = False
        Exit Function
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Orientation", Printer.Orientation
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Width", Printer.Width
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Height", Printer.Height
    zlPutPrinterSet = True
End Function

Public Function SwitchPrintSet(strSection As String, Optional blnSave As Boolean = False) As Boolean
    '------------------------------------------------
    '功能：读取用户对当前表的打印设置，作为当前缺省打印设置
    '   或将当前打印缺省设置保存为用户对当前表的打印设置
    '参数：
    '   strSection:当前表打印设置段名称
    '   blnSave:true,保存;false,读取
    '------------------------------------------------
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strSource As String, strObject As String
    
    If Printers.Count = 0 Then
        SwitchPrintSet = False
        Exit Function
    End If
    
    If blnSave Then
        strSource = "公共模块\" & App.ProductName & "\Default"
        strObject = "公共模块\" & App.ProductName & "\" & strSection
    Else
        strSource = "公共模块\" & App.ProductName & "\" & strSection
        strObject = "公共模块\" & App.ProductName & "\Default"
    End If
    
    strDeviceName = GetSetting("ZLSOFT", strSource, "DeviceName", Printer.DeviceName)
    intPaperSize = GetSetting("ZLSOFT", strSource, "PaperSize", Printer.PaperSize)
    lngWidth = GetSetting("ZLSOFT", strSource, "Width", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", strSource, "Height", Printer.Height)
    intPaperBin = GetSetting("ZLSOFT", strSource, "PaperBin", Printer.PaperBin)
    intOrientation = GetSetting("ZLSOFT", strSource, "Orientation", Printer.Orientation)
    
    SaveSetting "ZLSOFT", strObject, "DeviceName", strDeviceName
    SaveSetting "ZLSOFT", strObject, "PaperSize", intPaperSize
    SaveSetting "ZLSOFT", strObject, "PaperBin", intPaperBin
    SaveSetting "ZLSOFT", strObject, "Orientation", intOrientation
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Width", lngWidth
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Height", lngHeight
    
    SwitchPrintSet = True
End Function

Private Sub GetOEM()
'从注册表中得到产品注册信息
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    If gstrSysName = "" Then
        gstrSysName = "中联软件"
    End If
End Sub

Private Sub frmView_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub


