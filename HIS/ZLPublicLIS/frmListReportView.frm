VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmListReportView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "预览"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   DrawMode        =   2  'Blackness
   DrawStyle       =   2  'Dot
   FillStyle       =   0  'Solid
   HasDC           =   0   'False
   Icon            =   "frmListReportView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   134.673
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.815
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   9720
      MousePointer    =   99  'Custom
      ScaleHeight     =   1920
      ScaleWidth      =   1785
      TabIndex        =   7
      Top             =   2430
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.PictureBox picTmp 
      Height          =   2010
      Left            =   10125
      ScaleHeight     =   1950
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   4125
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1860
   End
   Begin VB.HScrollBar scrHsc 
      Height          =   250
      LargeChange     =   20
      Left            =   -180
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7260
      Width           =   8760
   End
   Begin VB.VScrollBar scrVsc 
      Height          =   4755
      LargeChange     =   20
      Left            =   9390
      Max             =   100
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Width           =   250
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   30
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   8760
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   2010
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   1
         Top             =   390
         Width           =   6990
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   1740
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   6990
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmListReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
''######################################################################################################################
''常量
'
''常量定义
Private Const OFFSET_LEFT = 1
Private Const OFFSET_TOP = 20
Private Const OFFSET_RIGHT = 20
Private Const OFFSET_BOTTOM = 20
Private Const SHADOW_W = 45 '阴影厚度
Private Const LOGPIXELSY = 90

Private Const LF_FACESIZE = 32
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const PROOF_QUALITY = 2
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0                 '     Don't   care   or   don't   know.
Private Const OEM_CHARSET = 255
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_TT_ONLY_PRECIS = 7

Private Type PRINTERPROPERTY
    PaperSize As Integer
    PaperWidth As Long
    PaperHeight As Long
    PaperLeft As Long
    PaperTop As Long
    PaperOrientation As Integer
End Type

Private mblnStartUp As Boolean
Private mstrTitle As String
Private mPrinterProperty As PRINTERPROPERTY
Private grsData As New ADODB.Recordset
Private grsPage As New ADODB.Recordset
Private mintCurPage As Integer
Private mintPage As Integer
Private mlngLeft As Long
Private mlngWidth As Long
Private mlngHeight As Long

Private mfrmListReportSet As New frmListReportSet

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type
Private Type DOCINFO
   cbSize As Long
   lpszDocName As String
   lpszOutput As String
   lpszDatatype As String
   fwType As Long
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long       ' or Boolean
Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long                                                                         '   or   Boolean

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Const BLACKONWHITE = 1
Private Const WHITEONBLACK = 2
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4
Private Const MAXSTRETCHBLTMODE = 4
Private Const STRETCH_ANDSCANS = BLACKONWHITE
Private Const STRETCH_ORSCANS = WHITEONBLACK
Private Const STRETCH_DELETESCANS = COLORONCOLOR
Private Const STRETCH_HALFTONE = HALFTONE
Private Const SRCCOPY = &HCC0020

Private Const DESIREDFONTSIZE = 12     ' Could use variable, TextBox, etc.
Private mstrRegisterPath    As String
Private mstrPrintPage As String
Public Event AfterPrinted()
''
'''######################################################################################################################
'''过程清单
''
Public Function InitReport(ByVal rsData As ADODB.Recordset, ByVal rsPage As ADODB.Recordset, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：初始化处理
    '参数：
    '返回：
    '******************************************************************************************************************
    cboPage.Clear
    mstrRegisterPath = strRegisterPath
    If CheckPrint(strRegisterPath) = False Then Exit Function

    '设置打印机
    '------------------------------------------------------------------------------------------------------------------
    mPrinterProperty.PaperSize = GetSetting("ZLSOFT", strRegisterPath, "纸张", Printer.PaperSize)
    If mPrinterProperty.PaperSize = 256 Then
        mPrinterProperty.PaperWidth = GetSetting("ZLSOFT", strRegisterPath, "宽度", Printer.Width)
        mPrinterProperty.PaperHeight = GetSetting("ZLSOFT", strRegisterPath, "高度", Printer.Height)
    End If

    mPrinterProperty.PaperLeft = GetSetting("ZLSOFT", strRegisterPath, "左边距", OFFSET_LEFT)
    mPrinterProperty.PaperTop = GetSetting("ZLSOFT", strRegisterPath, "上边距", OFFSET_TOP)
    mPrinterProperty.PaperOrientation = GetSetting("ZLSOFT", strRegisterPath, "纸向", Printer.Orientation)

    On Error Resume Next

    '纸张
    Printer.PaperSize = mPrinterProperty.PaperSize

    If mPrinterProperty.PaperSize = 256 Then
        Printer.Width = mPrinterProperty.PaperWidth
        Printer.Height = mPrinterProperty.PaperHeight
    Else
        '重新读取纸张宽和高
        mPrinterProperty.PaperWidth = Printer.Width
        mPrinterProperty.PaperHeight = Printer.Height
    End If

    Printer.Orientation = mPrinterProperty.PaperOrientation

    Set grsData = rsData
    Set grsPage = rsPage

    InitReport = True

End Function
''
Public Function ExportReport(Optional ByVal bytMode As Byte = 1, Optional ByVal strTitle As String, Optional ByVal strFile As String, Optional ByVal strPassWord As String, Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '功能：输出报告
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrTitle = strTitle & "预览"

    Select Case bytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 2                          '打印

        If PrintReport(strPage) Then
            RaiseEvent AfterPrinted
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case 1                          '预览

        Call PreviewReport(strPage)

    '------------------------------------------------------------------------------------------------------------------
'    Case 3                          '输出到PDF
'
'        If ExportPDFReport(strFile, strPassword, strPage) Then
'            RaiseEvent AfterPrinted
'        End If

    End Select
    
    ExportReport = True

End Function
'
Private Function PreviewReport(Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '功能：打印预览
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intTmpPage As Integer
    Dim strShow As String

    mblnStartUp = True

    '清空原来页面的数据
    For intLoop = picPage.UBound To 0 Step -1

        picPage(intLoop).Tag = ""
        If intLoop = 0 Then
            picPage(intLoop).Cls
        Else
            Unload picPage(intLoop)
        End If
    Next

    picPage(0).Width = Printer.Width
    picPage(0).Height = Printer.Height

    If grsPage.RecordCount > 0 Then
        intLoop = 0
        grsPage.Sort = "页号"
        grsPage.MoveFirst
        Do While Not grsPage.EOF

            intTmpPage = Val(grsPage("页号").Value) - 1
            If picPage.UBound < intTmpPage Then Call AddPage(intTmpPage)

            If grsPage("显示内容").Value <> "" Then
                strShow = grsPage("显示内容").Value
            Else
                strShow = "第 " & grsPage("虚拟页号").Value & " 页"
            End If

            cboPage.AddItem strShow

            grsPage.MoveNext
        Loop
    End If

    cboPage.Tag = intLoop

    '设置当前页面
    picPage(0).Visible = True
    picPage(0).ZOrder

    mblnStartUp = False

    If cboPage.ListCount > 0 Then cboPage.ListIndex = 0
    If gbtyModel = 2 Then
        Me.Caption = "检验报告批量打印"
        Me.BorderStyle = 2
        Me.WindowState = 2
        Me.Show 1
    Else
        Me.Caption = ""
        Me.BorderStyle = 0
    End If
    PreviewReport = True

End Function

'Private Function ExportPDFReport(Optional ByVal strFile As String, Optional ByVal strPassword As String, Optional ByVal strPage As String) As Boolean
'    '******************************************************************************************************************
'    '功能：输出为PDF
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intPages As Integer
'    Dim intPage As Integer
'    Dim clsPrintPDF As clsPDF
'    Dim blnFirst As Boolean
'
'    On Error GoTo errHand
'
'    grsPage.Filter = ""
'    If grsPage.RecordCount > 0 Then
'
'        For intLoop = 0 To Printers.count - 1
'            If Printers(intLoop).DeviceName = "TinyPDF" Then
'                Set Printer = Printers(intLoop): Exit For
'            End If
'        Next
'        If intLoop = Printers.count Then Exit Function
'
'        Set clsPrintPDF = New clsPDF
'        If clsPrintPDF.InitPDF(strFile, , , True, strPassword) = False Then Exit Function
'        blnFirst = True
'
'        intPages = grsPage.RecordCount
'        For intLoop = 1 To intPages
'            grsPage.Filter = "页号=" & intLoop
'
'            If strPage <> "" Then
'                '按虚拟页号计算
'                intPage = Val(grsPage("虚拟页号").Value)
'                If InStr("," & strPage & ",", "," & intPage & ",") = 0 Then
'                    grsPage.Filter = "页号=0"
'                End If
'            End If
'
'            If grsPage.RecordCount > 0 Then
'                If intLoop > 1 Then Printer.NewPage
'
'                If blnFirst Then
'                    blnFirst = False
'
'                    On Error Resume Next
'                    err = 0
'                    Printer.Print ""
'                    If err <> 0 Then
'                        Printer.EndDoc
'                        Exit Function
'                    End If
'                    On Error GoTo errHand
'                End If
'
'                Call ShowPage(Printer, intLoop)
'
'            End If
'
'        Next
'
'        Printer.EndDoc
'
'        Call clsPrintPDF.ResetPDF
'
'    End If
'
'    ExportPDFReport = True
'
'    Exit Function
'
'errHand:
'    If gobjComLib.ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
'
Private Function PrintReport(Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '功能：打印
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intPages As Integer
    Dim blnFirst As Boolean
    Dim intPage As Integer
    Dim intPrintPages As Integer
    
    strPage = "" ' GetSetting("ZLSOFT", strRegisterPath, "DeviceName", "")
    grsPage.Filter = ""
    If grsPage.RecordCount > 0 Then
        blnFirst = True
        intPages = grsPage.RecordCount
        For intLoop = 1 To intPages
            grsPage.Filter = "页号=" & intLoop

'            If strPage <> "" Then
'
'                '按虚拟页号计算
'                intPage = Val(grsPage("虚拟页号").Value)
'
'                If InStr("," & strPage & ",", "," & intPage & ",") = 0 Then
'                    '在求奇偶数页
'                    If strPage = "-1" Or strPage = "-2" Then
'                        If strPage = "-1" Then
'                            '奇数打印
'                            If intPage Mod 2 = 1 Or intPage = 0 Then
'                                '打印
'                            Else
'                                grsPage.Filter = "页号=0"
'                            End If
'                        Else
'                            '偶数打印
'                            If (intPage Mod 2 = 0) And intPage <> 0 Then
'                                '打印
'                            Else
'                                grsPage.Filter = "页号=0"
'                            End If
'                        End If
'                    Else
'                        grsPage.Filter = "页号=0"
'                    End If
'
'                End If
'
'            End If
            
            intPage = Val(grsPage("虚拟页号").Value)
'
            If mstrPrintPage = "" Then
            
            ElseIf InStr("," & mstrPrintPage & ",", "," & intPage & ",") > 0 Then
                
            Else
                    grsPage.Filter = "页号=0"
            End If

            
            
            If grsPage.RecordCount > 0 Then

                intPrintPages = intPrintPages + 1

                If intPrintPages > 1 Then Printer.NewPage
               
                    If blnFirst Then
                        blnFirst = False
    
                        On Error Resume Next
                        err = 0
                        Printer.Print ""
                        If err <> 0 Then
                            Printer.EndDoc
                            Exit Function
                        End If
                        On Error GoTo 0
                    End If
    
                    Call ShowPage(Printer, intLoop)
            End If
        Next
        Printer.EndDoc
    End If

    PrintReport = True

End Function

''######################################################################################################################
'
Private Function ShowPage(ByRef objDraw As Object, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：输出数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnFootHead As Boolean



    On Error Resume Next
    objDraw.Cls
    On Error GoTo 0

    '先画背景
    grsData.Filter = ""
    grsData.Filter = "页号=" & intPage & " And 对象='背景'"
    grsData.Sort = "序号"
    If grsData.RecordCount > 0 Then
        grsData.MoveFirst
        Do While Not grsData.EOF
            Call OutputObject(objDraw, grsData)
            grsData.MoveNext
        Loop
    End If

    '再画非背景
    grsData.Filter = ""
    grsData.Filter = "页号=" & intPage & " And 对象<>'背景'"
    grsData.Sort = "序号"
    If grsData.RecordCount > 0 Then
        grsData.MoveFirst
        Do While Not grsData.EOF
            Call OutputObject(objDraw, grsData)
            grsData.MoveNext
        Loop
    End If


    grsPage.Filter = ""
    grsPage.Filter = "页号=" & intPage
    grsData.Sort = "序号"
    If grsPage.RecordCount > 0 Then
        grsPage.MoveFirst
        If Val(grsPage("显示页眉").Value) = 1 Then
            grsData.Filter = ""
            grsData.Filter = "类别='页眉'"
            grsData.Sort = "序号"
            If grsData.RecordCount > 0 Then
                grsData.MoveFirst
                Do While Not grsData.EOF
                    Call OutputObject(objDraw, grsData)
                    grsData.MoveNext
                Loop
            End If
        End If


        If Val(grsPage("显示页脚").Value) = 1 Then
            grsData.Filter = ""
            grsData.Filter = "类别='页脚'"
            grsData.Sort = "序号"
            If grsData.RecordCount > 0 Then
                grsData.MoveFirst
                Do While Not grsData.EOF
                    Call OutputObject(objDraw, grsData, Val(grsPage("虚拟页号").Value), Val(grsPage("虚拟总页").Value))
                    grsData.MoveNext
                Loop
            End If
        End If

    End If

    ShowPage = True

End Function
'
Private Function AddPage(ByVal intPage As Integer) As Boolean

    Load picPage(intPage)
    picPage(intPage).ZOrder

End Function
'
Private Function OutputObject(ByRef objDraw As Object, ByVal rs As ADODB.Recordset, Optional ByVal intPage As Integer, Optional ByVal intPageTotal As Integer) As Boolean
    '******************************************************************************************************************
    '功能：输出数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPic As StdPicture
    Dim X As Long
    Dim Y As Long
    Dim Y1 As Long
    Dim X1 As Long
    Dim strTmp As String
    Dim strTmpLine As String
    Dim intLoop As Integer
    Dim intCharNumber As String
    Dim strChar As String

    strTmp = gobjCommFun.Nvl(rs("内容").Value)

    '调试
    If gobjCommFun.Nvl(rs("调试").Value, 0) = 1 Then
        strTmp = gobjCommFun.Nvl(rs("内容").Value)
    End If

    Select Case rs("对象").Value
    '------------------------------------------------------------------------------------------------------------------
    Case "文本", "页码", "续页"

        objDraw.FontName = Trim(rs("字体").Value)
        objDraw.FontSize = Val(rs("大小").Value)
        objDraw.FontBold = IIf(Val(rs("粗体").Value) = 1, True, False)
        objDraw.FontItalic = IIf(Val(rs("斜体").Value) = 1, True, False)

        If rs("对象").Value = "页码" Then

            If strTmp = "" Then
                strTmp = "第 n 页 / 共 m 页"
            End If

            If strTmp <> "" Then
                strTmp = Replace(strTmp, "n", intPage)
                strTmp = Replace(strTmp, "m", intPageTotal)
            End If

        Else
            strTmp = rs("内容").Value
        End If

        X = Val(rs("X0").Value)

        Select Case Val(rs("横向对齐").Value)
        Case 1

        Case 2
            X = X + (Val(rs("X1").Value) - Val(rs("X0").Value) - objDraw.TextWidth(strTmp)) / 2
        Case 3
            X = X + Val(rs("X1").Value) - Val(rs("X0").Value) - objDraw.TextWidth(strTmp)
        End Select

        If Val(rs("自动换行").Value) = 1 And Val(rs("行数").Value) > 1 Then

        ElseIf Val(rs("自动适应").Value) = 1 Then
            '按指定的区域大小打印，打不完时，缩小字号，直到能打印完全为止
            '关键是找出能打完的最大字号

        Else
            Select Case Val(rs("纵向对齐").Value)
            Case 1
                Y = Val(rs("Y0").Value)
            Case 2
                Y = Val(rs("Y0").Value) + (Val(rs("Y1").Value) - Val(rs("Y0").Value) - objDraw.TextHeight(strTmp)) / 2
            Case 3
                Y = Val(rs("Y0").Value) + Val(rs("Y1").Value) - Val(rs("Y0").Value) - objDraw.TextHeight(strTmp)
            End Select

            Select Case Val(rs("旋转角度").Value)
            Case 0                  '正常
                Call DrawText(objDraw, X, Y, strTmp, Val(rs("前景色").Value))
            Case 4                  '从下到上排并右转90度

                If Trim(strTmp) <> "" Then

                    Y1 = Val(rs("Y1").Value)
                    X1 = Val(rs("X0").Value)

                    intCharNumber = 0
                    For intLoop = 1 To Len(strTmp)

                        If Y1 > Val(rs("Y0").Value) Then
                            strChar = Mid(strTmp, intLoop, 1)

'                            If TypeName(objDraw) = "PictureBox" Then
                                Call DrawText(objDraw, X1 + 15, Y1 + 15, strChar, Val(rs("前景色").Value), 90)
'                            Else
'                                Call DrawRotationText(X1 + 15, Y1 + 15, strChar, Val(rs("前景色").Value), 90)
'                            End If

                            If Asc(strChar) < 0 Then
                                Y1 = Y1 - objDraw.TextWidth("AA")
                            Else
                                Y1 = Y1 - objDraw.TextWidth("A")
                            End If

                        End If
                    Next
                End If

'                Call DrawText(objDraw, x, y, strTmp, Val(rs("前景色").Value))

            Case Else
                Call DrawText(objDraw, X, Y, strTmp, Val(rs("前景色").Value))
            End Select

            If Val(grsData("下划线").Value) = 1 Then
                Call DrawLine(objDraw, Val(rs("X0").Value), Val(rs("Y1").Value) + 30, Val(rs("X1").Value), Val(rs("Y1").Value) + 30, Val(rs("前景色").Value), 0, 1, False)
            End If

        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "线条"

        Call DrawLine(objDraw, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value), Val(rs("Y1").Value), Val(rs("前景色").Value), Val(rs("线条类型").Value), Val(rs("线条宽度").Value), False)

    '------------------------------------------------------------------------------------------------------------------
    Case "线条1"

        Call DrawLine(objDraw, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value), Val(rs("Y1").Value), Val(rs("前景色").Value), Val(rs("线条类型").Value), Val(rs("线条宽度").Value), False)

'    '------------------------------------------------------------------------------------------------------------------
'    Case "条码"
'
'        If strTmp <> "" Then
'            Set objBarCode = New clsBarCode
'            Set objPic = objBarCode.DrawBarCode128(picTmp, 3, strTmp, False)
'            objDraw.PaintPicture objPic, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value) - Val(rs("X0").Value), Val(rs("Y1").Value) - Val(rs("Y0").Value), 0, 0, Val(rs("X1").Value) - Val(rs("X0").Value), Val(rs("Y1").Value) - Val(rs("Y0").Value)
'        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "背景"

        objDraw.Line (Val(rs("X0").Value), Val(rs("Y0").Value))-(Val(rs("X1").Value), Val(rs("Y1").Value)), Val(rs("背景色").Value), BF

    '------------------------------------------------------------------------------------------------------------------
    Case "图形"

        Call DrawPicture(objDraw, grsData("内容").Value, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value), Val(rs("Y1").Value))
'        Call DrawPicture(objDraw, grsData("内容").Value, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value), Val(rs("Y1").Value), Val(rs("横向对齐").Value))

    End Select

    OutputObject = True

End Function
'
Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    '******************************************************************************************************************
    '功能：根据相册大小自动等比例缩放相片文件
    '参数：缩放前的相片文件
    '返回：缩放后的相片文件
    '******************************************************************************************************************
    Dim strTmp As String
    Dim objMap As StdPicture
    Dim W As Single
    Dim H As Single
    Dim sglPerW As Single
    Dim sglPerH As Single
    Dim sglPer As Single
    Dim cx As Long
    Dim cy As Long

    On Error GoTo errHand

    If strFile = "" Then Exit Function

    cx = X2 - X1
    cy = Y2 - Y1

    Set objMap = VB.LoadPicture(strFile)

    W = objMap.Width * 0.566950910348006
    H = objMap.Height * 0.566950910348006

    sglPerW = 1
    sglPerH = 1

    If W > cx Then sglPerW = cx / W
    If H > cy Then sglPerH = cy / H

    If W > cx Or H > cy Then
        sglPer = IIf(sglPerW > sglPerH, sglPerH, sglPerW)
        W = W * sglPer
        H = H * sglPer
    End If

'    objDraw.PaintPicture objMap, X1, Y1, W, H
    objDraw.PaintPicture objMap, X1 + (cx - W) / 2, Y1 + (cy - H) / 2, W, H, 0, 0, objMap.Width * 0.566950910348006, objMap.Height * 0.566950910348006

    DrawPicture = True


    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    ShowSimpleMsg "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!"
    ShowSimpleMsg err.Description
'    Resume
End Function
'
'
Private Function CheckPrint(ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：检查打印机
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strPrintName As String
    Dim blnYesPrinter As Boolean
    Dim intLoop As Integer

    '打印机恢复及设置
    If Printers.count < 1 Then
        MsgBox "系统没有安装任何打印机不能继续打印！", vbInformation, "检验报告打印"
        Exit Function
    End If

    strPrintName = GetSetting("ZLSOFT", strRegisterPath, "DeviceName", "")

    If strPrintName = "" Then
        MsgBox "没有设置打印机,将使用系统默认打印机设置！", vbInformation, "zl9OpsFormat"
    Else

        '打印机
        blnYesPrinter = False
        If Printer.DeviceName <> strPrintName Then
            For intLoop = 0 To Printers.count - 1
                If Printers(intLoop).DeviceName = strPrintName Then Set Printer = Printers(intLoop): blnYesPrinter = True: Exit For
            Next
            If blnYesPrinter = False Then
                MsgBox "设置的打印机已不存在,将使用系统默认打印机设置！", vbInformation, "zl9OpsFormat"
            End If
        End If

    End If

    CheckPrint = True

End Function
'
Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    cbsMain.DeleteAll
    
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_PrintSet, "打印设置")
    If UCase(UserInfo.用户名) = "ZLHIS" Or UserInfo.姓名 Like "%管理%" Then
'        Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Order, "调整顺序")
'        Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Setting, "参数设置")
    End If
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigatebeginning, "首页", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateleft, "上页")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateright, "下页")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateend, "末页")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_PrintPageSet, "打印页码设置", True)
    If gbtyModel = 2 Then
        Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    End If

    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "页面")
    objControl.Flags = xtpFlagRightAlign

    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, conMenu_View_Page, "页面")
    cbrCustom.Handle = cboPage.hWnd
    cbrCustom.Flags = xtpFlagRightAlign

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
    End With

End Function

Private Sub cboPage_Click()
    Dim intLoop As Integer

    If mblnStartUp Then Exit Sub

    picBack.Cls

    For intLoop = 0 To picPage.UBound
        If intLoop = cboPage.ListIndex Then
            mintCurPage = intLoop

            picPage(intLoop).Visible = True
            picPage(intLoop).ZOrder

            If picPage(intLoop).Tag = "" Then
                picPage(intLoop).Tag = "装载"
                Call ShowPage(picPage(intLoop), intLoop + 1)
            End If

        Else

            picPage(intLoop).Visible = False
        End If
    Next

    Call cbsMain_Resize

End Sub
'
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim aryTmp          As Variant
    Dim objDrawReport   As Object

    Select Case Control.ID
        Case conMenu_File_PrintSet
            gobjPrintMode.zlPrintSet
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Print
            If CheckPrint(mstrRegisterPath) = False Then Exit Sub
            If PrintReport Then
                RaiseEvent AfterPrinted
            End If
    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Refresh               '刷新
             
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Navigatebeginning
            cboPage.ListIndex = 0
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Navigateleft
            If cboPage.ListIndex - 1 >= 0 Then cboPage.ListIndex = cboPage.ListIndex - 1
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Navigateright
            If cboPage.ListIndex + 1 < cboPage.ListCount Then cboPage.ListIndex = cboPage.ListIndex + 1
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Navigateend
            cboPage.ListIndex = cboPage.ListCount - 1
        Case conMenu_View_Order
'            frmSequence.ShowMe (1)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_PrintPageSet
            mstrPrintPage = frmPrintPageSet.ShowMe(cboPage.ListCount)
        Case Else
            Call CommandBarExecutePublic(Control, Me)
    End Select
End Sub
'
Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    If WindowState = 1 Then Exit Sub

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
'
    With picBack
        .Left = lngLeft
        .Top = lngTop - 510
        .Width = Me.ScaleWidth - scrVsc.Width
        .Height = Me.ScaleHeight - scrHsc.Height
    End With
'    picBack.Visible = True
'
'    picBack.Move lngLeft, lngTop - 510, lngRight - lngLeft - scrVsc.Width, lngBottom - lngTop - scrHsc.Height
    scrVsc.Move picBack.Width, picBack.Top, scrVsc.Width, picBack.Height
    scrHsc.Move picBack.Left, picBack.Top + picBack.Height, picBack.Width, scrHsc.Height

'    picShadow.Move picShadow.Left, picShadow.Top, picPage(mintCurPage).Width, picPage(mintCurPage).Height

    '调整预览页

    If picBack.ScaleWidth >= picPage(mintCurPage).Width + SHADOW_W * 4 Then
        picPage(mintCurPage).Left = (picBack.ScaleWidth - (picPage(mintCurPage).Width + SHADOW_W * 4)) / 2 + SHADOW_W * 2
        picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picPage(mintCurPage).Width + SHADOW_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.Value = 0
        scrHsc.Enabled = True
        scrhsc_Change
    End If
    If picBack.ScaleHeight >= picPage(mintCurPage).Height + SHADOW_W * 4 Then
        picPage(mintCurPage).Top = (picBack.ScaleHeight - (picPage(mintCurPage).Height + SHADOW_W * 4)) / 2 + SHADOW_W
        picShadow.Top = picPage(mintCurPage).Top + SHADOW_W
        scrVsc.Enabled = False
    Else
        scrVsc.Max = (picPage(mintCurPage).Height + SHADOW_W * 4 - picBack.ScaleHeight) / 15
        If scrVsc.Max / 3 < scrVsc.SmallChange Then
            scrVsc.LargeChange = scrVsc.SmallChange
        Else
            scrVsc.LargeChange = scrVsc.Max / 3
        End If
        scrVsc.Value = 0
        scrVsc.Enabled = True
        scrVsc_Change
    End If

End Sub
'
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errHand

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
        Control.Enabled = (cboPage.ListCount > 0)

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Navigatebeginning

        Control.Enabled = (cboPage.ListIndex > 0 And cboPage.ListCount > 0)

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Navigateleft

        Control.Enabled = (cboPage.ListIndex > 0 And cboPage.ListCount > 0)

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Navigateright

        Control.Enabled = (cboPage.ListIndex < cboPage.ListCount - 1 And cboPage.ListCount > 0)

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Navigateend

        Control.Enabled = (cboPage.ListIndex < cboPage.ListCount - 1 And cboPage.ListCount > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintPageSet
        Control.Enabled = (cboPage.ListCount > 0)
    Case Else
         Call CommandBarUpdatePublic(Control, Me)
    End Select
    Exit Sub
errHand:

End Sub
'
Private Sub Form_Load()
    mintCurPage = 0
    
    Call InitCommandBar

    mlngLeft = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7
    mlngWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    mintPage = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    mstrPrintPage = ""
    Call RestoreWinState(Me, "检验报告打印")
    glngTop = scrVsc.Top
    glngLeft = scrHsc.Left
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    glngTop = X
    glngLeft = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrintPage = ""
    Call gobjComLib.SaveWinState(Me, "检验报告打印")
End Sub
'
Private Sub scrVsc_Change()
    picPage((mintCurPage)).Top = -scrVsc.Value * 15# + SHADOW_W * 2
    picShadow.Top = picPage(mintCurPage).Top + SHADOW_W
    Me.Refresh
End Sub
'
Private Sub scrVsc_Scroll()
    picPage(mintCurPage).Top = -scrVsc.Value * 15# + SHADOW_W * 2
    picShadow.Top = picPage(mintCurPage).Top + SHADOW_W
    Me.Refresh
End Sub
'
Private Sub scrhsc_Change()
    picPage(mintCurPage).Left = -scrHsc.Value * 15# + SHADOW_W * 2
    picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
    Me.Refresh
End Sub
'
Private Sub scrhsc_Scroll()
    picPage(mintCurPage).Left = -scrHsc.Value * 15# + SHADOW_W * 2
    picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
    Me.Refresh
End Sub
'
Private Sub DrawLine(pic As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1, Optional ByVal blnEndArrow As Boolean)

    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    Dim lngLoop As Long

    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)

    If blnEndArrow Then

        If Y1 < Y2 Then
            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 - 75)
            Next
        Else

            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 + 75)
            Next

        End If
    End If

    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth

End Sub
'
Private Sub DrawText(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0, Optional Rotation As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long

    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        objDraw.FontTransparent = True
        .CurrentX = X
        .CurrentY = Y
        If Rotation = 0 Then
            objDraw.Print Text
        Else
            Call DrawRotateText(objDraw, .CurrentX, .CurrentY, Text, Rotation)
        End If

        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Sub DrawRotateText(mobjDrawObject As Object, X As Single, Y As Single, strText As String, ByVal Degrees As Long)
    '在(X,Y)处旋转输出Text文本
    Dim lngFont As Long
    Dim lngNewFont As Long
    Dim lf As LOGFONT
    Dim hWnd As Long
    Dim hDC As Long
    Dim hOldfont As Long
    Dim hPrintDc As Long

    If mobjDrawObject Is Printer Then
        hDC = Printer.hDC
    Else
        hWnd = GetDesktopWindow
        hDC = GetDC(hWnd)
    End If
    
    With lf
          .lfHeight = -(mobjDrawObject.Font.Size * GetDeviceCaps(hDC, LOGPIXELSY)) / 72
          .lfWidth = 0
          .lfEscapement = Degrees * 10
          .lfOrientation = .lfEscapement
          .lfWeight = mobjDrawObject.Font.Weight
          .lfItalic = mobjDrawObject.Font.Italic
          .lfUnderline = mobjDrawObject.Font.Underline
          .lfStrikeOut = mobjDrawObject.Font.Strikethrough
          .lfClipPrecision = CLIP_DEFAULT_PRECIS
          .lfQuality = PROOF_QUALITY
          .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
          .lfFaceName = mobjDrawObject.Font.Name & vbNullChar
          .lfCharSet = mobjDrawObject.Font.Charset
          If .lfCharSet = OEM_CHARSET Then
                If (Degrees Mod 360) <> 0 Then
                      .lfCharSet = ANSI_CHARSET
                End If
          End If
          If (Degrees Mod 360) <> 0 Then
                .lfOutPrecision = OUT_TT_ONLY_PRECIS
          Else
                .lfOutPrecision = OUT_DEFAULT_PRECIS
          End If
    End With

    If Not (mobjDrawObject Is Printer) Then
        Call ReleaseDC(hWnd, hDC)
    End If

    lngNewFont = CreateFontIndirect(lf)

    With mobjDrawObject
        If mobjDrawObject Is Printer Then
                X = X / Printer.TwipsPerPixelX
                Y = Y / Printer.TwipsPerPixelY
        End If
    End With

    hPrintDc = mobjDrawObject.hDC
    lngFont = SelectObject(hPrintDc, lngNewFont)
    If mobjDrawObject Is Printer Then
        Call TextOut(hPrintDc, X, Y, strText, LenB(StrConv(strText, vbFromUnicode)))
    Else
        mobjDrawObject.CurrentX = X
        mobjDrawObject.CurrentY = Y
        mobjDrawObject.Print strText
    End If

    Call SelectObject(hPrintDc, lngFont)
    Call DeleteObject(lngNewFont)

End Sub

Private Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：恢复窗体的状态，当左顶边界超出时，则自动设置为0
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object, objSub As Object
    Dim strIndex As String, blnDo As Boolean
    Dim blnAutoHIDe As Boolean
    
    On Error Resume Next
    
    blnDo = Val(gobjDatabase.GetPara("使用个性化风格")) <> 0
    blnAutoHIDe = Val(gobjDatabase.GetPara("界面区域隐藏")) = 1
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '恢复窗体的状态、位置、大小
    strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    If blnDo Then
        aryInfo = Split(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", strTmp), ",")
    Else
        aryInfo = Split(strTmp, ",")
    End If
    With objForm
        .WindowState = aryInfo(0)
        If UBound(aryInfo) = 4 Then
            .Left = IIf(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIf(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIf(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIf(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    '恢复窗体中各种控件的各种状态
    For Each objThis In objForm.Controls
        If blnDo Then
            strTmp = "": strIndex = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '特殊菜单的复选
                If objThis.Caption Like "标准按钮*" Or _
                    objThis.Caption Like "文本标签*" Or _
                    objThis.Caption Like "状态栏*" Or _
                    UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                    
                    strIndex = objThis.index
                    If err.Number <> 0 Then err.Clear: strIndex = ""
                    
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & strIndex & "状态", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strIndex = objThis.index
                If err.Number <> 0 Then err.Clear: strIndex = ""
                
                strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & strIndex & "状态", "")
                If strTmp <> "" Then
                    'Left,Top,Width、Height,Visible
                    If UBound(Split(strTmp, ",")) = 4 Then
                        If Split(strTmp, ",")(0) <> "-32767" Then objThis.Left = Split(strTmp, ",")(0)
                        If Split(strTmp, ",")(1) <> "-32767" Then objThis.Top = Split(strTmp, ",")(1)
                        If Split(strTmp, ",")(2) <> "-32767" Then objThis.Width = Split(strTmp, ",")(2)
                        If Split(strTmp, ",")(3) <> "-32767" Then objThis.Height = Split(strTmp, ",")(3)
                        If Split(strTmp, ",")(4) <> "-32767" Then objThis.Visible = Split(strTmp, ",")(4)
                    End If
                End If
            End If
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("CommandBars") 'CommandBar
                If blnDo Then
                    If objThis.ActiveMenuBar.Visible Then '有菜单栏的才处理
                        '状态栏
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "StatusBarVisible", "")
                        If strTmp <> "" Then objThis.StatusBar.Visible = Val(strTmp) <> 0
                        '标准按钮,文本标签
                        If objThis.count >= 2 Then
                            strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "Visible", "")
                            If strTmp <> "" Then
                                For i = 2 To objThis.count
                                    objThis(i).Visible = Val(strTmp) <> 0
                                Next
                            End If
                            strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "ButtonText", "")
                            If strTmp <> "" Then
                                For i = 2 To objThis.count
                                    For Each objSub In objThis(i).Controls
                                        objSub.Style = Val(strTmp)
                                    Next
                                Next
                            End If
                        End If
                        '大图标
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "LargeIcon", "")
                        If strTmp <> "" Then objThis.Options.LargeIcons = Val(strTmp) <> 0
                    End If
                End If
            Case UCase("DockingPane") 'DockingPane
                If blnDo Then
                    'strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, "")
                    'If strTmp <> "" Then objThis.LoadStateFromString strTmp
                End If
                If Not blnAutoHIDe Then
                    'PaneNoHIDeable(2) Or PaneNoCloseable(1) or PaneNoFloatable(4)，没有HIDe会自动出现Close,允许Float双击有问题
                    For i = 1 To objThis.PanesCount
                        'PaneNoCaption(8)
                        If objThis.Panes(i).Options And 8 = 8 Then
                            objThis.Panes(i).Options = 15
                        Else
                            objThis.Panes(i).Options = 7
                        End If
                    Next
                    DeleteSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis)
                    DeleteSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Replace(strProjectName, "\", "") & objForm.Name & strUserDef & "\" & TypeName(objThis)
                End If
            Case UCase("ReportControl") 'ReportControl
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, "")
                    If strTmp <> "" Then objThis.LoadSettings strTmp
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "文本", 1)
                        For i = 1 To objThis.Buttons.count
                            objThis.Buttons(i).Caption = IIf(strTmp = 1, objThis.Buttons(i).Tag, "")
                        Next
                    End If
                End If
            Case UCase("ListView")
                If blnDo Then
                    strIndex = objThis.index
                    If err.Number <> 0 Then err.Clear: strIndex = ""
                    gobjComLib.RestoreListViewState objThis, strProjectName & objForm.Name & strUserDef, strIndex
                End If
            Case UCase("CoolBar")
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "行序", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "可见栏", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).Visible = Split(strTmp, ",")(i)
                        Next
                    End If
                End If
            Case UCase("MSHFlexGrID"), UCase("BillEdit"), UCase("VSFlexGrID")
                If blnDo Then
                    gobjComLib.RestoreFlexState objThis, strProjectName & objForm.Name & strUserDef
                End If
            Case UCase("DataGrID")
                If blnDo Then
                    gobjComLib.RestoreDBGrIDState objThis, strProjectName & objForm.Name & strUserDef
                End If
        End Select
    Next
    RestoreWinState = True
End Function

Public Sub HideControl(ByVal lngPatientID As Long, ByVal blnState As Boolean)
    Dim intLoop As Integer
    
    '不管有没有数据先清除页码显示
    
    If lngPatientID = 0 Or Not blnState Then
        cboPage.Clear
        '清空原来页面的数据
'        For intLoop = picPage.UBound To 0 Step -1
'            picPage(intLoop).Visible = False
'        Next
        picBack.Visible = False
    Else
        '清空原来页面的数据
'        For intLoop = picPage.UBound To 0 Step -1
'            picPage(intLoop).Visible = True
'        Next
        picBack.Visible = True
    End If
    mstrPrintPage = ""
End Sub

Public Function GetPageCount() As Long
    GetPageCount = cboPage.ListCount
End Function

