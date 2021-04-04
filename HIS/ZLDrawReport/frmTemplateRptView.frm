VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTemplateRptView 
   Caption         =   "预览"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11160
   Icon            =   "frmTemplateRptView.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11160
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   8595
      MouseIcon       =   "frmTemplateRptView.frx":6852
      MousePointer    =   99  'Custom
      ScaleHeight     =   1920
      ScaleWidth      =   1785
      TabIndex        =   8
      Top             =   3870
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.PictureBox picTmp 
      Height          =   2010
      Left            =   10125
      ScaleHeight     =   1950
      ScaleWidth      =   3825
      TabIndex        =   7
      Top             =   4125
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   7515
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   405
      Width           =   1860
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmTemplateRptView.frx":69A4
      Height          =   250
      LargeChange     =   20
      Left            =   615
      Max             =   100
      MouseIcon       =   "frmTemplateRptView.frx":6CAE
      SmallChange     =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6030
      Width           =   8760
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmTemplateRptView.frx":6E00
      Height          =   4755
      LargeChange     =   20
      Left            =   9390
      Max             =   100
      MouseIcon       =   "frmTemplateRptView.frx":710A
      SmallChange     =   10
      TabIndex        =   4
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
      Left            =   375
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   8760
      Begin VB.PictureBox picAppendix 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   6030
         ScaleHeight     =   2715
         ScaleWidth      =   2700
         TabIndex        =   9
         Top             =   1470
         Width           =   2700
         Begin MSComctlLib.ListView lvwAppendix 
            Height          =   1290
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   2275
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "img16"
            SmallIcons      =   "img16"
            ColHdrIcons     =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "附件"
               Object.Width           =   4939
            EndProperty
         End
      End
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   270
         MouseIcon       =   "frmTemplateRptView.frx":725C
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   2
         Top             =   180
         Width           =   6990
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   6990
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6975
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTemplateRptView.frx":73AE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14605
            Object.ToolTipText     =   "打印机信息"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   4080
      MouseIcon       =   "frmTemplateRptView.frx":7C42
      MousePointer    =   99  'Custom
      ScaleHeight     =   1920
      ScaleWidth      =   1785
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplateRptView.frx":7D94
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemplateRptView.frx":80AE
            Key             =   "PDF"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmTemplateRptView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'常量

'常量定义
Private Const OFFSET_LEFT = 20
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

Private mclsPrint As New clsPrint

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
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long       ' or Boolean
Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long                                                                         '   or   Boolean

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
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
Private mclsImage As clsImage
Public Event AfterPrinted()
Private mblnPreview As Boolean
Private mbytMode As Byte

'######################################################################################################################
'过程清单

Public Function InitReport(ByVal rsData As ADODB.Recordset, ByVal rsPage As ADODB.Recordset, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：初始化处理
    '参数：
    '返回：
    '******************************************************************************************************************
    
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

Public Function ExportReport(Optional ByVal bytMode As Byte = 1, Optional ByVal strTitle As String, Optional ByVal strFile As String, Optional ByVal strPassword As String, Optional ByVal strPage As String, Optional ByVal bytPageStype As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：输出报告
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrTitle = strTitle & "预览"
    mbytMode = bytMode
    Select Case bytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 2                          '打印
        
        If PrintReport(strPage, bytPageStype) Then
            RaiseEvent AfterPrinted
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case 1                          '预览
        mblnPreview = True
        Call PreviewReport(strPage)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 3                          '输出到PDF
        
        If ExportPDFReport(strFile, strPassword, strPage, bytPageStype) Then
            RaiseEvent AfterPrinted
        End If
        
    End Select
    
    ExportReport = True
    
End Function

Private Function PreviewReport(Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '功能：打印预览
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intTmpPage As Integer
    Dim strShow As String
    Dim strTmp As String
    Dim varTmp As Variant
    Dim lngIndex As Long
    
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
            
            intTmpPage = Val(grsPage("页号").value) - 1
            If picPage.UBound < intTmpPage Then Call AddPage(intTmpPage)
                
            If grsPage("显示内容").value <> "" Then
                strShow = grsPage("显示内容").value
            Else
                strShow = "第 " & grsPage("虚拟页号").value & " 页"
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
    
    
    
    '加载附件
    intLoop = 1
    
    On Error GoTo errHand
    
    If grsData.RecordCount > 0 Then
        grsData.MoveFirst
        grsData.Filter = ""
        grsData.Filter = "类别='附件' or 类别='报表'"
        If grsPage.RecordCount > 0 Then
            
            Do While Not grsData.EOF
                strTmp = zlStr.NVL(grsData("内容").value)
                varTmp = Split(strTmp, "^")
                If InStr(strTmp, "^") > 0 Then
                    Call lvwAppendix.ListItems.Add(, grsData("类别") & "^" & grsData("内容").value, varTmp(0), , IIf(grsData("类别") = "附件", 2, 1))
                End If
                intLoop = intLoop + 1
                
                grsData.MoveNext
            Loop
        Else
            picAppendix.Visible = False
        End If
    End If
    
    Me.Caption = mstrTitle
    Me.Show 1
    
    PreviewReport = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub OutPutReport(ByVal str内容 As String)
    '******************************************************************************************************************
    '功能：输出报表
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objclsReport As Object
    Dim varStr As Variant
    Dim varTmp As Variant
    
    If Trim(str内容) = "" Then Exit Sub
    
    On Error Resume Next
    Set objclsReport = CreateObject("zl9Report.clsReport")
    Err.Clear
    
    On Error GoTo errHand
    
    If objclsReport Is Nothing Then Exit Sub
    varTmp = Split(str内容, "^")
    If UBound(varTmp) = 0 Then
        varStr = Split(varTmp(0), "'")
    Else
        varStr = Split(varTmp(1), "'")
    End If
    If UBound(varStr) <> 3 Then Exit Sub
    Call objclsReport.ReportOpen(gcnOracle, ParamInfo.系统号, varStr(0), Me, "任务id=" & Val(varStr(1)), "病人id=" & Val(varStr(2)), "清单id=" & Val(varStr(3)), "PrintEmpty=0", mbytMode)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ExportPDFReport(Optional ByVal strFile As String, Optional ByVal strPassword As String, Optional ByVal strPage As String, Optional ByVal bytPageStype As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：输出为PDF
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intPages As Integer
    Dim intPage As Integer
    Dim clsPrintPDF As clsPDF
    Dim blnFirst As Boolean
    Dim strPageCode As String
    
    On Error GoTo errHand
    
    grsPage.Filter = ""
    If grsPage.RecordCount > 0 Then

        For intLoop = 0 To Printers.count - 1
            If Printers(intLoop).DeviceName = "TinyPDF" Then
                Set Printer = Printers(intLoop): Exit For
            End If
        Next
        If intLoop = Printers.count Then Exit Function

        Set clsPrintPDF = New clsPDF
        If clsPrintPDF.InitPDF(strFile, , , True, strPassword) = False Then Exit Function
        blnFirst = True
        
        intPages = grsPage.RecordCount
        For intLoop = 1 To intPages
            grsPage.Filter = "页号=" & intLoop
            
            If strPage <> "" Then
                '按虚拟页号计算
                
                strPageCode = IIf(bytPageStype = 1, "虚拟页号", "页号")
                
                intPage = Val(grsPage(strPageCode).value)
                If InStr("," & strPage & ",", "," & intPage & ",") = 0 Then
                    grsPage.Filter = "页号=0"
                End If
            End If
            
            If grsPage.RecordCount > 0 Then
                If intLoop > 1 Then Printer.NewPage
                
                If blnFirst Then
                    blnFirst = False
                    
                    On Error Resume Next
                    Err = 0
                    Printer.Print ""
                    If Err <> 0 Then
                        Printer.EndDoc
                        Exit Function
                    End If
                    On Error GoTo errHand
                End If
                
                Call ShowPage(Printer, intLoop)
                
            End If
            
        Next
        
        Printer.EndDoc
        
        Call clsPrintPDF.ResetPDF
        
    End If

    ExportPDFReport = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrintReport(Optional ByVal strPage As String, Optional ByVal bytPageStype As Byte = 1) As Boolean
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
    Dim clsPDFOutput As clsPDFData
    Dim strErr As String
    Dim i As Integer
    Dim varTmp As Variant
    Dim strPageCode As String
    
    grsPage.Filter = ""
    If grsPage.RecordCount > 0 Then
        blnFirst = True
        intPages = grsPage.RecordCount
        For intLoop = 1 To intPages
            grsPage.Filter = "页号=" & intLoop
            
            If strPage <> "" And Val(strPage) <> 0 Then

                '按虚拟页号计算
                strPageCode = IIf(bytPageStype = 1, "虚拟页号", "页号")
                
                intPage = Val(grsPage(strPageCode).value)

                If InStr("," & strPage & ",", "," & intPage & ",") = 0 Then
                    '在求奇偶数页
                    If strPage = "-1" Or strPage = "-2" Then
                        If strPage = "-1" Then
                            '奇数打印
                            If intPage Mod 2 = 1 Or (intPage = 0 And bytPageStype = 0) Then
                                '打印
                            Else
                                grsPage.Filter = "页号=0"
                            End If
                        Else
                            '偶数打印
                            If (intPage Mod 2 = 0) And intPage <> 0 Then
                                '打印
                            Else
                                grsPage.Filter = "页号=0"
                            End If
                        End If
                    Else
                        grsPage.Filter = "页号=0"
                    End If
                
                End If

            End If

            If grsPage.RecordCount > 0 Then
                
                intPrintPages = intPrintPages + 1
                
                If intPrintPages > 1 Then Printer.NewPage
                
                If blnFirst Then
                    blnFirst = False
                    
                    On Error Resume Next
                    Err = 0
                    Printer.Print ""
                    If Err <> 0 Then
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
    
    '打印指定页不包含附件
    If strPage <> "" Then
        PrintReport = True
        Exit Function
    End If
    
    '打印附件
    
    grsData.Filter = ""
    If grsData.RecordCount > 0 Then
        grsData.Filter = ""
        grsData.Filter = "类别='附件' or 类别='报表'"
        If grsData.RecordCount > 0 Then
            grsData.MoveFirst
            Do While Not grsData.EOF
                
                If grsData("类别") = "附件" Then
                    '输出PDF文件
                    On Error GoTo errHand
                    If clsPDFOutput Is Nothing Then
                        Set clsPDFOutput = New clsPDFData
                    End If
                    varTmp = Split(zlStr.NVL(grsData("内容").value), "^")
                    If UBound(varTmp) > 0 Then
                        If Dir(varTmp(1), vbDirectory) = "" Then
                            ShowSimpleMsg "打印""" & varTmp(1) & """文件失败,该文件可能不存在!"
                            Err.Clear
                        Else
                            If UBound(varTmp) >= 3 Then
                                clsPDFOutput.FoxitPrint varTmp(3)
                            End If
                        End If
                    End If
                Else
                    '打印自定义报表
                    Call OutPutReport(grsData("内容").value)
                End If
                grsData.MoveNext
            Loop
            
            DoEvents
            Set clsPDFOutput = Nothing
            
            
        End If
    End If

    PrintReport = True
    
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function ExcelReport(Optional ByVal strFileName As String) As Boolean
    '******************************************************************************************************************
    '功能：输入出到Excel文件
    '参数：
    '返回：
    '******************************************************************************************************************
    
    
    ExcelReport = True
    
End Function


'######################################################################################################################

Private Function ShowPage(ByRef objDraw As Object, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：输出数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnFootHead As Boolean
    

    
    On Error Resume Next
    objDraw.Cls
    
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
    
    On Error GoTo 0
    
    grsPage.Filter = ""
    grsPage.Filter = "页号=" & intPage
    grsData.Sort = "序号"
    If grsPage.RecordCount > 0 Then
        grsPage.MoveFirst
        If Val(grsPage("显示页眉").value) = 1 Then
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
        

        If Val(grsPage("显示页脚").value) = 1 Then
            grsData.Filter = ""
            grsData.Filter = "类别='页脚'"
            grsData.Sort = "序号"
            If grsData.RecordCount > 0 Then
                grsData.MoveFirst
                Do While Not grsData.EOF
                    Call OutputObject(objDraw, grsData, Val(grsPage("虚拟页号").value), Val(grsPage("虚拟总页").value))
                    grsData.MoveNext
                Loop
            End If
        End If
        
    End If
    
    ShowPage = True
    
End Function

Private Function AddPage(ByVal intPage As Integer) As Boolean

    Load picPage(intPage)
    picPage(intPage).ZOrder

End Function

Private Function OutputObject(ByRef objDraw As Object, ByVal rs As ADODB.Recordset, Optional ByVal intPage As Integer, Optional ByVal intPageTotal As Integer) As Boolean
    '******************************************************************************************************************
    '功能：输出数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPic As StdPicture
    Dim x As Long
    Dim y As Long
    Dim Y1 As Long
    Dim X1 As Long
    Dim strTmp As String
    Dim strTmpLine As String
    Dim intLoop As Integer
    Dim intCharNumber As String
    Dim strChar As String
    
    strTmp = zlStr.NVL(rs("内容").value)
    
    '调试
    If zlStr.NVL(rs("调试").value, 0) = 1 Then
        strTmp = zlStr.NVL(rs("内容").value)
    End If
    
    Select Case rs("对象").value
    '------------------------------------------------------------------------------------------------------------------
    Case "文本", "页码", "续页"
        
        objDraw.FontName = Trim(rs("字体").value)
        objDraw.FontSize = Val(rs("大小").value)
        objDraw.FontBold = IIf(Val(rs("粗体").value) = 1, True, False)
        objDraw.FontItalic = IIf(Val(rs("斜体").value) = 1, True, False)
        
        If rs("对象").value = "页码" Then

            If strTmp = "" Then
                strTmp = "第 n 页 / 共 m 页"
            End If
            
            If strTmp <> "" Then
                strTmp = Replace(strTmp, "n", intPage)
                strTmp = Replace(strTmp, "m", intPageTotal)
            End If
            
        Else
            strTmp = rs("内容").value
        End If
        
        x = Val(rs("X0").value)
        
        Select Case Val(rs("横向对齐").value)
        Case 1

        Case 2
            x = x + (Val(rs("X1").value) - Val(rs("X0").value) - objDraw.TextWidth(strTmp)) / 2
        Case 3
            x = x + Val(rs("X1").value) - Val(rs("X0").value) - objDraw.TextWidth(strTmp)
        End Select
                
        If Val(rs("自动换行").value) = 1 And Val(rs("行数").value) > 1 Then
            For intLoop = 1 To Val(rs("行数").value)

                strTmp = GetLineText2(objDraw, rs("内容").value, intLoop, Val(rs("X1").value) - Val(rs("X0").value))

                x = Val(rs("X0").value)
                Select Case Val(rs("横向对齐").value)
                Case 1

                Case 2
                    x = x + (Val(rs("X1").value) - Val(rs("X0").value) - objDraw.TextWidth(strTmp)) / 2
                Case 3
                    x = x + Val(rs("X1").value) - Val(rs("X0").value) - objDraw.TextWidth(strTmp)
                End Select

                y = Val(rs("Y0").value) + Val(rs("B0").value) + (intLoop - 1) * (objDraw.TextHeight(strTmp) + Val(rs("R0").value))

                Select Case Val(rs("纵向对齐").value)
                Case 1                  '上边对齐
'                    Y = Val(rs("Y0").Value) + (intLoop - 1) * (objDraw.TextHeight(strTmp) + Val(rs("R0").Value))
                Case 2                  '居中对齐
                    y = y + (Val(rs("Y1").value) - Val(rs("Y0").value) - Val(rs("行数").value) * (objDraw.TextHeight(strTmp) + Val(rs("R0").value)) - Val(rs("R0").value)) / 2
                Case 3                  '下边对齐
                    y = y + Val(rs("Y1").value) - Val(rs("Y0").value) - Val(rs("行数").value) * (objDraw.TextHeight(strTmp) + Val(rs("R0").value)) - Val(rs("R0").value)
                End Select

                Call DrawText(objDraw, x, y, strTmp, Val(rs("前景色").value))

                If Val(grsData("下划线").value) = 1 Then
                    Call DrawLine(objDraw, Val(rs("X0").value), Val(rs("Y1").value) + 30, Val(rs("X1").value), Val(rs("Y1").value) + 30, Val(rs("前景色").value), 0, 1, False)
                End If

            Next

        ElseIf Val(rs("自动适应").value) = 1 Then
            '按指定的区域大小打印，打不完时，缩小字号，直到能打印完全为止
            '关键是找出能打完的最大字号
            
        Else
            Select Case Val(rs("纵向对齐").value)
            Case 1
                y = Val(rs("Y0").value)
            Case 2
                y = Val(rs("Y0").value) + (Val(rs("Y1").value) - Val(rs("Y0").value) - objDraw.TextHeight(strTmp)) / 2
            Case 3
                y = Val(rs("Y0").value) + Val(rs("Y1").value) - Val(rs("Y0").value) - objDraw.TextHeight(strTmp)
            End Select
            
            Select Case Val(rs("旋转角度").value)
            Case 0                  '正常
                Call DrawText(objDraw, x, y, strTmp, Val(rs("前景色").value))
            Case 4                  '从下到上排并右转90度
                
                If Trim(strTmp) <> "" Then
                
                    Y1 = Val(rs("Y1").value)
                    X1 = Val(rs("X0").value)
                                       
                    intCharNumber = 0
                    For intLoop = 1 To Len(strTmp)
                        
                        If Y1 > Val(rs("Y0").value) Then
                            strChar = Mid(strTmp, intLoop, 1)
                                                        
'                            If TypeName(objDraw) = "PictureBox" Then
                                Call DrawText(objDraw, X1 + 15, Y1 + 15, strChar, Val(rs("前景色").value), 90)
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
                Call DrawText(objDraw, x, y, strTmp, Val(rs("前景色").value))
            End Select
            
            If Val(grsData("下划线").value) = 1 Then
                Call DrawLine(objDraw, Val(rs("X0").value), Val(rs("Y1").value) + 30, Val(rs("X1").value), Val(rs("Y1").value) + 30, Val(rs("前景色").value), 0, 1, False)
            End If
        
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "线条"
    
        Call DrawLine(objDraw, Val(rs("X0").value), Val(rs("Y0").value), Val(rs("X1").value), Val(rs("Y1").value), Val(rs("前景色").value), Val(rs("线条类型").value), Val(rs("线条宽度").value), False)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "线条1"
        
        Call DrawLine(objDraw, Val(rs("X0").value), Val(rs("Y0").value), Val(rs("X1").value), Val(rs("Y1").value), Val(rs("前景色").value), Val(rs("线条类型").value), Val(rs("线条宽度").value), False)
        
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
        
        objDraw.Line (Val(rs("X0").value), Val(rs("Y0").value))-(Val(rs("X1").value), Val(rs("Y1").value)), Val(rs("背景色").value), BF
        
    '------------------------------------------------------------------------------------------------------------------
    Case "图形"
        
        Call DrawPicture(objDraw, grsData("内容").value, Val(rs("X0").value), Val(rs("Y0").value), Val(rs("X1").value), Val(rs("Y1").value), Val(rs("旋转角度").value))
'        Call DrawPicture(objDraw, grsData("内容").Value, Val(rs("X0").Value), Val(rs("Y0").Value), Val(rs("X1").Value), Val(rs("Y1").Value), Val(rs("横向对齐").Value))

    End Select
    
    OutputObject = True
    
End Function

Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lngAngle As Long) As Boolean
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
    Dim CX As Long
    Dim CY As Long

    On Error GoTo errHand

    If strFile = "" Then Exit Function

    CX = X2 - X1
    CY = Y2 - Y1
    
    picHide.Width = CX
    picHide.Height = CY - Y1 * 0.05

    Call mclsImage.InitImage(picHide)
    If lngAngle = 4 Then
        Call mclsImage.ShowImage(strFile, 自动缩放, True, 0, 0, 0, 0, Rotate270FlipxY)
    Else
        Call mclsImage.ShowImage(strFile, 自动缩放, True, 0, 0, 0, 0, RotateNoneFlipNone)
    End If
    Call mclsImage.DisposeImage
    objDraw.PaintPicture picHide.Image, X1, Y1 * 1.05
    DrawPicture = True


    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    ShowSimpleMsg "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!"
    ShowSimpleMsg Err.Description
'    Resume

End Function

Private Function DrawPictureEx(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
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
    Dim CX As Long
    Dim CY As Long

    On Error GoTo errHand

    If strFile = "" Then Exit Function

    CX = X2 - X1
    CY = Y2 - Y1

    Set objMap = VB.LoadPicture(strFile)

    W = objMap.Width * 0.566950910348006
    H = objMap.Height * 0.566950910348006

    sglPerW = 1
    sglPerH = 1

    If W > CX Then sglPerW = CX / W
    If H > CY Then sglPerH = CY / H

    If W > CX Or H > CY Then
        sglPer = IIf(sglPerW > sglPerH, sglPerH, sglPerW)
        W = W * sglPer
        H = H * sglPer
    End If

    objDraw.PaintPicture objMap, X1, Y1, W, H
    DrawPictureEx = True


    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    ShowSimpleMsg "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!"
    ShowSimpleMsg Err.Description
End Function



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function RotateImage(ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As String

    '------------------------------------------------------------------------------------------------------------------
    '第一种方法：
    RotateImage = RotateImage1(strFile, X1, Y1, X2, Y2)


End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'功能：
'参数：
'返回：
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function RotateImage1(ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As String

    picHide.AutoRedraw = True
    picHide.AutoSize = False
    picHide.ScaleMode = 1

    picHide.Width = X2 - X1
    picHide.Height = Y2 - Y1

    Call mclsImage.InitImage(picHide)

    Call mclsImage.LoadImageFile(strFile)

    picHide.Height = mclsImage.ImageWidth * 15
    picHide.Width = mclsImage.ImageHeight * 15

    Call mclsImage.ShowImage(strFile, True, 0, 0, mclsImage.ImageWidth, mclsImage.ImageHeight, Rotate270FlipxY)

    Call mclsImage.DisposeImage
    Set mclsImage = Nothing

    Call SavePicture(picHide.Image, strFile)

    RotateImage1 = strFile
End Function

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
        MsgBox "系统没有安装任何打印机不能继续打印，程序退出！", vbInformation, "zl9OpsFormat"
        Exit Function
    End If
    
    strPrintName = GetSetting("ZLSOFT", strRegisterPath, "打印机", "")
    
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

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, 1, "打印指定页(&C)")
    objControl.IconId = 103
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Navigatebeginning, "首页(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Navigateleft, "上页(&L)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Navigateright, "下页(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Navigateend, "末页(&E)")
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)

    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, 1, "打印指定页")
    objControl.IconId = 103
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigatebeginning, "首页", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateleft, "上页")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateright, "下页")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Navigateend, "末页")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "页面")
    objControl.Flags = xtpFlagRightAlign
            
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, conMenu_View_Page, "页面")
    cbrCustom.Handle = cboPage.hwnd
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
'                picPage(intLoop).Tag = "装载"
                If intLoop > 0 Then picPage(intLoop - 1).Cls
                If intLoop + 1 < picPage.UBound Then picPage(intLoop + 1).Cls
                Call ShowPage(picPage(intLoop), intLoop + 1)
            End If
            
        Else
            
            picPage(intLoop).Visible = False
        End If
    Next
    
    Call cbsMain_Resize
    
End Sub

Private Function ShowPrinterInfo() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    stbThis.Panels(2).Text = "打印机:" & Printer.DeviceName & "/纸张:" & _
        mclsPrint.GetPaperName(Printer.PaperSize) & "/尺寸:" & _
        CLng(Printer.Width / 56.7) & "×" & CLng(Printer.Height / 56.7) & "/纸向:" & _
        IIf(Printer.Orientation = 1, "纵向", "横向")
End Function
'

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim aryTmp As Variant
    Dim bytPageStype As Byte
    Dim strPage As String
    
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
    
        If PrintReport Then
            RaiseEvent AfterPrinted
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case 1
        
        If frmReportPrintPage.ShowDialog(Me, bytPageStype, strPage) Then
            If PrintReport(strPage, bytPageStype) Then RaiseEvent AfterPrinted
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
    '--------------------------------------------------------------------------------------------------------------
    Case Else
        Call CommandBarExecutePublic(Control, Me)
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    
    If WindowState = 1 Then Exit Sub
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    
    picBack.Move lngLeft, lngTop, lngRight - lngLeft - scrVsc.Width, lngBottom - lngTop - scrHsc.Height
    scrVsc.Move picBack.Width, picBack.Top, scrVsc.Width, picBack.Height
    scrHsc.Move picBack.Left, picBack.Top + picBack.Height, picBack.Width, scrHsc.Height
    
    picShadow.Move picShadow.Left, picShadow.Top, picPage(mintCurPage).Width, picPage(mintCurPage).Height
    
    '调整预览页
    
    If picBack.ScaleWidth >= picPage(mintCurPage).Width + SHADOW_W * 4 Then
        If grsData.RecordCount > 0 Then
            grsData.Filter = "类别='附件' or 类别='报表'"
            If grsData.RecordCount > 0 Then
                picPage(mintCurPage).Left = (picBack.ScaleWidth - (picPage(mintCurPage).Width + SHADOW_W * 4 + picAppendix.Width)) / 2 + SHADOW_W * 2
            Else
                picPage(mintCurPage).Left = (picBack.ScaleWidth - (picPage(mintCurPage).Width + SHADOW_W * 4)) / 2 + SHADOW_W * 2
            End If
        Else
            picPage(mintCurPage).Left = (picBack.ScaleWidth - (picPage(mintCurPage).Width + SHADOW_W * 4)) / 2 + SHADOW_W * 2
        End If
        
        picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picPage(mintCurPage).Width + SHADOW_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.value = 0
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
        scrVsc.value = 0
        scrVsc.Enabled = True
        scrVsc_Change
    End If
    picAppendix.Move picPage(mintCurPage).Left * 2 + picPage(mintCurPage).Width, 0
    picAppendix.Height = picBack.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error GoTo errHand
    
    Select Case Control.Id
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
    
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
    Case Else
         Call CommandBarUpdatePublic(Control, Me)
    End Select

errHand:

End Sub

Private Sub Form_Load()
    
    InitGDIPlus
    Set mclsImage = New clsImage
    
    If mblnPreview = False Then Exit Sub
    mintCurPage = 0
    
    Call InitCommandBar
    
    Call RestoreWinState(Me, App.ProductName)

    mlngLeft = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7
    mlngWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    mintPage = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)

    Call ShowPrinterInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Set mclsImage = Nothing
    TerminateGDIPlus
    
End Sub

Private Sub lvwAppendix_DblClick()
    Dim clsPDFOutput As New clsPDFData
    Dim strItem As String
    Dim varStr As Variant
    
    On Error GoTo errHand
    strItem = lvwAppendix.SelectedItem.Key
    varStr = Split(strItem, "^")
    If UBound(varStr) >= 2 Then
        If varStr(0) = "附件" Then
            '预览pdf
             If Not clsPDFOutput.HavePDF(Me) Then
                ShowSimpleMsg "附件加载失败，请检查是否完整安装Adobe Reader阅读器！"
                Exit Sub
            End If
            If UBound(varStr) >= 3 Then
                Call clsPDFOutput.ShellOpen(varStr(3))
            End If
        Else
            '预览报表
            Call OutPutReport(varStr(2))
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picAppendix_Resize()
    On Error Resume Next
    lvwAppendix.Move 200, 100, picAppendix.Width - 400, picAppendix.Height - 200
End Sub

Private Sub scrVsc_Change()
    picPage((mintCurPage)).Top = -scrVsc.value * 15# + SHADOW_W * 2
    picShadow.Top = picPage(mintCurPage).Top + SHADOW_W
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    picPage(mintCurPage).Top = -scrVsc.value * 15# + SHADOW_W * 2
    picShadow.Top = picPage(mintCurPage).Top + SHADOW_W
    Me.Refresh
End Sub

Private Sub scrhsc_Change()
    picPage(mintCurPage).Left = -scrHsc.value * 15# + SHADOW_W * 2
    picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    picPage(mintCurPage).Left = -scrHsc.value * 15# + SHADOW_W * 2
    picShadow.Left = picPage(mintCurPage).Left + SHADOW_W
    Me.Refresh
End Sub

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

Private Sub DrawText(objDraw As Object, ByVal x As Single, ByVal y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0, Optional Rotation As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    
    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        objDraw.FontTransparent = True
        .CurrentX = x
        .CurrentY = y
        If Rotation = 0 Then
            objDraw.Print Text
        Else
            Call DrawRotateText(objDraw, .CurrentX, .CurrentY, Text, Rotation)
        End If
        
        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Sub DrawRotateText(mobjDrawObject As Object, x As Single, y As Single, strText As String, ByVal Degrees As Long)
    '在(X,Y)处旋转输出Text文本
    Dim lngFont As Long
    Dim lngNewFont As Long
    Dim lf As LOGFONT
    Dim hwnd As Long
    Dim hdc As Long
    Dim hOldfont As Long
    Dim hPrintDc As Long
    
    If mobjDrawObject Is Printer Then
        hdc = Printer.hdc
    Else
        hwnd = GetDesktopWindow
        hdc = GetDC(hwnd)
    End If
       
    With lf
          .lfHeight = -(mobjDrawObject.Font.size * GetDeviceCaps(hdc, LOGPIXELSY)) / 72
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
        Call ReleaseDC(hwnd, hdc)
    End If
    
    lngNewFont = CreateFontIndirect(lf)
    
    With mobjDrawObject
        If mobjDrawObject Is Printer Then
                x = x / Printer.TwipsPerPixelX
                y = y / Printer.TwipsPerPixelY
        End If
    End With
    
    hPrintDc = mobjDrawObject.hdc
    lngFont = SelectObject(hPrintDc, lngNewFont)
    If mobjDrawObject Is Printer Then
        Call TextOut(hPrintDc, x, y, strText, LenB(StrConv(strText, vbFromUnicode)))
    Else
        mobjDrawObject.CurrentX = x
        mobjDrawObject.CurrentY = y
        mobjDrawObject.Print strText
    End If
    
    Call SelectObject(hPrintDc, lngFont)
    Call DeleteObject(lngNewFont)
    
End Sub
