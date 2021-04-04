VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form vsPrint 
   Caption         =   "打印"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "vsPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   10680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPageSetup 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   1620
      TabIndex        =   2
      Top             =   105
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   285
      TabIndex        =   1
      Top             =   105
      Width           =   1100
   End
   Begin VSPrinter8LibCtl.VSPrinter vp 
      Height          =   5955
      Left            =   195
      TabIndex        =   0
      Top             =   510
      Width           =   7020
      _cx             =   12382
      _cy             =   10504
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   30
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   1
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "整页(&P)|页宽(&W)|双面(&T)|所有(&n)"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
End
Attribute VB_Name = "vsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngIndex As Long
Private mlnghWnd As Long

Public Sub vsPrint(ByVal hWnd As Long, ByVal lngIndex As Long)
    Dim i As Long
    mlngIndex = lngIndex
    mlnghWnd = hWnd
    Me.Show
    Call ReadSetup
    '打印
    Call LoadPrintDoc
     
End Sub

Private Sub LoadPrintDoc()
    Dim strValue As String, iPage As Integer
    vp.Clear
    vp.StartDoc
    vp.RenderControl = mlnghWnd
    vp.EndDoc
    iPage = vp.PageCount
    vp.Clear
    
    '页脚
    vp.HdrFontName = "宋体"
    vp.HdrFontSize = 9
    strValue = ReadIni("Report" & mlngIndex, "打印页码", App.Path & "\PrintSetup.ini")
    If Val(strValue) = 1 Then
        vp.Footer = "第%d页共" & iPage & "页||"
    ElseIf Val(strValue) = 2 Then
        vp.Footer = "|第%d页共" & iPage & "页|"
    ElseIf Val(strValue) = 3 Then
        vp.Footer = "||第%d页共" & iPage & "页"
    End If
'
    vp.StartDoc
    vp.RenderControl = mlnghWnd
    vp.EndDoc

    
End Sub

Private Sub cmdPageSetup_Click()
    If vp.PrintDialog(pdPageSetup) Then
   
        ShowCaption
        Call LoadPrintDoc
    End If
End Sub

Private Sub cmdPrint_Click()
    If vp.PageCount > 0 Then vp.PrintDoc
End Sub

Private Sub Form_Resize()
    With Me.vp
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - Me.vp.Top
    End With
    
End Sub


Private Sub ReadSetup()

    Dim strValue As String, i As Integer
    '打印机
    strValue = ReadIni("Report" & mlngIndex, "打印机", App.Path & "\PrintSetup.ini")
    If strValue <> "" Then
        vp.Device = strValue
    End If
   
    '方向
    strValue = ReadIni("Report" & mlngIndex, "打印方向", App.Path & "\PrintSetup.ini")
    If Val(strValue) = 0 Then
        vp.Orientation = orPortrait
    Else
        vp.Orientation = orLandscape
    End If
    
    '边距
    strValue = ReadIni("Report" & mlngIndex, "上边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        vp.MarginTop = Format(Val(strValue), "0.0") & "mm"
    Else
        vp.MarginTop = "25.4mm"
    End If
    strValue = ReadIni("Report" & mlngIndex, "下边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        vp.MarginBottom = Format(Val(strValue), "0.0") & "mm"
    Else
        vp.MarginBottom = "25.4mm"
    End If
    strValue = ReadIni("Report" & mlngIndex, "左边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        vp.MarginLeft = Format(Val(strValue), "0.0") & "mm"
    Else
        vp.MarginLeft = "25.4mm"
    End If
    strValue = ReadIni("Report" & mlngIndex, "右边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        vp.MarginRight = Format(Val(strValue), "0.0") & "mm"
    Else
        vp.MarginRight = "25.4mm"
    End If
    
    '纸张大小
    strValue = ReadIni("Report" & mlngIndex, "纸张大小", App.Path & "\PrintSetup.ini")
    
    '纸张宽高
    If Val(strValue) <> 256 Then
        If Val(strValue) > 0 And Val(strValue) < 256 Then vp.PaperSize = Val(strValue)
    ElseIf Val(strValue) = 256 Then
        vp.PaperSize = pprUser
        strValue = ReadIni("Report" & mlngIndex, "纸张宽度", App.Path & "\PrintSetup.ini")
        vp.PaperWidth = Val(strValue)
        strValue = ReadIni("Report" & mlngIndex, "纸张高度", App.Path & "\PrintSetup.ini")
        vp.PaperHeight = Val(strValue)
    End If
     
    Call ShowCaption
End Sub

Private Sub ShowCaption()
    Me.Caption = vp.Device
    Me.Caption = Me.Caption & " (" & Format(vp.PageWidth / (1440 / 25.4), "0.00") & "×" & Format(vp.PageHeight / (1440 / 25.4), "0.00") & ") "
    If vp.Orientation = orPortrait Then
         Me.Caption = Me.Caption & " 纵向"
    Else
        Me.Caption = Me.Caption & " 横向"
    End If
End Sub


