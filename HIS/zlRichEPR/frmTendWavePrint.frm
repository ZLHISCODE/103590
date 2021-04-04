VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendWavePrint 
   Caption         =   "体温表预览"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9030
   Icon            =   "frmTendWavePrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9030
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmTendWavePrint.frx":5C12
      Height          =   4755
      LargeChange     =   20
      Left            =   8775
      Max             =   100
      MouseIcon       =   "frmTendWavePrint.frx":5F1C
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   250
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmTendWavePrint.frx":606E
      Height          =   250
      LargeChange     =   20
      Left            =   0
      Max             =   100
      MouseIcon       =   "frmTendWavePrint.frx":6378
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5475
      Width           =   8760
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   -180
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   8760
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   705
         MouseIcon       =   "frmTendWavePrint.frx":64CA
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   5
         Top             =   195
         Width           =   6990
         Begin VSFlex8Ctl.VSFlexGrid VsfData 
            Height          =   2610
            Left            =   2040
            TabIndex        =   6
            Top             =   0
            Width           =   4305
            _cx             =   7594
            _cy             =   4604
            Appearance      =   0
            BorderStyle     =   0
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483634
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   3
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   5000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendWavePrint.frx":661C
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   0   'False
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   4
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   810
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   6990
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendWavePrint.frx":667E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
            Object.ToolTipText     =   "打印机信息"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTendWavePrint.frx":6F12
   End
End
Attribute VB_Name = "frmTendWavePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngWidth As Long, mlngHeight As Long
Private mlngPreX As Long, mlngPreY As Long
Private Const Shadow_W = 45 '阴影厚度
Private Const mintCurveNullRow As Integer = 2

Private mintPrintRange As Integer
Private mlngBeginY As Long
Private mlngLeft As Long
Private mstrPrinter As String
Private msngScale As Single
Private mstrArrFromTo() As String
Private strSQL As String
Private mlngCaseRecordID As Long
Private mobjCombo As CommandBarComboBox
Private mobjPage As CommandBarComboBox
Private mintPage As Integer
Private Type Type_Printer
    intPage As Integer
    lngWidth As Long
    lngHeight As Long
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    intOrient As Integer
    intBin As Integer
End Type
Private gPrinter As Type_Printer

Public Event AfterPrint()

Public Function Preview(ByVal rsTemp As ADODB.Recordset, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal lngLeft As Long) As Boolean
    Dim lngTop As Long
    
    If DrawWaveStyle(picShow, rsTemp, False, lngTop) Then
        picShow.Height = Me.ScaleX(lngHeight, vbMillimeters, vbTwips)
        picShow.Width = Me.ScaleX(lngWidth, vbMillimeters, vbTwips)
        Call ShowTabBaby(rsTemp, lngLeft, lngTop)
        Me.Show 1
    End If
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)")
    With cbrMenuBar.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_Print, "打印(&P)..."
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '------------------------------------------------------------------------------------------------------------------
    Set cbrToolBar = cbsThis.Add("打印预览", xtpBarTop)
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.STYLE = xtpButtonIconAndCaption
    
        Set cbrControl = .Add(xtpControlButton, 1, "实际大小")
        cbrControl.BeginGroup = True
        cbrControl.Parameter = "1.00"
        
        Set mobjCombo = .Add(xtpControlComboBox, 3, "")
        mobjCombo.BeginGroup = True
        mobjCombo.AddItem "100%", 1
        mobjCombo.AddItem "75%", 2
        mobjCombo.AddItem "50%", 3
        mobjCombo.AddItem "30%", 4
        mobjCombo.AddItem "25%", 5
        mobjCombo.ListIndex = 1
        mobjCombo.Width = 80
        mobjCombo.DropDownWidth = 80
        mobjCombo.DropDownListStyle = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigatebeginning, "第一页"): cbrControl.BeginGroup = True: cbrControl.STYLE = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateleft, "前一页"): cbrControl.STYLE = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateright, "后一页"): cbrControl.STYLE = xtpButtonIcon
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Navigateend, "最后一页"): cbrControl.STYLE = xtpButtonIcon
                
        Set mobjPage = .Add(xtpControlComboBox, 6, "")
        mobjPage.BeginGroup = True
        
        mobjPage.ListIndex = 1
        mobjPage.Width = 80
        mobjPage.DropDownWidth = 80
        mobjPage.DropDownListStyle = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&E)"): cbrControl.STYLE = xtpButtonIconAndCaption
        cbrControl.BeginGroup = True
    End With
    
    
    '快键绑定
    '------------------------------------------------------------------------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    
    Select Case Control.ID
        
    Case 1
        
        msngScale = 1
        mobjCombo.Text = "100%"
        
    Case 2
        
        msngScale = Val(Control.Parameter)
        mobjCombo.Text = CStr(msngScale * 100) & "%"
    Case 3
        
        If Val(Control.Text) > 0 Then
            If Val(Control.Text) <= 200 Then
                msngScale = Val(Control.Text) / 100
                Control.Text = CStr(Val(Control.Text)) & "%"
            Else
                If Control.ListIndex = 0 Then
                    Control.ListIndex = 1
                End If
                
                msngScale = Val(Control.List(Control.ListIndex)) / 100
                Control.Text = Control.List(Control.ListIndex)
            End If
        Else
            If Control.ListIndex = 0 Then
                Control.ListIndex = 5
            End If
                
            msngScale = Val(Control.List(Control.ListIndex)) / 100
            Control.Text = Control.List(Control.ListIndex)
        End If
        
        
        DoEvents
        Control.SetFocus
        
    Case 6
        
        If Val(Control.Text) > 0 Then
            If Val(Control.Text) <= mobjPage.ListCount Then
                Control.ListIndex = Val(Control.Text)
            End If
        Else
        End If
        
        Control.Text = Control.List(Control.ListIndex)
        
        
        DoEvents
        Control.SetFocus
        
    Case conMenu_View_ToolBar_Button
    
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
        
    Case conMenu_View_StatusBar
    
        sta.Visible = Not sta.Visible
        cbsThis.RecalcLayout
            
    Case conMenu_Help_Help
    
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    
    Case conMenu_Help_About
        
        Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
        
    Case conMenu_Help_Web_Home
        
        Call zlHomePage(Me.hWnd)
        
    Case conMenu_Help_Web_Forum '中联论坛
    
        Call zlWebForum(Me.hWnd)
        
    Case conMenu_Help_Web_Mail
        
        Call zlMailTo(Me.hWnd)
    
    Case conMenu_File_Exit
        Unload Me
        Exit Sub
            
    End Select
    
End Sub


Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub cbsThis_Resize()
    
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    picBack.Move lngLeft, lngTop, lngRight - lngLeft - scrVsc.Width, lngBottom - lngTop - scrHsc.Height
    scrVsc.Move picBack.Left + picBack.Width, lngTop, scrVsc.Width, picBack.Height
    scrHsc.Move lngLeft, picBack.Top + picBack.Height, picBack.Width, scrHsc.Height

    picShadow.Width = picShow.Width
    picShadow.Height = picShow.Height
    
    '调整预览页

    If picBack.ScaleWidth >= picShow.Width + Shadow_W * 4 Then
        picShow.Left = (picBack.ScaleWidth - (picShow.Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picShow.Left + Shadow_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picShow.Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.Value = 0
        scrHsc.Enabled = True
        scrhsc_Change
    End If

    If picBack.ScaleHeight >= picShow.Height + Shadow_W * 4 Then
        picShow.Top = (picBack.ScaleHeight - (picShow.Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picShow.Top + Shadow_W
        scrVsc.Enabled = False
    Else
        scrVsc.Max = (picShow.Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
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

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 1, 2
        Control.Checked = (Val(Control.Parameter) = Val(msngScale))
        
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
                If Shift = 2 Then
                    scrVsc.Value = IIf(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIf(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
                End If
            End If
        Case vbKeyDown
            If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
                If Shift = 2 Then
                    scrVsc.Value = IIf(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
                Else
                    scrVsc.Value = IIf(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
                End If
            End If
        Case vbKeyLeft
            If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
                If Shift = 2 Then
                    scrHsc.Value = IIf(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIf(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
                End If
            End If
        Case vbKeyRight
            If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
                If Shift = 2 Then
                    scrHsc.Value = IIf(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
                Else
                    scrHsc.Value = IIf(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
                End If
            End If
            
    End Select
End Sub

Private Function PrintPage()
'功能：打印体温表
    Dim i As Long
    Dim intCol As Integer
    Dim strParam As String
    
    If Not ExistsPrinter Then MsgBox "系统中没有可用的打印机。", vbInformation: Exit Function
    If MsgBox("准备打印体温单，打印机是否已经准备就绪？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
    If Printer.DeviceName <> mstrPrinter Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = mstrPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
        
    On Error Resume Next
    
    '纸张
    If mintPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = mlngWidth
        Printer.Height = mlngHeight
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    On Error GoTo ErrPrintState
        
    '调用打印

    On Error Resume Next
    If mintPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = mlngWidth
        Printer.Height = mlngHeight
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    '打印机还原
    If IsWindowsNT And mintPage = 256 Then
         Call SetNTPrinterPaper(Me.hWnd, mlngWidth / conRatemmToTwip, mlngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    End If
        
        
    Call ShowPrinterInfo
    
    RaiseEvent AfterPrint
    Exit Function
ErrPrintState:
    MsgBox "打印机初始化失败！", vbExclamation, gstrSysName
End Function


Private Sub picback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.ID > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.IconId = cbrControl.IconId
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (Y - mlngPreY) / 15 > 0 Then
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - mlngPreY) / 15)
            Else
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - mlngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - mlngPreX) / 15 > 0 Then
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - mlngPreX) / 15)
            Else
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - mlngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button = 1 Then Set picShow.MouseIcon = scrHsc.MouseIcon
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
        
    Set cbrPopupBar = cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.ID > 0 Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.IconId = cbrControl.IconId
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    cbrPopupBar.ShowPopup
    
End Sub



Private Sub ShowPrinterInfo()
    sta.Panels(2).Text = "打印机:" & Printer.DeviceName & "/纸张:" & _
        GetPaperName(Printer.PaperSize) & "/尺寸:" & _
        CLng(Printer.Width / conRatemmToTwip) & "×" & CLng(Printer.Height / conRatemmToTwip) & "/纸向:" & _
        IIf(Printer.Orientation = 1, "纵向", "横向")
End Sub

Private Sub scrVsc_Change()
    picShow.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picShow.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    picShow.Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picShow.Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Change()
    picShow.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picShow.Left + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    picShow.Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picShow.Left + Shadow_W
    Me.Refresh
End Sub


Private Sub Form_Load()

    RestoreWinState Me, App.ProductName
    
    mlngLeft = gPrinter.lngLeft
    mlngWidth = gPrinter.lngWidth
    mlngHeight = gPrinter.lngHeight
    mintPage = gPrinter.intPage
    
    Call ShowPrinterInfo
    
End Sub


Private Function ShowTabBaby(ByVal rsTmp As ADODB.Recordset, ByVal lngTabLeft As Long, ByVal lngHeight As Long)
    Dim lngCurveRows As Long
    Dim lngMaxValue As Long, lngMinValue As Long
    Dim lngTotal As Long, lngCurveNull As Long
    Dim lngLeft As Long, lngCurveRowHeight As Long
    Dim lngTabBabyRowHeight As Long
    Dim lngRow As Long, lngDay As Long
    Dim lngId  As Long, lngTabBabyTitleID As Long, lngTabBabyNameID As Long
    Dim strSQL  As String
    Dim strBabyTitle As String, strTitleBabyFont As String
    Dim intTitleBabyTitleNum As Integer, i As Integer
    Dim BlnBaby As Boolean
    Dim objFont As StdFont
    
    Dim rsCurve As New ADODB.Recordset
    
    rsTmp.Filter = "父ID=NULL And 对象序号=1 And 内容文本='格式定义'"
    If rsTmp.RecordCount > 0 Then
        lngId = rsTmp!ID
        rsTmp.Filter = "父ID=" & lngId
        Do While Not rsTmp.EOF
            Select Case "" & rsTmp!要素名称
            Case "天数"
                lngDay = Val("" & rsTmp!内容文本)
            Case "婴儿表格左边距"
                lngLeft = Val("" & rsTmp!内容文本)
            Case "婴儿标题文本"
                strBabyTitle = "" & rsTmp!内容文本
            Case "婴儿标题字体"
                strTitleBabyFont = "" & rsTmp!内容文本
            Case "婴儿表格高度"
                lngTabBabyRowHeight = Val("" & rsTmp!内容文本)
            Case "表头层数"
                intTitleBabyTitleNum = Val("" & rsTmp!内容文本)
            Case "婴儿体温单"
                BlnBaby = Val("" & rsTmp!内容文本)
            Case "总列数"
                VsfData.Cols = Val("" & rsTmp!内容文本)
            End Select
            rsTmp.MoveNext
        Loop
    End If
    If Not BlnBaby Then VsfData.Visible = False: Exit Function
    
    rsTmp.Filter = "父ID=NULL And 对象序号=4 And 内容文本='婴儿体温单表头项目'"
    Do While Not rsTmp.EOF
        lngTabBabyTitleID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = "父ID=NULL And 对象序号=3 And 内容文本='表格项目定义'"
    Do While Not rsTmp.EOF
        lngTabBabyNameID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    
    
    With VsfData
        .Top = Me.ScaleX(lngHeight, vbPixels, vbTwips) + 20
        .Left = mlngWaveLeft + Me.ScaleX(lngLeft, vbMillimeters, vbTwips)
        .Width = picShow.Width
        .Height = picShow.Height
        .Rows = .FixedRows + lngDay + 1
        
        Select Case intTitleBabyTitleNum
            Case 1
                .RowHidden(2) = True
                .RowHidden(3) = True
            Case 2
                .RowHidden(3) = True
        End Select
        
        rsTmp.Filter = "父ID= " & lngTabBabyNameID
        rsTmp.Sort = "对象序号"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .ColWidth(Val(rsTmp!对象序号) - 1) = Split(rsTmp!对象属性, "`")(0)
                rsTmp.MoveNext
            Loop
        End If
        rsTmp.Filter = "父ID= " & lngTabBabyTitleID
        rsTmp.Sort = "对象序号"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .TextMatrix((Val(rsTmp!内容行次)), Val(rsTmp!对象序号) - 1) = NVL(rsTmp!内容文本)
                rsTmp.MoveNext
            Loop
        End If
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = strBabyTitle
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .CellBorderRange 1, 0, .Rows - 1, .Cols - 1, vbBlack, 1, 1, 1, 1, 1, 1
        .MergeCellsFixed = flexMergeFree
        .MergeCol(-1) = True
        .MergeRow(-1) = True
        
        Set objFont = New StdFont
        With objFont
            .Name = Split(strTitleBabyFont, ",")(0)
            .Size = Val(Split(strTitleBabyFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strTitleBabyFont, "粗") > 0 Then .Bold = True
            If InStr(1, strTitleBabyFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 0, .FixedCols, 0, .Cols - 1) = objFont
        .ROWHEIGHT(0) = objFont.Size * 20 + 150
        For i = 4 To .Rows - 1
        .ROWHEIGHT(i) = lngTabBabyRowHeight
        VsfData.Redraw = True
        Next
        
    End With
    
End Function

