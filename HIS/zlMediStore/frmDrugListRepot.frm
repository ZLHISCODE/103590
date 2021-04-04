VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugListRepot 
   BackColor       =   &H8000000C&
   Caption         =   "药品明细表"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame shpback 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "药品"
      Height          =   4065
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   5880
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdData 
         Height          =   2565
         Left            =   495
         TabIndex        =   3
         Top             =   945
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4524
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         AllowUserResizing=   3
         Appearance      =   0
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房:"
         Height          =   180
         Left            =   510
         TabIndex        =   7
         Top             =   735
         Width           =   450
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品明细表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   5
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label lbl期间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期间:   "
         Height          =   180
         Left            =   4620
         TabIndex        =   4
         Top             =   690
         Width           =   720
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   7785
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Object.ToolTipText     =   "重置条件"
               Object.Tag             =   "重置"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "图形"
               Key             =   "图形"
               Object.ToolTipText     =   "图形分析"
               Object.Tag             =   "图形"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "字体"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugListRepot.frx":04C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5085
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7964
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEXCEL 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "条件重置(&J)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "前景色(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "背景色(&B)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpZlWeb 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebSend 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)…"
      End
   End
End
Attribute VB_Name = "frmDrugListRepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public inDeptId As Long            '库房id
Public InDeptName  As String              '库房名称
Public inDrugType As Long          '药品类型id
Public inDrugTypeName As String        '药品类型名称

Dim dtpStartDate As String        '起止日期
Dim dtpEndDate As String        '终止日期
Dim strStartDate As String        '起止日期
Dim strEndDate As String        '终止日期
Dim DataRecordSet As ADODB.Recordset
Dim blnFirst As Boolean              '确定是否第一次使用本系统
Dim Bln西成药 As Boolean '表示是否具有查询西成药的权限
Dim Bln中成药 As Boolean '表示是否具有查询中成药的权限
Dim Bln中草药 As Boolean '表示是否具有查询中草药的权限
Dim Str材质 As String


Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    lbl库房.Caption = "库房：" & InDeptName
    lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
    ReFreshStru
    blnFirst = False
    If Not RefreshData Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    blnFirst = True
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD HH:mm:ss")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD HH:mm:ss")
    RestoreWinState Me
    If InStr(gstrStockSearchPrivs, "西成药") <> 0 Then
        Bln西成药 = True
    Else
        Bln西成药 = False
    End If
    
    If InStr(gstrStockSearchPrivs, "中成药") <> 0 Then
        Bln中成药 = True
    Else
        Bln中成药 = False
    End If
    
    If InStr(gstrStockSearchPrivs, "中草药") <> 0 Then
        Bln中草药 = True
    Else
        Bln中草药 = False
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Dim lngCbrHeight As Long, lngStbHeight As Long
    
    If Me.WindowState = 1 Then Exit Sub
    On Error Resume Next
    
    lngCbrHeight = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    lngStbHeight = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Me.shpback.Left = Me.ScaleLeft + 50
    Me.shpback.Width = Me.ScaleWidth - 100
    Me.shpback.Top = Me.ScaleTop + lngCbrHeight + 50
    Me.shpback.Height = Me.ScaleHeight - (lngCbrHeight + lngStbHeight + 100)
    
    Me.lblTitle.Top = 150
    Me.lblTitle.Left = 0
    Me.lblTitle.Width = Me.shpback.Width
    
    With lbl库房
        .Top = Me.lblTitle.Top + 500
        .Left = 200
    End With
    
    With Me.fgdData
        .Left = 200
        .Width = Me.shpback.Width - 400
        .Top = Me.lbl库房.Top + Me.lbl库房.Height + 45
        .Height = Me.shpback.Height - Me.fgdData.Top - 400
    End With
    Me.lbl期间.Top = Me.lblTitle.Top + 500
    Me.lbl期间.Left = Me.fgdData.Width + Me.fgdData.Left - Me.lbl期间.Width
    
    If Me.shpback.Width < Me.lblTitle.Width Then
        Me.lblTitle.Visible = False
        Me.fgdData.Visible = False
        Me.lbl期间.Visible = False
    Else
        Me.lblTitle.Visible = True
        Me.fgdData.Visible = True
        Me.lbl期间.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
    Unload frmDrugListRepotAsk
End Sub

Private Sub mnuEXCEL_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    objPrint.Title.Text = Me.lblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl库房.Caption
     objRow.Add Me.lbl期间.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objPrint.Body = fgdData
     
      Set objRow = New zlTabAppRow
     With objRow
        .Add "打印人:"
        .Add "打印时间:" & Format(Currentdate, "yyyy年MM月DD日")
     End With
     
     objPrint.BelowAppRows.Add objRow
    
     objPrint.PageFooter = 2
     
     zlPrintOrView1Grd objPrint, 3
     Set objPrint = Nothing

End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    RefreshData
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint True
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint False
End Sub
Private Sub grdPrint(blnIsPreview As Boolean)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：blnIsPreview false表示预览
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = Me.lblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl库房.Caption
     objRow.Add Me.lbl期间.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add "打印人:" & UserInfo.用户姓名
     objRow.Add "打印时间:" & Format(Currentdate, "yyyy年MM月DD日 HH:MM")
     objPrint.BelowAppRows.Add objRow
     Set objPrint.Body = fgdData
     
     objPrint.PageFooter = 2
     
    If Not blnIsPreview Then
             zlPrintOrView1Grd objPrint, 2
        Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    End If
    Set objPrint = Nothing
End Sub
Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub mnuHelpHelp_Click()
     Shell "hh.exe " & App.Path & "\zlMediBill.chm::/药库事务处理/药品库存查询.htm", vbNormalFocus
End Sub
Private Sub mnuHelpWebSend_Click()
    zlMailTo Me.hWnd
End Sub

Private Sub mnuHelpZlWeb_Click()
    zlHomePage Me.hWnd
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True
    
    Select Case Index
    Case 0
        Me.lblTitle.FontSize = 22
        Me.lbl库房.FontSize = 9
        Me.lbl期间.FontSize = 9
        Me.fgdData.Font.Size = 9
        Me.fgdData.FontFixed.Size = 9

     Case 1
        Me.lblTitle.FontSize = 24
        Me.lbl库房.FontSize = 11
        Me.lbl期间.FontSize = 11
        Me.fgdData.Font.Size = 11
        Me.fgdData.FontFixed.Size = 11

    Case 2
        Me.lblTitle.FontSize = 28
        Me.lbl库房.FontSize = 15
        Me.lbl期间.FontSize = 15
        Me.fgdData.Font.Size = 15
        Me.fgdData.FontFixed.Size = 15
    End Select
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.lblTitle.ForeColor)
    Me.lblTitle.ForeColor = lngForeColor
    Me.lbl库房.ForeColor = lngForeColor
    Me.lbl期间.ForeColor = lngForeColor
    Me.fgdData.ForeColor = lngForeColor
    Me.fgdData.ForeColorFixed = lngForeColor
End Sub
Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.fgdData.BackColor)
    Me.shpback.BackColor = lngBackColor
    Me.fgdData.BackColor = lngBackColor
    Me.fgdData.BackColorBkg = lngBackColor
    Me.fgdData.BackColorFixed = lngBackColor
End Sub


Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "预览"
            mnuFilePrintView_Click
        Case "打印"
            grdPrint True
        Case "重置"
            mnuFileOpen_Click
        Case "图形"
            
        Case "字体"
             PopupMenu mnuViewFont
        Case "前景色"
            mnuViewForeColor_Click
        Case "背景色" '
            mnuViewBackColor_Click
        Case "帮助"
            mnuHelpHelp_Click
        Case "退出"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Function RefreshData() As Boolean
    '-------------------------------------------------------------------------
    '--功能：刷新数据
    '--参数:                                                                --
    '--返回:                                                                --
    '-------------------------------------------------------------------------
    Dim strsql As String
    Dim Str单位 As String
    Dim Str系数 As String
    
    Dim lngRow As Long
    Dim i As Long
    Dim frmNewAsk As New frmDrugListRepotAsk
    Dim lng当前金额 As Double
    Dim lng当前数量 As Double
    Dim lng当前差价 As Double
    Dim lng期初金额 As Double
    Dim lng期初差价 As Double
    Dim lng期初数量 As Double
    Dim lng入库数量 As Double
    Dim lng入库金额 As Double
    Dim lng入库差价 As Double
    Dim lng出库金额 As Double
    Dim lng出库差价 As Double
    Dim lng出库数量 As Double
    Dim lng期末金额 As Double
    Dim lng期末差价 As Double
    Dim lng期末数量 As Double
    Dim lng调价金额 As Double
    Dim lng调价差价 As Double
    
    Dim dbl当前金额 As Double
    Dim dbl当前数量 As Double
    Dim dbl当前差价 As Double
    Dim dbl期初数量 As Double
    Dim Dbl期初金额 As Double
    Dim Dbl期初差价 As Double
    Dim dbl入库数量 As Double
    Dim dbl入库金额 As Double
    Dim dbl入库差价 As Double
    Dim dbl出库数量 As Double
    Dim dbl出库金额 As Double
    Dim dbl出库差价 As Double
    Dim dbl期末金额 As Double
    Dim dbl调价金额 As Double
    Dim dbl调价差价 As Double
    Dim dbl期末差价 As Double
    Dim dbl期末数量 As Double
    
    Dim str用途 As String
    
    On Error GoTo errHandle
    RefreshData = False
    Load frmDrugListRepotAsk
    With frmDrugListRepotAsk
        .inDeptId = inDeptId
        .Show 1, Me
        If Not .blnAskOk Then Exit Function
        inDeptId = .cbo库房.ItemData(.cbo库房.ListIndex)
        InDeptName = .cbo库房.Text
        strStartDate = Format(.dtpStartDate.Value, "yyyyMMDDHHmmss")
        strEndDate = Format(.dtpEndDate.Value, "yyyyMMDDHHmmss")
        dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD HH:mm:ss")
        dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD HH:mm:ss")
                
    End With
    
    Str材质 = "''"
    If Bln西成药 Then Str材质 = "'西成药'"
    If Bln中成药 Then
        If Bln西成药 Then
            Str材质 = Str材质 & ",'中成药'"
        Else
            Str材质 = "'中成药'"
        End If
    End If
    If Bln中草药 Then
        If Bln中成药 Or Bln西成药 Then
            Str材质 = Str材质 & ",'中草药'"
        Else
            Str材质 = "'中草药'"
        End If
    End If
    
    Select Case frmDrugQuery.intChoose级数
            Case 1
                Str单位 = "B.售价单位 As 单位,"
                Str系数 = "1"
            Case 2
                Str单位 = "B.门诊单位 As 单位,"
                Str系数 = "B.门诊包装"
            Case 3
                Str单位 = "B.药库单位 As 单位,"
                Str系数 = "B.药库包装"
            Case 4
                Str单位 = "B.住院单位 As 单位,"
                Str系数 = "B.住院包装"
    End Select
    
    
    str用途 = frmDrugQuery.tvwSection_S.SelectedItem.Key
    
'    StrSql = "Select Distinct A.当前数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 当前数量,A.当前金额 As 当前金额,A.当前差价 As 当前差价," & _
             " C.到当前发生数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到当前发生数量,C.到当前发生金额 As 到当前发生金额,C.到当前发生差价 As 到当前发生差价," & _
             " C.到期末入库数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到期末入库数量,C.到期末入库金额 As 到期末入库金额,C.到期末入库差价 As 到期末入库差价," & _
             " C.到期末出库数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到期末出库数量,C.到期末出库金额 As 到期末出库金额,C.到期末出库差价 As 到期末出库差价," & _
              Str单位 & "B.药品id,B.编码 As 编码,X.通用名称 As 名称,B.规格 As 规格" & _
            " From (Select 药品id,Sum(实际数量) As 当前数量,Sum(实际金额) As 当前金额,Sum(实际差价) As 当前差价 From 药品库存 Where 性质=1 " & IIf(inDeptId = 0, "", "And 库房id=" & inDeptId) & "Group by 药品id ) A," & _
            " (Select 药品id,Sum(到当前发生数量) As 到当前发生数量,Sum(到当前发生金额) As 到当前发生金额,Sum(到当前发生差价) As 到当前发生差价," & _
            "        Sum(到期末入库数量) As 到期末入库数量,Sum(到期末入库金额) As 到期末入库金额,Sum(到期末入库差价) As 到期末入库差价, " & _
            "        Sum(到期末出库数量) As 到期末出库数量,Sum(到期末出库金额) As 到期末出库金额,Sum(到期末出库差价) As 到期末出库差价" & _
            "  From (Select E.药品id,Sum(数量) As 到当前发生数量,Sum(E.金额) As 到当前发生金额,Sum(差价) As 到当前发生差价," & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,0,1)*Sum(E.数量)) as 到期末入库数量, " & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,-1,0)*Sum(E.数量)) as 到期末出库数量, " & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,0,1)*Sum(E.金额)) as 到期末入库金额, " & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,-1,0)*Sum(E.金额)) as 到期末出库金额, " & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,0,1)*Sum(E.差价)) as 到期末入库差价, " & _
            "       Decode(Sign(To_number(To_Char(E.日期,'yyyymmdd'))-" & strEndDate & "),-1,0,decode(F.系数,-1,-1,0)*Sum(E.差价)) as 到期末出库差价  " & _
            "        From 药品收发汇总 E,药品入出类别 F " & _
            "        Where To_number(To_Char(E.日期,'yyyymmdd'))>= " & strStartDate & IIf(inDeptId = 0, "", "And 库房id=" & inDeptId) & " And E.类别id=F.id " & _
            "        Group By E.药品id,E.日期,F.系数)" & _
            "  Group By 药品id)C," & _
            " 药品目录 B,药品信息 X" & _
            " Where B.药品id=A.药品id(+) And B.药品id=C.药品id(+) And B.药名id=X.药名id and (B.撤档时间 IS NULL OR TO_CHAR (B.撤档时间, 'yyyy-MM-dd') = '3000-01-01') " _
            & IIf(Left(str用途, 1) = "R", " and x.材质分类 In ('" & Mid(str用途, 2) & "')", " And x.用途分类id in ( Select id From 药品用途分类 Q start with Q.id= " & Mid(str用途, 2) & " connect by prior id=上级id)") _
            & " Order By B.药品id "
    
    strsql = "Select Distinct A.当前数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 当前数量,A.当前金额 As 当前金额,A.当前差价 As 当前差价," & _
             " C.到当前发生数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到当前发生数量,C.到当前发生金额 As 到当前发生金额,C.到当前发生差价 As 到当前发生差价," & _
             " C.到期末入库数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到期末入库数量,C.到期末入库金额 As 到期末入库金额,C.到期末入库差价 As 到期末入库差价," & _
             " C.到期末出库数量/Decode(" & Str系数 & ",0,1," & Str系数 & ") As 到期末出库数量,C.到期末出库金额 As 到期末出库金额,C.到期末出库差价 As 到期末出库差价," & _
              Str单位 & "B.药品id,B.编码 As 编码,X.通用名称 As 名称,B.规格 As 规格" & _
            " From (Select 药品id,Sum(实际数量) As 当前数量,Sum(实际金额) As 当前金额,Sum(实际差价) As 当前差价 From 药品库存 Where 性质=1 " & IIf(inDeptId = 0, "", "And 库房id=" & inDeptId) & "Group by 药品id ) A," & _
            " (SELECT 药品id," _
                & "(sum(DECODE(入出系数,-1,0,1)*实际数量)- sum(DECODE(入出系数,-1,1,0)*实际数量)) as 到当前发生数量," _
                & "(sum(DECODE(入出系数,-1,0,1)*零售金额)-sum(DECODE(入出系数,-1,1,0)*零售金额)) as 到当前发生金额," _
                & "(SUM(DECODE(入出系数,-1,0,1)*差价)-SUM(DECODE(入出系数,-1,1,0)*差价)) as  到当前发生差价," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数,-1,0,1)*实际数量)) AS 到期末入库数量," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数,-1,0,1)*零售金额)) AS 到期末入库金额," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数,-1,0,1)*差价)) AS 到期末入库差价," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数,-1,1,0)*实际数量)) AS 到期末出库数量," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数,-1,1,0)*零售金额)) AS 到期末出库金额," _
                & "sum(Decode(Sign(To_number(To_Char(审核日期,'yyyymmddhh24miss'))-" & strEndDate & "),1,0,DECODE(入出系数, -1, 1, 0) * 差价)) As 到期末出库差价 " _
                & " From 药品收发记录 " _
               & " WHERE 审核日期>=to_date('" & strStartDate & "','yyyy-mm-dd hh24:mi:ss') " _
               & IIf(inDeptId = 0, "", "And 库房id=" & inDeptId) & _
            "  Group By 药品id)C," & _
            " 药品目录 B,药品信息 X" & _
            " Where B.药品id=A.药品id(+) And B.药品id=C.药品id(+) And B.药名id=X.药名id and (B.撤档时间 IS NULL OR TO_CHAR (B.撤档时间, 'yyyy-MM-dd') = '3000-01-01') " _
            & IIf(Left(str用途, 1) = "R", " and x.材质分类 In ('" & Mid(str用途, 2) & "')", " And x.用途分类id in ( Select id From 药品用途分类 Q start with Q.id= " & Mid(str用途, 2) & " connect by prior id=上级id)") _
            & " Order By B.药品id "
    Set DataRecordSet = New ADODB.Recordset
    ShowFlash "正在装入数据，请稍候…", Me
    DoEvents
   
    With DataRecordSet
        Call SQLTest(App.Title, Me.Caption, strsql)
        Set DataRecordSet = zldatabase.OpenSQLRecord(strsql, "RefreshData")
        Call SQLTest
        If .RecordCount = 0 Then
            StopFlash
            MsgBox "在此条件和权限范围中,无任何明细表记录！", vbInformation, gstrSysName
            Exit Function
        End If
        lng期初金额 = 0: lng期初差价 = 0: lng入库金额 = 0: lng入库差价 = 0: lng出库金额 = 0: lng出库差价 = 0: lng期末金额 = 0:        lng期末差价 = 0
        lng调价金额 = 0: lng调价差价 = 0: lng期初数量 = 0
        Dbl期初金额 = 0: Dbl期初差价 = 0: dbl入库金额 = 0: dbl入库差价 = 0: dbl出库金额 = 0: dbl出库差价 = 0: dbl期末金额 = 0:        dbl期末差价 = 0
        dbl调价金额 = 0: dbl调价差价 = 0: dbl期初数量 = 0
        
        fgdData.rows = IIf(.RecordCount = 0, 1, .RecordCount) + 2
        Call RefreshGridColWidth(Me.fgdData, 0)
         i = 2
        Do While Not .EOF
            
            Dbl期初金额 = IIf(IsNull(.Fields("当前金额").Value), 0, .Fields("当前金额").Value) - IIf(IsNull(.Fields("到当前发生金额").Value), 0, .Fields("到当前发生金额").Value)
            Dbl期初差价 = IIf(IsNull(.Fields("当前差价").Value), 0, .Fields("当前差价").Value) - IIf(IsNull(.Fields("到当前发生差价").Value), 0, .Fields("到当前发生差价").Value)
            dbl期初数量 = IIf(IsNull(.Fields("当前数量").Value), 0, .Fields("当前数量").Value) - IIf(IsNull(.Fields("到当前发生数量").Value), 0, .Fields("到当前发生数量").Value)
            
            dbl入库金额 = IIf(IsNull(.Fields("到期末入库金额").Value), 0, .Fields("到期末入库金额").Value)
            dbl入库差价 = IIf(IsNull(.Fields("到期末入库差价").Value), 0, .Fields("到期末入库差价").Value)
            dbl入库数量 = IIf(IsNull(.Fields("到期末入库数量").Value), 0, .Fields("到期末入库数量").Value)
            dbl出库金额 = IIf(IsNull(.Fields("到期末出库金额").Value), 0, .Fields("到期末出库金额").Value)
            dbl出库差价 = IIf(IsNull(.Fields("到期末出库差价").Value), 0, .Fields("到期末出库差价").Value)
            dbl出库数量 = IIf(IsNull(.Fields("到期末出库数量").Value), 0, .Fields("到期末出库数量").Value)
                        
            dbl期末金额 = Dbl期初金额 + dbl入库金额 - dbl出库金额
            dbl期末差价 = Dbl期初差价 + dbl入库差价 - dbl出库差价
            dbl期末数量 = dbl期初数量 + dbl入库数量 - dbl出库数量
            
            
            lng期初金额 = lng期初金额 + Dbl期初金额
            lng期初差价 = lng期初差价 + Dbl期初差价
            lng入库金额 = lng入库金额 + dbl入库金额
            lng入库差价 = lng入库差价 + dbl入库差价
            lng出库金额 = lng出库金额 + dbl出库金额
            lng出库差价 = lng出库差价 + dbl出库差价
            lng期末金额 = lng期末金额 + dbl期末金额
            lng期末差价 = lng期末差价 + dbl期末差价
            
            fgdData.TextMatrix(i, 0) = IIf(IsNull(.Fields("编码").Value), "", .Fields("编码").Value)
            fgdData.TextMatrix(i, 1) = IIf(IsNull(.Fields("名称").Value), "", .Fields("名称").Value)
            fgdData.TextMatrix(i, 2) = IIf(IsNull(.Fields("规格").Value), "", .Fields("规格").Value)
            fgdData.TextMatrix(i, 3) = IIf(IsNull(.Fields("单位").Value), "", .Fields("单位").Value)
            
            fgdData.TextMatrix(i, 4) = " " & Format(dbl期初数量, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 5) = " " & Format(Dbl期初金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 6) = Format(Dbl期初差价, "##,###0.00;-##,###0.00; ; ")
            
            fgdData.TextMatrix(i, 7) = " " & Format(dbl入库数量, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 8) = " " & Format(dbl入库金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 9) = Format(dbl入库差价, "##,###0.00;-##,###0.00; ; ")
            
            fgdData.TextMatrix(i, 10) = " " & Format(dbl出库数量, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 11) = " " & Format(dbl出库金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 12) = Format(dbl出库差价, "##,###0.00;-##,###0.00; ; ")
'            fgdData.TextMatrix(i, 13) = " "
'            fgdData.TextMatrix(i, 14) = " "
            fgdData.TextMatrix(i, 13) = " " & Format(dbl期末数量, "##,###0.000;-##,###0.000; ; ")
            fgdData.TextMatrix(i, 14) = " " & Format(dbl期末金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 15) = Format(dbl期末差价, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, i)
            .MoveNext
            i = i + 1
        Loop
        If lng期初金额 <> 0 Or lng期初差价 <> 0 Or lng入库金额 <> 0 Or lng入库差价 <> 0 Or lng出库金额 <> 0 Or lng出库差价 <> 0 Or _
            lng期末金额 <> 0 Or lng期末差价 <> 0 Then
            fgdData.rows = fgdData.rows + 1
            fgdData.MergeRow(i) = True
            fgdData.TextMatrix(i, 0) = "合计"
            fgdData.TextMatrix(i, 1) = "合计"
            fgdData.TextMatrix(i, 2) = "合计"
            fgdData.TextMatrix(i, 3) = "合计"
            fgdData.TextMatrix(i, 4) = ""
            fgdData.TextMatrix(i, 5) = " " & Format(lng期初金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 6) = Format(lng期初差价, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 7) = "   "
            fgdData.TextMatrix(i, 8) = Format(lng入库金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 9) = "  " & Format(lng入库差价, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 10) = ""
            fgdData.TextMatrix(i, 11) = "  " & Format(lng出库金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 12) = Format(lng出库差价, "##,###0.00;-##,###0.00; ; ")
'            fgdData.TextMatrix(i, 13) = " "
'            fgdData.TextMatrix(i, 14) = " "
            fgdData.TextMatrix(i, 13) = "  "
            fgdData.TextMatrix(i, 14) = Format(lng期末金额, "##,###0.00;-##,###0.00; ; ")
            fgdData.TextMatrix(i, 15) = "  " & Format(lng期末差价, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, i)
        End If
        fgdData.Redraw = True
    End With
    lbl库房.Caption = "库房：" & InDeptName & Space(6) & "药品用途:" & inDrugTypeName
    lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
    StopFlash
    RefreshData = True
Exit Function
Err:
    StopFlash
    RefreshData = False
    Me.fgdData.Redraw = True
    MsgBox "在获取药品明细表时,出现了不可预知的错误!", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReFreshStru()
    '-------------------------------------------------------------------------
    '--功能:重新获的表头结构
    '--参数:
    '--返回:
    '-------------------------------------------------------------------------
    Dim IntCol As Long
    Me.Caption = "药品明细表"
    Me.lblTitle.Caption = GetUnitName & "药品明细表"

     With fgdData
            .Cols = 16
            .Redraw = False
            .rows = 6
            .FixedRows = 2
            .FixedCols = 0
            .MergeCells = flexMergeRestrictRows
            For IntCol = 0 To .Cols - 1
                .ColAlignmentFixed(IntCol) = 4
                If IntCol <= 3 Then
                    .ColAlignment(IntCol) = 1
                Else
                    .ColAlignment(IntCol) = 7
                End If
                If IntCol <= 3 Then
                    .ColWidth(IntCol) = IIf(IntCol <> 1, IIf(IntCol = 2, 1200, IIf(IntCol = 0, 600, 400)), 1400)
                Else
                    .ColWidth(IntCol) = 1000
                End If
            Next
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(2) = True
            .MergeCol(3) = True
            
            .TextMatrix(0, 0) = "编码"
            .TextMatrix(1, 0) = "编码"
            .TextMatrix(0, 1) = "名称"
            .TextMatrix(1, 1) = "名称"
            .TextMatrix(0, 2) = "规格"
            .TextMatrix(1, 2) = "规格"
            
            Select Case frmDrugQuery.intChoose级数
                Case 1
                    .TextMatrix(0, 3) = "售价单位"
                    .TextMatrix(1, 3) = "售价单位"
                Case 2
                    .TextMatrix(0, 3) = "门诊单位"
                    .TextMatrix(1, 3) = "门诊单位"
                Case 3
                    .TextMatrix(0, 3) = "库房单位"
                    .TextMatrix(1, 3) = "库房单位"
                Case 4
                    .TextMatrix(0, 3) = "住院单位"
                    .TextMatrix(1, 3) = "住院单位"
            End Select

            
            .TextMatrix(0, 4) = "期初"
            .TextMatrix(0, 5) = "期初"
            .TextMatrix(0, 6) = "期初"
            .TextMatrix(1, 4) = "数量"
            .TextMatrix(1, 5) = "金额"
            .TextMatrix(1, 6) = "差价"
            
            .TextMatrix(0, 7) = "本期入库"
            .TextMatrix(0, 8) = "本期入库"
            .TextMatrix(0, 9) = "本期入库"
            .TextMatrix(1, 7) = "数量"
            .TextMatrix(1, 8) = "金额"
            .TextMatrix(1, 9) = "差价"
            
            .TextMatrix(0, 10) = "本期出库"
            .TextMatrix(0, 11) = "本期出库"
            .TextMatrix(0, 12) = "本期出库"
            .TextMatrix(1, 10) = "数量"
            .TextMatrix(1, 11) = "金额"
            .TextMatrix(1, 12) = "差价"
            
            
            .TextMatrix(0, 13) = "期末"
            .TextMatrix(0, 14) = "期末"
            .TextMatrix(0, 15) = "期末"
            .TextMatrix(1, 13) = "数量"
            .TextMatrix(1, 14) = "金额"
            .TextMatrix(1, 15) = "差价"
             Call RefreshGridColWidth(Me.fgdData, 0)
            .Redraw = True
        End With
End Sub
