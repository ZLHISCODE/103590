VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugList 
   BackColor       =   &H8000000C&
   Caption         =   "药品明细帐"
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
      Left            =   540
      TabIndex        =   2
      Top             =   885
      Width           =   5880
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdData 
         Height          =   2565
         Left            =   495
         TabIndex        =   3
         Top             =   1170
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4524
         _Version        =   393216
         BackColor       =   16777215
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lbl单位 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "零售单位:"
         Height          =   180
         Left            =   3105
         TabIndex        =   10
         Top             =   945
         Width           =   2520
      End
      Begin VB.Label lbl规格 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "规格:"
         Height          =   180
         Left            =   2430
         TabIndex        =   9
         Top             =   930
         Width           =   1800
      End
      Begin VB.Label lbl药品 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品:"
         Height          =   180
         Left            =   510
         TabIndex        =   8
         Top             =   945
         Width           =   1365
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
         Caption         =   "药品明细帐"
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
            Picture         =   "frmDrugList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0234
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
            Picture         =   "frmDrugList.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugList.frx":04C6
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
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReFresh 
         Caption         =   "单据(&V) "
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
Attribute VB_Name = "frmDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public InDrugId As Long            '药品id
Public inDeptId As Long            '库房id
Public InDeptName  As String              '库房名称
Public InDrugName  As String       '药品名称
Public InDrugStAndard As String      '药品规格
Public InDrugUnit As String          '药品单位

Dim dtpStartDate As String        '起止日期
Dim dtpEndDate As String        '终止日期
Dim DataRecordSet As ADODB.Recordset
Dim RecTmpList As ADODB.Recordset
Dim blnFirst As Boolean              '确定是否第一次使用本系统

Private mlngLevel As Integer        '单位级数：1:兽价单位;2:门诊单位；3：库房单位； 4：住院单位



Private Sub fgdData_DblClick()
    If Me.fgdData.RowData(fgdData.Row) = 0 Then Exit Sub
    If Me.fgdData.TextMatrix(fgdData.Row, 1) = "" Then Exit Sub
        
    Dim rsTemp As New ADODB.Recordset
    Dim strsql As String
    Dim int记录状态 As Integer
    
    On Error GoTo errHandle
    With rsTemp
        strsql = "Select id,单据,NO,nvl(价格id,0) as 价格id From 药品收发记录 Where id=[1]"
        If .State = adStateOpen Then .Close
        Set rsTemp = zldatabase.OpenSQLRecord(strsql, "fgdData_DblClick", Me.fgdData.RowData(fgdData.Row))
        If .EOF Or .BOF Then Exit Sub
   '1-外购入库；2-自制入库；3-协药入库；4-其他入库；5-差价调整；6-库房移出；7-部门领用；8-收费处方；9-记帐单处方；10-记帐表处方；11-其他出库；12-盘点；13-调价变动
        int记录状态 = Me.fgdData.TextMatrix(fgdData.Row, 13)
        Select Case !单据
        Case 1
            frmPurchaseCard.ShowCard Me, !No, 4, int记录状态
        Case 2
            frmSelfMakeCard.ShowCard Me, !No, 4, int记录状态
        Case 3
            frmAccordDrugCard.ShowCard Me, !No, 4, int记录状态
        Case 4
            frmOtherInputCard.ShowCard Me, !No, 4, int记录状态
        Case 5
            frmDiffPriceAdjustCard.ShowCard Me, !No, 4, int记录状态
        Case 6
            frmTransferCard.ShowCard Me, !No, 4, int记录状态
        Case 7
            frmDrawCard.ShowCard Me, !No, 4, int记录状态
        Case 11
            frmOtherOutputCard.ShowCard Me, !No, 4, int记录状态
        Case 12
            frmCheckCard.ShowCard Me, !No, 4, int记录状态
        Case 13
            gstrUserName = UserInfo.用户姓名
            With frmAdjust
                .lngBillId = rsTemp!价格id
                .lngMediId = 1
                .Show 1, Me
            End With


        Case Else
            Frm单据See.byt单据 = !单据
            Frm单据See.strNo = !No
            Frm单据See.Show 1, Me
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    
    Select Case frmDrugQuery.intChoose级数
        Case 1
            lbl单位.Caption = "售价单位："
        Case 2
            lbl单位.Caption = "门诊单位："
        Case 3
            lbl单位.Caption = "药库单位："
        Case 4
            lbl单位.Caption = "住院单位："
    End Select
    
    
    Lbl规格.Caption = "规格：" & InDrugStAndard
    lbl库房.Caption = "库房：" & InDeptName
    lbl药品.Caption = "药品：" & InDrugName
    lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
    ReFreshStru
    blnFirst = False
    If Not RefreshData Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mlngLevel = GetSetting("ZLSOFT", "报表\药品库存查询", "库存-单位级数", 1)
    blnFirst = True
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD hh:mm:ss")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD hh:mm:ss")
    RestoreWinState Me
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
    
    Me.LblTitle.Top = 150
    Me.LblTitle.Left = 0
    Me.LblTitle.Width = Me.shpback.Width
    
    With lbl库房
        .Top = Me.LblTitle.Top + 500
        .Left = 200
    End With
    
    With lbl药品
        .Top = Me.lbl库房.Top + Me.lbl库房.Height + 45
        .Left = 200
    End With
    With Me.fgdData
        .Left = 200
        .Width = Me.shpback.Width - 400
        .Top = Me.lbl药品.Top + Me.lbl药品.Height + 45
        .Height = Me.shpback.Height - Me.fgdData.Top - 400
    End With
    
    With Lbl规格
        .Top = lbl药品.Top
        .Left = 200 + Abs((fgdData.Width - .Width)) / 2
    End With
    
    With lbl单位
        .Top = lbl药品.Top
        .Left = 200 + fgdData.Width - .Width
    End With

    Me.lbl期间.Top = Me.LblTitle.Top + 500
    Me.lbl期间.Left = Me.fgdData.Width + Me.fgdData.Left - Me.lbl期间.Width
    If Me.shpback.Width < Me.LblTitle.Width Then
        Me.LblTitle.Visible = False
        Me.fgdData.Visible = False
        Lbl规格.Visible = False
        lbl单位.Visible = False
        Me.lbl期间.Visible = False
    Else
        Me.LblTitle.Visible = True
        Lbl规格.Visible = True
        lbl单位.Visible = True
        Me.fgdData.Visible = True
        Me.lbl期间.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmDrugListAsk
    SaveWinState Me
End Sub

Private Sub mnuEXCEL_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    objPrint.Title.Text = Me.LblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl库房.Caption
     objRow.Add Me.lbl期间.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objPrint.Body = fgdData
     
      Set objRow = New zlTabAppRow
     With objRow
        .Add "打印人:" & UserInfo.用户姓名
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
    
    objPrint.Title.Text = Me.LblTitle.Caption
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl药品.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.Lbl规格.Caption
     objPrint.UnderAppRows.Add objRow
     
     Set objRow = New zlTabAppRow
     objRow.Add Me.lbl单位.Caption
     objRow.Add Me.lbl期间.Caption
     objRow.Add Me.lbl库房.Caption
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

Private Sub mnuFileReFresh_Click()
    fgdData_DblClick
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
        Me.LblTitle.FontSize = 22
        Me.lbl库房.FontSize = 9
        Me.lbl期间.FontSize = 9
        Me.lbl单位.FontSize = 9
        Me.Lbl规格.FontSize = 9
        Me.lbl药品.FontSize = 9
        Me.fgdData.Font.Size = 9
        Me.fgdData.FontFixed.Size = 9

     Case 1
        Me.LblTitle.FontSize = 24
        Me.lbl库房.FontSize = 11
        Me.lbl期间.FontSize = 11
        Me.lbl单位.FontSize = 11
        Me.Lbl规格.FontSize = 11
        Me.lbl药品.FontSize = 11
        Me.fgdData.Font.Size = 11
        Me.fgdData.FontFixed.Size = 11

    Case 2
        Me.LblTitle.FontSize = 28
        Me.lbl库房.FontSize = 15
        Me.lbl期间.FontSize = 15
        Me.lbl单位.FontSize = 15
        Me.Lbl规格.FontSize = 15
        Me.lbl药品.FontSize = 15
        Me.fgdData.Font.Size = 15
        Me.fgdData.FontFixed.Size = 15
    End Select
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.LblTitle.ForeColor)
    Me.LblTitle.ForeColor = lngForeColor
    Me.lbl库房.ForeColor = lngForeColor
    Me.lbl药品.ForeColor = lngForeColor
    Me.Lbl规格.ForeColor = lngForeColor
    Me.lbl单位.ForeColor = lngForeColor
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
'        Case "图形"
'            mnuViewchart_Click
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
    Dim lngRow As Long
    Dim dblCurrNum As Double      '当前余额数量
    Dim dblCurrMny As Double      '当前余额金额
    Dim dblCurrDf As Double      '当前余额差价
    Dim dblStartNum As Double   '开始时数量
    Dim dblStartMny As Double   '开始时金额
    Dim dblStartDf As Double   '开始时差价
    Dim dblinNum As Double     '入库数量
    Dim dblInMny As Double     '入库金额
    Dim dblinDf As Double      '入库差价
    Dim dblOutNum As Double     '出库数量
    Dim dblOutMny As Double     '出库金额
    Dim dblOutDf As Double      '出库差价
    Dim intLevel As Integer     '单位级数
        
    On Error GoTo errHandle
    dblCurrNum = 0: dblCurrMny = 0: dblCurrDf = 0
    dblinNum = 0: dblInMny = 0: dblinDf = 0
    dblOutNum = 0: dblOutMny = 0: dblOutDf = 0
    Load frmDrugListAsk
    With frmDrugListAsk
        .dtpStartDate.Value = CDate(dtpStartDate)
        .dtpEndDate.Value = CDate(dtpEndDate)
        .dtpEndDate.MaxDate = Currentdate()
        .dtpStartDate.MaxDate = .dtpEndDate.MaxDate
        .inDeptId = inDeptId
        
        .InDrugId = InDrugId
        .InDrugName = InDrugName
        .InDrugStAndard = InDrugStAndard
        .InDrugUnit = InDrugUnit
        .Show 1, Me
        RefreshData = False
    End With
    If frmDrugListAsk.blnAskOk = False Then
        Exit Function
    End If
    
    With frmDrugListAsk
        dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD hh:mm:ss")
        dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD hh:mm:ss")
        InDrugId = .InDrugId
        inDeptId = .cob库房.ItemData(.cob库房.ListIndex)
        InDeptName = .cob库房.Text
        InDrugName = .InDrugName
        InDrugStAndard = .InDrugStAndard
        InDrugUnit = .InDrugUnit
        intLevel = frmDrugQuery.intChoose级数
        
        
    End With
    
    '获取当前数据
    ShowFlash "正在装入数据，请稍候…", Me
    DoEvents
    On Error GoTo Err:
    
    Set RecTmpList = New ADODB.Recordset
    strsql = " Select Sum(实际数量)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as 当前数量,Sum(实际金额) as 当前金额,Sum(实际差价) as 当前差价" & _
             " From 药品库存 " & _
             " Where  性质=1 And 药品id=" & InDrugId & IIf(inDeptId = 0, "", "  And 库房id=" & inDeptId)
    With RecTmpList
    Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")

    If Not .EOF Then
        dblCurrNum = IIf(IsNull(.Fields("当前数量").Value), 0, .Fields("当前数量").Value)
        dblCurrMny = IIf(IsNull(.Fields("当前金额").Value), 0, .Fields("当前金额").Value)
        dblCurrDf = IIf(IsNull(.Fields("当前差价").Value), 0, .Fields("当前差价").Value)
    End If
     .Close
     
          
    '获取开始数据
     strsql = " Select sum(入库数量)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as 入库数量,sum(入库金额) as 入库金额,sum(入库差价) as 入库差价, " & _
            "        sum(出库数量)" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & " as 出库数量,sum(出库金额) as 出库金额,sum(出库差价) as 出库差价 " & _
            "  From ( " & _
            "        Select 'Aid' as RiD,id, " & _
            "          Decode(入出系数,1,1,0)*实际数量*付数 as 入库数量, " & _
            "          Decode(入出系数,1,1,0)*零售金额 as 入库金额, " & _
            "          Decode(入出系数,1,1,0)*差价 as 入库差价, " & _
            "          Decode(入出系数,-1,1,0)*实际数量*付数 as 出库数量, " & _
            "          Decode(入出系数,-1,1,0)*零售金额 as  出库金额, " & _
            "          Decode(入出系数,-1,1,0)*差价 as  出库差价 " & _
            "      From 药品收发记录 " & _
            "      Where 审核人 Is Not Null " & IIf(inDeptId = 0, " ", " And 库房id=" & inDeptId) & _
            "          And 药品id=" & InDrugId & "And 审核日期 >= " & " To_date('" & Format(dtpStartDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
            
      Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")
       
      
            dblinNum = IIf(IsNull(.Fields("入库数量").Value), 0, .Fields("入库数量").Value)
            dblInMny = IIf(IsNull(.Fields("入库金额").Value), 0, .Fields("入库金额").Value)
            dblinDf = IIf(IsNull(.Fields("入库差价").Value), 0, .Fields("入库差价").Value)
            dblOutNum = IIf(IsNull(.Fields("出库数量").Value), 0, .Fields("出库数量").Value)
            dblOutMny = IIf(IsNull(.Fields("出库金额").Value), 0, .Fields("出库金额").Value)
            dblOutDf = IIf(IsNull(.Fields("出库差价").Value), 0, .Fields("出库差价").Value)

            dblStartNum = dblCurrNum - dblinNum + dblOutNum
            dblStartMny = dblCurrMny - dblInMny + dblOutMny
            dblStartDf = dblCurrDf - dblinDf + dblOutDf

    End With
    
    '获取明细记录
    '1-外购入库；2-自制入库；3-协药入库；4-其他入库；5-差价调整；6-库房移出；7-部门领用；8-收费处方；9-记帐单处方；10-记帐表处方；11-其他出库；12-盘点；13-调价变动
    
    strsql = "Select max(a.id) as id, Decode(A.单据,1,'外购',2,'自制',3,'协定',4,'入库',5,'差价',6,'移库',7,'领用',8,'处方',9,'处方',10,'摆药',11,'出库',12,'盘点',13,'调价')||A.No as no,A.单据,A.审核日期,decode(a.记录状态,2,'冲销单据',rtrim(A.摘要)|| Decode(B.发票号,null,' ',' 发票号:')||发票号) as 摘要,A.批号,  " & _
            "       sum(Decode(入出系数,1,1,0)*A.实际数量*A.付数" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & ") as 入库数量,  " & _
            "       sum(Decode(入出系数,1,1,0)*A.零售金额) as 入库金额,  " & _
            "       sum(Decode(入出系数,1,1,0)*A.差价) as 入库差价,  " & _
            "       sum(Decode(入出系数,1,0,1)*A.实际数量*A.付数" & IIf(Val(Me.Tag) = 0, "/1", "/" & Me.Tag) & ") as 出库数量,  " & _
            "       sum(Decode(入出系数,1,0,1)*A.零售金额) as  出库金额,  " & _
            "       sum(Decode(入出系数,1,0,1)*A.差价) as  出库差价, A.记录状态  " & _
            " From    药品收发记录 A,药品应付记录 B  " & _
            " Where A.审核人 Is Not Null  And A.id=B.收发id(+) " & _
            "      And A.药品id= " & InDrugId & IIf(inDeptId = 0, "", " And A.库房id=" & inDeptId) & _
            "      And  A.审核日期 between To_date('" & Format(dtpStartDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') And To_date('" & Format(dtpEndDate, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')  " & _
            "    GROUP BY a.no, a.单据, a.审核日期, a.记录状态, a.摘要, b.发票号, a.批号,a.付数 " & _
            " order by A.审核日期 "
         'And Mod(A.记录状态,3)=1
    Set DataRecordSet = New ADODB.Recordset
    With DataRecordSet
        If .State = 1 Then .Close
        Set RecTmpList = zldatabase.OpenSQLRecord(strsql, "RefreshData")
'        If .RecordCount <> 0 Then
            ReFreshStru
'        End If
        
    
        
        lngRow = 2

'        If .RecordCount = 0 Then
'            StopFlash
'            MsgBox "药品在此期间无任何明细!", vbInformation, gstrSysName
'            Exit Function
'        End If
        
        ReFreshStru
        Me.fgdData.rows = Me.fgdData.rows + 1
        Me.fgdData.TextMatrix(lngRow, 0) = Format(dtpStartDate, "yyyy年MM月DD日")
        Me.fgdData.TextMatrix(lngRow, 1) = ""
        Me.fgdData.TextMatrix(lngRow, 2) = "期初余额"
        Me.fgdData.TextMatrix(lngRow, 3) = ""
        Me.fgdData.TextMatrix(lngRow, 4) = ""
        Me.fgdData.TextMatrix(lngRow, 5) = ""
        Me.fgdData.TextMatrix(lngRow, 6) = ""
        Me.fgdData.TextMatrix(lngRow, 7) = ""
        Me.fgdData.TextMatrix(lngRow, 8) = ""
        Me.fgdData.TextMatrix(lngRow, 9) = ""
        Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
        Call RefreshGridColWidth(Me.fgdData, lngRow)
        Me.fgdData.RowData(lngRow) = "0"
        lngRow = lngRow + 1
       
        Select Case intLevel
            Case 1
                lbl单位.Caption = "售价单位：" & InDrugUnit
            Case 2
                lbl单位.Caption = "门诊单位：" & InDrugUnit
            Case 3
                lbl单位.Caption = "药库单位：" & InDrugUnit
            Case 4
                lbl单位.Caption = "住院单位：" & InDrugUnit
        End Select

        Lbl规格.Caption = "规格：" & InDrugStAndard
        lbl库房.Caption = "库房：" & InDeptName
        lbl药品.Caption = "药品：" & InDrugName
        lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
'        lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
       
         If .RecordCount <> 0 Then
                Me.fgdData.rows = Me.fgdData.rows + .RecordCount
         End If
         
            dblinNum = 0
            dblInMny = 0
            dblinDf = 0
            dblOutNum = 0
            dblOutMny = 0
            dblOutDf = 0
         
         Do While Not .EOF
            dblStartNum = dblStartNum + IIf(IsNull(.Fields("入库数量").Value), 0, .Fields("入库数量").Value) - IIf(IsNull(.Fields("出库数量").Value), 0, .Fields("出库数量").Value)
            dblStartMny = dblStartMny + IIf(IsNull(.Fields("入库金额").Value), 0, .Fields("入库金额").Value) - IIf(IsNull(.Fields("出库金额").Value), 0, .Fields("出库金额").Value)
            dblStartDf = dblStartDf + IIf(IsNull(.Fields("入库差价").Value), 0, .Fields("入库差价").Value) - IIf(IsNull(.Fields("出库差价").Value), 0, .Fields("出库差价").Value)
            
            dblinNum = dblinNum + IIf(IsNull(.Fields("入库数量").Value), 0, .Fields("入库数量").Value)
            dblInMny = dblInMny + IIf(IsNull(.Fields("入库金额").Value), 0, .Fields("入库金额").Value)
            dblinDf = dblinDf + IIf(IsNull(.Fields("入库差价").Value), 0, .Fields("入库差价").Value)
            dblOutNum = dblOutNum + IIf(IsNull(.Fields("出库数量").Value), 0, .Fields("出库数量").Value)
            dblOutMny = dblOutMny + IIf(IsNull(.Fields("出库金额").Value), 0, .Fields("出库金额").Value)
            dblOutDf = dblOutDf + IIf(IsNull(.Fields("出库差价").Value), 0, .Fields("出库差价").Value)
            
            Me.fgdData.TextMatrix(lngRow, 0) = Format(.Fields("审核日期").Value, "yyyy年MM月DD日") & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 1) = IIf(IsNull(.Fields("no").Value), "", .Fields("no").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 2) = IIf(IsNull(.Fields("摘要").Value), "", .Fields("摘要").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 3) = IIf(IsNull(.Fields("批号").Value), "", .Fields("批号").Value) & IIf(lngRow Mod 2 = 0, "", " ")
            Me.fgdData.TextMatrix(lngRow, 4) = Format(.Fields("入库数量").Value, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 5) = Format(.Fields("入库金额").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 6) = Format(.Fields("入库差价").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 7) = Format(.Fields("出库数量").Value, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 8) = Format(.Fields("出库金额").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 9) = Format(.Fields("出库差价").Value, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
            Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
            Me.fgdData.TextMatrix(lngRow, 13) = .Fields("记录状态")
            Call RefreshGridColWidth(Me.fgdData, lngRow)
            Me.fgdData.RowData(lngRow) = .Fields("ID").Value
            lngRow = lngRow + 1
            .MoveNext
        Loop
    End With
'    dblCurrNum = dblStartNum
'    dblCurrMny = dblStartMny
'    dblCurrDf = dblStartDf
    
    If dblCurrNum <> 0 Or dblCurrMny <> 0 Or dblCurrDf <> 0 Or _
        dblinNum <> 0 Or dblInMny <> 0 Or dblinDf <> 0 Or _
        dblOutNum <> 0 Or dblOutMny <> 0 Or dblOutDf <> 0 Then
        Me.fgdData.TextMatrix(lngRow, 0) = Format(dtpEndDate, "yyyy年MM月DD日") & Space(lngRow Mod 2)
        Me.fgdData.TextMatrix(lngRow, 1) = ""
        Me.fgdData.TextMatrix(lngRow, 2) = "期末结存"
        Me.fgdData.TextMatrix(lngRow, 3) = ""
        Me.fgdData.TextMatrix(lngRow, 4) = Format(dblinNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 5) = Format(dblInMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 6) = Format(dblinDf, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 7) = Format(dblOutNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 8) = Format(dblOutMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 9) = Format(dblOutDf, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 10) = Format(dblStartNum, "###0.000;-###0.000; ; ")
        Me.fgdData.TextMatrix(lngRow, 11) = Format(dblStartMny, "###0.00;-###0.00; ; ")
        Me.fgdData.TextMatrix(lngRow, 12) = Format(dblStartDf, "###0.00;-###0.00; ; ")
        Call RefreshGridColWidth(Me.fgdData, lngRow)
        Me.fgdData.RowData(lngRow) = "0"
        lngRow = lngRow + 1
    End If
    Me.fgdData.ColWidth(13) = 0
    RefreshData = True
    StopFlash
Exit Function
Err:
   StopFlash
   RefreshData = False
   MsgBox "在获取明细帐数据时,出现了不可预知的错误!", vbInformation, gstrSysName
   Unload Me
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
    Me.Caption = "药品明细帐"
    Me.LblTitle.Caption = GetUnitName & "药品明细帐"
    With Me.fgdData
            .Redraw = False
            For IntCol = 0 To .rows - 1
                .MergeRow(IntCol) = False
            Next
             .Clear
             .Cols = 14
             .rows = 3
             .FixedRows = 2
             .MergeCells = flexMergeRestrictRows
             .SelectionMode = flexSelectionByRow
            For IntCol = 0 To .Cols - 1
                .FixedAlignment(IntCol) = 4
                If IntCol = 0 Then
                    .ColWidth(IntCol) = 1350
                ElseIf IntCol = 1 Then
                    .ColWidth(IntCol) = 800
                ElseIf IntCol = 2 Then
                    .ColWidth(IntCol) = 1200
                ElseIf IntCol = 3 Then
                    .ColWidth(IntCol) = 800
                Else
                    .ColWidth(IntCol) = 800
                End If
                If IntCol <= 3 Then
                    .ColAlignment(IntCol) = 1
                    .MergeCol(IntCol) = True
                Else
                    .MergeCol(IntCol) = False
                    .ColAlignment(IntCol) = 7
                End If
            Next
            .ColWidth(13) = 0
            .MergeCells = 1
            .TextMatrix(0, 0) = "日期"
            .TextMatrix(1, 0) = "日期"
            .TextMatrix(0, 1) = "单据号"
            .TextMatrix(1, 1) = "单据号"
            .TextMatrix(0, 2) = "摘要"
            .TextMatrix(1, 2) = "摘要"
            .TextMatrix(0, 3) = "批号"
            .TextMatrix(1, 3) = "批号"
            .TextMatrix(0, 4) = "入库"
            .TextMatrix(0, 5) = "入库"
            .TextMatrix(0, 6) = "入库"
            .TextMatrix(1, 4) = "数量"
            .TextMatrix(1, 5) = "金额"
            .TextMatrix(1, 6) = "差价"
            .MergeRow(0) = True
            .MergeRow(1) = True
            .TextMatrix(0, 7) = "出库"
            .TextMatrix(0, 8) = "出库"
            .TextMatrix(0, 9) = "出库"
            .TextMatrix(1, 7) = "数量"
            .TextMatrix(1, 8) = "金额"
            .TextMatrix(1, 9) = "差价"
        
            .TextMatrix(0, 10) = "结存"
            .TextMatrix(0, 11) = "结存"
            .TextMatrix(0, 12) = "结存"
            .TextMatrix(1, 10) = "数量"
            .TextMatrix(1, 11) = "金额"
            .TextMatrix(1, 12) = "差价"
'            Call RefreshGridColWidth(Me.fgdData, 0)
            .Redraw = True
    End With

End Sub

Private Function GetLevel(ByVal lng部门id As Long) As Integer
    '判断该部门只是药库而不是药房
    Dim rsTemp As New ADODB.Recordset
    Dim intChoose级数 As Integer
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "Select * From 部门性质说明 " & _
        " Where 部门id=[1] And 工作性质 IN ('西药库','中药库','成药库','制剂室','西药房','中药房','成药房') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "GetLevel", lng部门id)
    If Not rsTemp.EOF Then
        Select Case rsTemp!服务对象
            Case 0
                intChoose级数 = 3
            Case 1, 3
                intChoose级数 = 2
            Case 2
                intChoose级数 = 4
            Case Else
                intChoose级数 = 1
        End Select
    Else
        intChoose级数 = 1
    End If
   
    rsTemp.Close
    
    GetLevel = intChoose级数
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

