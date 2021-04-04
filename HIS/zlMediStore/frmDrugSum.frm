VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugSum 
   BackColor       =   &H8000000C&
   Caption         =   "药品总帐"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmDrugSum.frx":0000
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
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
         Caption         =   "药品总帐"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
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
      _Version        =   "6.0.8169"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
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
            Picture         =   "frmDrugSum.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0526
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":095C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":0FB0
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
            Picture         =   "frmDrugSum.frx":11CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":13E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":181E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSum.frx":1E72
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
            Picture         =   "frmDrugSum.frx":208E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7990
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
      Begin VB.Menu mnuViewBlc 
         Caption         =   "显示差价(&Z)"
         Checked         =   -1  'True
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
Attribute VB_Name = "frmDrugSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public inDeptId As Long            '库房id
Public InDeptName  As String              '库房名称
Public Bln差价 As Boolean        '是否改动差价选择
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



Private Sub fgdData_DblClick()

    If Me.fgdData.RowData(fgdData.Row) = 999999 Then Exit Sub
    If Me.fgdData.TextMatrix(fgdData.Row, 1) = "" Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrSQL As String
    With rsTemp
        StrSQL = "Select id,单据,NO,nvl(价格id,0) as 价格id" & _
                " From 药品收发记录" & _
                " Where No='" & Mid(Trim(Me.fgdData.TextMatrix(fgdData.Row, 1)), 3) & "'" & _
                "       And 单据=" & Me.fgdData.RowData(fgdData.Row)
        If .State = adStateOpen Then .Close
        .Open StrSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then Exit Sub
        
  '1-外购入库；2-自制入库；3-协药入库；4-其他入库；5-差价调整；6-库房移出；7-部门领用；8-收费处方；9-记帐单处方；10-记帐表处方；11-其他出库；12-盘点；13-调价变动
        
        Select Case !单据
        Case 1
            frmPurchaseCard.ShowCard Me, !No, 4
        Case 2
            frmSelfMakeCard.ShowCard Me, !No, 4
        Case 3
            frmAccordDrugCard.ShowCard Me, !No, 4
        Case 4
            frmOtherInputCard.ShowCard Me, !No, 4
        Case 5
            frmDiffPriceAdjustCard.ShowCard Me, !No, 4
        Case 6
            frmTransferCard.ShowCard Me, !No, 4
        Case 7
            frmDrawCard.ShowCard Me, !No, 4
        Case 11
            frmOtherOutputCard.ShowCard Me, !No, 4
        Case 12
            frmCheckCard.ShowCard Me, !No, 4
        Case 13
            gstrUserName = UserInfo.用户姓名
            With frmAdjust
                .lngBillId = rsTemp!价格id
                .lngMediId = 1
                .Show 1, Me
            End With
        Case Else
            Frm单据See.byt单据 = !单据
            Frm单据See.StrNo = !No
            Frm单据See.Show 1, Me
        End Select
    End With

End Sub

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
    Bln差价 = False
    dtpStartDate = Format(DateAdd("m", -1, Currentdate()), "yyyy-MM-DD")
    dtpEndDate = Format(Currentdate(), "yyyy-MM-DD")
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
    On Error Resume Next
    Unload frmDrugSumAsk
    SaveWinState Me
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
    zlMailTo Me.hwnd
End Sub

Private Sub mnuHelpZlWeb_Click()
    zlHomePage Me.hwnd
End Sub


Private Sub mnuViewBlc_Click()
    mnuViewBlc.Checked = Not mnuViewBlc.Checked
    Bln差价 = True
    Call ReFreshStru
    Call RefreshData
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
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
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
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
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
       
        Case "字体"
             PopupMenu mnuViewFont
        Case "前景色"
            mnuViewForeColor_Click
        Case "背景色" '
            mnuViewBackColor_Click
'        Case "帮助"
'            mnuHelpHelp_Click
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
    Dim StrSQL As String
    Dim lngRow As Long
    Dim rsRecord As ADODB.Recordset
'    Dim frmNewAsk As New frmDrugSumAsk
    Dim dblCurr差价 As Double
    Dim dblCurrMny As Double
    Dim dblOutMny As Double
    Dim dblInMny As Double
    Dim DblCgMny As Double
    
    Dim Dbl期初金额 As Double
    Dim Dbl期初差价 As Double
    Dim DBl实际金额 As Double
    Dim DBl实际差价 As Double
    RefreshData = False
    
    If Bln差价 = False Then
        Load frmDrugSumAsk
        With frmDrugSumAsk
            .inDeptId = inDeptId
            .Show 1, Me
            If Not .blnAskOk Then Exit Function
            inDeptId = .cbo库房.ItemData(.cbo库房.ListIndex)
            InDeptName = .cbo库房.Text
            strStartDate = Format(.dtpStartDate.Value, "yyyyMMDD")
            dtpStartDate = Format(.dtpStartDate.Value, "yyyy-MM-DD")
            strEndDate = Format(.dtpEndDate.Value, "yyyyMMDD")
            dtpEndDate = Format(.dtpEndDate.Value, "yyyy-MM-DD")
        End With
    
    Else: Bln差价 = False
    
    End If
    
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

    
    ShowFlash "正在装入数据，请稍候…", Me
    DoEvents
    
    '从现在的实际金额开始,求期初金额
     Set rsRecord = New ADODB.Recordset
     StrSQL = " Select Sum(实际金额) as 实际金额,Sum(实际差价) as 实际差价 " & _
            "From 药品库存 Where 性质=1 " & IIf(inDeptId = 0, "", " And 库房id =" & inDeptId) & _
            "And 药品id In (Select A.药品id From 药品目录 A,药品信息 B Where A.药名id=B.药名id And B.材质分类 In (" & Str材质 & ")) "
    Call SQLTest(App.Title, Me.Caption, StrSQL)
    rsRecord.Open StrSQL, gcnOracle
    Call SQLTest


          
     DBl实际金额 = IIf(IsNull(rsRecord!实际金额), 0, rsRecord!实际金额)
     DBl实际差价 = IIf(IsNull(rsRecord!实际差价), 0, rsRecord!实际差价)
     rsRecord.Close
    
'    StrSql = " Select  sum(A.金额) as 金额,sum(A.差价) as 差价" & _
            "  From ( Select  " & _
            "         B.金额,B.差价 " & _
            "         From 药品收发汇总 B,药品入出类别 C " & _
            "         Where B.日期 >=To_Date('" & strStartDate & "','yyyymmdd') And " & IIf(inDeptId = 0, "", "B.库房ID = " & inDeptId & " And ") & " B.类别id=C.id" & _
            "               And B.药品id In (Select X.药品id From 药品目录 X,药品信息 Y Where X.药名id=Y.药名id And Y.材质分类 In (" & Str材质 & "))) A"
    
    StrSQL = " Select  sum(金额) as 金额,sum(差价) as 差价" & _
            "    From 药品收发汇总 B " & _
            "   Where B.日期 >=To_Date('" & strStartDate & "','yyyymmdd') " _
              & IIf(inDeptId = 0, "", " and B.库房ID = " & inDeptId) _
              & " And B.药品id In " _
              & " (Select X.药品id From 药品目录 X,药品信息 Y Where X.药名id=Y.药名id And Y.材质分类 In (" & Str材质 & ")) "
    
    Set DataRecordSet = New ADODB.Recordset
    Call SQLTest(App.Title, Me.Caption, StrSQL)
    DataRecordSet.Open StrSQL, gcnOracle
    Call SQLTest
    
    
    Dbl期初金额 = DBl实际金额
    Dbl期初差价 = DBl实际差价
    With DataRecordSet
        
            Dbl期初金额 = Dbl期初金额 - IIf(IsNull(!金额), 0, !金额)
            Dbl期初差价 = Dbl期初差价 - IIf(IsNull(!差价), 0, !差价)
       
    End With
    
    
    
    '求期间发生数
    '1-外购入库；2-自制入库；3-协药入库；4-其他入库；5-差价调整；6-库房移出；7-部门领用；8-收费处方；9-记帐单处方；10-记帐表处方；11-其他出库；12-盘点；13-调价变动
    
        StrSQL = " Select A.审核日期,A.NO,单据,ltrim(C.名称) as 摘要, " & _
            "        abs(sum(A.采购金额)) as 采购金额,sum(A.入库金额) as 入库金额,sum(A.出库金额) As 出库金额 ,sum(A.差价) As 差价" & _
            " From ( " & _
            "     Select 'A'||id as RiD," & _
            "         单据 as 单据,  " & _
            "         Decode(单据,1,'外购',2,'自制',3,'协定',4,'入库',5,'差价',6,'移库',7,'领用',8,'处方',9,'处方',10,'摆药',11,'出库',12,'盘点',13,'调价')||No as no,  " & _
            "         审核日期 ,  " & _
            "         供药单位ID as 供药单位ID  ," & _
            "         库房id as 库房id ,入出系数*Decode(差价,null,0,差价) As 差价," & _
            "         (入出系数* Decode(单据,5,0,13,0,1)*成本金额) as 采购金额, " & _
            "         Decode(入出系数,-1,0,1)*零售金额 as 入库金额, " & _
            "         Decode(入出系数,1,0,1)*零售金额 as  出库金额 " & _
            "     From 药品收发记录 A, 药品目录 X, 药品信息 Y " & _
            "     Where A.药品id = X.药品id AND X.药名id = Y.药名id And Y.材质分类 In (" & Str材质 & ") " & _
            "         And 审核日期 <=to_date('" & strEndDate & "','yyyymmdd')+1" & "And 审核日期 >=to_date('" & strStartDate & "','yyyymmdd') " & _
            "         And 审核人 Is Not Null " & IIf(inDeptId = 0, "", " And 库房ID = " & inDeptId) & " And Mod(记录状态,3)=1 " & _
            "     ) A,部门表 B,药品供应商 C " & _
            " Where A.供药单位id=C.id(+) And A.库房id=B.id(+) " & _
            " Group by A.审核日期,A.单据,A.NO,C.名称 " & _
            " having sum(A.采购金额)<>0 or sum(A.入库金额) <>0 or sum(A.出库金额) <>0 order by A.审核日期"
            
            'Decode(单据,5,0,13,0,1)*成本金额 as 采购金额
    DataRecordSet.Close
    With DataRecordSet
        Call SQLTest(App.Title, Me.Caption, StrSQL)
        .Open StrSQL, gcnOracle
        Call SQLTest
'        If .RecordCount = 0 Then
'            StopFlash
'            MsgBox "在此条件和权限范围中,无任何总帐记录！", vbInformation, gstrSysName
'            Exit Function
'        End If
        dblInMny = 0
        DblCgMny = 0
        dblOutMny = 0
        dblCurrMny = Dbl期初金额
        dblCurr差价 = Dbl期初差价
        Dim colWidth As Long
        Me.fgdData.Rows = .RecordCount + 3
        lngRow = 2
        colWidth = 0
        Me.fgdData.Redraw = False
        
        Call RefreshGridColWidth(Me.fgdData, 0)
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
             dblInMny = dblInMny + IIf(IsNull(.Fields("入库金额").Value), 0, .Fields("入库金额").Value)
             DblCgMny = DblCgMny + IIf(IsNull(.Fields("采购金额").Value), 0, .Fields("采购金额").Value)
             dblOutMny = dblOutMny + IIf(IsNull(.Fields("出库金额").Value), 0, .Fields("出库金额").Value)
             dblCurrMny = dblCurrMny + IIf(IsNull(.Fields("入库金额").Value), 0, .Fields("入库金额").Value) - IIf(IsNull(.Fields("出库金额").Value), 0, .Fields("出库金额").Value)
             dblCurr差价 = dblCurr差价 + IIf(IsNull(.Fields("差价").Value), 0, .Fields("差价").Value)
             Me.fgdData.TextMatrix(lngRow, 0) = IIf(Format(.Fields("审核日期").Value, "yyyy-mm-dd") = "1932-09-09", dtpStartDate, Format(.Fields("审核日期").Value, "yyyy-mm-dd")) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 1) = IIf(IsNull(.Fields("no").Value), "", IIf(Format(.Fields("审核日期").Value, "yyyy-mm-dd") = "1932-09-09", "", .Fields("no").Value)) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 2) = IIf(Format(.Fields("审核日期").Value, "yyyy-mm-dd") = "1932-09-09", "期初发生额", IIf(IsNull(.Fields("摘要").Value), "", .Fields("摘要").Value)) & IIf(lngRow Mod 2 = 0, "", " ")
             Me.fgdData.TextMatrix(lngRow, 3) = " " & Format(.Fields("采购金额").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 4) = Format(.Fields("入库金额").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 5) = " " & Format(.Fields("出库金额").Value, "##,###0.00;-##,###0.00; ; ")
             Me.fgdData.TextMatrix(lngRow, 6) = Format(dblCurrMny, "##,###0.00;-##,###0.00; ; ")
             
             If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(lngRow, 7) = Format(dblCurr差价, "##,###0.00;-##,###0.00; ; ")
             
             Me.fgdData.RowData(lngRow) = IIf(IsNull(.Fields("单据").Value), 999999, .Fields("单据").Value)
             Call RefreshGridColWidth(Me.fgdData, lngRow)
             lngRow = lngRow + 1
             .MoveNext
           Loop
        End If
            Me.fgdData.MergeRow(1) = True
            Me.fgdData.TextMatrix(1, 0) = "期初"
            Me.fgdData.TextMatrix(1, 1) = "期初"
            Me.fgdData.TextMatrix(1, 2) = "期初"
            Me.fgdData.TextMatrix(1, 3) = ""
            Me.fgdData.TextMatrix(1, 4) = " "
            Me.fgdData.TextMatrix(1, 5) = ""
            Me.fgdData.TextMatrix(1, 6) = Format(Dbl期初金额, "##,###0.00;-##,###0.00; ; ")
            If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(1, 7) = Format(Dbl期初差价, "##,###0.00;-##,###0.00; ; ")
         
            
'        If dblInMny <> 0 Or DblCgMny <> 0 Or dblOutMny <> 0 Then
            Me.fgdData.MergeRow(Me.fgdData.Rows - 1) = True
            Me.fgdData.RowData(lngRow) = 999999
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 0) = "合计"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 1) = "合计"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 2) = "合计"
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 3) = " " & Format(DblCgMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 4) = Format(dblInMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 5) = " " & Format(dblOutMny, "##,###0.00;-##,###0.00; ; ")
            Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 6) = Format(dblCurrMny, "##,###0.00;-##,###0.00; ; ")
            If mnuViewBlc.Checked Then Me.fgdData.TextMatrix(Me.fgdData.Rows - 1, 7) = Format(dblCurr差价, "##,###0.00;-##,###0.00; ; ")
            Call RefreshGridColWidth(Me.fgdData, lngRow)
'        End If
        
        Me.fgdData.Redraw = True
    End With
    lbl库房.Caption = "库房：" & InDeptName
    lbl期间.Caption = "期间:" & dtpStartDate & "  至  " & dtpEndDate
    StopFlash
    RefreshData = True
Exit Function
Err:
    StopFlash
    RefreshData = False
    Me.fgdData.Redraw = True
    MsgBox "在获取药品总帐时,出现了不可预知的错误!", vbInformation, gstrSysName
End Function

Private Sub ReFreshStru()
    '-------------------------------------------------------------------------
    '--功能:重新获的表头结构
    '--参数:
    '--返回:
    '-------------------------------------------------------------------------
    Dim IntCol As Long
    Me.Caption = "药品总帐"
    Me.lblTitle.Caption = GetUnitName & "药品总帐"
    With Me.fgdData
            .Redraw = False
             .Clear
             .Cols = 7
             If mnuViewBlc.Checked Then .Cols = 8
             .Rows = 3
             .MergeCells = flexMergeRestrictRows
             For IntCol = 0 To .Rows - 1
                .MergeRow(IntCol) = False
                .CellAlignment = 1
             Next
            For IntCol = 0 To .Cols - 1
                .FixedAlignment(IntCol) = 4
            Next
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            If mnuViewBlc.Checked Then .ColAlignment(7) = 7
            
            .colWidth(0) = 400
            .colWidth(1) = 600
            .colWidth(2) = 400
            .colWidth(3) = 800
            .colWidth(4) = 800
            .colWidth(5) = 800
            .colWidth(6) = 800
            If mnuViewBlc.Checked Then .colWidth(7) = 800
            
            .TextMatrix(0, 0) = "日期"
            .TextMatrix(0, 1) = "单据号"
            .TextMatrix(0, 2) = "摘要"
            .TextMatrix(0, 3) = "采购金额"
            .TextMatrix(0, 4) = "入库金额"
            .TextMatrix(0, 5) = "出库金额"
            .TextMatrix(0, 6) = "结存金额"
            If mnuViewBlc.Checked Then .TextMatrix(0, 7) = "差价"
            .Redraw = True
    End With
End Sub
